import csv
import shutil
from py_pdf_parser.loaders import load_file

import os
script_dir = os.path.dirname(__file__)  # <-- absolute dir the script is in
rel_path = "SOC Temp"
SOC_DOC_PATH = os.path.join(script_dir, rel_path)

CSV_FILE = r"Saved_Jobs.csv"
CSV_FILE2 = r"Parsed_Jobs.csv"
SETUP_CSV = r"Setup.csv"
CLIENT_STRINGS = []
PRIMARY_SEARCH_TERM1 = []
PRIMARY_SEARCH_TERM2 = []
SECONDARY_SEARCH_TERM1 = []
SECONDARY_SEARCH_TERM2 = []
SECONDARY_SEARCH_OFFSET1 = []
SECONDARY_SEARCH_OFFSET2 = []

def getParams():
    global FILE_PATH
    if os.path.isfile(SETUP_CSV):   # check if CSV file exists, grab existing saved data
        with open(SETUP_CSV, 'r') as csvFile:
            csvReader = csv.reader(csvFile, delimiter=',')
            for data in csvReader:
                term = data[0].upper()
                if term == "PATH":
                    FILE_PATH = data[1]
                elif term == "CLIENT":
                    CLIENT_STRINGS.append(data[1])
                elif term == "PRIMARY":
                    PRIMARY_SEARCH_TERM1.append(data[1])
                    PRIMARY_SEARCH_TERM2.append(data[2])
                elif term == "SECONDARY":
                    SECONDARY_SEARCH_TERM1.append(data[1])
                    SECONDARY_SEARCH_TERM2.append(data[2])
                    SECONDARY_SEARCH_OFFSET1.append(int(data[3]))
                    SECONDARY_SEARCH_OFFSET2.append(int(data[4]))
    else:   # otherwise print that CSV doesn't exist
        print("\nCSV file {} doesn't exist! ".format(SETUP_CSV))


class File:   # Default File Class to hold and pass data around
    def __init__(self):
        self.file_path = r""
        self.is_rtu = False
        self.desc = ""
        self.job_num = ""
        self.po_num = ""
        self.ship_date = ""


def getPDFData(file_path):   # Function to scrape data in PDF file and determine if panel is of interest to you
    file_data = File()
    document = load_file(file_path)

    desc_element = document.elements.filter_by_text_contains("Description").extract_single_element()

    for index, term in enumerate(PRIMARY_SEARCH_TERM1):   # loop through list of primary search terms

        term1 = PRIMARY_SEARCH_TERM1[index].lower()   # get primary search term pair, secondary search term pair
        term2 = PRIMARY_SEARCH_TERM2[index].lower()   # and secondary offset pair for this primary search term pair
        term3 = SECONDARY_SEARCH_TERM1[index].lower()
        term4 = SECONDARY_SEARCH_TERM2[index].lower()
        offset1 = SECONDARY_SEARCH_OFFSET1[index]
        offset2 = SECONDARY_SEARCH_OFFSET2[index]
        if desc_element.text().__contains__(term1) and desc_element.text().__contains__(term2):
            desc_text = desc_element.text()
        else:
            desc_text = document.elements.below(desc_element)[0].text()

        desc = desc_text.lower()  # compare all text as lowercase

        if desc.__contains__(term1) and desc.__contains__(term2):   # if primary search term pairs are found, get data
            file_data.is_rtu = True   # set boolean is_rtu to true
            if term4 == "":   # if secondary search term2 is blank, don't strip end of description
                file_data.desc = desc_text[desc.find(term3)+offset1:].strip('\n').replace('\n', ' ').strip()
            else:   # else strip both ends of description based on terms
                file_data.desc = desc_text[
                                 desc.find(term3) + offset1:desc.find(term4)+offset2].strip(
                    '\n').replace('\n', ' ').strip()
            # Strip purchase order, job number and ship date strings from PDF data
            po_element = document.elements.filter_by_text_equal("Purchase Order No.").extract_single_element()
            po_text = document.elements.below(po_element)[0].text()
            job_element = document.elements.filter_by_text_equal("NovaTech Order No.").extract_single_element()
            job_text = document.elements.below(job_element)[0].text()
            ship_text = document.elements.filter_by_text_contains("Est. Ship Date").extract_single_element().text()
            ship_text = ship_text[ship_text.find("Date") + 4:].strip('\n').replace('\n', ' ')
            file_data.job_num = job_text.strip()   # store job number, po number and ship date
            file_data.po_num = po_text.strip()
            file_data.ship_date = ship_text.strip()

    return file_data   # return file data

def main():
    getParams()
    savedSet = set()   # initialize savedSet
    if os.path.isfile(CSV_FILE):   # check if CSV file exists, grab existing saved data
        with open(CSV_FILE, 'r+') as csvFile:
            csvReader = csv.reader(csvFile, delimiter=',')
            for row in csvReader:
                savedSet.add((row[0], row[1], row[2], row[3], row[4]))
    else:   # otherwise print that CSV doesn't exist and will be created
        print("\nNo CSV file {}, file will be created. ".format(CSV_FILE))
    parsedSet = set()  # initialize savedSet
    if os.path.isfile(CSV_FILE2):  # check if CSV file exists, grab existing saved data
        with open(CSV_FILE2, 'r+') as csvFile:
            csvReader = csv.reader(csvFile, delimiter=',')
            for row in csvReader:
                parsedSet.add((row[1]))
    else:  # otherwise print that CSV doesn't exist and will be created
        print("\nNo CSV file {}, file will be created. ".format(CSV_FILE2))

    nameSet = set()   # initialize nameSet to store files of interest
    for file in os.listdir(FILE_PATH):   # search FILE_PATH for all files, loop through list
        fullPath = os.path.join(FILE_PATH, file)   # get full path including filename and type
        if os.path.isfile(fullPath):   # if it's a file, check if file name contains any CLIENT_STRINGS as defined above
            if fullPath not in parsedSet:  # check if file path already saved, if not continue
                for client in CLIENT_STRINGS:
                    if file.__contains__(client):   # if it does, store full file path to check PDF data later
                        nameSet.add((file, fullPath))

    print("\nNew Jobs Found: {}".format(nameSet.__len__()))   # print total jobs matching CLIENT_STRINGS
    print("\nSaved RTU Jobs Found: {}".format(savedSet.__len__()))   # print total existing jobs saved

    retrievedSet = set()   # initialize retrievedSet to hold all files that are new jobs we care about
    for file in nameSet:
        pdf_data = File()
        PDFError = False
        try:
            pdf_data = getPDFData(file[1])   # for all files of interest, grab pdf data
        except:
            PDFError = True

        files_path = file[1]   # get SOC file full path
        if pdf_data.is_rtu and not PDFError:   # if file is of interest, add data to set
            retrievedSet.add((pdf_data.desc, pdf_data.job_num, pdf_data.po_num, pdf_data.ship_date, files_path))
        elif PDFError:
            retrievedSet.add(("", "", "", "", files_path))

    newSet = sorted(retrievedSet - savedSet)   # check for new jobs, any jobs in retrievedSet but not in savedSet
    if newSet.__len__() > 0:   # if there are new jobs, print their data and write to CSV file
        with open(CSV_FILE, 'a', newline='') as csvFile:   # open CSV file to write data to
            filesWriter = csv.writer(csvFile)
            print("\nNew RTU Jobs({})".format(newSet.__len__()))   # print total number of new jobs found
            for i in newSet:   # loop through new set, print the data, write to CSV file, and copy SOC to SOC_DOC_PATH
                if i[0] != "":
                    print("****************")
                    print("Job Description: {}".format(i[0]))
                    print("Job Number     : {}".format(i[1]))
                    print("Job PO Number  : {}".format(i[2]))
                    print("Job Ship Date  : {}".format(i[3]))
                else:
                    print("****************")
                    print("Job Couldn't be Read at file path: {}".format(i[4]))
            print("****************")
            response = (input("\nSave Jobs? (Y/N): "))
            if response == 'Y' or response == 'y':
                for i in newSet:   # loop through new set, update SavedJobs CSV file, and copy SOC to SOC_DOC_PATH
                    if i[0] != "":
                        filesWriter.writerow([i[0], i[1], i[2], i[3], i[4]])
                        shutil.copy(i[4], SOC_DOC_PATH)
                    else:
                        filesWriter.writerow([i[0], i[1], i[2], i[3], i[4]])
                        shutil.copy(i[4], SOC_DOC_PATH)
                if nameSet.__len__() > 0:  # if there are new jobs and new RTU, update ParsedJobs CSV file
                    with open(CSV_FILE2, 'a', newline='') as csvFile2:  # open CSV file to write data
                        filesWriter = csv.writer(csvFile2)
                        for i in nameSet:
                            filesWriter.writerow([i[0], i[1]])
                print("\nSaved Jobs updated and new SOCs copied.\n")
            else:
                print("\nNothing was saved.\n")
    else:   # else print no new jobs found
        if nameSet.__len__() > 0:  # if there are new jobs but no RTU, update CSV file
            with open(CSV_FILE2, 'a', newline='') as csvFile2:  # open CSV file to write data to
                filesWriter = csv.writer(csvFile2)
                for i in nameSet:
                    filesWriter.writerow([i[0], i[1]])
        print("\nNo New RTU Jobs Found.\n")


main()   # call main function when main.py is run
