from py_pdf_parser.loaders import load_file
from py_pdf_parser.visualise import visualise
from py_pdf_parser import tables

#document = load_file("SOC TEMP/23253 Florida Power & Light.pdf")
document = load_file("SOC TEMP/23244 Florida Power & Light.pdf")

#visualise(document)

desc_element = document.elements.filter_by_text_contains("Description").extract_single_element()

if desc_element.text().__contains__("RTU") or desc_element.text().__contains__("IACS") or desc_element.text().__contains__("SNW Retrofit"):
    desc_text = desc_element.text()
else:
    desc_text = document.elements.below(desc_element)[0].text()

print(desc_text)
