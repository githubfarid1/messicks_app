from PyPDF2 import PdfMerger, PdfReader, PdfWriter
fileinput = r"pdfs_result\86272.pdf"
reader = PdfReader(fileinput)
number_of_pages = len(reader.pages)
for i in range(0, number_of_pages):
    print(number_of_pages)
    page = reader.pages[2]
    lines = page.extract_text().split("\n")
    print(lines)