import sys
import os
current = os.path.dirname(os.path.realpath(__file__))
parent = os.path.dirname(current)
sys.path.append(parent)
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from slugify import slugify
from pyexcel_ods3 import get_data, save_data
from collections import OrderedDict
import settings as s
import glob

# IMPORTANT:
# replace ^.*$ with &
def parse():
    dlist = []
    data = OrderedDict()
    dlist.append(["NO", "NAME", "VENDOR",	"SECTION", "DIAGRAM", "FILENAME", "LINK"])
    no = 1
    for fileinput in glob.glob(s.PDF_RESULT_PATH + os.sep + "*.pdf"):
        reader = PdfReader(fileinput)
        number_of_pages = len(reader.pages)
        first = True
        name = vendor = section = diagram = ''
        for i in range(0, number_of_pages):
            # print(number_of_pages)
            page = reader.pages[i]
            lines = page.extract_text().split("\n")
            # input(lines)
            if len(lines) >= 1 and lines[0] == 'https://www.messicks.com':
                if not first:
                    filename = slugify("{}{}{}{}".format(name, vendor, section, diagram) )+".pdf"
                    writer.write(s.PDF_EXTRACT_PATH + os.sep + filename)
                    print(name, vendor, section, diagram)
                    link = '=HYPERLINK(CONCATENATE($Sheet2.$A$1,F{}),"OPEN PDF")'.format(i+1)
                    dlist.append([no, name, vendor, section, diagram, filename, link])
                    writer.close()
                    no+=1
                writer = PdfWriter()
                writer.add_page(page)
                first = False
                name = lines[2]
                vendor = lines[3].replace('VENDOR:', "").strip()
                section = lines[4].replace('SECTION:', "").strip()
                diagram = lines[5].replace('DIAGRAM:', "").strip()
            else:
                writer.add_page(page)
    data.update({"Sheet1": dlist})
    exloc = "file:///{}".format(s.PDF_EXTRACT_PATH + os.sep).replace("\\", "/")
    data.update({"Sheet2": [[exloc]]})
    return data
    


def main():
    for f in glob.glob(s.PDF_EXTRACT_PATH + os.sep + "*.pdf"):
        os.remove(f)
    data = parse()
    save_data(s.ODF_RESULT_PATH + os.sep +"result.ods", data)
    input("End Process..")    

if __name__ == '__main__':
    main()
