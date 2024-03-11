import os, glob
import sys
current = os.path.dirname(os.path.realpath(__file__))
parent = os.path.dirname(current)
sys.path.append(parent)
import time
from urllib.parse import urlparse
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from selenium.webdriver.support.select import Select
# import random
import settings as s
import argparse
import warnings
# import validators
# import pyexcel_ods3 as pods
from slugify import slugify
# import shutil
# from collections import OrderedDict
from html import unescape
import unicodedata
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import xlwings as xw
from urllib.parse import urlparse
from urllib.request import urlretrieve

warnings.filterwarnings("ignore", category=UserWarning)

def genfilename(title, section, diagram):
    pathname = s.PDF_EXTRACT_PATH + os.sep + slugify("{}{}{}".format(title, section, diagram))
    if len(pathname) < 255:
        return pathname + ".pdf"
    else:
        while True:
            second = str(int(time.time()))
            newpathname = pathname[0:244] + second
            if os.path.exists(newpathname + ".pdf"):
                continue
            return newpathname + ".pdf"

def parse(modelid, xlsheet2):
    diagramlist = []
    lastrow = xlsheet2.range('A' + str(xlsheet2.cells.last_cell.row)).end('up').row + 1
    success = 0 
    failed = 0
    trial = 0
    first = True
    start = 0
    end = 0
    # breakpoint()
    for row in range(2, lastrow + 1):
        if first and xlsheet2[f'B{row}'].value == modelid:
            first = False
            start = row
        if not first and xlsheet2[f'B{row}'].value != modelid:
            end = row
            break
    
    for row in range(start, end):
        # print(xlsheet2[f'B{row}'].value)
        dowloadurl = xlsheet2[f'G{row}'].value
        title, section, diagram  = xlsheet2[f'D{row}'].value, xlsheet2[f'E{row}'].value, xlsheet2[f'F{row}'].value
        pathname = genfilename(title, section, diagram)
        filename = os.path.basename(pathname)
        try:
            print("download for", title, section, diagram, end="... ", flush=True)
            if not os.path.exists(pathname):
                # breakpoint()
                urlretrieve(dowloadurl, pathname)
                print("OK")
            else:
                print("File Exists")
            link = '=HYPERLINK(CONCATENATE(Sheet3!$A$1,"{}"),"OPEN PDF")'.format(filename)
            diagramlist.append([section, diagram, link, dowloadurl, pathname])
            xlsheet2[f"H{row}"].value = link
            success += 1

        except:
            xlsheet2[f"H{lastrow}"].value = 'NOT FOUND'
            failed += 1


    return diagramlist, success, failed

def main():
    parser = argparse.ArgumentParser(description="Catalog Product Downloader")
    parser.add_argument('-i', '--input', type=str,help="Source File")
    
    args = parser.parse_args()
    isExist = os.path.exists(args.input)
    if isExist == False :
        input('Please check the XLS file')
        sys.exit()

    source = args.input
    # driver = browser_init()
    # driver.maximize_window()
    print('Opening the Source Excel File...', end="", flush=True)
    xlbook = xw.Book(source)
    xlsheet1 = xlbook.sheets["Sheet1"]
    xlsheet2 = xlbook.sheets["Sheet2"]
    print('OK')
    maxrow = xlsheet1.range('C' + str(xlsheet1.cells.last_cell.row)).end('up').row
    # breakpoint()
    for i in range(2, maxrow + 1):
        merger = PdfMerger()
        vendor = xlsheet1[f'A{i}'].value
        title = vendor + " " + xlsheet1[f'C{i}'].value + " Parts"
        if xlsheet1[f'E{i}'].value == 'NO':
            print('search',xlsheet1[f'B{i}'].value)
            diagramlist, success, failed = parse(modelid=xlsheet1[f'B{i}'].value, xlsheet2=xlsheet2)
            # break
            # diagramlist.append([section, diagram, link, dowloadurl, pathname])
            if len(diagramlist) > 0:
                for diagram in diagramlist:
                    filename = str(diagram[2]).split(",")[1].replace('"',"").replace(")","")
                    try:
                        merger.append(s.PDF_EXTRACT_PATH + os.sep + filename)
                        # print("merge", filename)
                    except:
                        urlretrieve(diagram[3], diagram[4])
                        merger.append(s.PDF_EXTRACT_PATH + os.sep + filename)
                        # breakpoint()

                merger.write(s.PDF_JOIN_PATH + os.sep + slugify(title) + ".pdf")
                link = '=HYPERLINK(CONCATENATE(Sheet3!$A$2,"{}"),"OPEN PDF")'.format(slugify(title) + ".pdf")
                xlsheet1[f'E{i}'].value = 'YES'
                xlsheet1[f'F{i}'].value = link
                xlsheet1[f'G{i}'].value = f"PDF Download Success={success}, Failed={failed}"
            else:
                xlsheet1[f'E{i}'].value = 'YES'
                xlsheet1[f'F{i}'].value = "FAILED"
                xlsheet1[f'G{i}'].value = "No PDF can not be Downloaded"
    # driver.quit()


    input("End Process..")    


if __name__ == '__main__':
    main()
