import os, glob
import sys
current = os.path.dirname(os.path.realpath(__file__))
parent = os.path.dirname(current)
sys.path.append(parent)
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager as CM
import time
from urllib.parse import urlparse
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from selenium.webdriver.support.select import Select
import random
import settings as s
import argparse
import warnings
import validators
import pyexcel_ods3 as pods
from slugify import slugify
import shutil
from collections import OrderedDict
from html import unescape
import unicodedata
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import xlwings as xw

warnings.filterwarnings("ignore", category=UserWarning)
def browser_init():
    # WARNING: AGAR BISA SET LOKASI DOWNLOAD PDF, MAKA NAMA PROFILE HARUS "DEFAULT" 
    cud = s.CHROME_USER_DATA
    cp = s.CHROME_PROFILE
    options = webdriver.ChromeOptions()
    options.add_argument("user-data-dir={}".format(cud))
    options.add_argument("profile-directory={}".format(cp))
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    options.add_argument("--window-size=800,600")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    profile = {"download.default_directory": s.CHROME_DOWNLOAD_PATH + os.sep, 
                "download.extensions_to_open": "applications/pdf",
                "download.prompt_for_download": False,
                'profile.default_content_setting_values.automatic_downloads': 1,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True                   
                }
    options.add_experimental_option("prefs", profile)

    return webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)


def getlatestfile(folder):
    list_of_files = glob.glob(folder + os.sep + "*.pdf") # * means all if need specific format then *.csv
    return max(list_of_files, key=os.path.getctime)    

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

def parse(url, driver, xlsheet2):
    driver.get(url)
    firstopt1txt = ''
    firstopt2txt = ''
    first = True
    time.sleep(2)
    curr=driver.current_window_handle
    title = driver.find_element(By.CSS_SELECTOR, "h1#model-title").text
    vendor = driver.find_element(By.XPATH, "/html/body/div[6]/button").get_attribute('data-brand')
    diagramlist = []
    lastrow = xlsheet2.range('A' + str(xlsheet2.cells.last_cell.row)).end('up').row + 1
    success = 0 
    failed = 0
    trial = 0
    while True:
        driver.find_element(By.CSS_SELECTOR, "span.ms-5 span.btn-prev-diagram").click()
        try:
            WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.cc-part-tile")))
            trial = 0
        except:
            trial += 1
            if trial == 3:
                return diagramlist, title, success, failed
                # break
        time.sleep(2)
        if first:
            first = False
            firstopt1txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select ms-3 me-3 section-list']")).first_selected_option.text
            firstopt2txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select diagram-list']")).first_selected_option.text
            # breakpoint()
            driver.find_element(By.CSS_SELECTOR, "span.print-pdf").click()
            time.sleep(3)
            time.sleep(random.randint(5, 7))
            print("download for", title, firstopt1txt, firstopt2txt)
            section, diagram  = firstopt1txt, firstopt2txt
            time.sleep(random.randint(2, 5))
            if  len(driver.window_handles) > 1:
                for handle in driver.window_handles:
                    if handle != curr:
                        driver.switch_to.window(handle)
                        time.sleep(2)
                        try:
                            driver.close()
                        except:
                            pass
                        if len(driver.window_handles) == 1:
                            break
                driver.switch_to.window(curr)
                xlsheet2[f"A{lastrow}"].value = vendor
                xlsheet2[f"B{lastrow}"].value = title 
                xlsheet2[f"C{lastrow}"].value = unicodedata.normalize('NFKC',unescape(section))
                xlsheet2[f"D{lastrow}"].value = unicodedata.normalize('NFKC',unescape(diagram))
                xlsheet2[f"E{lastrow}"].value = 'NOT FOUND'
                failed += 1
            else:
                # filename = slugify("{}{}{}".format(title, section, diagram) )+".pdf"
                filename = genfilename(title, section, diagram) + ".pdf"
                latestfile = getlatestfile(s.CHROME_DOWNLOAD_PATH)
                shutil.copyfile(latestfile, filename)
                # link = '=HYPERLINK(CONCATENATE($Sheet3.$A$1,"{}"),"OPEN PDF")'.format(filename)
                link = '=HYPERLINK(CONCATENATE(Sheet3!$A$1,"{}"),"OPEN PDF")'.format(filename)
                diagramlist.append([vendor, title, unicodedata.normalize('NFKC',unescape(section)), unicodedata.normalize('NFKC',unescape(diagram)), link])
                xlsheet2[f"A{lastrow}"].value = vendor
                xlsheet2[f"B{lastrow}"].value = title 
                xlsheet2[f"C{lastrow}"].value = unicodedata.normalize('NFKC',unescape(section))
                xlsheet2[f"D{lastrow}"].value = unicodedata.normalize('NFKC',unescape(diagram))
                xlsheet2[f"E{lastrow}"].value = link
                success += 1
            lastrow += 1
            continue
        opt1txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select ms-3 me-3 section-list']")).first_selected_option.text
        opt2txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select diagram-list']")).first_selected_option.text

        if opt1txt == firstopt1txt and opt2txt == firstopt2txt:
            break
        # breakpoint()
        driver.find_element(By.CSS_SELECTOR, "span.print-pdf").click()
        time.sleep(random.randint(5, 7))

        print("download for", title, opt1txt, opt2txt)
        section, diagram  = opt1txt, opt2txt
        time.sleep(random.randint(2, 5))
        if  len(driver.window_handles) > 1:
            for handle in driver.window_handles:
                if handle != curr:
                    driver.switch_to.window(handle)
                    time.sleep(2)
                    try:
                        driver.close()
                    except:
                        pass
                    if len(driver.window_handles) == 1:
                        break
            driver.switch_to.window(curr)
            xlsheet2[f"A{lastrow}"].value = vendor
            xlsheet2[f"B{lastrow}"].value = title 
            xlsheet2[f"C{lastrow}"].value = unicodedata.normalize('NFKC',unescape(section))
            xlsheet2[f"D{lastrow}"].value = unicodedata.normalize('NFKC',unescape(diagram))
            xlsheet2[f"E{lastrow}"].value = 'NOT FOUND'
            failed += 1
        else:
            # filename = slugify("{}{}{}".format(title, section, diagram) )+".pdf"
            filename = genfilename(title, section, diagram) + ".pdf"
            latestfile = getlatestfile(s.CHROME_DOWNLOAD_PATH)
            shutil.copyfile(latestfile, filename)
            link = '=HYPERLINK(CONCATENATE(Sheet3!$A$1,"{}"),"OPEN PDF")'.format(filename)

            diagramlist.append([vendor, title, unicodedata.normalize('NFKC',unescape(section)), unicodedata.normalize('NFKC',unescape(diagram)), link])
            xlsheet2[f"A{lastrow}"].value = vendor
            xlsheet2[f"B{lastrow}"].value = title 
            xlsheet2[f"C{lastrow}"].value = unicodedata.normalize('NFKC',unescape(section))
            xlsheet2[f"D{lastrow}"].value = unicodedata.normalize('NFKC',unescape(diagram))
            xlsheet2[f"E{lastrow}"].value = link
            success += 1
        lastrow += 1
    # driver.quit()
    return diagramlist, title, success, failed

def main():
    parser = argparse.ArgumentParser(description="Catalog Product Downloader")
    parser.add_argument('-i', '--input', type=str,help="Source File")
    
    args = parser.parse_args()
    isExist = os.path.exists(args.input)
    if isExist == False :
        input('Please check the XLS file')
        sys.exit()

    windp = s.CHROME_DOWNLOAD_PATH
    for f in glob.glob(windp + os.sep + "*"):
        os.remove(f)

    source = args.input
    
    driver = browser_init()
    driver.maximize_window()
    print('Opening the Source Excel File...', end="", flush=True)
    xlbook = xw.Book(source)
    xlsheet1 = xlbook.sheets["Sheet1"]
    xlsheet2 = xlbook.sheets["Sheet2"]
    print('OK')
    maxrow = xlsheet1.range('C' + str(xlsheet1.cells.last_cell.row)).end('up').row
    # breakpoint()
    for i in range(2, maxrow + 1):
        merger = PdfMerger()
        if xlsheet1[f'D{i}'].value == 'NO':
            diagramlist, title, success, failed = parse(url=xlsheet1[f'C{i}'].value, driver=driver, xlsheet2=xlsheet2)
            if len(diagramlist) > 0:
                for diagram in diagramlist:
                    filename = str(diagram[4]).split(",")[1].replace('"',"").replace(")","")
                    merger.append(s.PDF_EXTRACT_PATH + os.sep + filename)
                merger.write(s.PDF_JOIN_PATH + os.sep + slugify(title) + ".pdf")
                link = '=HYPERLINK(CONCATENATE(Sheet3!$A$2,"{}"),"OPEN PDF")'.format(slugify(title) + ".pdf")
                xlsheet1[f'D{i}'].value = 'YES'
                xlsheet1[f'E{i}'].value = link
                xlsheet1[f'F{i}'].value = f"PDF Download Success={success}, Failed={failed}"
            else:
                xlsheet1[f'D{i}'].value = 'YES'
                xlsheet1[f'E{i}'].value = "FAILED"
                xlsheet1[f'F{i}'].value = "No PDF can not be Downloaded"
    driver.quit()


    input("End Process..")    


if __name__ == '__main__':
    main()
