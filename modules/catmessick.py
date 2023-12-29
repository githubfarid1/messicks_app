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

def parse(url):
    windp = s.CHROME_DOWNLOAD_PATH
    for f in glob.glob(windp + os.sep + "pdf*.pdf"):
        os.remove(f)

    driver = browser_init()
    driver.maximize_window()
    o = urlparse(url)
    resfilename = "{}.pdf".format(o.path.split("/")[-1])
    driver.get(url)
    firstopt1txt = ''
    firstopt2txt = ''
    first = True
    merger = PdfMerger()
    time.sleep(2)
    curr=driver.current_window_handle
    # breakpoint()
    while True:
        driver.find_element(By.CSS_SELECTOR, "span.ms-5 span.btn-prev-diagram").click()
        WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.cc-part-tile")))
        time.sleep(2)
        if first:
            first = False
            firstopt1txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select ms-3 me-3 section-list']")).first_selected_option.text
            firstopt2txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select diagram-list']")).first_selected_option.text
            driver.find_element(By.CSS_SELECTOR, "span.print-pdf").click()
            time.sleep(3)
            time.sleep(random.randint(5, 7))

            print("download for ", firstopt1txt, firstopt2txt)
            time.sleep(random.randint(2, 5))
            continue
        opt1txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select ms-3 me-3 section-list']")).first_selected_option.text
        opt2txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select diagram-list']")).first_selected_option.text

        if opt1txt == firstopt1txt and opt2txt == firstopt2txt:
            break
        # breakpoint()
        
        driver.find_element(By.CSS_SELECTOR, "span.print-pdf").click()
        time.sleep(random.randint(5, 7))

        print("download for ", opt1txt, opt2txt)
        time.sleep(random.randint(2, 5))
        if  len(driver.window_handles) > 1:
            # breakpoint()
            # driver.switch_to.window(curr)
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
    print("PDF files merging ...")
    files = list(filter(os.path.isfile, glob.glob(windp + os.sep + "pdf*.pdf")))
    files.sort(key=lambda x: os.path.getmtime(x))
    for file in files:
        merger.append(file)
    
    merger.write(s.PDF_RESULT_PATH + os.sep + resfilename)


def main():
    parser = argparse.ArgumentParser(description="Catalog Product Downloader")
    parser.add_argument('-url', '--url', type=str,help="URL product")
    args = parser.parse_args()
    parse(url=args.url)
    input("End Process..")    


if __name__ == '__main__':
    main()
