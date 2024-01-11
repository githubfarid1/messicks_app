import sys
import os
current = os.path.dirname(os.path.realpath(__file__))
parent = os.path.dirname(current)
sys.path.append(parent)
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from slugify import slugify
import json
import  pyexcel_ods3 as pods
from collections import OrderedDict
import settings as s
import glob
import shutil
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager as CM
from selenium.webdriver.support.select import Select
import time
import random
from html import unescape
import unicodedata

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


data = pods.get_data(afile=s.ODF_RESULT_PATH + os.sep + "resulturls.ods")
sheet1 = data['Sheet1'].copy()
sheet2 = data['Sheet2'].copy()
sheet3 = data['Sheet3'].copy()
driver = browser_init()
driver.maximize_window()
diagramlist = []
diagramlist.append(["VENDOR", "MODEL NAME", "URL", "ISDOWNLOAD", "LINK"])
for idx, ds in enumerate(sheet1):
    if idx == 0:
        continue
    # if idx == 10:
    #     break
    title = f"{ds[0]}-{ds[1]}-parts"
    filename = slugify(title) + ".pdf"
    if os.path.exists(s.PDF_JOIN_PATH + os.sep + filename):
        link = '=HYPERLINK(CONCATENATE($Sheet3.$A$2,"{}"),"OPEN PDF")'.format(filename)
        sheet1[idx] = [ds[0], ds[1], ds[2], 'YES', link]
        url = ds[2]
        driver.get(url)
        # breakpoint()
        time.sleep(2)
        firstopt1txt = ''
        firstopt2txt = ''
        first = True
        title = driver.find_element(By.CSS_SELECTOR, "h1#model-title").text
        vendor = driver.find_element(By.XPATH, "/html/body/div[6]/button").get_attribute('data-brand')
        
        
        while True:
            driver.find_element(By.CSS_SELECTOR, "span.ms-5 span.btn-prev-diagram").click()
            WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.cc-part-tile")))
            time.sleep(2)
            if first:
                first = False
                firstopt1txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select ms-3 me-3 section-list']")).first_selected_option.text
                firstopt2txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select diagram-list']")).first_selected_option.text
                section, diagram  = firstopt1txt, firstopt2txt
                filename = slugify("{}{}{}".format(title, section, diagram) )+".pdf"
                # if os.path.exists(s.PDF_EXTRACT_PATH + os.sep + filename):
                #     continue

                # driver.find_element(By.CSS_SELECTOR, "span.print-pdf").click()
                # time.sleep(3)
                # time.sleep(random.randint(5, 7))
                # print("download for", title, section, diagram)
                # section, diagram  = firstopt1txt, firstopt2txt
                # time.sleep(random.randint(2, 5))
                # if  len(driver.window_handles) > 1:
                #     for handle in driver.window_handles:
                #         if handle != curr:
                #             driver.switch_to.window(handle)
                #             time.sleep(2)
                #             try:
                #                 driver.close()
                #             except:
                #                 pass
                #             if len(driver.window_handles) == 1:
                #                 break
                #     driver.switch_to.window(curr)
                # else:
                #     # filename = slugify("{}{}{}".format(title, section, diagram) )+".pdf"
                #     latestfile = getlatestfile(s.CHROME_DOWNLOAD_PATH)
                #     shutil.copyfile(latestfile, s.PDF_EXTRACT_PATH + os.sep + filename)
                link = '=HYPERLINK(CONCATENATE($Sheet3.$A$1,"{}"),"OPEN PDF")'.format(filename)
                if os.path.exists(s.PDF_EXTRACT_PATH + os.sep + filename):
                    diagramlist.append([vendor, title, unicodedata.normalize('NFKC',unescape(section)), unicodedata.normalize('NFKC',unescape(diagram)), link])
                continue

            opt1txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select ms-3 me-3 section-list']")).first_selected_option.text
            opt2txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select diagram-list']")).first_selected_option.text

            if opt1txt == firstopt1txt and opt2txt == firstopt2txt:
                break
            # breakpoint()
            section, diagram  = opt1txt, opt2txt
            filename = slugify("{}{}{}".format(title, section, diagram) )+".pdf"
            # if os.path.exists(s.PDF_EXTRACT_PATH + os.sep + filename):
            #     continue
            
            # driver.find_element(By.CSS_SELECTOR, "span.print-pdf").click()
            # time.sleep(random.randint(5, 7))

            # print("download for", title, opt1txt, opt2txt)
            section, diagram  = opt1txt, opt2txt
            # time.sleep(random.randint(2, 5))
            # if  len(driver.window_handles) > 1:
            #     for handle in driver.window_handles:
            #         if handle != curr:
            #             driver.switch_to.window(handle)
            #             time.sleep(2)
            #             try:
            #                 driver.close()
            #             except:
            #                 pass
            #             if len(driver.window_handles) == 1:
            #                 break
            #     driver.switch_to.window(curr)
            # else:
            filename = slugify("{}{}{}".format(title, section, diagram) )+".pdf"
            # latestfile = getlatestfile(s.CHROME_DOWNLOAD_PATH)
            # shutil.copyfile(latestfile, s.PDF_EXTRACT_PATH + os.sep + filename)
            link = '=HYPERLINK(CONCATENATE($Sheet3.$A$1,"{}"),"OPEN PDF")'.format(filename)
            if os.path.exists(s.PDF_EXTRACT_PATH + os.sep + filename):
                diagramlist.append([vendor, title, unicodedata.normalize('NFKC',unescape(section)), unicodedata.normalize('NFKC',unescape(diagram)), link])

        # driver.quit()
        # return diagramlist, title
data = OrderedDict()
data.update({"Sheet1": sheet1})
data.update({"Sheet2": diagramlist})
data.update({"Sheet3": sheet3})
pods.save_data(s.ODF_RESULT_PATH + os.sep + "resulturlsx.ods", data)
driver.quit()
        


        # for f in glob.glob(s.PDF_EXTRACT_PATH + os.sep + filename.replace(".pdf", "") + "*.pdf"):
        #     print(f)
        
        # input(filename)


