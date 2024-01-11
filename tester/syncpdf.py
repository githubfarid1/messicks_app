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
        breakpoint()
        for f in glob.glob(s.PDF_EXTRACT_PATH + os.sep + filename.replace(".pdf", "") + "*.pdf"):
            print(f)
        
        input(filename)


