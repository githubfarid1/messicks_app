import os, glob
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


cud = os.getcwd() + os.sep + "chrome-user-data"
cp = "Default"
options = webdriver.ChromeOptions()
options.add_argument("user-data-dir={}".format(cud))
options.add_argument("profile-directory={}".format(cp))
options.add_argument('--no-sandbox')
options.add_argument("--log-level=3")
options.add_argument("--window-size=800,600")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
# profile = {"download.default_directory": os.getcwd() + os.path.sep + "pdfs" + os.path.sep, 
#             "download.extensions_to_open": "applications/pdf",
#             "download.prompt_for_download": False,
#             'profile.default_content_setting_values.automatic_downloads': 1,
#             "download.directory_upgrade": True,
#             "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome                    
#             }
# options.add_experimental_option("prefs", profile)

# options.add_experimental_option("prefs", {
#   "download.default_directory": "C:\\tes\\",
#   "download.prompt_for_download": False,
#   "download.directory_upgrade": True,
# })

driver = webdriver.Chrome(service=Service(executable_path=os.path.join(os.getcwd(), "chromedriver", "chromedriver.exe")), options=options)
driver.maximize_window()
url = "https://messicks.com/KU/85674"
o = urlparse(url)
breakpoint()