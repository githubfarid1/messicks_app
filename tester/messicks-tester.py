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

windp = str(Path.home() / "Downloads")
for f in glob.glob(windp + os.sep + "pdf*.pdf"):
    os.remove(f)


cud = os.getcwd() + os.sep + "chrome-user-data"
cp = "Profile 1"
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
resfilename = "{}.pdf".format(o.path.split("/")[-1])
driver.get(url)
# breakpoint()
# breakpoint()
firstopt1txt = ''
firstopt2txt = ''
first = True
merger = PdfMerger()
# title = driver.find_element(By.CSS_SELECTOR, "h1.model-title").text
while True:
    driver.find_element(By.CSS_SELECTOR, "span.ms-5 span.btn-prev-diagram").click()
    # time.sleep(random.randint(2, 5))
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.cc-part-tile")))
    time.sleep(2)
    if first:
        first = False
        firstopt1txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select ms-3 me-3 section-list']")).first_selected_option.text
        firstopt2txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select diagram-list']")).first_selected_option.text
        driver.find_element(By.CSS_SELECTOR, "span.print-pdf").click()
        time.sleep(3)
        time.sleep(random.randint(5, 7))
        # if opt1txt == 'USEFUL PRODUCTS':
        #     break

        print("download for ", firstopt1txt, firstopt2txt)
        time.sleep(random.randint(2, 5))
        continue

    # curtitle = driver.find_element(By.CSS_SELECTOR, "h1.model-title").text
    
    # if curtitle != title:
    #     break
    opt1txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select ms-3 me-3 section-list']")).first_selected_option.text
    opt2txt = Select(driver.find_element(By.CSS_SELECTOR, "select[class='form-select diagram-list']")).first_selected_option.text
    if opt1txt == firstopt1txt and opt2txt == firstopt2txt:
        break
    driver.find_element(By.CSS_SELECTOR, "span.print-pdf").click()
    time.sleep(random.randint(5, 7))
    # if opt1txt == 'USEFUL PRODUCTS':
    #     break

    print("download for ", opt1txt, opt2txt)
    time.sleep(random.randint(2, 5))



files = list(filter(os.path.isfile, glob.glob(windp + os.sep + "pdf*.pdf")))
files.sort(key=lambda x: os.path.getmtime(x))
for file in files:
    merger.append(file)
merger.write(resfilename)