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
data = pods.get_data(afile=s.ODF_RESULT_PATH + os.sep + "resulturls.ods")
sheet1 = data['Sheet1'].copy()
sheet2 = data['Sheet2'].copy()
for idx, ds in enumerate(sheet1):
    if idx == 0:
        continue
    # if idx == 10:
    #     break
    filename = slugify(f"{ds[0]}-{ds[1]}-parts") +".pdf"
    if os.path.exists(s.PDF_JOIN_PATH + os.sep + filename):
        for f in glob.glob(s.PDF_EXTRACT_PATH + os.sep + filename.replace(".pdf", "")):
            print(f)
        
        input(filename)


