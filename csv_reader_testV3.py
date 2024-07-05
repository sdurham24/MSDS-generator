import csv
import re
import logging
import sys
import os
import shutil
# from docxtpl import DocxTemplate
# from docx2pdf import convert

logger = logging.getLogger(__name__)
logging.basicConfig(filename="process_info.log", format='%(levelname)s:%(message)s', encoding='utf-8', level=logging.DEBUG)

CMP_CODES = []
INVALID_CODES = []
PDF_NAMES = []
CSV_FILE_CHECK = ['shipping_export.csv', 'invalid_codes.csv']
CSV_FILE_EXIST = []
TEMPLATE = 'MSDS Template.docx'

def csv_exist():
    try:
        for name in CSV_FILE_CHECK:
            if os.path.exists(name):
                CSV_FILE_EXIST.append(name)
        if not CSV_FILE_EXIST:
            raise FileNotFoundError(f'There are no valid csv files in this folder.')
    except FileNotFoundError as e:
        print(e)
        logging.error(e)
        raise


def template_exist():
    try:
        if not os.path.exists(TEMPLATE):
            raise FileNotFoundError(f'The file "{TEMPLATE}" does not exist.')
    except FileNotFoundError as e:
        print(e)
        logging.error(e)
        raise

def create_folder(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

csv_exist()
template_exist()
print(CSV_FILE_EXIST)
create_folder('MSDS pdfs')
create_folder('MSDS raw files') 