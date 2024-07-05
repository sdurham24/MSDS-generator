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
CSV_NAMES = ['shipping_export.csv', 'invalid_codes.csv']
TEMPLATE = 'MSDS Template.docx'

def csv_exist():
    for name in CSV_NAMES:
        try:
            if not os.path.exists(name):
                raise FileNotFoundError(f'The file "{name}" does not exist.')
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

csv_exist()
template_exist()