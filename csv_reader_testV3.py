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

CSV_FILE_CHECK = ['shipping_export.csv', 'invalid_codes.csv']
CSV_FILE_EXIST = []
TEMPLATE = 'MSDS Template.docx'
CMP_CODES = []
INVALID_CODES = []
PDF_NAMES = []


#Check a valid csv file exists in current folder
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

#Check a valid template exists in current folder
def template_exist():
    try:
        if not os.path.exists(TEMPLATE):
            raise FileNotFoundError(f'The file "{TEMPLATE}" does not exist.')
    except FileNotFoundError as e:
        print(e)
        logging.error(e)
        raise

#Creates folders to hold pdfs and other raw files
def create_folder(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

def get_CMP_codes():
    global CMP_FILE_EXIST
    for csv in CMP_FILE_EXIST:
        try:
            with open(csv, 'r') as file: 
                reader = csv.DictReader(file)
                for col in reader:
                    try:
                        if re.search('CMP-.{8}-.{3}', col['FORMATTED_BATCH_ID']):
                            CMP_CODES.append(col['FORMATTED_BATCH_ID']) 
                    
                        else:
                            INVALID_CODES.append(col['FORMATTED_BATCH_ID'])
                    except KeyError as e:
                        print(f"Error: Missing expected column in the CSV file - {e}")
                        logging.error(f"Missing expected column in the CSV file - {e}")
                        raise
                        
            try:
                if len(INVALID_CODES) != 0:
                    with open('invalid_codes.csv', 'w') as f:
                            writer = csv.DictWriter(f, fieldnames=["FORMATTED_BATCH_ID"])
                            writer.writeheader()
                            writer.writerows({"FORMATTED_BATCH_ID":i} for i in INVALID_CODES)
                    print(f'The following codes were invalid, so have not been processed: {INVALID_CODES}\nThese have been added to a new csv file "invalid_codes.csv" for you to correct and re-run.')  
                    logging.warning(f'The following codes were invalid, so have not been processed: {INVALID_CODES}\nThese have been added to a new csv file "invalid_codes.csv" for you to correct and re-run.')
                    

            except IOError as e:
                print(f"Error writing to invalid_codes.csv: {e}")
                logging.error(f"Error writing to invalid_codes.csv: {e}")
                raise

        except FileNotFoundError as e:
            print(e)
            logging.error(e)
            raise
        except IOError as e:
            print(f"Error reading {csv}: {e}")
            logging.error(f"Error reading {csv}: {e}")
            raise
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            logging.error(f"An unexpected error occurred: {e}")
            raise
        
def get_phys_appearance():
    pass

csv_exist()
template_exist()
print(CSV_FILE_EXIST)
create_folder('MSDS pdfs')
create_folder('MSDS raw files') 
get_CMP_codes()
get_phys_appearance()