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
APPEARANCES = []
PDF_NAMES = []


#Check a valid csv file exists in current folder
def csv_exist():
    print('Checking csv file exists in folder...')
    logging.info('Checking csv file exists in folder...')
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
    print('Checking template exists in folder...')
    logging.info('Checking template exists in folder...')
    try:
        if not os.path.exists(TEMPLATE):
            raise FileNotFoundError(f'The file "{TEMPLATE}" does not exist.')
    except FileNotFoundError as e:
        print(e)
        logging.error(e)
        raise

#Creates folders to hold pdfs and other raw files
def create_folder(folder_path):
    print('Creating folders...')
    logging.info('Creating folders...')
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

#Checks list of available csvs, then reads and appends cmp code to list. If there are invalid codes, they are added to a new csv names accordingly.
def get_CMP_codes():
    print('Reading CMP codes from csv file...')
    logging.info('Reading CMP codes from csv file...')
    for csv_file in CSV_FILE_EXIST:
        try:
            with open(csv_file, 'r') as file: 
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
            print(f"Error reading {csv_file}: {e}")
            logging.error(f"Error reading {csv_file}: {e}")
            raise
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            logging.error(f"An unexpected error occurred: {e}")
            raise
        
def get_phys_appearance():
    print('Reading physical appearances from csv file...')
    logging.info('Reading physical appearances from csv file...')
    for csv_file in CSV_FILE_EXIST:
        try:
            with open(csv_file, 'r') as file: 
                reader = csv.DictReader(file)
                for col in reader:
                    try:
                        if col['FORMATTED_BATCH_ID'] in CMP_CODES:
                            APPEARANCES.append(col['COLOUR'] + col['FORM'])
                    except KeyError as e:
                        print(f"Error: Missing expected column in the CSV file - {e}")
                        logging.error(f"Missing expected column in the CSV file - {e}")
                        raise
        except FileNotFoundError as e:
            print(e)
            logging.error(e)
            raise
        except IOError as e:
            print(f"Error reading {csv_file}: {e}")
            logging.error(f"Error reading {csv_file}: {e}")
            raise
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            logging.error(f"An unexpected error occurred: {e}")
            raise

csv_exist()
template_exist()
print(CSV_FILE_EXIST)
create_folder('MSDS pdfs')
create_folder('MSDS raw files') 
get_CMP_codes()
get_phys_appearance()