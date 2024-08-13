import csv
import re
import logging
import sys
import os
import shutil
from docxtpl import DocxTemplate
from docx2pdf import convert

logger = logging.getLogger(__name__)
logging.basicConfig(filename="process_info.log", format='%(levelname)s:%(message)s', encoding='utf-8', level=logging.DEBUG)

CSV_FILE_CHECK = ['shipping_export.csv']
CSV_FILE_EXIST = []
TEMPLATE = 'MSDS Template.docx'
CMP_CODES = {}
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
    print(f'Creating folder "{folder_path}"...')
    logging.info(f'Creating folder "{folder_path}"...')
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
                        if col['FORMATTED_BATCH_ID'] not in CMP_CODES.items():
                            if re.match(r'^CMP-\d{8}-\d{3}$', col['FORMATTED_BATCH_ID']):
                                CMP_CODES.update({col['FORMATTED_BATCH_ID']:{'FORMATTED_BATCH_ID':col['FORMATTED_BATCH_ID'], 'COLOUR':col['COLOUR'], 'FORM_':col['FORM_']}})
                            
                            else:
                                INVALID_CODES.append(col['FORMATTED_BATCH_ID'])
                               
                        
                    except KeyError as e:
                        print(f"Error: Missing expected column in the CSV file - {e}")
                        logging.error(f"Missing expected column in the CSV file - {e}")
                        raise
                     
            try:
                if len(INVALID_CODES) != 0:
                    with open('invalid_codes.csv', 'w') as target, open('shipping_export.csv', 'r') as source:
                        reader = csv.DictReader(source)
                        writer = csv.DictWriter(target, fieldnames = reader.fieldnames)
                        writer.writeheader()
                        for col in reader:
                            if col['FORMATTED_BATCH_ID'] in INVALID_CODES:
                                writer.writerow(col)
                
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
# Reads csv file and adds appearance of compound to a list. Doubles as a check that there are no empty values.   
def get_phys_appearance():
    print('Reading physical appearances from csv file...')
    logging.info('Reading physical appearances from csv file...')
    for csv_file in CSV_FILE_EXIST:
        try:
            with open(csv_file, 'r') as file: 
                reader = csv.DictReader(file)
                for col in reader:
                    try:
                        if col['FORMATTED_BATCH_ID'] in CMP_CODES.items():
                            if not col['COLOUR'] or not col['FORM_']:
                                print(f"There are blank cells in the shipping_export file for {col['FORMATTED_BATCH_ID']}")
                                logging.error(f"There are blank cells in the shipping_export file for {col['FORMATTED_BATCH_ID']}")
                                raise KeyError
                                
                            else:
                                APPEARANCES.append(col['COLOUR'].title() + ' ' + col['FORM_'].title())
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

def make_doc():
    try:
        if CMP_CODES:
            print('Inserting data into template...')
            logging.info('Inserting data into template...')
        try:
            doc = DocxTemplate(TEMPLATE)
            for code, properties in CMP_CODES.items():
                doc.render(properties)
                doc.save(code +' MSDS'+'.docx')
                                
        except Exception as e:
            print('The data could not be inserted into the template.')
            logging.error(e)
            raise
        for i in CMP_CODES:
            PDF_NAMES.append(i +' MSDS'+'.docx')
    except Exception as e:
        print('The data could not be added to the template. Make sure that the template contains the phrase "{{FORMATTED_BATCH_ID}}", "{{COLOUR}}" and "{{FORM}}" before re-running.')
        logging.error(e)
        raise

#Converts docx to pdf file.        
def convert_to_pdf():
    print('Converting docx to pdf...')
    logging.info('Converting docx to pdf...')
    try:
        for i in PDF_NAMES:
            convert(i)
    except Exception as e:
        logging.error(e)
        raise

#Moves pdf files to folder, and all other raw files to separate folder. This reduces problems if any of the csv files need to be re-run.
def move_files():
    print('Attempting to move files to appropriate folders...')
    logging.info('Attempting to move files to appropriate folders...')
    source_dir = os.getcwd()
    pdf_dir = 'MSDS pdfs'
    raw_dir = 'MSDS raw files'
    files_to_move = os.listdir(source_dir)
  
    try:   
        for file in files_to_move:
            try:
                if file.endswith('pdf'):
                    source_path = os.path.join(source_dir, file)
                    pdf_path = os.path.join(pdf_dir, file)
                    shutil.move(source_path, pdf_path)
                
            except FileNotFoundError:
                logging.error('There are no pdf files in the current folder.')
                raise

        for file in files_to_move:
            try:
                if file != 'MSDS Template.docx' and file.endswith('docx'):
                    source_path = os.path.join(source_dir, file)
                    raw_path = os.path.join(raw_dir, file)
                    shutil.move(source_path, raw_path)
                elif file == 'shipping_export.csv': 
                    source_path = os.path.join(source_dir, file)
                    raw_path = os.path.join(raw_dir, file)
                    shutil.move(source_path, raw_path)   
                                
                        
            except FileNotFoundError:
                logging.error('Documents could not be moved into MSDS raw files folder')
                raise    
    except Exception as e:
        logging.error(e)
        raise

def exit():
        input('Press ENTER to exit\n')
        sys.exit()

csv_exist()
template_exist()
create_folder('MSDS pdfs')
create_folder('MSDS raw files') 
get_CMP_codes()
get_phys_appearance()
make_doc()
convert_to_pdf()
print(CMP_CODES.items())
move_files()
logging.shutdown()
print('Process complete.')

exit()
