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
#cmp_file = 'shipping_export.csv'
template = 'MSDS Template.docx'

def create_folder(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

class msds_generator():
    def __init__(self) -> None:
        pass
                
    #Function to convert csv into dictionary, then append the CMP codes to an iterable list. If code doesn't follow correct format, they are added to a new csv and a warning is saved into a logger.
    def get_CMP_codes(cmp_file):
        print('Getting CMP codes from csv file...')
        logging.info('Getting CMP codes from csv file.')
        try:
            if not os.path.exists(cmp_file):
                raise FileNotFoundError(f'The file "{cmp_file}" does not exist.')
                
            with open(cmp_file, 'r') as file: 
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
            print(f"Error reading {cmp_file}: {e}")
            logging.error(f"Error reading {cmp_file}: {e}")
            raise
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            logging.error(f"An unexpected error occurred: {e}")
            raise
    

   #Iterates through CMP_CODES list, and creates word docx file which inserts CMP code to template.
    # def make_doc(cmp_file):
    #     if CMP_CODES:
    #         print('Inserting CMP codes into document...')
    #         logging.info('Inserting CMP codes into document...')
            
    #         try:
    #             if not os.path.exists(cmp_file):
    #                 raise FileNotFoundError(f'The file "{cmp_file}" does not exist.')
                    
                
    #             try:
    #                 with open(cmp_file, 'r') as file: 
    #                     reader = csv.DictReader(file)
    #                     doc = DocxTemplate(template)
    #                     for dict in reader:
    #                       doc.render(dict)
    #                     for code in CMP_CODES:
    #                         doc.save(code +' MSDS'+'.docx')
    #             except Exception as e:
    #                 print('The CMP codes could not be inserted into the template.')
    #                 logging.error(e)
    #                 raise
    
                    
    #             for i in CMP_CODES:
    #                 PDF_NAMES.append(i +' MSDS'+'.docx')
    #         except Exception as e:
    #             print('The codes were not able to be added to the template. Make sure that the template contains the phrase "{{FORMATTED_BATCH_ID}}" before re-running.')
    #             logging.error(e)
    #             raise
            
    #Converts docx to pdf file.        
    # def convert_to_pdf():
    #     print('Converting docx to pdf...')
    #     logging.info('Converting docx to pdf...')
    #     try:
    #         for i in PDF_NAMES:
    #             convert(i)
    #     except Exception as e:
    #         logging.error(e)
    #         raise
        
    #Moves files to new folder
    def move_files():
        print('Attempting to move pdf files to folder "MSDS pdfs"...')
        logging.info('Attempting to move pdf files to folder "MSDS pdfs"...')
        source_dir = os.getcwd()
        pdf_dir = 'MSDS pdfs'
        raw_dir = 'MSDS raw files'
        files_to_move = os.listdir(source_dir)
        #print(files_to_move)
        
        
        for file in files_to_move:
            try:
                if file.endswith('pdf'):
                    source_path = os.path.join(source_dir, file)
                    pdf_path = os.path.join(pdf_dir, file)
                    shutil.move(source_path, pdf_path)

                # else:
                #     raise FileNotFoundError('There are no pdf files in the current folder.')
        
            except FileNotFoundError:
                logging.error('There are no pdf files in the current folder.')
                raise

        for file in files_to_move:
            try:
                if file.endswith('docx'):
                    source_path = os.path.join(source_dir, file)
                    raw_path = os.path.join(raw_dir, file)
                    shutil.move(source_path, raw_path)
                elif file == 'shipping_export.csv': 
                    source_path = os.path.join(source_dir, file)
                    raw_path = os.path.join(raw_dir, file)
                    shutil.move(source_path, raw_path)   
                elif file.endswith('log'):   
                    source_path = os.path.join(source_dir, file)
                    raw_path = os.path.join(raw_dir, file)
                    shutil.move(source_path, raw_path)
                # else:
                #     raise FileNotFoundError
            
            except FileNotFoundError:
                logging.error('Documents could not be moved into MSDS raw files folder')
                raise        
           
                
            
    #Allows user to exit upon completion of script.        
    def exit():
        input('Press ENTER to exit\n')
        logging.info('Process exited.')
        sys.exit()

create_folder('MSDS pdfs')
create_folder('MSDS raw files') 

if os.path.exists('shipping_export.csv'):   
    msds_generator.get_CMP_codes('shipping_export.csv')
else:
    msds_generator.get_CMP_codes('invalid_codes.csv')

# if os.path.exists('shipping_export.csv'):   
#     msds_generator.make_doc('shipping_export.csv')
# else:
#     msds_generator.make_doc('invalid_codes.csv')
#msds_generator.convert_to_pdf()
msds_generator.move_files()

print('Process complete.')
logging.info('Process complete.')
msds_generator.exit()

