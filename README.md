# MSDS Template Filler - a project built for my job role, which has the potential to be modified for more general usage. 

# An improved version is currently being worked on, where OOP is utilised.

This is an automation script to extract specified information about chemical compounds from a CSV and insert them into the corresponding fields of a generic MSDS template. For each row of data, an individual document is created as a docx file and then saved as a pdf. All generated documents are then organised into generated folders, depending on their file type and function.

The code requires the CSV file to have a specific name to read from, either 'shipping_export' or 'invalid_codes'. The 'invalid_codes' CSV file is generated as part of the error handling process, where any codes that do not meet the required format are moved into a new file for the user to correct and re-run without the programme failing. 

The fields the programme specifically inserts into the template are 'FORMATTED_BATCH_ID', 'COLOR', and 'FORM_'. In the absence of correctly formatted data in these fields, the programme will not run.

💾 The entire process is logged in 'process_info.log' to track any errors and when these might have occurred.

📂 The script should be run from the directory in which the shipping_export CSV exists and the generated files will be created.

📝 The requirements file can be found in this repository.
