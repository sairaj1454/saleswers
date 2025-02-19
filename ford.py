import pandas as pd
from docx import Document

# Paths to files
excel_file = r'C:\Users\sairaj.thikkisetty\Downloads\PROD_SALES_WANAB-CHD-2025-BroncoSport_20240625102720.xlsx'
word_file = r'C:\Users\sairaj.thikkisetty\Documents\test1.docx'
voci_excel_file = r'C:\Users\sairaj.thikkisetty\Documents\voci.xlsx'

# Load Excel data with header row specified (assuming actual headers are on the 7th row, index 6)
try:
    excel_data = pd.read_excel(excel_file, header=6)
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

# Ensure the column name matches
column_name = 'Feature WERS Code'
if column_name not in excel_data.columns:
    print(f"Column '{column_name}' not found in the Excel file.")
    exit()

# Extract codes from 'Feature WERS Code' column
codes_from_excel = excel_data[column_name].dropna().astype(str).tolist()

# Load Word document
try:
    doc = Document(word_file)
except Exception as e:
    print(f"Error reading Word file: {e}")
    exit()

# Extract text from Word document
text_content = []
for paragraph in doc.paragraphs:
    text_content.append(paragraph.text)

# Join all text into one string for searching
full_text = ' '.join(text_content)

# Print all content from Word document
print("Content from Word document:")
print(full_text)

# Find codes in Word document
codes_found_in_word = []
for code in codes_from_excel:
    if code in full_text:
        codes_found_in_word.append(code)

# Display found codes
print("Codes found in Word document:")
for code in codes_found_in_word:
    print(code)

# Load VOCI Excel data with header row specified (assuming actual headers are on the 11th row, index 10)
try:
    voci_data = pd.read_excel(voci_excel_file, header=11)
except Exception as e:
    print(f"Error reading VOCI Excel file: {e}")
    exit()

# Ensure the required columns are in the VOCI Excel file
required_columns = ['WERS Code', 'Sales Code']
for col in required_columns:
    if col not in voci_data.columns:
        print(f"Column '{col}' not found in the VOCI Excel file.")
        exit()

# Extract WERS Code and Sales Code columns
voci_codes = voci_data[['WERS Code', 'Sales Code']].dropna().astype(str)

# Compare and print matching Sales Codes
print("Matching Sales Codes for found WERS Codes:")
for code in codes_found_in_word:
    matching_sales_codes = voci_codes[voci_codes['WERS Code'] == code]['Sales Code'].tolist()
    if matching_sales_codes:
        for sales_code in matching_sales_codes:
            print(f"WERS Code: {code}, Sales Code: {sales_code}")
    else:
        print(f"WERS Code: {code}, Sales Code: {code} (No match found, using WERS code as Sales code)")

