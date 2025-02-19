from flask import Flask, render_template, request, send_file
import pandas as pd
from openpyxl import load_workbook
from docx import Document
import os
import re

app = Flask(__name__)

# Function to normalize WERS codes
def normalize_code(code):
    if code is None:
        return ""
    code = re.sub(r'[-_#]', ' ', code)  # Replace -, _, and # with spaces
    code = re.sub(r'\s*--\s*', ' ', code)  # Replace -- with a space
    code = re.sub(r'\s+', ' ', code)  # Replace multiple spaces with a single space
    return code.strip()

# Function to determine if a row corresponds to a single entry
def is_single_entry(row):
    wers_code = str(row['WERS Code'])
    # Check if the code contains only alphanumeric characters and is not empty
    return bool(wers_code and re.match(r'^[A-Za-z0-9]+$', wers_code))

@app.route('/')
def upload_files():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def process_files():
    uploads_dir = 'uploads'
    # Ensure the uploads directory exists
    if not os.path.exists(uploads_dir):
        os.makedirs(uploads_dir)

    if 'excel_file' not in request.files or 'word_file' not in request.files or 'voci_excel_file' not in request.files:
        return "Missing file(s)"

    # Save uploaded files
    excel_file = request.files['excel_file']
    word_file = request.files['word_file']
    voci_excel_file = request.files['voci_excel_file']
    excel_header = int(request.form['excel_header'])
    voci_header = int(request.form['voci_header'])

    excel_path = os.path.join(uploads_dir, excel_file.filename)
    word_path = os.path.join(uploads_dir, word_file.filename)
    voci_excel_path = os.path.join(uploads_dir, voci_excel_file.filename)

    excel_file.save(excel_path)
    word_file.save(word_path)
    voci_excel_file.save(voci_excel_path)

    try:
        excel_data = pd.read_excel(excel_path, header=excel_header - 1)
    except Exception as e:
        return f"Error reading Excel file: {e}"

    column_name = 'Feature WERS Code'
    if column_name not in excel_data.columns:
        return f"Column '{column_name}' not found in the Excel file."

    codes_from_excel = excel_data[column_name].dropna().astype(str).tolist()

    try:
        doc = Document(word_path)
    except Exception as e:
        return f"Error reading Word file: {e}"

    text_content = [paragraph.text for paragraph in doc.paragraphs]
    full_text = ' '.join(text_content)

    # Normalize WERS codes found in Excel
    normalized_codes_from_excel = [normalize_code(code) for code in codes_from_excel]

    # Find codes in Word document
    codes_found_in_word = []
    for code in codes_from_excel:
        normalized_code = normalize_code(code)
        if code in full_text or normalized_code in full_text:
            codes_found_in_word.append(code)

    try:
        voci_data = pd.read_excel(voci_excel_path, header=voci_header - 1)
    except Exception as e:
        return f"Error reading VOCI Excel file: {e}"

    required_columns = ['WERS Code', 'Sales Code']
    for col in required_columns:
        if col not in voci_data.columns:
            return f"Column '{col}' not found in the VOCI Excel file."

    voci_codes = voci_data[['WERS Code', 'Sales Code']].dropna().astype(str)

    # Create a mapping from WERS Code to Sales Code
    single_code_sales = {}
    group_code_sales = {}

    # First pass: identify all single entries
    for index, row in voci_codes.iterrows():
        wers_code = normalize_code(row['WERS Code'])
        sales_code = row['Sales Code']
        
        if wers_code and is_single_entry(row):
            single_code_sales[wers_code] = sales_code

    # Second pass: add group entries only if no single entry exists
    for index, row in voci_codes.iterrows():
        wers_code = normalize_code(row['WERS Code'])
        sales_code = row['Sales Code']
        
        if wers_code and not is_single_entry(row):
            if wers_code not in single_code_sales:  # Only add if no single entry exists
                group_code_sales[wers_code] = sales_code

    results = []
    for code in codes_found_in_word:
        normalized_code = normalize_code(code)
        
        # Determine sales code, prioritizing single sales codes
        chosen_sales_code = single_code_sales.get(normalized_code, group_code_sales.get(normalized_code, code))
        results.append((code, chosen_sales_code))

    # Load the existing Excel workbook
    wb = load_workbook(excel_path)
    ws = wb.active

    # Create a dictionary for results
    result_dict = {normalize_code(code): sales_code for code, sales_code in results}

    # Update Sales Code column (Column E)
    sales_code_col = 5  # Column E is the 5th column (1-based index)
    for row in range(2, ws.max_row + 1):  # Assuming row 1 is header
        wers_code_cell = ws.cell(row=row, column=3)  # Assuming WERS Code is in Column C (3rd column)
        sales_code_cell = ws.cell(row=row, column=sales_code_col)
        wers_code = wers_code_cell.value
        if wers_code in result_dict:
            sales_code_cell.value = result_dict[wers_code]

    # Save the updated Excel file
    updated_excel_filename = 'updated_' + excel_file.filename
    updated_excel_path = os.path.join(uploads_dir, updated_excel_filename)
    wb.save(updated_excel_path)

    # Debugging: Verify the saved path
    print(f"Updated Excel path: {updated_excel_path}")  # Debug print

    return render_template('results.html', results=results, updated_excel_path=updated_excel_filename)

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join('uploads', filename)
    print(f"Attempting to download: {file_path}")  # Debugging
    if not os.path.exists(file_path):
        return f"File not found: {file_path}", 404
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
