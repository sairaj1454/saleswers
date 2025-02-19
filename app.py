from flask import Flask, render_template, request, send_file
import pandas as pd
from openpyxl import load_workbook
from docx import Document
import os
import re

# Initialize Flask app
app = Flask(__name__)

# Ensure uploads directory exists
UPLOADS_DIR = 'uploads'
os.makedirs(UPLOADS_DIR, exist_ok=True)

# Function to normalize WERS codes
def normalize_code(code):
    if code is None:
        return ""
    code = re.sub(r'[-_#]', ' ', code)
    code = re.sub(r'\s*--\s*', ' ', code)
    code = re.sub(r'\s+', ' ', code)
    return code.strip()

# Function to determine if a row corresponds to a single entry
def is_single_entry(row):
    wers_code = str(row['WERS Code'])
    return bool(wers_code and re.match(r'^[A-Za-z0-9]+$', wers_code))

# Function to extract the alphanumeric code from the end of Feature WERS Description
def extract_end_code(description):
    if not isinstance(description, str):
        return None
    parts = re.split(r'[-\s]+', description.strip())
    return parts[-1] if parts else None

@app.route('/')
def upload_files():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def process_files():
    try:
        if 'excel_file' not in request.files or 'word_file' not in request.files or 'voci_excel_file' not in request.files:
            return "Missing file(s)", 400
        
        excel_file = request.files['excel_file']
        word_file = request.files['word_file']
        voci_excel_file = request.files['voci_excel_file']
        excel_header = int(request.form['excel_header'])
        voci_header = int(request.form['voci_header'])

        excel_path = os.path.join(UPLOADS_DIR, excel_file.filename)
        word_path = os.path.join(UPLOADS_DIR, word_file.filename)
        voci_excel_path = os.path.join(UPLOADS_DIR, voci_excel_file.filename)

        excel_file.save(excel_path)
        word_file.save(word_path)
        voci_excel_file.save(voci_excel_path)

        excel_data = pd.read_excel(excel_path, header=excel_header - 1)
        voci_data = pd.read_excel(voci_excel_path, header=voci_header - 1)

        required_excel_columns = ['Feature WERS Code', 'Feature WERS Description', 'Top Family WERS Code']
        for col in required_excel_columns:
            if col not in excel_data.columns:
                return f"Column '{col}' not found in the Excel file.", 400

        required_voci_columns = ['WERS Code', 'Sales Code']
        for col in required_voci_columns:
            if col not in voci_data.columns:
                return f"Column '{col}' not found in the VOCI Excel file.", 400

        # Extract mappings
        feature_description_map = {}
        top_family_map = {}

        for _, row in excel_data.iterrows():
            feature_code = str(row['Feature WERS Code']).strip() if pd.notna(row['Feature WERS Code']) else None
            top_family = str(row['Top Family WERS Code']).strip() if pd.notna(row['Top Family WERS Code']) else None
            description = str(row['Feature WERS Description']).strip() if pd.notna(row['Feature WERS Description']) else None

            if feature_code and description and top_family == 'YZA':
                top_family_map[feature_code] = top_family
                end_code = extract_end_code(description)
                if end_code:
                    feature_description_map[feature_code] = end_code

        single_code_sales = {}
        for _, row in voci_data.iterrows():
            wers_code = normalize_code(row['WERS Code'])
            sales_code = row['Sales Code']

            if wers_code and is_single_entry(row):
                if wers_code in top_family_map and wers_code in feature_description_map:
                    sales_code = feature_description_map[wers_code]
                single_code_sales[wers_code] = sales_code

        try:
            doc = Document(word_path)
        except Exception as e:
            return f"Error reading Word file: {e}", 500

        text_content = ' '.join(paragraph.text for paragraph in doc.paragraphs)
        codes_found_in_word = [code for code in excel_data['Feature WERS Code'].dropna().astype(str).tolist() if code in text_content]

        wb = load_workbook(excel_path)
        ws = wb.active
        sales_code_col = 5
        
        for row in range(2, ws.max_row + 1):
            wers_code_cell = ws.cell(row=row, column=3)
            sales_code_cell = ws.cell(row=row, column=sales_code_col)
            wers_code = wers_code_cell.value
            if wers_code in single_code_sales:
                sales_code_cell.value = single_code_sales[wers_code]

        updated_excel_filename = f"updated_{excel_file.filename}"
        updated_excel_path = os.path.join(UPLOADS_DIR, updated_excel_filename)
        wb.save(updated_excel_path)

        return render_template('results.html', updated_excel_path=updated_excel_filename)
    except Exception as e:
        return f"Internal Server Error: {e}", 500

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(UPLOADS_DIR, filename)
    if not os.path.exists(file_path):
        return "File not found", 404
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
