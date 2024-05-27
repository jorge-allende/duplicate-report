from flask import Flask, request, send_file, render_template, jsonify
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
PROCESSED_FILE = 'processed_file.xlsx'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(PROCESSED_FOLDER):
    os.makedirs(PROCESSED_FOLDER)

def process_excel(file_path, output_path):
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)
        print("Columns before dropping:", df.columns)

        # Step 1: Highlight and delete columns D, G, I, J, and K (0-indexed as 3, 6, 8, 9, 10)
        df.drop(df.columns[[3, 6, 8, 9, 10]], axis=1, inplace=True)
        print("Columns after dropping:", df.columns)

        # Save to a temporary file to use openpyxl for further processing
        temp_path = "temp.xlsx"
        df.to_excel(temp_path, index=False)

        # Load the workbook with openpyxl
        wb = load_workbook(temp_path)
        ws = wb.active

        # Step 2: Find the column letter for 'Content Checksum'
        col_name = 'Content Checksum'
        col_idx = df.columns.get_loc(col_name) + 1  # Get the column index and adjust for 1-based indexing
        col_letter = get_column_letter(col_idx)
        print(f"Column letter for '{col_name}': {col_letter}")

        # Step 3-8: Apply conditional formatting to the correct column
        fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        rule = FormulaRule(formula=[f"COUNTIF(${col_letter}$1:{col_letter}1, {col_letter}1)>1"], fill=fill)
        ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{ws.max_row}", rule)

        # Step 9-13: Apply filters and sort
        ws.auto_filter.ref = ws.dimensions

        # Save the workbook
        wb.save(temp_path)

        # Reload with pandas to apply further operations
        df = pd.read_excel(temp_path)
        print("Columns before finding duplicates:", df.columns)

        # Step 14-15: Highlight the cells in the correct columns based on the 'Content Checksum'
        # Keep the first occurrence unhighlighted, and highlight subsequent duplicates
        duplicate_counts = df.duplicated(subset=[col_name], keep='first')
        for idx in duplicate_counts[duplicate_counts].index:
            ws[f'D{idx + 2}'].fill = fill  # Highlight Agreement Name column cells
            ws[f'E{idx + 2}'].fill = fill  # Highlight Original File Name column cells

        # Step 16-17: Clear filter and hide the specific columns
        ws.column_dimensions[get_column_letter(df.columns.get_loc('Original File Name') + 1)].hidden = True
        ws.column_dimensions[get_column_letter(df.columns.get_loc('Content Checksum') + 1)].hidden = True

        # Save the workbook
        wb.save(output_path)
        print("Processing complete. Output saved to:", output_path)
    except Exception as e:
        print(f"Error processing the file: {e}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part', 400
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400
    if file:
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        # Process the file here
        processed_path = os.path.join(PROCESSED_FOLDER, PROCESSED_FILE)
        process_excel(filepath, processed_path)
        return jsonify(status='File uploaded and processed'), 200

@app.route('/download')
def download_file():
    processed_file_path = os.path.join(PROCESSED_FOLDER, PROCESSED_FILE)
    print("Downloading file from:", processed_file_path)
    if not os.path.exists(processed_file_path):
        print("File not found:", processed_file_path)
    return send_file(processed_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
