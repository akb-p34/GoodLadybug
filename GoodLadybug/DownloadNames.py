import fitz  # PyMuPDF
import openpyxl
from openpyxl import load_workbook
import re

# Ask for user input for file paths and sheet name
sheet_name = input("Enter the sheet name: ")
pdf_path = input("Enter the path to the PDF file: ")
excel_path = input("Enter the path to the Excel file: ")

# Open the PDF file
doc = fitz.open(pdf_path)

# Extract text from each page
lines = []
for page in doc:
    text = page.get_text()
    lines.extend(text.split('\n'))  # Split text into lines and add to list

# Close the PDF document
doc.close()

# Open the Excel workbook and select the sheet
wb = load_workbook(excel_path)
sheet = wb[sheet_name]

# Starting row in the Excel sheet
row_num = 10

# Process the lines to extract lot numbers and descriptions
for i, line in enumerate(lines):
    if line.startswith('Lot -'):
        lot_number_match = re.search(r'Lot - (\S+)', line)
        if lot_number_match:
            lot_number = lot_number_match.group(1)
            # The actual description is expected two lines below the 'Lot - XXX' line.
            # This accounts for the lot ID possibly repeating on the line immediately after.
            if i+2 < len(lines):
                description = lines[i+2].strip()
            else:
                description = "Description not found"
            print(f"Recognized Lot Number: {lot_number}")
            print(f"Recognized Description: {description}")
            print(f"Writing to Excel: Lot Number - Lot - {lot_number}, Description - {description}")
            
            sheet[f'B{row_num}'] = description
            sheet[f'K{row_num}'] = f"Lot - {lot_number}"
            row_num += 1

# Save the workbook
wb.save(excel_path)

# Provide feedback that the operation is complete
print(f"Items and lot identifiers have been successfully added to '{sheet_name}' in '{excel_path}'.")
