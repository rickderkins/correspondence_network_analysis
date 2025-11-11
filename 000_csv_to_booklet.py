import pandas as pd
from docx import Document
from docx.shared import Inches
import os

# --- Configuration ---
CSV_FILE = r'D:\heiBOX\Seafile\Masterarbeit_Ablage\hwsk_korr.csv'
OUTPUT_FILE = r'D:\heiBOX\Seafile\Masterarbeit_Ablage\All_Letters_One_Document.docx'

# Define a clean mapping for the extremely long original column headers
COLUMN_MAPPING = {
    'http://purl.org/dc/elements/1.1/title': 'ID',
    'http://purl.org/dc/elements/1.1/creator': 'Creator',
    'http://purl.org/dc/terms/audience': 'Recipient',
    'http://purl.org/dc/elements/1.1/date': 'Date',
    'http://purl.org/dc/elements/1.1/coverage': 'Location_From',
    'http://purl.org/dc/terms/spatial': 'Location_To',
    'http://purl.org/dc/elements/1.1/description': 'Description',
    'https://tropy.org/v1/tropy#text': 'Text_Full',
    'https://tropy.org/v1/tropy#tag': 'Tags',
    'https://tropy.org/v1/tropy#note': 'Note_Partial_Text'
    # Include all other columns you want to see here if not listed
}

# 1. Load the CSV file
try:
    df = pd.read_csv(CSV_FILE)
    # Rename columns to the clean, short versions
    df.rename(columns=COLUMN_MAPPING, inplace=True)
except FileNotFoundError:
    print(f"Error: CSV file '{CSV_FILE}' not found.")
    exit()

# 2. Initialize the Word Document
document = Document()
print(f"Loaded {len(df)} records. Starting document generation...")

# 3. Iterate over each row and create a page
for index, row in df.iterrows():
    # --- Start of the Letter/Page Content ---

    # 3.1. Add a title/heading for easy identification
    letter_heading = f"Letter {index + 1}: {row['Creator']} to {row['Recipient']} ({row['Date']})"
    document.add_heading(letter_heading, level=1)

    # 3.2. Add key metadata
    document.add_paragraph(f"Document ID: {row.get('ID', 'N/A')}")
    document.add_paragraph(f"Subject: {row.get('Description', 'N/A')}")
    document.add_paragraph(f"Tags: {row.get('Tags', 'None')}")
    document.add_paragraph(f"Sent From: {row.get('Location_From', 'N/A')} | To: {row.get('Location_To', 'N/A')}")
    document.add_paragraph("")  # Blank line separator

    # 3.3. Add the main letter text
    document.add_heading("Full Transcript:", level=3)
    main_text = str(row.get('Text_Full', row.get('Note_Partial_Text', 'NO TEXT FOUND FOR THIS ENTRY')))
    document.add_paragraph(main_text)

    # --- End of the Letter/Page Content ---

    # 3.4. Insert a page break (for all but the last record)
    if index < len(df) - 1:
        document.add_page_break()

# 4. Save the document
document.save(OUTPUT_FILE)
print(f"\n--- SUCCESS ---")
print(f"The document '{OUTPUT_FILE}' was created successfully.")
print(f"It contains {len(df)} letters, with one letter starting on each page.")