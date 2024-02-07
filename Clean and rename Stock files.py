import os
import pandas as pd
import re
from datetime import datetime

# Define the folder path
folder_path = r'C:\Users\Tom\OneDrive - Motor Depot\Apex\Stock'

# Function to process a single Excel file and create a new cleaned file
def process_excel_file(file_path):
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Remove contents from "Standard Equipment" and "Classified Features" columns
        df[['Standard Equipment', 'Classified Features', 'Notes']] = ''

        # Extract the filename without the extension
        base_filename = os.path.splitext(os.path.basename(file_path))[0]

        # Create a new filename for the cleaned data
        cleaned_file_path = os.path.join(folder_path, f'{base_filename}_cleaned.xlsx')

        # Save the cleaned data to the new Excel file
        df.to_excel(cleaned_file_path, index=False, engine='xlsxwriter')

        print(f"Processed: {file_path} => {cleaned_file_path}")
    except Exception as e:
        print(f"Error processing {file_path}: {str(e)}")

# Search for XLSX files only in the specified folder, not in subfolders
for file in os.listdir(folder_path):
    if file.startswith("vehicles-autoedit-") and file.endswith(".xlsx"):
        try:
            # Extract the timestamp from the filename using regex
            timestamp_match = re.search(r'vehicles-autoedit-(\d{14})', file)
            
            if timestamp_match:
                # Process the file and create the new cleaned file
                process_excel_file(os.path.join(folder_path, file))
            else:
                print(f"Error processing {file}: Unable to extract timestamp from filename")
        except Exception as e:
            print(f"Error processing {file}: {str(e)}")

print("Processing complete.")
