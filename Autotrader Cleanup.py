import os
import csv
import re

# Define the directory path
directory_path = r'C:\Users\Tom\OneDrive - Motor Depot\Pricing\Autotrader Forecourt Export\Logged Forecourt\Processed\Test'

# Function to process a CSV file and remove rows with Retail price of 0 and empty Auto Trader Retail Rating
def process_csv_file(file_path):
    with open(file_path, 'r', newline='', encoding='latin-1') as file:
        # Use 'latin-1' encoding to handle potential character encoding issues
        csv_reader = csv.DictReader(file)
        rows = [row for row in csv_reader if row.get('Auto Trader Retail Rating') not in ('', '0')]

    # Modify the file name to include "_cleaned"
    cleaned_file_path = re.sub(r'\.csv$', '_cleaned.csv', file_path)

    with open(cleaned_file_path, 'w', newline='', encoding='latin-1') as file:
        fieldnames = csv_reader.fieldnames
        csv_writer = csv.DictWriter(file, fieldnames=fieldnames)
        csv_writer.writeheader()
        csv_writer.writerows(rows)

# Recursively traverse the directory and its subdirectories
for root, _, files in os.walk(directory_path):
    for file_name in files:
        if file_name.endswith('.csv'):
            file_path = os.path.join(root, file_name)
            print(f'Processing file: {file_path}')
            process_csv_file(file_path)

print('Processing complete.')
