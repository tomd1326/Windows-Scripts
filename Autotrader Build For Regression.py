import os
import pandas as pd
from datetime import datetime

def extract_date_from_filename(filename):
    """Extract date from the filename in the format 'Exported Forecourt_yyyy_mm_dd.csv'."""
    try:
        date_part = filename.split('_')[1:]
        date_str = '_'.join(date_part)[:10]  # Extract yyyy_mm_dd
        return datetime.strptime(date_str, '%Y_%m_%d')
    except (IndexError, ValueError):
        return None

# Find the home directory of the current user
home_directory = os.path.expanduser('~')

# Define the OneDrive folder path based on the home directory
one_drive_folder = os.path.join(home_directory, 'OneDrive - Motor Depot')

# Define the subfolder path within OneDrive
subfolder_path = 'Pricing/Autotrader Forecourt Export/Logged/2024_01/Test'

# Combine the OneDrive folder path with the subfolder path
folder_path = os.path.join(one_drive_folder, subfolder_path)

# List all CSV files in the folder
csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]

# Filter files with valid dates and sort them by date (most recent first)
csv_files = sorted(
    [f for f in csv_files if extract_date_from_filename(f) is not None],
    key=extract_date_from_filename,
    reverse=True
)

# Create a dictionary to store DataFrames for each month
monthly_data = {}

# Read and combine CSV files for each month
for file in csv_files:
    file_path = os.path.join(folder_path, file)
    df = pd.read_csv(file_path, encoding='utf-8-sig')
    date = extract_date_from_filename(file)
    month_key = date.strftime('%Y_%m')
    
    if month_key not in monthly_data:
        monthly_data[month_key] = df
    else:
        monthly_data[month_key] = pd.concat([monthly_data[month_key], df])

    print(f"Processed file: {file}")

# Remove duplicates based on the StockID column for each month
for month_key, month_df in monthly_data.items():
    month_df.drop_duplicates(subset='StockId', keep='first', inplace=True)
    # Save the combined file for each month
    month_df.to_csv(os.path.join(folder_path, f'combined_csv_{month_key}.csv'), index=False, encoding='utf-8-sig')

print("All files have been processed, duplicates removed, and combined by month.")
