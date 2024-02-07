import os
import pandas as pd
from datetime import datetime
import shutil

# Folder path
folder_path = r'C:\Users\Tom\OneDrive - Motor Depot\Pricing\Autotrader Forecourt Export\Logged Forecourt'

# List all csv files in the folder
csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv') and 'Exported Forecourt' in f]
csv_files.sort()  # Sort files by name, which also sorts by date due to the naming convention

# Function to rename files by adding "_new" if they already exist
def rename_file_if_exists(file_path):
    base_name, file_ext = os.path.splitext(file_path)
    new_file_path = file_path
    counter = 1
    while os.path.exists(new_file_path):
        new_base_name = f"{base_name}_new_{counter}"
        new_file_path = f"{new_base_name}{file_ext}"
        counter += 1
    return new_file_path

# Find the old MASTER file (most recent based on date)
master_files = [f for f in os.listdir(folder_path) if f.startswith('__MASTER_Initial_AT_')]
if master_files:
    old_master_file = sorted(master_files, key=lambda x: datetime.strptime(x[20:30], '%Y_%m_%d'))[-1]
    old_master_path = os.path.join(folder_path, old_master_file)
else:
    print("No old MASTER file found.")
    old_master_path = None

# Today's date for the new MASTER file
today_date = datetime.now().strftime('%Y_%m_%d')
new_master_file = os.path.join(folder_path, f'__MASTER_Initial_AT_{today_date}.csv')

# Rename new MASTER file if it already exists
new_master_file = rename_file_if_exists(new_master_file)

# Initialize an empty DataFrame for the new MASTER file
new_master_df = pd.DataFrame()

# Process the old MASTER file
if old_master_path:
    try:
        old_master_df = pd.read_csv(old_master_path)
        print(f"Rows read from old MASTER file ({old_master_file}): {len(old_master_df)}")
        new_master_df = pd.concat([new_master_df, old_master_df], ignore_index=True)
    except UnicodeDecodeError:
        try:
            old_master_df = pd.read_csv(old_master_path, encoding='ISO-8859-1')
            print(f"Rows read from old MASTER file ({old_master_file}) with ISO-8859-1 encoding: {len(old_master_df)}")
            new_master_df = pd.concat([new_master_df, old_master_df], ignore_index=True)
        except Exception as e:
            print(f"Error reading {old_master_file} with alternate encoding: {e}")
    except Exception as e:
        print(f"Error reading {old_master_file}: {e}")

# Process each Exported Forecourt file
for file in csv_files:
    try:
        df = pd.read_csv(os.path.join(folder_path, file))
        original_row_count = len(df)
        df = df[df['Retail price'] != 0]  # Remove rows where Retail price is zero
        rows_removed = original_row_count - len(df)
        new_master_df = pd.concat([new_master_df, df], ignore_index=True)
        print(f"Rows removed from {file} with Retail Price of 0: {rows_removed}")
        print(f"Rows copied from {file}: {len(df)}")
    except Exception as e:
        print(f"Error processing {file}: {e}")

# Deduplicate based on VRM
initial_row_count = len(new_master_df)
new_master_df = new_master_df.drop_duplicates(subset=['VRM'])
final_row_count = len(new_master_df)

# Print summary
print(f"Rows removed by deduplication of VRM: {initial_row_count - final_row_count}")
print(f"Total rows in new MASTER file: {final_row_count}")

# Save the new MASTER file
try:
    new_master_df.to_csv(new_master_file, index=False, encoding='ISO-8859-1')
    print(f"New MASTER file saved as {new_master_file}")
except Exception as e:
    print(f"Error saving the new MASTER file: {e}")

# Move the old MASTER file to BACKUP folder
if old_master_path:
    backup_folder = os.path.join(folder_path, 'BACKUP')
    os.makedirs(backup_folder, exist_ok=True)
    shutil.move(old_master_path, os.path.join(backup_folder, old_master_file))
    print(f"Moved old MASTER file to {os.path.join(backup_folder, old_master_file)}")

# Move all processed Exported Forecourt files to Processed folder
processed_folder = os.path.join(folder_path, 'Processed')
os.makedirs(processed_folder, exist_ok=True)
for file in csv_files:
    shutil.move(os.path.join(folder_path, file), os.path.join(processed_folder, file))
    print(f"Moved {file} to {os.path.join(processed_folder, file)}")
