import os
import pandas as pd
from datetime import datetime
import shutil
import subprocess

# Function to remove BOM from a string
def remove_bom(text):
    return text.encode('utf-8').decode('utf-8-sig')

def remove_bom_from_csv(file_path):
    with open(file_path, 'rb') as file:
        content = file.read()
    if content.startswith(b'\xef\xbb\xbf'):
        content = content[3:]
    with open(file_path, 'wb') as file:
        file.write(content)

def check_and_correct_description_header(file_path):
    encodings_to_try = ['utf-8-sig', 'utf-8', 'ISO-8859-1']
    for encoding in encodings_to_try:
        try:
            df = pd.read_csv(file_path, encoding=encoding, low_memory=False)
            if 'ï»¿Description' in df.columns:
                df.rename(columns={'ï»¿Description': 'Description'}, inplace=True)
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
            return
        except Exception as e:
            pass
    print(f"1) Error checking and correcting 'Description' header in file '{file_path}': Unable to determine correct encoding")

def move_exported_files(exported_files, destination_dir, input_dir):
    for file in exported_files:
        source_path = os.path.join(input_dir, file)
        destination_path = os.path.join(destination_dir, file)
        try:
            shutil.move(source_path, destination_path)
            print(f"2) Moved '{source_path}' to '{destination_path}'")
        except Exception as e:
            print(f"2) Error moving '{source_path}' to '{destination_path}': {e}")

def process_forecourt_files():
    input_dir = r'C:\Users\Tom\OneDrive - Motor Depot\Pricing\Autotrader Forecourt Export'
    logged_forecourt_dir = r'C:\Users\Tom\OneDrive - Motor Depot\Pricing\Autotrader Forecourt Export\Logged Forecourt'
    backups_dir = r'C:\Users\Tom\OneDrive - Motor Depot\Pricing\Autotrader Forecourt Export\__BACKUPS'

    csv_files = [file for file in os.listdir(input_dir) if file.endswith('.csv')]
    encodings_to_try = ['utf-8-sig', 'utf-8', 'ISO-8859-1']

    for file_name in csv_files:
        full_path = os.path.join(input_dir, file_name)
        check_and_correct_description_header(full_path)

    master_df = pd.DataFrame()  # Initialize with an empty DataFrame
    total_rows_in_master = 0  # Initialize with a default value of 0
    master_files = [file for file in os.listdir(input_dir) if
                    file.startswith('__MASTER_Forecourt') and file.endswith('.csv')]
    if master_files:
        master_files.sort(reverse=True)
        latest_master_path = os.path.join(input_dir, master_files[0])

        # Attempt multiple encodings for reading master file
        for encoding in encodings_to_try:
            try:
                master_df = pd.read_csv(latest_master_path, encoding=encoding, header=0, low_memory=False)
                if 'Vehicle check issues' in master_df.columns:
                    master_df.drop('Vehicle check issues', axis=1, inplace=True)
                total_rows_in_master = len(master_df)  # Get the row count before modifications
                break  # Success, exit the loop
            except Exception as e:
                pass  # Try the next encoding
    print(f"3) Rows in original MASTER Forecourt file: {total_rows_in_master}")

    exported_files = [file for file in os.listdir(input_dir) if file.startswith("Exported Forecourt")]
    exported_files.sort(reverse=True)
    exported_dfs = []
    total_rows_from_exported_files = 0
    for file in exported_files:
        df = pd.read_csv(os.path.join(input_dir, file), encoding='utf-8-sig', header=0, low_memory=False)
        check_and_correct_description_header(os.path.join(input_dir, file))
        if 'Vehicle check issues' in df.columns:
            df.drop('Vehicle check issues', axis=1, inplace=True)
        total_rows_from_file = len(df)
        print(f"4) Rows in '{file}': {total_rows_from_file}")
        total_rows_from_exported_files += total_rows_from_file
        exported_dfs.append(df)

    print(f"5) Total rows of the MASTER Forecourt file plus the Exported Forecourt files: {total_rows_in_master + total_rows_from_exported_files}")

    exported_df = pd.concat(exported_dfs, ignore_index=True)
    final_df = pd.concat([exported_df, master_df], ignore_index=True)

    rows_before_dedup = len(final_df)
    final_df.drop_duplicates(subset="VRM", inplace=True)
    rows_after_dedup = len(final_df)
    print(f"6) Rows removed by deduplication: {rows_before_dedup - rows_after_dedup}")

    current_date = datetime.now().strftime('%Y_%m_%d')
    output_file_name = f'__MASTER_Forecourt_{current_date}.csv'
    count = 1
    while os.path.exists(os.path.join(input_dir, output_file_name)):
        output_file_name = f'__MASTER_Forecourt_{current_date}_new_{count}.csv'
        count += 1
    output_path = os.path.join(input_dir, output_file_name)
    
    final_df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"7) Rows on the new MASTER Forecourt file: {rows_after_dedup}")

    # Move the old MASTER Forecourt file to the backups directory
    if latest_master_path:
        backup_path = os.path.join(backups_dir, os.path.basename(latest_master_path))
        try:
            shutil.move(latest_master_path, backup_path)
            print(f"8) Moved old MASTER Forecourt file to '{backup_path}'")
        except Exception as e:
            print(f"8) Error moving old MASTER Forecourt file to '{backup_path}': {e}")

    move_exported_files(exported_files, logged_forecourt_dir, input_dir)

    # Running "Autotrader Initial.py" after completing the current process
    script_directory = os.path.dirname(os.path.abspath(__file__))  # Get the directory of the current script
    autotrader_initial_script = os.path.join(script_directory, "Autotrader Initial.py")
    subprocess.run(["python", autotrader_initial_script], check=True)  # Run "Autotrader Initial.py"

if __name__ == "__main__":
    process_forecourt_files()
