import os
import csv
import glob
import shutil
from datetime import datetime

def read_csv(file_path):
    with open(file_path, mode='r', newline='') as f:
        reader = csv.reader(f)
        return [row for row in reader]

def write_csv(file_path, data):
    with open(file_path, mode='w', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(data)

def sort_files_by_date(files):
    return sorted(files, key=lambda x: datetime.strptime(x.split('-')[-1].replace('.csv', ''), '%d_%m_%Y'), reverse=True)

# Path to the master CSV and the folder containing daily track CSVs
master_csv = "C:\\Users\\Tom\\OneDrive - Motor Depot\\Autoview\\__MASTER_Track.csv"
track_folder = "C:\\Users\\Tom\\OneDrive - Motor Depot\\Autoview\\"
archive_folder = "C:\\Users\\Tom\\OneDrive - Motor Depot\\Autoview\\Archived\\"

# Read master CSV data
master_data = read_csv(master_csv)

# Get a list of all track-dd_mm_yyyy.csv files and sort them by date
track_files = glob.glob(os.path.join(track_folder, "track-*.csv"))
sorted_track_files = sort_files_by_date(track_files)

# Iterate through each track CSV, starting with the most recent, and append its data to the master CSV
for track_file in sorted_track_files:
    track_data = read_csv(track_file)
    header = track_data[0]
    if header != master_data[0]:
        print(f"Skipping {track_file} due to mismatched header.")
        continue
    master_data[1:1] = track_data[1:]  # Skip appending the header row
    shutil.move(track_file, os.path.join(archive_folder, os.path.basename(track_file)))  # Move the file to the archive folder

# Write the updated master data back to the master CSV file
write_csv(master_csv, master_data)
