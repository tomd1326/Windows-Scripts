import os
import glob
import pandas as pd

# Directory containing track files
directory_path = r'C:\Users\tom.drayson\OneDrive - Motor Depot\Exports\Autoview\Track Files\Process'

# Pattern to match the desired track files
file_pattern = 'track_????_??_??.csv'

# Find all matching track files in the directory
file_list = glob.glob(os.path.join(directory_path, file_pattern))

# Function to extract the date from a track file name
def extract_date(file_name):
    return file_name[-14:-4]

# Sort the files based on the date extracted from the filename
file_list.sort(key=extract_date, reverse=True)

# Create an empty DataFrame to store the combined data
combined_data = pd.DataFrame()

# Iterate through the sorted file list and combine the data
for file_path in file_list:
    data = pd.read_csv(file_path)  # Read each CSV file
    combined_data = pd.concat([combined_data, data], ignore_index=True)  # Concatenate data

# Name of the combined output file in the "Process" folder
output_file_name = os.path.join(directory_path, 'combined_track_files.csv')

# Save the combined data to a new CSV file in the "Process" folder
combined_data.to_csv(output_file_name, index=False)

# Display the number of matching files found, rows added, and the name of the combined file
print(f"Number of matching files found: {len(file_list)}")
print(f"Number of rows added to the combined file: {len(combined_data)}")
print(f"Combined file '{output_file_name}' created in the 'Process' folder.")
