import os
import shutil
import re

# Define the source and destination file format
source_format = r'track-(\d{2})_(\d{2})_(\d{4}).csv'
destination_format = r'track_\3_\2_\1.csv'

# Specify the directory path
directory_path = r'C:\Users\tom.drayson\OneDrive - Motor Depot\Exports\Autoview\Archived\Process'

# Check if the directory exists
if not os.path.exists(directory_path):
    print(f"Directory '{directory_path}' does not exist.")
else:
    # Iterate through files in the specified directory
    for filename in os.listdir(directory_path):
        # Check if the file matches the source format
        match = re.match(source_format, filename)
        if match:
            # Extract date components from the source format
            day, month, year = match.groups()

            # Construct the destination filename using the destination format
            destination_filename = re.sub(source_format, destination_format, filename)

            # Create a copy of the file with the new name
            source_path = os.path.join(directory_path, filename)
            destination_path = os.path.join(directory_path, destination_filename)

            # Check if the destination file already exists
            if not os.path.exists(destination_path):
                shutil.copy(source_path, destination_path)
                print(f'Copied {filename} to {destination_filename}')
            else:
                print(f'Destination file {destination_filename} already exists, skipping.')

    print('Done')
