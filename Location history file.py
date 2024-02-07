import os
import pandas as pd
import glob

def process_location(file_path):
    df = pd.read_csv(file_path, encoding='utf-8-sig')
    original_row_count = len(df)

    df['Date Arrived'] = pd.to_datetime(df['Date Arrived'], dayfirst=True)
    df = df.sort_values(by='Date Arrived')
    
    # Convert 'Location' column values to uppercase
    df['Location'] = df['Location'].str.upper()
    
    valid_locations = ['FORECOURT', 'READY TODAY', 'GL COMPOUND', 'SF COMPOUND', 'TEMP LOAN CAR', 'COMPANY CAR']
    
    # Convert elements in valid_locations to uppercase
    valid_locations = [location.upper() for location in valid_locations]
    
    df_filtered = df[df['Location'].isin(valid_locations)]
    rows_removed = original_row_count - len(df_filtered)

    modified_path = file_path.replace('.csv', '_modified.csv')
    df_filtered.to_csv(modified_path, index=False, encoding='utf-8-sig')
    print(f'Location file processed and saved as {modified_path}')
    print(f'Removed {rows_removed} rows from Location file.')

def main(input_folder):
    os.chdir(input_folder)

    # Location File
    location_files = glob.glob('vehicles-location-history*.csv')
    if location_files:
        process_location(location_files[0])

if __name__ == "__main__":
    input_folder = r'C:\Users\Tom\OneDrive - Motor Depot\Pricing'
    main(input_folder)
