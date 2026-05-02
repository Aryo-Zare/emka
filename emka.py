
# pyxlsb : should be instaled in the enviornment to be able to read '.xlsb' files.

# %% copy excel files.

# only copy the excel files to a separate folder.

import shutil
from pathlib import Path

# %%%'

# 1. Define your base directories
source_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data")
dest_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\copy_excel")

# 2. Define the file extensions you want to capture
excel_extensions = {'.xls', '.xlsx', '.xlsb', '.xlsm'}

print("Starting file copy process...")
count = 0

# 3. Iterate through all files in the source directory and subdirectories
for file_path in source_dir.rglob('*'):
    # Check if the file is an Excel file
    if file_path.suffix.lower() in excel_extensions:
        
        # Figure out the relative path (e.g., ZC13\EMKA\Housing\data.xlsx)
        relative_path = file_path.relative_to(source_dir)
        
        # Create the exact same path in the destination directory
        new_file_path = dest_dir / relative_path
        
        # Ensure the subdirectories exist in the destination before copying
        new_file_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Copy the file (copy2 preserves file metadata like creation dates)
        shutil.copy2(file_path, new_file_path)
        # print(f"Copied: {relative_path}")
        count += 1

print(f"\nSuccess! Copied {count} Excel files to {dest_dir}")

# %%%'

# Starting file copy process...
# copied in 1s.
# Success! Copied 1452 Excel files to F:\OneDrive - Uniklinik RWTH Aachen\EMKA\copy_excel


# %% extract

# extract the data-segment from the excel files.

import pandas as pd
from pathlib import Path
import logging

# %%%'


# 1. Define your folders
base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel")
# base_dir = Path(r'F:\temp\4')
output_file = base_dir / "Master_Telemetry_Dataset.csv"
log_file = base_dir / "extraction_log.txt"

# 2. Set up the Logger (Writes to both console and a text file)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    force=True, # otherwise, it will save the log to the first file-locatio it was given to in this Python section ( python kernel restart ).
    handlers=[
        logging.FileHandler(log_file, mode='w', encoding='utf-8'), # Saves to file (overwrites old log)
        logging.StreamHandler() # Prints to console
    ]
)

all_data = []
logging.info("Starting robust omni-column data extraction...")

# 3. Process the files
for file_path in base_dir.rglob('*'):
    if file_path.suffix.lower() in ['.xls', '.xlsx', '.xlsb']:
        try:
            engine = 'pyxlsb' if file_path.suffix.lower() == '.xlsb' else 'openpyxl'
            df_raw = pd.read_excel(file_path, header=None, engine=engine)
            
            # STEP A: Search ALL columns for the "steps section" marker
            mask_steps = df_raw.astype(str).apply(lambda col: col.str.contains('steps section', case=False, na=False))
            step_rows = mask_steps.any(axis=1) 
            
            if step_rows.any():
                steps_row_idx = step_rows[step_rows].index[0]
                
                # STEP B: Search ALL columns for 'cpu-date' BELOW the "steps section"
                df_below_steps = df_raw.iloc[steps_row_idx:]
                mask_cpu = df_below_steps.astype(str).apply(lambda col: col.str.contains('cpu-date', case=False, na=False))
                cpu_rows = mask_cpu.any(axis=1)
                
                if cpu_rows.any():
                    header_idx = cpu_rows[cpu_rows].index[0]
                    
                    # Figure out exactly WHICH column contains 'cpu-date'
                    col_matches = mask_cpu.loc[header_idx]
                    cpu_date_col_idx = col_matches[col_matches].index[0]
                    
                    #==================================================================
                    #---- STEP C: 
                    # Combine Units, Deduplicate, and Find the End
                    
                    # 1. Combine Name and Unit for the column headers (e.g., "HR_aver (bpm)")
                    col_names = df_raw.iloc[header_idx].fillna('').astype(str).str.strip()
                    col_units = df_raw.iloc[header_idx + 1].fillna('').astype(str).str.strip()
                    
                    new_columns = []
                    for name, unit in zip(col_names, col_units):
                        if unit and unit.lower() != 'nan':
                            new_columns.append(f"{name}_({unit})")
                        else:
                            new_columns.append(name)

                    # 2. Deduplicate the column names (turns duplicate "HR_aver" into "HR_aver_1")
                    seen_cols = {}
                    deduped_columns = []
                    for col in new_columns:
                        if col in seen_cols:
                            seen_cols[col] += 1
                            deduped_columns.append(f"{col}_{seen_cols[col]}")
                        else:
                            seen_cols[col] = 0
                            deduped_columns.append(col)
                            
                    # 3. Find the first EMPTY ROW starting AFTER the units row (header_idx + 2)
                    df_below_units = df_raw.iloc[header_idx + 2:]
                    is_empty_row = pd.isna(df_below_units[cpu_date_col_idx]) | (df_below_units[cpu_date_col_idx].astype(str).str.strip() == '')
                    empty_indices = df_below_units[is_empty_row].index
                    
                    if not empty_indices.empty:
                        end_idx = empty_indices[0] # Back to 0! Because we started searching below the unit row.
                    else:
                        end_idx = len(df_raw) 
                        
                    # 4. Extract the clean chunk
                    df_chunk = df_raw.iloc[header_idx + 2 : end_idx].copy()
                    
                    # Apply the deduplicated, unit-inclusive column names
                    df_chunk.columns = deduped_columns 
                    
                    # Clean up: Drop completely empty columns
                    df_chunk.dropna(axis=1, how='all', inplace=True)
                    
                    
                    #==================================================================
                    #---- Metadata
                    df_chunk['Source_File'] = file_path.name
                    df_chunk['directory'] = str(file_path.parent)
                    
                    #---- DATE/TIME FIX BLOCK ----
                    # otherwise they will be saved as serial numbers :
                        # year : number of days from year 1900
                        # time : dcimal fraction of the time from a 24h period !
                    
                    # 1. Fix the 'cpu-date' column (Format: 20-Apr-21)
                    if 'cpu-date' in df_chunk.columns:
                        numeric_dates = pd.to_numeric(df_chunk['cpu-date'], errors='coerce')
                        df_chunk['cpu-date'] = pd.to_datetime(numeric_dates, unit='D', origin='1899-12-30').dt.strftime('%d-%b-%y')

                    # 2. Fix the 'cpu-time' column (Format: HH:MM:SS)
                    if 'cpu-time' in df_chunk.columns:
                        numeric_times = pd.to_numeric(df_chunk['cpu-time'], errors='coerce')
                        df_chunk['cpu-time'] = pd.to_datetime(numeric_times, unit='D', origin='1899-12-30').dt.strftime('%H:%M:%S')
                        
                    # 3. Fix the 'period-time' column (Format: HH:MM:SS)
                    if 'period-time' in df_chunk.columns:
                        numeric_periods = pd.to_numeric(df_chunk['period-time'], errors='coerce')
                        df_chunk['period-time'] = pd.to_datetime(numeric_periods, unit='D', origin='1899-12-30').dt.strftime('%H:%M:%S')
                        
                    #------------------------------------------
                    
                    all_data.append(df_chunk)
                    logging.info(f"SUCCESS: Extracted {len(df_chunk)} rows from {file_path.name}")
                else:
                    logging.warning(f"SKIPPED {file_path.name}: Found 'steps section', but no 'cpu-date' below it.")
            else:
                logging.warning(f"SKIPPED {file_path.name}: No 'steps section' found.")
                
        except Exception as e:
            logging.error(f"ERROR processing {file_path.name}: {e}")

# 4. Merge and Export
if all_data:
    final_dataset = pd.concat(all_data, ignore_index=True)
    final_dataset.to_csv(output_file, index=False)
    logging.info(f"DONE! Master dataset created with {len(final_dataset)} rows. Saved to: {output_file}")
else:
    logging.error("No data extracted. Please verify file paths and data structure.")

# %%% excel

# also save to excel.

output_file = base_dir / "Master_Telemetry_Dataset.xlsx"
final_dataset.to_excel(output_file, index=False)


final_dataset.shape
    # Out[11]: (66049, 18)

# %% log stats

# explore the save log-file & extract useful info !

import re
from pathlib import Path

# %%'

# 1. Point this to your log file
log_file = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\extraction_log__.txt")

# Counters for our statistics
stats = {
    'total_tested': 0,
    'success': 0,
    'skip_no_steps': 0,
    'skip_no_cpu_date': 0,
    'errors': 0,
    'total_rows': 0
}

# Lists to hold the names of problematic files
files_with_errors = []
files_missing_cpu = []

print("Analyzing log file...\n" + "-"*40)

# 2. Read the log file line by line
try:
    with open(log_file, 'r', encoding='utf-8') as file:
        for line in file:
            
            # Count Successes and extract row counts
            if "SUCCESS:" in line:
                stats['success'] += 1
                stats['total_tested'] += 1
                
                # Use regex to find the number between "Extracted" and "rows"
                match = re.search(r"Extracted (\d+) rows", line)
                if match:
                    stats['total_rows'] += int(match.group(1))
                    
            # Count "No steps section" warnings
            elif "No 'steps section' found" in line:
                stats['skip_no_steps'] += 1
                stats['total_tested'] += 1
                
            # Count "No cpu-date" warnings and save the filename
            elif "no 'cpu-date' below it" in line:
                stats['skip_no_cpu_date'] += 1
                stats['total_tested'] += 1
                
                # Extract filename using regex
                file_match = re.search(r"SKIPPED (.*?):", line)
                if file_match:
                    files_missing_cpu.append(file_match.group(1))
                    
            # Count real Errors and save the filename
            elif "ERROR processing" in line:
                stats['errors'] += 1
                stats['total_tested'] += 1
                
                file_match = re.search(r"ERROR processing (.*?):", line)
                if file_match:
                    files_with_errors.append(file_match.group(1))

    # 3. Print the Final Report
    print(f"OVERALL FILE STATISTICS")
    print(f"Total Files Processed : {stats['total_tested']}")
    print(f"  - Successfully Merged : {stats['success']} ({round((stats['success']/stats['total_tested'])*100, 1)}%)")
    print(f"  - Skipped (No steps)  : {stats['skip_no_steps']}")
    print(f"  - Skipped (No cpu)    : {stats['skip_no_cpu_date']}")
    print(f"  - Hard Errors         : {stats['errors']}")
    print("-" * 40)
    
    print(f"DATA VOLUME STATISTICS")
    print(f"Total Rows Extracted  : {stats['total_rows']}")
    if stats['success'] > 0:
        print(f"Average Rows per File : {round(stats['total_rows'] / stats['success'], 1)}")
    print("-" * 40)
    
    if stats['errors'] > 0:
        print(f"\nFiles that threw Hard Errors (Requires manual review):")
        for f in files_with_errors:
            print(f" - {f}")

except FileNotFoundError:
    print(f"Could not find the log file at: {log_file}")

# %%% out

# Analyzing log file...
# ----------------------------------------
# OVERALL FILE STATISTICS
# Total Files Processed : 1452
#   - Successfully Merged : 1206 (83.1%)
#   - Skipped (No steps)  : 234
#   - Skipped (No cpu)    : 0
#   - Hard Errors         : 12
# ----------------------------------------
# DATA VOLUME STATISTICS
# Total Rows Extracted  : 66049
# Average Rows per File : 54.8
# ----------------------------------------

# Files that threw Hard Errors (Requires manual review):
#  - ~$zc34_0b3f_2021_march_16_01.x00.xlsb
#  - zc33_0b3e_2021_march_17_01.x00.xlsb
#  - zc32_0a65_2021_march_16_01.x00.xlsb
#  - ~$zc32_0a65_2021_march_16_01.x00.xlsb
#  - ~$zc31_0ac6_2021_march_09_01.x00.xlsb
#  - ~$zc30_0a69_rx_back-housing_2021_02_04-2.x00.xlsb
#  - ~$zc28_0ac7_rx_front-housing_2021_01_19.x00.xlsb
#  - ~$zc28_0ac7_rx_front-housing_2021_01_22.x00.xlsb
#  - ~$zc26_0ae4_rx_front-housing_2020_11_16.x00.xlsb
#  - ~$zc17_0a66_rx_front-housing_2020_09_01.x00.xlsb
#  - ~$zc11_0a64_rx_of_2020_07_18.x01.xlsb
#  - ~$zc07_0a12_rx_front-housing_2020_06_12.x00.xlsb

# %% explore original column names

# test
# this is from the the older version of the extract-cell :
    # then, the column names & units were not merged
    # duplicate column names were also not re-named.

all_data[1].columns
    # Out[96]: 
    # Index(['cpu-date', 'cpu-time', 'period-time', 'mark-label', 'step-index',
    #        'BB__aver', 'HR__aver', 'DBP__aver', 'SBP__aver', 'MBP__aver',
    #        'aver__aver', 'aver__aver', 'HR__aver', 'Source_File', 'directory'],
    #       dtype='object', name=208)

all_data[100].columns
    # Out[97]: 
    # Index(['cpu-date', 'cpu-time', 'period-time', 'mark-label', 'step-index',
    #        'BB__aver', 'HR__aver', 'DBP__aver', 'SBP__aver', 'MBP__aver',
    #        'aver__aver', 'aver__aver', 'HR__aver', 'Source_File', 'directory'],
    #       dtype='object', name=208)

all_data[1000].columns
    # Out[98]: 
    # Index(['cpu-date', 'cpu-time', 'period-time', 'mark-label', 'step-index',
    #        'BB__aver', 'HR__aver', 'DBP__aver', 'SBP__aver', 'MBP__aver',
    #        'aver__aver', 'aver__aver', 'HR__aver', 'Source_File', 'directory'],
    #       dtype='object', name=209)

all_data[-1].columns
    # Out[99]: 
    # Index(['cpu-date', 'cpu-time', 'period-time', 'mark-label', 'step-index',
    #        'BB__aver', 'HR__aver', 'DBP__aver', 'SBP__aver', 'MBP__aver',
    #        'aver__aver', 'aver__aver', 'HR__aver', 'Source_File', 'directory'],
    #       dtype='object', name=208)


# %%% stat

# Create a set of all unique column structures in your list
unique_col_sets = set(tuple(df.columns) for df in all_data)

print(f"Found {len(unique_col_sets)} completely different column structures out of {len(all_data)} files.\n")

# Print them out to see the differences
for i, cols in enumerate(unique_col_sets):
    print(f"Structure {i+1} (Length: {len(cols)}):")
    print(cols)
    print("-" * 40)


    # Found 6 completely different column structures out of 1206 files.
    
    # Structure 1 (Length: 15):
    # ('cpu-date', 'cpu-time', 'period-time', 'mark-label', 'step-index', 'BB__aver', 'HR__aver', 'DBP__aver', 'SBP__aver', 'MBP__aver', 'aver__aver', 'aver__aver', 'HR__aver', 'Source_File', 'directory')
    # ----------------------------------------
    # Structure 2 (Length: 7):
    # ('cpu-date', 'cpu-time', 'period-time', 'step-index', 'HR__aver', 'Source_File', 'directory')
    # ----------------------------------------
    # Structure 3 (Length: 8):
    # ('cpu-date', 'cpu-time', 'period-time', 'mark-label', 'step-index', 'HR__aver', 'Source_File', 'directory')
    # ----------------------------------------
    # Structure 4 (Length: 16):
    # ('cpu-date', 'cpu-time', 'period-time', 'mark-label', 'step-index', 'BB__aver', 'HR__aver', 'DBP__aver', 'SBP__aver', 'MBP__aver', 'aver__aver', 'aver__aver', 'HR__aver', 'Abweichung in% HR vs HR', 'Source_File', 'directory')
    # ----------------------------------------
    # Structure 5 (Length: 16):
    # ('cpu-date', 'cpu-time', 'period-time', 'mark-label', 'step-index', 'BB__aver', 'HR__aver', 'DBP__aver', 'SBP__aver', 'MBP__aver', 'aver__aver', 'aver__aver', 'HR__aver', nan, 'Source_File', 'directory')
    # ----------------------------------------
    # Structure 6 (Length: 14):
    # ('cpu-date', 'cpu-time', 'period-time', 'step-index', 'BB__aver', 'HR__aver', 'DBP__aver', 'SBP__aver', 'MBP__aver', 'aver__aver', 'aver__aver', 'HR__aver', 'Source_File', 'directory')
    # ----------------------------------------


# %%% pickle

# dump the list of all the extracted data ( output of the loop ).

import pickle
from pathlib import Path

# Assuming base_dir is still defined from your previous script. 
# If not, just redefine it: base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\copy_excel")
backup_file = base_dir / "all_data_raw_backup.pkl"

print("Saving binary backup...")

# Open the file in 'wb' (Write Binary) mode
with open(backup_file, 'wb') as file:
    pickle.dump(all_data, file)

print(f"Success! Backed up exactly as it is in memory to:\n{backup_file}")


    # Saving binary backup...
    # Success! Backed up exactly as it is in memory to:
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\all_data_raw_backup.pkl

# %%%% load pickle

import pickle
from pathlib import Path

backup_file = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\copy_excel\all_data_raw_backup.pkl")

# Open the file in 'rb' (Read Binary) mode
with open(backup_file, 'rb') as file:
    all_data = pickle.load(file)

print(f"Loaded {len(all_data)} dataframes from backup!")

# %%'


