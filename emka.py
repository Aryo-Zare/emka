
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
base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
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
    # "and not ..." : to avoid it trying to open hidden Excel lock files ( prefix : ~$ ).
    if file_path.suffix.lower() in ['.xls', '.xlsx', '.xlsb'] and not file_path.name.startswith('~$'):
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

# %%% excel , pickle

# also save to excel.

output_file = base_dir / "Master_Telemetry_Dataset.xlsx"
final_dataset.to_excel(output_file, index=False)

output_file = base_dir / "Master_Telemetry_Dataset.pkl"
final_dataset.to_pickle(output_file)


final_dataset.shape
    # Out[11]: (66049, 18)
    
# %%'
    
base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
source_file = base_dir / "Master_Telemetry_Dataset.pkl"
df_master = pd.read_pickle(source_file)

# %% log stats

# explore the save log-file & extract useful info !

import re
from pathlib import Path

# %%%'

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

# %% columns

list(df_master.columns)
    # Out[14]: 
    # ['cpu-date',
    #  'cpu-time',
    #  'period-time',
    #  'mark-label',
    #  'step-index',
    #  'BB__aver_(ms)',
    #  'HR__aver_(bpm)',
    #  'DBP__aver_(mmHg)',
    #  'SBP__aver_(mmHg)',
    #  'MBP__aver_(mmHg)',
    #  'aver__aver_(°C)',
    #  'aver__aver_(%)',
    #  'HR__aver_(bpm)_1',
    #  'Source_File',
    #  'directory',
    #  'aver__aver_(g)',
    #  '_2',
    #  'Abweichung in% HR vs HR']

# %% suspicious_columns

# import pandas as pd
# from pathlib import Path


# List the suspicious columns you want to investigate
# (Make sure these exactly match the column names in your CSV)
suspicious_columns = ['aver__aver_(g)', '_2', 'Abweichung in% HR vs HR']

print("Hunting for data in suspicious columns...\n")

for col in suspicious_columns:
    if col in df_master.columns:
        # Filter for rows where the cell is NOT NaN AND NOT just an empty space
        mask_has_data = df_master[col].notna() & (df_master[col].astype(str).str.strip() != '')
        dirty_rows = df_master[mask_has_data]
        
        print(f"--- Column: '{col}' ---")
        print(f"Found {len(dirty_rows)} rows with actual data.")
        
        if len(dirty_rows) > 0:
            # Show the first 10 occurrences so you can see what the data actually is
            # We display the Source_File and directory so you know exactly where it came from
            preview = dirty_rows[['directory', 'Source_File', col]].head(10)
            print(preview.to_string(index=False))
        print("\n" + "="*50 + "\n")
    else:
        print(f"Column '{col}' not found in the dataset.\n")

# %%% out

    # Hunting for data in suspicious columns...
    
    # --- Column: 'aver__aver_(g)' ---
    # Found 17376 rows with actual data.
    #                                                             directory                         Source_File aver__aver_(g)
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\OP zc67_1a2c_rx_or_2023_10_23.x00.xlsb              0
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\OP zc67_1a2c_rx_or_2023_10_23.x00.xlsb              0
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\OP zc67_1a2c_rx_or_2023_10_23.x00.xlsb              0
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\OP zc67_1a2c_rx_or_2023_10_23.x00.xlsb          0.003
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\OP zc67_1a2c_rx_or_2023_10_23.x00.xlsb          0.009
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\OP zc67_1a2c_rx_or_2023_10_23.x00.xlsb              0
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\OP zc67_1a2c_rx_or_2023_10_23.x00.xlsb              0
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\OP zc67_1a2c_rx_or_2023_10_23.x00.xlsb          0.003
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\OP zc67_1a2c_rx_or_2023_10_23.x00.xlsb          0.004
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\OP zc67_1a2c_rx_or_2023_10_23.x00.xlsb          0.003
    
    # ==================================================
    
    # --- Column: '_2' ---
    # Found 222 rows with actual data.
    #                                                             directory                         Source_File                      _2
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC35\EMKA\OR zc35_0b72_2021_april_13_01.x01.xlsb Abweichung in% HR vs HR
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC35\EMKA\OR zc35_0b72_2021_april_13_01.x01.xlsb              338.577193
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC35\EMKA\OR zc35_0b72_2021_april_13_01.x01.xlsb              288.565303
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC35\EMKA\OR zc35_0b72_2021_april_13_01.x01.xlsb               262.37991
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC35\EMKA\OR zc35_0b72_2021_april_13_01.x01.xlsb                328.1783
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC35\EMKA\OR zc35_0b72_2021_april_13_01.x01.xlsb               319.52339
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC35\EMKA\OR zc35_0b72_2021_april_13_01.x01.xlsb              244.484411
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC35\EMKA\OR zc35_0b72_2021_april_13_01.x01.xlsb              281.566975
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC35\EMKA\OR zc35_0b72_2021_april_13_01.x01.xlsb              240.980898
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC35\EMKA\OR zc35_0b72_2021_april_13_01.x01.xlsb              240.213638
    
    # ==================================================
    
    # --- Column: 'Abweichung in% HR vs HR' ---
    # Found 161 rows with actual data.
    #                                                                               directory                           Source_File Abweichung in% HR vs HR
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC22\EMKA\Surgery\Implantation 0ac7_0ac7_2020_october_20_01.x01.xlsb               26.460404
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC22\EMKA\Surgery\Implantation 0ac7_0ac7_2020_october_20_01.x01.xlsb                19.99773
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC22\EMKA\Surgery\Implantation 0ac7_0ac7_2020_october_20_01.x01.xlsb               20.525496
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC22\EMKA\Surgery\Implantation 0ac7_0ac7_2020_october_20_01.x01.xlsb                5.273673
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC22\EMKA\Surgery\Implantation 0ac7_0ac7_2020_october_20_01.x01.xlsb                46.58261
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC22\EMKA\Surgery\Implantation 0ac7_0ac7_2020_october_20_01.x01.xlsb               19.032456
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC22\EMKA\Surgery\Implantation 0ac7_0ac7_2020_october_20_01.x01.xlsb                1.357057
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC22\EMKA\Surgery\Implantation 0ac7_0ac7_2020_october_20_01.x01.xlsb                6.625007
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC22\EMKA\Surgery\Implantation 0ac7_0ac7_2020_october_20_01.x01.xlsb                 6.18918
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC22\EMKA\Surgery\Implantation 0ac7_0ac7_2020_october_20_01.x01.xlsb                4.036679
    
    # ==================================================

# %% Metadata

# 2. Define a function to parse the path safely
def extract_path_info(path_string):
    # Convert the raw string into a Path object
    p = Path(str(path_string))
    
    try:
        # Find the exact index of 'copy_excel' in the folder hierarchy
        # e.g., ('F:\', 'OneDrive...', 'EMKA', 'data', 'copy_excel', 'ZC04', 'EMKA', 'Housing')
        anchor_idx = p.parts.index('copy_excel')
        
        # Grab the folders relative to the anchor
        sample_id = p.parts[anchor_idx + 1]  # 1 folder down
        setup = p.parts[anchor_idx + 3]      # 3 folders down
        
        return sample_id, setup
        
    except (ValueError, IndexError):
        # Failsafe: If the path is weirdly formatted, leave it blank rather than crashing
        # if 1 ValueError occurs, both values will be put as 'Unknown' !
        return "Unknown", "Unknown"
        
# %%% run

# 3. Apply the function to the 'directory' column to create the two new columns
# This unpacks the two extracted values directly into 'sample_ID' and 'setup'
df_master['sample_ID'], df_master['setup'] = zip(*df_master['directory'].apply(extract_path_info))

# %%% test

df_master['directory'].head()
    # Out[28]: 
    # 0    F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\Housing
    # 1    F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\Housing
    # 2    F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\Housing
    # 3    F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\Housing
    # 4    F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC69\EMKA\Housing
    # Name: directory, dtype: str

df_master[['sample_ID' , 'setup']].head()
    # Out[22]: 
    #   sample_ID    setup
    # 0      ZC69  Housing
    # 1      ZC69  Housing
    # 2      ZC69  Housing
    # 3      ZC69  Housing
    # 4      ZC69  Housing

# %%% timeline

# import pandas as pd
# from pathlib import Path

def extract_timeline(path_string):
    p = Path(str(path_string))
    
    try:
        anchor_idx = p.parts.index('copy_excel')

        # Safely check for the optional 4th folder (timeline)
        # We check if the total number of parts is strictly greater than the index we want to reach
        if len(p.parts) > (anchor_idx + 4):
            timeline = p.parts[anchor_idx + 4]
        else:
            timeline = "N/A"  # Or you can use None if you prefer actual blank cells
            
        return timeline
        
    except (ValueError, IndexError):
        return "Unknown"

# %%%%'

df_master['timeline'] = df_master['directory'].apply(extract_timeline)

list(df_master['timeline'].unique())
    # Out[30]: 
    # ['N/A',
    #  '1.Retraining',
    #  '2.Retraining',
    #  'POD 3',
    #  'POD 4',
    #  'POD 7',
    #  'POD 1',
    #  'Explantation',
    #  'Finale',
    #  'Implantation',
    #  'TI',
    #  '1.Wiederholung',
    #  '2.Wiederholung',
    #  'Ti',
    #  'Expl',
    #  'Impl',
    #  'Sacrifice',
    #  'Transponderimplantation',
    #  '1.re',
    #  '2.re',
    #  'Opening Seroma',
    #  'minütlicher meean',
    #  'Transponder Implantation']

# %%% problematic_directories

# Filter the dataframe for rows where the extraction failed
# failed_rows = df_master[df_master['sample_ID'] == 'Unknown']  : this also works , but the main problem is assignment to a setup directory !
    # 'sample_ID' is always present in the original directory !
    # so the problem is solely 'setup'.
failed_rows = df_master[df_master['setup'] == 'Unknown']

# get the unique combinations of 2 columns.
    # out : index is for the 1st occurence of the combination !
failed_rows[['directory','Source_File']].drop_duplicates()
    # Out[18]: 
    #                                                                 directory                                     Source_File
    # 36836  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC17\EMKA            0a66_0a66_2020_august_24_01.x00.xlsb
    # 43007  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA              0a11_0a11_2020_july_07_01.x00.xlsb
    # 60806  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC08\EMKA  zc08_0a13_rx_front-housing_2020_06_01.x00.xlsb
    # 62981  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC06\EMKA             0a11_0a11_rx_of_2020_03_15.x00.xlsb

failed_rows['sample_ID'].unique()
    # Out[41]: 
    # <StringArray>
    # ['Unknown']
    # Length: 1, dtype: str

# I checked them : there are single excel files inside : copy_excel\sample_ID\EMKA :
        # outside other subfolders ( 'Housing', ... ).
        # so they don't belong to any experimental-setup !

#======================================
#---- file check
# this is to check if the same file name appears in other directories !

unique_files = ['0a66_0a66_2020_august_24_01.x00.xlsb',
                '0a11_0a11_2020_july_07_01.x00.xlsb',
                'zc08_0a13_rx_front-housing_2020_06_01.x00.xlsb',
                '0a11_0a11_rx_of_2020_03_15.x00.xlsb'
                ]

df_unique_files = df_master[ df_master['Source_File'].isin(unique_files) ]

df_unique_files[['directory','Source_File']].drop_duplicates()
    # Out[44]: 
    #                                                                                directory                                     Source_File
    # 36836                 F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC17\EMKA            0a66_0a66_2020_august_24_01.x00.xlsb
    # 43007                 F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA              0a11_0a11_2020_july_07_01.x00.xlsb
    # 60465  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC09\EMKA\Surgery\Finale              0a11_0a11_2020_july_07_01.x00.xlsb
    # 60806                 F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC08\EMKA  zc08_0a13_rx_front-housing_2020_06_01.x00.xlsb
    # 61004         F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC08\EMKA\Housing  zc08_0a13_rx_front-housing_2020_06_01.x00.xlsb
    # 62981                 F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC06\EMKA             0a11_0a11_rx_of_2020_03_15.x00.xlsb

# as it turns out, 2 files are duplicated.
    # this is compatible wit Mareike's email on the nature of the duplicte files.
        # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC08\EMKA\Housing :
            # zc08_0a13_rx_front-housing_2020_06_01.x00.xlsb : should be assigned to POD TI+ 7 .
        # hence, all rows with the value 'Unknown' under the column 'setup' shoudl be deleted.

# %%%%'

# older method

    # # Get the unique original directory paths from those rows
    # problematic_directories = failed_rows['directory'].unique()
    
    # print(f"Found {len(problematic_directories)} problematic directories that didn't match the rule:\n")
    
    # # Print them out one by one so you can inspect them
    # for folder in problematic_directories:
    #     print(folder)
    
    # # out    
    #     # Found 4 problematic directories that didn't match the rule:
        
    #         # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC17\EMKA
    #         # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA
    #         # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC08\EMKA
    #         # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC06\EMKA


# %%%'

base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
output_file = base_dir / "Master_Telemetry_Dataset_2.pkl"
df_master.to_pickle(output_file)

# %%% unique
# %%%% setup unique

list(df_master["setup"].unique())
    # Out[32]: 
    # ['Housing',
    #  'OF',
    #  'OP',
    #  'Houisng',
    #  'OR',
    #  'housing',
    #  'Open Field',
    #  'Stoffwechselkäfig',
    #  'ZC30 housing',
    #  'ZC30 stoffwechselkäfig',
    #  'ZC30 or',
    #  'ZC30 Open Field',
    #  'ZC29 housing',
    #  'ZC29 stoffwechselkäfig',
    #  'ZC29 or',
    #  'ZC29 Open Field',
    #  'Surgery',
    #  'TI',
    #  'Unknown',
    #  'OP JoVe']

# %%%% sample_ID _ unique

list(df_master["sample_ID"].unique())
    # Out[33]: 
    # ['ZC69',
    #  'ZC68',
    #  'ZC67',
    #  'ZC66',
    #  'ZC65',
    #  'ZC64',
    #  'ZC63',
    #  'ZC62',
    #  'ZC61',
    #  'ZC60',
    #  'ZC38',
    #  'ZC37',
    #  'ZC36',
    #  'ZC35',
    #  'ZC34',
    #  'ZC33',
    #  'ZC32',
    #  'ZC31',
    #  'ZC30',
    #  'ZC29',
    #  'ZC28',
    #  'ZC27',
    #  'ZC26',
    #  'ZC25',
    #  'ZC24',
    #  'ZC23',
    #  'ZC22',
    #  'ZC21',
    #  'ZC20',
    #  'ZC19',
    #  'ZC18',
    #  'Unknown',
    #  'ZC17',
    #  'ZC16',
    #  'ZC15',
    #  'ZC14',
    #  'ZC13',
    #  'ZC12',
    #  'ZC11',
    #  'ZC10',
    #  'ZC09',
    #  'ZC08',
    #  'ZC07',
    #  'ZC06',
    #  'ZC05',
    #  'ZC04']

# %%%% mark-label

list(df_master["mark-label"].unique())
    # Out[59]: 
    # ['noname #1',
    #  nan,
    #  'Noname #1',
    #  '1.retraining',
    #  'POD 7',
    #  'Noname #2',
    #  'POD 3',
    #  'housing',
    #  'Noname #3',
    #  '2.retraining',
    #  'POD 1',
    #  'Noname #5',
    #  'POD 4',
    #  'noname #4']

# %%% clean

# 1. Create a dictionary of { 'Bad Name' : 'Good Name' }
corrections_Housing = {
    'Houisng': 'Housing',
    'housing': 'Housing',
    'ZC30 housing': 'Housing',
    'ZC29 housing': 'Housing'
}

# 2. Apply the replacement
df_master['setup'] = df_master['setup'].replace(corrections_Housing)

list(df_master["setup"].unique())
    # Out[35]: 
    # ['Housing',
    #  'OF',
    #  'OP',
    #  'OR',
    #  'Open Field',
    #  'Stoffwechselkäfig',
    #  'ZC30 stoffwechselkäfig',
    #  'ZC30 or',
    #  'ZC30 Open Field',
    #  'ZC29 stoffwechselkäfig',
    #  'ZC29 or',
    #  'ZC29 Open Field',
    #  'Surgery',
    #  'TI',
    #  'Unknown',
    #  'OP JoVe']

# %%%%'

# rest of the corrections.
corrections = {'ZC30 stoffwechselkäfig': 'Stoffwechselkäfig',
               'ZC30 or':'OR',
               'ZC30 Open Field':'OF',
               'ZC29 stoffwechselkäfig': 'Stoffwechselkäfig',
               'ZC29 or': 'OR',
               'ZC29 Open Field':'OF',
               'Open Field':'OF'
               }

# Apply the replacement
df_master['setup'] = df_master['setup'].replace(corrections)

list(df_master["setup"].unique())
    # Out[21]: 
    # ['Housing',
    #  'OF',
    #  'OP',
    #  'OR',
    #  'Stoffwechselkäfig',
    #  'Surgery',
    #  'TI',
    #  'Unknown',
    #  'OP JoVe']

# %%%%'

test = df_master[ df_master["setup"] == 'OP JoVe' ]

test['directory'][:4]
    # Out[25]: 
    # 39292    F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC16\EMKA\OP JoVe
    # 39293    F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC16\EMKA\OP JoVe
    # 39294    F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC16\EMKA\OP JoVe
    # 39295    F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC16\EMKA\OP JoVe
    # Name: directory, dtype: str

# %%%% ZC18

# this pig : TI is directly under EMKA :
    # copy_excel\ZC18\EMKA\TI
    # the normal should be :  copy_excel\ZC18\EMKA\surgery\TI

mask_setup_TI = df_master['setup'] == 'TI'
df_setup_TI = df_master[mask_setup_TI]
df_setup_TI.shape
    # Out[57]: (174, 22)

df_setup_TI['sample_ID'].unique()
    # Out[58]: 
    # <StringArray>
    # ['ZC18']
    # Length: 1, dtype: str


df_setup_TI[['sample_ID','setup','timeline']][:5]
    # Out[59]: 
    #       sample_ID setup timeline
    # 36662      ZC18    TI      N/A
    # 36663      ZC18    TI      N/A
    # 36664      ZC18    TI      N/A
    # 36665      ZC18    TI      N/A
    # 36666      ZC18    TI      N/A

# %%%%% fix

# 1. Create a mask that finds the exact rows with the mistake
# Use parentheses around each condition when combining with '&' (AND)
mask_mistake = (df_master['sample_ID'] == 'ZC18') & (df_master['setup'] == 'TI')

# 2. Use .loc[rows, columns] to overwrite the data
df_master.loc[mask_mistake, 'setup'] = 'Surgery'
df_master.loc[mask_mistake, 'timeline'] = 'TI'

# 3. Verify the changes worked
print("Verifying the fix for ZC18:")
mask_check = (df_master['sample_ID'] == 'ZC18') & (df_master['timeline'] == 'TI')
print(df_master.loc[mask_check, ['sample_ID', 'setup', 'timeline']].head())
    #       sample_ID    setup timeline
    # 36662      ZC18  Surgery       TI
    # 36663      ZC18  Surgery       TI
    # 36664      ZC18  Surgery       TI
    # 36665      ZC18  Surgery       TI
    # 36666      ZC18  Surgery       TI

# now, in the original dataset, under the column 'setup' : there would be no item as 'TI'.

# %%% unique_2

list(df_master["setup"].unique())
    # Out[14]: 
    # ['Housing',
    #  'OF',
    #  'OP',
    #  'OR',
    #  'Stoffwechselkäfig',
    #  'Surgery',
    #  'Unknown',
    #  'OP JoVe']


list(df_master["timeline"].unique())
    # Out[17]: 
    # ['N/A',
    #  '1.Retraining',
    #  '2.Retraining',
    #  'POD 3',
    #  'POD 4',
    #  'POD 7',
    #  'POD 1',
    #  'Explantation',
    #  'Finale',
    #  'Implantation',
    #  'TI',
    #  '1.Wiederholung',
    #  '2.Wiederholung',
    #  'Ti',
    #  'Expl',
    #  'Impl',
    #  'Sacrifice',
    #  'Transponderimplantation',
    #  '1.re',
    #  '2.re',
    #  'Opening Seroma',
    #  'minütlicher meean',
    #  'Transponder Implantation']

# %%%% correction-2

correction_setup = {'OP':'Surgery',
                    'OR':'Surgery'
                    }

df_master['setup'] = df_master['setup'].replace(correction_setup)

# byusing 'inplace=True' you get this warning :
    # c:\code\emka\emka.py:861: ChainedAssignmentError: A value is being set on a copy of a DataFrame or Series through chained assignment using an inplace method.
    # Such inplace method never works to update the original DataFrame or Series, because the intermediate object on which we are setting values always behaves as a copy (due to Copy-on-Write).

    # For example, when doing 'df[col].method(value, inplace=True)', try using 'df.method({col: value}, inplace=True)' instead, to perform the operation inplace on the original object, or try to avoid an inplace operation using 'df[col] = df[col].method(value)'.

    # See the documentation for a more detailed explanation: https://pandas.pydata.org/pandas-docs/stable/user_guide/copy_on_write.html
    #   df_master['setup'].replace(correction_setup, inplace=True)
# chatGPT recommendation : what is written above :
    # "
    # because:
    #     very readable
    #     avoids pandas inplace quirks (which pandas has been gradually discouraging)
    #     works consistently with Copy-on-Write
    # "

correction_timeline = {'POD 7':'Sacrifice',
                       'Finale':'Sacrifice',
                       
                       '1.Retraining':'Retraining_1',
                       '1.re' :'Retraining_1',
                       '1.Wiederholung':'Retraining_1',

                       '2.Retraining':'Retraining_2',
                       '2.re':'Retraining_2',
                       '2.Wiederholung':'Retraining_2',

                       'Expl':'Explantation',
                       'Impl':'Implantation',
                       
                       'Ti':'TI',
                       'Transponderimplantation':'TI',
                       'Transponder Implantation':'TI'
                       }

df_master['timeline'] = df_master['timeline'].replace(correction_timeline)


#============================================
#---- check

list(df_master["setup"].unique())
    # Out[27]: ['Housing', 'OF', 'Surgery', 'Stoffwechselkäfig', 'Unknown', 'OP JoVe']

list(df_master["timeline"].unique())
    # Out[26]: 
    # ['N/A',
    #  'Retraining_1',
    #  'Retraining_2',
    #  'POD 3',
    #  'POD 4',
    #  'Sacrifice',
    #  'POD 1',
    #  'Explantation',
    #  'Implantation',
    #  'TI',
    #  'Opening Seroma',
    #  'minütlicher meean']

# %%% delete

#---- inspect

df_Opening_Seroma = df_master[df_master["timeline"] == 'Opening Seroma']
df_Opening_Seroma['sample_ID'].unique()
    # Out[29]: 
    # <StringArray>
    # ['ZC13']
    # Length: 1, dtype: str


df_JoVe = df_master[df_master["setup"] == 'OP JoVe']
df_JoVe['sample_ID'].unique()
    # Out[31]: 
    # <StringArray>
    # ['ZC16']
    # Length: 1, dtype: str


# Mareike ( Wednesday, May 06, 2026 15:06 ) : "values from animal ZC06 should not be included."

#============================================
#---- delete

# 1. Define the list of animals to exclude
excluded_pigs = ['ZC06', 'ZC13', 'ZC16']

# Optional: Check how many rows you have before the deletion
print(f"Rows before deletion: {len(df_master)}")
    # Rows before deletion: 66049

# 2. Filter the dataframe
# The '~' symbol means "NOT". So this reads as: 
# "Keep rows where sample_ID is NOT IN the excluded_pigs list."
df_master = df_master[
                        ~df_master['sample_ID'].isin(excluded_pigs) 
                      ].copy()

# 3. Verify the changes
print(f"Rows after deletion: {len(df_master)}")
    # Rows after deletion: 64179

#============================================
#---- check
df_master['sample_ID'].unique()
    # Out[36]: 
    # <StringArray>
    # [   'ZC69',    'ZC68',    'ZC67',    'ZC66',    'ZC65',    'ZC64',    'ZC63',
    #     'ZC62',    'ZC61',    'ZC60',    'ZC38',    'ZC37',    'ZC36',    'ZC35',
    #     'ZC34',    'ZC33',    'ZC32',    'ZC31',    'ZC30',    'ZC29',    'ZC28',
    #     'ZC27',    'ZC26',    'ZC25',    'ZC24',    'ZC23',    'ZC22',    'ZC21',
    #     'ZC20',    'ZC19',    'ZC18', 'Unknown',    'ZC17',    'ZC15',    'ZC14',
    #     'ZC12',    'ZC11',    'ZC10',    'ZC09',    'ZC08',    'ZC07',    'ZC05',
    #     'ZC04']
    # Length: 43, dtype: str


list(df_master["setup"].unique())
    # Out[37]: ['Housing', 'OF', 'Surgery', 'Stoffwechselkäfig', 'Unknown']

list(df_master["timeline"].unique())
    # Out[38]: 
    # ['N/A',
    #  'Retraining_1',
    #  'Retraining_2',
    #  'POD 3',
    #  'POD 4',
    #  'Sacrifice',
    #  'POD 1',
    #  'Explantation',
    #  'Implantation',
    #  'TI',
    #  'minütlicher meean']

#============================================
#---- explore mask_Unknown numbers.

a = df_master['setup'] == 'Unknown'
b = ~( df_master['setup'] == 'Unknown' )

a.shape
    # Out[46]: (64179,)
b.shape
    # Out[49]: (64179,)

a.sum()
    # Out[51]: np.int64(382)
b.sum()
    # Out[50]: np.int64(63797)

#============================================
#---- exclude_Unkown

mask_exclude_Unkown = df_master['setup'] != 'Unknown'
df_master = df_master[mask_exclude_Unkown].copy()
df_master.shape
    # Out[61]: (63797, 22)

#============================================
#---- check

df_master['sample_ID'].unique()
    # Out[62]: 
    # <StringArray>
    # ['ZC69', 'ZC68', 'ZC67', 'ZC66', 'ZC65', 'ZC64', 'ZC63', 'ZC62', 'ZC61',
    #  'ZC60', 'ZC38', 'ZC37', 'ZC36', 'ZC35', 'ZC34', 'ZC33', 'ZC32', 'ZC31',
    #  'ZC30', 'ZC29', 'ZC28', 'ZC27', 'ZC26', 'ZC25', 'ZC24', 'ZC23', 'ZC22',
    #  'ZC21', 'ZC20', 'ZC19', 'ZC18', 'ZC17', 'ZC15', 'ZC14', 'ZC12', 'ZC11',
    #  'ZC10', 'ZC09', 'ZC08', 'ZC07', 'ZC05', 'ZC04']
    # Length: 42, dtype: str

# info
# type( df_master['sample_ID'].unique() )
    # Out[71]: pandas.arrays.StringArray

 # sorted() ias a built-in python function.
# Since sorted() basically just needs:
    # something iterable
    # elements that can be compared (<, >)
sorted( df_master['sample_ID'].unique() )
    # Out[72]: 
    # ['ZC04',
    #  'ZC05',
    #  'ZC07',
    #  'ZC08',
    #  'ZC09',
    #  'ZC10',
    #  'ZC11',
    #  'ZC12',
    #  'ZC14',
    #  'ZC15',
    #  'ZC17',
    #  'ZC18',
    #  'ZC19',
    #  'ZC20',
    #  'ZC21',
    #  'ZC22',
    #  'ZC23',
    #  'ZC24',
    #  'ZC25',
    #  'ZC26',
    #  'ZC27',
    #  'ZC28',
    #  'ZC29',
    #  'ZC30',
    #  'ZC31',
    #  'ZC32',
    #  'ZC33',
    #  'ZC34',
    #  'ZC35',
    #  'ZC36',
    #  'ZC37',
    #  'ZC38',
    #  'ZC60',
    #  'ZC61',
    #  'ZC62',
    #  'ZC63',
    #  'ZC64',
    #  'ZC65',
    #  'ZC66',
    #  'ZC67',
    #  'ZC68',
    #  'ZC69']
#_______________________
# note pigs ZC39 - 59 do not exist.
    # they were not implanted with transponder.


list(df_master["setup"].unique())
    # Out[63]: ['Housing', 'OF', 'Surgery', 'Stoffwechselkäfig']

list(df_master["timeline"].unique())
    # Out[64]: 
    # ['N/A',
    #  'Retraining_1',
    #  'Retraining_2',
    #  'POD 3',
    #  'POD 4',
    #  'Sacrifice',
    #  'POD 1',
    #  'Explantation',
    #  'Implantation',
    #  'TI',
    #  'minütlicher meean']

# %%% minütlicher_meean

# timeline | 'minütlicher_meean' : 
    # check its whereabouts !

df_minütlicher_meean = df_master[ df_master["timeline"] == 'minütlicher meean' ]
df_minütlicher_meean[['sample_ID','directory','Source_File']].drop_duplicates()
    # Out[16]: 
    #       sample_ID  \
    # 44177      ZC11   
    # 44311      ZC11   
    # 45751      ZC11   
    # 45756      ZC11   
    # 47196      ZC11   
    # 48636      ZC11   
    # 50076      ZC11   
    # 51516      ZC11   
    # 52956      ZC11   
    # 54396      ZC11   
    # 55836      ZC11   
    # 57166      ZC11   
    # 57355      ZC11   
    # 57366      ZC11   
    # 57367      ZC11   
    # 57368      ZC11   
    # 58808      ZC11   
    # 58809      ZC11   
    # 58822      ZC11   
    
    #                                                                                           directory  \
    # 44177  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 44311  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 45751  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 45756  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 47196  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 48636  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 50076  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 51516  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 52956  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 54396  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 55836  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 57166  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 57355  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 57366  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 57367  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 57368  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 58808  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 58809  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    # 58822  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean   
    
    #                                             Source_File  
    # 44177   zc11_0a64_rx_back-housing_2020_06_29-2.x00.xlsb  
    # 44311   zc11_0a64_rx_back-housing_2020_06_29-3.x00.xlsb  
    # 45751     zc11_0a64_rx_back-housing_2020_06_29.x00.xlsb  
    # 45756     zc11_0a64_rx_back-housing_2020_06_30.x00.xlsb  
    # 47196     zc11_0a64_rx_back-housing_2020_07_01.x00.xlsb  
    # 48636     zc11_0a64_rx_back-housing_2020_07_02.x00.xlsb  
    # 50076     zc11_0a64_rx_back-housing_2020_07_03.x00.xlsb  
    # 51516     zc11_0a64_rx_back-housing_2020_07_04.x00.xlsb  
    # 52956     zc11_0a64_rx_back-housing_2020_07_05.x00.xlsb  
    # 54396     zc11_0a64_rx_back-housing_2020_07_06.x00.xlsb  
    # 55836     zc11_0a64_rx_back-housing_2020_07_07.x00.xlsb  
    # 57166     zc11_0a64_rx_back-housing_2020_07_09.x00.xlsb  
    # 57355    zc11_0a64_rx_front-housing_2020_07_09.x00.xlsb  
    # 57366  zc11_0a64_rx_front-housing_2020_07_18-2.x00.xlsb  
    # 57367  zc11_0a64_rx_front-housing_2020_07_18-3.x00.xlsb  
    # 57368  zc11_0a64_rx_front-housing_2020_07_18-4.x00.xlsb  
    # 58808    zc11_0a64_rx_front-housing_2020_07_18.x00.xlsb  
    # 58809    zc11_0a64_rx_front-housing_2020_07_20.x00.xlsb  
    # 58822               zc11_0a64_rx_of_2020_07_18.x00.xlsb  


df_minütlicher_meean["sample_ID"].unique()
    # Out[18]: 
    # <StringArray>
    # ['ZC11']
    # Length: 1, dtype: str

df_minütlicher_meean["setup"].unique()
    # Out[17]: 
    # <StringArray>
    # ['Housing']
    # Length: 1, dtype: str

df_minütlicher_meean["directory"].unique()
    # Out[19]: 
    # <StringArray>
    # ['F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA\Housing\minütlicher meean']
    # Length: 1, dtype: str

#===========================================
#---- count

df_minütlicher_meean.shape
    # Out[20]: (14655, 22)

(df_master['sample_ID'] == 'ZC11').sum()
    # Out[22]: np.int64(15789)

15789 - 14655
    # Out[23]: 1134
# so for ZC11 : there are 1134 rows of '/hour' data, like other samples.

df_master['sample_ID'].value_counts().sort_index()
    # Out[24]: 
    # sample_ID
    # ZC04     1415
    # ZC05      985
    # ZC07     1131
    # ZC08     1020
    # ZC09      979
    # ZC10      995
    # ZC11    15789
    # ZC12      801
    # ZC14      907
    # ZC15      795
    # ZC17     2144
    # ZC18      662
    # ZC19     2369
    # ZC20     4073
    # ZC21     1941
    # ZC22     1680
    # ZC23     1166
    # ZC24     1072
    # ZC25     1159
    # ZC26     1088
    # ZC27      406
    # ZC28     1037
    # ZC29     1199
    # ZC30     1032
    # ZC31     2219
    # ZC32     2381
    # ZC33      539
    # ZC34      781
    # ZC35      958
    # ZC36      809
    # ZC37      745
    # ZC38      216
    # ZC60     1234
    # ZC61     1044
    # ZC62     1047
    # ZC63     1069
    # ZC64      944
    # ZC65      367
    # ZC66      586
    # ZC67      959
    # ZC68     1111
    # ZC69      943
    # Name: count, dtype: int64


#===========================================
#---- delete
df_master = df_master[df_master['timeline'] != 'minütlicher meean']
df_master.shape
    # Out[26]: (49142, 22)

list(df_master['timeline'].unique())
    # Out[28]: 
    # ['N/A',
    #  'Retraining_1',
    #  'Retraining_2',
    #  'POD 3',
    #  'POD 4',
    #  'Sacrifice',
    #  'POD 1',
    #  'Explantation',
    #  'Implantation',
    #  'TI']

# %%% N/A - TI

# this checks if for each pig, day-0 (TI) is registered.

(df_master['timeline'] == 'N/A').sum()
    # Out[29]: np.int64(27241)


# 1. Get a set of every single pig currently in your dataset.
set_all_pigs = set(df_master['sample_ID'].unique())

set_all_pigs
    # Out[31]: 
    # {'ZC04',
    #  'ZC05',
    #  'ZC07',
    #  'ZC08',
    #  'ZC09',
    #  'ZC10',
    #  'ZC11',
    #  'ZC12',
    #  'ZC14',
    #  'ZC15',
    #  'ZC17',
    #  'ZC18',
    #  'ZC19',
    #  'ZC20',
    #  'ZC21',
    #  'ZC22',
    #  'ZC23',
    #  'ZC24',
    #  'ZC25',
    #  'ZC26',
    #  'ZC27',
    #  'ZC28',
    #  'ZC29',
    #  'ZC30',
    #  'ZC31',
    #  'ZC32',
    #  'ZC33',
    #  'ZC34',
    #  'ZC35',
    #  'ZC36',
    #  'ZC37',
    #  'ZC38',
    #  'ZC60',
    #  'ZC61',
    #  'ZC62',
    #  'ZC63',
    #  'ZC64',
    #  'ZC65',
    #  'ZC66',
    #  'ZC67',
    #  'ZC68',
    #  'ZC69'}

len(set_all_pigs)
    # Out[82]: 42

# Filter the dataset to ONLY look at the 'TI' rows
df_TI = df_master[df_master['timeline'] == 'TI']
set_TI = set(df_TI['sample_ID'].unique())

set_TI
    # Out[35]: 
    # {'ZC07',
    #  'ZC08',
    #  'ZC09',
    #  'ZC10',
    #  'ZC11',
    #  'ZC12',
    #  'ZC14',
    #  'ZC15',
    #  'ZC17',
    #  'ZC18',
    #  'ZC19',
    #  'ZC20',
    #  'ZC21',
    #  'ZC22',
    #  'ZC23',
    #  'ZC24',
    #  'ZC25',
    #  'ZC26',
    #  'ZC28',
    #  'ZC29',
    #  'ZC30',
    #  'ZC31',
    #  'ZC32',
    #  'ZC34'}

# pigs in which they do not have the entry 'TI' under the column 'timeline'.
missing_pigs = set_all_pigs - set_TI

missing_pigs
    # Out[37]: 
    # {'ZC04',
    #  'ZC05',
    #  'ZC27',
    #  'ZC33',
    #  'ZC35',
    #  'ZC36',
    #  'ZC37',
    #  'ZC38',
    #  'ZC60',
    #  'ZC61',
    #  'ZC62',
    #  'ZC63',
    #  'ZC64',
    #  'ZC65',
    #  'ZC66',
    #  'ZC67',
    #  'ZC68',
    #  'ZC69'}

len(missing_pigs)
    # Out[39]: 18

#===========================================
#---- timeline_availability

# this is to check, for each pig of those missing 'Ti' under the column 'timeline', is at least one of other 'timelines' ( of-course not N/A ) is available !

# 1. Filter the dataset to only include rows from our 18 missing pigs
df_missing_pigs = df_master[df_master['sample_ID'].isin(missing_pigs)]

# 2. Group by the pig ID and grab all unique values in their 'timeline' column
timeline_availability = df_missing_pigs.groupby('sample_ID')['timeline'].unique()

timeline_availability
    # Out[42]: 
    # sample_ID
    # ZC04                                                         [N/A]
    # ZC05                                                         [N/A]
    # ZC27                                                         [N/A]
    # ZC33         [Explantation, Implantation, POD 3, POD 4, Sacrifice]
    # ZC35               [N/A, Retraining_1, Retraining_2, POD 1, POD 3]
    # ZC36                                                         [N/A]
    # ZC37    [N/A, Retraining_1, Retraining_2, POD 3, POD 4, Sacrifice]
    # ZC38                                                         [N/A]
    # ZC60                                                         [N/A]
    # ZC61                                                         [N/A]
    # ZC62                                                         [N/A]
    # ZC63                                                         [N/A]
    # ZC64                                                         [N/A]
    # ZC65                                                         [N/A]
    # ZC66                                                         [N/A]
    # ZC67                                                         [N/A]
    # ZC68                                                         [N/A]
    # ZC69                                                         [N/A]
    # Name: timeline, dtype: object

#===========================================
#---- Fetch date_TI

# the non-standard 'ZC6' entry was renamed to 'ZC06'
overview_3 = pd.read_excel(  r'F:\OneDrive - Uniklinik RWTH Aachen\kidney\overview_3.xlsx' , header=[0,1] , index_col=0 )
overview_3.iloc[:7,:7]
    # Out[55]: 
    #           Sample ID:          Treatment             Group:         BW Eingang  \
    #   Unnamed: 0_level_1 Unnamed: 1_level_1 Unnamed: 2_level_1 Unnamed: 3_level_1   
    # 0               ZC04            DBD-HTK                  1               25.6   
    # 1               ZC05         DBD-Ecosol                  2               20.2   
    # 2               ZC06            DBD-HTK                  1                 21   
    # 3               ZC07         DBD-Ecosol                  2               18.7   
    # 4               ZC08            DBD-HTK                  1                 19   
    # 5               ZC09         DBD-Ecosol                  2               22.9   
    # 6               ZC10            DBD-HTK                  1               22.3   
    
    #              Ear tag   Operation date Ti: Operation TI incisicion time:  
    #   Unnamed: 4_level_1   Unnamed: 5_level_1            Unnamed: 6_level_1  
    # 0                140  2020-02-03 00:00:00                      09:00:00  
    # 1                142  2020-02-25 00:00:00                      09:30:00  
    # 2                143  2020-02-25 00:00:00                      06:40:00  
    # 3                158  2020-05-25 00:00:00                      08:15:00  
    # 4                157  2020-05-25 00:00:00                      10:42:00  
    # 5                159  2020-06-15 00:00:00                      08:55:00  
    # 6                160  2020-06-15 00:00:00                      06:27:00  

overview_3.iloc[:7,[0,5]]
    # Out[58]: 
    #           Sample ID:   Operation date Ti:
    #   Unnamed: 0_level_1   Unnamed: 5_level_1
    # 0               ZC04  2020-02-03 00:00:00
    # 1               ZC05  2020-02-25 00:00:00
    # 2               ZC06  2020-02-25 00:00:00
    # 3               ZC07  2020-05-25 00:00:00
    # 4               ZC08  2020-05-25 00:00:00
    # 5               ZC09  2020-06-15 00:00:00
    # 6               ZC10  2020-06-15 00:00:00

df_date_TI = overview_3.iloc[:, [0, 5]].copy()
df_date_TI.columns = ['sample_ID', 'date_TI']

df_date_TI.shape
    # Out[63]: (82, 2)

df_date_TI[:5]
    # Out[64]: 
    #   sample_ID              date_TI
    # 0      ZC04  2020-02-03 00:00:00
    # 1      ZC05  2020-02-25 00:00:00
    # 2      ZC06  2020-02-25 00:00:00
    # 3      ZC07  2020-05-25 00:00:00
    # 4      ZC08  2020-05-25 00:00:00

df_date_TI[-5:]
    # Out[65]: 
    #    sample_ID date_TI
    # 77       NaN     NaN
    # 78       NaN     NaN
    # 79       NaN     NaN
    # 80       NaN     NaN
    # 81       NaN     NaN

# how='all' means: drop a row only if all columns in that row are NaN
df_date_TI_2 = df_date_TI.dropna(how='all')
df_date_TI_2.shape
    # Out[67]: (66, 2)

df_date_TI_2[-5:]
    # Out[68]: 
    #    sample_ID              date_TI
    # 61      ZC65  2023-07-24 00:00:00
    # 62      ZC66  2023-08-21 00:00:00
    # 63      ZC67  2023-08-29 00:00:00
    # 64      ZC68  2023-09-25 00:00:00
    # 65      ZC69  2023-10-09 00:00:00

df_date_TI_2.info()
    # <class 'pandas.DataFrame'>
    # RangeIndex: 66 entries, 0 to 65
    # Data columns (total 2 columns):
    #  #   Column     Non-Null Count  Dtype 
    # ---  ------     --------------  ----- 
    #  0   sample_ID  66 non-null     str   
    #  1   date_TI    66 non-null     object
    # dtypes: object(1), str(1)
    # memory usage: 1.2+ KB

# converting the date objects to the standard pandas datetime object.
# there are '-' values ( pigs not implanted with transponder {ZC39-59})   =>  errors='coerce'.
df_date_TI_2['date_TI'] = pd.to_datetime(df_date_TI_2['date_TI'],
                                         errors='coerce')

df_date_TI_3 = df_date_TI_2.dropna(subset=['date_TI']).reset_index(drop=True).copy()

df_date_TI_3.shape
    # Out[75]: (44, 2)

df_date_TI_3[:4]
    # Out[76]: 
    #   sample_ID    date_TI
    # 0      ZC04 2020-02-03
    # 1      ZC05 2020-02-25
    # 2      ZC06 2020-02-25
    # 3      ZC07 2020-05-25

df_date_TI_3.info()
    # <class 'pandas.DataFrame'>
    # RangeIndex: 44 entries, 0 to 43
    # Data columns (total 2 columns):
    #  #   Column     Non-Null Count  Dtype         
    # ---  ------     --------------  -----         
    #  0   sample_ID  44 non-null     str           
    #  1   date_TI    44 non-null     datetime64[us]
    # dtypes: datetime64[us](1), str(1)
    # memory usage: 836.0 bytes

df_overview_TI.rename(columns={'date_TI': 'TI_start_date_overview'}, inplace=True)

#---- save
base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
file_name = 'df_overview_TI'   # formerly : 'df_date_TI_3'
output_file = base_dir / f"{file_name}.pkl"
df_overview_TI.to_pickle(output_file)

# read
df_overview_TI = pd.read_pickle( base_dir / f"{file_name}.pkl" )

#=====================================================
#---- DIFFERENCE

# diffeence of sample_IDs in the EMKA dataset with the one in the excel sheet 'overview'.

set_date_TI_3 = set(df_date_TI_3['sample_ID'])

missing_pigs - set_date_TI_3
    # Out[96]: set()
# => all pigs' Ti-date is now known !

#=====================================================
#---- compare

# checking for consistency.
# this compares if the TI-start-dates in pigs that have them in df_master is consistent with that from the 'overview' dataset.
    # for those pigs that have TI-start-date in both dataframes.

df_master_TI = df_master[df_master['timeline'] == 'TI']

# Find the exact baseline date for each pig
# We group by the pig and find the minimum (earliest) timestamp in their 'TI' data
df_master_TI_baselines = df_master_TI.groupby('sample_ID')['timestamp'].min().reset_index()
df_master_TI_baselines.rename(columns={'timestamp': 'TI_start_date'}, inplace=True)

df_master_TI_baselines.shape
    # Out[30]: (24, 2)

df_master_TI_baselines
    # Out[31]: 
    #    sample_ID       TI_start_date
    # 0       ZC07 2020-05-25 08:49:39
    # 1       ZC08 2020-05-25 11:26:03
    # 2       ZC09 2020-06-15 09:30:41
    # 3       ZC10 2020-06-15 12:58:57
    # ...

df_overview_TI
    # Out[32]: 
    #    sample_ID    date_TI
    # 0       ZC04 2020-02-03
    # 1       ZC05 2020-02-25
    # 2       ZC06 2020-02-25
    # ...

# Merge your calculated baselines with the adjunct dataframe
# 'inner' ensures we only compare pigs that exist in BOTH lists
comparison_df = pd.merge(df_master_TI_baselines, 
                         df_overview_TI, 
                         on='sample_ID', 
                         how='inner')

comparison_df.shape
    # Out[35]: (23, 3)

comparison_df
    # Out[34]: 
    #    sample_ID       TI_start_date    date_TI
    # 0       ZC07 2020-05-25 08:49:39 2020-05-25
    # 1       ZC08 2020-05-25 11:26:03 2020-05-25
    # 2       ZC09 2020-06-15 09:30:41 2020-06-15
    # ...

comparison_df.rename(columns={'TI_start_date': 'TI_start_date_master', 
                              'date_TI':'TI_start_date_overview'
                              }, 
                     inplace=True)


# Strip the hours/minutes/seconds away from both sides to ensure a fair comparison
comparison_df['TI_start_date_master_2'] = comparison_df['TI_start_date_master'].dt.normalize()

comparison_df[['TI_start_date_master_2','TI_start_date_overview']]
    # Out[40]: 
    #    TI_start_date_master_2 TI_start_date_overview
    # 0              2020-05-25             2020-05-25
    # 1              2020-05-25             2020-05-25
    # 2              2020-06-15             2020-06-15
    # 3              2020-06-15             2020-06-15
    # 4              2020-06-29             2020-06-29
    # 5              2020-07-27             2020-07-27
    # 6              2020-08-03             2020-08-03
    # 7              2020-08-31             2020-08-31
    # 8              2020-08-31             2020-08-31
    # 9              2020-09-07             2020-09-07
    # 10             2020-09-07             2020-09-07
    # 11             2020-10-05             2020-10-05
    # 12             2020-10-05             2020-10-05
    # 13             2020-11-02             2020-11-02
    # 14             2020-11-02             2020-11-02
    # 15             2020-11-09             2020-11-09
    # 16             2020-11-09             2020-11-09
    # 17             2021-01-18             2021-01-18
    # 18             2021-01-25             2021-01-25
    # 19             2021-01-25             2021-01-25
    # 20             2021-02-22             2021-02-22
    # 21             2021-02-22             2021-02-22
    # 22             2021-03-01             2021-03-01

comparison_df['is_consistent'] = comparison_df['TI_start_date_master_2'] == comparison_df['TI_start_date_overview']

comparison_df['is_consistent'].unique()
    # Out[43]: array([ True])
# all are 'True'.

# df_TI_ID_date = df_TI[['sample_ID','timeline','cpu-date']].drop_duplicates()

#=====================================================
#---- start-dates _ all

# get the start-dates fro all pigs in a separate dataframe.

# pigs that have TI-dates in the master dataframe, & their TI-dates.
df_master_TI_baselines
    # Out[31]: 
    #    sample_ID       TI_start_date
    # 0       ZC07 2020-05-25 08:49:39
    # 1       ZC08 2020-05-25 11:26:03
    # 2       ZC09 2020-06-15 09:30:41
    # 3       ZC10 2020-06-15 12:58:57
    # ...

df_master_TI_baselines['TI_start_date_master'] = df_master_TI_baselines['TI_start_date'].dt.normalize()
df_master_TI_baselines.drop(columns=['TI_start_date'], inplace=True)
df_master_TI_baselines.head()
    # Out[58]: 
    #   sample_ID TI_start_date_master
    # 0      ZC07           2020-05-25
    # 1      ZC08           2020-05-25
    # 2      ZC09           2020-06-15
    # 3      ZC10           2020-06-15
    # 4      ZC11           2020-06-29

set_missing_pigs = {'ZC04',
                    'ZC05',
                    'ZC27',
                    'ZC33',
                    'ZC35',
                    'ZC36',
                    'ZC37',
                    'ZC38',
                    'ZC60',
                    'ZC61',
                    'ZC62',
                    'ZC63',
                    'ZC64',
                    'ZC65',
                    'ZC66',
                    'ZC67',
                    'ZC68',
                    'ZC69'}

# the adjunct dataframe containing TI dates from the pigs missing it.
df_overview_TI.head()
    # Out[51]: 
    #   sample_ID TI_start_date_overview
    # 0      ZC04             2020-02-03
    # 1      ZC05             2020-02-25
    # 2      ZC06             2020-02-25
    # 3      ZC07             2020-05-25
    # 4      ZC08             2020-05-25


# 1. Isolate ONLY the missing pigs from the adjunct dataframe
# This ensures pigs like ZC07 and ZC08 (which are in both) don't get duplicated
df_missing_dates = df_overview_TI[df_overview_TI['sample_ID'].isin(set_missing_pigs)].copy()

# 2. Standardize the date column names
# pandas pd.concat() needs the columns to have the exact same name to stack them perfectly
df_master_TI_baselines = df_master_TI_baselines.rename(columns={'TI_start_date_master': 'TI_start_date'})
df_missing_dates = df_missing_dates.rename(columns={'TI_start_date_overview': 'TI_start_date'})

# 3. Stack (concatenate) the two dataframes vertically
df_TI_start_dates_all = pd.concat([df_master_TI_baselines, df_missing_dates], ignore_index=True)

# 4. Sort alphabetically by sample_ID and clean up the index for a professional look
df_TI_start_dates_all = df_TI_start_dates_all.sort_values(by='sample_ID').reset_index(drop=True)

df_TI_start_dates_all.shape
    # Out[64]: (42, 2)

df_TI_start_dates_all
    # Out[65]: 
    #    sample_ID TI_start_date
    # 0       ZC04    2020-02-03
    # 1       ZC05    2020-02-25
    # 2       ZC07    2020-05-25
    # 3       ZC08    2020-05-25
    # 4       ZC09    2020-06-15
    # 5       ZC10    2020-06-15
    # 6       ZC11    2020-06-29
    # 7       ZC12    2020-06-29
    # 8       ZC14    2020-07-27
    # 9       ZC15    2020-08-03
    # 10      ZC17    2020-08-31
    # 11      ZC18    2020-08-31
    # 12      ZC19    2020-09-07
    # 13      ZC20    2020-09-07
    # 14      ZC21    2020-10-05
    # 15      ZC22    2020-10-05
    # 16      ZC23    2020-11-02
    # 17      ZC24    2020-11-02
    # 18      ZC25    2020-11-09
    # 19      ZC26    2020-11-09
    # 20      ZC27    2021-01-18
    # 21      ZC28    2021-01-18
    # 22      ZC29    2021-01-25
    # 23      ZC30    2021-01-25
    # 24      ZC31    2021-02-22
    # 25      ZC32    2021-02-22
    # 26      ZC33    2021-03-01
    # 27      ZC34    2021-03-01
    # 28      ZC35    2021-03-29
    # 29      ZC36    2021-03-29
    # 30      ZC37    2021-04-19
    # 31      ZC38    2021-04-19
    # 32      ZC60    2023-05-22
    # 33      ZC61    2023-06-15
    # 34      ZC62    2023-06-21
    # 35      ZC63    2023-07-10
    # 36      ZC64    2023-07-17
    # 37      ZC65    2023-07-24
    # 38      ZC66    2023-08-21
    # 39      ZC67    2023-08-29
    # 40      ZC68    2023-09-25
    # 41      ZC69    2023-10-09

# check
set_df_TI_start_dates_all = set(df_TI_start_dates_all['sample_ID'].unique())

# check if all pigs are covered.
set_all_pigs.symmetric_difference(set_df_TI_start_dates_all)
    # Out[69]: set()
# => all pigs in df_master now have a TI-start date available.

# save
base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
file_name = 'df_TI_start_dates_all'   # formerly : 'df_date_TI_3'
output_file = base_dir / f"{file_name}.pkl"
df_TI_start_dates_all.to_pickle(output_file)

# read
df_TI_start_dates_all = pd.read_pickle(output_file)


# %%%% timetag

# --- STEP 1: MERGE TI DATES INTO MASTER ---
# Merge the clean dates into the master dataframe
df_master = df_master.merge(df_TI_start_dates_all, on='sample_ID', how='left')

# --- STEP 2: CALCULATE ELAPSED CALENDAR DAYS ---
# We use .dt.normalize() on the timestamp to strip away the hours/minutes.
# This ensures we are counting strict calendar days, not 24-hour periods!
day_of_recording = df_master['timestamp'].dt.normalize()
day_of_TI = df_master['TI_start_date'].dt.normalize()

# Calculate the difference and extract the integer number of days
df_master['days_since_TI'] = (day_of_recording - day_of_TI).dt.days

#================================
# --- STEP 3: APPLY THE NAMING CONVENTION ---
def map_timetag(days):
    # Handle any rows where dates might be missing (results in NaN)
    if pd.isna(days):
        return 'Unknown'
        
    days = int(days)
    
    if days == 0:
        return 'TI'
    elif 1 <= days <= 11:
        return f'TI_{days}'
    elif days == 12:
        return 'Retraining_1'
    elif days == 13:
        return 'Retraining_2'
    elif days == 14:
        return 'Explantation'
    elif days == 15:
        return 'Implantation'
    elif 16 <= days <= 21:
        return f'POD_{days - 15}' # e.g., Day 16 - 15 = POD_1
    elif days == 22:
        return 'Sacrifice'
    #================================
    # Failsafe
    elif days > 22:
        return f'Post_Sacrifice_{days-22}' # Failsafe for extra data
    elif days < 0:
        return f'Pre_TI_{days}' # Failsafe for data recorded before TI
    else:
        return 'Unknown'
#================================

# Apply the function to create the new column
df_master['timetag'] = df_master['days_since_TI'].apply(map_timetag)

#---- VERIFICATION ---
# Display a randomized sample of 10 rows to see the different tags in action
preview_cols = ['sample_ID', 'timestamp', 'TI_start_date', 'days_since_TI', 'timetag']
# if not wrapping it in : print() : the output will be jammed !
print(df_master.dropna(subset=['days_since_TI'])[preview_cols].sample(10).to_string(index=False))
    # sample_ID           timestamp TI_start_date  days_since_TI      timetag
    #      ZC09 2020-07-05 01:24:11    2020-06-15             20        POD_5
    #      ZC65 2023-08-02 17:59:40    2023-07-24              9         TI_9
    #      ZC07 2020-06-08 11:31:07    2020-05-25             14 Explantation
    #      ZC31 2021-03-07 23:52:54    2021-02-22             13 Retraining_2
    #      ZC60 2023-06-05 10:43:03    2023-05-22             14 Explantation
    #      ZC17 2020-09-15 11:29:26    2020-08-31             15 Implantation
    #      ZC65 2023-08-10 12:56:30    2023-07-24             17        POD_2
    #      ZC08 2020-06-08 14:49:28    2020-05-25             14 Explantation
    #      ZC29 2021-02-16 12:57:51    2021-01-25             22    Sacrifice
    #      ZC11 2020-07-13 07:49:42    2020-06-29             14 Explantation

# another run
    # sample_ID           timestamp TI_start_date  days_since_TI      timetag
    #      ZC32 2021-03-08 06:11:54    2021-02-22             14 Explantation
    #      ZC60 2023-06-05 10:04:03    2023-05-22             14 Explantation
    #      ZC21 2020-10-18 03:08:51    2020-10-05             13 Retraining_2
    #      ZC07 2020-06-08 10:21:06    2020-05-25             14 Explantation
    #      ZC33 2021-03-15 11:18:27    2021-03-01             14 Explantation
    #      ZC23 2020-11-24 21:50:35    2020-11-02             22    Sacrifice
    #      ZC31 2021-03-09 10:32:43    2021-02-22             15 Implantation
    #      ZC21 2020-10-22 21:10:48    2020-10-05             17        POD_2
    #      ZC20 2020-09-10 21:01:30    2020-09-07              3         TI_3
    #      ZC37 2021-05-03 09:41:09    2021-04-19             14 Explantation

list(df_master['timetag'].unique())
    # Out[39]: 
    # ['TI_1',
    #  'TI_2',
    #  'TI_3',
    #  'TI_4',
    #  'TI_5',
    #  'TI_6',
    #  'TI_7',
    #  'TI_8',
    #  'TI_9',
    #  'TI_10',
    #  'TI_11',
    #  'Retraining_1',
    #  'Retraining_2',
    #  'Explantation',
    #  'Implantation',
    #  'POD_1',
    #  'POD_2',
    #  'TI',
    #  'POD_3',
    #  'POD_4',
    #  'POD_5',
    #  'POD_6',
    #  'Sacrifice',
    #  'Post_Sacrifice_13',
    #  'Post_Sacrifice_27',
    #  'Post_Sacrifice_34',
    #  'Post_Sacrifice_1']

# Show a quick summary of how many rows exist for each timetag
df_master['timetag'].value_counts()
    # Out[40]: 
    # timetag
    # Explantation         11795
    # Implantation          9100
    # TI                    4064
    # Retraining_2          3061
    # POD_1                 2618
    # Sacrifice             1971
    # POD_2                 1880
    # Retraining_1          1245
    # POD_3                 1237
    # POD_4                  966
    # TI_6                   944
    # TI_5                   943
    # TI_1                   908
    # TI_4                   876
    # TI_2                   849
    # TI_3                   840
    # TI_7                   831
    # TI_11                  809
    # TI_10                  783
    # TI_9                   765
    # TI_8                   746
    # POD_6                  730
    # POD_5                  624
    # Post_Sacrifice_27      423
    # Post_Sacrifice_13       82
    # Post_Sacrifice_34       36
    # Post_Sacrifice_1        16
    # Name: count, dtype: int64


# %% timestamp

# check the original time data.

df_master[['cpu-date', 'cpu-time']].head()
    # Out[44]: 
    #     cpu-date  cpu-time
    # 0  10-Oct-23  12:01:10
    # 1  10-Oct-23  13:01:10
    # 2  10-Oct-23  14:01:09
    # 3  10-Oct-23  15:01:10
    # 4  10-Oct-23  16:01:10

df_master[['cpu-date', 'cpu-time']][1000:1005]
    # Out[45]: 
    #        cpu-date  cpu-time
    # 1000  12-Oct-23  21:02:21
    # 1001  12-Oct-23  22:02:21
    # 1002  12-Oct-23  23:02:22
    # 1003  13-Oct-23  00:02:21
    # 1004  13-Oct-23  01:02:21

df_master[['cpu-date', 'cpu-time']][10000:10005]
    # Out[46]: 
    #         cpu-date  cpu-time
    # 10000  04-May-21  09:56:56
    # 10001  04-May-21  09:57:56
    # 10002  04-May-21  09:58:55
    # 10003  04-May-21  09:59:56
    # 10004  04-May-21  10:00:55

df_master[['cpu-date', 'cpu-time']][20000:20005]
    # Out[47]: 
    #         cpu-date  cpu-time
    # 20000  09-Feb-21  12:16:42
    # 20001  09-Feb-21  12:17:42
    # 20002  09-Feb-21  12:18:43
    # 20003  09-Feb-21  12:19:42
    # 20004  09-Feb-21  12:20:43


df_master[['cpu-date', 'cpu-time']][30000:30005]
    # Out[48]: 
    #         cpu-date  cpu-time
    # 30000  19-Sep-20  03:40:25
    # 30001  19-Sep-20  04:40:24
    # 30002  19-Sep-20  05:40:25
    # 30003  19-Sep-20  06:40:25
    # 30004  19-Sep-20  07:40:24

df_master[['cpu-date', 'cpu-time']][40000:40005]
    # Out[49]: 
    #         cpu-date  cpu-time
    # 40000  03-Aug-20  11:14:48
    # 40001  03-Aug-20  11:15:49
    # 40002  03-Aug-20  11:16:49
    # 40003  03-Aug-20  11:17:48
    # 40004  03-Aug-20  11:18:49


df_master[['cpu-date', 'cpu-time']][50000:50005]
    # Out[50]: 
    #         cpu-date  cpu-time
    # 50000  03-Jul-20  19:27:05
    # 50001  03-Jul-20  19:28:06
    # 50002  03-Jul-20  19:29:05
    # 50003  03-Jul-20  19:30:06
    # 50004  03-Jul-20  19:31:06

df_master[['cpu-date', 'cpu-time']][60000:60005]
    # Out[51]: 
    #         cpu-date  cpu-time
    # 60000  22-Jun-20  19:17:34
    # 60001  22-Jun-20  20:17:33
    # 60002  22-Jun-20  21:17:33
    # 60003  22-Jun-20  22:17:34
    # 60004  22-Jun-20  23:17:33

# %%% convert

# convert it to a a pandas datetime object.

# 1. Combine the date and time strings into a single column with a space in between
combined_datetime_str = df_master['cpu-date'] + ' ' + df_master['cpu-time']

combined_datetime_str.head()
    # Out[53]: 
    # 0    10-Oct-23 12:01:10
    # 1    10-Oct-23 13:01:10
    # 2    10-Oct-23 14:01:09
    # 3    10-Oct-23 15:01:10
    # 4    10-Oct-23 16:01:10
    # dtype: str

# 2. Convert the combined string into a true pandas datetime object
    # The format '%d-%b-%y %H:%M:%S' tells pandas exactly how to read your specific layout:
    # %d = 2-digit day (10)
    # %b = Abbreviated month (Oct)
    # %y = 2-digit year (23)
    # %H:%M:%S = Hours:Minutes:Seconds
df_master['timestamp'] = pd.to_datetime(combined_datetime_str, format='%d-%b-%y %H:%M:%S')

# 3. Verify the conversion worked by checking the data types
print(df_master[['cpu-date', 'cpu-time', 'timestamp']].head())
    #     cpu-date  cpu-time           timestamp
    # 0  10-Oct-23  12:01:10 2023-10-10 12:01:10
    # 1  10-Oct-23  13:01:10 2023-10-10 13:01:10
    # 2  10-Oct-23  14:01:09 2023-10-10 14:01:09
    # 3  10-Oct-23  15:01:10 2023-10-10 15:01:10
    # 4  10-Oct-23  16:01:10 2023-10-10 16:01:10


# Data Types:
print(df_master[['cpu-date', 'cpu-time', 'timestamp']].dtypes)
    # cpu-date                str
    # cpu-time                str
    # timestamp    datetime64[us]
    # dtype: object

# %%% save

base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
output_file = base_dir / "Master_Telemetry_Dataset_2.pkl"
df_master.to_pickle(output_file)


base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
source_file = base_dir / "Master_Telemetry_Dataset_2.pkl"
df_master = pd.read_pickle(source_file)

# %% reorder columns

ID_columns = ['sample_ID', 'setup', 'timetag', 'timeline', 'timestamp']

cols = ID_columns + [col for col in df_master.columns
                     if col not in ID_columns
                     ]

df_master = df_master[cols]


# %% filter

# 1. Create the 'Housing' subset
# We use .copy() to ensure this new dataframe is completely independent of the master dataset
df_housing = df_master[df_master['setup'] == 'Housing'].copy()

# 2. Verify the extraction
print(f"Total rows in Master Dataset: {len(df_master)}")
print(f"Total rows in Housing Subset: {len(df_housing)}")

# Optional: Double-check that ONLY 'Housing' exists in this new dataframe
print("\nUnique setups in new dataframe:")
print(df_housing['setup'].unique())

# %%% out

    # Total rows in Master Dataset: 66049
    # Total rows in Housing Subset: 32464
    
    # Unique setups in new dataframe:
    # <StringArray>
    # ['Housing']
    # Length: 1, dtype: str

# %%%'

samples = list(df_housing['sample_ID'].unique())

len(samples)
    # Out[42]: 43

samples
    # Out[39]: 
    # ['ZC69',
    #  'ZC68',
    #  'ZC67',
    #  'ZC66',
    #  'ZC65',
    #  'ZC64',
    #  'ZC63',
    #  'ZC62',
    #  'ZC61',
    #  'ZC60',
    #  'ZC38',
    #  'ZC37',
    #  'ZC36',
    #  'ZC35',
    #  'ZC34',
    #  'ZC32',
    #  'ZC31',
    #  'ZC30',
    #  'ZC29',
    #  'ZC28',
    #  'ZC27',
    #  'ZC26',
    #  'ZC25',
    #  'ZC24',
    #  'ZC23',
    #  'ZC22',
    #  'ZC21',
    #  'ZC20',
    #  'ZC19',
    #  'ZC18',
    #  'ZC17',
    #  'ZC15',
    #  'ZC14',
    #  'ZC13',
    #  'ZC12',
    #  'ZC11',
    #  'ZC10',
    #  'ZC09',
    #  'ZC08',
    #  'ZC07',
    #  'ZC06',
    #  'ZC05',
    #  'ZC04']

# %% explore

df_master.info()
    # <class 'pandas.DataFrame'>
    # Index: 38069 entries, 0 to 41991
    # Data columns (total 26 columns):
    #  #   Column                   Non-Null Count  Dtype          
    # ---  ------                   --------------  -----          
    #  0   sample_ID                38069 non-null  str            
    #  1   setup                    38069 non-null  str            
    #  2   timetag                  38069 non-null  str            
    #  3   timeline                 38069 non-null  str            
    #  4   timestamp                38069 non-null  datetime64[us] 
    #  5   cpu-date                 38069 non-null  str            
    #  6   cpu-time                 38069 non-null  str            
    #  7   period-time              37479 non-null  str            
    #  8   mark-label               36435 non-null  object         
    #  9   step-index               38069 non-null  object         
    #  10  BB__aver_(ms)            37964 non-null  object         
    #  11  HR__aver_(bpm)           37964 non-null  object         
    #  12  DBP__aver_(mmHg)         37964 non-null  object         
    #  13  SBP__aver_(mmHg)         37964 non-null  object         
    #  14  MBP__aver_(mmHg)         37964 non-null  object         
    #  15  aver__aver_(°C)          38068 non-null  object         
    #  16  aver__aver_(%)           34425 non-null  object         
    #  17  HR__aver_(bpm)_1         38069 non-null  object         
    #  18  Source_File              38069 non-null  str            
    #  19  directory                38069 non-null  str            
    #  20  aver__aver_(g)           3644 non-null   object         
    #  21  _2                       219 non-null    object         
    #  22  Abweichung in% HR vs HR  161 non-null    object         
    #  23  TI_start_date            38069 non-null  datetime64[us] 
    #  24  days_since_TI            38069 non-null  int64          
    #  25  time_diff                38027 non-null  timedelta64[us]
    # dtypes: datetime64[us](2), int64(1), object(13), str(9), timedelta64[us](1)
    # memory usage: 7.8+ MB


df_master.iloc[:4,:6]
    # Out[14]: 
    #   sample_ID    setup timetag timeline           timestamp   cpu-date
    # 0      ZC04  Housing    TI_4      N/A 2020-02-07 15:53:53  07-Feb-20
    # 1      ZC04  Housing    TI_4      N/A 2020-02-07 15:53:53  07-Feb-20
    # 2      ZC04  Housing    TI_4      N/A 2020-02-07 15:53:53  07-Feb-20
    # 3      ZC04  Housing    TI_4      N/A 2020-02-07 15:53:53  07-Feb-20


list(df_master.columns)
    # Out[65]: 
    # ['sample_ID',
    #  'setup',
    #  'timetag',
    #  'timeline',
    #  'timestamp',
    #  'cpu-date',
    #  'cpu-time',
    #  'period-time',
    #  'mark-label',
    #  'step-index',
    #  'BB__aver_(ms)',
    #  'HR__aver_(bpm)',
    #  'DBP__aver_(mmHg)',
    #  'SBP__aver_(mmHg)',
    #  'MBP__aver_(mmHg)',
    #  'aver__aver_(°C)',
    #  'aver__aver_(%)',
    #  'HR__aver_(bpm)_1',
    #  'Source_File',
    #  'directory',
    #  'aver__aver_(g)',
    #  '_2',
    #  'Abweichung in% HR vs HR',
    #  'TI_start_date',
    #  'days_since_TI']

#===============================
#---- N/A timeline

# do not wonder why some values under the column 'timeline' is N/A.
    # while the corresponding values under the column 'timetag' exist.
    # not every 'timeline' was input by the data acquisitors.

df_master['timeline'].unique()
    # Out[81]: 
    # <StringArray>
    # ['N/A', 'TI', 'Retraining_1', 'Retraining_2', 'Explantation', 'Implantation', 'POD 1', 'POD 3', 'POD 4', 'Sacrifice']
    # Length: 10, dtype: str


df_master[['timetag','timeline']][:4]
    # Out[82]: 
    #   timetag timeline
    # 0    TI_4      N/A
    # 1    TI_4      N/A
    # 2    TI_4      N/A
    # 3    TI_4      N/A

# %% duplicate _ HR__aver_(bpm)

# removing duplicate rows : from duplicate excel files only containing data from 'HR__aver_(bpm)'.
    # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\screenshot\duplicate

df_master.shape
    # Out[54]: (49142, 25)

df_master[['BB__aver_(ms)','HR__aver_(bpm)', 'DBP__aver_(mmHg)','SBP__aver_(mmHg)','MBP__aver_(mmHg)', 'aver__aver_(°C)']][29280:29310]
    # Out[50]: 
    #       BB__aver_(ms) HR__aver_(bpm) DBP__aver_(mmHg) SBP__aver_(mmHg)  \
    # 29280             0              0                0                0   
    # 29281             0              0                0                0   
    # 29282             0              0                0                0   
    # 29283             0              0                0                0   
    # 29284             0              0                0                0   
    # 29285             0              0                0                0   
    # 29286             0              0                0                0   
    # 29287             0              0                0                0   
    # 29288             0              0                0                0   
    # 29289             0              0                0                0   
    # 29290           562        106.763            74.85           117.95   
    # 29291             0              0                0                0   
    # 29292             0              0                0                0   
    # 29293           596        100.671             73.9            118.1   
    # 29294       598.857        100.217           75.479          119.887   
    # 29295       587.185        102.229           76.672          120.296   
    # 29296           NaN        165.428              NaN              NaN   
    # 29297           NaN         252.84              NaN              NaN   
    # 29298           NaN        248.695              NaN              NaN   
    # 29299           NaN        243.782              NaN              NaN   
    # 29300           NaN        239.963              NaN              NaN   
    # 29301           NaN        231.924              NaN              NaN   
    # 29302           NaN        196.038              NaN              NaN   
    # 29303           NaN        159.949              NaN              NaN   
    # 29304           NaN        169.062              NaN              NaN   
    # 29305           NaN        217.474              NaN              NaN   
    # 29306           NaN        207.589              NaN              NaN   
    # 29307           NaN        186.549              NaN              NaN   
    # 29308           NaN        147.362              NaN              NaN   
    # 29309           NaN        197.368              NaN              NaN   
    
    #       MBP__aver_(mmHg) aver__aver_(°C)  
    # 29280                0              37  
    # 29281                0               0  
    # 29282                0          37.031  
    # 29283                0               0  
    # 29284                0               0  
    # 29285                0               0  
    # 29286                0               0  
    # 29287                0               0  
    # 29288                0               0  
    # 29289                0               0  
    # 29290           93.515          37.146  
    # 29291                0               0  
    # 29292                0               0  
    # 29293           93.277          37.219  
    # 29294           94.664          37.254  
    # 29295           95.802          37.284  
    # 29296              NaN             NaN  
    # 29297              NaN             NaN  
    # 29298              NaN             NaN  
    # 29299              NaN             NaN  
    # 29300              NaN             NaN  
    # 29301              NaN             NaN  
    # 29302              NaN             NaN  
    # 29303              NaN             NaN  
    # 29304              NaN             NaN  
    # 29305              NaN             NaN  
    # 29306              NaN             NaN  
    # 29307              NaN             NaN  
    # 29308              NaN             NaN  
    # 29309              NaN             NaN  


# 1. Define the specific columns that act as your "junk filter"
columns_to_check = [
    'BB__aver_(ms)', 
    'DBP__aver_(mmHg)',
    'SBP__aver_(mmHg)',
    'MBP__aver_(mmHg)', 
    'aver__aver_(°C)'
]

# 2. Drop the rows where ALL of those specific columns are NaN
# how='all' is the magic word here. It ensures it only drops rows missing everything in the subset.
df_master = df_master.dropna(subset=columns_to_check, how='all').copy()

df_master.shape
    # Out[56]: (41992, 25)

# Optional: Reset the index so your row numbers are clean and continuous again
df_master = df_master.reset_index(drop=True)

# %% 0

####################################################
#---- delete
# dete rows awith all value of parameter columns = 0.

# deleting '0' value rows.

df_master.shape
    # Out[33]: (41992, 26)

parameter_columns = [
    'BB__aver_(ms)',
    'HR__aver_(bpm)',
    'DBP__aver_(mmHg)',
    'SBP__aver_(mmHg)',
    'MBP__aver_(mmHg)',
    'aver__aver_(°C)',
    'aver__aver_(%)',
    'HR__aver_(bpm)_1',
    'aver__aver_(g)',   # is this important ?
]


# .all(axis=1) : checks whether all columns in that row are zero.
rows_to_delete = (df_master[parameter_columns] == 0).all(axis=1)

rows_to_delete.shape
    # Out[36]: (41992,)

rows_to_delete.sum()
    # Out[37]: np.int64(0)

# => non of the rows have all parameter_columns values = 0

#==============================

# here, this column is ignored :'aver__aver_(g)'.

parameter_columns_2 = [
    'BB__aver_(ms)',
    'HR__aver_(bpm)',
    'DBP__aver_(mmHg)',
    'SBP__aver_(mmHg)',
    'MBP__aver_(mmHg)',
    'aver__aver_(°C)',
    'aver__aver_(%)',
    'HR__aver_(bpm)_1',
]

rows_to_delete_2 = (df_master[parameter_columns_2] == 0).all(axis=1)

rows_to_delete_2.sum()
    # Out[39]: np.int64(2640)

df_master = df_master[~rows_to_delete_2]

df_master.shape
    # Out[41]: (39352, 26)

####################################################
#---- 0 => NaN

# Convert true 0s to NaNs for the physiological columns
for col in parameter_columns_2 :
    # Replace 0 with NaN
    df_master[col] = df_master[col].replace(0, pd.NA)
    


# %% time-diff

# exploring the time-diffs throught the whole dataset.
    # how many are per-hour ?
    # how many per-minute ?
    # ...

# 1. Sort the data to guarantee it is in perfect chronological order for each pig
df_master = df_master.sort_values(by=['sample_ID', 'timestamp']).reset_index(drop=True)

# 2. Calculate the exact time difference between consecutive rows, strictly within each pig's data
    # The .diff() operation returns a Series with the same index as the original DataFrame, not a shorter Series. For each group:
        # First row of each group gets NaN (since no previous row)
        # Subsequent rows get the differences
        # Total length = original length
    # Index alignment: 
        # When assigning back to df_master['time_diff'], pandas aligns by the index, not by position. 
        # So the NaN values and differences are placed in their correct original rows.
df_master['time_diff'] = df_master.groupby('sample_ID')['timestamp'].diff()

# 3. Count how often each specific time gap occurs
resolution_counts = df_master['time_diff'].value_counts()

resolution_counts.shape
    # Out[76]: (744,)

resolution_counts[:4]
    # Out[77]: 
    # time_diff
    # 0 days 00:00:59    8973
    # 0 days 00:01:01    8944
    # 0 days 00:01:00    5780
    # 0 days 00:59:59    5075
    # Name: count, dtype: int64

# Round the time differences to the nearest minute ('1min' or 'min')
rounded_diffs = df_master['time_diff'].dt.round('1min')

# the first row is NaT : diff can not be calculated from any previous value.
rounded_diffs[:10]
    # Out[80]: 
    # 0               NaT
    # 1   0 days 00:00:00
    # 2   0 days 00:00:00
    # 3   0 days 00:00:00
    # 4   0 days 01:00:00
    # 5   0 days 00:00:00
    # 6   0 days 00:00:00
    # 7   0 days 00:00:00
    # 8   0 days 01:00:00
    # 9   0 days 00:00:00
    # Name: time_diff, dtype: timedelta64[us]


rounded_diffs.value_counts().shape
    # Out[82]: (192,)

rounded_diffs.value_counts()
    # Out[79]: 
    # time_diff
    # 0 days 00:01:00    23912
    # 0 days 01:00:00    15374
    # 0 days 00:00:00     1634
    # 0 days 01:01:00      212
    # 0 days 01:02:00       38
     
    # 0 days 01:21:00        1
    # 0 days 17:52:00        1
    # 0 days 03:01:00        1
    # 0 days 23:59:00        1
    # 0 days 02:01:00        1
    # Name: count, Length: 192, dtype: int64

rounded_diffs.value_counts()[:10]
    # Out[81]: 
    # time_diff
    # 0 days 00:01:00    23912
    # 0 days 01:00:00    15374
    # 0 days 00:00:00     1634
    # 0 days 01:01:00      212
    # 0 days 01:02:00       38
    # 0 days 00:13:00       33
    # 0 days 00:08:00       32
    # 0 days 00:47:00       29
    # 0 days 00:52:00       29
    # 0 days 00:25:00       24
    # Name: count, dtype: int64

#=======================================================================
#---- inspect duplicate timestamps

# Extracting simultaneous records.

# Find all rows that share the same sample_ID and timestamp.
# keep=False : show us the original row AND the duplicate row.
mask_simultaneous = df_master.duplicated(subset=['sample_ID', 
                                                 'timestamp'], 
                                         keep=False)

# Create a dataframe of just these anomalies, sorted so the matching pairs are stacked right next to each other
df_simultaneous = df_master[mask_simultaneous].sort_values(by=['sample_ID', 'timestamp'])

df_simultaneous.shape
    # Out[45]: (2298, 26)  :  after removing rows with all parameter-columns having '0' values.
    # Out[17]: (2532, 26)

# Choose a few relevant columns to inspect to see if the data values are identical or different
inspect_cols = [
    'sample_ID', 
    'timestamp', 
    'Source_File', 
    'HR__aver_(bpm)', 
    'SBP__aver_(mmHg)',
    'aver__aver_(°C)'
]

# PREVIEW OF ZERO-DIFFERENCE PAIRS

# more examples  =>  explore_emka.py  |   duplicate timestamps

df_simultaneous[inspect_cols].head(10)
    # Out[31]: 
    #   sample_ID           timestamp                                       Source_File HR__aver_(bpm) SBP__aver_(mmHg)
    # 0      ZC04 2020-02-07 15:53:53  zc04_0a0f_rx_front-housing_2020_02_07-9.x00.xlsb        132.788          125.356
    # 1      ZC04 2020-02-07 15:53:53  zc04_0a0f_rx_front-housing_2020_02_07-9.x01.xlsb        132.788          125.356
    # 2      ZC04 2020-02-07 15:53:53  zc04_0a0f_rx_front-housing_2020_02_07-9.x02.xlsb        132.788          125.356
    # 3      ZC04 2020-02-07 15:53:53  zc04_0a0f_rx_front-housing_2020_02_07-9.x03.xlsb        132.788          125.356
    # 4      ZC04 2020-02-07 16:53:52  zc04_0a0f_rx_front-housing_2020_02_07-9.x00.xlsb        127.771          113.238
    # 5      ZC04 2020-02-07 16:53:52  zc04_0a0f_rx_front-housing_2020_02_07-9.x01.xlsb        127.771          113.238
    # 6      ZC04 2020-02-07 16:53:52  zc04_0a0f_rx_front-housing_2020_02_07-9.x02.xlsb        127.771          113.238
    # 7      ZC04 2020-02-07 16:53:52  zc04_0a0f_rx_front-housing_2020_02_07-9.x03.xlsb        127.771          113.238
    # 8      ZC04 2020-02-07 17:53:53  zc04_0a0f_rx_front-housing_2020_02_07-9.x00.xlsb        113.823          137.825
    # 9      ZC04 2020-02-07 17:53:53  zc04_0a0f_rx_front-housing_2020_02_07-9.x01.xlsb        113.823          137.825


df_simultaneous[inspect_cols][-10:]
    # after removing rows with all parameter-columns having '0' values.
    # Out[47]: 
    #       sample_ID           timestamp                          Source_File HR__aver_(bpm) SBP__aver_(mmHg)
    # 32608      ZC38 2021-04-19 12:37:16  zc38_0a65_2021_april_19_01.x01.xlsb        115.088          117.398
    # 32609      ZC38 2021-04-19 12:37:16  zc38_0a65_2021_april_19_01.x02.xlsb        115.088          117.398
    # 32610      ZC38 2021-04-19 12:38:17  zc38_0a65_2021_april_19_01.x01.xlsb        116.262          115.048
    # 32611      ZC38 2021-04-19 12:38:17  zc38_0a65_2021_april_19_01.x02.xlsb        116.262          115.048
    # 32612      ZC38 2021-04-19 12:39:16  zc38_0a65_2021_april_19_01.x01.xlsb        118.724          115.388
    # 32613      ZC38 2021-04-19 12:39:16  zc38_0a65_2021_april_19_01.x02.xlsb        118.724          115.388
    # 32614      ZC38 2021-04-19 12:40:16  zc38_0a65_2021_april_19_01.x01.xlsb          120.4          116.829
    # 32615      ZC38 2021-04-19 12:40:16  zc38_0a65_2021_april_19_01.x02.xlsb          120.4          116.829
    # 32616      ZC38 2021-04-19 12:41:17  zc38_0a65_2021_april_19_01.x01.xlsb        122.454          115.195
    # 32617      ZC38 2021-04-19 12:41:17  zc38_0a65_2021_april_19_01.x02.xlsb        122.454          115.195

    #===================================================================================================================

    # before removing rows with all parameter-columns having '0' values.
        # last 6 rows :
            # the duplicates are those with values 0 under the HR & Bp columns !
    # Out[25]: 
    #       sample_ID           timestamp                          Source_File HR__aver_(bpm) SBP__aver_(mmHg)
    # 32614      ZC38 2021-04-19 12:40:16  zc38_0a65_2021_april_19_01.x01.xlsb          120.4          116.829
    # 32615      ZC38 2021-04-19 12:40:16  zc38_0a65_2021_april_19_01.x02.xlsb          120.4          116.829
    # 32616      ZC38 2021-04-19 12:41:17  zc38_0a65_2021_april_19_01.x01.xlsb        122.454          115.195
    # 32617      ZC38 2021-04-19 12:41:17  zc38_0a65_2021_april_19_01.x02.xlsb        122.454          115.195
    # 33524      ZC60 2023-06-05 14:39:03        0bb9_2023_june_05_03.x00.xlsb              0                0
    # 33525      ZC60 2023-06-05 14:39:03   0bb9_0bb9_2023_june_05_01.x00.xlsb         71.155           76.324
    # 33707      ZC60 2023-06-05 17:39:03        0bb9_2023_june_05_03.x00.xlsb              0                0
    # 33708      ZC60 2023-06-05 17:39:03   0bb9_0bb9_2023_june_05_01.x00.xlsb        112.103           85.013
    # 39886      ZC67 2023-09-16 11:03:29   zc67_2023_september_16_03.x00.xlsb         98.042           98.626
    # 39887      ZC67 2023-09-16 11:03:29  zc63_1a2c_rx_of_2023_09_16.x00.xlsb              0                0


# finding out how many samples-files are creating duplicates.
df_simultaneous[['sample_ID', 'Source_File']].drop_duplicates()
    # after removing rows with all parameter-columns having '0' values.
    # Out[49]: 
    #       sample_ID                                       Source_File
    # 0          ZC04  zc04_0a0f_rx_front-housing_2020_02_07-9.x00.xlsb
    # 1          ZC04  zc04_0a0f_rx_front-housing_2020_02_07-9.x01.xlsb
    # 2          ZC04  zc04_0a0f_rx_front-housing_2020_02_07-9.x02.xlsb
    # 3          ZC04  zc04_0a0f_rx_front-housing_2020_02_07-9.x03.xlsb
    # 96         ZC04    zc04_0a0f_rx_front-housing_2020_02_08.x00.xlsb
    #         ...                                               ...
    # 31272      ZC36    zc36_0b73_rx_front-housing_2021_04_13.x01.xlsb
    # 31727      ZC37               zc37_0ac5_2021_april_19_01.x01.xlsb
    # 31728      ZC37               zc37_0ac5_2021_april_19_01.x02.xlsb
    # 32473      ZC38               zc38_0a65_2021_april_19_01.x01.xlsb
    # 32474      ZC38               zc38_0a65_2021_april_19_01.x02.xlsb
    
    # [100 rows x 2 columns]    

    #===================================================================================================================

    # before removing rows with all parameter-columns having '0' values.
    # Out[21]: 
    #       sample_ID                                       Source_File
    # 0          ZC04  zc04_0a0f_rx_front-housing_2020_02_07-9.x00.xlsb
    # 1          ZC04  zc04_0a0f_rx_front-housing_2020_02_07-9.x01.xlsb
    # 2          ZC04  zc04_0a0f_rx_front-housing_2020_02_07-9.x02.xlsb
    # 3          ZC04  zc04_0a0f_rx_front-housing_2020_02_07-9.x03.xlsb
    # 96         ZC04    zc04_0a0f_rx_front-housing_2020_02_08.x00.xlsb
    #         ...                                               ...
    # 32474      ZC38               zc38_0a65_2021_april_19_01.x02.xlsb
    # 33524      ZC60                     0bb9_2023_june_05_03.x00.xlsb
    # 33525      ZC60                0bb9_0bb9_2023_june_05_01.x00.xlsb
    # 39886      ZC67                zc67_2023_september_16_03.x00.xlsb
    # 39887      ZC67               zc63_1a2c_rx_of_2023_09_16.x00.xlsb
    
    # [104 rows x 2 columns]

# %%% drop duplicate rows

df_master.shape
    # Out[56]: (39352, 26)

# This ensures that pandas always encounters the .x00.xlsb file first and keeps it as the "master" row, while discarding the .x01 and .x02 duplicates.
    # This keeps your Source_File column nice and organized for tracing data back to its origin!
df_master = df_master.sort_values(by=['sample_ID', 'timestamp', 'Source_File'])

# 1. Define the specific columns that must match exactly.
duplicate_subset = [
    'sample_ID',
    'timestamp',
    'HR__aver_(bpm)',
    'MBP__aver_(mmHg)',
    'aver__aver_(°C)'
]

# 2. Drop the duplicates
# keep='first' tells pandas to keep the first row it finds (e.g., the .x00 file) and delete the rest
df_master = df_master.drop_duplicates(subset=duplicate_subset, 
                                      keep='first').copy()

df_master.shape
    # Out[60]: (38079, 26)

39352 - 38079
    # Out[61]: 1273 : this number of rows were removed.

#=================================================
#---- time-diff : repeated 

# You can re-run your time_diff calculation here to verify the '0 days' gap is gone!

df_master['time_diff'] = df_master.groupby('sample_ID')['timestamp'].diff()

df_master['time_diff'][:4]
    # Out[72]: 
    # 0                NaT
    # 4    0 days 00:59:59
    # 8    0 days 01:00:01
    # 12   0 days 01:00:00
    # Name: time_diff, dtype: timedelta64[us]

# 3. Count how often each specific time gap occurs
resolution_counts = df_master['time_diff'].value_counts()

resolution_counts.shape
    # Out[64]: (773,)

resolution_counts[:10]
    # Out[65]: 
    # time_diff
    # 0 days 00:00:59    8099
    # 0 days 00:01:01    8089
    # 0 days 00:01:00    5235
    # 0 days 00:59:59    4972
    # 0 days 01:00:01    4940
    # 0 days 01:00:00    4908
    # 0 days 00:02:00      96
    # 0 days 00:02:01      51
    # 0 days 00:03:00      49
    # 0 days 00:01:59      46
    # Name: count, dtype: int64

#======================================================

rounded_diffs = df_master['time_diff'].dt.round('1min')

rounded_diffs.value_counts().shape
    # Out[67]: (211,)

rounded_diffs.value_counts()[:10]
    # Out[69]: 
    # time_diff
    # 0 days 00:01:00    21510
    # 0 days 01:00:00    15065
    # 0 days 01:01:00      209
    # 0 days 00:02:00      201
    # 0 days 00:00:00       98  # number of rows with equal timestamps ( but different values in : "sample_ID", HR, BP, temperature ).
    # 0 days 00:03:00       81
    # 0 days 00:04:00       42
    # 0 days 01:02:00       38
    # 0 days 00:13:00       35
    # 0 days 00:47:00       27
    # Name: count, dtype: int64

#======================================================
#---- time duplicates : re-check

# repeat the time-duplicate find-out step ran previously to check these duplicates !
    # those duplicated rows : duplicated by these 2 columns : 'sample_ID', 'timestamp'.

mask_simultaneous = df_master.duplicated(subset=['sample_ID', 'timestamp'], keep=False)

# Create a dataframe of just these anomalies, sorted so the matching pairs are stacked right next to each other
df_simultaneous = df_master[mask_simultaneous].sort_values(by=['sample_ID', 'timestamp'])

df_simultaneous.shape
    # Out[75]: (20, 26)

# Choose a few relevant columns to inspect to see if the data values are identical or different
inspect_cols = [
    'sample_ID', 
    'timestamp', 
    'Source_File', 
    'HR__aver_(bpm)', 
    'MBP__aver_(mmHg)',
    'aver__aver_(°C)'
]

# PREVIEW OF ZERO-DIFFERENCE PAIRS.
# these rows have the same sample-ID & timestamps.
    # but different values in : HR or BP or temperature.
df_simultaneous[inspect_cols]
    # Out[77]: 
    #       sample_ID           timestamp                                           Source_File HR__aver_(bpm) MBP__aver_(mmHg) aver__aver_(°C)
    # 1107       ZC04 2020-02-18 12:18:39  ZC04_0a0f_2020_february_18_01-2Implantation.x00.xlsb              0                0               0
    # 1106       ZC04 2020-02-18 12:18:39  zc04_0a0f_rx_front-housing_2020_02_18_POD 1.x00.xlsb              0                0           36.84
    # 5443       ZC09 2020-07-07 11:27:44                   0a11_0a11_rx_of_2020_07_07.x00.xlsb              0                0          38.774
    # 5442       ZC09 2020-07-07 11:27:44        zc09_0a11_rx_front-housing_2020_07_06.x00.xlsb              0                0          39.047
    # 9684       ZC15 2020-08-22 11:07:35         zc15_0a67_rx_back-housing_2020_08_22.x00.xlsb         134.61          102.479          39.059
    # 9685       ZC15 2020-08-22 11:07:35                   zc15_0a67_rx_of_2020_08_22.x00.xlsb        156.492            99.11          39.218
    # 30042      ZC35 2021-04-10 20:29:15        zc35_0b72_rx_front-housing_2021_04_10.x00.xlsb        137.354          108.304          37.192
    # 30043      ZC35 2021-04-10 20:29:15        zc35_0b72_rx_front-housing_2021_04_10.x01.xlsb        137.354          108.304          37.196
    # 30124      ZC35 2021-04-12 08:29:34        zc35_0b72_rx_front-housing_2021_04_11.x00.xlsb        162.963           87.661          37.024
    # 30125      ZC35 2021-04-12 08:29:34        zc35_0b72_rx_front-housing_2021_04_11.x01.xlsb        162.971           87.659          37.024
    # 30364      ZC35 2021-04-13 08:29:53        zc35_0b72_rx_front-housing_2021_04_12.x00.xlsb        276.159           99.648          35.964
    # 30365      ZC35 2021-04-13 08:29:53        zc35_0b72_rx_front-housing_2021_04_12.x01.xlsb          257.8          100.067          36.016
    # 30706      ZC35 2021-04-15 04:18:59        zc35_0b72_rx_front-housing_2021_04_14.x00.xlsb         98.167           89.701          35.254
    # 30707      ZC35 2021-04-15 04:18:59        zc35_0b72_rx_front-housing_2021_04_14.x01.xlsb          98.76           89.717           35.26
    # 30728      ZC35 2021-04-15 15:19:18        zc35_0b72_rx_front-housing_2021_04_15.x00.xlsb         85.348           81.075          35.707
    # 30729      ZC35 2021-04-15 15:19:18        zc35_0b72_rx_front-housing_2021_04_15.x01.xlsb         85.392           81.062          35.707
    # 31249      ZC36 2021-04-12 21:29:54        zc36_0b73_rx_front-housing_2021_04_12.x00.xlsb         96.695           60.082          36.258
    # 31250      ZC36 2021-04-12 21:29:54        zc36_0b73_rx_front-housing_2021_04_12.x01.xlsb         96.704           60.079          36.258
    # 31523      ZC36 2021-04-13 23:18:41        zc36_0b73_rx_front-housing_2021_04_13.x00.xlsb        160.495           40.255          36.294
    # 31524      ZC36 2021-04-13 23:18:41        zc36_0b73_rx_front-housing_2021_04_13.x01.xlsb        155.985           40.141          36.294

#================================
#---- drop duplicates : sample_ID & timestamp

df_master.shape 
    # Out[79]: (38079, 26)

df_master = df_master.drop_duplicates(subset=['sample_ID', 'timestamp'], keep='first').copy()

df_master.shape
    # Out[81]: (38069, 26)


38079 - 38069
    # Out[82]: 10
        # half of the duplicate numbers were removed.
        # => df_simultaneous.shape

#================================
#---- recheck time-diffs.

df_master['time_diff'] = df_master.groupby('sample_ID')['timestamp'].diff()

# 3. Count how often each specific time gap occurs
resolution_counts = df_master['time_diff'].value_counts()

resolution_counts.shape
    # Out[18]: (772,)

# time_diff is the index of the resolution_counts Series.
resolution_counts[:20]
    # Out[31]: 
    # time_diff
    # 0 days 00:00:59    8099   # 1 minute group.
    # 0 days 00:01:01    8089
    # 0 days 00:01:00    5235
    ###################################
    # 0 days 00:59:59    4972   # 1 hour group.
    # 0 days 01:00:01    4940
    # 0 days 01:00:00    4908
    ####################################
    # 0 days 00:02:00      96   # miscellaneous !
    # 0 days 00:02:01      51
    # 0 days 00:03:00      49
    # 0 days 00:01:59      46
    # 0 days 01:00:20      35
    # 0 days 01:00:18      32
    # 0 days 00:03:59      20
    # 0 days 01:00:17      20
    # 0 days 01:00:29      18
    # 0 days 01:00:36      17
    # 0 days 01:00:21      17
    # 0 days 01:00:13      16
    # 0 days 01:00:28      16
    # 0 days 00:12:40      15
    # Name: count, dtype: int64

resolution_counts_sorted = resolution_counts.sort_index(ascending=False)

resolution_counts_sorted[:10]
    # Out[29]: 
    # time_diff
    # 21 days 02:17:37    1
    # 13 days 17:46:49    1
    # 11 days 23:16:50    1
    # 10 days 23:05:02    1
    # 10 days 11:42:15    1
    # 6 days 20:35:36     1
    # 5 days 03:00:32     1
    # 5 days 02:40:38     1
    # 5 days 00:52:13     1
    # 3 days 16:45:40     1
    # Name: count, dtype: int64

resolution_counts_sorted[-10:]
    # Out[30]: 
    # time_diff
    # 0 days 00:00:11    3
    # 0 days 00:00:10    1
    # 0 days 00:00:08    4
    # 0 days 00:00:07    2
    # 0 days 00:00:06    3
    # 0 days 00:00:05    9
    # 0 days 00:00:04    5
    # 0 days 00:00:03    1
    # 0 days 00:00:02    2
    # 0 days 00:00:01    5
    # Name: count, dtype: int64

#######################################
#---- !

resolution_counts.value_counts().shape
    # Out[24]: (31,)

resolution_counts.value_counts()
    # Out[25]: 
    # count
    # 1       547
    # 2       105
    # 3        34
    # 4        17
    # 5        13
    # 14        6
    # 6         6
    # 8         5
    # 7         5
    # 10        4
    # 9         4
    # 12        3
    # 20        2
    # 17        2
    # 16        2
    # 13        2
    # 8099      1
    # 8089      1
    # 5235      1
    # 4972      1
    # 4940      1
    # 4908      1
    # 96        1
    # 51        1
    # 49        1
    # 46        1
    # 35        1
    # 32        1
    # 18        1
    # 15        1
    # 11        1
    # Name: count, dtype: int64

###############################

# Round the time differences to the nearest minute ('1min' or 'min')
rounded_diffs = df_master['time_diff'].dt.round('1min')

rounded_diffs.value_counts().shape
    # Out[21]: (211,)

rounded_diffs.value_counts()[:10]
    # Out[22]: 
    # time_diff
    # 0 days 00:01:00    21510
    # 0 days 01:00:00    15065
    # 0 days 01:01:00      209
    # 0 days 00:02:00      201
    # 0 days 00:00:00       88   #  this is a product of rounding : original value is not exactly 0 !  =>  see the next section.
    # 0 days 00:03:00       81
    # 0 days 00:04:00       42
    # 0 days 01:02:00       38
    # 0 days 00:13:00       35
    # 0 days 00:47:00       27
    # Name: count, dtype: int64

###########
# check for duplicate timestamps.
mask_simultaneous = df_master.duplicated(subset=['sample_ID', 'timestamp'], keep=False)

# Create a dataframe of just these anomalies, sorted so the matching pairs are stacked right next to each other
df_simultaneous = df_master[mask_simultaneous].sort_values(by=['sample_ID', 'timestamp'])

df_simultaneous.shape
    # Out[23]: (0, 26)  #  no duplicate timestamps.

# %% delete columns

# delete junk columns.

df_master.shape
    # Out[36]: (38069, 26)



df_master.info()
    # <class 'pandas.DataFrame'>
    # Index: 38069 entries, 0 to 41991
    # Data columns (total 22 columns):
    #  #   Column            Non-Null Count  Dtype          
    # ---  ------            --------------  -----          
    #  0   sample_ID         38069 non-null  str            
    #  1   setup             38069 non-null  str            
    #  2   timetag           38069 non-null  str            
    #  3   timeline          38069 non-null  str            
    #  4   timestamp         38069 non-null  datetime64[us] 
    #  5   cpu-date          38069 non-null  str            
    #  6   cpu-time          38069 non-null  str            
    #  7   period-time       37479 non-null  str            
    #  8   mark-label        36435 non-null  object         
    #  9   step-index        38069 non-null  object         
    #  10  BB__aver_(ms)     29790 non-null  object         
    #  11  HR__aver_(bpm)    29790 non-null  object         
    #  12  DBP__aver_(mmHg)  29789 non-null  object         
    #  13  SBP__aver_(mmHg)  29790 non-null  object         
    #  14  MBP__aver_(mmHg)  29790 non-null  object         
    #  15  aver__aver_(°C)   36705 non-null  object         
    #  16  aver__aver_(%)    34424 non-null  object         
    #  17  Source_File       38069 non-null  str            
    #  18  directory         38069 non-null  str            
    #  19  TI_start_date     38069 non-null  datetime64[us] 
    #  20  days_since_TI     38069 non-null  int64          
    #  21  time_diff         38027 non-null  timedelta64[us]
    # dtypes: datetime64[us](2), int64(1), object(9), str(9), timedelta64[us](1)
    # memory usage: 6.7+ MB



# 1. Create a list of the exact column names to delete
junk_columns = [
    'HR__aver_(bpm)_1',
    'aver__aver_(g)',
    '_2',
    'Abweichung in% HR vs HR',
    'cpu-date', 
    'cpu-time', 
    'time_diff', # Aggregating time_diff makes no sense because the time difference between your new rows will automatically be exactly 1 hour.
    'period-time', 
    'mark-label', 
    'step-index'
]

# 2. Drop them from the dataframe
# Using errors='ignore' is a great safety measure. It tells pandas: 
# "If one of these columns is already gone, just ignore it and don't crash."
df_master = df_master.drop(columns=junk_columns, errors='ignore')

df_master.shape
    # Out[42]: (38069, 17)

#---- type-cast : object => numeric

df_master.info()
    # <class 'pandas.DataFrame'>
    # Index: 38069 entries, 0 to 41991
    # Data columns (total 17 columns):
    #  #   Column            Non-Null Count  Dtype         
    # ---  ------            --------------  -----         
    #  0   sample_ID         38069 non-null  str           
    #  1   setup             38069 non-null  str           
    #  2   timetag           38069 non-null  str           
    #  3   timeline          38069 non-null  str           
    #  4   timestamp         38069 non-null  datetime64[us]
    #  5   cpu-date          38069 non-null  str           
    #  6   BB__aver_(ms)     29790 non-null  object        
    #  7   HR__aver_(bpm)    29790 non-null  object        
    #  8   DBP__aver_(mmHg)  29789 non-null  object        
    #  9   SBP__aver_(mmHg)  29790 non-null  object        
    #  10  MBP__aver_(mmHg)  29790 non-null  object        
    #  11  aver__aver_(°C)   36705 non-null  object        
    #  12  aver__aver_(%)    34424 non-null  object        
    #  13  Source_File       38069 non-null  str           
    #  14  directory         38069 non-null  str           
    #  15  TI_start_date     38069 non-null  datetime64[us]
    #  16  days_since_TI     38069 non-null  int64         
    # dtypes: datetime64[us](2), int64(1), object(7), str(7)
    # memory usage: 5.2+ MB

'''
    Why are physiological columns showing as object?
    In pandas, the object dtype usually means "text string."
    
    When you import data from hundreds of messy Excel files, 
        if even one single cell in the HR__aver_(bpm) column contains a non-number 
        (like a blank space " ", a dash "-", or a machine error code like "Error"), pandas panics. 
        To prevent the data from breaking, pandas imports the entire column as text (object) to accommodate that one rogue character.
    
    Why this matters right now: 
            If you try to run the .agg('mean') function on an object column, pandas will either throw an error or silently skip it.
    The Fix: We must force these columns into float (decimal numbers) before we do the downsampling. 
        We can do this using pd.to_numeric(..., errors='coerce'), which turns the numbers into floats and turns any rogue text into safe NaN blanks.

'''

# physiological columns ( of interest )
phys_cols = [
    'BB__aver_(ms)',
    'HR__aver_(bpm)',
    'DBP__aver_(mmHg)',
    'SBP__aver_(mmHg)',
    'MBP__aver_(mmHg)',
    'aver__aver_(°C)',
    'aver__aver_(%)'
]

# Converting physiological data to numeric floats.
for col in phys_cols:
    df_master[col] = pd.to_numeric(df_master[col], errors='coerce')

df_master.info()
    # <class 'pandas.DataFrame'>
    # Index: 38069 entries, 0 to 41991
    # Data columns (total 16 columns):
    #  #   Column            Non-Null Count  Dtype         
    # ---  ------            --------------  -----         
    #  0   sample_ID         38069 non-null  str           
    #  1   setup             38069 non-null  str           
    #  2   timetag           38069 non-null  str           
    #  3   timeline          38069 non-null  str           
    #  4   timestamp         38069 non-null  datetime64[us]
    #  5   BB__aver_(ms)     29790 non-null  float64       
    #  6   HR__aver_(bpm)    29790 non-null  float64       
    #  7   DBP__aver_(mmHg)  29789 non-null  float64       
    #  8   SBP__aver_(mmHg)  29790 non-null  float64       
    #  9   MBP__aver_(mmHg)  29790 non-null  float64       
    #  10  aver__aver_(°C)   36705 non-null  float64       
    #  11  aver__aver_(%)    34424 non-null  float64       
    #  12  Source_File       38069 non-null  str           
    #  13  directory         38069 non-null  str           
    #  14  TI_start_date     38069 non-null  datetime64[us]
    #  15  days_since_TI     38069 non-null  int64         
    # dtypes: datetime64[us](2), float64(7), int64(1), str(6)
    # memory usage: 4.9 MB


# %% down-re-sample

df_master.shape
    # Out[31]: (38069, 16)

#---- dictionary of functions
# Dynamically build an aggregation dictionary of functions :
    # 2 functions : mean , first.
# This tells pandas: "If it's a number, average it. If it's text, keep the first one."
# first :
    # 'first' is a built-in pandas aggregation command.
    # "Look at all the rows inside this 1-hour bucket. 
        # Just grab the top-most value you see and throw the rest away."
agg_dict = {}
for col in df_master.columns:
    # Skip the grouping columns
    if col in ['sample_ID', 'timestamp']:
        continue
    # If it is a phyiological column, calculate the mean
    elif col in phys_cols:
        agg_dict[col] = 'mean'
    # If the column is text (timetag, setup, Source_File), take the first entry in that hour
    else:
        agg_dict[col] = 'first'

# just to be able to compare it later to the new dataframe.
df_master_original_timestamp = df_master.copy()

#====================================================================================
#---- down-re-sample
# Perform the grouping and downsampling
# freq='1h' creates strict 1-hour buckets starting exactly at the top of the hour (e.g., 01:00:00)
# grouping :
    # First, isolate the data by the individual pig (sample_ID).
    # Then, within that specific pig's data, group it further into 1-hour time buckets (pd.Grouper(...))( = re-sample , regularization ).
df_master = df_master_original_timestamp.groupby(by=['sample_ID', 
                                                     pd.Grouper(key='timestamp', 
                                                                freq='1h'
                                                                )
                                                     ]
                                                 ).agg(agg_dict).reset_index()

#====================================================================================
# Verification Report

df_master.shape
    # Out[30]: (16226, 16)

preview_cols = ['sample_ID', 'timestamp', 'timetag', 'HR__aver_(bpm)', 'MBP__aver_(mmHg)']

df_master[preview_cols][:4]
    # Out[33]: 
    #   sample_ID           timestamp timetag  HR__aver_(bpm)  MBP__aver_(mmHg)
    # 0      ZC04 2020-02-07 15:00:00    TI_4      132.788000        117.624000
    # 1      ZC04 2020-02-07 16:00:00    TI_4      127.771000        106.015000
    # 2      ZC04 2020-02-07 17:00:00    TI_4      113.823000        128.930000
    # 3      ZC04 2020-02-07 18:00:00    TI_4      123.002000        114.023000

df_master[preview_cols][1000:1004]
    # Out[34]: 
    #      sample_ID           timestamp timetag  HR__aver_(bpm)  MBP__aver_(mmHg)
    # 1000      ZC07 2020-06-13 12:00:00   POD_4      115.130000         85.628000
    # 1001      ZC07 2020-06-13 13:00:00   POD_4      215.535667         58.838667
    # 1002      ZC07 2020-06-13 14:00:00   POD_4      120.956000         86.434000
    # 1003      ZC07 2020-06-13 15:00:00   POD_4      113.916000         90.711000

df_master[preview_cols][-4:]
    # Out[35]: 
    #       sample_ID           timestamp timetag  HR__aver_(bpm)  MBP__aver_(mmHg)
    # 16222      ZC69 2023-10-26 10:00:00   POD_2       93.704000         92.649000
    # 16223      ZC69 2023-10-26 11:00:00   POD_2       95.267857         90.756857
    # 16224      ZC69 2023-10-26 12:00:00   POD_2             NaN               NaN
    # 16225      ZC69 2023-10-26 13:00:00   POD_2      151.515000        177.322000

# %% matrix

# Data Overview Matrix

#===================================================================
#---- sample matrix

# number of individual samples ( pigs ) in each setup-timetag intersection.
# Create the pivot table
overview_matrix_sample = pd.pivot_table(
    df_master,                   # Use your newly regularized dataset
    values='sample_ID',          # The column we want to count
    index='timetag',             # Rows of the matrix
    columns='setup',             # Columns of the matrix
    aggfunc='nunique',           # 'nunique' counts the number of unique pigs!
    fill_value=0                 # Replaces empty intersections with 0 instead of NaN
)

# problem : the index ( timetag ) isnot chronologically ordred.
overview_matrix_sample
    # Out[41]: 
    # setup              Housing  OF  Stoffwechselkäfig  Surgery
    # timetag                                                   
    # Explantation            33   1                  4       36
    # Implantation            33   5                  0       27
    # POD_1                   33  22                  0        1
    # POD_2                   32   5                  0        4
    # POD_3                   29  24                  0        1
    # POD_4                   28  18                  0        0
    # POD_5                   28   0                  0        0
    # POD_6                   24   3                  0        3
    # Post_Sacrifice_13        0   0                  0        1
    # Post_Sacrifice_27        0   0                  0        1
    # Post_Sacrifice_34        0   0                  0        1
    # Retraining_1            34  28                  0        0
    # Retraining_2            32  28                  4        1
    # Sacrifice               20  14                  0       19
    # TI                      31   0                  0       35
    # TI_1                    34   0                  0        2
    # TI_10                   33   0                  0        0
    # TI_11                   34   0                  0        0
    # TI_2                    32   0                  0        2
    # TI_3                    34   0                  0        0
    # TI_4                    35   0                  0        0
    # TI_5                    35   0                  0        0
    # TI_6                    35   0                  0        0
    # TI_7                    34   0                  0        0
    # TI_8                    30   0                  0        0
    # TI_9                    32   0                  0        0

#=================================================
#---- chronological order of 'timetag'
df_tag_order = df_master[['timetag','days_since_TI']].drop_duplicates().sort_values(by=['days_since_TI'])

df_tag_order
    # Out[46]: 
    #                  timetag  days_since_TI
    # 424                   TI              0
    # 438                 TI_1              1
    # 592                 TI_2              2
    # 606                 TI_3              3
    # 0                   TI_4              4
    # 9                   TI_5              5
    # 33                  TI_6              6
    # 55                  TI_7              7
    # 79                  TI_8              8
    # 103                 TI_9              9
    # 126                TI_10             10
    # 150                TI_11             11
    # 174         Retraining_1             12
    # 198         Retraining_2             13
    # 222         Explantation             14
    # 246         Implantation             15
    # 270                POD_1             16
    # 294                POD_2             17
    # 318                POD_3             18
    # 342                POD_4             19
    # 366                POD_5             20
    # 390                POD_6             21
    # 414            Sacrifice             22
    # 14067  Post_Sacrifice_13             35
    # 14069  Post_Sacrifice_27             49
    # 14077  Post_Sacrifice_34             56

# 1. Extract the sorted list of timetags from your ordering dataframe
list_ordered_timetags = df_tag_order['timetag'].tolist()

list_ordered_timetags
    # Out[66]: 
    # ['TI',
    #  'TI_1',
    #  'TI_2',
    #  'TI_3',
    #  'TI_4',
    #  'TI_5',
    #  'TI_6',
    #  'TI_7',
    #  'TI_8',
    #  'TI_9',
    #  'TI_10',
    #  'TI_11',
    #  'Retraining_1',
    #  'Retraining_2',
    #  'Explantation',
    #  'Implantation',
    #  'POD_1',
    #  'POD_2',
    #  'POD_3',
    #  'POD_4',
    #  'POD_5',
    #  'POD_6',
    #  'Sacrifice',
    #  'Post_Sacrifice_13',
    #  'Post_Sacrifice_27',
    #  'Post_Sacrifice_34']

# 2. Filter the list to only include tags that actually exist in the matrix's index
# (This is a safety measure to prevent pandas from adding blank rows if a tag is missing)
valid_ordered_timetags = [tag 
                          for tag in list_ordered_timetags 
                          if tag in overview_matrix_sample.index]

# 3. Apply the specific order to the index (rows) of the matrix
overview_matrix_sample = overview_matrix_sample.loc[valid_ordered_timetags]

# the matrix with the index (timetag) with chronological order.
overview_matrix_sample
    # Out[50]: 
    # setup              Housing  OF  Stoffwechselkäfig  Surgery
    # timetag                                                   
    # TI                      31   0                  0       35
    # TI_1                    34   0                  0        2
    # TI_2                    32   0                  0        2
    # TI_3                    34   0                  0        0
    # TI_4                    35   0                  0        0
    # TI_5                    35   0                  0        0
    # TI_6                    35   0                  0        0
    # TI_7                    34   0                  0        0
    # TI_8                    30   0                  0        0
    # TI_9                    32   0                  0        0
    # TI_10                   33   0                  0        0
    # TI_11                   34   0                  0        0
    # Retraining_1            34  28                  0        0
    # Retraining_2            32  28                  4        1
    # Explantation            33   1                  4       36
    # Implantation            33   5                  0       27
    # POD_1                   33  22                  0        1
    # POD_2                   32   5                  0        4
    # POD_3                   29  24                  0        1
    # POD_4                   28  18                  0        0
    # POD_5                   28   0                  0        0
    # POD_6                   24   3                  0        3
    # Sacrifice               20  14                  0       19
    # Post_Sacrifice_13        0   0                  0        1
    # Post_Sacrifice_27        0   0                  0        1
    # Post_Sacrifice_34        0   0                  0        1

#---- save
base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER\matrix")
file_name = 'overview_matrix_sample'
overview_matrix_sample.to_pickle( base_dir / f"{file_name}.pkl" )
overview_matrix_sample.to_excel( base_dir / f"{file_name}.xlsx" )


#=================================================
#---- volume

# TOTAL RECORDED HOURS PER SETUP & TIMETAG

# Since your dataset is now perfectly downsampled to 1-hour intervals, 
    # counting the number of rows is mathematically identical to counting the total hours of recorded data.

# 1. Create the pivot table using 'count' to get the number of rows
volume_matrix = pd.pivot_table(
    df_master,
    values='sample_ID',          # We can still count the ID column, but now it counts every instance
    index='timetag',             # Keeping your preferred orientation
    columns='setup',
    aggfunc='count',             # <--- The magic change: counts total rows instead of unique pigs
    fill_value=0
)

# 2. Re-apply your chronological sorting
valid_ordered_timetags_volume = [tag 
                                 for tag in list_ordered_timetags 
                                 if tag in volume_matrix.index]
volume_matrix = volume_matrix.loc[valid_ordered_timetags_volume]

# TOTAL RECORDED HOURS PER SETUP & TIMETAG.
volume_matrix
    # Out[69]: 
    # setup              Housing  OF  Stoffwechselkäfig  Surgery
    # timetag                                                   
    # TI                     335   0                  0       99
    # TI_1                   734   0                  0       48
    # TI_2                   724   0                  0       32
    # TI_3                   788   0                  0        0
    # TI_4                   823   0                  0        0
    # TI_5                   813   0                  0        0
    # TI_6                   804   0                  0        0
    # TI_7                   738   0                  0        0
    # TI_8                   710   0                  0        0
    # TI_9                   733   0                  0        0
    # TI_10                  746   0                  0        0
    # TI_11                  783   0                  0        0
    # Retraining_1           764  29                  0        0
    # Retraining_2           641  29                 60       15
    # Explantation           535   1                 31      178
    # Implantation           644   5                  0      125
    # POD_1                  743  23                  0        2
    # POD_2                  703   5                  0        8
    # POD_3                  645  24                  0        2
    # POD_4                  650  19                  0        0
    # POD_5                  611   0                  0        0
    # POD_6                  509   3                  0        3
    # Sacrifice              254  15                  0       29
    # Post_Sacrifice_13        0   0                  0        2
    # Post_Sacrifice_27        0   0                  0        8
    # Post_Sacrifice_34        0   0                  0        1

#---- save
base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER\matrix")
file_name = 'volume_matrix'
volume_matrix.to_pickle( base_dir / f"{file_name}.pkl" )
volume_matrix.to_excel( base_dir / f"{file_name}.xlsx" )


# %% I/O

# 7 : 'timetag' was added.
# 8 : duplicate rows ( minute inputs ) were deleted.
# 10 : 
    # 0 => NaN
    # delete the junk columns.
# 11 : 
        # junk columns removed.
        # physiological columns : type-casting : object => numeric
        # down-re-sampling

#---- address / name
base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
file_name = 'Master_Telemetry_Dataset_11'

#======================================================================
#---- save

output_file = base_dir / f"{file_name}.pkl"
df_master.to_pickle(output_file)


output_file = base_dir / f"{file_name}.xlsx"
df_master.to_excel(output_file, index=False)

#==================================================================
#---- read

source_file = base_dir / f"{file_name}.pkl"
df_master = pd.read_pickle(source_file)

#==================
# df_TI_start_dates_all   =>  top

# %%'

