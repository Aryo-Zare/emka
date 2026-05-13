
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

ID_columns = ['sample_ID', 'setup', 'timeline', 'timestamp']

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

# %%'

df_master.iloc[:5,:4]
    # Out[83]: 
    #   sample_ID    setup timeline           timestamp
    # 0      ZC69  Housing      N/A 2023-10-10 12:01:10
    # 1      ZC69  Housing      N/A 2023-10-10 13:01:10
    # 2      ZC69  Housing      N/A 2023-10-10 14:01:09
    # 3      ZC69  Housing      N/A 2023-10-10 15:01:10
    # 4      ZC69  Housing      N/A 2023-10-10 16:01:10


# %% I/O

#---- address / name
base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
file_name = 'Master_Telemetry_Dataset_6'

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


# %%'

