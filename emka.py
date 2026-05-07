
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
    
# %%
    
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
failed_rows = df_master[df_master['sample_ID'] == 'Unknown']

# Get the unique original directory paths from those rows
problematic_directories = failed_rows['directory'].unique()

print(f"Found {len(problematic_directories)} problematic directories that didn't match the rule:\n")

# Print them out one by one so you can inspect them
for folder in problematic_directories:
    print(folder)

# out    
    # Found 4 problematic directories that didn't match the rule:
    
        # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC17\EMKA
        # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC11\EMKA
        # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC08\EMKA
        # F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC06\EMKA

# I checked them : there are single excel files inside : copy_excel\sample_ID\EMKA :
        # outside other subfolders ( 'Housing', ... ).
        # so they don't belong to any experiental-setup !

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

# %% I/O

base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
output_file = base_dir / "Master_Telemetry_Dataset_3.pkl"
df_master.to_pickle(output_file)


base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
output_file = base_dir / "Master_Telemetry_Dataset_3.xlsx"
df_master.to_excel(output_file, index=False)

base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
source_file = base_dir / "Master_Telemetry_Dataset_3.pkl"
df_master = pd.read_pickle(source_file)


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

# %%'



