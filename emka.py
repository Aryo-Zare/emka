
# pyxlsb : should be instaled in the enviornment to be able to read '.xlsb' files.

# %%

import shutil
from pathlib import Path

# %%


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

# %%

# Starting file copy process...
# copied in 1s.
# Success! Copied 1452 Excel files to F:\OneDrive - Uniklinik RWTH Aachen\EMKA\copy_excel

# %%

file_path = r'F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC09\EMKA\Housing\zc09_0a11_rx_front-housing_2020_06_16.x00.xlsb'
df_test = pd.read_excel(file_path)

# for opening '.xlsb' files, this is needed to be installed.
# ImportError: `Import pyxlsb` failed.  Use pip or conda to install the pyxlsb package.

df_test.shape
    # Out[20]: (249, 20)

df_test.head()
    # Out[21]: 
    #                 Unnamed: 0           Unnamed: 1  ... Unnamed: 18 Unnamed: 19
    # 0  ecgAUTO analysis report                  NaN  ...         NaN         NaN
    # 1                      NaN                  NaN  ...         NaN         NaN
    # 2                        0  main-header section  ...         NaN         NaN
    # 3                      NaN                  NaN  ...         NaN         NaN
    # 4                      NaN                  NaN  ...         NaN         NaN
    
    # [5 rows x 20 columns]

# %%

df_test[0] == 'cpu-date'
    # KeyError: 0

# %%

step_matches = df_test[df_test[0].astype(str).str.contains('steps section', case=False, na=False)].index

# %%

import pandas as pd
from pathlib import Path
import logging

# %%


# 1. Define your folders
base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel")
# base_dir = Path(r'F:\temp\4')
output_file = base_dir / "Master_Telemetry_Dataset.csv"
log_file = base_dir / "extraction_log.txt"

# 2. Set up the Logger (Writes to both console and a text file)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
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
                    
                    # STEP C: Find the first EMPTY ROW in the specific 'cpu-date' column
                    df_below_header = df_raw.iloc[header_idx + 1:]
                    is_empty_row = pd.isna(df_below_header[cpu_date_col_idx]) | (df_below_header[cpu_date_col_idx].astype(str).str.strip() == '')
                    empty_indices = df_below_header[is_empty_row].index
                    
                    if not empty_indices.empty:
                        end_idx = empty_indices[1]  # 2nd empty row.  1st empty row is right after 'cpu-time' ( just for organization ).
                    else:
                        end_idx = len(df_raw) 
                        
                    # Extract the clean chunk
                    df_chunk = df_raw.iloc[header_idx + 2 : end_idx].copy() # + 2 : so that the 'units' (mmHg , ...) row would not be selected.
                    df_chunk.columns = df_raw.iloc[header_idx]  # this fetches the header from the original file, not the chunked one.
                    df_chunk.dropna(axis=1, how='all', inplace=True)
                    
                    # Add Metadata
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

# %%

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


# %%

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


# %%


