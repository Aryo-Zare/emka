
# %%'

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

# %%'

df_test[0] == 'cpu-date'
    # KeyError: 0

# %%'

step_matches = df_test[df_test[0].astype(str).str.contains('steps section', case=False, na=False)].index

# %%'

df_raw_0.shape
    # Out[17]: (249, 20)

mask_steps = df_raw_0.astype(str).apply(lambda col: col.str.contains('steps section', case=False, na=False))


mask_steps
    # Out[19]: 
    #         0      1      2      3      4   ...     15     16     17     18     19
    # 0    False  False  False  False  False  ...  False  False  False  False  False
    # 1    False  False  False  False  False  ...  False  False  False  False  False
    # 2    False  False  False  False  False  ...  False  False  False  False  False
    # 3    False  False  False  False  False  ...  False  False  False  False  False
    # 4    False  False  False  False  False  ...  False  False  False  False  False
    # ..     ...    ...    ...    ...    ...  ...    ...    ...    ...    ...    ...
    # 244  False  False  False  False  False  ...  False  False  False  False  False
    # 245  False  False  False  False  False  ...  False  False  False  False  False
    # 246  False  False  False  False  False  ...  False  False  False  False  False
    # 247  False  False  False  False  False  ...  False  False  False  False  False
    # 248  False  False  False  False  False  ...  False  False  False  False  False
    
    # [249 rows x 20 columns]

mask_steps.sum()
    # Out[20]: 
    # 0     0
    # 1     1
    # 2     0
    # 3     0
    # 4     0
    # 5     0
    # 6     0
    # 7     0
    # 8     0
    # 9     0
    # 10    0
    # 11    0
    # 12    0
    # 13    0
    # 14    0
    # 15    0
    # 16    0
    # 17    0
    # 18    0
    # 19    0
    # dtype: int64

step_rows = mask_steps.any(axis=1) 

step_rows
    # Out[22]: 
    # 0      False
    # 1      False
    # 2      False
    # 3      False
    # 4      False
     
    # 244    False
    # 245    False
    # 246    False
    # 247    False
    # 248    False
    # Length: 249, dtype: bool

type(mask_steps)
    # Out[23]: pandas.DataFrame

step_rows.sum()
    # Out[24]: np.int64(1)

step_rows[200:210]
    # Out[25]: 
    # 200    False
    # 201    False
    # 202     True
    # 203    False
    # 204    False
    # 205    False
    # 206    False
    # 207    False
    # 208    False
    # 209    False
    # dtype: bool


step_rows.any()
    # Out[26]: np.True_



step_rows[step_rows]
    # Out[27]: 
    # 202    True
    # dtype: bool


step_rows[step_rows].index[0]
    # Out[28]: 202


df_below_steps = df_raw_0.iloc[steps_row_idx:]



df_below_steps.iloc[: , 3:10]
    # Out[36]: 
    #                                                      3  ...          9
    # 202                                                NaN  ...        NaN
    # 203                               no mark, no comment.  ...        NaN
    # 204               missing values reported as 0 (zero).  ...        NaN
    # 205  missing values reported as 0 (zero).all steps ...  ...        NaN
    # 206                                                NaN  ...        NaN
    # 207                                                NaN  ...        NaN
    # 208                                           cpu-time  ...  DBP__aver
    # 209                                                NaN  ...       mmHg
    # 210                                           0.578449  ...     71.612
    # 211                                           0.620116  ...     75.052
    # 212                                           0.661782  ...     73.769
    # 213                                           0.703449  ...     76.028
    # 214                                           0.745116  ...     75.311
    # 215                                           0.786782  ...     68.771
    # 216                                           0.828449  ...     75.251
    # 217                                           0.870116  ...     78.892
    # 218                                           0.911782  ...     79.711
    # 219                                           0.953449  ...     79.466
    # 220                                           0.995116  ...     98.019
    # 221                                           0.036782  ...    103.007
    # 222                                           0.078449  ...     78.063
    # 223                                           0.120116  ...     88.584
    # 224                                           0.161782  ...     99.876
    # 225                                           0.203449  ...     99.813
    # 226                                           0.245116  ...     99.037
    # 227                                           0.286782  ...     91.621
    # 228                                           0.328449  ...     84.828
    # 229                                           0.370116  ...          0
    # 230                                           0.411782  ...          0
    # 231                                           0.453449  ...          0
    # 232                                           0.495116  ...          0
    # 233                                           0.536782  ...          0
    # 234                                                NaN  ...        NaN
    # 235                                                NaN  ...        NaN
    # 236                                           cpu-time  ...  mark-unit
    # 237                                                NaN  ...        NaN
    # 238                                                NaN  ...        NaN
    # 239                                           cpu-time  ...  mark-unit
    # 240                                                NaN  ...        NaN
    # 241                                                NaN  ...        NaN
    # 242                                                NaN  ...        NaN
    # 243                                                val  ...        NaN
    # 244                                                NaN  ...        NaN
    # 245                                                NaN  ...        NaN
    # 246                                                NaN  ...        NaN
    # 247                                                NaN  ...        NaN
    # 248                                                NaN  ...        NaN
    
    # [47 rows x 7 columns]


# as seen : the date & time  are shown as fractions !
    # 44306/365
        # Out[73]: 121.38630136986302    :  121 years : start-year = 1900 .
    
    # minutes
    # 52/60
        # Out[70]: 0.8666666666666667

    # hour + minute.
    # 14.86/24
        # Out[71]: 0.6191666666666666
df_below_steps.iloc[6:12 , 2:8]
    # Out[72]: 
    #             2         3            4           5           6         7
    # 208  cpu-date  cpu-time  period-time  mark-label  step-index  BB__aver
    # 209       NaN       NaN          NaN         NaN         NaN        ms
    # 210     44306  0.578449            0         NaN           1   623.091
    # 211     44306  0.620116     0.041664     housing           2   665.767
    # 212     44306  0.661782     0.083331     housing           3   671.761
    # 213     44306  0.703449     0.124997     housing           4   654.305

mask_cpu = df_below_steps.astype(str).apply(lambda col: col.str.contains('cpu-date', case=False, na=False))

mask_cpu
    # Out[38]: 
    #         0      1      2      3      4   ...     15     16     17     18     19
    # 202  False  False  False  False  False  ...  False  False  False  False  False
    # 203  False  False  False  False  False  ...  False  False  False  False  False
    # 204  False  False  False  False  False  ...  False  False  False  False  False
    # 205  False  False  False  False  False  ...  False  False  False  False  False
    # 206  False  False  False  False  False  ...  False  False  False  False  False
    # 207  False  False  False  False  False  ...  False  False  False  False  False
    # 208  False  False   True  False  False  ...  False  False  False  False  False
    # 209  False  False  False  False  False  ...  False  False  False  False  False
    # 210  False  False  False  False  False  ...  False  False  False  False  False
    # 211  False  False  False  False  False  ..


cpu_rows = mask_cpu.any(axis=1)

cpu_rows
    # Out[40]: 
    # 202    False
    # 203    False
    # 204    False
    # 205    False
    # 206    False
    # 207    False
    # 208     True
    # 209    False
    # 210    False
    # 211    False


cpu_rows.any()
    # Out[42]: np.True_


cpu_rows[cpu_rows]
    # Out[43]: 
    # 208    True
    # 236    True
    # 239    True
    # dtype: bool


cpu_rows[cpu_rows].index[0]
    # Out[44]: np.int64(208)


# row index.
header_idx = cpu_rows[cpu_rows].index[0]


col_matches = mask_cpu.loc[header_idx]


col_matches
    # Out[55]: 
    # 0     False
    # 1     False
    # 2      True
    # 3     False
    # 4     False
    # 5     False
    # 6     False
    # 7     False
    # 8     False
    # 9     False
    # 10    False
    # 11    False
    # 12    False
    # 13    False
    # 14    False
    # 15    False
    # 16    False
    # 17    False
    # 18    False
    # 19    False
    # Name: 208, dtype: bool

col_matches[col_matches]
    # Out[58]: 
    # 2    True
    # Name: 208, dtype: bool

# column index
cpu_date_col_idx = col_matches[col_matches].index[0]

# %%'

# explore

header_idx
    # Out[61]: np.int64(208)

cpu_date_col_idx
    # Out[60]: np.int64(2)

# %%'

df_below_header = df_raw_0.iloc[header_idx + 1:]

df_below_header.iloc[ :5, :5]
    # Out[74]: 
    #        0    1      2         3         4
    # 209  NaN  NaN    NaN       NaN       NaN
    # 210  NaN  NaN  44306  0.578449         0
    # 211  NaN  NaN  44306  0.620116  0.041664
    # 212  NaN  NaN  44306  0.661782  0.083331
    # 213  NaN  NaN  44306  0.703449  0.124997

# this was the source of problem :
    # you should not search for the 1st empty row.
    # but, the 2nd empty row !
is_empty_row = pd.isna(df_below_header[cpu_date_col_idx]) | (df_below_header[cpu_date_col_idx].astype(str).str.strip() == '')
    # Out[76]: 
    # 209     True
    # 210    False
    # 211    False
    # 212    False
    # ...


empty_indices = df_below_header[is_empty_row].index

empty_indices
    # Out[78]: Index([209, 234, 235, 237, 238, 240, 241, 244, 246, 247, 248], dtype='int64')


# %%'

# after running the mani program on a test file, with the corrected dnd-index :
final_dataset.iloc[:6,:6]
    # Out[82]: 
    # 208 cpu-date  cpu-time period-time mark-label step-index BB__aver
    # 0        NaN       NaN         NaN        NaN        NaN       ms
    # 1      44306  0.578449           0        NaN          1  623.091
    # 2      44306  0.620116    0.041664    housing          2  665.767
    # 3      44306  0.661782    0.083331    housing          3  671.761
    # 4      44306  0.703449    0.124997    housing          4  654.305
    # 5      44306  0.745116    0.166664    housing          5  624.076

# %%'

# practice : .loc

# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.loc.html

df.loc[["viper", "sidewinder"] , "shield" ]
    # Out[53]: 
    # viper         5
    # sidewinder    8
    # Name: shield, dtype: int64



# %%'


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


# %% count
# %%% lines

# counting the number of lines in the log text-file.

from pathlib import Path

file_path = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER\extraction_log__.txt")

with file_path.open("r", encoding="utf-8") as f:
    num_lines = sum(1 for _ in f)

print(num_lines)

    # Out :
    # 1455

# %%% excel files

# count the number of Excel files ion the folder.

from pathlib import Path

folder = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel")

excel_extensions = {".xls", ".xlsx", ".xlsb"}

num_excel_files = sum(
    1 
    for file in folder.rglob("*")
    if file.is_file() and file.suffix.lower() in excel_extensions
)

print(num_excel_files)

    # out
    # 1453

# %% duplicate timestamps

df_simultaneous[inspect_cols][-20:-10]
    # Out[48]: 
    #       sample_ID           timestamp                          Source_File HR__aver_(bpm) SBP__aver_(mmHg)
    # 32598      ZC38 2021-04-19 12:32:16  zc38_0a65_2021_april_19_01.x01.xlsb        123.207           99.181
    # 32599      ZC38 2021-04-19 12:32:16  zc38_0a65_2021_april_19_01.x02.xlsb        123.207           99.181
    # 32600      ZC38 2021-04-19 12:33:17  zc38_0a65_2021_april_19_01.x01.xlsb              0                0
    # 32601      ZC38 2021-04-19 12:33:17  zc38_0a65_2021_april_19_01.x02.xlsb              0                0
    # 32602      ZC38 2021-04-19 12:34:16  zc38_0a65_2021_april_19_01.x01.xlsb        108.077             99.3
    # 32603      ZC38 2021-04-19 12:34:16  zc38_0a65_2021_april_19_01.x02.xlsb        108.077             99.3
    # 32604      ZC38 2021-04-19 12:35:17  zc38_0a65_2021_april_19_01.x01.xlsb        116.368          114.829
    # 32605      ZC38 2021-04-19 12:35:17  zc38_0a65_2021_april_19_01.x02.xlsb        116.368          114.829
    # 32606      ZC38 2021-04-19 12:36:17  zc38_0a65_2021_april_19_01.x01.xlsb        116.622          117.615
    # 32607      ZC38 2021-04-19 12:36:17  zc38_0a65_2021_april_19_01.x02.xlsb        116.622          117.615

df_simultaneous[inspect_cols][-110:-100]
    # Out[50]: 
    #       sample_ID           timestamp                          Source_File HR__aver_(bpm) SBP__aver_(mmHg)
    # 32497      ZC38 2021-04-19 11:42:17  zc38_0a65_2021_april_19_01.x01.xlsb        104.738           99.067
    # 32498      ZC38 2021-04-19 11:42:17  zc38_0a65_2021_april_19_01.x02.xlsb        104.738           99.067
    # 32499      ZC38 2021-04-19 11:43:17  zc38_0a65_2021_april_19_01.x01.xlsb        109.978             98.9
    # 32500      ZC38 2021-04-19 11:43:17  zc38_0a65_2021_april_19_01.x02.xlsb        109.978             98.9
    # 32501      ZC38 2021-04-19 11:44:16  zc38_0a65_2021_april_19_01.x01.xlsb        108.418            98.65
    # 32502      ZC38 2021-04-19 11:44:16  zc38_0a65_2021_april_19_01.x02.xlsb        108.418            98.65
    # 32503      ZC38 2021-04-19 11:45:17  zc38_0a65_2021_april_19_01.x01.xlsb        106.321           97.431
    # 32504      ZC38 2021-04-19 11:45:17  zc38_0a65_2021_april_19_01.x02.xlsb        106.321           97.431
    # 32506      ZC38 2021-04-19 11:46:16  zc38_0a65_2021_april_19_01.x01.xlsb        104.738          102.725
    # 32507      ZC38 2021-04-19 11:46:16  zc38_0a65_2021_april_19_01.x02.xlsb        104.738          102.725

df_simultaneous[inspect_cols][-1010:-1000]
    # Out[51]: 
    #       sample_ID           timestamp                                       Source_File HR__aver_(bpm) SBP__aver_(mmHg)
    # 15826      ZC22 2020-10-27 10:07:34  zc22_0ac7_rx_front-housing_2020_10_26-2.x00.xlsb        109.206          121.358
    # 15827      ZC22 2020-10-27 10:07:34  zc22_0ac7_rx_front-housing_2020_10_26-2.x01.xlsb        109.206          121.358
    # 15828      ZC22 2020-10-27 11:07:35  zc22_0ac7_rx_front-housing_2020_10_26-2.x00.xlsb        109.937          123.869
    # 15829      ZC22 2020-10-27 11:07:35  zc22_0ac7_rx_front-housing_2020_10_26-2.x01.xlsb        109.937          123.869
    # 15840      ZC22 2020-10-27 12:07:35  zc22_0ac7_rx_front-housing_2020_10_26-2.x00.xlsb         97.789          123.758
    # 15841      ZC22 2020-10-27 12:07:35  zc22_0ac7_rx_front-housing_2020_10_26-2.x01.xlsb         97.789          123.758
    # 19343      ZC26 2020-11-09 15:13:38    zc26_0ae4_rx_front-housing_2020_11_09.x00.xlsb        147.406          122.498
    # 19344      ZC26 2020-11-09 15:13:38    zc26_0ae4_rx_front-housing_2020_11_09.x01.xlsb        147.406          122.498
    # 19345      ZC26 2020-11-09 16:13:37    zc26_0ae4_rx_front-housing_2020_11_09.x00.xlsb        146.827          114.671
    # 19346      ZC26 2020-11-09 16:13:37    zc26_0ae4_rx_front-housing_2020_11_09.x01.xlsb        146.827          114.671

#=========================================
#---- + temperature

df_simultaneous[inspect_cols].head(10)
    # Out[53]: 
    #   sample_ID           timestamp                                       Source_File HR__aver_(bpm) SBP__aver_(mmHg) aver__aver_(°C)
    # 0      ZC04 2020-02-07 15:53:53  zc04_0a0f_rx_front-housing_2020_02_07-9.x00.xlsb        132.788          125.356          38.899
    # 1      ZC04 2020-02-07 15:53:53  zc04_0a0f_rx_front-housing_2020_02_07-9.x01.xlsb        132.788          125.356          38.899
    # 2      ZC04 2020-02-07 15:53:53  zc04_0a0f_rx_front-housing_2020_02_07-9.x02.xlsb        132.788          125.356          38.899
    # 3      ZC04 2020-02-07 15:53:53  zc04_0a0f_rx_front-housing_2020_02_07-9.x03.xlsb        132.788          125.356          38.899
    # 4      ZC04 2020-02-07 16:53:52  zc04_0a0f_rx_front-housing_2020_02_07-9.x00.xlsb        127.771          113.238          38.804
    # 5      ZC04 2020-02-07 16:53:52  zc04_0a0f_rx_front-housing_2020_02_07-9.x01.xlsb        127.771          113.238          38.804
    # 6      ZC04 2020-02-07 16:53:52  zc04_0a0f_rx_front-housing_2020_02_07-9.x02.xlsb        127.771          113.238          38.804
    # 7      ZC04 2020-02-07 16:53:52  zc04_0a0f_rx_front-housing_2020_02_07-9.x03.xlsb        127.771          113.238          38.804
    # 8      ZC04 2020-02-07 17:53:53  zc04_0a0f_rx_front-housing_2020_02_07-9.x00.xlsb        113.823          137.825          38.964
    # 9      ZC04 2020-02-07 17:53:53  zc04_0a0f_rx_front-housing_2020_02_07-9.x01.xlsb        113.823          137.825          38.964

    # df_simultaneous[inspect_cols][-10:]
    # Out[54]: 
    #       sample_ID           timestamp                          Source_File HR__aver_(bpm) SBP__aver_(mmHg) aver__aver_(°C)
    # 32608      ZC38 2021-04-19 12:37:16  zc38_0a65_2021_april_19_01.x01.xlsb        115.088          117.398          36.075
    # 32609      ZC38 2021-04-19 12:37:16  zc38_0a65_2021_april_19_01.x02.xlsb        115.088          117.398          36.075
    # 32610      ZC38 2021-04-19 12:38:17  zc38_0a65_2021_april_19_01.x01.xlsb        116.262          115.048          36.036
    # 32611      ZC38 2021-04-19 12:38:17  zc38_0a65_2021_april_19_01.x02.xlsb        116.262          115.048          36.036
    # 32612      ZC38 2021-04-19 12:39:16  zc38_0a65_2021_april_19_01.x01.xlsb        118.724          115.388          36.006
    # 32613      ZC38 2021-04-19 12:39:16  zc38_0a65_2021_april_19_01.x02.xlsb        118.724          115.388          36.006
    # 32614      ZC38 2021-04-19 12:40:16  zc38_0a65_2021_april_19_01.x01.xlsb          120.4          116.829          35.991
    # 32615      ZC38 2021-04-19 12:40:16  zc38_0a65_2021_april_19_01.x02.xlsb          120.4          116.829          35.991
    # 32616      ZC38 2021-04-19 12:41:17  zc38_0a65_2021_april_19_01.x01.xlsb        122.454          115.195          35.982
    # 32617      ZC38 2021-04-19 12:41:17  zc38_0a65_2021_april_19_01.x02.xlsb        122.454          115.195          35.982

df_simultaneous[inspect_cols][-1010:-1000]
    # Out[55]: 
    #       sample_ID           timestamp                                       Source_File HR__aver_(bpm) SBP__aver_(mmHg) aver__aver_(°C)
    # 15826      ZC22 2020-10-27 10:07:34  zc22_0ac7_rx_front-housing_2020_10_26-2.x00.xlsb        109.206          121.358          38.928
    # 15827      ZC22 2020-10-27 10:07:34  zc22_0ac7_rx_front-housing_2020_10_26-2.x01.xlsb        109.206          121.358          38.928
    # 15828      ZC22 2020-10-27 11:07:35  zc22_0ac7_rx_front-housing_2020_10_26-2.x00.xlsb        109.937          123.869          39.007
    # 15829      ZC22 2020-10-27 11:07:35  zc22_0ac7_rx_front-housing_2020_10_26-2.x01.xlsb        109.937          123.869          39.007
    # 15840      ZC22 2020-10-27 12:07:35  zc22_0ac7_rx_front-housing_2020_10_26-2.x00.xlsb         97.789          123.758          38.837
    # 15841      ZC22 2020-10-27 12:07:35  zc22_0ac7_rx_front-housing_2020_10_26-2.x01.xlsb         97.789          123.758          38.837
    # 19343      ZC26 2020-11-09 15:13:38    zc26_0ae4_rx_front-housing_2020_11_09.x00.xlsb        147.406          122.498          36.625
    # 19344      ZC26 2020-11-09 15:13:38    zc26_0ae4_rx_front-housing_2020_11_09.x01.xlsb        147.406          122.498          36.625
    # 19345      ZC26 2020-11-09 16:13:37    zc26_0ae4_rx_front-housing_2020_11_09.x00.xlsb        146.827          114.671          36.859
    # 19346      ZC26 2020-11-09 16:13:37    zc26_0ae4_rx_front-housing_2020_11_09.x01.xlsb        146.827          114.671          36.859


# %%'

# 11
list(df_master.columns)
    # Out[13]: 
    # ['sample_ID',
    #  'timestamp',
    #  'setup',
    #  'timetag',
    #  'timeline',
    #  'BB__aver_(ms)',
    #  'HR__aver_(bpm)',
    #  'DBP__aver_(mmHg)',
    #  'SBP__aver_(mmHg)',
    #  'MBP__aver_(mmHg)',
    #  'aver__aver_(°C)',
    #  'aver__aver_(%)',
    #  'Source_File',
    #  'directory',
    #  'TI_start_date',
    #  'days_since_TI']

# %% err__

# %%% post-sacrifice

mask_post_sacrifice = df_master['days_since_TI'] > 22
df_post_sacrifice = df_master[mask_post_sacrifice]
df_post_sacrifice
    # Out[19]: 
    #       sample_ID           timestamp    setup            timetag timeline  BB__aver_(ms)  HR__aver_(bpm)  DBP__aver_(mmHg)  SBP__aver_(mmHg)  MBP__aver_(mmHg)  \
    # 14067      ZC64 2023-08-21 13:00:00  Surgery  Post_Sacrifice_13      N/A     666.461133       96.739733         51.390267         86.539867         66.121267   
    # 14068      ZC64 2023-08-21 14:00:00  Surgery  Post_Sacrifice_13      N/A     653.043649       93.807491         53.534333         95.839193         71.817877   
    # 14069      ZC64 2023-09-04 08:00:00  Surgery  Post_Sacrifice_27      N/A     595.207529      103.628000         48.645235         71.106941         58.801412   
    # 14070      ZC64 2023-09-04 09:00:00  Surgery  Post_Sacrifice_27      N/A     702.243169       98.352864         62.390254         88.781169         73.526847   
    # 14071      ZC64 2023-09-04 10:00:00  Surgery  Post_Sacrifice_27      N/A     707.271508       93.272017         55.503492         81.216339         66.546763   
    # 14072      ZC64 2023-09-04 11:00:00  Surgery  Post_Sacrifice_27      N/A     609.266767       99.266817         62.233467         94.359633         76.431867   
    # 14073      ZC64 2023-09-04 12:00:00  Surgery  Post_Sacrifice_27      N/A     638.315650       96.435017         62.837050         95.538233         77.096550   
    # 14074      ZC64 2023-09-04 13:00:00  Surgery  Post_Sacrifice_27      N/A     623.082200      115.559783         62.642467         93.188500         75.668650   
    # 14075      ZC64 2023-09-04 14:00:00  Surgery  Post_Sacrifice_27      N/A     573.933217      126.587467         76.899517        111.443583         92.409283   
    # 14076      ZC64 2023-09-04 15:00:00  Surgery  Post_Sacrifice_27      N/A     518.287356      138.945644         75.823978        109.951022         91.381400   
    # 14077      ZC64 2023-09-11 12:00:00  Surgery  Post_Sacrifice_34      N/A     855.690067       72.667433         39.609400         64.483733         49.739433   
    
    #        aver__aver_(°C)  aver__aver_(%)                              Source_File                                                              directory  \
    # 14067        33.412571       99.888040     zc64_1a2d_2023_august_21_01.x00.xlsb  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC64\EMKA\OR   
    # 14068        34.314298       99.969228   zc64_1a2d_2023_august_21_01-2.x00.xlsb  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC64\EMKA\OR   
    # 14069        35.417176       97.003706  zc64_1a2d_2023_september_04_01.x00.xlsb  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC64\EMKA\OR   
    # 14070        35.925317       81.784733  zc64_1a2d_2023_september_04_01.x00.xlsb  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC64\EMKA\OR   
    # 14071        36.321817       89.017200  zc64_1a2d_2023_september_04_01.x00.xlsb  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC64\EMKA\OR   
    # 14072        35.941283       99.802783  zc64_1a2d_2023_september_04_01.x00.xlsb  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC64\EMKA\OR   
    # 14073        35.993833       99.167917  zc64_1a2d_2023_september_04_01.x00.xlsb  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC64\EMKA\OR   
    # 14074        36.112650       94.907217  zc64_1a2d_2023_september_04_01.x00.xlsb  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC64\EMKA\OR   
    # 14075        36.277600       87.502350  zc64_1a2d_2023_september_04_01.x00.xlsb  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC64\EMKA\OR   
    # 14076        36.073870       95.255783  zc64_1a2d_2023_september_04_01.x00.xlsb  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC64\EMKA\OR   
    # 14077        36.127353       97.396944  zc64_1a2d_2023_september_11_01.x00.xlsb  F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\ZC64\EMKA\OR   
    
    #       TI_start_date  days_since_TI  
    # 14067    2023-07-17             35  
    # 14068    2023-07-17             35  
    # 14069    2023-07-17             49  
    # 14070    2023-07-17             49  
    # 14071    2023-07-17             49  
    # 14072    2023-07-17             49  
    # 14073    2023-07-17             49  
    # 14074    2023-07-17             49  
    # 14075    2023-07-17             49  
    # 14076    2023-07-17             49  
    # 14077    2023-07-17             56  


base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER\check")
file_name = 'post_sacrifice'
df_post_sacrifice.to_excel( base_dir / f"{file_name}.xlsx" )  # index=False

# %%%'

mask_post_surgery_TI_1 = ( df_master['setup'] == 'Surgery' ) & ( df_master['timetag'] == 'TI_1' )
df_post_surgery_TI_1 = df_master[mask_post_surgery_TI_1 ]
df_post_surgery_TI_1.shape
    # Out[29]: (48, 16)]

df_post_surgery_TI_1

# %% read

base_dir = Path(r"F:\OneDrive - Uniklinik RWTH Aachen\EMKA\data\copy_excel\MASTER")
file_name = 'Master_Telemetry_Dataset_10'
source_file = base_dir / f"{file_name}.pkl"
df_master_10 = pd.read_pickle(source_file)


df_master_10.shape
    # Out[24]: (38069, 22)

df_master_10['timestamp'][:4]
    # Out[33]: 
    # 0    2020-02-07 15:53:53
    # 4    2020-02-07 16:53:52
    # 8    2020-02-07 17:53:53
    # 12   2020-02-07 18:53:53
    # Name: timestamp, dtype: datetime64[us]

list(df_master_10.columns)
    # Out[26]: 
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
    #  'Source_File',
    #  'directory',
    #  'TI_start_date',
    #  'days_since_TI',
    #  'time_diff']

#=================

df_master_11.shape
    # Out[25]: (16226, 16)

# %%'

