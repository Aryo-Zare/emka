

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

# %%

# explore

header_idx
    # Out[61]: np.int64(208)

cpu_date_col_idx
    # Out[60]: np.int64(2)

# %%

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


# %%

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

# %%

# practice : .loc

# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.loc.html

df.loc[["viper", "sidewinder"] , "shield" ]
    # Out[53]: 
    # viper         5
    # sidewinder    8
    # Name: shield, dtype: int64



# %%



