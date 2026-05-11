

# %%

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# %%


# --- 1. FIND A CANDIDATE PIG ---
print("Finding recording durations for each pig...")

# Calculate the start and end times for each pig
durations = df_master.groupby('sample_ID')['timestamp'].agg(['min', 'max'])
durations['duration_hours'] = (durations['max'] - durations['min']).dt.total_seconds() / 3600

# Filter for pigs with at least 24 hours of data
good_candidates = durations[durations['duration_hours'] >= 24]

type(good_candidates)
    # Out[16]: pandas.DataFrame

good_candidates.shape
    # Out[17]: (45, 3)

good_candidates[:10]
    # Out[19]: 
    #                           min                 max  duration_hours
    # sample_ID                                                        
    # Unknown   2020-03-15 11:35:17 2020-08-24 15:35:03     3891.996111
    # ZC04      2020-02-07 15:53:53 2020-02-25 09:25:14      425.522500
    # ZC05      2020-02-25 09:38:24 2020-03-18 12:40:43      531.038611
    # ZC06      2020-02-25 13:02:46 2020-03-16 15:08:38      482.097778
    # ZC07      2020-05-25 08:49:39 2020-06-16 13:09:55      532.337778
    # ZC08      2020-05-25 10:07:26 2020-06-16 15:50:50      533.723333
    # ZC09      2020-06-15 09:30:41 2020-07-07 18:27:45      536.951111
    # ZC10      2020-06-15 12:58:57 2020-07-07 18:27:45      533.480000
    # ZC11      2020-06-29 17:02:57 2020-07-21 13:03:31      524.009444
    # ZC12      2020-06-29 18:28:24 2020-07-14 14:48:48      356.340000

# %%



# --- 1. SLICE DATA FOR ZC12 ---
chosen_pig = 'ZC12'
print(f"Extracting data for {chosen_pig}...")

df_pig = df_master[df_master['sample_ID'] == chosen_pig].copy()

# ---> THE FIX: Force the graphing columns to be pure numbers <---
numeric_cols = ['HR__aver_(bpm)', 'SBP__aver_(mmHg)', 'DBP__aver_(mmHg)', 'MBP__aver_(mmHg)']
for col in numeric_cols:
    df_pig[col] = pd.to_numeric(df_pig[col], errors='coerce')

df_pig = df_pig.sort_values('timestamp')

# --- 2. FIND THE FIRST MIDNIGHT ---
first_timestamp = df_pig['timestamp'].iloc[0]
start_time = first_timestamp.normalize() + pd.Timedelta(days=1)
end_time = start_time + pd.Timedelta(hours=24)

print(f"Plotting 24h cycle from {start_time} to {end_time}...")

# Slice the 24-hour window
df_24h = df_pig[(df_pig['timestamp'] >= start_time) & (df_pig['timestamp'] < end_time)]

if df_24h.empty:
    print(f"Warning: No data found for {chosen_pig} in this exact 24-hour window.")
else:
    # --- 3. CREATE THE DUAL-AXIS PLOT ---
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax1 = plt.subplots(figsize=(12, 6))

    # Plot Heart Rate (LEFT axis)
    color_hr = '#d62728'
    ax1.plot(df_24h['timestamp'], df_24h['HR__aver_(bpm)'], 
             color=color_hr, 
             linewidth=5,              # Made the line significantly thicker
             # marker='o',                 # Added circular dots at each exact data point
             # markersize=6,               # Size of the dots
             # markeredgecolor='white',    # Adds a crisp white border around the dots to break up the blue background
             label='Heart Rate')
    
    ax1.set_xlabel(f'Time of Day (Date: {start_time.strftime("%d %b %Y")})', fontsize=12, fontweight='bold', labelpad=10)
    ax1.set_ylabel('Heart Rate (bpm)', color=color_hr, fontsize=12, fontweight='bold')
    ax1.tick_params(axis='y', labelcolor=color_hr, labelsize=11)

    # Plot Blood Pressure (RIGHT axis)
    ax2 = ax1.twinx()
    color_bp = '#1f77b4'
    sbp_col = 'SBP__aver_(mmHg)'
    dbp_col = 'DBP__aver_(mmHg)'
    mbp_col = 'MBP__aver_(mmHg)'

    # Fill and plot lines
    ax2.fill_between(df_24h['timestamp'], df_24h[dbp_col], df_24h[sbp_col], color=color_bp, alpha=0.15)
    ax2.plot(df_24h['timestamp'], df_24h[mbp_col], color=color_bp, linewidth=2, label='Mean BP')
    ax2.plot(df_24h['timestamp'], df_24h[sbp_col], color='#000000', linewidth=0.8, linestyle=':', alpha=0.6)
    ax2.plot(df_24h['timestamp'], df_24h[dbp_col], color='#000000', linewidth=0.8, linestyle=':', alpha=0.6)

    ax2.set_ylabel('Blood Pressure (mmHg)', color=color_bp, fontsize=12, fontweight='bold')
    ax2.tick_params(axis='y', labelcolor=color_bp, labelsize=11)

    # Titles and formatting
    plt.title(f'24-Hour Telemetry Profile (Pig: {chosen_pig})', fontsize=16, fontweight='bold', pad=15)
    ax1.set_xlim([start_time, end_time])
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
    ax1.xaxis.set_major_locator(mdates.HourLocator(interval=2))
    plt.xticks(fontsize=11)

    # Legends and Grid
    lines_1, labels_1 = ax1.get_legend_handles_labels()
    lines_2, labels_2 = ax2.get_legend_handles_labels()
    ax1.legend(lines_1 + lines_2, labels_1 + labels_2, loc='upper left', frameon=True, shadow=True)
    ax2.grid(False)

    # Save and show
    plt.tight_layout()

    print(f"Done! Presentation plot saved to:\n{plot_filename}")
    plt.show()

# %%

# if you put 1 '\' at the end of the string  =>  
    # SyntaxError: unterminated string literal (detected at line 1); perhaps you escaped the end quote?
plot_path = r'F:\OneDrive - Uniklinik RWTH Aachen\EMKA\analysis__EMKA\plot__scand-LAS\\'
plot_name = f'DualAxis_24h_{chosen_pig}_2'
plt.savefig(plot_path + plot_name + '.pdf')
plt.savefig(plot_path + plot_name + '.svg')
    

# %%

