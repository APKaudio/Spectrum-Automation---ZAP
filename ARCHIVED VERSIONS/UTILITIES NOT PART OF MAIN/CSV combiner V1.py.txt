# NB you need to pip install pandas

import os
import pandas as pd
from datetime import datetime

# Use raw string to avoid issues with backslashes
folder_path = r"C:\Users\4483\N9340 Scans"

print("🔍 Scanning folder for CSV files...")
csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]
csv_files.sort()

print(f"✅ Found {len(csv_files)} CSV files.")
if len(csv_files) == 0:
    print("❌ No CSV files found. Exiting.")
    exit()
# Load the first CSV
print(f"📥 Reading first CSV file: {csv_files[0]}")
first_df = pd.read_csv(os.path.join(folder_path, csv_files[0]), header=None, skiprows=1)

base_df = first_df.iloc[:, 0:2].copy()
column_blocks = []

# Loop through all files and gather additional columns
col_counter = 3
for file in csv_files:
    print(f"🔄 Processing: {file}")
    df = pd.read_csv(os.path.join(folder_path, file), header=None, skiprows=1)


    start_col = 2 if file == csv_files[0] else 1
    new_cols = df.iloc[:, start_col:]

    new_cols.columns = [f"col{col_counter + i}" for i in range(new_cols.shape[1])]
    col_counter += new_cols.shape[1]

    column_blocks.append(new_cols)

# Combine everything in one go
mashed_df = pd.concat([base_df] + column_blocks, axis=1)

# Save mashed file
timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
mashed_filename = f"Mashed-{timestamp}.csv"
mashed_path = os.path.join(folder_path, mashed_filename)
print(f"💾 Saving mashed file as: {mashed_filename}")
mashed_df.to_csv(mashed_path, index=False)

# Create averaged file
print("🧮 Calculating average of all columns (except the first)...")
average_df = pd.DataFrame()
average_df[mashed_df.columns[0]] = mashed_df.iloc[:, 0]  # Keep first column
average_df["Average"] = mashed_df.iloc[:, 1:].mean(axis=1)

# Save average file
average_filename = f"Average-{timestamp}.csv"
average_path = os.path.join(folder_path, average_filename)
print(f"💾 Saving average file as: {average_filename}")
average_df.to_csv(average_path, index=False)

print("✅ All tasks completed successfully!")
