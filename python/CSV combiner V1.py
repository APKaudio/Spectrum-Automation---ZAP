#Installation of dependancies
#NB you need to pip install pandas

import os
import pandas as pd
from datetime import datetime

def process_csv_files(folder_path):
    """
    Scans a specified folder for CSV files, merges them into a single file,
    and calculates the average of all columns (except the first two) into another file.

    Args:
        folder_path (str): The path to the folder containing the CSV files.
    """
    print("🔍 Scanning folder for CSV files...")
    # List all files in the given folder and filter for CSV files (case-insensitive)
    csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]
    csv_files.sort() # Sort the files to ensure consistent processing order

    print(f"✅ Found {len(csv_files)} CSV files.")
    if len(csv_files) == 0:
        print("❌ No CSV files found. Exiting function.")
        return # Exit if no CSV files are found

    # Load the first CSV file to establish the base structure (first two columns)
    print(f"📥 Reading first CSV file: {csv_files[0]}")
    # Read CSV, skipping the first row (header), and ensure no header is inferred
    first_df = pd.read_csv(os.path.join(folder_path, csv_files[0]), header=None, skiprows=1)

    # The base DataFrame will contain the first two columns from the first CSV
    base_df = first_df.iloc[:, 0:2].copy()
    column_blocks = [] # List to store DataFrames of additional columns

    # Loop through all found CSV files to gather their data
    for file in csv_files:
        print(f"🔄 Processing: {file}")
        # Read each CSV file, skipping the first row (header), and no header inference
        df = pd.read_csv(os.path.join(folder_path, file), header=None, skiprows=1)

        # Determine the starting column for new data:
        # If it's the first file, start from the 3rd column (index 2) as first two are base.
        # Otherwise, start from the 2nd column (index 1) as other files often repeat first column.
        start_col = 2 if file == csv_files[0] else 1
        new_cols = df.iloc[:, start_col:] # Select relevant columns

        # Assign new column names using the base filename and a numerical suffix
        # This makes the source of the data more explicit in the mashed file.
        new_cols.columns = [f"{os.path.basename(file)}_{i}" for i in range(new_cols.shape[1])]

        column_blocks.append(new_cols) # Add the new columns DataFrame to the list

    # Combine the base DataFrame with all collected column blocks along the column axis
    mashed_df = pd.concat([base_df] + column_blocks, axis=1)

    # Generate a timestamp for unique filenames
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    mashed_filename = f"Mashed-{timestamp}.csv"
    mashed_path = os.path.join(folder_path, mashed_filename)
    print(f"💾 Saving mashed file as: {mashed_filename}")
    # Save the mashed DataFrame to a new CSV file, without the DataFrame index
    mashed_df.to_csv(mashed_path, index=False)

    # Create a DataFrame for the averaged data
    print("🧮 Calculating average of all columns (except the first two)...")
    average_df = pd.DataFrame()
    # Keep the first column from the mashed DataFrame
    average_df[mashed_df.columns[0]] = mashed_df.iloc[:, 0]
    # Calculate the mean of all columns starting from the second one (index 1) across rows
    average_df["Average"] = mashed_df.iloc[:, 1:].mean(axis=1)

    # Save the averaged file
    average_filename = f"Average-{timestamp}.csv"
    average_path = os.path.join(folder_path, average_filename)
    print(f"💾 Saving average file as: {average_filename}")
    # Save the average DataFrame to a new CSV file, without the DataFrame index
    average_df.to_csv(average_path, index=False)

    print("✅ All tasks completed successfully!")

if __name__ == "__main__":
    # IMPORTANT: Replace this with the actual path to your folder
    # Use a raw string (r"...") to avoid issues with backslashes in Windows paths.
    # For example: r"C:\Users\YourUser\Documents\MyCSVFiles"
    # Or for macOS/Linux: "/Users/YourUser/Documents/MyCSVFiles"
    my_folder_path = r"C:\Users\4483\N9340 Scans\IKE - V4 N9340 Scans\N9340 Scans"

    # Call the function to process the CSV files
    process_csv_files(my_folder_path)
