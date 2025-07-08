import os
import pandas as pd

def combine_csv_files(directory_path, output_filename="STICHED.csv"):
    """
    Combines all CSV files in a given directory into a single CSV file.
    It ignores headers in input CSVs and does not create a header in the output CSV.

    Args:
        directory_path (str): The path to the directory containing the CSV files.
        output_filename (str): The name of the output CSV file.
    """
    all_files = []
    combined_df = pd.DataFrame() # Initialize an empty DataFrame to store combined data

    print(f"Searching for CSV files in: {directory_path}")

    # Check if the directory exists
    if not os.path.isdir(directory_path):
        print(f"Error: Directory not found at '{directory_path}'")
        return

    # List all files in the directory
    try:
        files_in_directory = os.listdir(directory_path)
    except OSError as e:
        print(f"Error accessing directory '{directory_path}': {e}")
        return

    # Filter for CSV files
    csv_files = [f for f in files_in_directory if f.endswith('.csv')]

    if not csv_files:
        print(f"No CSV files found in the directory: {directory_path}")
        return

    print(f"Found {len(csv_files)} CSV files.")

    # Read each CSV file and append its DataFrame to a list, ignoring headers
    for filename in csv_files:
        file_path = os.path.join(directory_path, filename)
        try:
            print(f"Reading {filename} (ignoring header)...")
            # Read CSV without a header (header=None)
            df = pd.read_csv(file_path, header=None)
            all_files.append(df)
        except pd.errors.EmptyDataError:
            print(f"Warning: {filename} is empty and will be skipped.")
        except Exception as e:
            print(f"Error reading {filename}: {e}")

    if not all_files:
        print("No valid CSV files were read to combine.")
        return

    # Concatenate all DataFrames into a single DataFrame
    try:
        combined_df = pd.concat(all_files, ignore_index=True)
        print("All CSV files combined successfully.")
    except Exception as e:
        print(f"Error during concatenation: {e}")
        return

    # Define the output file path
    output_file_path = os.path.join(directory_path, output_filename)

    # Save the combined DataFrame to a new CSV file without a header (header=False)
    try:
        combined_df.to_csv(output_file_path, index=False, header=False)
        print(f"Combined CSV saved to: {output_file_path}")
    except Exception as e:
        print(f"Error saving combined CSV to {output_file_path}: {e}")

# Specify the directory path
# IMPORTANT: Make sure this path is correct on your system
directory = r"C:\Users\4483\N9340 Scans\Trad Scans\New folder"

# Call the function to combine the CSV files
if __name__ == "__main__":
    combine_csv_files(directory)
