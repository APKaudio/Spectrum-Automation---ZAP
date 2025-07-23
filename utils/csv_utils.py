# csv_utils.py
#
# This module provides utility functions for writing spectrum scan data to CSV files.
# It encapsulates the logic for handling file paths, directory creation, and data formatting
# for CSV output, ensuring consistent data storage for analysis and historical tracking.
#
# Author: Anthony Peter Kuzub
# Blog: www.Like.audio (Contributor to this project)
#
# Professional services for customizing and tailoring this software to your specific
# application can be negotiated. There is no change to use, modify, or fork this software.
#
# Build Log: https://like.audio/category/software/spectrum-scanner/
# Source Code: https://github.com/APKaudio/
#
import csv
import os

def write_scan_data_to_csv(file_path, header, data, append_mode=False):
    """
    Writes scan data to a CSV file. This function is designed to write raw frequency
    and amplitude data collected from the spectrum analyzer. It handles creating
    the necessary directory structure if it doesn't exist.

    Inputs:
        file_path (str): The full path to the CSV file where the data will be written.
        header (list): A list of strings representing the CSV header row.
                       Note: This function *never* writes the header itself.
                       It's included for compatibility but is ignored.
        data (list): A list of lists or tuples, where each inner list/tuple represents a row of data.
                     Expected format for each data point: `(frequency_mhz, amplitude_dbm)`.
        append_mode (bool): If True, data will be appended to the file if it exists.
                            If False (default), the file will be overwritten.

    Process:
        1. **Directory Creation**: Extracts the directory path from `file_path`. If the directory
           does not exist, it creates all necessary parent directories using `os.makedirs(exist_ok=True)`.
        2. **File Open Mode**: Determines the file opening mode (`'a'` for append, `'w'` for write/overwrite)
           based on `append_mode` and whether the file already exists.
        3. **CSV Writing**: Opens the CSV file with the determined mode and `newline=''` (important for CSVs).
           Initializes a `csv.writer`.
        4. **Data Iteration and Write**: Iterates through each `(freq_mhz, level_dbm)` pair in the `data` list.
           For each pair, it formats the frequency and level to three decimal places and writes them as a row
           to the CSV file.
        5. **Error Handling**: Includes `try-except` blocks to catch `IOError` (e.g., permission issues, disk full)
           and general `Exception` during file operations, re-raising them to allow the calling function
           to handle recovery or display a user-friendly error.

    Outputs:
        None. (Side effect: Creates or modifies a CSV file at the specified `file_path`.
               Prints messages to console for directory creation and errors.)

    Notes:
        - **Header Handling**: The `header` argument is explicitly ignored by this function.
          This design choice assumes that if a header is needed, it is handled by the
          calling context (e.g., written once when the file is first created) or that
          the file is intended for raw data only without a header. This simplifies
          the `write_scan_data_to_csv` function, making it purely for data rows.
        - **Float Formatting**: Frequencies and amplitudes are formatted to three decimal places
          (`:.3f`) to ensure consistent precision in the CSV output.
    """
    try:
        # Ensure the directory exists before attempting to open the file
        output_dir = os.path.dirname(file_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created directory: {output_dir}")

        # The header is now explicitly NOT written by this function.
        # It's assumed that if a header is needed, it's handled by the calling context
        # or the file is for raw data only.
        mode = 'a' if append_mode and os.path.exists(file_path) else 'w'

        with open(file_path, mode, newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            
            # Write data rows
            for freq_mhz, level_dbm in data:
                csv_writer.writerow([f"{freq_mhz:.3f}", f"{level_dbm:.3f}"])
    except IOError as e:
        print(f"❌ I/O Error writing to CSV file {file_path}: {e}")
        raise # Re-raise to allow calling function to handle
    except Exception as e:
        print(f"❌ An unexpected error occurred while writing to CSV {file_path}: {e}")
        raise # Re-raise to allow calling function to handle
