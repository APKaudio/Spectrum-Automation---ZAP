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
    the necessary directory structure if it doesn't exist and conditionally writes
    the header.

    Inputs:
        file_path (str): The full path to the CSV file where the data will be written.
        header (list): A list of strings representing the CSV header row.
                       This header will be written if the file is new or not in append mode.
        data (list): A list of lists or tuples, where each inner list/tuple represents a row of data.
                     Expected format for each data point: (frequency_mhz, level_dbm).
        append_mode (bool, optional): If True, data will be appended to an existing file.
                                      If False, a new file will be created (overwriting if exists).
                                      Defaults to False.
    Process:
        - **Directory Check**: Ensures the output directory exists, creating it if necessary.
        - **File Open Mode**: Determines whether to open in 'w' (write/overwrite) or 'a' (append) mode.
        - **Header Writing**: If the file is opened in 'w' mode (new file or overwrite),
                              the provided `header` is written as the first row.
        - **Data Writing**: Iterates through the `data` and writes each row to the CSV file.
        - **Float Formatting**: Frequencies and amplitudes are formatted to three decimal places
          (`:.3f`) to ensure consistent precision in the CSV output.
    """
    try:
        # Ensure the directory exists before attempting to open the file
        output_dir = os.path.dirname(file_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created directory: {output_dir}")

        # Determine if the file exists before deciding the mode and if header needs to be written
        file_exists = os.path.exists(file_path)
        
        # If not in append_mode, or if in append_mode but file doesn't exist, open in write mode.
        # Otherwise, open in append mode.
        mode = 'a' if append_mode and file_exists else 'w'
        
        # Flag to indicate if header needs to be written
        write_header = (mode == 'w') # Write header if we are creating a new file or overwriting

        with open(file_path, mode, newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            
            if write_header:
                csv_writer.writerow(header)
                print(f"Wrote header to CSV file: {file_path}")
            
            # Write data rows
            for freq_mhz, level_dbm in data:
                csv_writer.writerow([f"{freq_mhz:.3f}", f"{level_dbm:.3f}"])
    except IOError as e:
        print(f"❌ I/O Error writing to CSV file {file_path}: {e}")
        raise # Re-raise to allow calling function to handle
    except Exception as e:
        print(f"❌ An unexpected error occurred while writing to CSV file {file_path}: {e}")
        raise # Re-raise to allow calling function to handle
