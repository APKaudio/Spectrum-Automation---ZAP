# csv_utils.py

import csv
import os

def write_scan_data_to_csv(file_path, header, data, append_mode=False):
    """
    Writes scan data to a CSV file.

    Args:
        file_path (str): The full path to the CSV file to write.
        header (list): A list of strings for the CSV header row (ignored if append_mode is True).
        data (list): A list of lists, where each inner list is a row of data.
                     Expected format for data points: (frequency_mhz, amplitude_dbm).
        append_mode (bool): If True, appends to the file. If False (default), overwrites.
                            Header is NEVER written by this function.
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
