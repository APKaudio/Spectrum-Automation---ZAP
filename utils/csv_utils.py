# csv_utils.py

import csv
import os

def write_scan_data_to_csv(file_path, data, write_header=False):
    """
    Writes scan data to a CSV file. If write_header is True, it will write the header.
    This function is primarily for initial file creation or full overwrite.

    Args:
        file_path (str): The full path to the CSV file to write.
        data (list): A list of lists, where each inner list is a row of data.
                     Expected format for data points: (frequency_hz, amplitude_dbm).
                     This function will convert frequency to MHz for the CSV.
        write_header (bool): If True, writes the header row. Defaults to False.
    """
    try:
        # Ensure the directory exists before attempting to open the file
        output_dir = os.path.dirname(file_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created directory: {output_dir}")

        # Define header if needed for initial write (even if not written to file)
        header = ["Frequency (MHz)", "Level (dBm)"]

        with open(file_path, 'w', newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            if write_header:
                csv_writer.writerow(header)

            # Write data rows, converting frequency from Hz to MHz
            MHZ_TO_HZ = 1_000_000 # Define locally if not imported globally

            for freq_hz, level_dbm in data:
                csv_writer.writerow([f"{freq_hz / MHZ_TO_HZ:.3f}", f"{level_dbm:.3f}"])
        print(f"✅ Data successfully written to CSV: {file_path}")
    except IOError as e:
        print(f"❌ I/O Error writing to CSV file {file_path}: {e}")
        raise # Re-raise to allow calling function to handle
    except Exception as e:
        print(f"❌ An unexpected error occurred while writing CSV: {e}")
        raise # Re-raise to allow calling function to handle

def append_scan_data_to_csv(file_path, data):
    """
    Appends scan data to an existing CSV file. If the file does not exist,
    it creates it without a header.

    Args:
        file_path (str): The full path to the CSV file to append to.
        data (list): A list of lists, where each inner list is a row of data.
                     Expected format for data points: (frequency_hz, amplitude_dbm).
                     This function will convert frequency to MHz for the CSV.
    """
    try:
        # Ensure the directory exists before attempting to open the file
        output_dir = os.path.dirname(file_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created directory: {output_dir}")

        # Check if file exists to determine if header needs to be written (only for first write)
        file_exists = os.path.exists(file_path)

        with open(file_path, 'a', newline='') as csv_file: # Open in append mode
            csv_writer = csv.writer(csv_file)

            # No header written when appending, as per user's previous request for no headers in CSVs.
            # If a header was ever desired for the *very first* write, that logic would go here,
            # but for continuous writing, we assume no header is desired for any part.

            # Write data rows, converting frequency from Hz to MHz
            MHZ_TO_HZ = 1_000_000 # Define locally if not imported globally

            for freq_hz, level_dbm in data:
                csv_writer.writerow([f"{freq_hz / MHZ_TO_HZ:.3f}", f"{level_dbm:.3f}"])
        # print(f"✅ Data successfully appended to CSV: {file_path}") # This can be noisy, keep only for debugging
    except IOError as e:
        print(f"❌ I/O Error appending to CSV file {file_path}: {e}")
        raise # Re-raise to allow calling function to handle
    except Exception as e:
        print(f"❌ An unexpected error occurred while appending CSV: {e}")
        raise # Re-raise to allow calling function to handle

