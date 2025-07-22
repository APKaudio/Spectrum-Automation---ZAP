# csv_utils.py

import csv
import os

def write_scan_data_to_csv(file_path, header, data):
    """
    Writes scan data to a CSV file.

    Args:
        file_path (str): The full path to the CSV file to write.
        header (list): A list of strings for the CSV header row.
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

        with open(file_path, 'w', newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow(header) # Write the header row
            
            # Write data rows, converting frequency from Hz to MHz
            # Assuming MHZ_TO_HZ is available or passed, for now, hardcode if not
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

