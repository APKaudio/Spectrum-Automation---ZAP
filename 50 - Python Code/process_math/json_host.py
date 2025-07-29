# process_math/json_host.py
#
# This module provides a simple Flask-based JSON API to serve scan data
# from CSV files located in the 'scan_data' directory. It is designed
# to be queried by external applications, such as an Electron-based GUI.
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
import sys
import os
import csv
import inspect
from flask import Flask, jsonify, send_from_directory, request

# --- Path Configuration for Imports and Data Folder ---
# This block ensures that the project root is in sys.path, allowing imports
# from 'utils' or other top-level packages to resolve correctly,
# especially when this script is run directly.

# Get the directory of the current script (json_host.py)
current_script_dir = os.path.dirname(os.path.abspath(__file__))

# Navigate up to the project root.
# From process_math/json_host.py, it's two levels up:
# current_script_dir (process_math) -> parent (src) -> parent (project_root)
project_root = os.path.abspath(os.path.join(current_script_dir, '..', '..'))

# Add the project root to sys.path if it's not already there.
# This makes 'src', 'utils', 'ref', etc., discoverable as top-level packages.
if project_root not in sys.path:
    sys.path.insert(0, project_root)
    # debug_print is not available yet, so use print for initial path setup debug
    print(f"Added {project_root} to sys.path for module discovery.")


# Now that sys.path is configured, import debug_print.
# This import expects 'utils' to be a package directly under project_root.
try:
    from utils.instrument_control import debug_print
except ImportError as e:
    print(f"Failed to import debug_print: {e}")
    print(f"Current sys.path: {sys.path}")
    # Define a dummy debug_print if the import fails, to prevent further errors
    def debug_print(*args, **kwargs):
        pass # Do nothing if debug_print cannot be imported


app = Flask(__name__)

# Define the base directory where scan CSVs are stored.
# This should match the 'output_folder' setting in main_app.py, which defaults to 'scan_data'.
# It's assumed to be a sibling of the 'src' directory, at the project root level.
SCAN_DATA_FOLDER = os.path.join(project_root, 'scan_data')
MARKERS_FILE = os.path.join(project_root, 'ref', 'MARKERS.csv') # Assuming MARKERS.csv is in ref/

# Ensure the scan data directory exists. This is primarily for robustness;
# the main application should create it during scans.
os.makedirs(SCAN_DATA_FOLDER, exist_ok=True)

debug_print(f"JSON API initialized. Serving data from: {SCAN_DATA_FOLDER}", file=__file__, function="init")

def _read_markers_data():
    """
    Reads marker data from MARKERS.csv.
    Assumes MARKERS.csv has a header row.
    Returns a list of dictionaries, where each dictionary is a marker.
    """
    markers_data = []
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not os.path.exists(MARKERS_FILE):
        debug_print(f"Markers file not found: {MARKERS_FILE}", file=current_file, function=current_function)
        return []

    try:
        with open(MARKERS_FILE, 'r', newline='') as csv_file:
            csv_reader = csv.DictReader(csv_file) # Use DictReader to read with headers
            for row in csv_reader:
                # Convert relevant fields to float if they are numeric
                marker = {}
                for key, value in row.items():
                    try:
                        marker[key] = float(value)
                    except ValueError:
                        marker[key] = value # Keep as string if not a number
                markers_data.append(marker)
        debug_print(f"Successfully read {len(markers_data)} markers from {MARKERS_FILE}", file=current_file, function=current_function)
    except Exception as e:
        debug_print(f"Error reading markers file {MARKERS_FILE}: {e}", file=current_file, function=current_function)
    return markers_data


@app.route('/api/scan_data/<filename>', methods=['GET'])
def get_scan_data(filename):
    """
    API endpoint to retrieve scan data from a specified CSV file.
    The data is returned as a JSON object containing headers, scan data, and marker data.

    Example usage:
    GET /api/scan_data/MyScan_RBW10K_HOLD3_Offset0_20231027_103000.csv

    Returns:
        JSON response: An object with "headers", "data", and "markers" arrays, or an error message.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Received request for file: {filename}", file=current_file, function=current_function)

    # Construct the full path to the requested CSV file
    file_path = os.path.join(SCAN_DATA_FOLDER, filename)

    # Validate filename to prevent directory traversal attacks
    if not os.path.abspath(file_path).startswith(os.path.abspath(SCAN_DATA_FOLDER)):
        debug_print(f"Attempted directory traversal detected: {filename}", file=current_file, function=current_function)
        return jsonify({"error": "Invalid filename or path"}), 400

    if not os.path.exists(file_path):
        debug_print(f"File not found: {file_path}", file=current_file, function=current_function)
        return jsonify({"error": "File not found"}), 404

    scan_data = []
    try:
        with open(file_path, 'r', newline='') as csv_file:
            csv_reader = csv.reader(csv_file)
            # As per previous instructions, CSVs are written without headers.
            # The headers "Frequency_MHz" and "Power_dBm" are fixed.
            
            for row in csv_reader:
                if len(row) == 2: # Expecting two columns: Frequency_MHz, Power_dBm
                    try:
                        freq_mhz = float(row[0])
                        power_dbm = float(row[1])
                        scan_data.append([freq_mhz, power_dbm]) # Append as a list of values
                    except ValueError:
                        debug_print(f"Skipping malformed row in {filename}: {row}", file=current_file, function=current_function)
                        continue
                else:
                    debug_print(f"Skipping row with incorrect number of columns in {filename}: {row}", file=current_file, function=current_function)
                    continue

        debug_print(f"Successfully read {len(scan_data)} data points from {filename}", file=current_file, function=current_function)

        # Read marker data
        markers_data = _read_markers_data()

        response_data = {
            "headers": ["Frequency_MHz", "Power_dBm"],
            "data": scan_data,
            "markers": markers_data
        }
        return jsonify(response_data), 200

    except Exception as e:
        debug_print(f"Error reading CSV file {file_path}: {e}", file=current_file, function=current_function)
        return jsonify({"error": f"Error processing file: {e}"}), 500

@app.route('/api/list_scans', methods=['GET'])
def list_scans():
    """
    API endpoint to list all available CSV scan files in the SCAN_DATA_FOLDER.

    Returns:
        JSON response: A list of filenames or an error message.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Received request to list scan files.", file=current_file, function=current_function)

    try:
        scan_files = [f for f in os.listdir(SCAN_DATA_FOLDER) if f.endswith('.csv')]
        # Sort files by modification time, latest first
        scan_files.sort(key=lambda f: os.path.getmtime(os.path.join(SCAN_DATA_FOLDER, f)), reverse=True)
        debug_print(f"Found {len(scan_files)} scan files.", file=current_file, function=current_function)
        return jsonify(scan_files), 200
    except Exception as e:
        debug_print(f"Error listing scan files: {e}", file=current_file, function=current_function)
        return jsonify({"error": f"Error listing files: {e}"}), 500

@app.route('/')
def index():
    """
    Basic endpoint to confirm the API is running.
    """
    return "Scan Data API is running. Use /api/scan_data/<filename> to get data or /api/list_scans to list files."

# --- Running the Flask Application ---
if __name__ == '__main__':
    # This block allows you to run the Flask application directly for testing.
    # In a real Electron application, you would typically start this Python script
    # as a subprocess and communicate with it (e.g., using child_process in Node.js).
    
    # For development, set debug=True to enable Flask's debugger and auto-reloader.
    # For production deployment, always set debug=False for security and performance.
    debug_print(f"Starting JSON API server from {current_script_dir}", file=__file__, function="main")
    debug_print(f"Serving files from {SCAN_DATA_FOLDER}", file=__file__, function="main")
    app.run(debug=False, port=5000) # You can change the port if needed
