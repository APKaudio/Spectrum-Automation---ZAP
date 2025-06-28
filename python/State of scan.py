import pyvisa
import time
import argparse # Import argparse for command-line argument parsing
import struct # Import struct for unpacking binary data
import numpy as np # For numerical operations, especially mean (already imported)
import os # Import os for directory creation and path manipulation
import csv # Import csv for writing CSV files
from datetime import datetime # Import datetime for timestamping files
import pandas as pd # Import pandas for DataFrame operations
import plotly.express as px # Import plotly for interactive plotting

# Define constants for better readability and easier modification
MHZ_TO_HZ = 1_000_000  # Conversion factor from MHz to Hz
DEFAULT_SEGMENT_WIDTH_HZ = 10_000_000  # 10 MHz segment for sweeping (this constant is now less relevant for auto-segmenting)
DEFAULT_RBW_STEP_SIZE_HZ = 10000  # 10 kHz RBW resolution desired per data point
DEFAULT_CYCLE_WAIT_TIME_SECONDS = 10  # 10 seconds wait between full scan cycles
DEFAULT_MAXHOLD_TIME_SECONDS = 3 # Default max hold time for the new argument


# Define the frequency bands to scan
# Frequencies are stored in MHz for readability, will be converted to Hz for instrument commands
BAND_RANGES = [
    {"Band Name": "Low VHF+FM", "Start MHz": 50.000, "Stop MHz": 110.000},
    {"Band Name": "High VHF+216", "Start MHz": 170.000, "Stop MHz": 220.000},
    {"Band Name": "UHF -1", "Start MHz": 400.000, "Stop MHz": 700.000},
    {"Band Name": "UHF -2", "Start MHz": 700.000, "Stop MHz": 900.000},
    {"Band Name": "900 ISM-STL", "Start MHz": 900.000, "Stop MHz": 970.000},
    {"Band Name": "AFTRCC-1", "Start MHz": 1430.000, "Stop MHz": 1540.000},
    {"Band Name": "DECT-ALL", "Start MHz": 1880.000, "Stop MHz": 2000.000},
    {"Band Name": "2 GHz Cams", "Start MHz": 2000.000, "Stop MHz": 2390.000},
]

# Define VISA addresses for different users/instruments
VISA_ADDRESSES = {
    'apk': 'USB0::0x0957::0xFFEF::CN03480580::0::INSTR',
    'zap': 'USB1::0x0957::0xFFEF::SG05300002::0::INSTR' # Original user provided 'USB1', keeping it.
}

def query_safe(inst, command):
    """
    Safely queries the instrument, handling PyVISA errors.
    Args:
        inst (pyvisa.resources.Resource): The PyVISA instrument object.
        command (str): The SCPI command to query.
    Returns:
        str: The response from the instrument, stripped of whitespace,
             or "[Not Supported or Timeout]" if an error occurs.
    """
    try:
        return inst.query(command).strip()
    except pyvisa.VisaIOError:
        return "[Not Supported or Timeout]"

def write_safe(inst, command):
    """
    Safely writes a command to the instrument, handling PyVISA errors.
    Args:
        inst (pyvisa.resources.Resource): The PyVISA instrument object.
        command (str): The SCPI command to write.
    Returns:
        bool: True if the write was successful, False otherwise.
    """
    try:
        inst.write(command)
        return True
    except pyvisa.VisaIOError as e:
        print(f"Error writing command '{command}': {e}")
        return False

def setup_arguments():
    """
    Parses command-line arguments for the spectrum analyzer sweep.
    Allows customization of filename prefix, frequency range, step size, and user.
    """
    parser = argparse.ArgumentParser(description="Spectrum Analyzer Sweep and CSV Export")

    parser.add_argument('--SCANname', type=str, default="25kz scan ",
                        help='Prefix for the output CSV filename')
    parser.add_argument('--startFreq', type=float, default=None,
                        help='Start frequency in Hz (overrides default bands if provided with --endFreq)')
    parser.add_argument('--endFreq', type=float, default=None,
                        help='End frequency in Hz (overrides default bands if provided with --startFreq)')
    parser.add_argument('--stepSize', type=float, default=DEFAULT_RBW_STEP_SIZE_HZ, # Updated to new variable name
                        help='Step size in Hz')
    parser.add_argument('--user', type=str, choices=['apk', 'zap'], default='zap',
                        help='Specify who is running the program: "apk" or "zap". Default is "zap".')
    parser.add_argument('--maxHoldTime', type=float, default=DEFAULT_MAXHOLD_TIME_SECONDS,
                        help='Duration in seconds for which MAX Hold should be active during scans. Set to 0 to disable. (Note: Instrument\'s MAX Hold is typically a continuous mode; this value serves as a flag for enablement during the entire scan duration).')

    return parser.parse_args()

def initialize_instrument(visa_address):
    """
    Establishes a VISA connection to the instrument and performs initial configuration.

    Args:
        visa_address (str): The VISA address of the spectrum analyzer.

    Returns:
        pyvisa.resources.Resource: The instrument object if connection is successful, else None.
    """
    rm = pyvisa.ResourceManager()
    inst = None
    try:
        inst = rm.open_resource(visa_address)
        inst.timeout = 5000  # Set timeout to 5 seconds for queries
        print(f"\nSuccessfully connected to: {query_safe(inst, '*IDN?')}\n")

        # Clear and reset the instrument
        write_safe(inst, "*CLS")
        write_safe(inst, "*RST")
        query_safe(inst, "*OPC?") # Wait for operations to complete
        print("Instrument cleared and reset.")

        # Configure preamplifier for high sensitivity
        write_safe(inst, ":SENS:POW:GAIN ON") # Equivalent to ':POWer:GAIN 1' for most Keysight instruments
        print("Preamplifier turned ON for high sensitivity.")

        # Configure display and marker settings
        write_safe(inst, ":DISP:WIND:TRAC:Y:SCAL LOG") # Logarithmic scale (dBm)
        write_safe(inst, ":DISP:WIND:TRAC:Y:SCAL:RLEVel -30DBM") # Reference level (corrected from :DISP:WIND:TRAC:Y:RLEV)

        # Turn on all six markers and set them to position mode (neutral state)
        # They will be used by the peak table functionality per segment.
        # Note: N9340B typically supports 4 markers. Attempting 6, but instrument might ignore extra.
        for i in range(1, 7): # Markers 1 through 6
            write_safe(inst, f":CALC:MARK{i}:STATe ON")
            write_safe(inst, f":CALC:MARK{i}:MODE POS") 
            print(f"Marker {i} enabled in position mode (for peak table display).")
        
        print("Display set to logarithmic scale, reference level -30 dBm, and Markers 1-6 enabled.")

        time.sleep(2) # Allow instrument to settle
        return inst
    except pyvisa.VisaIOError as e:
        print(f"VISA Error: Could not connect to or communicate with the instrument at {visa_address}: {e}")
        print("Please ensure the instrument is connected, powered on, and the VISA address is correct.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred during instrument initialization: {e}")
        return None

def scan_bands(inst, csv_writer, max_hold_time):
    """
    Iterates through predefined frequency bands, sets the start/stop frequencies,
    reduces RBW to 10000 Hz, and triggers a sweep for each band.
    It extracts trace data, writes it directly to the provided CSV writer,
    and also returns it for further processing (plotting).
    This function now dynamically segments bands to maintain a consistent
    effective resolution bandwidth per trace point.

    Args:
        inst (pyvisa.resources.Resource): The PyVISA instrument object.
        csv_writer (csv.writer): The CSV writer object to write data to.
        max_hold_time (float): Duration in seconds for which MAX Hold should be active.
                                If > 0, MAX Hold mode is enabled for the scan.
    Returns:
        list: A list of dictionaries, where each dictionary represents a data point
              with 'Band Name', 'Frequency (Hz)', and 'Level (dBm)'.
    """
    all_scan_data = [] # To store all data points across all bands for plotting

    print("\n--- Starting Band Scan ---")
    
    # Configure MAX Hold mode based on max_hold_time argument
    if max_hold_time > 0:
        write_safe(inst, ":TRAC1:MODE MAXHold")
        print("Trace 1 set to MAX Hold mode for the duration of the scan.")
    else:
        # Ensure normal/clear write mode if max hold is not requested
        write_safe(inst, ":TRAC1:MODE WRITe")
        print("Trace 1 set to normal WRITE mode (MAX Hold disabled).")


    try:
        actual_sweep_points = int(float(query_safe(inst, ":SENS:SWE:POINts?")))
        print(f"Actual sweep points set by instrument: {actual_sweep_points}.")
    except ValueError:
        print(f"  Could not parse actual sweep points from instrument response: '{query_safe(inst, ":SENS:SWE:POINts?")}'. Defaulting to 401.")
        actual_sweep_points = 401 # Fallback to a common default if query fails

    if actual_sweep_points <= 1: # Ensure we have at least 2 points to calculate a span
        print(f"Warning: Instrument returned {actual_sweep_points} sweep points. Cannot effectively segment bands. Skipping scan.")
        return []

    # Calculate the optimal span for each segment to achieve desired RBW per point
    # We want (Segment Span / (Actual Points - 1)) = Desired RBW
    # So, Segment Span = Desired RBW * (Actual Points - 1)
    optimal_segment_span_hz = DEFAULT_RBW_STEP_SIZE_HZ * (actual_sweep_points - 1)
    print(f"Optimal segment span to achieve {DEFAULT_RBW_STEP_SIZE_HZ/1000:.0f} kHz effective RBW per point: {optimal_segment_span_hz / MHZ_TO_HZ:.3f} MHz.")


    # Set trace format to REAL (binary)
    write_safe(inst, ":TRAC:FORMat REAL")
    print("Set trace data format to REAL (binary) for efficient data transfer.")

    # Save current read_termination and encoding
    original_read_termination = inst.read_termination
    original_encoding = inst.encoding
    
    # Set read_termination to empty string for raw binary data reads.
    # This is crucial to prevent truncation if binary data contains bytes
    # that could be interpreted as termination characters.
    inst.read_termination = ''
    # Use latin-1 or iso-8859-1 for raw byte decoding if necessary,
    # though for binary, decoding should happen after struct.unpack
    inst.encoding = 'latin-1' 

    for band in BAND_RANGES:
        band_name = band["Band Name"]
        band_start_freq_hz = band["Start MHz"] * MHZ_TO_HZ
        band_stop_freq_hz = band["Stop MHz"] * MHZ_TO_HZ

        print(f"\nProcessing Band: {band_name} (Total Range: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {band_stop_freq_hz/MHZ_TO_HZ:.3f} MHz)")

        current_segment_start_freq_hz = band_start_freq_hz
        segment_counter = 0

        while current_segment_start_freq_hz < band_stop_freq_hz:
            segment_counter += 1
            segment_stop_freq_hz = min(current_segment_start_freq_hz + optimal_segment_span_hz, band_stop_freq_hz)
            actual_segment_span_hz = segment_stop_freq_hz - current_segment_start_freq_hz

            if actual_segment_span_hz <= 0: # Avoid infinite loop if start == stop or negative span
                break
            
            # If the last segment is very small, it might result in less than 2 points,
            # which could cause issues with frequency step calculation.
            # We ensure minimum span if points are fixed.
            if actual_sweep_points > 1 and actual_segment_span_hz < (DEFAULT_RBW_STEP_SIZE_HZ * (actual_sweep_points - 1)):
                if segment_stop_freq_hz == band_stop_freq_hz: # This is the very last segment
                     pass # Allow smaller span for last segment, it will have fewer effective points but covers the end
                else:
                    # For intermediate segments, if the calculated span is too small, skip to next full segment
                    current_segment_start_freq_hz += optimal_segment_span_hz
                    continue # Skip this tiny segment, move to next potential full segment

            print(f"  Scanning Segment {segment_counter}: {current_segment_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz (Span: {actual_segment_span_hz/MHZ_TO_HZ:.3f} MHz)")
            
            # Set instrument frequency range for the current segment
            write_safe(inst, f":SENS:FREQ:STAR {current_segment_start_freq_hz}")
            write_safe(inst, f":SENS:FREQ:STOP {segment_stop_freq_hz}")
            # Explicitly set span to ensure the instrument uses it for the sweep
            write_safe(inst, f":SENS:FREQ:SPAN {actual_segment_span_hz}")

            time.sleep(0.1) # Small delay to ensure commands are processed
            
            # Add settling time for max hold values to show up, if max hold is enabled
            if max_hold_time > 0:
                print(f"Waiting {max_hold_time} seconds for MAX hold to settle...", end='')
                for i in range(int(max_hold_time), 0, -1):
                    print(f"\rWaiting {i} seconds for MAX hold to settle...   ", end='') # \r to overwrite line
                    time.sleep(1)
                print("\rMAX hold settle time complete.                            ") # Clear the line after countdown

            query_safe(inst, "*OPC?") # Wait for the sweep to complete
            print(f"  Sweep completed for segment {segment_counter}.")

            # Read and process trace data
            trace_data = []
            raw_bytes = b''
            try:
                inst.write(":TRAC1:DATA?")
                raw_bytes = inst.read_raw()
                
                # Assuming raw binary float32 values without IEEE 488.2 header
                num_values_received = len(raw_bytes) // 4
                if num_values_received == 0:
                    print("    No trace data bytes received for this segment.")
                    current_segment_start_freq_hz = segment_stop_freq_hz
                    continue # Move to the next segment if no data

                trace_data = list(struct.unpack('<' + 'f' * num_values_received, raw_bytes))

                # --- Peak Table Implementation ---
                # 1. Enable Peak Table State
                write_safe(inst, ":CALC:MARK:PEAK:TABLe:STATE ON")
                # 2. Set Peak Table Threshold and Max Peaks
                write_safe(inst, ":CALC:MARK:PEAK:TABLe:THReshold -90DBM") # Set threshold to -90 dBm
                write_safe(inst, ":CALC:MARK:PEAK:TABLe:MAX 6") # Limit to max 6 peaks
                # 3. Generate Peak Table (find all peaks based on instrument settings)
                write_safe(inst, ":CALC:MARK:PEAK:TABLe:ALL")
                # 4. Query Peak Table Data
                peak_table_data_str = query_safe(inst, ":CALC:MARK:PEAK:TABLe:DATA?")

                segment_found_peaks = []
                # Parse the peak table data string (e.g., "F1,A1,F2,A2,...")
                if peak_table_data_str and peak_table_data_str != "[Not Supported or Timeout]":
                    try:
                        peak_values = [float(x) for x in peak_table_data_str.split(',')]
                        for j in range(0, len(peak_values), 2):
                            if j + 1 < len(peak_values):
                                segment_found_peaks.append({"freq": peak_values[j], "level": peak_values[j+1]})
                    except ValueError:
                        print(f"    Warning: Could not parse peak table data string: '{peak_table_data_str}'.")

                # Assign the found peaks from the peak table to the available markers (up to 6)
                # The markers were already turned ON and set to POS mode in initialize_instrument
                for idx, peak_info in enumerate(segment_found_peaks[:6]): # Take up to the first 6 peaks
                    marker_num = idx + 1
                    write_safe(inst, f":CALC:MARK{marker_num}:X {peak_info['freq']}")
                    write_safe(inst, f":CALC:MARK{marker_num}:Y {peak_info['level']}")
                
                # Optionally, turn off unused markers if fewer than 6 peaks are found
                for idx in range(len(segment_found_peaks), 6):
                    marker_num = idx + 1
                    write_safe(inst, f":CALC:MARK{marker_num}:STATe OFF")

                # 5. Disable Peak Table State after using its data to assign markers
                write_safe(inst, ":CALC:MARK:PEAK:TABLe:STATE OFF")
                
                if segment_found_peaks:
                    print(f"    Peaks found in segment {segment_counter}:")
                    for p in segment_found_peaks:
                        print(f"      Frequency: {p['freq']/MHZ_TO_HZ:.3f} MHz, Level: {p['level']:.2f} dBm")
                else:
                    print("    No peaks found in this segment's peak table.")


                # Use the actual number of points received for frequency calculation for this segment
                num_trace_points_actual = len(trace_data)
                
                if num_trace_points_actual > 1:
                    freq_step_per_point_actual = actual_segment_span_hz / (num_trace_points_actual - 1)
                elif num_trace_points_actual == 1:
                    freq_step_per_point_actual = 0 # Single point, no step
                else:
                    freq_step_per_point_actual = 0 # No points

                # Loop to append data to all_scan_data and write to CSV
                for i, amp_value in enumerate(trace_data):
                    current_freq_for_point_hz = current_segment_start_freq_hz + (i * freq_step_per_point_actual)
                    
                    peak_indicator = ""
                    # Check if this point is a peak from the peak table
                    for found_peak in segment_found_peaks:
                        # Use a small tolerance for floating point comparison (e.g., half the step size)
                        if abs(current_freq_for_point_hz - found_peak["freq"]) < (freq_step_per_point_actual / 2.0):
                            peak_indicator = "Peak from Table"
                            break # Mark only once if multiple peaks are very close

                    # Append to list for plotting later
                    all_scan_data.append({
                        "Frequency (MHz)": current_freq_for_point_hz / MHZ_TO_HZ, # Store in MHz for DataFrame consistency
                        "Level (dBm)": amp_value,
                        "Band Name": band_name,
                        "Peak Indicator": peak_indicator # Add the new column
                    })
                    
                    # Write directly to CSV file with desired order and units
                    csv_writer.writerow([
                        f"{current_freq_for_point_hz / MHZ_TO_HZ:.2f}",  # Frequency in MHz
                        f"{amp_value:.2f}",                               # Level in dBm
                        band_name,                                        # Band Name
                        peak_indicator                                    # Peak Indicator
                    ])
                
            except pyvisa.VisaIOError as e:
                print(f"    Error reading trace data (PyVISA IO Error): {e}")
                print(f"    Raw bytes potentially causing error: {raw_bytes[:100]}...")
            except ValueError as e:
                print(f"    Error processing binary data (ValueError): {e}")
                print(f"    Raw bytes potentially causing error: {raw_bytes[:100]}...")
            except struct.error as e:
                print(f"    Error unpacking binary data (Struct Error - incorrect format/length): {e}")
                print(f"    Raw bytes for struct unpack: {raw_bytes[:100]}...")
            except Exception as e:
                print(f"    An unexpected error occurred during trace processing: {e}")

            current_segment_start_freq_hz = segment_stop_freq_hz # Move to the start of the next segment

    # Restore original read_termination and encoding
    inst.read_termination = original_read_termination
    inst.encoding = original_encoding
    print("\n--- Band Scan Complete ---")
    return all_scan_data # Return the collected data


def plot_spectrum_data(df, output_html_filename):
    """
    Generates an interactive Plotly Express line plot from the spectrum analyzer data.
    The plot is saved as an HTML file.
    """
    print(f"\n--- Generating Interactive Plot: {output_html_filename} ---")

    # The DataFrame already contains 'Frequency (MHz)'
    # Create Interactive Plot with Plotly Express
    fig = px.line(df,
                  x="Frequency (MHz)",
                  y="Level (dBm)",
                  color="Band Name", # Differentiate lines by band name
                  title="Spectrum Analyzer Scans: Amplitude vs Frequency",
                  labels={"Frequency (MHz)": "Frequency (MHz)", "Level (dBm)": "Amplitude (dBm)"},
                  # Adjust hover_data to reflect the new DataFrame structure if needed,
                  # but for now, it matches 'Frequency (MHz)' and 'Level (dBm)' which are already plotted.
                  # Added Band Name to hover_data explicitly.
                  hover_data={"Frequency (MHz)": ':.2f', "Level (dBm)": ':.2f', "Band Name": True, "Peak Indicator": True}
                 )

    # Set X-axis (Frequency) to Logarithmic Scale
    fig.update_xaxes(type="log",
                     title="Frequency (MHz)",
                     showgrid=True, gridwidth=1,
                     tickformat = None) # Setting tickformat to None or "" often helps prevent scientific notation

    # Apply Dark Mode Theme
    fig.update_layout(template="plotly_dark")

    # Set Y-axis (Amplitude) Maximum to 0 dBm
    y_min = df['Level (dBm)'].min() # Get overall min across all plotted amplitude series
    # Ensure y-axis goes up to 0 dBm, with a sensible minimum (e.g., -100 dBm if data is higher)
    fig.update_yaxes(range=[y_min if y_min < -100 else -100, 0],
                     title="Amplitude (dBm)",
                     showgrid=True, gridwidth=1)

    # Customize Colors and Line Styles (Optional) - Plotly automatically assigns distinct colors
    # No specific customization applied by default, but can be added here if desired.

    fig.write_html(output_html_filename, auto_open=True)
    print(f"\n--- Plotly Express Interactive Plot Generated and saved to {output_html_filename} ---")


def main():
    """
    Main function to connect to the N9340B Spectrum Analyzer,
    run initial setup, perform band scans, and then read final configuration.
    """
    args = setup_arguments() # Parse command-line arguments

    # Determine instrument address based on user argument
    instrument_address = VISA_ADDRESSES.get(args.user)

    if not instrument_address:
        print(f"Error: No VISA address configured for user '{args.user}'.")
        print("Available devices:", pyvisa.ResourceManager().list_resources()) # Use ResourceManager directly here
        return

    # Initialize the instrument using the new function
    inst = initialize_instrument(instrument_address)

    if inst is None:
        print("Instrument initialization failed. Exiting.")
        return

    # File & directory setup for CSV and HTML plot
    scan_dir = os.path.join(os.getcwd(), "N9340 Scans")
    os.makedirs(scan_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H-%M-%S")
    csv_filename = os.path.join(scan_dir, f"{args.SCANname.strip()}_{timestamp}.csv")
    html_plot_filename = os.path.join(scan_dir, f"{args.SCANname.strip()}_{timestamp}.html")

    try:
        print(f"Opening CSV file for writing: {csv_filename}")
        with open(csv_filename, mode='w', newline='') as csvfile:
            csv_writer = csv.writer(csvfile)
            # Write header row to CSV with desired order and units
            csv_writer.writerow(["Frequency (MHz)", "Level (dBm)", "Band Name", "Peak Indicator"])

            # After successful initialization, proceed with scanning bands
            # scan_bands now takes the csv_writer and still returns the collected data for plotting
            all_scan_data = scan_bands(inst, csv_writer, args.maxHoldTime) # Pass maxHoldTime argument

        if not all_scan_data:
            print("No scan data collected. Skipping plotting.")
            return

        # Convert collected data to pandas DataFrame
        df = pd.DataFrame(all_scan_data)

        # Call the plotting function
        plot_spectrum_data(df, html_plot_filename)

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    finally:
        if inst and inst.session: # Check if inst object exists and has an active session
            inst.close()
            print("\nConnection to N9340B closed.")

if __name__ == '__main__':
    main()
