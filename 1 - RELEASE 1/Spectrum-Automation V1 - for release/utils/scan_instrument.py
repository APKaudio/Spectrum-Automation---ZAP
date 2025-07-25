# scan_instrument.py
#
# This module contains the core logic for controlling the spectrum analyzer
# to perform frequency sweeps across specified bands. It handles the low-level
# communication with the instrument via PyVISA, processes raw trace data,
# and saves it to CSV files. This is a critical component for data acquisition.
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
# scan_instrument.py

import pyvisa
import time
import numpy as np
import struct # Not used in current version, but kept if needed for future binary data handling
import re
import tkinter as tk # For messagebox, used in _debug_mode_enabled and within scan_bands
import datetime # Added import for datetime
import os # Added for path manipulation
import pandas as pd # Added for data manipulation (deduplication)

# Import instrument control functions
from utils.instrument_control import query_safe, write_safe, debug_print, initialize_instrument # Added initialize_instrument
# Import CSV utility functions
from utils.csv_utils import write_scan_data_to_csv # Changed to write_scan_data_to_csv for single file write

# Import constants from frequency_bands.py
try:
    from utils.frequency_bands import (
        MHZ_TO_HZ,
        VBW_RBW_RATIO # Import VBW_RBW_RATIO
    )
except ImportError:
    print("Error: frequency_bands.py not found or VBW_RBW_RATIO missing. Ensure it's in the utils directory.")
    MHZ_TO_HZ = 1_000_000
    VBW_RBW_RATIO = 1/3 # Define dummy value to prevent errors

def _process_raw_scan_data(raw_data, overall_start_freq_hz, overall_stop_freq_hz):
    """
    Processes raw scan data by de-duplicating entries, sorting by frequency,
    and interpolating missing frequency points to ensure a continuous spectrum.

    Inputs:
        raw_data (list): A list of (frequency_hz, level_dbm) tuples, potentially with duplicates.
        overall_start_freq_hz (float): The lowest frequency in the entire scan sweep.
        overall_stop_freq_hz (float): The highest frequency in the entire scan sweep.
    Process:
        1. Converts raw data to a Pandas DataFrame.
        2. Removes duplicate frequency entries, keeping the first occurrence.
        3. Sorts the DataFrame by frequency.
        4. Creates a complete range of frequencies from `overall_start_freq_hz` to `overall_stop_freq_hz`
           with a step size determined by the unique frequencies in the raw data.
        5. Merges the raw data with the complete frequency range, filling missing values using linear interpolation.
        6. Handles NaN values by filling them with the nearest valid data point.
        7. Converts frequencies to MHz for consistency in the output.
    Outputs:
        list: A list of (frequency_mhz, level_dbm) tuples representing the processed,
              de-duplicated, sorted, and interpolated scan data.
    """
    if not raw_data:
        debug_print("No raw data to process.")
        return []

    # Convert to DataFrame for easier manipulation
    df = pd.DataFrame(raw_data, columns=['Frequency_Hz', 'Level_dBm'])

    # Remove duplicates, keeping the first occurrence
    df.drop_duplicates(subset=['Frequency_Hz'], keep='first', inplace=True)

    # Sort by frequency
    df.sort_values(by='Frequency_Hz', inplace=True)

    # Create a complete frequency range for interpolation
    # Determine step size from the existing data or use a default if not enough points
    if len(df['Frequency_Hz']) > 1:
        # Calculate the average step size from the unique frequencies
        avg_step = np.mean(np.diff(df['Frequency_Hz'].unique()))
    else:
        # Fallback if there's not enough data to calculate a step
        avg_step = 1000 # Default step of 1 kHz, or adjust as appropriate for your instrument

    # Ensure the step is not zero or too small to prevent infinite loops or huge arrays
    if avg_step <= 0:
        avg_step = 1000 # Fallback to a safe default

    # Create a new, dense frequency range
    new_freq_range = np.arange(overall_start_freq_hz, overall_stop_freq_hz + avg_step, avg_step)
    
    # Create a DataFrame for the new frequency range
    df_full_range = pd.DataFrame({'Frequency_Hz': new_freq_range})

    # Merge the original data with the full frequency range, then interpolate
    df_merged = pd.merge(df_full_range, df, on='Frequency_Hz', how='left')
    
    # Interpolate missing values
    df_merged['Level_dBm'] = df_merged['Level_dBm'].interpolate(method='linear')

    # Fill any remaining NaN values (e.g., at the very start/end if interpolation can't reach)
    # Use ffill (forward fill) then bfill (backward fill) to fill from nearest valid data
    df_merged['Level_dBm'].ffill(inplace=True)
    df_merged['Level_dBm'].bfill(inplace=True)

    # Drop any rows where Level_dBm is still NaN (shouldn't happen with ffill/bfill but as a safeguard)
    df_merged.dropna(subset=['Level_dBm'], inplace=True)

    # Convert frequencies to MHz for output
    df_merged['Frequency_MHz'] = df_merged['Frequency_Hz'] / MHZ_TO_HZ

    return df_merged[['Frequency_MHz', 'Level_dBm']].values.tolist()


def scan_bands(app_instance, inst, selected_bands, output_folder, scan_name, freq_offset, scan_rbw_segmentation, vbw_config_val, max_hold_time):
    """
    Performs a spectrum scan across a list of selected frequency bands.
    This function controls the instrument to sweep each band, collect trace data,
    apply frequency offset, and save the raw data to a CSV file.

    Inputs:
        app_instance (App): The main application instance, used for console updates and stopping scan.
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        selected_bands (list): A list of dictionaries, each defining a frequency band with
                               "Band Name", "Start MHz", and "Stop MHz" keys.
        output_folder (str): The directory where CSV scan data will be saved.
        scan_name (str): The base name for the scan files.
        freq_offset (float): The frequency offset in Hz to apply to all scan frequencies.
        scan_rbw_segmentation (float): The Resolution Bandwidth (RBW) to use for each scan segment.
        vbw_config_val (int): The Video Bandwidth (VBW) to configure on the instrument.
        max_hold_time (float): Duration in seconds for which MAX Hold should be active (0 if disabled).
    Process:
        1. Checks if the instrument is connected.
        2. Calculates overall start and stop frequencies for the entire sweep.
        3. Initializes the instrument with `rbw_config_val`, `vbw_config_val`, and other settings.
        4. Loops through each selected band:
           - Sets instrument's start and stop frequencies for the current segment.
           - Configures Max Hold if enabled.
           - Initiates a single sweep (`:INIT:IMMed; *WAI`).
           - Fetches trace data (`:TRACe:DATA? TRACe1`).
           - Parses the trace data (comma-separated frequencies and levels).
           - Applies `freq_offset` to each frequency point.
           - Stores raw data.
           - Handles various errors during communication and data processing.
        5. After all bands are scanned (or if interrupted), processes the raw data
           using `_process_raw_scan_data`.
        6. Saves the processed data to a CSV file.
        7. Returns the processed data and the CSV filename.
    Outputs:
        tuple: (list, str) A tuple containing:
               - A list of (frequency_mhz, level_dbm) tuples for the full sweep.
               - The filename of the saved CSV.
               Returns (None, None) if no data is collected or an error occurs.
    """
    if not inst:
        print("🚫 Instrument not connected. Cannot perform scan.")
        return None, None

    raw_scan_data_for_current_sweep = []
    current_scan_csv_path = None
    overall_start_freq_hz = float('inf')
    overall_stop_freq_hz = float('-inf')

    if not selected_bands:
        print("🚫 No frequency bands selected for scan.")
        return None, None

    # Determine overall start and stop frequencies across all selected bands
    for band in selected_bands:
        start_hz = band['Start MHz'] * MHZ_TO_HZ
        stop_hz = band['Stop MHz'] * MHZ_TO_HZ
        overall_start_freq_hz = min(overall_start_freq_hz, start_hz)
        overall_stop_freq_hz = max(overall_stop_freq_hz, stop_hz)

    # Initialize instrument settings once for the entire scan operation
    # This assumes initialize_instrument can handle being called multiple times or
    # that its settings persist across segments.
    # It's better to ensure the instrument is initialized with the current RBW/VBW
    # before starting the sweep.
    # The app_instance.instrument_model is passed from main_app.py
    if not app_instance.instrument_model:
        print("🚫 Instrument model not identified. Cannot initialize instrument.")
        return None, None

    # Re-initialize instrument with current settings for the sweep
    # This will set RBW, VBW, Ref Level, Preamp, etc.
    # Corrected call: initialize_instrument is a standalone function
    if not initialize_instrument(
        app_instance.inst,
        float(app_instance.desired_ref_level_var.get()),
        app_instance.high_sensitivity_var.get(),
        app_instance.desired_preamp_var.get(),
        int(scan_rbw_segmentation), # Ensure RBW is int for instrument command
        vbw_config_val,
        app_instance.instrument_model
    ):
        print("🚫 Failed to initialize instrument for scan. Aborting.")
        return None, None

    # Generate a unique filename for the current scan
    scan_name_safe = re.sub(r'[^\w\-_\. ]', '_', scan_name) # Sanitize scan name
    
    # Debug print for type and value of scan_rbw_segmentation
    debug_print(f"Type of scan_rbw_segmentation: {type(scan_rbw_segmentation)}, Value: {scan_rbw_segmentation}")
    
    # Ensure scan_rbw_segmentation is float before division for display purposes
    rbw_str = f"RBW{int(float(scan_rbw_segmentation)/1000):04d}K" if float(scan_rbw_segmentation) >= 1000 else f"RBW{int(float(scan_rbw_segmentation))}Hz"
    hold_str = f"HOLD{int(max_hold_time)}s" if max_hold_time > 0 else "NOHOLD"
    offset_str = f"Offset{int(freq_offset)}Hz" if freq_offset != 0 else "NoOffset"
    # Corrected: Use datetime.datetime.now()
    timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    current_scan_csv_filename = f"{scan_name_safe}_{rbw_str}_{hold_str}_{offset_str}_{timestamp_str}.csv"
    current_scan_csv_path = os.path.join(output_folder, current_scan_csv_filename)

    # Removed the direct header writing here. It's now handled by write_scan_data_to_csv.
    csv_header = ["Frequency (MHz)", "Level (dBm)"] # Define header for passing to write_scan_data_to_csv

    last_successful_band_index = -1

    current_segment_start_freq_hz = overall_start_freq_hz # Start from the overall beginning

    for i, band in enumerate(selected_bands):
        if not app_instance.scanning: # Check for stop signal before processing each band
            print(f"Scan stopped by user during band {band['Band Name']}.")
            break

        band_start_freq_mhz = band['Start MHz']
        band_stop_freq_mhz = band['Stop MHz']
        band_name = band['Band Name']

        segment_start_freq_hz = band_start_freq_mhz * MHZ_TO_HZ + freq_offset
        segment_stop_freq_hz = band_stop_freq_mhz * MHZ_TO_HZ + freq_offset

        # Ensure segment frequencies are within the instrument's capabilities if known
        # (This is a generic check, actual limits depend on the instrument)
        # For example, if instrument has a hard lower limit of 9 kHz, adjust.
        # This part might need more specific logic based on the instrument model.

        print(f"\nScanning Band: {band_name} ({band_start_freq_mhz:.3f} MHz to {band_stop_freq_mhz:.3f} MHz)")
        print(f"  Applied Offset: {freq_offset} Hz")
        print(f"  Effective Scan Range: {segment_start_freq_hz / MHZ_TO_HZ:.3f} MHz to {segment_stop_freq_hz / MHZ_TO_HZ:.3f} MHz")

        try:
            # Set Start and Stop Frequencies for the current segment
            if not write_safe(inst, f":SENSE:FREQ:STAR {segment_start_freq_hz}"): raise Exception("Failed to set start frequency.")
            if not write_safe(inst, f":SENSE:FREQ:STOP {segment_stop_freq_hz}"): raise Exception("Failed to set stop frequency.")
            
            # Configure Max Hold if enabled
            if max_hold_time > 0:
                if not write_safe(inst, ":TRACe:TYPE MAXH"): raise Exception("Failed to set trace type to Max Hold.")
                # For N9340B, there isn't a direct "max hold time" command. Max hold is usually continuous until reset.
                # If a specific hold time is needed, it might involve polling or a custom sweep.
                # For now, we'll just enable Max Hold. The 'max_hold_time' might be used for a software-based hold.
                print(f"  Max Hold Enabled (Instrument side).")
            else:
                if not write_safe(inst, ":TRACe:TYPE NORM"): raise Exception("Failed to set trace type to Normal.")
                print("  Max Hold Disabled (Instrument side).")

            # Initiate a single sweep and wait for it to complete
            if not write_safe(inst, ":INIT:IMMed"): raise Exception("Failed to initiate immediate sweep.")
            if not write_safe(inst, "*WAI"): raise Exception("Failed to wait for sweep completion.")
            
            # Fetch trace data (Trace1)
            trace_data_str = query_safe(inst, ":TRACe:DATA? TRACe1")
            
            if not trace_data_str:
                print(f"🚫 No trace data received for band {band_name}.")
                continue

            # Parse the trace data string
            # Example format: "FREQ1,LEVEL1,FREQ2,LEVEL2,..."
            data_points = [float(x) for x in trace_data_str.split(',')]
            
            if len(data_points) % 2 != 0:
                print(f"🚫 Incomplete trace data received for band {band_name}. Skipping this band.")
                continue

            # Pair frequencies and levels, apply offset, and convert frequency to MHz for storage
            # The instrument returns frequencies in Hz, so no need to convert from MHz to Hz here.
            # We convert to MHz for storage in CSV.
            current_band_data = []
            for j in range(0, len(data_points), 2):
                freq_hz = data_points[j]
                level_dbm = data_points[j+1]
                
                # Apply the current cycle's frequency offset
                adjusted_freq_hz = freq_hz # Offset is already handled by setting start/stop freqs
                
                current_band_data.append((adjusted_freq_hz, level_dbm))
            
            raw_scan_data_for_current_sweep.extend(current_band_data)
            last_successful_band_index = i
            print(f"✅ Data acquired for band {band_name}. Points: {len(current_band_data)}")

        except pyvisa.errors.VisaIOError as e:
            print(f"🚫 VISA IO Error during scan for band {band_name}: {e}")
            print(f"🐛 Raw data string potentially causing error: {e}") # This line seems misplaced
            raise # Re-raise the exception to be caught by the main loop for recovery
        except ValueError as e:
            print(f"🚫 Error processing ASCII trace data (ValueError - cannot convert/unpack) for band {band_name}: {e}")
            print(f"🐞 Raw data string for parsing: '{trace_data_str}'")
        except Exception as e:
            print(f"🚨 An unexpected error occurred during trace processing for band {band_name}: {e}")

        # Move to the start of the next segment (no need if bands are distinct)
        # current_segment_start_freq_hz = segment_stop_freq_hz # This is not needed for distinct bands

    # --- After all bands are scanned, or if interrupted, process the collected raw data for the full sweep ---\
    if not raw_scan_data_for_current_sweep:
        print("🚫 No raw data collected across all bands for the current sweep.")
        return None, None

    print("\n--- 🎉 Band Scan Complete! Processing collected data... ---")
    final_sweep_data_for_plotting = _process_raw_scan_data(
        raw_scan_data_for_current_sweep,
        overall_start_freq_hz, # Use the calculated overall start
        overall_stop_freq_hz # Use the calculated overall stop
    )

    if not final_sweep_data_for_plotting:
        print(f"🚫 No data collected for full scan after de-duplication attempt for band: {selected_bands[last_successful_band_index]['Band Name'] if last_successful_band_index != -1 else 'N/A'}.")
    else:
        print(f"✅ De-duplicated and filtered {len(final_sweep_data_for_current_sweep)} points for plotting for full scan.")

    # Save the final processed data to CSV
    try:
        # write_scan_data_to_csv expects (freq_mhz, level_dbm) and now handles the header
        write_scan_data_to_csv(current_scan_csv_path, csv_header, final_sweep_data_for_plotting, append_mode=False)
        print(f"✅ Full sweep data saved to: {current_scan_csv_path}")
    except Exception as e:
        print(f"❌ Error saving full sweep data to CSV: {e}")
        return None, None

    return final_sweep_data_for_plotting, current_scan_csv_path
