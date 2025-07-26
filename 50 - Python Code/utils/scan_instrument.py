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
from tkinter import messagebox # Direct import for messagebox
import datetime # Added import for datetime
import os # Added for path manipulation
import pandas as pd # Added for data manipulation (deduplication)
import csv # Added import for csv module
import inspect # Import inspect module for debug_print

# Import instrument control functions
from utils.instrument_control import query_safe, write_safe, debug_print, initialize_instrument # Added initialize_instrument
# Import CSV utility functions
from utils.csv_utils import write_scan_data_to_csv # Changed to write_scan_data_to_csv for single file write
from utils.frequency_bands import MHZ_TO_HZ, VBW_RBW_RATIO # Import constants

def _process_raw_scan_data(raw_data, overall_start_freq_hz, overall_stop_freq_hz, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Processes raw frequency and level data to remove duplicates and sort it.
    Converts frequencies to MHz and levels to dBm.
    Ensures data is within the overall scan range.
    """
    if not raw_data:
        debug_print("No raw data to process.", file=file, function=function)
        return []

    # Convert to a DataFrame for efficient processing
    # Assuming raw_data is a list of (frequency_hz, level_dbm) tuples/lists
    df = pd.DataFrame(raw_data, columns=['Frequency_Hz', 'Power_dBm'])

    # Convert Frequency to MHz
    df['Frequency_MHz'] = df['Frequency_Hz'] / MHZ_TO_HZ

    # Remove duplicates based on Frequency_MHz, keeping the last (or first, depending on preference)
    df.drop_duplicates(subset='Frequency_MHz', keep='last', inplace=True)

    # Sort by Frequency_MHz
    df.sort_values(by='Frequency_MHz', inplace=True)

    # Filter data to be within the overall scan range (in MHz for comparison)
    overall_start_freq_mhz = overall_start_freq_hz / MHZ_TO_HZ
    overall_stop_freq_mhz = overall_stop_freq_hz / MHZ_TO_HZ
    df = df[(df['Frequency_MHz'] >= overall_start_freq_mhz) & (df['Frequency_MHz'] <= overall_stop_freq_mhz)]

    # Return as a list of (freq_mhz, level_dbm) tuples
    return list(zip(df['Frequency_MHz'], df['Power_dBm']))


def scan_bands(app_instance_ref, inst, stop_event, pause_event, instrument_model, # FIX: Added app_instance_ref
               rbw_val, cycle_wait_time_val, maxhold_time_val,
               reference_level_val, freq_shift_val, maxhold_enabled_val,
               high_sensitivity_val, preamp_on_val, scan_rbw_segmentation_val,
               scan_name, output_folder, selected_bands, current_scan_cycle_count,
               file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Performs a frequency sweep across selected bands using the connected instrument.
    This function is designed to be run in a separate thread.

    Inputs:
        app_instance_ref (object): A reference to the main Tkinter App instance for GUI updates.
                                   Expected to have .scanning, .paused, .after, ._update_console_line, .instrument_model attributes.
        inst (pyvisa.resources.Resource): The connected PyVISA instrument object.
        stop_event (threading.Event): Event to signal the scan to stop.
        pause_event (threading.Event): Event to signal the scan to pause/resume.
        instrument_model (str): The detected model of the instrument.
        rbw_val (float): Resolution Bandwidth value in Hz.
        cycle_wait_time_val (float): Time to wait between scan cycles in seconds.
        maxhold_time_val (float): Max Hold time in seconds.
        reference_level_val (float): Reference Level in dBm.
        freq_shift_val (float): Frequency shift in Hz.
        maxhold_enabled_val (bool): True if Max Hold is enabled.
        high_sensitivity_val (bool): True if high sensitivity mode is enabled.
        preamp_on_val (bool): True if preamplifier is ON.
        scan_rbw_segmentation_val (float): RBW segmentation value in Hz.
        scan_name (str): Name of the current scan.
        output_folder (str): Directory to save scan data.
        selected_bands (list): List of dictionaries, each with "Band Name", "Start MHz", "Stop MHz".
        current_scan_cycle_count (int): The current scan cycle number.

    Returns:
        tuple: (last_successful_band_index, csv_filename_current_cycle)
               last_successful_band_index (int): Index of the last band successfully scanned.
                                                 -1 if no bands were scanned.
               csv_filename_current_cycle (str or None): Full path to the CSV file
                                                         for the current cycle, or None if no data.
    """
    debug_print(f"Starting scan_bands for cycle {current_scan_cycle_count}...", file=file, function=function)
    
    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Define common instrument commands based on model
    is_n9340b = (instrument_model == "N9340B")

    # Store all raw data for the current full sweep across all bands
    raw_scan_data_for_current_sweep = []
    
    overall_start_freq_hz = float('inf')
    overall_stop_freq_hz = float('-inf')

    last_successful_band_index = -1

    for band_index, band in enumerate(selected_bands):
        if stop_event.is_set():
            debug_print("Stop event detected, exiting band scan loop.", file=file, function=function)
            break # Exit if stop is requested

        while pause_event.is_set():
            debug_print("Pause event detected, pausing band scan.", file=file, function=function)
            time.sleep(0.1) # Small sleep to avoid busy-waiting

        band_name = band["Band Name"]
        start_freq_mhz = band["Start MHz"]
        stop_freq_mhz = band["Stop MHz"]

        start_freq_hz = start_freq_mhz * MHZ_TO_HZ
        stop_freq_hz = stop_freq_mhz * MHZ_TO_HZ

        # Update overall scan range
        overall_start_freq_hz = min(overall_start_freq_hz, start_freq_hz)
        overall_stop_freq_hz = max(overall_stop_freq_hz, stop_freq_hz)

        debug_print(f"Scanning band: {band_name} ({start_freq_mhz:.3f} MHz - {stop_freq_mhz:.3f} MHz)", file=file, function=function)

        try:
            # Set frequency range
            if not write_safe(inst, f":FREQuency:STARt {start_freq_hz}"): return last_successful_band_index, None
            if not write_safe(inst, f":FREQuency:STOP {stop_freq_hz}"): return last_successful_band_index, None
            
            # Set RBW and VBW
            if not write_safe(inst, f":BANDwidth {rbw_val}"): return last_successful_band_index, None
            vbw_val = rbw_val * VBW_RBW_RATIO
            if not write_safe(inst, f":BANDwidth:VIDeo {vbw_val}"): return last_successful_band_index, None

            # Set Reference Level
            if not write_safe(inst, f":DISPlay:WINdow:TRACe:Y:RLEVel {reference_level_val}DBM"): return last_successful_band_index, None

            # Set Max Hold (if enabled)
            if maxhold_enabled_val:
                if not write_safe(inst, f":DISPlay:WINDow:TRACe:TYPE MAXH"): return last_successful_band_index, None
                # Max hold time is usually set implicitly by the sweep time or continuous sweep
                # For specific max hold time, instrument might need a different command or a loop
            else:
                if not write_safe(inst, f":DISPlay:WINDow:TRACe:TYPE NORM"): return last_successful_band_index, None

            # Set Preamplifier (if enabled)
            if preamp_on_val:
                if not write_safe(inst, ":INPut:ATTenuation:PREamp ON"): return last_successful_band_index, None
            else:
                if not write_safe(inst, ":INPut:ATTenuation:PREamp OFF"): return last_successful_band_index, None

            # Set High Sensitivity (if applicable, N9340B specific)
            if is_n9340b:
                if high_sensitivity_val:
                    if not write_safe(inst, ":SENSe:POWer:RF:HSENs ON"): return last_successful_band_index, None
                else:
                    if not write_safe(inst, ":SENSe:POWer:RF:HSENs OFF"): return last_successful_band_index, None

            # Set Frequency Shift (if applicable)
            if freq_shift_val != 0:
                if not write_safe(inst, f":FREQuency:SHIFt {freq_shift_val}"): return last_successful_band_index, None
            else:
                if not write_safe(inst, ":FREQuency:SHIFt 0"): return last_successful_band_index, None # Ensure it's off if 0

            # Set Sweep Time to Auto
            if not write_safe(inst, ":SWEep:TIME:AUTO ON"): return last_successful_band_index, None
            
            # Initiate a single sweep and wait for completion
            if not write_safe(inst, ":INITiate:IMMediate"): return last_successful_band_index, None
            if not query_safe(inst, "*OPC?"): # Wait for operation complete
                debug_print("Operation not complete after initiate immediate.", file=file, function=function)
                return last_successful_band_index, None

            # Fetch trace data
            trace_data_str = query_safe(inst, ":TRACe:DATA? TRACE1")
            if trace_data_str is None:
                debug_print(f"Failed to query trace data for band {band_name}.", file=file, function=function)
                continue # Skip to next band if data query fails

            # Parse trace data
            # The data format is typically comma-separated values (CSV format from instrument)
            # Example: "FREQ1,LEVEL1,FREQ2,LEVEL2,..." or "LEVEL1,LEVEL2,..."
            # Assuming it's pairs of frequency and level
            data_points = [float(x) for x in trace_data_str.strip().split(',')]
            
            # Check if the data is in pairs (freq, level) or just levels
            # Most modern instruments return frequency and amplitude pairs
            if len(data_points) % 2 == 0:
                # Assuming (frequency_hz, level_dbm) pairs
                for i in range(0, len(data_points), 2):
                    freq_hz = data_points[i]
                    level_dbm = data_points[i+1]
                    raw_scan_data_for_current_sweep.append((freq_hz, level_dbm))
            else:
                # If only levels are returned, we need to calculate frequencies
                # This requires knowing start_freq_hz, stop_freq_hz, and number of points
                # For simplicity, we'll assume pairs for now, or handle based on instrument manual.
                # For now, if odd, log a warning and skip, or try to infer.
                debug_print(f"Warning: Unexpected trace data format for band {band_name}. Odd number of points: {len(data_points)}", file=file, function=function)
                # Fallback: If only levels, assume equally spaced frequencies
                # This is a simplification and might not be accurate for all instruments.
                num_points = len(data_points)
                if num_points > 0:
                    frequencies_hz = np.linspace(start_freq_hz, stop_freq_hz, num_points)
                    for i in range(num_points):
                        raw_scan_data_for_current_sweep.append((frequencies_hz[i], data_points[i]))


            last_successful_band_index = band_index
            debug_print(f"Collected data for band {band_name}.", file=file, function=function)

        except pyvisa.errors.VisaIOError as e:
            debug_print(f"VISA error during scan for band {band_name}: {e}", file=file, function=function)
            app_instance_ref.after(0, messagebox.showwarning, "VISA Error", f"VISA error during scan for band {band_name}: {e}")
            # Continue to next band or stop, depending on desired robustness
            continue # Try next band
        except Exception as e:
            debug_print(f"An unexpected error occurred during scan for band {band_name}: {e}", file=file, function=function)
            app_instance_ref.after(0, messagebox.showwarning, "Scan Error", f"An unexpected error occurred during scan for band {band_name}: {e}")
            continue # Try next band

    # After iterating through all bands, process the collected raw data
    debug_print("All bands scanned or loop interrupted. Processing collected data...", file=file, function=function)
    final_sweep_data_for_plotting = _process_raw_scan_data(
        raw_scan_data_for_current_sweep,
        overall_start_freq_hz, # Use the calculated overall start
        overall_stop_freq_hz # Use the calculated overall stop
    )

    csv_filename_current_cycle = None
    if final_sweep_data_for_plotting:
        debug_print(f"✅ De-duplicated and filtered {len(final_sweep_data_for_plotting)} points for plotting for full scan.", file=file, function=function)

        # Save the final processed data to CSV
        try:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_filename_current_cycle = os.path.join(output_folder, f"{scan_name}_Cycle{current_scan_cycle_count}_{timestamp}.csv")
            
            # write_scan_data_to_csv expects (freq_mhz, level_dbm)
            # _process_raw_scan_data now returns (freq_mhz, level_dbm)
            # Pass header as None to indicate no header should be written
            write_scan_data_to_csv(csv_filename_current_cycle, None, final_sweep_data_for_plotting, append_mode=False)
            debug_print(f"✅ Full sweep data saved to: {csv_filename_current_cycle}", file=file, function=function)
            
        except Exception as e:
            debug_print(f"❌ Error saving full sweep data to CSV: {e}", file=file, function=function)
            app_instance_ref.after(0, messagebox.showerror, "CSV Save Error", f"Could not save full sweep data to CSV: {e}")
            csv_filename_current_cycle = None # Indicate failure to save CSV
    else:
        debug_print(f"🚫 No data collected for full scan after de-duplication attempt for band: {selected_bands[last_successful_band_index]['Band Name'] if last_successful_band_index != -1 else 'N/A'}.", file=file, function=function)

    return last_successful_band_index, csv_filename_current_cycle
