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

    # Ensure Power_dBm is numeric (it might be read as string from CSV if not handled)
    # This line is crucial for preventing the TypeError
    df['Power_dBm'] = pd.to_numeric(df['Power_dBm'], errors='coerce')
    df.dropna(subset=['Power_dBm'], inplace=True) # Remove rows where conversion failed


    # Remove duplicates based on Frequency_MHz, keeping the last (or first, depending on preference)
    df.drop_duplicates(subset='Frequency_MHz', keep='last', inplace=True)

    # Sort by Frequency_MHz
    df.sort_values(by='Frequency_MHz', inplace=True)

    # Filter data to be within the overall scan range (in MHz for comparison)
    overall_start_freq_mhz = overall_start_freq_hz / MHZ_TO_HZ
    overall_stop_freq_mhz = overall_stop_freq_hz / MHZ_TO_HZ
    # Add a small epsilon for float comparison at stop to include the exact stop frequency
    df = df[(df['Frequency_MHz'] >= overall_start_freq_mhz) & (df['Frequency_MHz'] <= overall_stop_freq_mhz + 1e-9)]

    # Return as a list of (freq_mhz, level_dbm) tuples
    return list(zip(df['Frequency_MHz'], df['Power_dBm']))


def scan_bands(app_instance_ref, inst, stop_event, pause_event, instrument_model,
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
    debug_print(f"scan_bands parameters: "
                f"instrument_model={instrument_model}, "
                f"rbw_val={rbw_val}, "
                f"cycle_wait_time_val={cycle_wait_time_val}, "
                f"maxhold_time_val={maxhold_time_val}, "
                f"reference_level_val={reference_level_val}, "
                f"freq_shift_val={freq_shift_val}, "
                f"maxhold_enabled_val={maxhold_enabled_val}, "
                f"high_sensitivity_val={high_sensitivity_val}, "
                f"preamp_on_val={preamp_on_val}, "
                f"scan_rbw_segmentation_val={scan_rbw_segmentation_val}, "
                f"scan_name='{scan_name}', "
                f"output_folder='{output_folder}', "
                f"selected_bands={selected_bands}, "
                f"current_scan_cycle_count={current_scan_cycle_count}",
                file=file, function=function)

    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Define common instrument commands based on model
    is_n9340b = (instrument_model == "N9340B")

    # Store all raw data for the current full sweep across all bands
    raw_scan_data_for_current_sweep = []
    
    overall_start_freq_hz = float('inf')
    overall_stop_freq_hz = float('-inf')

    last_successful_band_index = -1

    # Determine the CSV filename for this scan session (continuous raw data)
    timestamp_hm = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename_current_cycle = os.path.join(output_folder, f"{scan_name}_Cycle{current_scan_cycle_count}_{timestamp_hm}.csv")

    app_instance_ref.after(0, app_instance_ref._update_console_line, "\n--- 📡 Starting Band Scan ---\n")
    app_instance_ref.after(0, app_instance_ref._update_console_line, "💾 Assuming ASCII data format for trace data.\n")

    # Calculate overall start and stop frequencies for the entire selected_bands list
    if selected_bands:
        # Sort bands by start frequency to ensure overall range is correct
        sorted_bands = sorted(selected_bands, key=lambda x: x["Start MHz"])
        overall_start_freq_hz = (sorted_bands[0]["Start MHz"] * MHZ_TO_HZ) + freq_shift_val
        overall_stop_freq_hz = (sorted_bands[-1]["Stop MHz"] * MHZ_TO_HZ) + freq_shift_val
    else:
        app_instance_ref.after(0, app_instance_ref._update_console_line, "🚫 No bands selected for scanning.\n")
        return -1, None

    # --- Apply global instrument settings once before starting band-specific sweeps ---
    # Set RBW and VBW
    # Convert rbw_val to int to ensure no decimal point is sent to the instrument
    if not write_safe(inst, f":SENSE:BAND:RES {int(rbw_val)}"): return last_successful_band_index, None
    vbw_calculated = rbw_val * VBW_RBW_RATIO
    # Convert vbw_calculated to int to ensure no decimal point is sent to the instrument
    if not write_safe(inst, f":SENSE:BAND:VID {int(vbw_calculated)}"): return last_successful_band_index, None
    app_instance_ref.after(0, app_instance_ref._update_console_line, 
                           f"📏 Global RBW set to {rbw_val/1000:.0f} kHz, VBW to {vbw_calculated:.0f} Hz.\n")

    # Set Reference Level
    if not write_safe(inst, f":DISPlay:WINdow:TRACe:Y:RLEVel {reference_level_val}DBM"): return last_successful_band_index, None
    app_instance_ref.after(0, app_instance_ref._update_console_line, 
                           f"⬆️ Reference Level set to {reference_level_val} dBm.\n")
    
    # Set Preamplifier
    preamp_cmd = ":INPut:ATTenuation:PREamp ON" if preamp_on_val else ":INPut:ATTenuation:PREamp OFF"
    if not write_safe(inst, preamp_cmd): return last_successful_band_index, None
    app_instance_ref.after(0, app_instance_ref._update_console_line, 
                           f"⚡ Preamplifier set to {'ON' if preamp_on_val else 'OFF'}.\n")

    # Set High Sensitivity (N9340B specific)
    if is_n9340b:
        hs_cmd = ":SENSe:POWer:RF:HSENs ON" if high_sensitivity_val else ":SENSe:POWer:RF:HSENs OFF"
        if not write_safe(inst, hs_cmd): return last_successful_band_index, None
        app_instance_ref.after(0, app_instance_ref._update_console_line, 
                               f"🔬 High Sensitivity set to {'ON' if high_sensitivity_val else 'OFF'}.\n")

    # Set Frequency Shift
    if freq_shift_val != 0:
        if not write_safe(inst, f":FREQuency:SHIFt {freq_shift_val}"): return last_successful_band_index, None
        app_instance_ref.after(0, app_instance_ref._update_console_line, 
                               f"↔️ Frequency Shift set to {freq_shift_val} Hz.\n")
    else:
        if not write_safe(inst, ":FREQuency:SHIFt 0"): return last_successful_band_index, None
        app_instance_ref.after(0, app_instance_ref._update_console_line, 
                               "↔️ Frequency Shift set to 0 Hz.\n")

    # Set Sweep Time to Auto
    if not write_safe(inst, ":SWEep:TIME:AUTO ON"): return last_successful_band_index, None
    app_instance_ref.after(0, app_instance_ref._update_console_line, 
                           "⏱️ Sweep Time set to Auto.\n")
    # Set initial trace modes (blank all, then set Trace2 for Max Hold/Normal)
    if not write_safe(inst, ":TRAC1:MODE BLANk;:TRAC2:MODE BLANk;:TRAC3:MODE BLANk"): return last_successful_band_index, None
    app_instance_ref.after(0, app_instance_ref._update_console_line, 
                           "✨ Traces 1, 2, 3 set to Blank mode.\n")
    # --- End global instrument settings ---


    for band_index, band in enumerate(selected_bands):
        if stop_event.is_set():
            debug_print("Stop event detected, exiting band scan loop.", file=file, function=function)
            break # Exit if stop is requested

        band_name = band["Band Name"]
        # Apply the frequency offset here
        band_start_freq_hz = (band["Start MHz"] * MHZ_TO_HZ) + freq_shift_val
        band_stop_freq_hz = (band["Stop MHz"] * MHZ_TO_HZ) + freq_shift_val

        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        app_instance_ref.after(0, app_instance_ref._update_console_line, 
                               f"\n📈 [{current_time}] Processing Band: {band_name} (Shifted Range: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {band_stop_freq_hz/MHZ_TO_HZ:.3f} MHz)\n")

        # Determine actual_sweep_points based on instrument model
        if instrument_model == "N9340B":
            expected_sweep_points = 461
        elif instrument_model == "N9342CN":
            expected_sweep_points = 500
        else:
            expected_sweep_points = 500 # Default for unknown models
        app_instance_ref.after(0, app_instance_ref._update_console_line, 
                               f"📊 Using {expected_sweep_points} sweep points per trace for {band_name} ({instrument_model if instrument_model else 'Unknown'} detected).\n")


        # --- Reverted Segment Calculation Logic to previous version ---
        full_band_span_hz = band_stop_freq_hz - band_start_freq_hz
        
        if full_band_span_hz <= 0:
            total_segments_in_band = 1
            # For a single point or zero span, the segment span is just the full band span
            optimal_segment_span_hz = full_band_span_hz
        else:
            # Recalculate optimal_segment_span_hz to perfectly divide the band into equal segments
            # This ensures all segments have the same span, even if it slightly deviates from scan_rbw_segmentation
            total_segments_in_band = int(np.ceil(full_band_span_hz / (scan_rbw_segmentation_val * (expected_sweep_points - 1))))
            if total_segments_in_band == 0: # Ensure at least one segment for non-zero spans
                total_segments_in_band = 1
            
            # Now calculate the actual segment span to ensure equal division
            optimal_segment_span_hz = full_band_span_hz / total_segments_in_band
            # Ensure it's at least the minimum possible span for expected_sweep_points > 1
            if expected_sweep_points > 1 and optimal_segment_span_hz < (scan_rbw_segmentation_val * (expected_sweep_points - 1)):
                optimal_segment_span_hz = scan_rbw_segmentation_val * (expected_sweep_points - 1)
        # --- End Reverted Segment Calculation Logic ---


        # The effective scan stop frequency for the segment loop
        # This should be the actual end of the current band, not an extended range.
        effective_scan_stop_freq_hz = band_start_freq_hz + (total_segments_in_band * optimal_segment_span_hz)
        app_instance_ref.after(0, app_instance_ref._update_console_line, 
                               f"🎯 Optimal segment span for {band_name}: {np.ceil(optimal_segment_span_hz / MHZ_TO_HZ * 10) / 10:.1f} MHz.\n") # Rounded up to 1 decimal place
        app_instance_ref.after(0, app_instance_ref._update_console_line, 
                               f"📏 Effective scanned range for equal segments: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {effective_scan_stop_freq_hz/MHZ_TO_HZ:.3f} MHz.\n")


        current_segment_start_freq_hz = band_start_freq_hz
        segment_counter = 0

        # Loop until the current segment start frequency reaches or exceeds the band's actual stop frequency
        while current_segment_start_freq_hz < band_stop_freq_hz and not stop_event.is_set():
            # Handle pause/resume
            while pause_event.is_set():
                app_instance_ref.after(0, app_instance_ref._update_console_line, "Scan Paused. Click Resume to continue.\n")
                time.sleep(0.1)
                if stop_event.is_set(): # Check stop event even when paused
                    debug_print(f"Stop event detected during pause in segment {segment_counter + 1}.", file=file, function=function)
                    return last_successful_band_index, csv_filename_current_cycle

            if stop_event.is_set(): # Check stop event after pause loop
                debug_print(f"Stop event detected, exiting band scan loop.", file=file, function=function)
                break

            segment_counter += 1
            segment_stop_freq_hz = current_segment_start_freq_hz + optimal_segment_span_hz

            # Ensure segment stop frequency does not exceed the band's actual stop frequency
            segment_stop_freq_hz = min(segment_stop_freq_hz, band_stop_freq_hz)

            # --- Minimal Instrument Settings per segment ---
            # Only set frequency range and trace mode per segment
            # Convert frequencies to int and round up using np.ceil() before sending to the instrument
            if not write_safe(inst, f":SENS:FREQ:STAR {int(np.ceil(current_segment_start_freq_hz))};:SENS:FREQ:STOP {int(np.ceil(segment_stop_freq_hz))};:TRAC2:MODE Blank"):
                return last_successful_band_index, None
            
            # Set trace mode (only Trace2 needs to be set to MAXH or NORM)
            if maxhold_enabled_val:
                if not write_safe(inst, ":TRAC2:MODE MAXHold;"): return last_successful_band_index, None
            else:
                if not write_safe(inst, ":TRAC2:MODE NORM;"): return last_successful_band_index, None # Use Normal if max hold is off

            time.sleep(0.1)

           
            # Add settling time for max hold values to show up, if max hold is enabled
            if maxhold_enabled_val and maxhold_time_val > 0:
                for _ in range(int(maxhold_time_val * 10)): # Check every 0.1 seconds
                    while pause_event.is_set():
                        app_instance_ref.after(0, app_instance_ref._update_console_line, "Scan Paused. Click Resume to continue.\n")
                        time.sleep(0.1) # Sleep briefly while paused
                        if stop_event.is_set():
                            debug_print(f"Stop event detected during pause in max hold for segment {segment_counter + 1}.", file=file, function=function)
                            return last_successful_band_index, csv_filename_current_cycle
                    if stop_event.is_set():
                        debug_print(f"Stop event detected during max hold for segment {segment_counter + 1}.", file=file, function=function)
                        return last_successful_band_index, csv_filename_current_cycle

                    # Update display for countdown (only update every second for cleaner output)
                    if _ % 10 == 0: # Every 10 iterations (1 second)
                        sec_remaining = int(maxhold_time_val - (_ / 10))
                        display_text = f"⏳ {sec_remaining}"
                        app_instance_ref.after(0, app_instance_ref._update_console_line, display_text) # Added \n
                    time.sleep(0.1)
                app_instance_ref.after(0, app_instance_ref._update_console_line, "✅") # Added \n

            if stop_event.is_set():
                debug_print(f"Stop event detected after max hold for segment {segment_counter + 1}.", file=file, function=function)
                break

            # Calculate progress for the emoji bar - Using more compatible ASCII characters
            progress_percentage = (segment_counter / total_segments_in_band)
            bar_length = 20 # Total number of characters in the bar
            filled_length = int(round(bar_length * progress_percentage))
            # Using '█' (U+2588 Full Block) and '-' (Hyphen) for better compatibility
            progressbar = '█' * filled_length + '-' * (bar_length - filled_length)

            # Calculate display values, rounding up to match instrument commands
            # Ensure optimal_segment_span_hz is divided by MHZ_TO_HZ before rounding for display
            display_span_mhz = np.ceil(optimal_segment_span_hz / MHZ_TO_HZ * 10) / 10
            display_start_freq_mhz = np.ceil(current_segment_start_freq_hz) / MHZ_TO_HZ
            display_stop_freq_mhz = np.ceil(segment_stop_freq_hz) / MHZ_TO_HZ

            # Combined print statement as per user request, now using _update_console_line without overwrite
            # Added '\r' to ensure the next message overwrites this line correctly
            progress_message = (f"{progressbar}🔍 Span:📊{display_span_mhz:.1f} MHz--" # Formatted to 1 decimal place
                                f"📈{display_start_freq_mhz:.3f} MHz to " # Formatted to 3 decimal places
                                f"📉{display_stop_freq_mhz:.3f} MHz   ✅{segment_counter} of {total_segments_in_band}\n") # Formatted to 3 decimal places
            app_instance_ref.after(0, app_instance_ref._update_console_line, progress_message + '\n')


            # Read and process trace data
            trace_data = []
            segment_raw_data = [] # To store raw data for the current segment before filtering
            try:
                # Conditional trace data query based on instrument model
                if instrument_model == "N9340B":
                    trace_data_str = query_safe(inst, ":TRAC2:DATA?")
                elif instrument_model == "N9342CN":
                    trace_data_str = query_safe(inst, ":TRACe:DATA? TRACe2")
                else: # Fallback for unknown models
                    trace_data_str = query_safe(inst, ":TRACe:DATA? TRACe2")


                if trace_data_str is None or "[Not Supported or Timeout]" in trace_data_str or not trace_data_str.strip():
                    app_instance_ref.after(0, app_instance_ref._update_console_line, "🚫 No valid trace data string received for this segment.\n")
                    current_segment_start_freq_hz = segment_stop_freq_hz
                    continue # Move to the next segment if no data

                # Split the string by comma and convert each part to float
                trace_data = [float(val) for val in trace_data_str.split(',')]

                num_trace_points_actual = len(trace_data)

                if num_trace_points_actual > 1:
                    freq_step_per_point_actual = optimal_segment_span_hz / (num_trace_points_actual - 1)
                elif num_trace_points_actual == 1:
                    freq_step_per_point_actual = 0 # Single point, no step
                else:
                    freq_step_per_point_actual = 0 # No points
                    app_instance_ref.after(0, app_instance_ref._update_console_line, "🚫 No trace data received for this segment after parsing.\n")
                    current_segment_start_freq_hz = segment_stop_freq_hz
                    continue # Move to the next segment if no data

                # Loop to append data to raw_scan_data_for_current_sweep (for later de-duplication for plotting)
                # and to segment_raw_data (for immediate CSV writing after filtering)
                for j, amp_value in enumerate(trace_data):
                    current_freq_for_point_hz = current_segment_start_freq_hz + (j * freq_step_per_point_actual)
                    # Convert amp_value to float here to ensure numeric type
                    segment_raw_data.append((current_freq_for_point_hz, float(amp_value)))
                    raw_scan_data_for_current_sweep.append((current_freq_for_point_hz, float(amp_value)))


                # *** Filter segment data before writing to CSV ***
                # Ensure data written to CSV is strictly within the original band's bounds
                filtered_segment_data_for_csv = []
                for freq_hz, amp_value in segment_raw_data:
                    # IMPORTANT CHANGE: Filter based on the original band's start and stop frequencies
                    # This ensures no "leftover" data outside the selected band is written to CSV.
                    if freq_hz >= band_start_freq_hz and freq_hz <= band_stop_freq_hz + 1e-9:
                        filtered_segment_data_for_csv.append((freq_hz, amp_value))

                # Write filtered segment data to CSV immediately after processing
                if filtered_segment_data_for_csv: # Only write if there's data after filtering
                    header = ["Frequency_MHz", "Power_dBm"] # Define header for this CSV
                    # Convert frequencies to MHz for the CSV
                    csv_data_to_write = [(f / MHZ_TO_HZ, amp) for f, amp in filtered_segment_data_for_csv]
                    # This will create/append to the CSV file defined at the start of scan_bands
                    write_scan_data_to_csv(csv_filename_current_cycle, None, csv_data_to_write, append_mode=True)
                    # Removed the print statement here as requested


                # Update last_successful_band_index after successfully processing a band
                last_successful_band_index = band_index

            except pyvisa.errors.VisaIOError as e:
                app_instance_ref.after(0, app_instance_ref._update_console_line, f"🛑 Error reading trace data (PyVISA IO Error): {e}\n")
                app_instance_ref.after(0, app_instance_ref._update_console_line, f"🐛 Raw data string potentially causing error: {e}\n")
                # Do not raise, allow the scan to continue to the next segment/band if possible
                break # Exit current band, try next if not stopping entirely
            except ValueError as e:
                app_instance_ref.after(0, app_instance_ref._update_console_line, f"🚫 Error processing ASCII trace data (ValueError - cannot convert/unpack): {e}\n")
                app_instance_ref.after(0, app_instance_ref._update_console_line, f"🐞 Raw data string for parsing: {e}\n")
                continue # Move to the next segment if parsing fails
            except Exception as e:
                app_instance_ref.after(0, app_instance_ref._update_console_line, f"🚨 An unexpected error occurred during trace processing: {e}\n")
                continue # Move to the next segment if other error occurs

            # Move to the start of the next segment
            current_segment_start_freq_hz = segment_stop_freq_hz

    # --- After all bands are scanned, or if interrupted, process the collected raw data for the full sweep ---
    app_instance_ref.after(0, app_instance_ref._update_console_line, "\n--- 🎉 Band Scan Complete! Processing collected data... ---\n")
    final_sweep_data_for_plotting = _process_raw_scan_data(
        raw_scan_data_for_current_sweep,
        overall_start_freq_hz, # Use the calculated overall start
        overall_stop_freq_hz # Use the calculated overall stop
    )

    if not final_sweep_data_for_plotting:
        app_instance_ref.after(0, app_instance_ref._update_console_line, 
                               f"🚫 No data collected for full scan after de-duplication attempt for band: {selected_bands[last_successful_band_index]['Band Name'] if last_successful_band_index != -1 else 'N/A'}.\n")
        return last_successful_band_index, None # Return None for filename if no data after processing
    else:
        app_instance_ref.after(0, app_instance_ref._update_console_line, 
                               f"✅ De-duplicated and filtered {len(final_sweep_data_for_plotting)} points for plotting for full scan.\n")
        return last_successful_band_index, csv_filename_current_cycle
