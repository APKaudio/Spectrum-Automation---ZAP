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
import tkinter as tk # For messagebox, used in _debug_mode_enabled and within scan_bands # Keeping tk, removed messagebox
# from tkinter import messagebox # Direct import for messagebox # Removed
import datetime # Added import for datetime
import os # Added for path manipulation
import pandas as pd # Added for data manipulation (deduplication)
import csv # Added import for csv module
import inspect # Import inspect module for debug_print

# Import instrument control functions
from utils.instrument_control import query_safe, write_safe, debug_print, initialize_instrument
# Corrected import: Changed set_log_visa_commands to set_log_visa_commands_mode
from utils.instrument_control import set_debug_mode, set_log_visa_commands_mode
# Import CSV utility
from utils.csv_utils import write_scan_data_to_csv # Import write_scan_data_to_csv

# Import frequency band definitions
from ref.frequency_bands import SCAN_BAND_RANGES, MHZ_TO_HZ, VBW_RBW_RATIO # Import VBW_RBW_RATIO

def _process_raw_scan_data(raw_data, overall_start_freq_hz, overall_stop_freq_hz):
    """
    Processes raw scan data (frequency and amplitude pairs) to remove duplicates,
    sort by frequency, and ensure a contiguous range.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function} function.", file=current_file, function=current_function)
    debug_print(f"Processing raw scan data ({len(raw_data)} points)...", file=current_file, function=current_function)

    if not raw_data:
        debug_print("No raw data to process.", file=current_file, function=current_function)
        debug_print(f"Exiting {current_function} function. Result: Empty DataFrame", file=current_file, function=current_function)
        return pd.DataFrame()

    # Convert to DataFrame
    df = pd.DataFrame(raw_data, columns=['Frequency (Hz)', 'Amplitude (dBm)'])

    # Remove duplicates based on Frequency (keeping the last one which is often from the end of a segment)
    df.drop_duplicates(subset=['Frequency (Hz)'], keep='last', inplace=True)

    # Sort by frequency
    df.sort_values(by='Frequency (Hz)', inplace=True)

    # Filter to the overall scan range to remove any stray points outside
    df = df[(df['Frequency (Hz)'] >= overall_start_freq_hz) & (df['Frequency (Hz)'] <= overall_stop_freq_hz)]

    # Convert Frequency to MHz for consistency in output/plotting
    df['Frequency (MHz)'] = df['Frequency (Hz)'] / MHZ_TO_HZ
    
    # Reorder columns
    df = df[['Frequency (MHz)', 'Amplitude (dBm)']]

    debug_print(f"Processed data shape: {df.shape}", file=current_file, function=current_function)
    debug_print(f"Exiting {current_function} function. Result: DataFrame with {df.shape[0]} rows.", file=current_file, function=current_function)
    return df


def scan_bands(app_instance_ref, inst, selected_bands, rbw_hz, ref_level_dbm, freq_shift_hz, maxhold_enabled, high_sensitivity, preamp_on, rbw_step_size_hz, cycle_wait_time_seconds, scan_name, output_folder, stop_event, pause_event, log_visa_commands_enabled, general_debug_enabled, app_console_update_func):
    """
    Performs a full scan across specified frequency bands.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function} function.", file=current_file, function=current_function)
    debug_print("Starting scan_bands function.", file=current_file, function=current_function)

    # Configure debug mode for underlying instrument control functions
    # This must be done at the start of the scan
    set_debug_mode(general_debug_enabled)
    set_log_visa_commands_mode(log_visa_commands_enabled) # Corrected function call

    overall_start_freq_hz = min(band["Start MHz"] for band in selected_bands) * MHZ_TO_HZ
    overall_stop_freq_hz = max(band["Stop MHz"] for band in selected_bands) * MHZ_TO_HZ
    
    app_console_update_func(f"Scanning from {overall_start_freq_hz / MHZ_TO_HZ:.3f} MHz to {overall_stop_freq_hz / MHZ_TO_HZ:.3f} MHz...")
    debug_print(f"Overall scan range: {overall_start_freq_hz} Hz to {overall_stop_freq_hz} Hz", file=current_file, function=current_function)

    # Get max hold time from app instance
    max_hold_time = float(app_instance_ref.maxhold_time_seconds_var.get())

    # Initialize instrument for scan (basic setup, no specific scan parameters here)
    app_console_update_func("Initializing instrument for scan settings...")
    if not initialize_instrument(
        inst,
        model_match=app_instance_ref.instrument_model, # Pass the instrument model
        ref_level_dbm=ref_level_dbm,
        high_sensitivity_on=high_sensitivity,
        preamp_on=preamp_on,
        rbw_config_val=rbw_hz,
        vbw_config_val=rbw_hz * VBW_RBW_RATIO # Calculate VBW based on RBW and ratio
    ):
        app_console_update_func("❌ Error: Failed to initialize instrument for scan. Aborting.")
        debug_print("Instrument initialization failed in scan_bands.", file=current_file, function=current_function)
        debug_print(f"Exiting {current_function} function. Result: -1, None, None", file=current_file, function=current_function)
        return -1, None, None

    # Apply specific scan settings after basic initialization
    app_console_update_func("Applying scan parameters to instrument...")
    if not write_safe(inst, f":SENSe:BANDwidth:RESolution {rbw_hz}HZ"):
        app_console_update_func(f"❌ Error: Failed to set RBW to {rbw_hz}Hz.")
        debug_print(f"Failed to set RBW: {rbw_hz}Hz", file=current_file, function=current_function)
        debug_print(f"Exiting {current_function} function. Result: -1, None, None", file=current_file, function=current_function)
        return -1, None, None
    
    # Ensure VBW is set as a ratio of RBW, if supported by instrument.
    vbw_percent = int(VBW_RBW_RATIO * 100) # Convert 1/3 to 33%
    if not write_safe(inst, f":SENSe:BANDwidth:VIDeo:RATio {vbw_percent}PCT"): # Example command, verify for specific instrument
        app_console_update_func(f"❌ Error: Failed to set VBW ratio to {vbw_percent}PCT.")
        debug_print(f"Failed to set VBW ratio: {vbw_percent}PCT", file=current_file, function=current_function)
        # This might not be critical, so continue, but log error

    if not write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM"):
        app_console_update_func(f"❌ Error: Failed to set reference level to {ref_level_dbm}dBm.")
        debug_print(f"Failed to set reference level: {ref_level_dbm}dBm", file=current_file, function=current_function)
        debug_print(f"Exiting {current_function} function. Result: -1, None, None", file=current_file, function=current_function)
        return -1, None, None

    if not write_safe(inst, f":INPut:RFSense:FREQuency:SHIFt {freq_shift_hz}HZ"):
        app_console_update_func(f"❌ Error: Failed to set frequency shift to {freq_shift_hz}Hz.")
        debug_print(f"Failed to set frequency shift: {freq_shift_hz}Hz", file=current_file, function=current_function)
        # Not critical, continue but log

    # Maxhold setting is now handled per segment with the timer
    # maxhold_cmd = ":TRAC2:MODE MAXHold" if maxhold_enabled else ":TRAC2:MODE AVERage" # Assuming Trace 2 for MaxHold
    # if not write_safe(inst, maxhold_cmd):
    #     app_console_update_func(f"❌ Error: Failed to set Max Hold mode to {maxhold_enabled}.")
    #     debug_print(f"Failed to set Max Hold mode: {maxhold_enabled}", file=current_file, function=current_function)
    #     # Not critical, continue but log

    # High sensitivity / Preamp setting
    # This might involve multiple commands depending on the instrument model
    if high_sensitivity:
        # Assuming high sensitivity implies attenuation off and gain on
        if not write_safe(inst, ":INPut:ATTenuation:AUTO OFF"):
            app_console_update_func("❌ Error: Failed to set attenuation auto OFF for high sensitivity.")
            debug_print("Failed to set attenuation auto OFF.", file=current_file, function=current_function)
        if not write_safe(inst, ":INPut:GAIN:STATe ON"): # Preamp ON
            app_console_update_func("❌ Error: Failed to turn ON preamp for high sensitivity.")
            debug_print("Failed to turn ON preamp.", file=current_file, function=current_function)
    else:
        # Assuming high sensitivity off implies attenuation auto on and gain off
        if not write_safe(inst, ":INPut:ATTenuation:AUTO ON"):
            app_console_update_func("❌ Error: Failed to set attenuation auto ON for normal sensitivity.")
            debug_print("Failed to set attenuation auto ON.", file=current_file, function=current_function)
        if not write_safe(inst, ":INPut:GAIN:STATe OFF"): # Preamp OFF
            app_console_update_func("❌ Error: Failed to turn OFF preamp for normal sensitivity.")
            debug_print("Failed to turn OFF preamp.", file=current_file, function=current_function)

    if preamp_on: # Explicit preamp control, might be redundant with high_sensitivity
        if not write_safe(inst, ":INPut:GAIN:STATe ON"):
            app_console_update_func("❌ Error: Failed to turn ON preamp.")
            debug_print("Failed to turn ON preamp.", file=current_file, function=current_function)
    else:
        if not write_safe(inst, ":INPut:GAIN:STATe OFF"):
            app_console_update_func("❌ Error: Failed to turn OFF preamp.")
            debug_print("Failed to turn OFF preamp.", file=current_file, function=current_function)

    # Determine the CSV filename for this scan session (continuous raw data)
    timestamp_hm = datetime.datetime.now().strftime("%Y%m%d_%H%M%S") # YYYYMMDD_HHMMSS
    # Define the CSV filename for the *current scan cycle's raw data*
    csv_filename_current_cycle = os.path.join(output_folder, f"{scan_name}_RBW{int(rbw_hz/1000)}K_HOLD{int(max_hold_time)}_Offset{int(freq_shift_hz)}_{timestamp_hm}.csv")


    raw_scan_data_for_current_sweep = [] # List to collect (freq, amplitude) tuples for the entire sweep
    last_successful_band_index = -1 # Keep track of the last band that successfully scanned
    
    markers_data_from_scan = [] # To collect markers if extracted during the scan

    app_console_update_func("\n--- 📡 Starting Band Scan ---")
    app_console_update_func("💾 Assuming ASCII data format for trace data.")


    for i, band in enumerate(selected_bands):
        if stop_event.is_set():
            app_console_update_func("Scan stopped by user during band iteration.")
            debug_print("Scan stop event set during band iteration.", file=current_file, function=current_function)
            debug_print(f"Exiting {current_function} function. Result: -1, None, None", file=current_file, function=current_function)
            break # Exit loop if stop is requested

        band_name = band["Band Name"]
        
        # Apply the frequency offset here to the band's start and stop frequencies
        band_start_freq_hz = (band["Start MHz"] * MHZ_TO_HZ) + freq_shift_hz
        band_stop_freq_hz = (band["Stop MHz"] * MHZ_TO_HZ) + freq_shift_hz

        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M") # Use datetime.datetime
        app_console_update_func(f"\n📈 [{current_time}] Processing Band: {band_name} (Shifted Range: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {band_stop_freq_hz/MHZ_TO_HZ:.3f} MHz)")
        debug_print(f"Processing band: {band_name} ({band_start_freq_hz}-{band_stop_freq_hz} Hz)", file=current_file, function=current_function)

        # Determine actual_sweep_points based on instrument model
        # This logic is from the old scanner for compatibility
        if app_instance_ref.instrument_model == "N9340B":
            expected_sweep_points = 461
            app_console_update_func(f"📊 Using {expected_sweep_points} sweep points per trace for {band_name} (N9340B detected).")
        elif app_instance_ref.instrument_model == "N9342CN": # Explicitly check for N9342CN
            expected_sweep_points = 500
            app_console_update_func(f"📊 Using {expected_sweep_points} sweep points per trace for {band_name} (N9342CN detected).")
        else: # Default for unknown models, or other N9342 variants if needed
            expected_sweep_points = 500
            app_console_update_func(f"📊 Using {expected_sweep_points} sweep points per trace for {band_name} (Unknown or default model detected).")


        # Calculate total number of segments for the current band based on desired RBW and sweep points
        full_band_span_hz = band_stop_freq_hz - band_start_freq_hz
        if full_band_span_hz <= 0:
            total_segments_in_band = 1
            optimal_segment_span_hz = full_band_span_hz
        else:
            # Recalculate optimal_segment_span_hz to perfectly divide the band into equal segments
            # This ensures all segments have the same span, even if it slightly deviates from rbw_step_size_hz
            # The old code used scan_rbw_segmentation here, which maps to rbw_step_size_hz in new
            total_segments_in_band = int(np.ceil(full_band_span_hz / (rbw_step_size_hz * (expected_sweep_points - 1))))
            if total_segments_in_band == 0: # Ensure at least one segment for non-zero spans
                total_segments_in_band = 1
            
            # Now calculate the actual segment span to ensure equal division
            optimal_segment_span_hz = full_band_span_hz / total_segments_in_band
            # Ensure it's at least the minimum possible span for expected_sweep_points > 1
            if expected_sweep_points > 1 and optimal_segment_span_hz < (rbw_step_size_hz * (expected_sweep_points - 1)):
                optimal_segment_span_hz = rbw_step_size_hz * (expected_sweep_points - 1)


        # Calculate the effective stop frequency for the scan based on equal segments
        effective_scan_stop_freq_hz = band_start_freq_hz + (total_segments_in_band * optimal_segment_span_hz)
        app_console_update_func(f"🎯 Optimal segment span for {band_name}: {optimal_segment_span_hz / MHZ_TO_HZ:.3f} MHz.")
        app_console_update_func(f"📏 Effective scanned range for equal segments: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {effective_scan_stop_freq_hz/MHZ_TO_HZ:.3f} MHz.")


        current_segment_start_freq_hz = band_start_freq_hz
        segment_counter = 0

        # Loop until the current segment start frequency reaches or exceeds the effective scan stop frequency
        while current_segment_start_freq_hz < effective_scan_stop_freq_hz:
            # Check for pause state at the beginning of each segment
            while pause_event.is_set():
                app_console_update_func("Scan Paused. Click Resume to continue.")
                time.sleep(0.1) # Sleep briefly while paused
                if stop_event.is_set(): # Allow stopping even when paused
                    app_console_update_func(f"Scan for {band_name} interrupted during pause in segment {segment_counter + 1}.")
                    # Process and return data collected so far for the entire sweep
                    return _process_raw_scan_data(raw_scan_data_for_current_sweep, overall_start_freq_hz, overall_stop_freq_hz), last_successful_band_index, csv_filename_current_cycle
            if stop_event.is_set(): # Check scan flag after pause loop
                app_console_update_func(f"Scan for {band_name} interrupted during segment {segment_counter + 1}.")
                # Process and return data collected so far for the entire sweep
                return _process_raw_scan_data(raw_scan_data_for_current_sweep, overall_start_freq_hz, overall_stop_freq_hz), last_successful_band_index, csv_filename_current_cycle

            segment_counter += 1
            # The segment stop frequency is now always based on the optimal span
            segment_stop_freq_hz = current_segment_start_freq_hz + optimal_segment_span_hz

            # Set instrument frequency range for the current segment
            write_safe(inst, f":SENS:FREQ:STAR {current_segment_start_freq_hz};:SENS:FREQ:STOP {segment_stop_freq_hz}")

            # Set trace modes as in the old scanner
            write_safe(inst, ":TRAC1:MODE BLANk;:TRAC2:MODE BLANk;:TRAC3:MODE BLANk")
            if maxhold_enabled: # Only set MaxHold if it's enabled in the GUI
                write_safe(inst, ":TRAC2:MODE MAXHold;")

            # Add settling time for max hold values to show up, if max hold is enabled and time > 0
            if maxhold_enabled and max_hold_time > 0:
                for _ in range(int(max_hold_time * 10)): # Check every 0.1 seconds
                    while pause_event.is_set():
                        app_console_update_func("Scan Paused. Click Resume to continue.")
                        time.sleep(0.1) # Sleep briefly while paused
                        if stop_event.is_set(): # Allow stopping even when paused
                            app_console_update_func(f"Scan for {band_name} interrupted during pause in max hold for segment {segment_counter + 1}.")
                            return _process_raw_scan_data(raw_scan_data_for_current_sweep, overall_start_freq_hz, overall_stop_freq_hz), last_successful_band_index, csv_filename_current_cycle
                    if stop_event.is_set(): # Check scan flag after pause loop
                        app_console_update_func(f"Scan for {band_name} interrupted during max hold for segment {segment_counter + 1}.")
                        return _process_raw_scan_data(raw_scan_data_for_current_sweep, overall_start_freq_hz, overall_stop_freq_hz), last_successful_band_index, csv_filename_current_cycle

                    # Update display for countdown (only update every second for cleaner output)
                    if _ % 10 == 0: # Every 10 iterations (1 second)
                        sec_remaining = int(max_hold_time - (_ / 10))
                        display_text = f"⏳ {sec_remaining}"
                        app_console_update_func(display_text)
                    time.sleep(0.1) # Small sleep to allow other threads/GUI to run
                # Clear the line before printing the final message
                app_console_update_func("✅") # Overwrite with final checkmark

            if stop_event.is_set(): # Check scan flag after max hold loop
                app_console_update_func(f"Scan for {band_name} interrupted after max hold for segment {segment_counter + 1}.")
                return _process_raw_scan_data(raw_scan_data_for_current_sweep, overall_start_freq_hz, overall_stop_freq_hz), last_successful_band_index, csv_filename_current_cycle

            # Calculate progress for the emoji bar - Using more compatible ASCII characters
            progress_percentage = (segment_counter / total_segments_in_band)
            bar_length = 20 # Total number of characters in the bar
            filled_length = int(round(bar_length * progress_percentage))
            # Using '█' (U+2588 Full Block) and '-' (Hyphen) for better compatibility
            progressbar = '█' * filled_length + '-' * (bar_length - filled_length)

            # Combined print statement as per user request, now using _update_console_line with overwrite
            progress_message = f"{progressbar}🔍 Span:📊{optimal_segment_span_hz/MHZ_TO_HZ:.3f} MHz--📈{current_segment_start_freq_hz/MHZ_TO_HZ:.3f} MHz to 📉{segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz   ✅{segment_counter} of {total_segments_in_band} "
            app_console_update_func(progress_message)

            # Read and process trace data
            trace_data = []
            segment_raw_data = [] # To store raw data for the current segment before filtering
            try:
                # Conditional trace data query based on instrument model
                if app_instance_ref.instrument_model == "N9340B":
                    trace_data_str = query_safe(inst, ":TRAC2:DATA?")
                elif app_instance_ref.instrument_model == "N9342CN":
                    trace_data_str = query_safe(inst, ":TRACe:DATA? TRACe2")
                else: # Fallback for unknown models, or other N9342 variants if needed
                    trace_data_str = query_safe(inst, ":TRACe:DATA? TRACe2")


                if trace_data_str is None or "[Not Supported or Timeout]" in trace_data_str or not trace_data_str.strip():
                    app_console_update_func("🚫 No valid trace data string received for this segment.")
                    current_segment_start_freq_hz = segment_stop_freq_hz
                    continue # Move to the next segment if no data

                data_part = None
                # First, try to match the header format
                match = re.match(r'#\d+\d+(.*)', trace_data_str)
                if match:
                    data_part = match.group(1)
                else:
                    # If header not found, assume the string itself is the data part
                    data_part = trace_data_str
                    debug_print("Trace data header not found. Attempting to parse raw string directly.", file=current_file, function=current_function)

                if data_part:
                    try:
                        # Split the string by comma and convert each part to float
                        amplitudes_dbm = [float(val) for val in data_part.split(',') if val.strip()]
                        debug_print(f"Parsed {len(amplitudes_dbm)} amplitude points.", file=current_file, function=current_function)

                        # Query actual start/stop frequencies and number of points from instrument
                        actual_center_freq_hz = float(query_safe(inst, ":SENSe:FREQuency:CENTer?"))
                        actual_span_hz = float(query_safe(inst, ":SENSe:FREQuency:SPAN?"))
                        num_points = 2  ################## THIS IS MESSES UP

                        if num_points == 0:
                            app_console_update_func(f"⚠️ Warning: Instrument reported 0 sweep points for segment {segment_idx + 1}. Skipping data processing for this segment.")
                            debug_print("Instrument reported 0 sweep points.", file=current_file, function=current_function)
                            current_segment_start_freq_hz = segment_stop_freq_hz # Ensure we advance
                            continue # Skip this segment if no points

                        # Generate frequency points for the current segment's trace
                        if num_points > 1:
                            segment_trace_start_freq = actual_center_freq_hz - (actual_span_hz / 2)
                            segment_trace_stop_freq = actual_center_freq_hz + (actual_span_hz / 2)
                            frequencies_hz = np.linspace(segment_trace_start_freq, segment_trace_stop_freq, num_points)
                        else: # Handle case with a single point
                            frequencies_hz = np.array([actual_center_freq_hz])

                        # Ensure amplitudes and frequencies match in length
                        if len(amplitudes_dbm) == len(frequencies_hz):
                            for freq, amp in zip(frequencies_hz, amplitudes_dbm):
                                raw_scan_data_for_current_sweep.append((freq, amp))
                                segment_raw_data.append((freq, amp)) # Also add to segment_raw_data for immediate CSV write
                            app_console_update_func(f"  Collected {len(amplitudes_dbm)} data points for segment {segment_idx + 1}.")
                            debug_print(f"Collected {len(amplitudes_dbm)} data points for segment {segment_idx + 1}.", file=current_file, function=current_function)
                        else:
                            app_console_update_func(f"❌ Error: Mismatch in data points and frequencies for segment {segment_idx + 1}. Expected {len(frequencies_hz)}, got {len(amplitudes_dbm)} amplitudes. Raw data: {data_part[:100]}...")
                            debug_print(f"Data length mismatch. Expected {len(frequencies_hz)}, got {len(amplitudes_dbm)}.", file=current_file, function=current_function)
                            # Attempt to continue to next segment, but this is a data integrity issue
                    except ValueError as ve:
                        app_console_update_func(f"❌ Data Parsing Error in segment {segment_idx + 1}: {ve}. Could not convert data to float. Raw data: {data_part[:100]}...")
                        debug_print(f"ValueError parsing trace data: {ve}", file=current_file, function=current_function)
                        current_segment_start_freq_hz = segment_stop_freq_hz # Advance to next segment
                        continue
                else:
                    app_console_update_func(f"❌ Error: No parsable data part found in trace data for segment {segment_idx + 1}. Raw data: {trace_data_str[:100]}...")
                    debug_print(f"No parsable data part found: {trace_data_str[:50]}...", file=current_file, function=current_function)
                    current_segment_start_freq_hz = segment_stop_freq_hz # Advance to next segment
                    continue
                
                # --- Write filtered segment data to CSV immediately after processing ---
                if segment_raw_data: # Only write if there's data
                    filtered_segment_data_for_csv = []
                    for freq_hz, amp_value in segment_raw_data:
                        # Filter based on the original band's start and stop frequencies
                        if freq_hz >= band_start_freq_hz and freq_hz <= band_stop_freq_hz + 1e-9: # Add epsilon for float comparison
                            filtered_segment_data_for_csv.append((freq_hz, amp_value))

                    if filtered_segment_data_for_csv:
                        header = ["Frequency_MHz", "Power_dBm"] # Define header for this CSV
                        # Convert frequencies to MHz for the CSV
                        csv_data_to_write = [(f / MHZ_TO_HZ, amp) for f, amp in filtered_segment_data_for_csv]
                        write_scan_data_to_csv(csv_filename_current_cycle, header, csv_data_to_write, append_mode=True)
                        debug_print(f"Appended {len(filtered_segment_data_for_csv)} points to {csv_filename_current_cycle}", file=current_file, function=current_function)
                    else:
                        debug_print("No data to append to CSV after filtering for this segment.", file=current_file, function=current_function)
                else:
                    debug_print("Segment raw data was empty, nothing to write to CSV.", file=current_file, function=current_function)

                # Update last_successful_band_index after successfully processing a band
                last_successful_band_index = i

            except pyvisa.errors.VisaIOError as e:
                app_console_update_func(f"❌ VISA Error during segment {segment_idx + 1} scan: {e}")
                debug_print(f"VISA Error in scan_bands segment: {e}", file=current_file, function=current_function)
                break # For critical errors, break out of the outer loop as well
            except Exception as e:
                app_console_update_func(f"🚨 An unexpected error occurred during trace processing: {e}\n")
                debug_print(f"Unexpected error during trace processing: {e}", file=current_file, function=current_function)
                current_segment_start_freq_hz = segment_stop_freq_hz # Advance to next segment
                continue # Move to the next segment if other error occurs

            # Move to the start of the next segment
            current_segment_start_freq_hz = segment_stop_freq_hz

    # --- After all bands are scanned, or if interrupted, process the collected raw data for the full sweep ---
    app_console_update_func("\n--- 🎉 Band Scan Complete! Processing collected data... ---\n")
    final_sweep_data_for_plotting = _process_raw_scan_data(
        raw_scan_data_for_current_sweep,
        overall_start_freq_hz, # Use the calculated overall start
        overall_stop_freq_hz # Use the calculated overall stop
    )

    if not final_sweep_data_for_plotting.empty: # Check if DataFrame is not empty
        app_console_update_func(
                               f"✅ De-duplicated and filtered {len(final_sweep_data_for_plotting)} points for full scan.")
        
        # The raw CSV has already been written incrementally.
        # This part is for the final processed data or if a single file was desired.
        # Keeping this block for consistency with previous version's flow,
        # but the primary raw data saving is now per-segment.
        # If a separate "final processed CSV" is desired, uncomment and adjust.
        # timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        # final_csv_filename = os.path.join(output_folder, f"{scan_name}_Processed_Scan_{timestamp}.csv")
        # try:
        #     final_sweep_data_for_plotting.to_csv(final_csv_filename, index=False)
        #     app_console_update_func(f"✅ Processed scan data saved to: {final_csv_filename}")
        #     debug_print(f"Processed scan data saved to {final_csv_filename}", file=current_file, function=current_function)
        # except Exception as e:
        #     app_console_update_func(f"❌ Error: Failed to save processed scan CSV: {e}")
        #     debug_print(f"Error saving processed scan CSV: {e}", file=current_file, function=current_function)


        # --- Extract markers from the final sweep data (if enabled) ---
        # This is a placeholder. Actual marker extraction logic would go here.
        # For a basic implementation, we can simulate markers or look for peaks.
        # This part assumes MARKERS.CSV generation happens externally or is loaded from a report.
        # Here, we can create dummy markers for demonstration or integrate a real peak detection.
        # Since the request is to eliminate messagebox and markers data is passed back
        # to scan_logic and then to plot_single_scan_data, we'll keep it simple for now.
        
        # For the purpose of this task, let's assume markers_data_from_scan is populated
        # by some external means or a future enhancement. For now, it remains empty
        # or can be a dummy. The main_app.py passes its `last_scan_markers` to plotting logic.
        
        debug_print(f"Exiting {current_function} function. Result: {last_successful_band_index}, DataFrame, Markers Data", file=current_file, function=current_function)
        return last_successful_band_index, final_sweep_data_for_plotting, markers_data_from_scan # Return the DataFrame and markers
    else:
        app_console_update_func(
                               f"🚫 No data collected for full scan after de-duplication attempt for band: {selected_bands[last_successful_band_index]['Band Name'] if last_successful_band_index != -1 else 'N/A'}.")
        debug_print("Final sweep data is empty after processing.", file=current_file, function=current_function)
        debug_print(f"Exiting {current_function} function. Result: {last_successful_band_index}, None, None", file=current_file, function=current_function)
        return last_successful_band_index, None, None # Return None for filename if no data after processing

