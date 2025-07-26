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
import csv # Added import for csv module

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


def scan_bands(app_instance_ref, inst, selected_bands, scan_rbw_segmentation, rbw_config_val, vbw_config_val, max_hold_time, current_freq_offset, last_scanned_band_index=0):
    """
    Iterates through predefined frequency bands, sets the start/stop frequencies,
    reduces RBW to 10000 Hz, and triggers a sweep for each band.
    It extracts trace data, and also returns it for further processing (plotting and CSV writing).
    This function now dynamically segments bands to maintain a consistent
    effective resolution bandwidth per trace point, ensuring equal spans for all segments.
    It also displays the time of day for each band scanned.

    Added last_scanned_band_index to resume from where it left off.
    CSV data is now continuously written after each segment.

    Args:
        app_instance_ref (object): A reference to the main Tkinter App instance for GUI updates.
                                    Expected to have .scanning, .paused, .after, ._update_console_line, .instrument_model attributes.
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        selected_bands (list): A list of band dictionaries to scan.
        scan_rbw_segmentation (float): Resolution Bandwidth for segmenting bands (from new GUI input).
        rbw_config_val (float): Resolution Bandwidth value to configure on the instrument.
        vbw_config_val (float): Video Bandwidth value to configure on the instrument.
        max_hold_time (float): Duration in seconds for which MAX Hold should be active.
                                If > 0, MAX Hold mode is enabled for the scan.
        current_freq_offset (float): The frequency offset (in Hz) to apply to all band frequencies.
        last_scanned_band_index (int): Index of the band to start scanning from.
                                        Used for resuming scans after an error.
    Returns:
        tuple: (list: all_scan_data_cleaned, int: last_successful_band_index, str: csv_filename_current_cycle)
               all_scan_data_cleaned: A list of de-duplicated and filtered data points
                                      for the entire sweep (or up to interruption).
               last_successful_band_index: The index of the last band that was fully
                                           or partially scanned successfully.
               csv_filename_current_cycle: The full path to the CSV file where the scan data was saved.
    """
    # This will accumulate all raw data points for the *current single scan sweep*
    # before final de-duplication for plotting.
    raw_scan_data_for_current_sweep = []
    last_successful_band_index = last_scanned_band_index

    print("\n--- 📡 Starting Band Scan ---")

    print("💾 Assuming ASCII data format for trace data.")

    # Apply RBW and VBW settings here once at the start of the scan
    # Use scan_rbw_segmentation for the instrument's RBW during the scan
    write_safe(inst, f":SENSE:BAND:RES {scan_rbw_segmentation}")
    print(f"📏 Set RBW to {scan_rbw_segmentation/1000:.0f} kHz for scan (from Scan RBW setting).")
    write_safe(inst, f":SENSE:BAND:VID {vbw_config_val}")
    print(f"📺 Set VBW to {vbw_config_val} Hz for scan.")

    # Determine the CSV filename for this scan session (continuous raw data)
    # This file is now explicitly for the raw data of the current scan cycle.
    scan_name = app_instance_ref.scan_name_var.get()
    output_folder = app_instance_ref.output_folder_var.get()
    timestamp_hm = datetime.datetime.now().strftime("%Y%m%d_%H%M%S") # YYYYMMDD_HHMM (no seconds)
    
    # Define the CSV filename for the *current scan cycle's raw data*
    csv_filename_current_cycle = os.path.join(output_folder, f"{scan_name}_RBW{int(rbw_config_val/1000)}K_HOLD{int(max_hold_time)}_Offset{int(current_freq_offset)}_{timestamp_hm}.csv")


    # Ensure output directory exists before starting to write
    os.makedirs(output_folder, exist_ok=True)

    # Calculate overall start and stop frequencies for the entire selected_bands list
    overall_start_freq_hz = (selected_bands[0]["Start MHz"] * MHZ_TO_HZ) + current_freq_offset
    overall_stop_freq_hz = (selected_bands[-1]["Stop MHz"] * MHZ_TO_HZ) + current_freq_offset


    # *** Use selected_bands for scanning the instrument ***
    # Iterate through bands starting from last_scanned_band_index
    for i in range(last_scanned_band_index, len(selected_bands)):
        band = selected_bands[i]
        band_name = band["Band Name"]
        # Apply the frequency offset here
        band_start_freq_hz = (band["Start MHz"] * MHZ_TO_HZ) + current_freq_offset
        band_stop_freq_hz = (band["Stop MHz"] * MHZ_TO_HZ) + current_freq_offset

        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M") # Use datetime.datetime
        print(f"\n📈 [{current_time}] Processing Band: {band_name} (Shifted Range: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {band_stop_freq_hz/MHZ_TO_HZ:.3f} MHz)")

        # Determine actual_sweep_points based on instrument model
        if app_instance_ref.instrument_model == "N9340B":
            expected_sweep_points = 461
            print(f"📊 Using {expected_sweep_points} sweep points per trace for {band_name} (N9340B detected).")
        elif app_instance_ref.instrument_model == "N9342CN": # Explicitly check for N9342CN
            expected_sweep_points = 500
            print(f"📊 Using {expected_sweep_points} sweep points per trace for {band_name} (N9342CN detected).")
        else: # Default for unknown models, or other N9342 variants if needed
            expected_sweep_points = 500
            print(f"📊 Using {expected_sweep_points} sweep points per trace for {band_name} (Unknown or default model detected).")


        # Calculate total number of segments for the current band based on desired RBW and sweep points
        full_band_span_hz = band_stop_freq_hz - band_start_freq_hz
        if full_band_span_hz <= 0:
            total_segments_in_band = 1
            # For a single point or zero span, the segment span is just the full band span
            optimal_segment_span_hz = full_band_span_hz
        else:
            # Recalculate optimal_segment_span_hz to perfectly divide the band into equal segments
            # This ensures all segments have the same span, even if it slightly deviates from scan_rbw_segmentation
            total_segments_in_band = int(np.ceil(full_band_span_hz / (scan_rbw_segmentation * (expected_sweep_points - 1))))
            if total_segments_in_band == 0: # Ensure at least one segment for non-zero spans
                total_segments_in_band = 1
            
            # Now calculate the actual segment span to ensure equal division
            optimal_segment_span_hz = full_band_span_hz / total_segments_in_band
            # Ensure it's at least the minimum possible span for expected_sweep_points > 1
            if expected_sweep_points > 1 and optimal_segment_span_hz < (scan_rbw_segmentation * (expected_sweep_points - 1)):
                optimal_segment_span_hz = scan_rbw_segmentation * (expected_sweep_points - 1)


        # Calculate the effective stop frequency for the scan based on equal segments
        # This will be used to control the while loop, ensuring we cover the full band with equal segments
        effective_scan_stop_freq_hz = band_start_freq_hz + (total_segments_in_band * optimal_segment_span_hz)
        print(f"🎯 Optimal segment span for {band_name}: {optimal_segment_span_hz / MHZ_TO_HZ:.3f} MHz.")
        print(f"📏 Effective scanned range for equal segments: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {effective_scan_stop_freq_hz/MHZ_TO_HZ:.3f} MHz.")


        current_segment_start_freq_hz = band_start_freq_hz
        segment_counter = 0

        # Loop until the current segment start frequency reaches or exceeds the effective scan stop frequency
        while current_segment_start_freq_hz < effective_scan_stop_freq_hz:
            # Check for pause state at the beginning of each segment
            while app_instance_ref.paused:
                app_instance_ref.after(100, app_instance_ref._update_console_line, "Scan Paused. Click Resume to continue.", False)
                time.sleep(0.1) # Sleep briefly while paused
                if not app_instance_ref.scanning: # Allow stopping even when paused
                    print(f"Scan for {band_name} interrupted during pause in segment {segment_counter + 1}.")
                    # Process and return data collected so far for the entire sweep
                    return _process_raw_scan_data(raw_scan_data_for_current_sweep, overall_start_freq_hz, overall_stop_freq_hz), last_successful_band_index, csv_filename_current_cycle
            if not app_instance_ref.scanning: # Check scan flag after pause loop
                print(f"Scan for {band_name} interrupted during segment {segment_counter + 1}.")
                # Process and return data collected so far for the entire sweep
                return _process_raw_scan_data(raw_scan_data_for_current_sweep, overall_start_freq_hz, overall_stop_freq_hz), last_successful_band_index, csv_filename_current_cycle

            segment_counter += 1
            # The segment stop frequency is now always based on the optimal span
            segment_stop_freq_hz = current_segment_start_freq_hz + optimal_segment_span_hz

            # Set instrument frequency range for the current segment
            write_safe(inst, f":SENS:FREQ:STAR {current_segment_start_freq_hz};:SENS:FREQ:STOP {segment_stop_freq_hz}")

            write_safe(inst, ":TRAC1:MODE BLANk;:TRAC2:MODE BLANk;:TRAC3:MODE BLANk")
            write_safe(inst, ":TRAC2:MODE MAXHold;")

            # Add settling time for max hold values to show up, if max hold is enabled
            if max_hold_time > 0:
                for _ in range(int(max_hold_time * 10)): # Check every 0.1 seconds
                    while app_instance_ref.paused:
                        app_instance_ref.after(100, app_instance_ref._update_console_line, "Scan Paused. Click Resume to continue.", False)
                        time.sleep(0.1) # Sleep briefly while paused
                        if not app_instance_ref.scanning: # Allow stopping even when paused
                            print(f"Scan for {band_name} interrupted during pause in max hold for segment {segment_counter + 1}.")
                            # Process and return data collected so far for the entire sweep
                            return _process_raw_scan_data(raw_scan_data_for_current_sweep, overall_start_freq_hz, overall_stop_freq_hz), last_successful_band_index, csv_filename_current_cycle
                    if not app_instance_ref.scanning: # Check scan flag after pause loop
                        print(f"Scan for {band_name} interrupted during max hold for segment {segment_counter + 1}.")
                        # Process and return data collected so far for the entire sweep
                        return _process_raw_scan_data(raw_scan_data_for_current_sweep, overall_start_freq_hz, overall_stop_freq_hz), last_successful_band_index, csv_filename_current_cycle

                    # Update display for countdown (only update every second for cleaner output)
                    if _ % 10 == 0: # Every 10 iterations (1, second)
                        sec_remaining = int(max_hold_time - (_ / 10))
                        display_text = f"⏳ {sec_remaining}"
                        app_instance_ref.after(0, app_instance_ref._update_console_line, display_text, False)
                    time.sleep(0.1) # Small sleep to allow other threads/GUI to run
                # Clear the line before printing the final message
                app_instance_ref.after(0, app_instance_ref._update_console_line, "✅", False) # Overwrite with final checkmark

            if not app_instance_ref.scanning: # Check scan flag after max hold loop
                print(f"Scan for {band_name} interrupted after max hold for segment {segment_counter + 1}.")
                # Process and return data collected so far for the entire sweep
                return _process_raw_scan_data(raw_scan_data_for_current_sweep, overall_start_freq_hz, overall_stop_freq_hz), last_successful_band_index, csv_filename_current_cycle

            # Calculate progress for the emoji bar - Using more compatible ASCII characters
            progress_percentage = (segment_counter / total_segments_in_band)
            bar_length = 20 # Total number of characters in the bar
            filled_length = int(round(bar_length * progress_percentage))
            # Using '█' (U+2588 Full Block) and '-' (Hyphen) for better compatibility
            progressbar = '█' * filled_length + '-' * (bar_length - filled_length)

            # Combined print statement as per user request, now using _update_console_line with overwrite
            progress_message = f"{progressbar}🔍 Span:📊{optimal_segment_span_hz/MHZ_TO_HZ:.3f} MHz--📈{current_segment_start_freq_hz/MHZ_TO_HZ:.3f} MHz to 📉{segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz   ✅{segment_counter} of {total_segments_in_band} \n"
            # Add '\r' to ensure the next message overwrites this line correctly
            app_instance_ref.after(0, app_instance_ref._update_console_line, progress_message + '\r', True)

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
                    print("🚫 No valid trace data string received for this segment.")
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
                    print("🚫 No trace data received for this segment after parsing.")
                    current_segment_start_freq_hz = segment_stop_freq_hz
                    continue # Move to the next segment if no data

                # Loop to append data to raw_scan_data_for_current_sweep (for later de-duplication for plotting)
                # and to segment_raw_data (for immediate CSV writing after filtering)
                for j, amp_value in enumerate(trace_data):
                    current_freq_for_point_hz = current_segment_start_freq_hz + (j * freq_step_per_point_actual)
                    segment_raw_data.append((current_freq_for_point_hz, amp_value))
                    raw_scan_data_for_current_sweep.append((current_freq_for_point_hz, amp_value))


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
                    write_scan_data_to_csv(csv_filename_current_cycle, header, csv_data_to_write, append_mode=True)
                    # Removed the print statement here as requested


                # Update last_successful_band_index after successfully processing a band
                last_successful_band_index = i

            except pyvisa.errors.VisaIOError as e:
                print(f"🛑 Error reading trace data (PyVISA IO Error): {e}")
                print(f"🐛 Raw data string potentially causing error: {e}")
                raise # Re-raise the exception to be caught by the main loop for recovery
            except ValueError as e:
                print(f"🚫 Error processing ASCII trace data (ValueError - cannot convert/unpack): {e}")
                print(f"🐞 Raw data string for parsing: {e}")
            except Exception as e:
                print(f"🚨 An unexpected error occurred during trace processing: {e}")

            # Move to the start of the next segment
            current_segment_start_freq_hz = segment_stop_freq_hz

    # --- After all bands are scanned, or if interrupted, process the collected raw data for the full sweep ---
    if not raw_scan_data_for_current_sweep:
        print("🚫 No raw data collected across all bands for the current sweep.")
        return None, -1, None # Return -1 for last_successful_band_index if no data

    print("\n--- 🎉 Band Scan Complete! Processing collected data... ---")
    final_sweep_data_for_plotting = _process_raw_scan_data(
        raw_scan_data_for_current_sweep,
        overall_start_freq_hz, # Use the calculated overall start
        overall_stop_freq_hz # Use the calculated overall stop
    )

    if not final_sweep_data_for_plotting:
        print(f"🚫 No data collected for full scan after de-duplication attempt for band: {selected_bands[last_successful_band_index]['Band Name'] if last_successful_band_index != -1 else 'N/A'}.")
    else:
        # Corrected: Use final_sweep_data_for_plotting instead of final_sweep_data_for_current_sweep
        print(f"✅ De-duplicated and filtered {len(final_sweep_data_for_plotting)} points for plotting for full scan.")

    # Save the final processed data to CSV
    try:
        # write_scan_data_to_csv expects (freq_mhz, level_dbm)
        # _process_raw_scan_data already returns (freq_mhz, level_dbm)
        write_scan_data_to_csv(csv_filename_current_cycle, ["Frequency_MHz", "Power_dBm"], final_sweep_data_for_plotting, append_mode=False)
        print(f"✅ Full sweep data saved to: {csv_filename_current_cycle}")
    except Exception as e:
        print(f"❌ Error saving full sweep data to CSV: {e}")
        return None, last_successful_band_index, None

    return final_sweep_data_for_plotting, last_successful_band_index, csv_filename_current_cycle