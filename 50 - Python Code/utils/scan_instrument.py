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
from utils.frequency_bands import SCAN_BAND_RANGES, MHZ_TO_HZ, VBW_RBW_RATIO # Import VBW_RBW_RATIO

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

    # Initialize instrument for scan (basic setup, no specific scan parameters here)
    app_console_update_func("Initializing instrument for scan settings...")
    if not initialize_instrument(
        inst,
        debug_mode=general_debug_enabled,
        model_match=app_instance_ref.instrument_model # Pass the instrument model
    ):
        app_console_update_func("❌ Error: Failed to initialize instrument for scan. Aborting.")
        app_instance_ref.after(0, lambda: print("❌ Error: Failed to initialize instrument for scan. Aborting."))
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

    # Maxhold setting
    maxhold_cmd = ":TRAC2:MODE MAXHold" if maxhold_enabled else ":TRAC2:MODE AVERage" # Assuming Trace 2 for MaxHold
    if not write_safe(inst, maxhold_cmd):
        app_console_update_func(f"❌ Error: Failed to set Max Hold mode to {maxhold_enabled}.")
        debug_print(f"Failed to set Max Hold mode: {maxhold_enabled}", file=current_file, function=current_function)
        # Not critical, continue but log

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


    raw_scan_data_for_current_sweep = [] # List to collect (freq, amplitude) tuples for the entire sweep
    last_successful_band_index = -1 # Keep track of the last band that successfully scanned
    
    markers_data_from_scan = [] # To collect markers if extracted during the scan

    for i, band in enumerate(selected_bands):
        if stop_event.is_set():
            app_console_update_func("Scan stopped by user during band iteration.")
            debug_print("Scan stop event set during band iteration.", file=current_file, function=current_function)
            debug_print(f"Exiting {current_function} function. Result: {last_successful_band_index}, None, None", file=current_file, function=current_function)
            break # Exit loop if stop is requested

        band_name = band["Band Name"]
        start_freq_mhz = band["Start MHz"]
        stop_freq_mhz = band["Stop MHz"]
        start_freq_hz = start_freq_mhz * MHZ_TO_HZ
        stop_freq_hz = stop_freq_mhz * MHZ_TO_HZ

        app_console_update_func(f"\n--- Scanning Band: {band_name} ({start_freq_mhz:.3f} MHz - {stop_freq_mhz:.3f} MHz) ---")
        debug_print(f"Processing band: {band_name} ({start_freq_hz}-{stop_freq_hz} Hz)", file=current_file, function=current_function)

        # Break down band into segments if the span is too large for a single trace
        # Example: max span of 100 MHz for some instruments
        max_segment_span_hz = 100 * MHZ_TO_HZ # Arbitrary limit, adjust as per instrument capability
        # Use rbw_step_size_hz to determine segmentation or default to a fixed size
        # This is where 'default_scan_rbw_segmentation' from config.ini would be useful.
        # For now, let's use a dynamic calculation based on current RBW, aiming for approx 1000 points
        
        # Calculate segments based on rbw_step_size_hz (which is max RBW for a single segment)
        # This is a simplification; a real instrument might have a fixed max span.
        segment_step = rbw_step_size_hz # Assuming rbw_step_size_hz is the max desired RBW for a segment
        if segment_step == 0:
            segment_step = 1000000 # Default to 1 MHz segments if 0 for some reason
        
        # Dynamically determine number of points to request per segment.
        # This calculation needs to be aligned with how the instrument provides trace points.
        # For Agilent/Keysight, it's often 401, 801, etc., points.
        # If we divide band by segment_step, we get number of "resolution units".
        # Let's target approximately 1000 points per segment if instrument supports it.
        # Or simply use fixed segments if the instrument has a strict max span.
        
        # For now, let's use the explicit rbw_step_size_hz for segmenting if it's large enough.
        # Otherwise, default to max_segment_span_hz.
        actual_segment_span_hz = max(rbw_step_size_hz, max_segment_span_hz)

        num_segments = int(np.ceil((stop_freq_hz - start_freq_hz) / actual_segment_span_hz))
        if num_segments == 0: num_segments = 1 # Ensure at least one segment for very small bands

        current_segment_start_freq_hz = start_freq_hz

        for segment_idx in range(num_segments):
            if stop_event.is_set():
                app_console_update_func("Scan stopped by user during segment iteration.")
                debug_print("Scan stop event set during segment iteration.", file=current_file, function=current_function)
                break

            while pause_event.is_set():
                app_console_update_func("Scan paused. Waiting to resume...")
                debug_print("Scan paused, waiting for resume.", file=current_file, function=current_function)
                time.sleep(1) # Check every second if unpaused

            segment_stop_freq_hz = min(current_segment_start_freq_hz + actual_segment_span_hz, stop_freq_hz)
            segment_center_freq_hz = (current_segment_start_freq_hz + segment_stop_freq_hz) / 2
            segment_span_hz = segment_stop_freq_hz - current_segment_start_freq_hz

            app_console_update_func(f"  Segment {segment_idx + 1}/{num_segments}: {current_segment_start_freq_hz/MHZ_TO_HZ:.3f} MHz - {segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz (Span: {segment_span_hz/MHZ_TO_HZ:.3f} MHz)")
            debug_print(f"  Configuring segment: Center={segment_center_freq_hz} Hz, Span={segment_span_hz} Hz", file=current_file, function=current_function)

            try:
                # Set instrument's center frequency and span for the current segment
                if not write_safe(inst, f":FREQuency:CENTer {segment_center_freq_hz}HZ"):
                    app_console_update_func(f"❌ Error: Failed to set center frequency for segment {segment_idx + 1}.")
                    debug_print(f"Failed to set center freq: {segment_center_freq_hz}Hz", file=current_file, function=current_function)
                    break # Abort band if critical setting fails

                if not write_safe(inst, f":FREQuency:SPAN {segment_span_hz}HZ"):
                    app_console_update_func(f"❌ Error: Failed to set span for segment {segment_idx + 1}.")
                    debug_print(f"Failed to set span: {segment_span_hz}Hz", file=current_file, function=current_function)
                    break # Abort band if critical setting fails
                
                # Apply RBW and VBW (assuming they are consistent across segments or dynamically adjusted)
                # Ensure VBW is set as a ratio of RBW, if supported by instrument.
                vbw_percent = int(VBW_RBW_RATIO * 100) # Convert 1/3 to 33%
                if not write_safe(inst, f":BANDwidth:RESolution {rbw_hz}HZ"):
                    app_console_update_func(f"❌ Error: Failed to set RBW for segment {segment_idx + 1}.")
                    debug_print(f"Failed to set RBW: {rbw_hz}Hz", file=current_file, function=current_function)
                    break
                if not write_safe(inst, f":BANDwidth:VIDeo:RATio {vbw_percent}PCT"): # Example command, verify for specific instrument
                    app_console_update_func(f"❌ Error: Failed to set VBW for segment {segment_idx + 1}.")
                    debug_print(f"Failed to set VBW ratio: {vbw_percent}PCT", file=current_file, function=current_function)
                    # This might not be critical, so continue, but log error
                
                # Take single trace measurement
                # Initiate a single sweep and wait for it to complete
                write_safe(inst, ":INITiate:IMMediate")
                write_safe(inst, "*WAI") # Wait for operation to complete

                # Fetch trace data
                trace_data_str = query_safe(inst, ":TRACe:DATA? TRACE1")
                debug_print(f"Raw trace data string length: {len(trace_data_str) if trace_data_str else 0}", file=current_file, function=current_function)

                if trace_data_str:
                    # Parse the trace data string (usually comma-separated floats)
                    # The format can vary. Keysight N9340B typically returns ASCII data
                    # like '#<num_digits><total_bytes><data_bytes>...'.
                    # We need to extract the actual data portion.
                    match = re.match(r'#\d+\d+(.*)', trace_data_str)
                    if match:
                        data_part = match.group(1)
                        amplitudes_dbm = [float(x) for x in data_part.strip().split(',') if x.strip()]
                        debug_print(f"Parsed {len(amplitudes_dbm)} amplitude points.", file=current_file, function=current_function)

                        # Query actual start/stop frequencies and number of points from instrument
                        # This is more robust than calculating locally, especially with auto settings
                        actual_center_freq_hz = float(query_safe(inst, ":SENSe:FREQuency:CENTer?"))
                        actual_span_hz = float(query_safe(inst, ":SENSe:FREQuency:SPAN?"))
                        num_points = int(query_safe(inst, ":SWEep:POINts?"))

                        if num_points == 0:
                            app_console_update_func(f"⚠️ Warning: Instrument reported 0 sweep points for segment {segment_idx + 1}. Skipping data processing for this segment.")
                            debug_print("Instrument reported 0 sweep points.", file=current_file, function=current_function)
                            continue # Skip this segment if no points

                        # Generate frequency points for the current segment's trace
                        # This assumes linear sweep over the span with 'num_points'
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
                            app_console_update_func(f"  Collected {len(amplitudes_dbm)} data points for segment {segment_idx + 1}.")
                            debug_print(f"Collected {len(amplitudes_dbm)} data points for segment {segment_idx + 1}.", file=current_file, function=current_function)
                        else:
                            app_console_update_func(f"❌ Error: Mismatch in data points and frequencies for segment {segment_idx + 1}. Expected {len(frequencies_hz)}, got {len(amplitudes_dbm)} amplitudes.")
                            debug_print(f"Data length mismatch. Expected {len(frequencies_hz)}, got {len(amplitudes_dbm)}.", file=current_file, function=current_function)
                            # Attempt to continue to next segment, but this is a data integrity issue
                    else:
                        app_console_update_func(f"❌ Error: Could not parse trace data format from instrument for segment {segment_idx + 1}.")
                        debug_print(f"Trace data regex mismatch: {trace_data_str[:50]}...", file=current_file, function=current_function)
                else:
                    app_console_update_func(f"⚠️ Warning: No trace data received from instrument for segment {segment_idx + 1}.")
                    debug_print("No trace data received.", file=current_file, function=current_function)

            except pyvisa.errors.VisaIOError as e:
                app_console_update_func(f"❌ VISA Error during segment {segment_idx + 1} scan: {e}")
                app_instance_ref.after(0, lambda: print(f"❌ VISA Error: A VISA communication error occurred during segment scan: {e}"))
                debug_print(f"VISA Error in scan_bands segment: {e}", file=current_file, function=current_function)
                # For critical errors, consider breaking out of the outer loop as well
                break
            except ValueError as e:
                app_console_update_func(f"❌ Data Parsing Error in segment {segment_idx + 1}: {e}. Trace data might be invalid.")
                debug_print(f"ValueError parsing trace data: {e}", file=current_file, function=current_function)
                # Continue to next segment if data parsing fails for one segment
            except Exception as e:
                app_console_update_func(f"🚨 An unexpected error occurred during trace processing: {e}\n")
                debug_print(f"Unexpected error during trace processing: {e}", file=current_file, function=current_function)
                continue # Move to the next segment if other error occurs

            # Move to the start of the next segment
            current_segment_start_freq_hz = segment_stop_freq_hz
        
        # If the band was completed without critical errors, update last_successful_band_index
        if not stop_event.is_set() and segment_idx + 1 == num_segments: # Check if inner loop completed all segments
            last_successful_band_index = i

    # --- After all bands are scanned, or if interrupted, process the collected raw data for the full sweep ---
    app_instance_ref.after(0, app_instance_ref._update_console_line, "\n--- 🎉 Band Scan Complete! Processing collected data... ---\n")
    final_sweep_data_for_plotting = _process_raw_scan_data(
        raw_scan_data_for_current_sweep,
        overall_start_freq_hz, # Use the calculated overall start
        overall_stop_freq_hz # Use the calculated overall stop
    )

    if not final_sweep_data_for_plotting.empty: # Check if DataFrame is not empty
        app_instance_ref.after(0, app_instance_ref._update_console_line, 
                               f"✅ De-duplicated and filtered {len(final_sweep_data_for_plotting)} points for full scan.")
        
        # --- Save the combined raw data to a CSV file ---
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        raw_csv_filename = os.path.join(output_folder, f"{scan_name}_Raw_Scan_{timestamp}.csv")
        
        try:
            # Prepare data as list of lists for csv_utils
            csv_data = final_sweep_data_for_plotting[['Frequency (MHz)', 'Amplitude (dBm)']].values.tolist()
            csv_header = ['Frequency (MHz)', 'Amplitude (dBm)']
            
            write_scan_data_to_csv(raw_csv_filename, csv_header, csv_data, append_mode=False)
            app_instance_ref.after(0, app_instance_ref._update_console_line, f"✅ Raw scan data saved to: {raw_csv_filename}")
            debug_print(f"Raw scan data saved to {raw_csv_filename}", file=current_file, function=current_function)
        except Exception as e:
            app_instance_ref.after(0, app_instance_ref._update_console_line, f"❌ Error: Failed to save raw scan CSV: {e}")
            app_instance_ref.after(0, lambda: print(f"❌ Error: Could not save raw scan CSV: {e}"))
            debug_print(f"Error saving raw scan CSV: {e}", file=current_file, function=current_function)
            # Do not return, as we still want to return the DataFrame even if saving failed

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
        app_instance_ref.after(0, app_instance_ref._update_console_line, 
                               f"🚫 No data collected for full scan after de-duplication attempt for band: {selected_bands[last_successful_band_index]['Band Name'] if last_successful_band_index != -1 else 'N/A'}.")
        debug_print("Final sweep data is empty after processing.", file=current_file, function=current_function)
        debug_print(f"Exiting {current_function} function. Result: {last_successful_band_index}, None, None", file=current_file, function=current_function)
        return last_successful_band_index, None, None # Return None for filename if no data after processing
