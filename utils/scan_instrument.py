# scan_instrument.py

import pyvisa
import time
import numpy as np
import struct # Not used in current version, but kept if needed for future binary data handling
import re
import tkinter as tk # For messagebox, used in _debug_mode_enabled and within scan_bands
import datetime # Added import for datetime

# Import instrument control functions
from utils.instrument_control import query_safe, write_safe

# Import constants from frequency_bands.py
try:
    from utils.frequency_bands import MHZ_TO_HZ
except ImportError:
    print("Error: frequency_bands.py not found. Please ensure it's in the same directory.")
    MHZ_TO_HZ = 1_000_000 # Define dummy value

# Note: _debug_mode_enabled and set_debug_mode are now in instrument_control.py


# initialize_instrument has been moved to instrument_control.py

# FIX: Removed csv_writer from the function signature
def scan_bands(app_instance_ref, inst, selected_bands, scan_rbw_segmentation, rbw_config_val, vbw_config_val, max_hold_time, current_freq_offset, last_scanned_band_index=0):
    """
    Iterates through predefined frequency bands, sets the start/stop frequencies,
    reduces RBW to 10000 Hz, and triggers a sweep for each band.
    It extracts trace data, and also returns it for further processing (plotting and CSV writing).
    This function now dynamically segments bands to maintain a consistent
    effective resolution bandwidth per trace point.
    It also displays the time of day for each band scanned.
    
    Added last_scanned_band_index to resume from where it left off.

    Args:
        app_instance_ref (object): A reference to the main Tkinter App instance for GUI updates.
                                   Expected to have .scanning, .paused, .after, .update_progress_label, ._update_console_line, .instrument_model attributes.
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
        tuple: (list: all_scan_data, int: last_successful_band_index)
               all_scan_data: A list of dictionaries, where each dictionary represents a data point
                              with 'Band Name', 'Frequency (Hz)', and 'Level (dBm)'.
               last_successful_band_index: The index of the last band that was fully
                                           or partially scanned successfully.
    """
    all_scan_data = [] # To store all data points across all bands for plotting
    last_successful_band_index = last_scanned_band_index

    print("\n--- 📡 Starting Band Scan ---")

    print("💾 Assuming ASCII data format for trace data.")

    # Apply RBW and VBW settings here once at the start of the scan
    # Use scan_rbw_segmentation for the instrument's RBW during the scan
    write_safe(inst, f":SENSE:BAND:RES {scan_rbw_segmentation}")
    print(f"📏 Set RBW to {scan_rbw_segmentation/1000:.0f} kHz for scan (from Scan RBW setting).")
    write_safe(inst, f":SENSE:BAND:VID {vbw_config_val}")
    print(f"📺 Set VBW to {vbw_config_val} Hz for scan.")


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
            actual_sweep_points = 461
            print(f"📊 Using {actual_sweep_points} sweep points per trace for {band_name} (N9340B detected).")
        else: # Default for N9342CN or unknown
            actual_sweep_points = 500
            print(f"📊 Using {actual_sweep_points} sweep points per trace for {band_name} (N9342CN or Unknown detected).")


        # Calculate the optimal span for each segment to achieve desired RBW per point for *this band*
        # We want (Segment Span / (Actual Points - 1)) = Desired Scan RBW (segmentation)
        # So, Segment Span = Desired Scan RBW * (Actual Points - 1)
        optimal_segment_span_hz = scan_rbw_segmentation * (actual_sweep_points - 1)
        print(f"🎯 Optimal segment span to achieve {scan_rbw_segmentation/1000:.0f} kHz effective RBW per point for {band_name}: {optimal_segment_span_hz / MHZ_TO_HZ:.3f} MHz.")


        # Calculate total number of segments for the current band
        full_band_span_hz = band_stop_freq_hz - band_start_freq_hz
        if full_band_span_hz <= 0:
            total_segments_in_band = 1 # A single point or zero span, still one "segment" to process
        else:
            total_segments_in_band = int(np.ceil(full_band_span_hz / optimal_segment_span_hz))
            # Ensure at least one segment if the band has any span
            if total_segments_in_band == 0:
                total_segments_in_band = 1

        # Now, explicitly set the START and STOP for the *first segment* of this new band
        current_segment_start_freq_hz = band_start_freq_hz # Initialize for the loop
        first_segment_stop_freq_hz = min(current_segment_start_freq_hz + optimal_segment_span_hz, band_stop_freq_hz)
        print(f"🚀 Instrument forced to initial segment range for new band: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {first_segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz.")

        segment_counter = 0

        while current_segment_start_freq_hz < band_stop_freq_hz:
            # Check for pause state at the beginning of each segment
            while app_instance_ref.paused:
                app_instance_ref.after(100, app_instance_ref.update_progress_label, "Scan Paused. Click Resume to continue.")
                time.sleep(0.1) # Sleep briefly while paused
                if not app_instance_ref.scanning: # Allow stopping even when paused
                    print(f"Scan for {band_name} interrupted during pause in segment {segment_counter + 1}.")
                    return all_scan_data, last_successful_band_index # IMMEDIATE RETURN
            
            if not app_instance_ref.scanning: # Check scan flag after pause loop
                print(f"Scan for {band_name} interrupted during segment {segment_counter + 1}.")
                return all_scan_data, last_successful_band_index # IMMEDIATE RETURN

            segment_counter += 1
            segment_stop_freq_hz = min(current_segment_start_freq_hz + optimal_segment_span_hz, band_stop_freq_hz)
            actual_segment_span_hz = segment_stop_freq_hz - current_segment_start_freq_hz

            if actual_segment_span_hz <= 0: # Avoid infinite loop if start == stop or negative span
                break

            # If the last segment is very small, it might result in less than 2 points,
            # which could cause issues with frequency step calculation.
            # We ensure minimum span if points are fixed.
            if actual_sweep_points > 1 and actual_segment_span_hz < (scan_rbw_segmentation * (actual_sweep_points - 1)):
                if segment_stop_freq_hz == band_stop_freq_hz: # This is the very last segment
                    pass # Allow smaller span for last segment, it will have fewer effective points but covers the end
                else:
                    # For intermediate segments, if the calculated span is too small, skip to next full segment
                    current_segment_start_freq_hz += optimal_segment_span_hz
                    continue # Skip this tiny segment, move to next potential full segment

            # Set instrument frequency range for the current segment
            write_safe(inst, f":SENS:FREQ:STAR {current_segment_start_freq_hz};:SENS:FREQ:STOP {segment_stop_freq_hz}")

            write_safe(inst, ":TRAC1:MODE BLANk;:TRAC2:MODE BLANk;:TRAC3:MODE BLANk")
            write_safe(inst, ":TRAC2:MODE MAXHold;")

            # Add a small delay after setting frequencies to allow instrument to configure
            time.sleep(0.1)
            time.sleep(0.5) # Add a small delay for data processing within the instrument

            # Add settling time for max hold values to show up, if max hold is enabled
            if max_hold_time > 0:
                for _ in range(int(max_hold_time * 10)): # Check every 0.1 seconds
                    while app_instance_ref.paused:
                        app_instance_ref.after(100, app_instance_ref.update_progress_label, "Scan Paused. Click Resume to continue.")
                        time.sleep(0.1) # Sleep briefly while paused
                        if not app_instance_ref.scanning: # Allow stopping even when paused
                            print(f"Scan for {band_name} interrupted during pause in max hold for segment {segment_counter + 1}.")
                            return all_scan_data, last_successful_band_index # IMMEDIATE RETURN
                    
                    if not app_instance_ref.scanning: # Check scan flag after pause loop
                        print(f"Scan for {band_name} interrupted during max hold for segment {segment_counter + 1}.")
                        return all_scan_data, last_successful_band_index # IMMEDIATE RETURN

                    # Update display for countdown (only update every second for cleaner output)
                    if _ % 10 == 0: # Every 10 iterations (1 second)
                        sec_remaining = int(max_hold_time - (_ / 10))
                        display_text = f"⏳ {sec_remaining}"
                        app_instance_ref.after(0, app_instance_ref._update_console_line, display_text, False) 
                    time.sleep(0.1) # Small sleep to allow other threads/GUI to run
                # Clear the line before printing the final message
                app_instance_ref.after(0, app_instance_ref._update_console_line, "✅", False) # Overwrite with final checkmark
            
            if not app_instance_ref.scanning: # Check scan flag after max hold loop
                print(f"Scan for {band_name} interrupted after max hold for segment {segment_counter + 1}.")
                return all_scan_data, last_successful_band_index # IMMEDIATE RETURN
            
            # Calculate progress for the emoji bar - Using more compatible ASCII characters
            progress_percentage = (segment_counter / total_segments_in_band)
            bar_length = 20 # Total number of characters in the bar
            filled_length = int(round(bar_length * progress_percentage))
            # Using '█' (U+2588 Full Block) and '-' (Hyphen) for better compatibility
            progressbar = '█' * filled_length + '-' * (bar_length - filled_length)

            # Combined print statement as per user request, now using _update_console_line with overwrite
            progress_message = f"{progressbar}🔍 Span:📊{actual_segment_span_hz/MHZ_TO_HZ:.3f} MHz--📈{current_segment_start_freq_hz/MHZ_TO_HZ:.3f} MHz to 📉{segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz   ✅{segment_counter} of {total_segments_in_band} \n"
            # Add '\r' to ensure the next message overwrites this line correctly
            app_instance_ref.after(0, app_instance_ref._update_console_line, progress_message + '\r', True) 

            # Read and process trace data
            trace_data = []
            try:
                # Conditional trace data query based on instrument model
                if app_instance_ref.instrument_model == "N9340B":
                    trace_data_str = query_safe(inst, ":TRAC2:DATA?")
                else: # Default to N9342CN command or if model is unknown
                    trace_data_str = query_safe(inst, ":TRACe:DATA? TRACe2")

                if trace_data_str is None or "[Not Supported or Timeout]" in trace_data_str or not trace_data_str.strip():
                    print("🚫 No valid trace data string received for this segment.")
                    current_segment_start_freq_hz = segment_stop_freq_hz
                    continue # Move to the next segment if no data

                # Split the string by comma and convert each part to float
                trace_data = [float(val) for val in trace_data_str.split(',')]

                num_trace_points_actual = len(trace_data)

                if num_trace_points_actual > 1:
                    freq_step_per_point_actual = actual_segment_span_hz / (num_trace_points_actual - 1)
                elif num_trace_points_actual == 1:
                    freq_step_per_point_actual = 0 # Single point, no step
                else:
                    freq_step_per_point_actual = 0 # No points
                    print("🚫 No trace data received for this segment after parsing.")
                    current_segment_start_freq_hz = segment_stop_freq_hz
                    continue # Move to the next segment if no data

                # Loop to append data to all_scan_data
                for j, amp_value in enumerate(trace_data):
                    current_freq_for_point_hz = current_segment_start_freq_hz + (j * freq_step_per_point_actual)

                    # Append to list for plotting later (using Hz for consistency with plot_data expectation)
                    all_scan_data.append((current_freq_for_point_hz, amp_value))
                
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

            current_segment_start_freq_hz = segment_stop_freq_hz # Move to the start of the next segment

    print("\n--- 🎉 Band Scan Complete! ---")
    return all_scan_data, last_successful_band_index # Return the collected data and last successful index
