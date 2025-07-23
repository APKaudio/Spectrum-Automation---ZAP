# Updated ScanV10.4.2.py with new plot options, auto-open functionality, and enhanced scan_bands

import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, TclError
from tkinter import ttk  # Import ttk for Treeview - FIX: Added this import
import pyvisa
import time
import argparse
import struct
import numpy as np
import os
import csv
from datetime import datetime
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import sys
import subprocess
import threading
import re
import webbrowser # Added for opening HTML plots

# Import the frequency band definitions from the new file
# Assuming frequency_bands.py exists in the same directory
try:
    from frequency_bands import (
        MHZ_TO_HZ,
        SCAN_BAND_RANGES,
        TV_PLOT_BAND_MARKERS,
        GOV_PLOT_BAND_MARKERS
    )
except ImportError:
    print("Error: frequency_bands.py not found. Please ensure it's in the same directory.")
    # Define dummy values to prevent errors if file is missing
    MHZ_TO_HZ = 1_000_000
    SCAN_BAND_RANGES = [
        {"Band Name": "Dummy Band 1", "Start MHz": 100, "Stop MHz": 200},
        {"Band Name": "Dummy Band 2", "Start MHz": 400, "Stop MHz": 500}
    ]
    TV_PLOT_BAND_MARKERS = []
    GOV_PLOT_BAND_MARKERS = []


# Updated wait time variable and its usage for the continuous loop
DEFAULT_RBW_STEP_SIZE_HZ = 1000000 # Corrected default RBW to 10 kHz as per typical usage
DEFAULT_CYCLE_WAIT_TIME_SECONDS = 0 # 30 seconds wait between full scan cycles
DEFAULT_MAXHOLD_TIME_SECONDS = 3 # Default max hold time for the new argument
DEFAULT_SCAN_RBW_HZ = 10000 # Changed default to 100 kHz as per user request for "Scan RBW" for segmentation
DEFAULT_REFERENCE_LEVEL_DBM = -40 # Default reference level as requested

# Global variable for debug mode, controlled by the GUI
_debug_mode_enabled = False

# --- Utility Functions --- 

def check_and_install_dependencies():
    """Checks for necessary libraries and installs them if missing."""
    required_packages = ['pyvisa', 'pandas', 'plotly', 'numpy']
    missing_packages = []
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)

    if missing_packages:
        response = messagebox.askyesno(
            "Missing Dependencies",
            f"The following packages are missing: {', '.join(missing_packages)}.\n"
            "Do you want to install them now? This requires an internet connection."
        )
        if response:
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", *missing_packages])
                messagebox.showinfo("Installation Complete", "Dependencies installed successfully. Please restart the application.")
                return False # Indicate that app should not proceed directly
            except Exception as e:
                messagebox.showerror("Installation Failed", f"Failed to install dependencies: {e}")
                return False
        else:
            messagebox.showwarning("Dependencies Missing", "Application cannot run without required dependencies.")
            return False
    return True # All dependencies are met

def query_safe(inst, command):
    """Safely queries the instrument, handling VISA errors."""
    try:
        response = inst.query(command).strip()
        global _debug_mode_enabled # Access the global variable
        if _debug_mode_enabled: # Conditional print
            print(f"Query: '{command}' -> Response: '{response}'")
        return response
    except pyvisa.errors.VisaIOError as e:
        print(f"VISA Error during query '{command}': {e}")
        return None
    except Exception as e:
        print(f"Error parsing response for '{command}': {e}")
        return None

def write_safe(inst, command):
    """Safely writes to the instrument, handling VISA errors."""
    try:
        inst.write(command)
        global _debug_mode_enabled # Access the global variable
        if _debug_mode_enabled: # Conditional print
            print(f"Write: '{command}'")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"VISA Error during write '{command}': {e}")
        return False

# --- Instrument Control Functions ---

def initialize_instrument(inst, ref_level_dbm, preamp_on, display_log_scale, rbw_config_val, vbw_config_val, max_hold_on):
    """Initializes the Keysight N9340B spectrum analyzer with basic settings."""
    print("✨ Initializing instrument with desired settings...")
    try:
        write_safe(inst, "SYSTem:DISPlay:UPDate ON") # Ensure display updates
    

        # Set reference level
        write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM")
        print(f"✅ Set reference level to {ref_level_dbm} dBm.")

        # Set preamplifier
        if preamp_on:
            write_safe(inst, ":SENSe:POWer:RF:GAIN:STATe ON") # Corrected SCPI command for preamp
            print("✅ Preamplifier ON.")
        else:
            write_safe(inst, ":SENSe:POWer:RF:GAIN:STATe OFF") # Corrected SCPI command for preamp
            print("✅ Preamplifier OFF.")


        write_safe(inst, ":TRAC1:MODE WRITe")
        print(f"✅ Trace 1 sent to write")


        write_safe(inst, ":TRAC2:MODE MAXHold")
        print(f"✅ Trace 2 sent to MAX HOLD")



        

        # Set display scale
        if display_log_scale:
            write_safe(inst, ":DISPlay:WINDow:TRACe:Y:SCALe:SPACing LOGarithmic")
            print("✅ Display scale set to LOG.")
        else:
            write_safe(inst, ":DISPlay:WINDow:TRACe:Y:SCALe:SPACing LINear")
            print("✅ Display scale set to LIN.")
        

        # Set RBW and VBW (These are now passed in, but the UI fields are removed)
        write_safe(inst, ":SENSe:BANDwidth:RESolution:AUTO OFF")
        write_safe(inst, f":SENSe:BANDwidth:RESolution {rbw_config_val}")
        

        write_safe(inst, ":SENSe:BANDwidth:VIDeo:AUTO OFF")
        write_safe(inst, f":SENSe:BANDwidth:VIDeo {vbw_config_val}")
        
        print(f"✅ Set RBW to {rbw_config_val} Hz, VBW to {vbw_config_val} Hz.")
        
        # Set sweep time to auto
        write_safe(inst, ":SENSe:SWEep:TIME:AUTO ON")
        
        print("✅ Sweep time set to AUTO.")

        # Configure trace 1 to clear write or max hold based on user selection
        if max_hold_on:
            write_safe(inst, ":TRAC2:MODE MAXHold")
            print("⏸️ Trace 1 set to max hold.")
        else:
            write_safe(inst, ":TRAC1:MODE WRITe") # Or NORMal, depending on desired default
            print("▶️ Trace 1 set to normal/write mode.")
        
        # Explicitly set data format to ASCII for :TRACe:DATA? query
        write_safe(inst, ":FORMat:DATA ASCii") 
        print("✅ Set trace data format to ASCII for data transfer.")

      
        print("🎉 Instrument initialized successfully with desired settings.")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"🛑 Failed to initialize instrument with desired settings: {e}")
        return False

def scan_bands(app_instance, csv_writer, selected_bands, scan_rbw_segmentation, rbw_config_val, vbw_config_val, max_hold_time, last_scanned_band_index=0):
    """
    Iterates through predefined frequency bands, sets the start/stop frequencies,
    reduces RBW to 10000 Hz, and triggers a sweep for each band.
    It extracts trace data, writes it directly to the provided CSV writer,
    and also returns it for further processing (plotting).
    This function now dynamically segments bands to maintain a consistent
    effective resolution bandwidth per trace point.
    It also displays the time of day for each band scanned.
    
    Added last_scanned_band_index to resume from where it left off.

    Args:
        app_instance (App): The main Tkinter App instance for GUI updates.
        csv_writer (csv.writer): The CSV writer object to write data to.
        selected_bands (list): A list of band dictionaries to scan.
        scan_rbw_segmentation (float): Resolution Bandwidth for segmenting bands (from new GUI input).
        rbw_config_val (float): Resolution Bandwidth value to configure on the instrument.
        vbw_config_val (float): Video Bandwidth value to configure on the instrument.
        max_hold_time (float): Duration in seconds for which MAX Hold should be active.
                                If > 0, MAX Hold mode is enabled for the scan.
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
    write_safe(app_instance.inst, f":SENSE:BAND:RES {scan_rbw_segmentation}")
    print(f"📏 Set RBW to {scan_rbw_segmentation/1000:.0f} kHz for scan (from Scan RBW setting).")
    write_safe(app_instance.inst, f":SENSE:BAND:VID {vbw_config_val}")
    print(f"📺 Set VBW to {vbw_config_val} Hz for scan.")


    # *** Use selected_bands for scanning the instrument ***
    # Iterate through bands starting from last_scanned_band_index
    for i in range(last_scanned_band_index, len(selected_bands)):
        band = selected_bands[i]
        band_name = band["Band Name"]
        band_start_freq_hz = band["Start MHz"] * MHZ_TO_HZ
        band_stop_freq_hz = band["Stop MHz"] * MHZ_TO_HZ

        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"\n📈 [{current_time}] Processing Band: {band_name} (Total Range: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {band_stop_freq_hz/MHZ_TO_HZ:.3f} MHz)")

        # Determine actual_sweep_points based on instrument model
        if app_instance.instrument_model == "N9340B":
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
            # This line had a typo: `total_segments_in_segments` should be `total_segments_in_band`
            if total_segments_in_band == 0: # Corrected variable name here
                total_segments_in_band = 1

        # Now, explicitly set the START and STOP for the *first segment* of this new band
        # This will be the first segment's actual range.
        current_segment_start_freq_hz = band_start_freq_hz # Initialize for the loop
        first_segment_stop_freq_hz = min(current_segment_start_freq_hz + optimal_segment_span_hz, band_stop_freq_hz)
        #write_safe(app_instance.inst, f":SENS:FREQ:STAR {current_segment_start_freq_hz}")
        #write_safe(app_instance.inst, f":SENS:FREQ:STOP {first_segment_stop_freq_hz}")
           #query_safe(app_instance.inst, "*OPC?") # Wait for operation to complete
        #app_instance.inst.clear() # Flush buffer after OPC query
        #time.sleep(0.2) # Small delay for the instrument to settle
        print(f"🚀 Instrument forced to initial segment range for new band: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {first_segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz.")
        # --- END RE-ADDED BLOCK ---

        segment_counter = 0

        while current_segment_start_freq_hz < band_stop_freq_hz:
            # Check for pause state at the beginning of each segment
            while app_instance.paused:
                app_instance.after(100, app_instance.update_progress_label, "Scan Paused. Click Resume to continue.")
                time.sleep(0.1) # Sleep briefly while paused
                if not app_instance.scanning: # Allow stopping even when paused
                    print(f"Scan for {band_name} interrupted during pause in segment {segment_counter + 1}.")
                    break # Exit pause loop if scan is stopped

            if not app_instance.scanning: # Check scan flag after pause loop
                print(f"Scan for {band_name} interrupted during segment {segment_counter + 1}.")
                break

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
            write_safe(app_instance.inst, f":SENS:FREQ:STAR {current_segment_start_freq_hz};:SENS:FREQ:STOP {segment_stop_freq_hz}")

            #write_safe(app_instance.inst, f":SENS:FREQ:STOP {segment_stop_freq_hz}")

            write_safe(app_instance.inst, ":TRAC1:MODE BLANk;:TRAC2:MODE BLANk;:TRAC3:MODE BLANk")
            
            
            write_safe(app_instance.inst, ":TRAC2:MODE MAXHold;:TRAC3:MODE MINHold")


            
            # Add a small delay after setting frequencies to allow instrument to configure
            time.sleep(0.1)

            #query_safe(app_instance.inst, "*OPC?") # Wait for the sweep to completed
            #app_instance.inst.clear() # Flush buffer after OPC query
            time.sleep(0.5) # Add a small delay for data processing within the instrument

            # Add settling time for max hold values to show up, if max hold is enabled
            if max_hold_time > 0:
                # FIX: Replaced single time.sleep with a loop to allow pause checks
                for _ in range(int(max_hold_time * 10)): # Check every 0.1 seconds
                    while app_instance.paused:
                        app_instance.after(100, app_instance.update_progress_label, "Scan Paused. Click Resume to continue.")
                        time.sleep(0.1) # Sleep briefly while paused
                        if not app_instance.scanning: # Allow stopping even when paused
                            print(f"Scan for {band_name} interrupted during pause in max hold for segment {segment_counter + 1}.")
                            break # Exit pause loop if scan is stopped
                    
                    if not app_instance.scanning: # Check scan flag after pause loop
                        print(f"Scan for {band_name} interrupted during max hold for segment {segment_counter + 1}.")
                        break # Exit max hold loop if scan is stopped

                    # Update display for countdown (only update every second for cleaner output)
                    if _ % 10 == 0: # Every 10 iterations (1 second)
                        sec_remaining = int(max_hold_time - (_ / 10))
                        display_text = f"⏳ {sec_remaining}"
                        app_instance.after(0, app_instance._update_console_line, display_text, False) 
                    time.sleep(0.1) # Small sleep to allow other threads/GUI to run
                # Clear the line before printing the final message
                app_instance.after(0, app_instance._update_console_line, "✅", False) # Overwrite with final checkmark
            
            if not app_instance.scanning: # Check scan flag after max hold loop
                print(f"Scan for {band_name} interrupted after max hold for segment {segment_counter + 1}.")
                break

         #   query_safe(app_instance.inst, "*OPC?") # Wait for the sweep to complete
            #app_instance.inst.clear() # Flush buffer after OPC query
            
            # Calculate progress for the emoji bar - Using more compatible ASCII characters
            progress_percentage = (segment_counter / total_segments_in_band)
            bar_length = 20 # Total number of characters in the bar
            filled_length = int(round(bar_length * progress_percentage))
            # Using '█' (U+2588 Full Block) and '-' (Hyphen) for better compatibility
            progressbar = '█' * filled_length + '-' * (bar_length - filled_length)

            # Combined print statement as per user request, now using _update_console_line with overwrite
            progress_message = f"{progressbar}🔍 Span:📊{actual_segment_span_hz/MHZ_TO_HZ:.3f} MHz--📈{current_segment_start_freq_hz/MHZ_TO_HZ:.3f} MHz to 📉{segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz   ✅{segment_counter} of {total_segments_in_band} \n"
            # Add '\r' to ensure the next message overwrites this line correctly
            app_instance.after(0, app_instance._update_console_line, progress_message + '\r', True) 

            # Read and process trace data
            trace_data = []
            try:
                # Conditional trace data query based on instrument model
                if app_instance.instrument_model == "N9340B":
                    trace_data_str = query_safe(app_instance.inst, ":TRAC1:DATA?")
                else: # Default to N9342CN command or if model is unknown
                    trace_data_str = query_safe(app_instance.inst, ":TRACe:DATA? TRACe2")

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

                # Loop to append data to all_scan_data and write to CSV
                for j, amp_value in enumerate(trace_data):
                    current_freq_for_point_hz = current_segment_start_freq_hz + (j * freq_step_per_point_actual)

                    # Append to list for plotting later (using Hz for consistency with plot_data expectation)
                    all_scan_data.append((current_freq_for_point_hz, amp_value))

                    # Write directly to CSV file with desired order and units
                    csv_writer.writerow([
                        f"{current_freq_for_point_hz / MHZ_TO_HZ:.2f}",  # Frequency in MHz for CSV
                        f"{amp_value:.2f}"
                    ])
                
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

def plot_single_scan_data(scanned_data, plot_title_suffix, include_tv_markers=True, include_gov_markers=True):
    """
    Generates an interactive Plotly HTML plot from scanned data,
    including overlays for TV and Government frequency bands based on flags.
    """
    if not scanned_data:
        print("No data to plot.")
        return None, None # Return None for both fig and filename

    # Ensure scanned_data is in the expected (Frequency_Hz, Power_dBm) tuple format
    df = pd.DataFrame(scanned_data, columns=['Frequency_Hz', 'Power_dBm'])
    df['Frequency_MHz'] = df['Frequency_Hz'] / MHZ_TO_HZ

    fig = px.line(df, x='Frequency_MHz', y='Power_dBm', 
                  title=f'RF Spectrum Scan ({plot_title_suffix})',
                  labels={'Frequency_MHz': 'Frequency (MHz)', 'Power_dBm': 'Power (dBm)'},
                  line_shape='linear') # 'linear' connects points with straight lines

    # Determine Y-axis range for marker positioning
    y_range_min = df['Power_dBm'].min()
    # Set max Y-axis to 0 dBm as requested
    y_range_max = 0 # Fixed maximum to 0 dBm
    # Add some padding to the y-range for better text visibility (only to min, max is fixed)
    y_padding = (y_range_max - y_range_min) * 0.1
    y_range_min -= y_padding
    
    fig.update_layout(yaxis_range=[y_range_min, y_range_max])


    # Determine X-axis range for marker visibility check
    x_min_data = df['Frequency_MHz'].min()
    x_max_data = df['Frequency_MHz'].max()

    # --- Add TV Band Markers ---
    if include_tv_markers:
        # Define colors for the TV band markers and text
        tv_marker_line_color = "rgba(255, 255, 0, 0.7)"  # Bright yellow, semi-transparent
        tv_marker_text_color = "yellow"
        tv_band_fill_color = "rgba(255, 255, 0, 0.05)"    # Very light yellow, highly transparent fill

        for band in TV_PLOT_BAND_MARKERS:
            # Only add markers if they are within the actual scanned frequency range
            if band["Start MHz"] < x_max_data and band["Stop MHz"] > x_min_data:
                # Add a shaded rectangle to represent the frequency band allocation
                fig.add_shape(
                    type="rect",
                    x0=band["Start MHz"],
                    y0=y_range_min,  # Span full Y-axis range
                    x1=band["Stop MHz"],
                    y1=y_range_max,  # Span full Y-axis range
                    line=dict(
                        color=tv_marker_line_color,
                        width=0.3,
                        dash="dot",
                    ),
                    fillcolor=tv_band_fill_color,
                    layer="below",
                )

                # Add text markers using go.Scatter with mode='text'
                x_center = (band["Start MHz"] + band["Stop MHz"]) / 2
                y_text_position = y_range_max - (y_range_max - y_range_min) * 0.05

                fig.add_trace(go.Scatter(
                    x=[x_center],
                    y=[y_text_position],
                    mode='text',
                    text=[f"{band['Band Name']}<br>{band['Start MHz']:.1f}-{band['Stop MHz']:.1f} MHz"],
                    textfont=dict(
                        size=8,
                        color=tv_marker_text_color
                    ),
                    showlegend=False,
                    hoverinfo='text',
                    name=f"Band Label: {band['Band Name']}"
                ))

    # --- Add Government Band Markers ---
    if include_gov_markers:
        # Define colors for the Government band markers and text
        gov_marker_line_color = "rgba(255, 0, 0, 0.9)"  # Red, semi-transparent
        gov_marker_text_color = "red"
        gov_band_fill_color = "rgba(255, 0, 0, 0.1)"    # Very light red, highly transparent fill

        # Define the four y-offsets for staggering
        y_offset_level_1 = 0.20
        y_offset_level_2 = 0.25
        y_offset_level_3 = 0.30
        y_offset_level_4 = 0.35
        y_offset_levels = [y_offset_level_1, y_offset_level_2, y_offset_level_3, y_offset_level_4]

        for i, band in enumerate(GOV_PLOT_BAND_MARKERS):
            # Only add markers if they are within the actual scanned frequency range
            if band["Start MHz"] < x_max_data and band["Stop MHz"] > x_min_data:
                # Add a shaded rectangle to represent the frequency band allocation
                fig.add_shape(
                    type="rect",
                    x0=band["Start MHz"],
                    y0=y_range_min,
                    x1=band["Stop MHz"],
                    y1=y_range_max,
                    line=dict(
                        color=gov_marker_line_color,
                        width=0.3,
                        dash="dot",
                    ),
                    fillcolor=gov_band_fill_color,
                    layer="below",
                )

                # Add text markers using go.Scatter with mode='text'
                x_center = (band["Start MHz"] + band["Stop MHz"]) / 2

                # Determine the Y position based on staggering using modulo for 4 levels
                current_y_offset = y_offset_levels[i % len(y_offset_levels)]
                y_text_position = y_range_max - (y_range_max - y_range_min) * current_y_offset

                fig.add_trace(go.Scatter(
                    x=[x_center],
                    y=[y_text_position],
                    mode='text',
                    text=[f"{band['Band Name']}<br>{band['Start MHz']:.1f}-{band['Stop MHz']:.1f} MHz"],
                    textfont=dict(
                        size=8,
                        color=gov_marker_text_color
                    ),
                    showlegend=False,
                    hoverinfo='text',
                    name=f"Band Label: {band['Band Name']}"
                ))

    # Apply Dark Mode Theme
    fig.update_layout(template="plotly_dark")

    fig.update_layout(hovermode="x unified") # Show all traces on hover

    plot_filename = f"spectrum_scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    # This function will now return the path to be saved by the calling function
    return fig, plot_filename # Return figure and suggested filename

def plot_multi_trace_data(aggregated_df, plot_title_suffix, include_tv_markers=True, include_gov_markers=True, historical_dfs_with_names=None, output_html_path=None):
    """
    Generates an interactive Plotly HTML plot from aggregated scan data (Average, Median, Range),
    including overlays for TV and Government frequency bands based on flags.
    Optionally includes historical scan data as additional layers.

    Args:
        aggregated_df (pd.DataFrame): DataFrame with 'Frequency_Hz', 'Average_dBm', 'Median_dBm', 'Range_dBm'.
        plot_title_suffix (str): Suffix for the plot title.
        include_tv_markers (bool): Whether to include TV band markers.
        include_gov_markers (bool): Whether to include Government band markers.
        historical_dfs_with_names (list of tuple): Optional. List of (DataFrame, name) tuples for historical scans.
        output_html_path (str): The full path including filename for the HTML output.
    Returns:
        tuple: (plotly.graph_objects.Figure, str) The Plotly figure object and the output HTML path.
    """
    if aggregated_df.empty and not historical_dfs_with_names:
        print("No data to plot (neither aggregated nor historical).")
        return None, None

    fig = go.Figure()

    # Add Average trace
    if not aggregated_df.empty:
        # The 'Frequency_MHz' column should already be present in aggregated_df
        # from the generate_average_plot function.
        # This check is a safeguard but should ideally not be needed if data preparation is correct.
        if 'Frequency_MHz' not in aggregated_df.columns:
            # If it's truly missing, and Frequency_Hz is present, derive it.
            # This indicates an issue in data preparation before calling this function.
            if 'Frequency_Hz' in aggregated_df.columns:
                aggregated_df['Frequency_MHz'] = aggregated_df['Frequency_Hz'] / MHZ_TO_HZ
            else:
                print("Error: Neither 'Frequency_MHz' nor 'Frequency_Hz' found in aggregated_df.")
                return None, None # Cannot plot without frequency data

        fig.add_trace(go.Scatter(x=aggregated_df['Frequency_MHz'], y=aggregated_df['Average_dBm'],
                                 mode='lines', name='Average Power (dBm)',
                                 line=dict(color='cyan', width=3))) # Make average line thicker

        # Add Median trace (now solid)
        fig.add_trace(go.Scatter(x=aggregated_df['Frequency_MHz'], y=aggregated_df['Median_dBm'],
                                 mode='lines', name='Median Power (dBm)',
                                 line=dict(color='lightgreen', width=2))) # Removed dash='dot'

        # Add Range trace (Max-Min difference) (now solid)
        fig.add_trace(go.Scatter(x=aggregated_df['Frequency_MHz'], y=aggregated_df['Range_dBm'],
                                 mode='lines', name='Range (Max - Min) (dB)',
                                 line=dict(color='yellow', width=2))) # Removed dash='dash'

    # Add historical data as additional layers
    if historical_dfs_with_names:
        for hist_df, hist_name in historical_dfs_with_names:
            # FIX: Ensure we are operating on a copy to avoid SettingWithCopyWarning
            hist_df_copy = hist_df.copy() 
            # The 'Frequency_MHz' column should already be present in hist_df_copy
            if 'Frequency_MHz' not in hist_df_copy.columns:
                # If it's truly missing, and Frequency_Hz is present, derive it.
                if 'Frequency_Hz' in hist_df_copy.columns:
                    hist_df_copy['Frequency_MHz'] = hist_df_copy['Frequency_Hz'] / MHZ_TO_HZ
                else:
                    print(f"Error: Neither 'Frequency_MHz' nor 'Frequency_Hz' found in historical_df for {hist_name}. Skipping.")
                    continue # Skip this historical DataFrame if frequency data is missing

            fig.add_trace(go.Scatter(x=hist_df_copy['Frequency_MHz'], y=hist_df_copy['Power_dBm'],
                                     mode='lines', name=f"Scan: {hist_name}",
                                     line=dict(color='rgba(100, 100, 255, 0.5)', width=1, dash='dot'), # Lighter, thinner, dashed
                                     showlegend=True)) # Ensure legend entry for each historical scan


    fig.update_layout(title=f'RF Spectrum Scan - {plot_title_suffix}',
                      xaxis_title='Frequency (MHz)',
                      yaxis_title='Power / Range (dBm)', # Y-axis label accommodates range and power
                      template="plotly_dark",
                      hovermode="x unified")

    # Determine Y-axis range for marker positioning, considering all plotted traces
    all_y_values = []
    if not aggregated_df.empty:
        all_y_values.extend(aggregated_df['Average_dBm'].tolist())
        all_y_values.extend(aggregated_df['Median_dBm'].tolist())
        all_y_values.extend(aggregated_df['Range_dBm'].tolist()) # Include Range_dBm in max calculation
    if historical_dfs_with_names:
        for hist_df, _ in historical_dfs_with_names:
            # Ensure 'Power_dBm' exists before extending
            if 'Power_dBm' in hist_df.columns:
                all_y_values.extend(hist_df['Power_dBm'].tolist())

    if all_y_values:
        # Calculate y_range_max as the maximum of all relevant y-values, ensuring it's at least 0
        y_range_max = max(0, max(all_y_values))
    else:
        y_range_max = 0 # Default if no data

    y_range_min = min(all_y_values) if all_y_values else -100 # Default if no data
    y_padding = (y_range_max - y_range_min) * 0.1
    y_range_min -= y_padding
    
    fig.update_layout(yaxis_range=[y_range_min, y_range_max]) # Apply the updated Y-axis range


    # Determine X-axis range for marker visibility check
    x_min_data = float('inf')
    x_max_data = float('-inf')

    if not aggregated_df.empty:
        if 'Frequency_MHz' in aggregated_df.columns:
            x_min_data = min(x_min_data, aggregated_df['Frequency_MHz'].min())
            x_max_data = max(x_max_data, aggregated_df['Frequency_MHz'].max())
    if historical_dfs_with_names:
        for hist_df, _ in historical_dfs_with_names:
            if 'Frequency_MHz' in hist_df.columns:
                x_min_data = min(x_min_data, hist_df['Frequency_MHz'].min())
                x_max_data = max(x_max_data, hist_df['Frequency_MHz'].max())

    # --- Add TV Band Markers ---
    if include_tv_markers:
        # Define colors for the TV band markers and text
        tv_marker_line_color = "rgba(255, 255, 0, 0.7)"  # Bright yellow, semi-transparent
        tv_marker_text_color = "yellow"
        tv_band_fill_color = "rgba(255, 255, 0, 0.05)"    # Very light yellow, highly transparent fill

        for band in TV_PLOT_BAND_MARKERS:
            # Only add markers if they are within the actual scanned frequency range
            if band["Start MHz"] < x_max_data and band["Stop MHz"] > x_min_data:
                # Add a shaded rectangle to represent the frequency band allocation
                fig.add_shape(
                    type="rect",
                    x0=band["Start MHz"],
                    y0=y_range_min,  # Span full Y-axis range
                    x1=band["Stop MHz"],
                    y1=y_range_max,  # Span full Y-axis range
                    line=dict(
                        color=tv_marker_line_color,
                        width=0.3,
                        dash="dot",
                    ),
                    fillcolor=tv_band_fill_color,
                    layer="below",
                )

                # Add text markers using go.Scatter with mode='text'
                x_center = (band["Start MHz"] + band["Stop MHz"]) / 2
                y_text_position = y_range_max - (y_range_max - y_range_min) * 0.05

                fig.add_trace(go.Scatter(
                    x=[x_center],
                    y=[y_text_position],
                    mode='text',
                    text=[f"{band['Band Name']}<br>{band['Start MHz']:.1f}-{band['Stop MHz']:.1f} MHz"],
                    textfont=dict(
                        size=8,
                        color=tv_marker_text_color
                    ),
                    showlegend=False,
                    hoverinfo='text',
                    name=f"Band Label: {band['Band Name']}"
                ))

    # --- Add Government Band Markers ---
    if include_gov_markers:
        # Define colors for the Government band markers and text
        gov_marker_line_color = "rgba(255, 0, 0, 0.9)"  # Red, semi-transparent
        gov_marker_text_color = "red"
        gov_band_fill_color = "rgba(255, 0, 0, 0.1)"    # Very light red, highly transparent fill

        # Define the four y-offsets for staggering
        y_offset_level_1 = 0.20
        y_offset_level_2 = 0.25
        y_offset_level_3 = 0.30
        y_offset_level_4 = 0.35
        y_offset_levels = [y_offset_level_1, y_offset_level_2, y_offset_level_3, y_offset_level_4]

        for i, band in enumerate(GOV_PLOT_BAND_MARKERS):
            # Only add markers if they are within the actual scanned frequency range
            if band["Start MHz"] < x_max_data and band["Stop MHz"] > x_min_data:
                # Add a shaded rectangle to represent the frequency band allocation
                fig.add_shape(
                    type="rect",
                    x0=band["Start MHz"],
                    y0=y_range_min,
                    x1=band["Stop MHz"],
                    y1=y_range_max,
                    line=dict(
                        color=gov_marker_line_color,
                        width=0.3,
                        dash="dot",
                    ),
                    fillcolor=gov_band_fill_color,
                    layer="below",
                )

                # Add text markers using go.Scatter with mode='text'
                x_center = (band["Start MHz"] + band["Stop MHz"]) / 2

                # Determine the Y position based on staggering using modulo for 4 levels
                current_y_offset = y_offset_levels[i % len(y_offset_levels)]
                y_text_position = y_range_max - (y_range_max - y_range_min) * current_y_offset

                fig.add_trace(go.Scatter(
                    x=[x_center],
                    y=[y_text_position],
                    mode='text',
                    text=[f"{band['Band Name']}<br>{band['Start MHz']:.1f}-{band['Stop MHz']:.1f} MHz"],
                    textfont=dict(
                        size=8,
                        color=gov_marker_text_color
                    ),
                    showlegend=False,
                    hoverinfo='text',
                    name=f"Band Label: {band['Band Name']}"
                ))

    # Apply Dark Mode Theme
    fig.update_layout(template="plotly_dark")

    fig.update_layout(hovermode="x unified") # Show all traces on hover

    plot_filename = f"spectrum_scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    # This function will now return the path to be saved by the calling function
    return fig, plot_filename # Return figure and suggested filename


# --- GUI Classes ---

class TextRedirector(object):
    """A class to redirect stdout/stderr to a Tkinter scrolled text widget."""
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag
        self.last_char_was_cr = False

    def write(self, str_val):
        self.widget.config(state=tk.NORMAL) # Enable editing
        
        # Handle carriage returns for in-place updates (like progress bars)
        if '\r' in str_val:
            parts = str_val.split('\r')
            for i, part in enumerate(parts):
                if self.last_char_was_cr and i == 0:
                    # If last char was CR, delete current line and then write
                    self.widget.delete("end-1c linestart", "end-1c")
                self.widget.insert(tk.END, part, (self.tag,))
                self.widget.see(tk.END)
                if i < len(parts) - 1: # If not the last part, it means there was a \r
                    self.last_char_was_cr = True
                else:
                    self.last_char_was_cr = False
        else:
            self.widget.insert(tk.END, str_val, (self.tag,))
            self.widget.see(tk.END)
            self.last_char_was_cr = False

        self.widget.config(state=tk.DISABLED) # Disable editing
        self.widget.update_idletasks() # Force update

    def flush(self):
        # This is typically called after write; ensures content is displayed
        # For a ScrolledText widget with update_idletasks, this might not need explicit action.
        pass

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        # FIX: Set root window background to black
        self.configure(bg="black") 
        # Initial title, will be updated upon connection
        self.title(f"RF Spectrum Analyzer Controller - {os.path.basename(sys.argv[0])}")
        self.geometry("1400x780") # Increased width to accommodate new list

        self.rm = pyvisa.ResourceManager()
        self.instrument_list = []
        self.inst = None
        self.scanning = False # Flag to control scanning thread
        self.paused = False # Flag to control pausing the scan
        self.last_scan_data = None # Store the last scan data for plotting
        self.last_csv_file_path = None # Store path of last saved CSV
        self.current_csv_file = None # To hold the open CSV file object
        self.collected_scans_dataframes = [] # To store pandas DataFrames of each completed scan
        self.instrument_model = None # Initialize instrument model to None

        # Dictionary to hold Entry widgets for coloring
        self.desired_setting_entries = {}

        # Initialize Tkinter variables for Scan Configuration (User Input, can be pushed)
        self.desired_ref_level_var = tk.StringVar(self, value=str(DEFAULT_REFERENCE_LEVEL_DBM)) # Set default reference level
        self.desired_preamp_var = tk.BooleanVar(self, value=True)
        self.desired_log_scale_var = tk.BooleanVar(self, value=True)
        self.desired_max_hold_var = tk.BooleanVar(self, value=True) # Changed to True by default
        self.desired_max_hold_time_var = tk.StringVar(self, value=str(DEFAULT_MAXHOLD_TIME_SECONDS))
        # RBW and VBW variables are still needed for instrument configuration, but UI fields are removed
        self.desired_rbw_var = tk.StringVar(self, value=str(DEFAULT_RBW_STEP_SIZE_HZ))
        self.desired_vbw_display_var = tk.StringVar(self, value=str(int(float(DEFAULT_RBW_STEP_SIZE_HZ) / 3))) 
        self.desired_cycle_wait_time_var = tk.StringVar(self, value=str(DEFAULT_CYCLE_WAIT_TIME_SECONDS))
        self.output_folder_var = tk.StringVar(self, value="scan_data") # This is now the base output directory
        self.scan_name_var = tk.StringVar(self, value="MyScan") # New variable for Scan Name
        self.resource_var = tk.StringVar(self) # For VISA resource dropdown

        # Variables for plot markers and auto-open (NEW)
        self.include_gov_markers_var = tk.BooleanVar(self, value=True)
        self.include_tv_markers_var = tk.BooleanVar(self, value=True)
        self.open_html_after_complete_var = tk.BooleanVar(self, value=True) # New variable for "Open HTML after complete"

        # New variable for Scan RBW (for segmentation, distinct from instrument RBW)
        self.desired_scan_rbw_segmentation_var = tk.StringVar(self, value=str(DEFAULT_SCAN_RBW_HZ))

        # Debug mode variable
        self.debug_mode_var = tk.BooleanVar(self, value=False)
        self.debug_mode_var.trace_add("write", self._update_debug_mode_global)


        # Create two main frames: one for the GUI, one for the console output
        self.main_frame = tk.Frame(self, bg="black") # FIX: Set background
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.console_frame = tk.Frame(self, width=700, bg="black")
        self.console_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.console_frame.pack_propagate(False)

        self.console_control_frame = tk.Frame(self.console_frame, bg="black")
        self.console_control_frame.pack(side=tk.TOP, fill=tk.X, pady=5)

        # Using lambda for button commands to ensure correct binding
        self.start_scan_button = tk.Button(self.console_control_frame, text="Start Scan", command=lambda: self.start_scan_thread(), state=tk.DISABLED, bg="green", fg="white", height=2)
        self.start_scan_button.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill=tk.X)

        # New Pause/Resume Button
        self.pause_resume_button = tk.Button(self.console_control_frame, text="Pause Scan", command=self.toggle_pause_scan, state=tk.DISABLED, bg="orange", fg="white", height=2)
        self.pause_resume_button.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill=tk.X)


        self.stop_scan_button = tk.Button(self.console_control_frame, text="Stop Scan", command=lambda: self.stop_scan(), state=tk.DISABLED, bg="red", fg="white", height=2)
        self.stop_scan_button.pack(side=tk.RIGHT, padx=5, pady=5, expand=True, fill=tk.X)

        self.console_output = scrolledtext.ScrolledText(self.console_frame, wrap=tk.WORD, bg="black", fg="white", font=("Consolas", 10))
        self.console_output.pack(expand=True, fill=tk.BOTH)
        self.console_output.configure(state="disabled")

        sys.stdout = TextRedirector(self.console_output, "stdout")
        sys.stderr = TextRedirector(self.console_output, "stderr")

        print("--- RF Spectrum Scanner GUI Initialized ---")
        self.create_widgets()
        # Defer populate_resources call to prevent AttributeError during __init__
        self.after(0, self.populate_resources)
        # self.actual_sweep_points = 401 # No longer a fixed constant, determined dynamically

    def _update_debug_mode_global(self, *args):
        """Updates the global debug mode variable when the checkbox state changes."""
        global _debug_mode_enabled
        _debug_mode_enabled = self.debug_mode_var.get()
        print(f"Debug Mode: {'Enabled' if _debug_mode_enabled else 'Disabled'}")

    def create_widgets(self):
        # FIX: Configure ttk style for dark mode
        style = ttk.Style(self)
        style.theme_use('clam') # Use a theme that allows customization

        # General Treeview styling
        style.configure("Treeview", 
                        background="grey", 
                        foreground="white", 
                        fieldbackground="grey",
                        bordercolor="black",
                        lightcolor="grey",
                        darkcolor="grey")
        style.map("Treeview", 
                  background=[("selected", "blue")], 
                  foreground=[("selected", "white")])
        
        # Treeview Heading styling
        style.configure("Treeview.Heading", 
                        background="darkgrey", 
                        foreground="white",
                        font=('TkDefaultFont', 10, 'bold')) # Make headings bold for visibility

        # Scrollbar styling (for Treeview and Canvas)
        style.configure("Vertical.TScrollbar", 
                        background="darkgrey", 
                        troughcolor="black", 
                        bordercolor="black",
                        arrowcolor="white")
        style.map("Vertical.TScrollbar",
                  background=[('active', 'gray')])


        # Resource selection
        resource_frame = tk.LabelFrame(self.main_frame, text="Instrument Connection", padx=10, pady=10, bg="black", fg="white") # FIX: Set background and foreground
        resource_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(resource_frame, text="VISA Resource:", bg="black", fg="white").grid(row=0, column=0, sticky=tk.W, pady=2) # FIX: Set background and foreground
        self.resource_dropdown = tk.OptionMenu(resource_frame, self.resource_var, "No Resources Found")
        self.resource_dropdown.config(bg="grey", fg="white", highlightbackground="grey", highlightcolor="grey", activebackground="darkgrey", activeforeground="white") # FIX: Set colors
        self.resource_dropdown.grid(row=0, column=1, sticky=tk.EW, pady=2)
        # FIX: Configure the dropdown menu itself
        self.resource_dropdown["menu"].config(bg="grey", fg="white", activebackground="darkgrey", activeforeground="white")


        # Using lambda for button commands to ensure correct binding
        self.refresh_button = tk.Button(resource_frame, text="Refresh", command=lambda: self.populate_resources(), bg="darkgrey", fg="white") # FIX: Set colors
        self.refresh_button.grid(row=0, column=2, padx=5, pady=2)

        self.connect_button = tk.Button(resource_frame, text="Connect", command=lambda: self.connect_instrument(), bg="darkgrey", fg="white") # FIX: Set colors
        self.connect_button.grid(row=0, column=3, padx=5, pady=2)
        
        self.disconnect_button = tk.Button(resource_frame, text="Disconnect", command=lambda: self.disconnect_instrument(), state=tk.DISABLED, bg="darkgrey", fg="white") # FIX: Set colors
        self.disconnect_button.grid(row=0, column=4, padx=5, pady=2)

        # Scan Configuration (User Input) - Moved to be directly under Instrument Connection
        scan_settings_frame = tk.LabelFrame(self.main_frame, text="Scan Configuration (Push to Device)", padx=10, pady=10, bg="black", fg="white") # FIX: Set background and foreground
        scan_settings_frame.pack(pady=10, padx=10, fill=tk.X) # Pack it here, after resource_frame

        row_idx = 0
        tk.Label(scan_settings_frame, text="Reference Level (dBm):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2) # FIX: Set background and foreground
        entry_ref_level = tk.Entry(scan_settings_frame, textvariable=self.desired_ref_level_var, bg="grey", fg="white", insertbackground="white") # FIX: Set colors
        entry_ref_level.grid(row=row_idx, column=1, sticky=tk.EW, pady=2)
        self.desired_setting_entries["ref_level"] = entry_ref_level

        row_idx += 1
        tk.Label(scan_settings_frame, text="Preamplifier (ON/OFF):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2) # FIX: Set background and foreground
        check_preamp = tk.Checkbutton(scan_settings_frame, variable=self.desired_preamp_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white") # FIX: Set colors
        check_preamp.grid(row=row_idx, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["preamp"] = check_preamp # Associate for coloring

        row_idx += 1
        tk.Label(scan_settings_frame, text="Display Scale (LOG/LIN):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2) # FIX: Set background and foreground
        check_log_scale = tk.Checkbutton(scan_settings_frame, variable=self.desired_log_scale_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white") # FIX: Set colors
        check_log_scale.grid(row=row_idx, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["log_scale"] = check_log_scale

        # New "Scan RBW" for segmentation
        row_idx += 1
        tk.Label(scan_settings_frame, text="Scan RBW (for Segmentation, Hz):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2) # FIX: Set background and foreground
        entry_scan_rbw_segmentation = tk.Entry(scan_settings_frame, textvariable=self.desired_scan_rbw_segmentation_var, bg="grey", fg="white", insertbackground="white") # FIX: Set colors
        entry_scan_rbw_segmentation.grid(row=row_idx, column=1, sticky=tk.EW, pady=2)
        self.desired_setting_entries["scan_rbw_segmentation"] = entry_scan_rbw_segmentation


        row_idx += 1
        tk.Label(scan_settings_frame, text="Max Hold Enabled:", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2) # FIX: Set background and foreground
        check_max_hold = tk.Checkbutton(scan_settings_frame, variable=self.desired_max_hold_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white") # FIX: Set colors
        check_max_hold.grid(row=row_idx, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["max_hold_enabled"] = check_max_hold
        
        row_idx += 1
        tk.Label(scan_settings_frame, text="Max Hold Time (s):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2) # FIX: Set background and foreground
        entry_max_hold_time = tk.Entry(scan_settings_frame, textvariable=self.desired_max_hold_time_var, bg="grey", fg="white", insertbackground="white") # FIX: Set colors
        entry_max_hold_time.grid(row=row_idx, column=1, sticky=tk.EW, pady=2)
        self.desired_setting_entries["max_hold_time"] = entry_max_hold_time

        row_idx += 1
        tk.Label(scan_settings_frame, text="Cycle Wait Time (s):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2) # FIX: Set background and foreground
        entry_cycle_wait = tk.Entry(scan_settings_frame, textvariable=self.desired_cycle_wait_time_var, bg="grey", fg="white", insertbackground="white") # FIX: Set colors
        entry_cycle_wait.grid(row=row_idx, column=1, sticky=tk.EW, pady=2)
        self.desired_setting_entries["cycle_wait_time"] = entry_cycle_wait
        
        # New: Scan Name
        row_idx += 1
        tk.Label(scan_settings_frame, text="Scan Name:", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2) # FIX: Set background and foreground
        entry_scan_name = tk.Entry(scan_settings_frame, textvariable=self.scan_name_var, bg="grey", fg="white", insertbackground="white") # FIX: Set colors
        entry_scan_name.grid(row=row_idx, column=1, sticky=tk.EW, pady=2)
        self.desired_setting_entries["scan_name"] = entry_scan_name

        # Output Folder with Open Folder Button
        row_idx += 1
        output_folder_frame = tk.Frame(scan_settings_frame, bg="black") # FIX: Set background
        output_folder_frame.grid(row=row_idx, column=0, columnspan=2, sticky=tk.EW, pady=2)
        tk.Label(output_folder_frame, text="Output Folder (Base):", bg="black", fg="white").pack(side=tk.LEFT, padx=(0, 5)) # FIX: Set background and foreground
        entry_output_folder = tk.Entry(output_folder_frame, textvariable=self.output_folder_var, bg="grey", fg="white", insertbackground="white") # FIX: Set colors
        entry_output_folder.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.desired_setting_entries["output_folder"] = entry_output_folder
        open_folder_button = tk.Button(output_folder_frame, text="Open Folder", command=lambda: self.open_output_folder(), bg="darkgrey", fg="white") # FIX: Set colors
        open_folder_button.pack(side=tk.RIGHT, padx=(5, 0))

        # Plot Options
        row_idx += 1
        tk.Label(scan_settings_frame, text="Include TV Band Markers:", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2) # FIX: Set background and foreground
        chk_tv_markers = tk.Checkbutton(scan_settings_frame, variable=self.include_tv_markers_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white") # FIX: Set colors
        chk_tv_markers.grid(row=row_idx, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["include_tv_markers"] = chk_tv_markers

        row_idx += 1
        tk.Label(scan_settings_frame, text="Include Government Band Markers:", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2) # FIX: Set background and foreground
        chk_gov_markers = tk.Checkbutton(scan_settings_frame, variable=self.include_gov_markers_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white") # FIX: Set colors
        chk_gov_markers.grid(row=row_idx, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["include_gov_markers"] = chk_gov_markers

        row_idx += 1
        tk.Label(scan_settings_frame, text="Open HTML Plot After Each Cycle:", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2) # FIX: Set background and foreground
        chk_open_html = tk.Checkbutton(scan_settings_frame, variable=self.open_html_after_complete_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white") # FIX: Set colors
        chk_open_html.grid(row=row_idx, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["open_html_after_complete"] = chk_open_html


        row_idx += 1
        # Frame for Apply Settings and Generate Plot buttons
        button_row_frame = tk.Frame(scan_settings_frame, bg="black") # FIX: Set background
        button_row_frame.grid(row=row_idx, column=0, columnspan=2, pady=10, sticky=tk.EW)
        button_row_frame.grid_columnconfigure(0, weight=1)
        button_row_frame.grid_columnconfigure(1, weight=1)

        self.apply_button = tk.Button(button_row_frame, text="Apply Settings to Device", command=lambda: self.apply_settings_to_device(), state=tk.DISABLED, bg="darkgrey", fg="white") # FIX: Set colors
        self.apply_button.grid(row=0, column=0, padx=5, sticky=tk.EW)

        # Plot button moved here
        self.plot_button = tk.Button(button_row_frame, text="Generate Plot (Average)", command=lambda: self.generate_average_plot(), state=tk.NORMAL, bg="blue", fg="white")
        self.plot_button.grid(row=0, column=1, padx=5, sticky=tk.EW)


        # Frame to hold both Band Selection and Preset Files side-by-side
        bands_and_presets_frame = tk.Frame(self.main_frame, bg="black") # FIX: Set background
        bands_and_presets_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        bands_and_presets_frame.grid_columnconfigure(0, weight=1)
        bands_and_presets_frame.grid_columnconfigure(1, weight=1)
        bands_and_presets_frame.grid_rowconfigure(0, weight=1)

        # Band Selection - Moved to appear after scan settings
        band_selection_frame = tk.LabelFrame(bands_and_presets_frame, text="Frequency Band Selection", padx=10, pady=10, bg="black", fg="white") # FIX: Set background and foreground
        band_selection_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        # Crucial change: added expand=True here to ensure it takes vertical space

        self.band_checkboxes = []
        self.band_vars = []

        band_canvas = tk.Canvas(band_selection_frame, bg="black", highlightbackground="black") # FIX: Set background
        band_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # FIX: Use ttk.Scrollbar for consistency with styling
        band_scrollbar = ttk.Scrollbar(band_selection_frame, orient="vertical", command=band_canvas.yview)
        band_scrollbar.pack(side=tk.RIGHT, fill="y")

        band_canvas.configure(yscrollcommand=band_scrollbar.set)
        band_canvas.bind('<Configure>', lambda e: band_canvas.configure(scrollregion = band_canvas.bbox("all")))

        self.inner_band_frame = tk.Frame(band_canvas, bg="black") # FIX: Set background
        band_canvas.create_window((0, 0), window=self.inner_band_frame, anchor="nw")

        for i, band in enumerate(SCAN_BAND_RANGES):
            var = tk.BooleanVar(self)
            chk = tk.Checkbutton(self.inner_band_frame, text=f"{band['Band Name']} ({band['Start MHz']:.3f}-{band['Stop MHz']:.3f} MHz)", variable=var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white") # FIX: Set colors
            chk.grid(row=i, column=0, sticky=tk.W, padx=5, pady=1)
            var.set(True) # Default all bands to selected
            self.band_checkboxes.append(chk)
            self.band_vars.append({"band": band, "var": var}) # Store band info with var
        
        # New: Device Preset Files Frame
        preset_files_frame = tk.LabelFrame(bands_and_presets_frame, text="Device Preset Files (C:\\PRESETS\\)", padx=10, pady=10, bg="black", fg="white") # FIX: Set background and foreground
        preset_files_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

        # Button to load selected preset
        self.load_preset_button = tk.Button(preset_files_frame, text="Load Selected Preset", command=self.load_selected_preset, state=tk.DISABLED, bg="darkgrey", fg="white") # FIX: Set colors
        self.load_preset_button.pack(pady=5)

        # Treeview for displaying preset files
        self.preset_tree = ttk.Treeview(preset_files_frame, columns=("Name",), show="headings", selectmode="browse")
        self.preset_tree.heading("Name", text="Preset File Name")
        self.preset_tree.column("Name", width=200, anchor="w")
        self.preset_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configure a style for blue text for "MON" presets
        self.preset_tree.tag_configure("Mon", foreground="blue")


        # Scrollbar for the Treeview
        preset_scrollbar = ttk.Scrollbar(preset_files_frame, orient="vertical", command=self.preset_tree.yview)
        preset_scrollbar.pack(side=tk.RIGHT, fill="y")
        self.preset_tree.configure(yscrollcommand=preset_scrollbar.set)

        # Bind selection event to enable/disable load button
        self.preset_tree.bind("<<TreeviewSelect>>", self._on_preset_select)

        self.progress_label = tk.Label(self.main_frame, text="Ready.", bg="black", fg="white") # FIX: Set background and foreground
        self.progress_label.pack(pady=5)

        # Debug Mode Checkbox at the very bottom
        debug_frame = tk.Frame(self.main_frame, bg="black") # FIX: Set background
        debug_frame.pack(pady=10, padx=10, fill=tk.X)
        tk.Checkbutton(debug_frame, text="Enable Debug Mode (Log VISA Commands)", variable=self.debug_mode_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white").pack(anchor=tk.W) # FIX: Set colors


        # Configure columns to expand
        for i in range(5): # Adjust if more columns are added
            resource_frame.grid_columnconfigure(i, weight=1)
        # Only 2 columns for scan_settings_frame now
        scan_settings_frame.grid_columnconfigure(0, weight=1)
        scan_settings_frame.grid_columnconfigure(1, weight=1)

        # Update VBW display initially (even if UI fields are removed, variables are still used)
        self.update_vbw_display()

    # Removed _check_and_enable_average_plot_button as it's always enabled now.


    def update_progress_label(self, message):
        """Updates the progress label on the GUI."""
        self.progress_label.config(text=message)

    def update_vbw_display(self):
        """Updates the VBW display based on the current RBW setting."""
        try:
            # VBW is now derived from Scan RBW for segmentation, not Instrument RBW
            scan_rbw_val = float(self.desired_scan_rbw_segmentation_var.get())
            self.desired_vbw_display_var.set(str(int(scan_rbw_val / 3)))
        except ValueError:
            self.desired_vbw_display_var.set("Invalid RBW")


    def open_output_folder(self):
        """Opens the specified output folder in the file explorer."""
        folder_path = self.output_folder_var.get()
        if not os.path.exists(folder_path):
            messagebox.showwarning("Folder Not Found", f"The folder '{folder_path}' does not exist.")
            print(f"🚫 Folder not found: {folder_path}")
            return
        
        try:
            if sys.platform == "win32":
                os.startfile(folder_path)
            elif sys.platform == "darwin": # macOS
                subprocess.run(['open', folder_path])
            else: # linux variants
                subprocess.run(['xdg-open', folder_path])
            print(f"✅ Opened folder: {folder_path}")
        except Exception as e:
            messagebox.showerror("Error Opening Folder", f"Failed to open folder '{folder_path}': {e}")
            print(f"❌ Error opening folder: {e}")

    def connect_instrument(self):
        """Establishes connection to the selected instrument and queries its settings."""
        selected_resource = self.resource_var.get()
        if selected_resource == "No resources found" or "Error listing resources" in selected_resource:
            messagebox.showwarning("Connection Warning", "Please select a valid VISA resource.")
            return

        if self.inst:
            try:
                self.inst.close()
                print("🔌 Closed existing connection.")
            except Exception as e:
                print(f"Error closing existing connection: {e}")

        try:
            self.inst = self.rm.open_resource(selected_resource)
            self.inst.timeout = 5000 # 5 seconds timeout
            self.inst.read_termination = '\n' # N9340B typically terminates with newline
            self.inst.write_termination = '\n'
            
            # Query instrument ID
            instrument_id = query_safe(self.inst, "*IDN?")
            if instrument_id:
                print(f"✅ Connected to: {instrument_id.strip()}")
                # Updated regex to specifically match N9342CN or N9340B
                model_match = re.search(r'(N9342CN|N9340B)', instrument_id) 
                self.instrument_model = model_match.group(0) if model_match else "Unknown Model"
                
                # Update window title with detected model
                self.title(f"RF Spectrum Analyzer Controller - {self.instrument_model} - {os.path.basename(sys.argv[0])}")

                # Removed success dialog box: messagebox.showinfo("Connection Successful", f"Connected to: {instrument_id.strip()}")
                
                # No longer querying and displaying 'current' settings on GUI, but keeping the print for console
                self.query_instrument_settings() 
                
                # Now, push the desired settings from the GUI to the instrument (initial configuration)
                ref_level = float(self.desired_ref_level_var.get())
                preamp_on = self.desired_preamp_var.get()
                display_log = self.desired_log_scale_var.get()
                # RBW and VBW values are now taken from scan_rbw_segmentation and derived VBW
                rbw_config = int(float(self.desired_scan_rbw_segmentation_var.get())) # Use Scan RBW for instrument RBW
                vbw_config = int(float(self.desired_vbw_display_var.get())) # Use the calculated VBW
                max_hold_on = self.desired_max_hold_var.get() # Get max hold state from GUI

                # Reset the instrument to a known state using *RST first during connection
                write_safe(self.inst, "*RST")
                query_safe(self.inst, "*OPC?")
                time.sleep(1) # Give it a moment after reset

                if initialize_instrument(self.inst, ref_level, preamp_on, display_log, rbw_config, vbw_config, max_hold_on):
                    self.start_scan_button.config(state=tk.NORMAL)
                    self.stop_scan_button.config(state=tk.DISABLED)
                    self.pause_resume_button.config(state=tk.DISABLED) # Initially disabled
                    self.disconnect_button.config(state=tk.NORMAL)
                    self.apply_button.config(state=tk.NORMAL) # Enable apply button after successful connection
                    self.reset_setting_colors() # Settings pushed, so revert colors to black
                    # No longer re-querying to update GUI, as current settings display is removed

                    # NEW: Query and populate device preset files
                    self.query_device_presets()
                    self.load_preset_button.config(state=tk.DISABLED) # Disable load button until a preset is selected
                else:
                    messagebox.showerror("Initialization Failed", "Instrument initialization with desired settings failed.")
                    self.inst.close()
                    self.inst = None
                    self.start_scan_button.config(state=tk.DISABLED)
                    self.stop_scan_button.config(state=tk.DISABLED)
                    self.pause_resume_button.config(state=tk.DISABLED)
                    self.disconnect_button.config(state=tk.DISABLED)
                    self.apply_button.config(state=tk.DISABLED)
                    self.load_preset_button.config(state=tk.DISABLED)
            else:
                messagebox.showerror("Connection Failed", "Could not query instrument ID. Check connection or address.")
                if self.inst:
                    self.inst.close()
                self.inst = None
                self.start_scan_button.config(state=tk.DISABLED)
                self.stop_scan_button.config(state=tk.DISABLED)
                self.pause_resume_button.config(state=tk.DISABLED)
                self.disconnect_button.config(state=tk.DISABLED)
                self.apply_button.config(state=tk.DISABLED)
                self.load_preset_button.config(state=tk.DISABLED)
        except pyvisa.errors.VisaIOError as e:
            messagebox.showerror("VISA Error", f"Failed to connect to {selected_resource}: {e}")
            self.inst = None
            self.start_scan_button.config(state=tk.DISABLED)
            self.stop_scan_button.config(state=tk.DISABLED)
            self.pause_resume_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.DISABLED)
            self.apply_button.config(state=tk.DISABLED)
            self.load_preset_button.config(state=tk.DISABLED)
        except ValueError as e:
            messagebox.showerror("Input Error", f"Invalid numeric value for instrument settings: {e}. Please check Reference Level, RBW, and Max Hold Time.")
            self.start_scan_button.config(state=tk.DISABLED)
            self.stop_scan_button.config(state=tk.DISABLED)
            self.pause_resume_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.DISABLED)
            self.apply_button.config(state=tk.DISABLED)
            self.load_preset_button.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
            self.inst = None
            self.start_scan_button.config(state=tk.DISABLED)
            self.apply_button.config(state=tk.DISABLED)
            self.load_preset_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.DISABLED)
            self.pause_resume_button.config(state=tk.DISABLED)

            
    def disconnect_instrument(self):
        """Closes the connection to the instrument."""
        if self.inst:
            try:
                self.inst.close()
                self.inst = None
                self.instrument_model = None # Reset model on disconnect
                print("🔌 Instrument disconnected.")
                # Removed success dialog box: messagebox.showinfo("Disconnected", "Instrument disconnected successfully.")
                self.start_scan_button.config(state=tk.DISABLED)
                self.stop_scan_button.config(state=tk.DISABLED)
                self.pause_resume_button.config(state=tk.DISABLED) # Disable on disconnect
                self.disconnect_button.config(state=tk.DISABLED)
                self.apply_button.config(state=tk.DISABLED) # Disable apply button on disconnect
                self.update_progress_label("Disconnected.")
                # Clear the preset treeview on disconnect
                for item in self.preset_tree.get_children():
                    self.preset_tree.delete(item)
                self.load_preset_button.config(state=tk.DISABLED) # Disable load button on disconnect
                # The plot button is now always enabled, so no need to disable it here.
                # After disconnecting, refresh the resource list to allow reconnection
                self.populate_resources() 
                # Revert title to default
                self.title(f"RF Spectrum Analyzer Controller - {os.path.basename(sys.argv[0])}")
            except Exception as e:
                messagebox.showerror("Disconnect Error", f"Error disconnecting instrument: {e}")
                print(f"Error disconnecting: {e}")
        else:
            messagebox.showwarning("Disconnect Warning", "No instrument is currently connected.")
            
    def apply_settings_to_device(self):
        """Applies the desired settings from the GUI to the connected instrument."""
        if not self.inst:
            messagebox.showwarning("Not Connected", "Please connect to an instrument first to apply settings.")
            return

        try:
            ref_level = float(self.desired_ref_level_var.get())
            preamp_on = self.desired_preamp_var.get()
            display_log = self.desired_log_scale_var.get()
            # RBW and VBW values are now taken from scan_rbw_segmentation and derived VBW
            rbw_config = int(float(self.desired_scan_rbw_segmentation_var.get())) # Use Scan RBW for instrument RBW
            vbw_config = int(float(self.desired_vbw_display_var.get())) # Use the calculated VBW
            max_hold_on = self.desired_max_hold_var.get() # Get max hold state from GUI

            if initialize_instrument(self.inst, ref_level, preamp_on, display_log, rbw_config, vbw_config, max_hold_on):
                # Removed success dialog box: messagebox.showinfo("Settings Applied", "Desired settings successfully applied to the instrument.")
                print("Desired settings successfully applied to the instrument.")
                self.reset_setting_colors() # Revert colors to black after successful application
                # No longer re-querying to update GUI, as current settings display is removed
            else:
                messagebox.showerror("Apply Failed", "Failed to apply settings to the instrument. Check console for details.")
        except ValueError as e:
            messagebox.showerror("Input Error", f"Invalid numeric value for instrument settings: {e}. Please check Reference Level, RBW, and Max Hold Time.")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred while applying settings: {e}")
            print(f"❌ Error applying settings: {e}")

    def reset_setting_colors(self):
        """Resets the text color of all desired setting entries to black."""
        for key, entry_widget in self.desired_setting_entries.items():
            if isinstance(entry_widget, tk.Entry):
                entry_widget.config(fg="white") # Keep white text on grey background
            # For Checkbuttons, there's no direct foreground change for the text
            # unless a custom widget is implemented.

    def query_instrument_settings(self):
        """
        Queries the instrument for its current settings and prints them to the console.
        This function no longer updates GUI labels as the 'Current Device Configuration' section is removed.
        """
        if not self.inst:
            print("Not connected to instrument, cannot query settings.")
            return

        print("\nQuerying current instrument settings from device (console log only)...")
        try:
            print(f"  Reference Level (dBm): {query_safe(self.inst, ':DISPlay:WINDow:TRACe:Y:RLEVel?')}")
            print(f"  Preamplifier (ON/OFF): {query_safe(self.inst, ':SENSe:POWer:RF:GAIN:STATe?')}")
            print(f"  Display Scale (LOG/LIN): {query_safe(self.inst, ':DISPlay:WINDow:TRACe:Y:SCALe:SPACing?')}")
            
            # Query RBW in Hz and convert to MHz for display
            rbw_hz = query_safe(self.inst, ':SENSe:BANDwidth:RESolution?')
            if rbw_hz:
                try:
                    print(f"  RBW: {float(rbw_hz) / MHZ_TO_HZ:.3f} MHz")
                except ValueError:
                    print(f"  RBW: {rbw_hz} (could not convert to MHz)")

            print(f"  VBW (Hz): {query_safe(self.inst, ':SENSe:BANDwidth:VIDeo?')}")
            print(f"  Sweep Time Auto (ON/OFF): {query_safe(self.inst, ':SENSe:SWEep:TIME:AUTO?')}")
            
            # Query Start Freq in Hz and convert to MHz for display
            start_freq_hz = query_safe(self.inst, ':SENSe:FREQuency:STARt?')
            if start_freq_hz:
                try:
                    print(f"  Start Freq: {float(start_freq_hz) / MHZ_TO_HZ:.3f} MHz")
                except ValueError:
                    print(f"  Start Freq: {start_freq_hz} (could not convert to MHz)")
            
            # Query Stop Freq in Hz and convert to MHz for display
            stop_freq_hz = query_safe(self.inst, ':SENSe:FREQuency:STOP?')
            if stop_freq_hz:
                try:
                    print(f"  Stop Freq: {float(stop_freq_hz) / MHZ_TO_HZ:.3f} MHz")
                    # For N9340B, the stop frequency might be the end of the span, not necessarily the actual stop.
                    # This is more for general info, the actual scan logic handles spans.
                except ValueError:
                    print(f"  Stop Freq: {stop_freq_hz} (could not convert to MHz)")
            
            print("✅ Current instrument settings queried successfully.")

        except Exception as e:
            print(f"🛑 Error querying instrument settings: {e}")
            # No messagebox here as the GUI display for these is removed.

    def query_device_presets(self):
        """
        Queries the connected instrument for a list of preset files in "C:\\PRESETS\\"
        and populates the preset_tree Treeview.
        """
        if not self.inst:
            print("Not connected to instrument, cannot query device presets.")
            return

        # Clear existing items in the Treeview
        for item in self.preset_tree.get_children():
            self.preset_tree.delete(item)

        print("\nQuerying device preset files from C:\\PRESETS\\...")
        try:
            # Send the SCPI command to catalog presets
            response = query_safe(self.inst, 'MMEMory:CATalog? "C:\\\\PRESETS\\\\"')

            if response is None:
                print("🚫 No response received for preset catalog query.")
                self.preset_tree.insert("", "end", values=("No presets found or device error.",))
                self.load_preset_button.config(state=tk.DISABLED) # Disable load button if no presets
                return

            # Example data: <- 9090,54010,31,450WT470.STA,STA,1488,2022/03/12 16:28, ...
            # Split the response by commas
            parts = response.split(',')

            if len(parts) < 3: # Expect at least mem_used, mem_free, total_items
                print(f"🚫 Unexpected response format for preset catalog: {response}")
                self.preset_tree.insert("", "end", values=("Error parsing presets.",))
                self.load_preset_button.config(state=tk.DISABLED)
                return

            # The actual item listings start after the first 3 parts
            # Each item is 4 parts long: name, type, size, modified_time
            # So, we start from index 3 and take groups of 4
            
            # The total items count is at index 2 (0-indexed)
            try:
                total_items = int(parts[2])
            except ValueError:
                print(f"Warning: Could not parse total items count from response: {parts[2]}")
                total_items = 0 # Assume 0 if parsing fails

            preset_files = []
            # Iterate through the item listings, which start from index 3
            # Each item is 4 parts long: name, type, size, modified_time
            for i in range(3, len(parts), 4):
                if i + 3 < len(parts): # Ensure there are enough parts for a full item entry
                    name = parts[i].strip()
                    item_type = parts[i+1].strip()
                    # Only interested in .STA files as per user request
                    if item_type.upper() == "STA" and name.upper().endswith(".STA"):
                        preset_files.append(name)
                else:
                    print(f"Warning: Incomplete item entry found at index {i} in preset catalog response.")
                    break # Stop if an incomplete item entry is encountered

            if preset_files:
                for preset_name in sorted(preset_files): # Sort alphabetically for better readability
                    tags = ()
                    if "MON" in preset_name.upper(): # Check for "MON" (case-insensitive)
                        tags = ("Mon",) # Apply the "Mon" tag
                    self.preset_tree.insert("", "end", values=(preset_name,), tags=tags)
                print(f"✅ Found {len(preset_files)} '.STA' preset files.")
                self.load_preset_button.config(state=tk.DISABLED) # Initially disabled until selection
            else:
                self.preset_tree.insert("", "end", values=("No .STA preset files found.",))
                print("🚫 No '.STA' preset files found in C:\\PRESETS\\.")
                self.load_preset_button.config(state=tk.DISABLED)

        except pyvisa.errors.VisaIOError as e:
            messagebox.showerror("VISA Error", f"Failed to query device presets: {e}")
            print(f"🛑 VISA Error querying device presets: {e}")
            self.preset_tree.insert("", "end", values=(f"VISA Error: {e}",))
            self.load_preset_button.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred while querying presets: {e}")
            print(f"❌ Error querying device presets: {e}")
            self.preset_tree.insert("", "end", values=(f"Error: {e}",))
            self.load_preset_button.config(state=tk.DISABLED)

    def _on_preset_select(self, event):
        """Enables the Load Preset button if a preset is selected."""
        selected_items = self.preset_tree.selection()
        if selected_items and self.inst: # Only enable if an item is selected AND instrument is connected
            self.load_preset_button.config(state=tk.NORMAL)
        else:
            self.load_preset_button.config(state=tk.DISABLED)

    def load_selected_preset(self):
        """
        Loads the selected preset file onto the instrument.
        """
        if not self.inst:
            messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
            return

        selected_items = self.preset_tree.selection()
        if not selected_items:
            messagebox.showwarning("No Preset Selected", "Please select a preset file from the list to load.")
            return

        # Get the selected item's values (which is a tuple, we need the first element for the name)
        selected_preset_name = self.preset_tree.item(selected_items[0], 'values')[0]

        # Construct the full path
        preset_path = f"C:\\\\PRESETS\\\\{selected_preset_name}"
        
        # Construct the SCPI command
        command = f':MMEMory:LOAD STA,"{preset_path}"'
        
        print(f"\nAttempting to load preset: {selected_preset_name}")
        try:
            if write_safe(self.inst, command):
                print(f"✅ Preset '{selected_preset_name}' loaded successfully.")
                
                # Query and display current instrument settings after loading preset
                print("\n--- Current Instrument Settings after Preset Load ---")
                rbw_hz = query_safe(self.inst, ':SENSe:BANDwidth:RESolution?')
                start_freq_hz = query_safe(self.inst, ':SENSe:FREQuency:STARt?')
                stop_freq_hz = query_safe(self.inst, ':SENSe:FREQuency:STOP?')

                if rbw_hz:
                    try:
                        print(f"  RBW: {float(rbw_hz) / MHZ_TO_HZ:.3f} MHz")
                    except ValueError:
                        print(f"  RBW: {rbw_hz} (could not convert to MHz)")
                if start_freq_hz:
                    try:
                        print(f"  Start Freq: {float(start_freq_hz) / MHZ_TO_HZ:.3f} MHz")
                    except ValueError:
                        print(f"  Start Freq: {start_freq_hz} (could not convert to MHz)")
                if stop_freq_hz:
                    try:
                        print(f"  Stop Freq: {float(stop_freq_hz) / MHZ_TO_HZ:.3f} MHz")
                    except ValueError:
                        print(f"  Stop Freq: {stop_freq_hz} (could not convert to MHz)")
                print("--------------------------------------------------")

            else:
                messagebox.showerror("Load Preset Failed", f"Failed to load preset: {selected_preset_name}. Check console for details.")
                print(f"❌ Failed to load preset '{selected_preset_name}'.")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred while loading preset: {e}")
            print(f"❌ Error loading preset: {e}")


    def start_scan_thread(self):
        """Starts the scanning process in a separate thread."""
        if not self.inst:
            messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
            return
        
        if self.scanning:
            messagebox.showwarning("Scan in Progress", "A scan is already running.")
            return

        # Disable buttons during scan
        self.start_scan_button.config(state=tk.DISABLED)
        self.stop_scan_button.config(state=tk.NORMAL)
        self.pause_resume_button.config(state=tk.NORMAL) # Enable pause button when scan starts
        self.connect_button.config(state=tk.DISABLED)
        self.disconnect_button.config(state=tk.DISABLED)
        self.apply_button.config(state=tk.DISABLED)
        self.load_preset_button.config(state=tk.DISABLED) # Disable load button during scan
        # Plot button is now always enabled, no need to disable it here.

        self.scanning = True
        self.paused = False # Ensure not paused at start
        self.pause_resume_button.config(text="Pause Scan") # Set text to Pause

        print("\nStarting continuous spectrum scan...")
        
        # Get scan configuration from GUI
        max_hold_enabled = self.desired_max_hold_var.get()
        max_hold_time = float(self.desired_max_hold_time_var.get()) if max_hold_enabled else 0
        
        # Get the new "Scan RBW" for segmentation
        scan_rbw_segmentation = float(self.desired_scan_rbw_segmentation_var.get())

        # RBW and VBW values for instrument configuration are now derived from scan_rbw_segmentation
        rbw_config_val = scan_rbw_segmentation # Instrument RBW is set to Scan RBW
        vbw_config_val = int(rbw_config_val / 3) # VBW is RBW / 3

        # Get selected bands
        selected_bands = [item["band"] for item in self.band_vars if item["var"].get()]
        if not selected_bands:
            messagebox.showwarning("No Bands Selected", "Please select at least one frequency band to scan.")
            print("🚫 No bands selected for scan.")
            self.stop_scan() # Reset GUI state
            return

        # Ensure base output directory exists
        base_output_dir = self.output_folder_var.get()
        if not os.path.exists(base_output_dir):
            os.makedirs(base_output_dir)
            print(f"Created base output directory: {base_output_dir}")

        # Pass all necessary arguments to _run_scan
        scan_thread = threading.Thread(target=self._run_scan, 
                                       args=(selected_bands, 
                                             scan_rbw_segmentation, rbw_config_val, 
                                             vbw_config_val, max_hold_time))
        scan_thread.daemon = True # Allow program to exit even if thread is running
        scan_thread.start()

    def toggle_pause_scan(self):
        """Toggles the paused state of the scan."""
        if self.scanning: # Only allow pause/resume if a scan is active
            self.paused = not self.paused
            if self.paused:
                self.pause_resume_button.config(text="Resume Scan", bg="blue")
                self.update_progress_label("Scan Paused.")
                print("Scan paused.")
            else:
                self.pause_resume_button.config(text="Pause Scan", bg="orange")
                self.update_progress_label("Scan Resumed.")
                print("Scan resumed.")
        else:
            messagebox.showwarning("Scan Not Active", "No scan is currently running to pause or resume.")


    def _run_scan(self, selected_bands, scan_rbw_segmentation, rbw_config_val, vbw_config_val, max_hold_time):
        """Internal method to run the scan logic, called by the thread."""
        try:
            # Loop for continuous scanning
            while self.scanning:
                # Check for pause state before starting a new cycle or segment
                while self.paused:
                    self.after(100, self.update_progress_label, "Scan Paused. Click Resume to continue.")
                    time.sleep(0.5) # Sleep briefly while paused
                    if not self.scanning: # Allow stopping even when paused
                        break # Exit pause loop if scan is stopped

                if not self.scanning: # Check again after pause loop
                    print("\nScan process finished (interrupted).")
                    self.after(0, self.update_progress_label, "Scan interrupted by user.")
                    break # Exit the while loop


                # Get current scan name, RBW, HOLD time, and timestamp
                scan_name = self.scan_name_var.get()
                if not scan_name:
                    scan_name = "UnnamedScan" # Fallback if user leaves it blank
                
                # Format RBW and HOLD for filename
                rbw_str = f"RBW{int(scan_rbw_segmentation/1000):04d}K" # e.g., RBW0010K for 10kHz
                hold_str = f"HOLD{int(max_hold_time):02d}" # e.g., HOLD03 for 3 seconds

                # Date and time for filename
                datetime_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                # CSV filename using the new convention
                csv_filename = f"{scan_name}_{rbw_str}_{hold_str}_{datetime_str}.csv"
                current_csv_file_path = os.path.join(self.output_folder_var.get(), csv_filename)
                self.last_csv_file_path = current_csv_file_path # Update for single plotting

                try:
                    with open(current_csv_file_path, 'w', newline='') as current_csv_file:
                        csv_writer = csv.writer(current_csv_file)
                        # Updated header to reflect data being in MHz
                        csv_writer.writerow(["Frequency (MHz)", "Level (dBm)"])
                        print(f"CSV file created for this cycle: {current_csv_file_path}")

                        # Perform the scan for this cycle
                        scanned_data, last_successful_band_index = scan_bands(
                            self, csv_writer, selected_bands, 
                            scan_rbw_segmentation, rbw_config_val, 
                            vbw_config_val, max_hold_time
                        ) 
                        
                        # Convert to DataFrame and append to collected scans for averaging
                        if scanned_data:
                            # When creating df_scan, ensure columns match what's written to CSV
                            # The data in scanned_data is (Frequency_Hz, Power_dBm)
                            # We need to convert Frequency_Hz to Frequency_MHz for the DataFrame
                            df_scan = pd.DataFrame(scanned_data, columns=['Frequency_Hz', 'Power_dBm'])
                            df_scan['Frequency_MHz'] = df_scan['Frequency_Hz'] / MHZ_TO_HZ
                            # Only keep the columns relevant for historical storage/plotting
                            self.collected_scans_dataframes.append(df_scan[['Frequency_MHz', 'Power_dBm']].copy())
                            print(f"✅ Stored scan data for averaging. Total scans collected: {len(self.collected_scans_dataframes)}")

                        self.last_scan_data = scanned_data # Store for single plot
                        
                        print(f"Cycle scan finished. Data saved to: {current_csv_file_path}")

                        # Auto-plot after each cycle completes (if enabled)
                        if self.open_html_after_complete_var.get():
                            # Pass the scan_name and formatted RBW/HOLD/Date_time for the plot title suffix
                            plot_suffix = f"{scan_name}_{rbw_str}_{hold_str}_{datetime_str}"
                            self.after(0, self.generate_single_scan_plot_and_open_wrapper, scanned_data, plot_suffix) 
                        
                        # Removed: Trigger the averaged CSV and plot generation after each scan cycle completes.
                        # This will now ONLY be triggered by the "Generate Plot (Average)" button.
                        # if self.collected_scans_dataframes: # Only if there's data to average
                        #     self.after(0, self._generate_current_cycle_average_csv_and_plot) 
                        
                        # Check if scan was stopped before waiting
                        if not self.scanning:
                            print("\nScan process finished (interrupted).")
                            self.after(0, self.update_progress_label, "Scan interrupted by user.")
                            break # Exit the while loop
                        
                        # Wait for the next cycle
                        wait_time = float(self.desired_cycle_wait_time_var.get())
                        if wait_time > 0:
                            self.after(0, self.update_progress_label, f"Waiting {wait_time} seconds for next cycle...")
                            print(f"Waiting {wait_time} seconds before next scan cycle...")
                            # FIX: Replaced single time.sleep with a loop to allow pause checks
                            for _ in range(int(wait_time * 10)): # Check every 0.1 seconds
                                while self.paused:
                                    self.after(100, self.update_progress_label, "Scan Paused. Click Resume to continue.")
                                    time.sleep(0.1) # Sleep briefly while paused
                                    if not self.scanning: # Allow stopping even when paused
                                        print("\nScan process finished (interrupted during pause in wait).")
                                        self.after(0, self.update_progress_label, "Scan interrupted during pause in wait.")
                                        break # Exit pause loop if scan is stopped
                                
                                if not self.scanning: # Check again after pause loop
                                    print("\nScan process finished (interrupted during wait).")
                                    self.after(0, self.update_progress_label, "Scan interrupted during wait.")
                                    break # Exit wait loop if scan is stopped
                                time.sleep(0.1) # Small sleep to allow other threads/GUI to run

                except Exception as e:
                    self.after(0, messagebox.showerror, "Scan Cycle Error", f"An error occurred during scan cycle: {e}")
                    print(f"❌ Scan cycle encountered an error: {e}")
                    self.after(0, self.update_progress_label, f"Scan cycle error: {e}")
                    self.scanning = False # Stop scanning on error
                    break # Exit the while loop

            # End of while loop
            print("\nContinuous scan process terminated.")
            self.after(0, self.update_progress_label, "Continuous scan terminated.")
            
        except Exception as e:
            self.after(0, messagebox.showerror, "Scan Thread Error", f"An unexpected error occurred in main scan thread: {e}")
            print(f"❌ Main scan thread encountered an error: {e}")
            self.after(0, self.update_progress_label, f"Main scan thread error: {e}")
        finally:
            self.scanning = False
            self.paused = False # Ensure paused flag is reset
            self.after(100, self.reset_scan_buttons) # Reset GUI buttons

    def populate_resources(self):
        """Populates the VISA resource dropdown."""
        try:
            self.instrument_list = self.rm.list_resources()
            if self.instrument_list:
                self.resource_var.set(self.instrument_list[0])
                menu = self.resource_dropdown["menu"]
                menu.delete(0, "end")
                for resource in self.instrument_list:
                    menu.add_command(label=resource, command=tk._setit(self.resource_var, resource))
            else:
                self.resource_var.set("No resources found")
                menu = self.resource_dropdown["menu"]
                menu.delete(0, "end")
                menu.add_command(label="No resources found", command=tk._setit(self.resource_var, "No resources found"))
            self.start_scan_button.config(state=tk.DISABLED) # Disable until connected
            self.disconnect_button.config(state=tk.DISABLED)
            self.apply_button.config(state=tk.DISABLED) # Disable apply until connected
            self.connect_button.config(state=tk.NORMAL) # Enable connect button
            self.load_preset_button.config(state=tk.DISABLED) # Disable load button until connected and selected
            self.pause_resume_button.config(state=tk.DISABLED) # Disable pause button
            # Plot button is always enabled, no need to check here.
        except Exception as e:
            messagebox.showerror("Error", f"Failed to list VISA resources: {e}")
            self.resource_var.set("Error listing resources")
            self.start_scan_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.DISABLED)
            self.apply_button.config(state=tk.DISABLED)
            self.connect_button.config(state=tk.DISABLED) # Disable connect if error
            self.load_preset_button.config(state=tk.DISABLED)
            self.pause_resume_button.config(state=tk.DISABLED)
            # Plot button is always enabled, no need to disable here.

    def stop_scan(self):
        """Stops the ongoing scan."""
        self.scanning = False
        self.paused = False # Ensure paused state is reset on stop
        print("\nAttempting to stop scan... Please wait for current sweep to finish.")
        self.stop_scan_button.config(state=tk.DISABLED) # Disable stop button immediately
        self.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED) # Reset pause button text and disable

    def reset_scan_buttons(self):
        """Resets the state of scan-related buttons after a scan completes or stops."""
        self.start_scan_button.config(state=tk.NORMAL)
        if self.inst: # Only enable disconnect/apply if connected
            self.disconnect_button.config(state=tk.NORMAL)
            self.apply_button.config(state=tk.NORMAL)
            # Re-enable load preset button if a preset is selected and instrument is connected
            if self.preset_tree.selection():
                self.load_preset_button.config(state=tk.NORMAL)
        self.stop_scan_button.config(state=tk.DISABLED)
        self.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED) # Reset pause button text and disable
        # Plot button is always enabled, no need to change its state here.


    def generate_single_scan_plot_and_open_wrapper(self, scanned_data, plot_title_suffix):
        """Wrapper to call plot_single_scan_data and handle saving/opening."""
        if not scanned_data:
            print("No single scan data available to plot.")
            return

        print("Generating single scan plot...")
        try:
            fig, plot_filename = plot_single_scan_data( # Get figure and suggested filename
                scanned_data, 
                plot_title_suffix, # Pass the suffix for the title
                include_tv_markers=self.include_tv_markers_var.get(),
                include_gov_markers=self.include_gov_markers_var.get()
            )
            
            # Save the plot in the base output folder
            plot_path = os.path.join(self.output_folder_var.get(), plot_filename)
            fig.write_html(plot_path)

            print(f"✅ Single scan plot generation complete: {plot_path}")
            if self.open_html_after_complete_var.get():
                try:
                    webbrowser.open(plot_path)
                    print(f"✅ Opened single scan plot in browser: {plot_path}")
                except Exception as e:
                    print(f"❌ Failed to open single scan plot in browser: {e}")
                    messagebox.showwarning("Open Plot Error", f"Could not open single scan plot in web browser automatically: {e}")
        except Exception as e:
            messagebox.showerror("Single Plot Error", f"Failed to generate single scan plot: {e}")
            print(f"❌ Error generating single scan plot: {e}")

    def _generate_current_cycle_average_csv_and_plot(self): # Renamed function
        """
        Calculates average, median, and range from collected scan data (from current scan cycle),
        saves them to a CSV, and plots them with overlays.
        This function is called on the main Tkinter thread via self.after().
        """
        if not self.collected_scans_dataframes:
            print("No scan data collected for current cycle averaging.")
            return

        print("\n📊 Generating averaged, median, and range data for current cycle...")

        # Combine all scan data into a single DataFrame for easier processing
        # Concatenate the collected_scans_dataframes vertically
        combined_current_scans_df = pd.concat(self.collected_scans_dataframes)
        
        # Now group by Frequency_MHz and apply the aggregations on 'Power_dBm'
        # The collected_scans_dataframes already have 'Frequency_MHz'
        aggregated_df = combined_current_scans_df.groupby('Frequency_MHz')['Power_dBm'].agg(
            Average_dBm='mean',
            Median_dBm='median',
            Max_dBm='max', # Intermediate for Range
            Min_dBm='min'  # Intermediate for Range
        ).reset_index() # Reset index to make Frequency_MHz a column again

        # Calculate Range (Max - Min)
        aggregated_df['Range_dBm'] = aggregated_df['Max_dBm'] - aggregated_df['Min_dBm']

        # Drop intermediate Max_dBm and Min_dBm columns
        aggregated_df = aggregated_df.drop(columns=['Max_dBm', 'Min_dBm'])

        # Frequency_MHz is already present from the groupby operation, no need to re-add.

        # Generate filename based on current time (HMS) and scan name
        timestamp_hms = datetime.now().strftime("%H%M%S") # HMS format
        base_filename = f"{self.scan_name_var.get()}_{timestamp_hms}"

        csv_filename = os.path.join(self.output_folder_var.get(), f"{base_filename}_averaged_cycle.csv") # Added _cycle to distinguish
        html_filename = os.path.join(self.output_folder_var.get(), f"{base_filename}_averaged_cycle.html") # Added _cycle to distinguish

        # Ensure output directory exists
        os.makedirs(self.output_folder_var.get(), exist_ok=True)

        # Save to CSV
        try:
            # Select columns for CSV: Frequency_MHz, Average_dBm, Median_dBm, Range_dBm
            aggregated_df.to_csv(csv_filename, index=False, float_format='%.2f',
                                 columns=['Frequency_MHz', 'Average_dBm', 'Median_dBm', 'Range_dBm'])
            print(f"✅ Averaged data for current cycle saved to: {csv_filename}")
            self.last_csv_file_path = csv_filename
        except Exception as e:
            print(f"❌ Failed to save averaged CSV for current cycle: {e}")
            messagebox.showerror("CSV Save Error", f"Could not save averaged CSV for current cycle: {e}")
            return

        # Plotting the averaged, median, and range data
        try:
            fig, plot_html_path_return = plot_multi_trace_data(
                aggregated_df,
                f"{self.scan_name_var.get()} - Averaged, Median & Range (Current Cycle {timestamp_hms})", # Include HMS in plot title
                self.include_tv_markers_var.get(),
                self.include_gov_markers_var.get(),
                output_html_path=html_filename # Pass the desired full path for the HTML file
            )

            if fig:
                fig.write_html(plot_html_path_return, auto_open=False)
                print(f"✅ Averaged plot for current cycle saved to: {plot_html_path_return}")
                if self.open_html_after_complete_var.get():
                    self._open_plot_in_browser(plot_html_path_return)
            else:
                print("🚫 Plotly figure was not generated for current cycle averaged data.")

        except Exception as e:
            print(f"❌ Failed to generate or save current cycle averaged plot: {e}")
            messagebox.showerror("Plot Error", f"Could not generate or save current cycle averaged plot: {e}")

        # Clear collected dataframes after processing to prepare for the next set of scans
        self.collected_scans_dataframes = []
        print("🗑️ Cleared collected scan data for next averaging cycle.")


    def generate_average_plot(self):
        """
        Generates an average, median, and range plot from ALL relevant CSV files
        found in the current output folder base. This is triggered by the button.
        This plot also includes all individual historical scans as overlay layers.
        """
        if self.scanning:
            messagebox.showwarning("Plotting Error", "Cannot generate historical average plot while a scan is in progress.")
            return

        base_output_dir = self.output_folder_var.get()
        if not os.path.exists(base_output_dir):
            messagebox.showwarning("Folder Not Found", f"The output folder '{base_output_dir}' does not exist. Please ensure it exists and contains scan data.")
            return

        all_historical_dfs = [] # To store DataFrames for aggregation
        historical_dfs_for_overlays = [] # To store (DataFrame, name) for plotting as overlays

        # Regex to extract timestamp (time only) from filename: e.g., MyScan_RBW####_HOLD##_YYYYMMDD_HHMMSS.csv
        # Updated regex to capture RBW, HOLD, and the full datetime string
        filename_pattern = re.compile(r'^(.*)_RBW(\d{4}K)_HOLD(\d{2})_(\d{8}_\d{6})\.csv$')

        # Iterate directly through files in the base output directory
        for file_name in os.listdir(base_output_dir):
            # Filter for CSVs that match the scan name prefix and are not the 'averaged_cycle' files
            if file_name.endswith(".csv") and file_name.startswith(self.scan_name_var.get() + "_") and "_averaged_cycle.csv" not in file_name and "_HISTORICAL_" not in file_name:
                csv_path = os.path.join(base_output_dir, file_name)
                try:
                    # Read CSV, assuming the header is "Frequency (MHz)", "Level (dBm)"
                    df = pd.read_csv(csv_path).copy()

                    # NEW: Strip whitespace from column names to handle potential invisible characters
                    df.columns = df.columns.str.strip()
                    print(f"Actual columns in {file_name} after stripping: {df.columns.tolist()}") # Debug print

                    # Ensure columns are as expected and rename for internal consistency
                    if "Frequency (MHz)" in df.columns and "Level (dBm)" in df.columns:
                        df.rename(columns={"Frequency (MHz)": "Frequency_MHz", "Level (dBm)": "Power_dBm"}, inplace=True)
                        # The data in "Frequency (MHz)" is already in MHz as per your scan_bands function
                        # So, we just rename it to Frequency_MHz. No division by MHZ_TO_HZ needed here.

                        # Append to list for aggregation, ensuring 'Frequency_MHz' and 'Power_dBm' columns
                        all_historical_dfs.append(df[['Frequency_MHz', 'Power_dBm']].copy())

                        # Extract components for layer name
                        match = filename_pattern.match(file_name)
                        if match:
                            scan_name_prefix, rbw_val, hold_val, datetime_val = match.groups()
                            # Format datetime for display
                            display_datetime = datetime.strptime(datetime_val, "%Y%m%d_%H%M%S").strftime("%Y-%m-%d %H:%M:%S")
                            display_name = f"{scan_name_prefix}_{rbw_val}_{hold_val} ({display_datetime})"
                            historical_dfs_for_overlays.append((df[['Frequency_MHz', 'Power_dBm']].copy(), display_name))
                        else:
                            historical_dfs_for_overlays.append((df[['Frequency_MHz', 'Power_dBm']].copy(), file_name)) # Fallback to full filename

                    else:
                        print(f"Skipping {file_name}: Missing expected columns or incorrect format. Expected 'Frequency (MHz)' and 'Level (dBm)'.")
                except Exception as e:
                    print(f"Error reading historical CSV {csv_path}: {e}")

        if not all_historical_dfs:
            messagebox.showwarning("No Data Found", f"No valid historical scan data CSV files with prefix '{self.scan_name_var.get()}_' found in '{base_output_dir}' to generate an average plot.")
            return

        print("📊 Generating historical averaged, median, and range data from CSV files...")

        # Concatenate all historical DataFrames vertically for aggregation
        combined_all_scans_df = pd.concat(all_historical_dfs)
        
        # Now group by Frequency_MHz and apply the aggregations on 'Power_dBm'
        # The key change is grouping by 'Frequency_MHz' directly.
        aggregated_df = combined_all_scans_df.groupby('Frequency_MHz')['Power_dBm'].agg(
            Average_dBm='mean',
            Median_dBm='median',
            Max_dBm='max', # Intermediate for Range
            Min_dBm='min'  # Intermediate for Range
        ).reset_index() # Reset index to make Frequency_MHz a column again

        # Calculate Range (Max - Min)
        aggregated_df['Range_dBm'] = aggregated_df['Max_dBm'] - aggregated_df['Min_dBm']

        # Drop intermediate Max_dBm and Min_dBm columns
        aggregated_df = aggregated_df.drop(columns=['Max_dBm', 'Min_dBm'])

        # Frequency_MHz is already present from the groupby operation, no need to re-add.

        # Generate filename based on current time (HMS) and scan name
        timestamp_hms = datetime.now().strftime("%H%M%S") # HMS format
        base_filename = f"{self.scan_name_var.get()}_HISTORICAL_{timestamp_hms}" # Distinct filename for historical average

        csv_filename = os.path.join(self.output_folder_var.get(), f"{base_filename}_averaged.csv")
        html_filename = os.path.join(self.output_folder_var.get(), f"{base_filename}_averaged.html")

        # Save to CSV
        try:
            # Select columns for CSV: Frequency_MHz, Average_dBm, Median_dBm, Range_dBm
            aggregated_df.to_csv(csv_filename, index=False, float_format='%.2f',
                                 columns=['Frequency_MHz', 'Average_dBm', 'Median_dBm', 'Range_dBm'])
            print(f"✅ Historical averaged data saved to: {csv_filename}")
            self.last_csv_file_path = csv_filename
        except Exception as e:
            print(f"❌ Failed to save historical averaged CSV: {e}")
            messagebox.showerror("CSV Save Error", f"Could not save historical averaged CSV: {e}")
            return

        # Plotting the historical averaged, median, and range data, PLUS historical overlays
        try:
            fig, plot_html_path_return = plot_multi_trace_data(
                aggregated_df,
                f"{self.scan_name_var.get()} - Historical Averaged, Median & Range ({timestamp_hms})", # Include HMS in plot title
                self.include_tv_markers_var.get(),
                self.include_gov_markers_var.get(),
                historical_dfs_with_names=historical_dfs_for_overlays, # Pass historical data for overlays
                output_html_path=html_filename # Pass the desired full path for the HTML file
            )

            if fig:
                fig.write_html(plot_html_path_return, auto_open=False)
                print(f"✅ Historical averaged plot saved to: {plot_html_path_return}")
                if self.open_html_after_complete_var.get():
                    self._open_plot_in_browser(plot_html_path_return)
            else:
                print("🚫 Plotly figure was not generated for historical averaged data.")

        except Exception as e:
            messagebox.showerror("Plot Error", f"Could not generate or save historical averaged plot: {e}")
            print(f"❌ Failed to generate or save historical averaged plot: {e}")


    def _open_plot_in_browser(self, plot_path):
        """Helper to open an HTML plot in the default web browser."""
        try:
            webbrowser.open(plot_path)
            print(f"✅ Opened plot in browser: {plot_path}")
        except Exception as e:
            print(f"❌ Failed to open plot in browser: {e}")
            messagebox.showwarning("Open Plot Error", f"Could not open plot in web browser automatically: {e}")

    def _update_console_line(self, text_to_display, overwrite=False):
        """
        Helper function to update the console output safely from any thread,
        handling line overwriting.
        """
        self.console_output.config(state=tk.NORMAL)
        if overwrite:
            try:
                # Delete from the start of the last line to the end
                self.console_output.delete("end-1c linestart", "end-1c")
            except TclError:
                # Handle case where there's no previous line to delete (e.g., first message)
                pass
        self.console_output.insert(tk.END, text_to_display)
        self.console_output.see(tk.END)
        self.console_output.config(state=tk.DISABLED)
        self.console_output.update_idletasks()


# The actual entry point of the script
if __name__ == '__main__':
    # Ensure dependencies are installed before running the app
    if check_and_install_dependencies():
        app = App()
        app.mainloop()
    else:
        print("Critical dependencies missing. Please install them to run the application.")
        messagebox.showerror("Dependency Error", "Some required Python packages are missing. Please install them manually and try again.")
