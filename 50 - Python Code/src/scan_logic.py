# src/scan_logic.py
import tkinter as tk
from tkinter import messagebox # Ensure messagebox is directly imported
import threading
import time
import os
from datetime import datetime
import pandas as pd
import pyvisa # Ensure pyvisa is imported

from utils.scan_instrument import scan_bands
from utils.frequency_bands import MHZ_TO_HZ # Assuming MHZ_TO_HZ is needed here
from utils.instrument_control import log_visa_command, debug_print # Import debug_print

def start_scan_thread_logic(app_instance):
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
        return
    
    if app_instance.scanning:
        messagebox.showwarning("Scan in Progress", "A scan is already running.")
        return

    # Save configuration when scan starts - this will save current GUI settings as LAST_USED
    from src.config_manager import save_config # Import locally to avoid circular dependency
    save_config(app_instance)

    app_instance.start_scan_button.config(state=tk.DISABLED)
    app_instance.stop_scan_button.config(state=tk.NORMAL)
    app_instance.pause_resume_button.config(state=tk.NORMAL)
    app_instance.connect_button.config(state=tk.DISABLED)
    app_instance.disconnect_button.config(state=tk.DISABLED)
    app_instance.apply_button.config(state=tk.DISABLED)
    app_instance.plot_button.config(state=tk.DISABLED)
    app_instance.load_preset_button.config(state=tk.DISABLED)
    app_instance._stop_connect_button_blink() # Stop blinking if it was on

    app_instance.scanning = True
    app_instance.paused = False
    app_instance.scan_cycle_count = 0 # Reset cycle count for a new scan
    app_instance.current_freq_offset = 0 # Reset frequency offset for a new scan
    app_instance.collected_scans_dataframes = [] # Clear previous scan data
    app_instance.last_scanned_band_index = 0 # Reset last scanned band index

    selected_bands = [item["band"] for item in app_instance.band_vars if item["var"].get()]
    if not selected_bands:
        messagebox.showwarning("No Bands Selected", "Please select at least one frequency band to scan.")
        app_instance.stop_scan() # Reset buttons
        return

    # Collect parameters from Tkinter variables for the scan
    # Ensure these variables match the expected types and values for scan_bands
    scan_rbw_segmentation_val = app_instance.desired_scan_rbw_segmentation_var.get()
    freq_shift_val = app_instance.shift_freq_var.get()
    
    # This is the RBW Step Size from the GUI, used as rbw_config_val in scan_bands
    rbw_config_val = app_instance.desired_rbw_var.get() 
    
    # This is the calculated VBW from the GUI, used as vbw_config_val in scan_bands
    vbw_config_val = app_instance.desired_vbw_display_var.get() 
    
    max_hold_time = app_instance.desired_max_hold_time_var.get()

    # Debug prints to verify values before starting the thread
    print(f"DEBUG: Starting scan with parameters:")
    print(f"DEBUG:   selected_bands: {[b['Band Name'] for b in selected_bands]}")
    print(f"DEBUG:   scan_rbw_segmentation_val: {scan_rbw_segmentation_val} (Type: {type(scan_rbw_segmentation_val)})")
    print(f"DEBUG:   freq_shift_val: {freq_shift_val} (Type: {type(freq_shift_val)})")
    print(f"DEBUG:   rbw_config_val: {rbw_config_val} (Type: {type(rbw_config_val)})")
    print(f"DEBUG:   vbw_config_val: {vbw_config_val} (Type: {type(vbw_config_val)})")
    print(f"DEBUG:   max_hold_time: {max_hold_time} (Type: {type(max_hold_time)})")

    # Start the scan in a separate thread
    app_instance.scan_thread = threading.Thread(target=app_instance._run_scan, args=(
        selected_bands,
        scan_rbw_segmentation_val,
        freq_shift_val,
        rbw_config_val,  # Correctly pass rbw_config_val
        vbw_config_val,  # Correctly pass vbw_config_val
        max_hold_time    # Correctly pass max_hold_time
    ))
    app_instance.scan_thread.daemon = True # Allow the thread to exit with the main app
    app_instance.scan_thread.start()
    print("Scan thread started.")


def toggle_pause_scan_logic(app_instance):
    if app_instance.scanning:
        app_instance.paused = not app_instance.paused
        if app_instance.paused:
            app_instance.pause_resume_button.config(text="Resume Scan")
            app_instance.after(100, app_instance._update_console_line, "Scan Paused. Click Resume to continue.", False) # Don't overwrite immediately
        else:
            app_instance.pause_resume_button.config(text="Pause Scan")
            app_instance.after(100, app_instance._update_console_line, "Scan Resumed.", True) # Overwrite the previous pause message
    else:
        messagebox.showwarning("No Scan in Progress", "No scan is currently running to pause or resume.")

def run_scan_logic(app_instance, selected_bands, scan_rbw_segmentation, freq_shift_value, rbw_config_val, vbw_config_val, max_hold_time):
    """
    Executes the main scan logic in a loop, handling multiple scan cycles.
    This function runs in a separate thread to keep the GUI responsive.

    Inputs:
        app_instance (App): The main application instance.
        selected_bands (list): List of frequency bands to scan.
        scan_rbw_segmentation (float): RBW for segmenting bands.
        freq_shift_value (float): Frequency offset to apply per cycle.
        rbw_config_val (float): RBW value to configure on the instrument.
        vbw_config_val (float): VBW value to configure on the instrument.
        max_hold_time (float): Max hold time for the instrument.
    Process:
        1. Loops indefinitely while `app_instance.scanning` is True.
        2. Increments `scan_cycle_count` and updates `current_freq_offset`.
        3. Calls `scan_bands` from `utils.scan_instrument` to perform the actual sweep.
        4. Handles `VisaIOError` for instrument communication issues, attempting to reconnect.
        5. Processes collected scan data, generates plots, and saves CSVs.
        6. Implements `cycle_wait_time_seconds` delay between cycles.
        7. Resets GUI buttons upon scan completion or interruption.
    Outputs: None (modifies app_instance state, generates files)
    """
    try:
        while app_instance.scanning:
            app_instance.scan_cycle_count += 1
            app_instance.current_freq_offset += freq_shift_value # Accumulate offset

            print(f"\n--- Starting Scan Cycle {app_instance.scan_cycle_count} (Offset: {app_instance.current_freq_offset:.0f} Hz) ---")

            try:
                # Call scan_bands with the correct parameters
                final_sweep_data, last_successful_band_index, csv_filename = scan_bands(
                    app_instance,
                    app_instance.inst,
                    selected_bands,
                    scan_rbw_segmentation,
                    rbw_config_val,
                    vbw_config_val,
                    max_hold_time,
                    app_instance.current_freq_offset,
                    last_scanned_band_index=app_instance.last_scanned_band_index
                )
                app_instance.last_scanned_band_index = last_successful_band_index # Update for potential resume

                if final_sweep_data:
                    # Convert the list of tuples to a DataFrame for plotting and storage
                    df = pd.DataFrame(final_sweep_data, columns=["Frequency_MHz", "Power_dBm"])
                    app_instance.collected_scans_dataframes.append(df)
                    print(f"✅ Data from Cycle {app_instance.scan_cycle_count} collected and stored.")

                    # Generate plot for the single scan
                    plot_title_suffix = f"Cycle {app_instance.scan_cycle_count} (Offset: {app_instance.current_freq_offset/MHZ_TO_HZ:.3f} MHz)"
                    output_html_path = os.path.join(app_instance.output_folder_var.get(), f"{app_instance.scan_name_var.get()}_Cycle{app_instance.scan_cycle_count}.html")
                    
                    app_instance.after(0, app_instance.generate_single_scan_plot_and_open_wrapper,
                                       csv_filename, # Use the CSV filename returned by scan_bands
                                       plot_title_suffix,
                                       output_html_path,
                                       app_instance.open_html_after_complete_var.get())
                else:
                    print(f"🚫 No data collected for Cycle {app_instance.scan_cycle_count}.")

            except pyvisa.errors.VisaIOError as e:
                print(f"❌ VISA communication error during scan: {e}")
                messagebox.showerror("VISA Error", f"Lost connection to instrument or communication error: {e}\nAttempting to reconnect...")
                app_instance.after(0, app_instance._reset_gui_on_disconnect_or_error) # Reset GUI on main thread
                app_instance.scanning = False # Stop scanning loop
                # Attempt to reconnect
                app_instance.after(1000, lambda: app_instance.connect_instrument_logic(app_instance)) # Try to reconnect after 1 sec
                break # Exit scan loop

            except Exception as e:
                print(f"❌ Main scan thread encountered an error: {e}")
                messagebox.showerror("Scan Error", f"An unexpected error occurred during scan: {e}")
                app_instance.after(0, app_instance._reset_gui_on_disconnect_or_error) # Reset GUI on main thread
                app_instance.scanning = False # Stop scanning loop
                break # Exit scan loop

            # Wait for the next cycle, respecting pause
            if app_instance.scanning and app_instance.cycle_wait_time_seconds_var.get() > 0:
                print(f"Waiting for {app_instance.cycle_wait_time_seconds_var.get()} seconds before next cycle...")
                for remaining in range(int(app_instance.cycle_wait_time_seconds_var.get()), 0, -1):
                    while app_instance.paused:
                        app_instance.after(100, app_instance._update_console_line, "Scan Paused. Click Resume to continue.", False)
                        time.sleep(0.1) # Sleep briefly while paused
                        if not app_instance.scanning: # Allow stopping even when paused
                            print("Scan interrupted during wait time.")
                            break
                    if not app_instance.scanning:
                        break # Exit inner loop if scanning stopped
                    app_instance.after(0, app_instance._update_console_line, f"Next cycle in {remaining} seconds...\r", True)
                    time.sleep(1)
                app_instance.after(0, app_instance._update_console_line, "                                               \r", True) # Clear line
            
            if not app_instance.scanning: # Check again after wait time
                print("Scan loop terminated.")
                break

    finally:
        app_instance.scanning = False
        app_instance.paused = False
        app_instance.after(100, reset_scan_buttons_logic, app_instance)
        if not app_instance.inst and app_instance.instrument_list and app_instance.resource_var.get() != "No resources found":
            app_instance._start_connect_button_blink()
        app_instance._stop_pause_button_blink() # Ensure blinking stops on scan end/stop

def stop_scan_logic(app_instance):
    app_instance.scanning = False
    app_instance.paused = False
    print("Attempting to stop scan... Please wait for current sweep to finish.")
    app_instance.stop_scan_button.config(state=tk.DISABLED)
    app_instance.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED)

def pause_resume_scan_logic(app_instance):
    if app_instance.scanning:
        app_instance.paused = not app_instance.paused
        if app_instance.paused:
            app_instance.pause_resume_button.config(text="Resume Scan")
            app_instance.after(100, app_instance._update_console_line, "Scan Paused. Click Resume to continue.", False)
        else:
            app_instance.pause_resume_button.config(text="Pause Scan")
            app_instance.after(100, app_instance._update_console_line, "Scan Resumed.", True)
    else:
        messagebox.showwarning("No Scan in Progress", "No scan is currently running to pause or resume.")


def reset_scan_buttons_logic(app_instance):
    app_instance.start_scan_button.config(state=tk.NORMAL)
    if app_instance.inst:
        app_instance.disconnect_button.config(state=tk.NORMAL)
        app_instance.apply_button.config(state=tk.NORMAL)
        if app_instance.preset_tree.selection() and app_instance.instrument_model != "N9340B":
            app_instance.load_preset_button.config(state=tk.NORMAL)
    else:
        app_instance.disconnect_button.config(state=tk.DISABLED)
        app_instance.apply_button.config(state=tk.DISABLED)
        app_instance.load_preset_button.config(state=tk.DISABLED)
    app_instance.stop_scan_button.config(state=tk.DISABLED)
    app_instance.pause_resume_button.config(state=tk.DISABLED)
    app_instance.plot_button.config(state=tk.NORMAL) # Re-enable generate plot button
