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
    app_instance.load_preset_button.config(state=tk.DISABLED)
    
    # Clear previous scan data
    app_instance.collected_scans_dataframes = []
    app_instance.scan_cycle_count = 0
    app_instance.current_freq_offset = 0

    # Get selected bands from checkboxes
    selected_bands = [item["band"] for item in app_instance.band_vars if item["var"].get()]
    if not selected_bands:
        messagebox.showwarning("No Bands Selected", "Please select at least one frequency band to scan.")
        reset_scan_buttons_logic(app_instance)
        return

    scan_rbw_segmentation = app_instance.desired_scan_rbw_segmentation_var.get()
    freq_shift_value = app_instance.shift_freq_var.get()
    max_hold_time = app_instance.desired_max_hold_time_var.get()
    
    # Get RBW and VBW values for instrument configuration
    rbw_config_val = int(scan_rbw_segmentation)
    vbw_config_val = int(rbw_config_val / 3) # VBW is typically 1/3 of RBW

    # Start the scan in a new thread
    scan_thread = threading.Thread(target=app_instance._run_scan, args=(selected_bands, scan_rbw_segmentation, freq_shift_value, rbw_config_val, vbw_config_val, max_hold_time))
    scan_thread.daemon = True # Allow the thread to exit with the main application
    app_instance.scanning = True
    scan_thread.start()
    print("Scan thread started.")

def toggle_pause_scan_logic(app_instance):
    """
    Toggles the paused state of the scan.

    Inputs:
        app_instance (App): The main application instance.
    Process:
        1. If a scan is running, toggles `app_instance.paused`.
        2. Updates the "Pause/Resume Scan" button text and color.
        3. Prints a message to the console, overwriting the previous line.
    Outputs: None
    """
    if app_instance.scanning:
        if app_instance.paused:
            app_instance.paused = False
            app_instance.after(0, app_instance.pause_resume_button.config, text="Pause Scan", bg="orange")
            app_instance.after(0, app_instance._update_console_line, "Scan Resumed.\n", overwrite=True) # Overwrite
        else:
            app_instance.paused = True
            app_instance.after(0, app_instance.pause_resume_button.config, text="Resume Scan", bg="blue")
            app_instance.after(0, app_instance._update_console_line, "Scan Paused. Click Resume to continue.\n", overwrite=True) # Overwrite

def run_scan_logic(app_instance, selected_bands, scan_rbw_segmentation, freq_shift_value, rbw_config_val, vbw_config_val, max_hold_time):
    """
    Executes the main scan loop. This function runs in a separate thread.

    Inputs:
        app_instance (App): The main application instance.
        selected_bands (list): List of dictionaries, each defining a frequency band.
        scan_rbw_segmentation (float): The RBW to use for each scan segment.
        freq_shift_value (float): The frequency offset (in Hz) to apply per scan cycle.
        rbw_config_val (int): RBW value to configure on the instrument.
        vbw_config_val (int): VBW value to configure on the instrument.
        max_hold_time (float): Duration in seconds for which MAX Hold should be active (0 if disabled).
    Process:
        1. Loops as long as `app_instance.scanning` is True.
        2. Handles pausing: waits while `app_instance.paused` is True.
        3. Calculates current frequency offset and updates `app_instance.current_freq_offset`.
        4. Calls `scan_bands` from `utils.scan_instrument` to perform the actual scan.
        5. Appends collected data to `app_instance.collected_scans_dataframes`.
        6. Updates scan cycle count.
        7. Generates and opens a plot after each cycle if `open_html_after_complete_var` is True.
        8. Handles exceptions during scanning, showing error messages.
        9. Ensures buttons are reset after the scan completes or is stopped.
    Outputs: None
    """
    try:
        while app_instance.scanning:
            # Handle pause state
            while app_instance.paused:
                # Use _update_console_line with overwrite to prevent spam
                app_instance.after(0, app_instance._update_console_line, "Scan Paused. Click Resume to continue.\n", overwrite=True)
                time.sleep(0.1) # Sleep briefly while paused
                if not app_instance.scanning: # If stop button pressed while paused
                    app_instance.after(0, app_instance._update_console_line, "\nScan process finished (interrupted).\n", overwrite=False) # New line for final message
                    break
            
            if not app_instance.scanning: # Check again after pause loop
                app_instance.after(0, app_instance._update_console_line, "\nScan process finished (interrupted).\n", overwrite=False) # New line for final message
                break

            app_instance.scan_cycle_count += 1
            app_instance.current_freq_offset = app_instance.scan_cycle_count * freq_shift_value
            
            app_instance._update_console_line(f"--- Starting Scan Cycle {app_instance.scan_cycle_count} (Offset: {app_instance.current_freq_offset} Hz) ---", overwrite=True)

            # Perform the scan for the selected bands
            current_sweep_data, output_csv_filename = scan_bands(
                app_instance, # Pass app_instance
                app_instance.inst,
                selected_bands,
                app_instance.output_folder_var.get(),
                app_instance.scan_name_var.get(),
                app_instance.current_freq_offset,
                rbw_config_val,
                vbw_config_val,
                max_hold_time
            )

            if current_sweep_data:
                # Convert raw data to DataFrame and store
                df = pd.DataFrame(current_sweep_data, columns=["Frequency (MHz)", "Level (dBm)"])
                app_instance.collected_scans_dataframes.append({
                    "df": df,
                    "cycle_count": app_instance.scan_cycle_count,
                    "offset_hz": app_instance.current_freq_offset,
                    "timestamp": datetime.now()
                })
                app_instance.last_scan_data = df # Update last_scan_data for plotting
                app_instance._update_console_line(f"✅ Data collected for cycle {app_instance.scan_cycle_count}. Saved to: {output_csv_filename}\n", overwrite=False) # Ensure newline for final message

                if app_instance.open_html_after_complete_var.get():
                    plot_title = f"{app_instance.scan_name_var.get()} - Cycle {app_instance.scan_cycle_count} (Offset: {app_instance.current_freq_offset} Hz)"
                    app_instance.after(0, app_instance.generate_single_scan_plot_and_open_wrapper,
                                       os.path.join(app_instance.output_folder_var.get(), output_csv_filename),
                                       plot_title,
                                       os.path.join(app_instance.output_folder_var.get(), f"{os.path.splitext(output_csv_filename)[0]}.html"),
                                       True) # auto_open_browser=True
            else:
                app_instance._update_console_line(f"🚫 No data collected for scan cycle {app_instance.scan_cycle_count}.\n", overwrite=False) # Ensure newline

            # Wait for the specified cycle wait time before the next scan
            wait_time = app_instance.desired_cycle_wait_time_var.get()
            if wait_time > 0:
                app_instance._update_console_line(f"Waiting for {wait_time} seconds before next cycle...", overwrite=True)
                for _ in range(int(wait_time * 10)): # Check every 0.1 seconds
                    while app_instance.paused:
                        app_instance.after(0, app_instance._update_console_line, "Scan Paused. Click Resume to continue.\n", overwrite=True)
                        time.sleep(0.1)
                        if not app_instance.scanning:
                            app_instance.after(0, app_instance._update_console_line, "\nScan process finished (interrupted during pause in wait).\n", overwrite=False)
                            break
                    
                    if not app_instance.scanning:
                        app_instance.after(0, app_instance._update_console_line, "\nScan process finished (interrupted during wait).\n", overwrite=False)
                        break
                    time.sleep(0.1)
                if app_instance.scanning: # Only if not interrupted during wait
                    app_instance.after(0, app_instance._update_console_line, f"Finished waiting {wait_time} seconds.\n", overwrite=True)


    except pyvisa.errors.VisaIOError as e:
        app_instance.after(0, messagebox.showerror, "VISA Scan Error", f"A VISA communication error occurred during scan: {e}")
        print(f"❌ VISA Scan Error: {e}")
    except Exception as e:
        app_instance.after(0, messagebox.showerror, "Scan Thread Error", f"An unexpected error occurred in main scan thread: {e}")
        print(f"❌ Main scan thread encountered an error: {e}")
        print(f"Main scan thread error: {e}")
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
    print("\nAttempting to stop scan... Please wait for current sweep to finish.")
    app_instance.stop_scan_button.config(state=tk.DISABLED)
    app_instance.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED)

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
    app_instance.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED)
