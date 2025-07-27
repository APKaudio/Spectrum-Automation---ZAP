# src/scan_logic.py
import tkinter as tk
from tkinter import messagebox # Ensure messagebox is directly imported
import threading
import time
import os
from datetime import datetime # Correct import: from datetime import datetime
import pandas as pd
import pyvisa # Ensure pyvisa is imported
import inspect # FIX: Import inspect module

from utils.scan_instrument import scan_bands
from utils.frequency_bands import MHZ_TO_HZ # Assuming MHZ_TO_HZ is needed here
from utils.instrument_control import log_visa_command, debug_print # Import debug_print
from src.plot_logic import generate_single_scan_plot_and_open_wrapper_logic # FIX: Import the plotting logic

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
    app_instance.load_preset_button.config(state=tk.DISABLED) # Disable load preset button during scan

    # Disable query presets button
    if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
        app_instance.preset_files_tab.query_presets_button.config(state=tk.DISABLED)

    app_instance.scanning = True
    app_instance.paused = False
    app_instance.stop_event.clear()
    app_instance.pause_event.clear()

    # Clear previous scan dataframes for a new scan
    app_instance.collected_scans_dataframes = []
    app_instance.current_scan_cycle_count = 0 # Initialize here for the first cycle

    # Get current settings from Tkinter variables
    # Ensure these are retrieved as floats where necessary
    # Use .get() on the Tkinter variables
    rbw_val = float(app_instance.desired_rbw_var.get())
    cycle_wait_time_val = float(app_instance.desired_cycle_wait_time_var.get())
    maxhold_time_val = float(app_instance.desired_max_hold_time_var.get())
    reference_level_val = float(app_instance.desired_reference_level_var.get())
    freq_shift_val = float(app_instance.desired_freq_shift_var.get())
    maxhold_enabled_val = app_instance.desired_maxhold_enabled_var.get()
    high_sensitivity_val = app_instance.desired_high_sensitivity_var.get()
    preamp_on_val = app_instance.desired_preamp_on_var.get()
    scan_rbw_segmentation_val = float(app_instance.desired_scan_rbw_segmentation_var.get())
    
    scan_name = app_instance.scan_name_var.get()
    output_folder = app_instance.output_folder_var.get()

    # Get selected bands
    selected_bands = [item["band"] for item in app_instance.band_vars if item["var"].get()]
    if not selected_bands:
        messagebox.showwarning("No Bands Selected", "Please select at least one frequency band to scan.")
        stop_scan_logic(app_instance) # Reset buttons if no bands selected
        return

    # Start the scan in a new thread
    scan_thread = threading.Thread(target=run_scan_logic, args=(
        app_instance, rbw_val, cycle_wait_time_val, maxhold_time_val,
        reference_level_val, freq_shift_val, maxhold_enabled_val,
        high_sensitivity_val, preamp_on_val, scan_rbw_segmentation_val,
        scan_name, output_folder, selected_bands
    ))
    scan_thread.daemon = True # Allow the thread to exit with the main program
    scan_thread.start()
    app_instance._update_console_line("Scan started in background...")


def run_scan_logic(app_instance, rbw_val, cycle_wait_time_val, maxhold_time_val,
                   reference_level_val, freq_shift_val, maxhold_enabled_val,
                   high_sensitivity_val, preamp_on_val, scan_rbw_segmentation_val,
                   scan_name, output_folder, selected_bands):
    """
    Executes the main scan logic in a separate thread.
    """
    try:
        # Initial setup for the scan (e.g., setting instrument to sweep mode)
        # This part should be handled by apply_settings_to_device_logic if not already done
        # or specific commands here if they are part of scan initiation.
        # For now, rely on apply_settings_to_device_logic being called before scan.

        while app_instance.scanning and not app_instance.stop_event.is_set():
            if app_instance.paused:
                app_instance.pause_event.wait() # Wait until unpaused
                app_instance.pause_event.clear() # Clear event after resuming
                if app_instance.stop_event.is_set(): # Check stop event again after pause
                    break

            app_instance.current_scan_cycle_count += 1
            # Removed 'True' argument as _update_console_line no longer takes 'overwrite'
            app_instance._update_console_line(f"Starting scan cycle {app_instance.current_scan_cycle_count}...")

            # Call the scan_bands function from scan_instrument.py
            last_successful_band_index, current_cycle_csv_filename = scan_bands(
                app_instance, # Pass app_instance_ref
                app_instance.inst, app_instance.stop_event, app_instance.pause_event,
                app_instance.instrument_model, rbw_val, cycle_wait_time_val,
                maxhold_time_val, reference_level_val, freq_shift_val,
                maxhold_enabled_val, high_sensitivity_val, preamp_on_val,
                scan_rbw_segmentation_val, scan_name, output_folder,
                selected_bands, app_instance.current_scan_cycle_count # FIX: Pass current_scan_cycle_count
            )

            if app_instance.stop_event.is_set():
                # Removed 'False' argument
                app_instance._update_console_line("Scan stopped by user.")
                break

            if current_cycle_csv_filename:
                # Load the current cycle's data into a DataFrame and add to list
                try:
                    df = pd.read_csv(current_cycle_csv_filename, header=None, names=['Frequency_MHz', 'Power_dBm'])
                    app_instance.collected_scans_dataframes.append(df)
                    # Removed 'True' argument
                    app_instance._update_console_line(f"Cycle {app_instance.current_scan_cycle_count} data collected. Total dataframes: {len(app_instance.collected_scans_dataframes)}")

                    # Generate single scan plot after each cycle if desired
                    if app_instance.open_html_after_complete_var.get():
                        # FIX: Corrected datetime.datetime.now() to datetime.now()
                        plot_html_output_path = os.path.join(output_folder, f"{scan_name}_Cycle{app_instance.current_scan_cycle_count}_Plot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
                        # Use app_instance.after to schedule plot generation on the main thread
                        app_instance.after(100, generate_single_scan_plot_and_open_wrapper_logic,
                                           app_instance, current_cycle_csv_filename, plot_html_output_path, True)

                except Exception as e:
                    # Removed 'False' argument
                    app_instance._update_console_line(f"❌ Error processing CSV for cycle {app_instance.current_scan_cycle_count}: {e}")
                    debug_print(f"Error processing CSV for cycle {app_instance.current_scan_cycle_count}: {e}", file=__file__, function=inspect.currentframe().f_code.co_name)
                    # FIX: Schedule messagebox.showerror on the main thread
                    app_instance.after(100, messagebox.showerror, "Error Processing CSV", f"Error processing CSV for cycle {app_instance.current_scan_cycle_count}: {e}")
                    break # Stop scan on critical error

            else:
                # Removed 'False' argument
                app_instance._update_console_line(f"🚫 No data collected for cycle {app_instance.current_scan_cycle_count}.")
                # If a band failed, we might want to continue to the next or stop.
                # For now, let's continue if no data was collected for a band, but log it.

            time.sleep(cycle_wait_time_val) # Wait before next cycle

    except pyvisa.errors.VisaIOError as e:
        # Removed 'False' argument
        app_instance._update_console_line(f"❌ VISA I/O Error during scan: {e}")
        # FIX: Schedule messagebox.showerror on the main thread
        app_instance.after(100, messagebox.showerror, "VISA Error", f"A VISA I/O error occurred during scan: {e}")
        debug_print(f"VISA I/O Error during scan: {e}", file=__file__, function=inspect.currentframe().f_code.co_name)
    except Exception as e:
        # Removed 'False' argument
        app_instance._update_console_line(f"❌ An unexpected error occurred during scan: {e}")
        # FIX: Schedule messagebox.showerror on the main thread
        app_instance.after(100, messagebox.showerror, "Scan Error", f"An unexpected error occurred during scan: {e}")
        debug_print(f"Unexpected error during scan: {e}", file=__file__, function=inspect.currentframe().f_code.co_name)
    finally:
        # Ensure GUI elements are reset after scan completes or stops due to error
        app_instance.after(100, reset_scan_buttons_logic, app_instance)
        app_instance.after(100, app_instance._stop_pause_button_blink)
        app_instance.scanning = False # Ensure scanning flag is reset
        app_instance.paused = False # Ensure paused flag is reset
        app_instance.stop_event.clear() # Clear stop event
        app_instance.pause_event.clear() # Clear pause event

        # Enable plot button if data was collected
        if app_instance.collected_scans_dataframes:
            # FIX: Correctly pass arguments to config method
            app_instance.after(100, lambda: app_instance.plot_button.config(state=tk.NORMAL))
            # Removed 'False' argument
            app_instance._update_console_line("Scan finished. Plot button enabled.")
        else:
            # Removed 'False' argument
            app_instance._update_console_line("Scan finished. No data collected for plotting.")


def stop_scan_logic(app_instance):
    """
    Signals the running scan thread to stop.
    """
    if app_instance.scanning:
        app_instance.stop_event.set()
        app_instance.pause_event.set() # Also set pause event to unblock if paused
        # Removed 'False' argument
        app_instance._update_console_line("Stopping scan...")
    else:
        messagebox.showwarning("No Scan in Progress", "No scan is currently running to stop.")

def pause_resume_scan_logic(app_instance):
    if app_instance.scanning:
        app_instance.paused = not app_instance.paused
        if app_instance.paused:
            app_instance.pause_resume_button.config(text="Resume Scan")
            app_instance._start_pause_button_blink() # Start blinking when paused
            # Removed 'False' argument
            app_instance.after(100, app_instance._update_console_line, "Scan Paused. Click Resume to continue.")
        else:
            app_instance.pause_event.set() # Signal to resume
            app_instance.pause_resume_button.config(text="Pause Scan")
            app_instance._stop_pause_button_blink() # Stop blinking when resumed
            # Removed 'True' argument
            app_instance.after(100, app_instance._update_console_line, "Scan Resumed.")
    else:
        messagebox.showwarning("No Scan in Progress", "No scan is currently running to pause or resume.")


def reset_scan_buttons_logic(app_instance):
    app_instance.start_scan_button.config(state=tk.NORMAL)
    app_instance.stop_scan_button.config(state=tk.DISABLED)
    app_instance.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED) # Reset text and state
    if app_instance.inst:
        app_instance.disconnect_button.config(state=tk.NORMAL)
        app_instance.apply_button.config(state=tk.NORMAL)
        if app_instance.preset_files_tab.get_selected_preset() and app_instance.instrument_model != "N9340B":
            app_instance.load_preset_button.config(state=tk.NORMAL)
        # Enable query presets button
        if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
            app_instance.preset_files_tab.query_presets_button.config(state=tk.NORMAL)
    else:
        app_instance.disconnect_button.config(state=tk.DISABLED)
        app_instance.apply_button.config(state=tk.DISABLED)
        app_instance.load_preset_button.config(state=tk.DISABLED)
        # Disable query presets button
        if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
            app_instance.preset_files_tab.query_presets_button.config(state=tk.DISABLED)

