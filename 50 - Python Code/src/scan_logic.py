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

# Import specific plotting functions directly from their source files
# Import _open_plot_in_browser directly to avoid NameError
from src.plot_logic import plot_single_scan_data, _open_plot_in_browser 
from utils.averaging_utils import generate_current_cycle_average_csv_and_plot as generate_average_plot_logic # Import from averaging_utils and alias


def start_scan_thread_logic(app_instance):
    if not app_instance.inst:
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: messagebox.showwarning("Not Connected", "Please connect to an instrument first."))
        return
    
    if app_instance.scanning:
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: messagebox.showwarning("Scan in Progress", "A scan is already running."))
        return

    # Save configuration when scan starts - this will save current GUI settings as LAST_USED
    from src.config_manager import save_config # Import locally to avoid circular dependency
    save_config(app_instance)

    # Disable buttons at the start of the scan
    app_instance.start_scan_button.config(state=tk.DISABLED)
    app_instance.connect_button.config(state=tk.DISABLED)
    app_instance.disconnect_button.config(state=tk.DISABLED)
    app_instance.apply_button.config(state=tk.DISABLED)
    app_instance.load_preset_button.config(state=tk.DISABLED) # Disable load preset button
    app_instance.plot_button.config(state=tk.DISABLED) # Disable plot button during scan

    # Disable query presets button
    if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
        app_instance.preset_files_tab.query_presets_button.config(state=tk.DISABLED)

    app_instance.stop_scan_button.config(state=tk.NORMAL)
    app_instance.pause_resume_button.config(state=tk.NORMAL)
    app_instance.pause_resume_button.config(text="Pause Scan") # Ensure text is "Pause Scan"

    app_instance.scanning = True
    app_instance.paused = False
    app_instance.stop_event.clear()
    app_instance.pause_event.clear() # Ensure pause event is clear at start

    # Clear previous scan data
    app_instance.collected_scans_dataframes = []
    app_instance.current_scan_cycle_count = 0

    # Start the scan in a new thread
    scan_thread = threading.Thread(target=run_scan_logic, args=(app_instance,))
    scan_thread.daemon = True # Allow the application to exit even if thread is running
    scan_thread.start()
    print("Scan initiated in background thread.")


def run_scan_logic(app_instance):
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    try:
        num_cycles = app_instance.num_scan_cycles_var.get()
        scan_name = app_instance.scan_name_var.get()
        output_folder = app_instance.output_folder_var.get()

        # Ensure output directory exists
        os.makedirs(output_folder, exist_ok=True)

        selected_bands_for_scan = [
            item["band"] for item in app_instance.band_vars if item["var"].get()
        ]

        if not selected_bands_for_scan:
            app_instance.after(0, lambda: messagebox.showwarning("No Bands Selected", "Please select at least one frequency band to scan."))
            app_instance.scanning = False
            return # Exit the scan thread

        # Get current settings from Tkinter variables
        scan_rbw_hz = float(app_instance.desired_rbw_var.get())
        maxhold_time_seconds = float(app_instance.desired_max_hold_time_var.get())
        cycle_wait_time_seconds = float(app_instance.desired_cycle_wait_time_var.get())
        reference_level_dbm = float(app_instance.desired_reference_level_var.get())
        freq_shift_hz = float(app_instance.desired_freq_shift_var.get())
        maxhold_enabled = app_instance.desired_maxhold_enabled_var.get()
        high_sensitivity = app_instance.desired_high_sensitivity_var.get()
        preamp_on = app_instance.desired_preamp_on_var.get()
        scan_rbw_segmentation = float(app_instance.desired_scan_rbw_segmentation_var.get())
        default_focus_width = float(app_instance.desired_default_focus_width_var.get())
        include_gov_markers = app_instance.include_gov_markers_var.get()
        include_tv_markers = app_instance.include_tv_markers_var.get()
        open_html_after_complete = app_instance.open_html_after_complete_var.get()
        include_markers_from_csv = app_instance.include_markers_var.get() # Get state of new checkbox


        app_instance.after(0, app_instance._update_console_line, f"Starting scan for {num_cycles} cycles...\n")
        app_instance.after(0, app_instance._update_console_line, f"Scan Name: {scan_name}\n")
        app_instance.after(0, app_instance._update_console_line, f"Output Folder: {output_folder}\n")
        app_instance.after(0, app_instance._update_console_line, f"Selected Bands: {[b['Band Name'] for b in selected_bands_for_scan]}\n")
        app_instance.after(0, app_instance._update_console_line, f"RBW: {scan_rbw_hz} Hz, Max Hold Time: {maxhold_time_seconds}s, Cycle Wait: {cycle_wait_time_seconds}s\n")
        app_instance.after(0, app_instance._update_console_line, f"Ref Level: {reference_level_dbm} dBm, Freq Shift: {freq_shift_hz} Hz\n")
        app_instance.after(0, app_instance._update_console_line, f"Max Hold Enabled: {maxhold_enabled}, High Sensitivity: {high_sensitivity}, Preamp ON: {preamp_on}\n")
        app_instance.after(0, app_instance._update_console_line, f"RBW Segmentation: {scan_rbw_segmentation} Hz, Default Focus Width: {default_focus_width} Hz\n")
        app_instance.after(0, app_instance._update_console_line, f"Include Gov Markers: {include_gov_markers}, Include TV Markers: {include_tv_markers}\n")
        app_instance.after(0, app_instance._update_console_line, f"Open HTML After Complete: {open_html_after_complete}\n")
        app_instance.after(0, app_instance._update_console_line, f"Include Custom Markers (MARKERS.CSV): {include_markers_from_csv}\n") # Log new setting


        for cycle in range(num_cycles):
            if app_instance.stop_event.is_set():
                app_instance.after(0, app_instance._update_console_line, "\nScan stopped by user.\n")
                break

            # Pause logic
            if app_instance.paused:
                app_instance.after(0, app_instance._update_console_line, f"\nScan paused. Waiting to resume...\n")
                app_instance.after(0, app_instance._start_pause_button_blink)
                app_instance.pause_event.wait() # Wait until resume is signaled
                app_instance.after(0, app_instance._stop_pause_button_blink)
                if app_instance.stop_event.is_set(): # Check stop event again after resume
                    app_instance.after(0, app_instance._update_console_line, "\nScan stopped by user during pause.\n")
                    break
                app_instance.after(0, app_instance._update_console_line, "\nScan resumed.\n")
                app_instance.paused = False # Reset paused flag

            app_instance.current_scan_cycle_count = cycle + 1
            app_instance.after(0, app_instance._update_console_line, f"\n--- Starting Cycle {cycle + 1}/{num_cycles} ---\n")

            # Perform the scan for the selected bands
            # scan_bands returns (last_successful_band_index, output_csv_filename)
            last_successful_band_index, current_cycle_csv_filename = scan_bands(
                app_instance, # app_instance_ref
                app_instance.inst, # inst
                app_instance.stop_event, # stop_event
                app_instance.pause_event, # pause_event
                app_instance.instrument_model, # instrument_model
                scan_rbw_hz, # rbw_val
                cycle_wait_time_seconds, # cycle_wait_time_val
                maxhold_time_seconds, # maxhold_time_val
                reference_level_dbm, # reference_level_val
                freq_shift_hz, # freq_shift_val
                maxhold_enabled, # maxhold_enabled_val
                high_sensitivity, # high_sensitivity_val
                preamp_on, # preamp_on_val
                scan_rbw_segmentation, # scan_rbw_segmentation_val
                scan_name, # scan_name
                output_folder, # output_folder
                selected_bands_for_scan, # selected_bands
                app_instance.current_scan_cycle_count, # current_scan_cycle_count
                file=current_file, # Explicitly pass as keyword
                function=current_function # Explicitly pass as keyword
            )

            if app_instance.stop_event.is_set():
                app_instance.after(0, app_instance._update_console_line, "\nScan stopped by user.\n")
                break

            if current_cycle_csv_filename:
                app_instance.after(0, app_instance._update_console_line, f"Cycle {cycle + 1} data saved to: {current_cycle_csv_filename}\n")
                debug_print(f"Attempting to read CSV for plotting: {current_cycle_csv_filename}", file=current_file, function=current_function)
                
                # Load the CSV into a DataFrame for plotting and aggregation
                try:
                    # Read CSV without header, so columns are 0 and 1
                    # Explicitly name columns during read to prevent pandas from inferring headers
                    df_to_plot = pd.read_csv(current_cycle_csv_filename, header=None, names=['Frequency (MHz)', 'Amplitude (dBm)'])
                    debug_print(f"CSV read successfully. Initial DataFrame columns: {df_to_plot.columns.tolist()}", file=current_file, function=current_function)
                    debug_print(f"Initial DataFrame head:\n{df_to_plot.head()}", file=current_file, function=current_function)

                    app_instance.collected_scans_dataframes.append(df_to_plot)

                    # Generate a single plot for the current scan cycle
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    single_scan_html_filename = os.path.join(
                        output_folder,
                        f"{scan_name}_Cycle{cycle + 1}_{timestamp}_SingleScan.html"
                    )

                    # Determine the path for MARKERS.CSV to be in the output folder
                    markers_csv_path = os.path.join(output_folder, "MARKERS.CSV")
                    debug_print(f"Checking for MARKERS.CSV at: {markers_csv_path}", file=current_file, function=current_function)


                    fig, plot_html_path_return = plot_single_scan_data(
                        df_to_plot,
                        f"{scan_name} - Cycle {cycle + 1} Single Scan",
                        single_scan_html_filename,
                        include_tv_markers=include_tv_markers,
                        include_gov_markers=include_gov_markers,
                        include_markers_from_csv=include_markers_from_csv, # Pass the new variable
                        markers_csv_path=markers_csv_path # Pass the path
                    )

                    if fig and open_html_after_complete:
                        # Use the directly imported _open_plot_in_browser function
                        app_instance.after(0, lambda path=plot_html_path_return: _open_plot_in_browser(path))

                except Exception as e:
                    app_instance.after(0, app_instance._update_console_line, f"❌ Error processing or plotting scan data for cycle {cycle + 1}: {e}\n")
                    debug_print(f"Error in scan_logic plotting: {e}", file=current_file, function=current_function)
            else:
                app_instance.after(0, app_instance._update_console_line, f"🚫 No data collected for cycle {cycle + 1}.\n")

            if cycle < num_cycles - 1:
                app_instance.after(0, app_instance._update_console_line, f"Waiting for {cycle_wait_time_seconds} seconds before next cycle...\n")
                time.sleep(cycle_wait_time_seconds)

    except Exception as e:
        app_instance.after(0, app_instance._update_console_line, f"❌ An error occurred during scan: {e}\n")
        debug_print(f"Error in run_scan_logic: {e}", file=current_file, function=current_function)
    finally:
        app_instance.scanning = False
        app_instance.after(0, app_instance._update_console_line, "\nScan process finished.\n")
        app_instance.after(0, lambda: reset_scan_buttons_logic(app_instance)) # Reset buttons on main thread
        # Enable plot button if any data was collected
        if app_instance.collected_scans_dataframes:
            app_instance.after(0, lambda: app_instance.plot_button.config(state=tk.NORMAL))


def stop_scan_logic(app_instance):
    if app_instance.scanning:
        app_instance.stop_event.set()
        app_instance.pause_event.set() # Also set pause event to unblock if paused
        app_instance.after(0, app_instance._update_console_line, "Stop signal sent. Finishing current operation...\n")
    else:
        app_instance.after(0, lambda: messagebox.showwarning("No Scan in Progress", "No scan is currently running to stop."))


def pause_resume_scan_logic(app_instance):
    if app_instance.scanning:
        if app_instance.paused:
            app_instance.paused = False
            app_instance.pause_event.set() # Signal to resume
            app_instance.pause_resume_button.config(text="Pause Scan")
            app_instance.after(0, app_instance._stop_pause_button_blink)
            app_instance.after(0, app_instance._update_console_line, "Scan resume signal sent.\n")
        else:
            app_instance.paused = True
            app_instance.pause_event.clear() # Block the thread
            app_instance.pause_resume_button.config(text="Resume Scan")
            app_instance.after(0, app_instance._start_pause_button_blink)
            app_instance.after(0, app_instance._update_console_line, "Scan pause signal sent.\n")
    else:
        app_instance.after(0, lambda: messagebox.showwarning("No Scan in Progress", "No scan is currently running to pause/resume."))


def reset_scan_buttons_logic(app_instance):
    app_instance.start_scan_button.config(state=tk.NORMAL)
    app_instance.stop_scan_button.config(state=tk.DISABLED)
    app_instance.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED) # Reset text and state
    if app_instance.inst:
        app_instance.disconnect_button.config(state=tk.NORMAL)
        app_instance.apply_button.config(state=tk.NORMAL)
        if hasattr(app_instance, 'preset_files_tab') and app_instance.preset_files_tab.get_selected_preset() and app_instance.instrument_model != "N9340B":
            app_instance.load_preset_button.config(state=tk.NORMAL)
        # Enable query presets button
        if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
            app_instance.preset_files_tab.query_presets_button.config(state=tk.NORMAL)
    else:
        app_instance.disconnect_button.config(state=tk.DISABLED)
        app_instance.apply_button.config(state=tk.DISABLED)
        app_instance.load_preset_button.config(state=tk.DISABLED) # Disable load preset if no instrument
        # Disable query presets button if no instrument
        if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
            app_instance.preset_files_tab.query_presets_button.config(state=tk.DISABLED)
