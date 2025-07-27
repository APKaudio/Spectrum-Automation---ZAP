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
from src.plot_logic import plot_single_scan_data # Import plot_single_scan_data directly
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

    app_instance.start_scan_button.config(state=tk.DISABLED)
    app_instance.stop_scan_button.config(state=tk.NORMAL)
    app_instance.pause_resume_button.config(state=tk.NORMAL)

    app_instance.scanning = True
    app_instance.paused = False
    app_instance.stop_event.clear()
    app_instance.pause_event.clear()

    app_instance.current_scan_cycle_count = 0 # Reset cycle count for a new scan session
    app_instance.collected_scans_dataframes = [] # Clear collected dataframes for a new scan session

    app_instance.after(0, app_instance._update_console_line, "Scan started in background...\n")

    # Start the scan in a new thread
    app_instance.scan_thread = threading.Thread(
        target=run_scan_logic, # Changed target from _run_scan to run_scan_logic
        args=(app_instance,)
    )
    app_instance.scan_thread.daemon = True # Allow the program to exit even if thread is running
    app_instance.scan_thread.start()


def run_scan_logic(app_instance): # Renamed from _run_scan to run_scan_logic
    """
    Internal function to run the scan process in a separate thread.
    This function orchestrates the scanning, data collection, and plotting.
    """
    file = __file__
    function = inspect.currentframe().f_code.co_name

    try:
        # Get current settings from Tkinter variables
        # Assuming num_scan_cycles_var will be defined in main_app.py
        num_scan_cycles = app_instance.num_scan_cycles_var.get()
        rbw_val = app_instance.desired_rbw_var.get()
        cycle_wait_time_val = app_instance.desired_cycle_wait_time_var.get()
        maxhold_time_val = app_instance.desired_max_hold_time_var.get()
        reference_level_val = app_instance.desired_reference_level_var.get()
        freq_shift_val = app_instance.desired_freq_shift_var.get()
        maxhold_enabled_val = app_instance.desired_maxhold_enabled_var.get()
        high_sensitivity_val = app_instance.desired_high_sensitivity_var.get()
        preamp_on_val = app_instance.desired_preamp_on_var.get()
        scan_rbw_segmentation_val = app_instance.desired_scan_rbw_segmentation_var.get()
        scan_name = app_instance.scan_name_var.get()
        output_folder = app_instance.output_folder_var.get()
        selected_bands = [band["band"] for band in app_instance.band_vars if band["var"].get()]
        
        # Ensure output folder exists
        os.makedirs(output_folder, exist_ok=True)

        # Main scan loop (e.g., for multiple cycles)
        app_instance.after(0, app_instance._update_console_line, f"Total scan cycles configured: {num_scan_cycles}\n")

        for cycle in range(1, num_scan_cycles + 1):
            if app_instance.stop_event.is_set():
                app_instance.after(0, app_instance._update_console_line, f"Scan stopped by user after cycle {cycle-1}.\n")
                break

            app_instance.current_scan_cycle_count = cycle
            app_instance.after(0, app_instance._update_console_line, f"Starting scan cycle {cycle}...\n")

            # Call the scan_bands function from utils.scan_instrument
            last_successful_band_index, csv_filename_current_cycle = scan_bands(
                app_instance,
                app_instance.inst,
                app_instance.stop_event,
                app_instance.pause_event,
                app_instance.instrument_model,
                float(rbw_val), # Ensure these are passed as floats
                float(cycle_wait_time_val),
                float(maxhold_time_val),
                float(reference_level_val),
                float(freq_shift_val),
                bool(maxhold_enabled_val),
                bool(high_sensitivity_val),
                bool(preamp_on_val),
                float(scan_rbw_segmentation_val),
                scan_name,
                output_folder,
                selected_bands,
                cycle # Pass current cycle number
            )

            if csv_filename_current_cycle:
                # Load the CSV into a DataFrame and add to collected_scans_dataframes
                try:
                    df = pd.read_csv(csv_filename_current_cycle, header=None)
                    df.columns = ['Frequency_MHz', 'Power_dBm']
                    df['Frequency_MHz'] = pd.to_numeric(df['Frequency_MHz'], errors='coerce')
                    df['Power_dBm'] = pd.to_numeric(df['Power_dBm'], errors='coerce')
                    df.dropna(subset=['Frequency_MHz', 'Power_dBm'], inplace=True)
                    
                    if not df.empty:
                        app_instance.collected_scans_dataframes.append(df)
                        app_instance.after(0, app_instance._update_console_line, f"Cycle {cycle} data collected. Total dataframes: {len(app_instance.collected_scans_dataframes)}\n")
                    else:
                        app_instance.after(0, app_instance._update_console_line, f"Warning: Cycle {cycle} CSV was empty or contained no valid data after processing. Not added to collected data.\n")

                except Exception as e:
                    app_instance.after(0, app_instance._update_console_line, f"Error loading CSV for cycle {cycle}: {e}\n")
                    debug_print(f"Error loading CSV for cycle {cycle}: {e}", file=file, function=function)
                    # Schedule messagebox to run on the main thread
                    if app_instance.after(0, lambda: messagebox.askyesno("CSV Load Error", f"Error loading CSV for cycle {cycle}: {e}\nDo you want to continue with the next cycle?")):
                        continue
                    else:
                        break # Stop scan if user chooses not to continue

                # Generate and open plot for the current single scan cycle
                output_html_path_single = os.path.join(output_folder, f"{scan_name}_Cycle{cycle}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
                
                # Use the directly imported function: plot_single_scan_data
                app_instance.after(0, lambda: plot_single_scan_data(
                    csv_file_path=csv_filename_current_cycle, # Pass the CSV file path
                    output_html_path=output_html_path_single,
                    auto_open_browser=app_instance.open_html_after_complete_var.get(), # Pass auto_open_browser setting
                    include_tv_markers=app_instance.include_tv_markers_var.get(), # Pass TV markers setting
                    include_gov_markers=app_instance.include_gov_markers_var.get() # Pass Gov markers setting
                ))

            else:
                app_instance.after(0, app_instance._update_console_line, f"🚫 No CSV file generated for cycle {cycle}. Skipping plot generation for this cycle.\n")

            # Pause between cycles if not the last cycle
            if cycle < num_scan_cycles and not app_instance.stop_event.is_set():
                app_instance.after(0, app_instance._update_console_line, f"Waiting for {cycle_wait_time_val} seconds before next cycle...\n")
                time.sleep(float(cycle_wait_time_val)) # Ensure float conversion for time.sleep

        # After all cycles, generate the averaged plot if data was collected
        if app_instance.collected_scans_dataframes:
            # Use the directly imported and aliased function
            app_instance.after(0, lambda: generate_average_plot_logic(
                app_instance.collected_scans_dataframes,
                app_instance.scan_name_var,
                app_instance.output_folder_var,
                app_instance.open_html_after_complete_var,
                app_instance.include_tv_markers_var,
                app_instance.include_gov_markers_var
            ))
        else:
            app_instance.after(0, app_instance._update_console_line, "No data collected across all cycles to generate an average plot.\n")

    except pyvisa.errors.VisaIOError as e:
        app_instance.after(0, app_instance._update_console_line, f"🛑 VISA communication error during scan: {e}\n")
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: messagebox.showerror("VISA Error", f"A VISA communication error occurred during the scan: {e}"))
        debug_print(f"VISA Error in _run_scan: {e}", file=file, function=function)
    except Exception as e:
        app_instance.after(0, app_instance._update_console_line, f"❌ An unexpected error occurred during scan: {e}\n")
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: messagebox.showerror("Scan Error", f"An unexpected error occurred during the scan: {e}"))
        debug_print(f"Unexpected Error in _run_scan: {e}", file=file, function=function)
    finally:
        app_instance.scanning = False
        app_instance.paused = False
        app_instance.after(0, reset_scan_buttons_logic, app_instance)
        app_instance.after(0, app_instance._update_console_line, "Scan process finished.\n")


def pause_resume_scan_logic(app_instance):
    if app_instance.scanning:
        if app_instance.paused:
            app_instance.paused = False
            app_instance.pause_event.clear()
            app_instance.pause_resume_button.config(text="Pause Scan")
            app_instance.after(100, app_instance._update_console_line, "Scan Resumed.\n")
        else:
            app_instance.paused = True
            app_instance.pause_event.set()
            app_instance.pause_resume_button.config(text="Resume Scan")
            app_instance.after(100, app_instance._update_console_line, "Scan Paused.\n")
    else:
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: messagebox.showwarning("No Scan in Progress", "No scan is currently running to pause or resume."))


def stop_scan_logic(app_instance):
    if app_instance.scanning:
        app_instance.after(0, app_instance._update_console_line, "Stopping scan...\n")
        app_instance.stop_event.set() # Signal the thread to stop
        app_instance.pause_event.clear() # Clear pause in case it was paused
        app_instance.paused = False # Reset paused state
        # The thread will eventually terminate, and finally block will reset buttons
    else:
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: messagebox.showwarning("No Scan in Progress", "No scan is currently running to stop."))


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
