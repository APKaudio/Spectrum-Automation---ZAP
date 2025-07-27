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

    # Disable buttons during scan
    app_instance.start_scan_button.config(state=tk.DISABLED)
    app_instance.stop_scan_button.config(state=tk.NORMAL)
    app_instance.pause_resume_button.config(state=tk.NORMAL)
    app_instance.disconnect_button.config(state=tk.DISABLED)
    app_instance.apply_button.config(state=tk.DISABLED)
    app_instance.load_preset_button.config(state=tk.DISABLED)
    
    # Disable plot button when scan starts
    if hasattr(app_instance, 'plotting_tab') and hasattr(app_instance.plotting_tab, 'plot_button'):
        app_instance.plotting_tab.plot_button.config(state=tk.DISABLED)

    # Disable query presets button
    if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
        app_instance.preset_files_tab.query_presets_button.config(state=tk.DISABLED)

    app_instance.scanning = True
    app_instance.stop_event.clear()
    app_instance.pause_event.clear() # Ensure pause event is clear at start

    # Start the scan in a separate thread
    scan_thread = threading.Thread(target=run_scan_logic, args=(app_instance,))
    scan_thread.daemon = True # Allow the thread to exit with the main application
    scan_thread.start()
    print("🚀 Scan started...")


def run_scan_logic(app_instance):
    """
    Executes the spectrum scan logic in a separate thread.
    This function continuously scans selected frequency bands, collects data,
    and updates the GUI with progress.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    try:
        app_instance.collected_scans_dataframes = [] # Clear previous scan data
        app_instance.current_scan_cycle_count = 0
        num_scan_cycles = app_instance.num_scan_cycles_var.get()

        debug_print(f"Starting scan with {num_scan_cycles} cycles.", file=current_file, function=current_function)

        for cycle_num in range(1, num_scan_cycles + 1):
            if app_instance.stop_event.is_set():
                debug_print(f"Scan cycle {cycle_num} interrupted by stop event.", file=current_file, function=current_function)
                break # Exit loop if stop is requested

            app_instance.after(0, app_instance._update_console_line, f"\n--- Starting Scan Cycle {cycle_num}/{num_scan_cycles} ---\n")
            debug_print(f"Starting Scan Cycle {cycle_num}/{num_scan_cycles}", file=current_file, function=current_function)

            # Wait if paused
            while app_instance.paused:
                app_instance.after(0, app_instance._update_console_line, "Scan Paused. Click 'Pause Scan' to resume.\n")
                debug_print("Scan Paused.", file=current_file, function=current_function)
                app_instance.pause_event.wait(1.0) # Wait with a timeout to re-check stop_event
                if app_instance.stop_event.is_set():
                    debug_print("Scan stopped while paused.", file=current_file, function=current_function)
                    break # Exit pause if stop is requested

            if app_instance.stop_event.is_set():
                break # Exit loop if stop was requested while paused

            # Get current settings from Tkinter variables
            selected_bands = [item["band"] for item in app_instance.band_vars if item["var"].get()]
            scan_rbw_hz = float(app_instance.desired_rbw_var.get())
            ref_level_dbm = float(app_instance.desired_reference_level_var.get())
            freq_shift_hz = float(app_instance.desired_freq_shift_var.get())
            maxhold_enabled = app_instance.desired_maxhold_enabled_var.get()
            maxhold_time_seconds = float(app_instance.desired_max_hold_time_var.get())
            high_sensitivity = app_instance.desired_high_sensitivity_var.get()
            preamp_on = app_instance.desired_preamp_on_var.get()
            scan_rbw_segmentation = float(app_instance.desired_scan_rbw_segmentation_var.get())
            
            # Pass the _update_console_line method to scan_bands for real-time updates
            last_successful_band_index, current_scan_filename = scan_bands(
                app_instance.inst,
                selected_bands,
                scan_rbw_hz,
                ref_level_dbm,
                freq_shift_hz,
                maxhold_enabled,
                maxhold_time_seconds,
                high_sensitivity,
                preamp_on,
                scan_rbw_segmentation,
                app_instance.scan_name_var.get(),
                app_instance.output_folder_var.get(),
                app_instance.stop_event,
                app_instance.pause_event,
                app_instance._update_console_line, # Pass the callback here
                app_instance # Pass the app_instance itself for access to properties
            )

            if app_instance.stop_event.is_set():
                app_instance.after(0, app_instance._update_console_line, f"\n--- Scan Cycle {cycle_num} Interrupted ---\n")
                debug_print(f"Scan Cycle {cycle_num} Interrupted.", file=current_file, function=current_function)
                break
            
            if current_scan_filename:
                # Load the newly generated CSV into a DataFrame and add to collected_scans_dataframes
                try:
                    df = pd.read_csv(current_scan_filename)
                    app_instance.collected_scans_dataframes.append(df)
                    app_instance.current_scan_cycle_count = cycle_num # Update current cycle count
                    app_instance.after(0, app_instance._update_console_line, f"✅ Data for Cycle {cycle_num} loaded from {os.path.basename(current_scan_filename)}\n")
                    debug_print(f"Data for Cycle {cycle_num} loaded from {os.path.basename(current_scan_filename)}", file=current_file, function=current_function)

                    # Generate single scan plot for the current cycle
                    plot_title = f"{app_instance.scan_name_var.get()} - Cycle {cycle_num} - {datetime.now().strftime('%Y-%m-%d %H-%M-%S')}"
                    plot_output_path = os.path.join(app_instance.output_folder_var.get(), app_instance.scan_name_var.get(), f"plot_cycle_{cycle_num}.html")

                    # Use the app_instance's boolean vars for plotting options
                    fig, plot_html_path_return = plot_single_scan_data(
                        df,
                        plot_title,
                        app_instance.include_tv_markers_var.get(),
                        app_instance.include_gov_markers_var.get(),
                        app_instance.include_markers_var.get(), # Pass custom markers flag
                        app_instance.output_folder_var.get(), # Pass output folder for markers.csv lookup
                        output_html_path=plot_output_path
                    )
                    if fig:
                        app_instance.after(0, app_instance._update_console_line, f"✅ Plot for Cycle {cycle_num} saved to: {plot_html_path_return}\n")
                        debug_print(f"Plot for Cycle {cycle_num} saved to: {plot_html_path_return}", file=current_file, function=current_function)
                        if app_instance.open_html_after_complete_var.get():
                            app_instance.after(0, lambda p=plot_html_path_return: _open_plot_in_browser(p))
                    else:
                        app_instance.after(0, app_instance._update_console_line, f"🚫 Plot for Cycle {cycle_num} was not generated.\n")
                        debug_print(f"Plot for Cycle {cycle_num} was not generated.", file=current_file, function=current_function)

                except Exception as e:
                    app_instance.after(0, app_instance._update_console_line, f"❌ Error processing data for Cycle {cycle_num}: {e}\n")
                    debug_print(f"Error processing data for Cycle {cycle_num}: {e}", file=current_file, function=current_function)
                    # Decide if you want to stop the scan on error or continue
                    # For now, let's continue to the next cycle
            else:
                app_instance.after(0, app_instance._update_console_line, f"🚫 No data collected for Cycle {cycle_num}.\n")
                debug_print(f"No data collected for Cycle {cycle_num}.", file=current_file, function=current_function)

            if cycle_num < num_scan_cycles and not app_instance.stop_event.is_set():
                wait_time = float(app_instance.desired_cycle_wait_time_var.get())
                app_instance.after(0, app_instance._update_console_line, f"Waiting {wait_time} seconds before next cycle...\n")
                debug_print(f"Waiting {wait_time} seconds before next cycle.", file=current_file, function=current_function)
                time.sleep(wait_time) # Wait between cycles

    except pyvisa.errors.VisaIOError as e:
        app_instance.after(0, lambda: messagebox.showerror("VISA Error", f"A VISA communication error occurred during scan: {e}"))
        app_instance.after(0, app_instance._update_console_line, f"❌ VISA error during scan: {e}\n")
        debug_print(f"VISA error during scan: {e}", file=current_file, function=current_function)
    except Exception as e:
        app_instance.after(0, lambda: messagebox.showerror("Scan Error", f"An unexpected error occurred during scan: {e}"))
        app_instance.after(0, app_instance._update_console_line, f"❌ Unexpected error during scan: {e}\n")
        debug_print(f"Unexpected error during scan: {e}", file=current_file, function=current_function)
    finally:
        app_instance.scanning = False
        app_instance.stop_event.clear()
        app_instance.pause_event.clear()
        app_instance.after(0, app_instance._stop_pause_button_blink) # Ensure blinking stops
        app_instance.after(0, lambda: reset_scan_buttons_logic(app_instance)) # Reset buttons on main thread

        # After scan completes (or stops), generate the average plot if data exists
        if app_instance.collected_scans_dataframes:
            app_instance.after(0, app_instance._update_console_line, "\n--- Generating Final Averaged Plot and CSVs ---\n")
            debug_print("Generating Final Averaged Plot and CSVs.", file=current_file, function=current_function)
            try:
                # Call the generate_current_cycle_average_csv_and_plot function with current app instance variables
                generate_average_plot_logic(
                    app_instance.collected_scans_dataframes,
                    app_instance.scan_name_var,
                    app_instance.output_folder_var,
                    app_instance.include_tv_markers_var,
                    app_instance.include_gov_markers_var,
                    app_instance.include_markers_var, # Pass custom markers flag
                    app_instance.open_html_after_complete_var
                )
                app_instance.after(0, app_instance._update_console_line, "✅ Final averaged plot and CSVs generated.\n")
                debug_print("Final averaged plot and CSVs generated.", file=current_file, function=current_function)
            except Exception as e:
                app_instance.after(0, app_instance._update_console_line, f"❌ Error generating final averaged plot/CSVs: {e}\n")
                debug_print(f"Error generating final averaged plot/CSVs: {e}", file=current_file, function=current_function)
        else:
            app_instance.after(0, app_instance._update_console_line, "🚫 No scan data collected for final plot generation.\n")
            debug_print("No scan data collected for final plot generation.", file=current_file, function=current_function)

        app_instance.after(0, app_instance._update_console_line, "--- Scan process finished. ---\n")
        debug_print("Scan process finished.", file=current_file, function=current_function)


def stop_scan_logic(app_instance):
    """
    Signals the scanning thread to stop.
    """
    if app_instance.scanning:
        app_instance.stop_event.set()
        app_instance.pause_event.set() # Also set pause event to unblock if paused
        app_instance.after(0, app_instance._update_console_line, "🛑 Stop signal sent. Waiting for scan to terminate...\n")
        debug_print("Stop signal sent.", file=__file__, function=inspect.currentframe().f_code.co_name)
    else:
        app_instance.after(0, lambda: messagebox.showwarning("No Scan in Progress", "No scan is currently running to stop."))


def pause_resume_scan_logic(app_instance):
    """
    Toggles the pause/resume state of the scan.
    """
    if app_instance.scanning:
        app_instance.paused = not app_instance.paused
        if app_instance.paused:
            app_instance.pause_event.clear() # Clear event to block thread
            app_instance.pause_resume_button.config(text="Resume Scan", style='Red.TButton')
            app_instance.after(0, app_instance._start_pause_button_blink)
            app_instance.after(0, app_instance._update_console_line, "⏸️ Scan paused.\n")
            debug_print("Scan paused.", file=__file__, function=inspect.currentframe().f_code.co_name)
        else:
            app_instance.pause_event.set() # Set event to unblock thread
            app_instance.pause_resume_button.config(text="Pause Scan", style='Orange.TButton')
            app_instance.after(0, app_instance._stop_pause_button_blink)
            app_instance.after(0, app_instance._update_console_line, "▶️ Scan resumed.\n")
            debug_print("Scan resumed.", file=__file__, function=inspect.currentframe().f_code.co_name)
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
        
        # Enable plot button if there's data to plot
        if hasattr(app_instance, 'plotting_tab') and hasattr(app_instance.plotting_tab, 'plot_button'):
            if app_instance.collected_scans_dataframes:
                app_instance.plotting_tab.plot_button.config(state=tk.NORMAL)
            else:
                app_instance.plotting_tab.plot_button.config(state=tk.DISABLED)

    else:
        app_instance.disconnect_button.config(state=tk.DISABLED)
        app_instance.apply_button.config(state=tk.DISABLED)
        app_instance.load_preset_button.config(state=tk.DISABLED) # Disable load preset if no instrument
        # Disable query presets button
        if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
            app_instance.preset_files_tab.query_presets_button.config(state=tk.DISABLED)
        
        # Disable plot button if no instrument
        if hasattr(app_instance, 'plotting_tab') and hasattr(app_instance.plotting_tab, 'plot_button'):
            app_instance.plotting_tab.plot_button.config(state=tk.DISABLED)

