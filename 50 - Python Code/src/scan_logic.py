# src/scan_logic.py
import tkinter as tk
import threading
import time
import os
from datetime import datetime
import pandas as pd
import pyvisa
import inspect

from utils.scan_instrument import scan_bands
from utils.frequency_bands import MHZ_TO_HZ
from utils.instrument_control import log_visa_command, debug_print
from src.config_manager import save_config # Import save_config

from src.plot_logic import plot_single_scan_data, _open_plot_in_browser
from utils.averaging_utils import generate_current_cycle_average_csv_and_plot as generate_average_plot_logic
from src.instrument_logic import query_current_instrument_settings_logic # Import the logic function


def start_scan_thread_logic(app_instance):
    """
    Starts the spectrum analyzer scan in a new thread to keep the GUI responsive.
    Manages scan state and button enablement.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Attempting to start scan thread...", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)

    if not app_instance.inst:
        app_instance.after(0, lambda: app_instance._print_to_gui_console("⚠️ Warning: Please connect to an instrument first."))
        debug_print("Scan start failed: No instrument connected.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
        return
    
    if app_instance.scanning:
        app_instance.after(0, lambda: app_instance._print_to_gui_console("⚠️ Warning: A scan is already running."))
        debug_print("Scan start failed: Scan already in progress.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
        return

    # Save configuration when scan starts (this includes scan name and output folder)
    save_config(app_instance) # Corrected function call
    app_instance.after(0, lambda: app_instance._print_to_gui_console("Configuration saved before scan."))
    debug_print("Configuration saved before scan.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)

    app_instance.scanning = True
    app_instance.stop_scan_event.clear()
    app_instance.pause_scan_event.clear()

    # Disable relevant buttons during scan
    app_instance.start_scan_button.config(state=tk.DISABLED)
    app_instance.connect_button.config(state=tk.DISABLED) # Disable connect button
    app_instance.disconnect_button.config(state=tk.DISABLED) # Disable disconnect button
    app_instance.apply_button.config(state=tk.DISABLED) # Disable apply settings button
    app_instance.load_preset_button.config(state=tk.DISABLED) # Disable load preset button
    
    # Disable query presets button
    if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
        app_instance.preset_files_tab.query_presets_button.config(state=tk.DISABLED)
    
    # Disable plot button
    if hasattr(app_instance, 'plotting_tab') and hasattr(app_instance.plotting_tab, 'plot_button'):
        app_instance.plotting_tab.plot_button.config(state=tk.DISABLED)

    app_instance.pause_scan_button.config(state=tk.NORMAL)
    app_instance.stop_scan_button.config(state=tk.NORMAL)

    app_instance.after(0, lambda: app_instance._print_to_gui_console("Starting scan..."))

    # Get selected bands
    selected_bands = [
        item["band"] for item in app_instance.band_vars if item["var"].get()
    ]
    if not selected_bands:
        app_instance.after(0, lambda: app_instance._print_to_gui_console("⚠️ Warning: No bands selected for scan. Please select at least one band."))
        debug_print("No bands selected for scan.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
        _reset_scan_buttons(app_instance)
        return

    # Retrieve settings from Tkinter variables
    try:
        scan_rbw_hz = float(app_instance.scan_rbw_hz_var.get())
        rbw_step_size_hz = float(app_instance.rbw_step_size_hz_var.get())
        maxhold_time_seconds = float(app_instance.maxhold_time_seconds_var.get())
        maxhold_enabled = app_instance.maxhold_enabled_var.get()
        high_sensitivity = app_instance.high_sensitivity_var.get()
        preamp_on = app_instance.preamp_on_var.get()
        scan_rbw_segmentation = float(app_instance.scan_rbw_segmentation_var.get())
        num_scan_cycles = app_instance.num_scan_cycles_var.get()
        output_folder = app_instance.output_folder_var.get()
        scan_name = app_instance.scan_name_var.get()

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            app_instance.after(0, lambda: app_instance._print_to_gui_console(f"Created output directory: {output_folder}"))
            debug_print(f"Created output directory: {output_folder}", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)

    except ValueError as e:
        app_instance.after(0, lambda: app_instance._print_to_gui_console(f"❌ Error: Invalid input for scan settings: {e}"))
        debug_print(f"Invalid input for scan settings: {e}", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
        _reset_scan_buttons(app_instance)
        return

    app_instance.collected_scans_dataframes = [] # Clear previous scan data

    # Start the scan in a new thread
    app_instance.scan_thread = threading.Thread(
        target=_run_scan,
        args=(
            app_instance,
            selected_bands,
            scan_rbw_hz,
            rbw_step_size_hz,
            maxhold_time_seconds,
            output_folder,
            scan_name,
            app_instance.stop_scan_event,
            app_instance.pause_scan_event,
            app_instance._update_console_text, # Pass the direct update function
            maxhold_enabled,
            high_sensitivity,
            preamp_on,
            scan_rbw_segmentation,
            num_scan_cycles,
            app_instance._print_to_gui_console # Pass console_print_func
        )
    )
    app_instance.scan_thread.daemon = True # Allow the thread to exit with the main app
    app_instance.scan_thread.start()


def _run_scan(app_instance, selected_bands, scan_rbw_hz, rbw_step_size_hz, maxhold_time_seconds,
              output_folder, scan_name, stop_event, pause_event, update_console_line_func,
              maxhold_enabled, high_sensitivity, preamp_on, scan_rbw_segmentation, num_scan_cycles, console_print_func):
    """
    Internal function to run the scan cycles. Designed to be run in a separate thread.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Executing _run_scan in separate thread...", file=current_file, function=current_function, console_print_func=console_print_func)

    try:
        for cycle_num in range(num_scan_cycles):
            if stop_event.is_set():
                console_print_func(f"Scan stopped during cycle {cycle_num + 1}.")
                break

            while pause_event.is_set():
                update_console_line_func("Scan paused. Press 'Resume Scan' to continue.")
                time.sleep(0.5)
                if stop_event.is_set():
                    console_print_func(f"Scan stopped during pause in cycle {cycle_num + 1}.")
                    break
            if stop_event.is_set():
                break

            console_print_func(f"\n--- Starting Scan Cycle {cycle_num + 1}/{num_scan_cycles} ---")
            debug_print(f"Starting Scan Cycle {cycle_num + 1}", file=current_file, function=current_function, console_print_func=console_print_func)

            last_successful_band_index, output_csv_filename, markers_data = scan_bands( # Added markers_data to return
                app_instance_ref=app_instance, # Pass the app_instance itself
                inst=app_instance.inst,
                selected_bands=selected_bands,
                rbw_hz=scan_rbw_hz, # Corrected to use local variable
                ref_level_dbm=float(app_instance.reference_level_dbm_var.get()), # Corrected to use Tkinter var
                freq_shift_hz=float(app_instance.freq_shift_hz_var.get()), # Corrected to use Tkinter var
                maxhold_enabled=maxhold_enabled,
                high_sensitivity=high_sensitivity,
                preamp_on=preamp_on,
                rbw_step_size_hz=rbw_step_size_hz, # Corrected to use local variable
                cycle_wait_time_seconds=float(app_instance.cycle_wait_time_seconds_var.get()), # Corrected to use Tkinter var
                scan_name=scan_name,
                output_folder=output_folder,
                stop_event=stop_event,
                pause_event=pause_event,
                log_visa_commands_enabled=app_instance.log_visa_commands_enabled_var.get(),
                general_debug_enabled=app_instance.general_debug_enabled_var.get(),
                app_console_update_func=update_console_line_func, # Pass the console update method
                scan_rbw_segmentation=scan_rbw_segmentation # Pass scan_rbw_segmentation
            )

            if output_csv_filename:
                try:
                    # Load the newly created CSV into a DataFrame
                    df = pd.read_csv(output_csv_filename)
                    app_instance.collected_scans_dataframes.append(df)
                    app_instance.last_scan_markers = markers_data # Store markers for plotting
                    console_print_func(f"Collected data from '{os.path.basename(output_csv_filename)}' for averaging.")
                    debug_print(f"Added {os.path.basename(output_csv_filename)} to collected_scans_dataframes.", file=current_file, function=current_function, console_print_func=console_print_func)

                    # Generate single scan plot for the current cycle
                    current_time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                    # Filename includes cycle number and timestamp
                    plot_filename = os.path.join(output_folder, f"{scan_name}_Cycle{cycle_num}_{current_time_str}_scan.html")

                    fig, plot_html_path_return = plot_single_scan_data(
                        df, # Use the DataFrame for the current cycle
                        f"{scan_name} - Cycle {cycle_num} Scan",
                        output_html_path=plot_filename,
                        include_tv_markers=app_instance.include_tv_markers_var.get(),
                        include_gov_markers=app_instance.include_gov_markers_var.get(),
                        include_markers=app_instance.include_markers_var.get(),
                        last_scan_markers=app_instance.last_scan_markers
                    )
                    if fig and app_instance.open_html_after_complete_var.get():
                        _open_plot_in_browser(plot_html_path_return)

                except Exception as e:
                    console_print_func(f"❌ Error loading scan data into DataFrame or plotting: {e}")
                    debug_print(f"Error loading CSV to DataFrame or plotting: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                console_print_func(f"🚫 No data file generated for cycle {cycle_num}.")
                debug_print(f"No data file for cycle {cycle_num}.", file=current_file, function=current_function, console_print_func=console_print_func)

            # Wait time between cycles (if not the last cycle and not stopped)
            if cycle_num < num_scan_cycles and not stop_event.is_set():
                wait_time = float(app_instance.cycle_wait_time_seconds_var.get()) # Corrected to use Tkinter var
                console_print_func(f"Cycle {cycle_num} complete. Waiting {wait_time} seconds before next cycle...")
                time.sleep(wait_time)

    except pyvisa.errors.VisaIOError as e:
        console_print_func(f"❌ Error: VISA error during scan: {e}")
        debug_print(f"VISA Error during scan: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
    except Exception as e:
        console_print_func(f"❌ Error: An unexpected error occurred during scan: {e}")
        debug_print(f"Unexpected error during scan: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
    finally:
        app_instance.scanning = False
        app_instance.after(0, lambda: console_print_func("\n--- Scan process finished ---"))
        debug_print("Scan process finished.", file=current_file, function=current_function, console_print_func=console_print_func)
        app_instance.after(0, lambda: _reset_scan_buttons(app_instance)) # Reset buttons on main thread

        # After scan, if data was collected, generate the averaged plot
        if app_instance.collected_scans_dataframes:
            app_instance.after(0, lambda: generate_average_plot_logic(
                app_instance.collected_scans_dataframes,
                app_instance.scan_name_var,
                app_instance.output_folder_var,
                app_instance.open_html_after_complete_var,
                app_instance.include_tv_markers_var,
                app_instance.include_gov_markers_var,
                app_instance.include_markers_var,
                app_instance._print_to_gui_console # Pass console_print_func
            ))
        else:
            app_instance.after(0, lambda: console_print_func("🚫 No data collected across all cycles for plotting."))
        
        # Ensure instrument settings are queried and updated after scan completes
        if app_instance.inst:
            query_current_instrument_settings_logic(app_instance, console_print_func) # Call the logic function


def pause_scan_logic(app_instance):
    """
    Toggles the pause state of the scan.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    if app_instance.scanning:
        if app_instance.pause_scan_event.is_set():
            app_instance.pause_scan_event.clear()
            app_instance.pause_scan_button.config(text="Pause Scan", style='Orange.TButton')
            app_instance.after(0, lambda: app_instance._print_to_gui_console("Scan resumed."))
            debug_print("Scan resumed.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
        else:
            app_instance.pause_scan_event.set()
            app_instance.pause_scan_button.config(text="Resume Scan", style='Green.TButton')
            app_instance.after(0, lambda: app_instance._print_to_gui_console("Scan paused."))
            debug_print("Scan paused.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
    else:
        app_instance.after(0, lambda: app_instance._print_to_gui_console("⚠️ Warning: No scan is currently running to pause/resume."))
        debug_print("No scan to pause/resume.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)


def resume_scan_logic(app_instance):
    """
    Resumes a paused scan by clearing the pause event.
    This function simply calls pause_scan_logic, which handles the toggle.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Attempting to resume scan...", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
    pause_scan_logic(app_instance) # Call the existing toggle function


def stop_scan_logic(app_instance):
    """
    Signals the scan thread to stop.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    if app_instance.scanning:
        app_instance.stop_scan_event.set()
        app_instance.pause_scan_event.clear() # Clear pause in case it was paused
        app_instance.after(0, lambda: app_instance._print_to_gui_console("Stopping scan... Please wait for current operation to complete."))
        debug_print("Stop scan event set.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
        # Buttons will be reset by _run_scan's finally block
    else:
        app_instance.after(0, lambda: app_instance._print_to_gui_console("⚠️ Warning: No scan is currently running to stop."))
        debug_print("No scan to stop.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)


def _reset_scan_buttons(app_instance):
    """
    Resets the state of scan control buttons after a scan completes or is stopped.
    This function should be called on the main Tkinter thread using app_instance.after().
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Resetting scan buttons...", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
    app_instance.start_scan_button.config(state=tk.NORMAL)
    app_instance.pause_scan_button.config(state=tk.DISABLED, text="Pause Scan", style='Orange.TButton')
    app_instance.stop_scan_button.config(state=tk.DISABLED)
    
    # Re-enable disconnect and apply buttons
    app_instance.disconnect_button.config(state=tk.NORMAL)
    app_instance.apply_button.config(state=tk.NORMAL)
    
    # Re-enable load preset button based on instrument model
    if app_instance.inst and app_instance.instrument_model != "N9340B":
        app_instance.load_preset_button.config(state=tk.NORMAL)
    else:
        app_instance.load_preset_button.config(state=tk.DISABLED) # Disable if no instrument or N9340B
    
    # Re-enable query presets button
    if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
        if app_instance.inst:
            app_instance.preset_files_tab.query_presets_button.config(state=tk.NORMAL)
        else:
            app_instance.preset_files_tab.query_presets_button.config(state=tk.DISABLED)
        
    # Enable plot button if there's data to plot
    if hasattr(app_instance, 'plotting_tab') and hasattr(app_instance.plotting_tab, 'plot_button'):
        if app_instance.collected_scans_dataframes:
            app_instance.plotting_tab.plot_button.config(state=tk.NORMAL)
        else:
            app_instance.plotting_tab.plot_button.config(state=tk.DISABLED)

    # Ensure connect button visibility is correct after scan
    _update_button_states_on_connection(app_instance)


def _update_button_states_on_connection(app_instance):
    """
    Updates the state of connection-related buttons based on the instrument's connection status.
    This function should be called on the main Tkinter thread using app_instance.after().
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Updating button states based on connection...", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)

    if app_instance.inst: # Instrument is connected
        app_instance.connect_button.grid_forget() # Hide connect button
        app_instance.disconnect_button.config(state=tk.NORMAL)
        # Ensure disconnect button is visible if it was hidden
        # This line is redundant if the button is already gridded, but harmless.
        # app_instance.disconnect_button.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        app_instance.apply_button.config(state=tk.NORMAL)
        # Enable load preset button based on instrument model
        if app_instance.instrument_model != "N9340B":
            app_instance.load_preset_button.config(state=tk.NORMAL)
        else:
            app_instance.load_preset_button.config(state=tk.DISABLED)
        
        # Enable query presets button
        if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
            app_instance.preset_files_tab.query_presets_button.config(state=tk.NORMAL)
        
        # Enable plot button if there's data to plot
        if hasattr(app_instance, 'plotting_tab') and hasattr(app_instance.plotting_tab, 'plot_button'):
            if app_instance.collected_scans_dataframes:
                app_instance.plotting_tab.plot_button.config(state=tk.NORMAL)
            else:
                app_instance.plotting_tab.plot_button.config(state=tk.DISABLED)

    else: # Instrument is disconnected
        app_instance.connect_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew") # Show connect button
        app_instance.disconnect_button.config(state=tk.DISABLED)
        app_instance.apply_button.config(state=tk.DISABLED)
        app_instance.load_preset_button.config(state=tk.DISABLED) # Disable load preset if no instrument
        # Disable query presets button
        if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
            app_instance.preset_files_tab.query_presets_button.config(state=tk.DISABLED)
        
        # Disable plot button if no instrument
        if hasattr(app_instance, 'plotting_tab') and hasattr(app_instance.plotting_tab, 'plot_button'):
            app_instance.plotting_tab.plot_button.config(state=tk.DISABLED)

