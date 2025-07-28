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

    if app_instance.is_scanning:
        app_instance.after(0, lambda: app_instance._print_to_gui_console("ℹ️ Info: Scan is already in progress."))
        debug_print("Scan already in progress.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
        return

    # Disable buttons during scan
    app_instance.start_scan_button.config(state=tk.DISABLED)
    app_instance.stop_scan_button.config(state=tk.NORMAL)
    
    # Disable buttons in InstrumentTab
    if hasattr(app_instance, 'instrument_tab'):
        app_instance.instrument_tab.connect_button.config(state=tk.DISABLED)
        app_instance.instrument_tab.disconnect_button.config(state=tk.DISABLED)
        app_instance.instrument_tab.apply_settings_button.config(state=tk.DISABLED) # Updated button name
        app_instance.instrument_tab.query_settings_button.config(state=tk.DISABLED)

    # Disable relevant buttons in other tabs
    if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
        app_instance.preset_files_tab.query_presets_button.config(state=tk.DISABLED)
    if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'load_preset_button'):
        app_instance.preset_files_tab.load_preset_button.config(state=tk.DISABLED)
    if hasattr(app_instance, 'plotting_tab') and hasattr(app_instance.plotting_tab, 'plot_button'):
        app_instance.plotting_tab.plot_button.config(state=tk.DISABLED)
    if hasattr(app_instance, 'plotting_tab') and hasattr(app_instance.plotting_tab, 'plot_average_button'):
        app_instance.plotting_tab.plot_average_button.config(state=tk.DISABLED)


    app_instance.is_scanning = True
    app_instance.scan_thread = threading.Thread(target=_run_scan, args=(app_instance,))
    app_instance.scan_thread.daemon = True # Allow the thread to exit with the main application
    app_instance.scan_thread.start()
    app_instance.after(0, lambda: app_instance._print_to_gui_console("▶️ Scan started..."))
    debug_print("Scan thread started.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)


def stop_scan_logic(app_instance):
    """
    Stops the currently running scan.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Attempting to stop scan...", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)

    if app_instance.is_scanning:
        app_instance.is_scanning = False # Signal the thread to stop
        # Wait for the thread to finish (optional, but good for cleanup)
        if app_instance.scan_thread and app_instance.scan_thread.is_alive():
            app_instance.scan_thread.join(timeout=1.0) # Wait for up to 1 second
            if app_instance.scan_thread.is_alive():
                app_instance.after(0, lambda: app_instance._print_to_gui_console("⚠️ Warning: Scan thread did not terminate gracefully."))
                debug_print("Scan thread did not terminate gracefully.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
        app_instance.after(0, lambda: app_instance._print_to_gui_console("⏹️ Scan stopped by user."))
        debug_print("Scan stopped by user.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
    else:
        app_instance.after(0, lambda: app_instance._print_to_gui_console("ℹ️ Info: No scan is currently running."))
        debug_print("No scan is currently running.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
    
    # Re-enable buttons after stopping
    # Pass the InstrumentTab buttons directly
    if hasattr(app_instance, 'instrument_tab'):
        app_instance.after(0, lambda: update_connection_status_logic(
            app_instance,
            True, # Assuming connection is still active after stopping scan
            app_instance._print_to_gui_console,
            connect_btn=app_instance.instrument_tab.connect_button,
            disconnect_btn=app_instance.instrument_tab.disconnect_button,
            apply_btn=app_instance.instrument_tab.apply_settings_button,
            query_btn=app_instance.instrument_tab.query_settings_button
        ))


def _run_scan(app_instance):
    """
    The main scan loop, run in a separate thread.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Entering _run_scan function.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)

    try:
        selected_bands = [item["band"] for item in app_instance.band_vars if item["var"].get()]
        if not selected_bands:
            app_instance.after(0, lambda: app_instance._print_to_gui_console("⚠️ Warning: No frequency bands selected for scan."))
            debug_print("No frequency bands selected.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
            app_instance.is_scanning = False
            # Re-enable buttons
            if hasattr(app_instance, 'instrument_tab'):
                app_instance.after(0, lambda: update_connection_status_logic(
                    app_instance, True, app_instance._print_to_gui_console,
                    connect_btn=app_instance.instrument_tab.connect_button,
                    disconnect_btn=app_instance.instrument_tab.disconnect_button,
                    apply_btn=app_instance.instrument_tab.apply_settings_button,
                    query_btn=app_instance.instrument_tab.query_settings_button
                ))
            return

        num_scan_cycles = int(app_instance.num_scan_cycles_var.get())
        cycle_wait_time = float(app_instance.cycle_wait_time_seconds_var.get()) # Use correct var name
        scan_name = app_instance.scan_name_var.get()
        output_dir = app_instance.output_folder_var.get()
        open_html_after_complete = app_instance.open_html_after_complete_var.get()
        include_tv_markers = app_instance.include_tv_markers_var.get()
        include_gov_markers = app_instance.include_gov_markers_var.get()
        include_markers = app_instance.include_markers_var.get() # Get state of include_markers checkbox

        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        debug_print(f"Ensured output directory exists: {output_dir}", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)

        app_instance.collected_scans_dataframes = [] # Clear previous scan data
        app_instance.last_scan_markers = [] # Clear previous markers

        for cycle in range(num_scan_cycles):
            if not app_instance.is_scanning:
                app_instance.after(0, lambda: app_instance._print_to_gui_console(f"Scan cycle {cycle + 1}/{num_scan_cycles} interrupted."))
                debug_print(f"Scan cycle {cycle + 1}/{num_scan_cycles} interrupted.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
                break
            
            app_instance.after(0, lambda c=cycle: app_instance._print_to_gui_console(f"Scanning Cycle {c + 1}/{num_scan_cycles}..."))
            debug_print(f"Starting scan cycle {cycle + 1}/{num_scan_cycles}.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)

            # Perform the actual scan
            last_successful_band_index, scan_df, markers_data = scan_bands(
                app_instance.inst,
                selected_bands,
                app_instance.scan_rbw_hz_var, # Use correct var name
                app_instance.rbw_segmentation_var,
                app_instance.reference_level_dbm_var, # Use correct var name
                app_instance.freq_shift_hz_var, # Use correct var name
                app_instance.max_hold_enabled_var,
                app_instance.high_sensitivity_var,
                app_instance.preamp_on_var,
                app_instance.cycle_wait_time_seconds_var, # Pass cycle_wait_time_seconds_var
                app_instance.maxhold_time_seconds_var,    # Pass maxhold_time_seconds_var
                app_instance._print_to_gui_console, # Pass console_print_func
                app_instance # Pass app_instance_ref
            )

            if scan_df is not None and not scan_df.empty:
                app_instance.collected_scans_dataframes.append(scan_df)
                app_instance.last_scan_markers = markers_data # Store markers from this scan
                app_instance.after(0, lambda: app_instance._print_to_gui_console(f"✅ Data collected for cycle {cycle + 1}."))
                debug_print(f"Data collected for cycle {cycle + 1}. DataFrame shape: {scan_df.shape}", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)

                # Generate and save single scan plot
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                plot_filename = os.path.join(output_dir, f"{scan_name}_Scan_{timestamp}.html")
                
                app_instance.after(0, lambda: app_instance._print_to_gui_console(f"Generating plot for cycle {cycle + 1}..."))
                debug_print(f"Generating plot for cycle {cycle + 1} to {plot_filename}.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
                
                fig, html_path = plot_single_scan_data(
                    scan_df,
                    f"{scan_name} - Cycle {cycle + 1} - {timestamp}",
                    include_tv_markers,
                    include_gov_markers,
                    include_markers, # Pass the include_markers flag
                    output_html_path=plot_filename,
                    console_print_func=app_instance._print_to_gui_console
                )
                if fig:
                    app_instance.after(0, lambda: app_instance._print_to_gui_console(f"✅ Plot saved to: {html_path}"))
                    debug_print(f"Plot saved to: {html_path}", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
                    app_instance.plotting_tab.last_plot_path = html_path # Update last plot path in plotting tab
                    if open_html_after_complete:
                        app_instance.after(0, lambda p=html_path: _open_plot_in_browser(p, app_instance._print_to_gui_console))
                else:
                    app_instance.after(0, lambda: app_instance._print_to_gui_console(f"🚫 Failed to generate plot for cycle {cycle + 1}."))
                    debug_print(f"Failed to generate plot for cycle {cycle + 1}.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
            else:
                app_instance.after(0, lambda: app_instance._print_to_gui_console(f"🚫 No data collected for cycle {cycle + 1}."))
                debug_print(f"No data collected for cycle {cycle + 1}.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)

            if app_instance.is_scanning and cycle < num_scan_cycles - 1:
                app_instance.after(0, lambda: app_instance._print_to_gui_console(f"Waiting {cycle_wait_time} seconds before next cycle..."))
                debug_print(f"Waiting {cycle_wait_time} seconds before next cycle.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
                time.sleep(cycle_wait_time)

        app_instance.is_scanning = False
        app_instance.after(0, lambda: app_instance._print_to_gui_console("✅ Scan complete."))
        debug_print("Scan complete. Re-enabling buttons.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)

        # After scan, if there's collected data, update the markers tab
        if app_instance.last_scan_markers:
            # Assuming markers_data is a list of dicts, and plot_single_scan_data returns it correctly
            # The headers might need to be inferred or passed explicitly if not consistent
            if app_instance.last_scan_markers: # Check if list is not empty
                # Assuming all markers have the same keys, take headers from the first marker
                headers = list(app_instance.last_scan_markers[0].keys())
                app_instance.after(0, lambda h=headers, r=app_instance.last_scan_markers: app_instance.markers_tab.update_markers_data(h, r))
                app_instance.after(0, lambda: app_instance._print_to_gui_console(f"📊 Markers data updated in Markers tab with {len(app_instance.last_scan_markers)} entries."))
                debug_print(f"Markers data updated in Markers tab with {len(app_instance.last_scan_markers)} entries.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
            else:
                app_instance.after(0, lambda: app_instance._print_to_gui_console("ℹ️ No markers extracted during scan to update Markers tab."))
                debug_print("No markers extracted during scan.", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
        
        # Re-enable buttons after scan completes
        # Pass the InstrumentTab buttons directly
        if hasattr(app_instance, 'instrument_tab'):
            app_instance.after(0, lambda: update_connection_status_logic(
                app_instance,
                True, # Assuming connection is still active after stopping scan
                app_instance._print_to_gui_console,
                connect_btn=app_instance.instrument_tab.connect_button,
                disconnect_btn=app_instance.instrument_tab.disconnect_button,
                apply_btn=app_instance.instrument_tab.apply_settings_button,
                query_btn=app_instance.instrument_tab.query_settings_button
            ))

    except Exception as e:
        app_instance.is_scanning = False
        app_instance.after(0, lambda: app_instance._print_to_gui_console(f"❌ An error occurred during scan: {e}"))
        debug_print(f"Error during scan: {e}", file=current_file, function=current_function, console_print_func=app_instance._print_to_gui_console)
        # Re-enable buttons on error
        if hasattr(app_instance, 'instrument_tab'):
            app_instance.after(0, lambda: update_connection_status_logic(
                app_instance,
                True, # Assuming connection is still active despite error
                app_instance._print_to_gui_console,
                connect_btn=app_instance.instrument_tab.connect_button,
                disconnect_btn=app_instance.instrument_tab.disconnect_button,
                apply_btn=app_instance.instrument_tab.apply_settings_button,
                query_btn=app_instance.instrument_tab.query_settings_button
            ))


def update_connection_status_logic(app_instance, is_connected, console_print_func,
                                   connect_btn=None, disconnect_btn=None,
                                   apply_btn=None, query_btn=None):
    """
    Updates the GUI elements based on the instrument connection status.
    This logic is centralized here.
    Accepts specific button widgets for InstrumentTab.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Updating connection status GUI elements. Connected: {is_connected}", file=current_file, function=current_function, console_print_func=console_print_func)

    # Main scan control buttons
    app_instance.start_scan_button.config(state=tk.NORMAL if is_connected and not app_instance.is_scanning else tk.DISABLED)
    app_instance.stop_scan_button.config(state=tk.NORMAL if is_connected and app_instance.is_scanning else tk.DISABLED)

    # InstrumentTab specific buttons
    if connect_btn and disconnect_btn and apply_btn and query_btn:
        if is_connected:
            connect_btn.grid_remove() # Hide connect button
            disconnect_btn.config(state=tk.NORMAL)
            apply_btn.config(state=tk.NORMAL)
            query_btn.config(state=tk.NORMAL)
        else: # Instrument is disconnected
            connect_btn.grid() # Show connect button
            disconnect_btn.config(state=tk.DISABLED)
            apply_btn.config(state=tk.DISABLED)
            query_btn.config(state=tk.DISABLED)
    else:
        debug_print("InstrumentTab buttons not provided to update_connection_status_logic.", file=current_file, function=current_function, console_print_func=console_print_func)


    # Preset tab buttons
    if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
        app_instance.preset_files_tab.query_presets_button.config(state=tk.NORMAL if is_connected else tk.DISABLED)
    if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'load_preset_button'):
         app_instance.preset_files_tab.load_preset_button.config(state=tk.NORMAL if is_connected else tk.DISABLED)
    
    # Plotting tab buttons
    if hasattr(app_instance, 'plotting_tab') and hasattr(app_instance.plotting_tab, 'plot_button'):
        # Plot buttons should be enabled if there's data OR if connected (to allow plotting new data)
        # For simplicity, let's enable if connected and not scanning, or if data exists.
        # The plotting tab itself should manage its button states more precisely based on data.
        # Here, we just ensure they are disabled if no instrument.
        if is_connected and not app_instance.is_scanning:
            app_instance.plotting_tab.plot_button.config(state=tk.NORMAL)
            app_instance.plotting_tab.plot_average_button.config(state=tk.NORMAL)
        else:
            app_instance.plotting_tab.plot_button.config(state=tk.DISABLED)
            app_instance.plotting_tab.plot_average_button.config(state=tk.DISABLED)

