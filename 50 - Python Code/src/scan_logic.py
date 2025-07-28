# src/scan_logic.py
import tkinter as tk
import inspect

# Import debug_print from utils
from utils.instrument_control import debug_print


def update_connection_status_logic(app_instance, is_connected, console_print_func):
    """
    Updates the GUI elements based on the instrument connection status and scan status.
    This is the central function for managing button states across all relevant tabs.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Updating connection status GUI elements. Connected: {is_connected}", file=current_file, function=current_function, console_print_func=console_print_func)

    # Get references to tabs and their buttons (ensure they exist)
    instrument_tab = getattr(app_instance, 'instrument_tab', None)
    scan_control_tab = getattr(app_instance, 'scan_control_tab', None)
    preset_files_tab = getattr(app_instance, 'preset_files_tab', None)
    plotting_tab = getattr(app_instance, 'plotting_tab', None)
    markers_display_tab = getattr(app_instance, 'markers_display_tab', None) # Added for completeness

    is_scanning = scan_control_tab.is_scanning if scan_control_tab else False
    is_paused = scan_control_tab.is_paused if scan_control_tab else False

    # --- InstrumentTab buttons ---
    if instrument_tab:
        if is_connected:
            instrument_tab.connect_button.grid_remove() # Hide connect button
            instrument_tab.disconnect_button.config(state=tk.NORMAL)
            instrument_tab.apply_settings_button.config(state=tk.NORMAL)
            instrument_tab.query_settings_button.config(state=tk.NORMAL)
        else: # Instrument is disconnected
            instrument_tab.connect_button.grid() # Show connect button
            instrument_tab.disconnect_button.config(state=tk.DISABLED)
            instrument_tab.apply_settings_button.config(state=tk.DISABLED)
            instrument_tab.query_settings_button.config(state=tk.DISABLED)
    else:
        debug_print("InstrumentTab not found when updating connection status.", file=current_file, function=current_function, console_print_func=console_print_func)

    # --- ScanControlTab buttons ---
    if scan_control_tab:
        if is_connected:
            if is_scanning and not is_paused:
                scan_control_tab.start_button.config(state=tk.DISABLED)
                scan_control_tab.pause_button.config(state=tk.NORMAL)
                scan_control_tab.stop_button.config(state=tk.NORMAL)
            elif is_scanning and is_paused:
                scan_control_tab.start_button.config(state=tk.NORMAL) # To resume
                scan_control_tab.pause_button.config(state=tk.DISABLED)
                scan_control_tab.stop_button.config(state=tk.NORMAL)
            else: # Connected but not scanning/paused
                scan_control_tab.start_button.config(state=tk.NORMAL)
                scan_control_tab.pause_button.config(state=tk.DISABLED)
                scan_control_tab.stop_button.config(state=tk.DISABLED)
        else: # Not connected
            scan_control_tab.start_button.config(state=tk.DISABLED)
            scan_control_tab.pause_button.config(state=tk.DISABLED)
            scan_control_tab.stop_button.config(state=tk.DISABLED)
    else:
        debug_print("ScanControlTab not found when updating connection status.", file=current_file, function=current_function, console_print_func=console_print_func)


    # --- Preset tab buttons ---
    if preset_files_tab:
        # The query_presets_button state depends on connection and scanning status
        preset_files_tab.query_presets_button.config(state=tk.NORMAL if is_connected and not is_scanning else tk.DISABLED)
        # The load_preset_button state is managed internally by PresetFilesTab's _select_preset
        # and update_preset_list methods, so we don't directly configure it here.
        # It should be disabled by default and enabled only when a preset is selected.
    else:
        debug_print("PresetFilesTab not found when updating connection status.", file=current_file, function=current_function, console_print_func=console_print_func)
    
    # --- Plotting tab buttons ---
    if plotting_tab:
        # Plot buttons should be enabled if there's data OR if connected (to allow plotting new data)
        # Here, we just ensure they are disabled if no instrument or scanning.
        if is_connected and not is_scanning:
            plotting_tab.plot_button.config(state=tk.NORMAL)
            plotting_tab.plot_average_button.config(state=tk.NORMAL)
        else:
            plotting_tab.plot_button.config(state=tk.DISABLED)
            plotting_tab.plot_average_button.config(state=tk.DISABLED)
    else:
        debug_print("PlottingTab not found when updating connection status.", file=current_file, function=current_function, console_print_func=console_print_func)

    # --- Markers Display Tab buttons (if any) ---
    # Assuming MarkersDisplayTab might have buttons that depend on connection/scan status
    if markers_display_tab:
        # Example: if you have a button to "Query Markers from Instrument"
        # markers_display_tab.query_markers_button.config(state=tk.NORMAL if is_connected and not is_scanning else tk.DISABLED)
        pass # No specific buttons to update in MarkersDisplayTab based on provided code
    else:
        debug_print("MarkersDisplayTab not found when updating connection status.", file=current_file, function=current_function, console_print_func=console_print_func)

