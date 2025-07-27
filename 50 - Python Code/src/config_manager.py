# src/config_manager.py
import configparser
import os
import tkinter as tk # Import tkinter for update_idletasks
from utils.instrument_control import debug_print # Import debug_print
import inspect # Import inspect module for debug_print

def load_config(app_instance):
    """
    Loads configuration settings from `config.ini`.
    If the file or sections are missing, it ensures default settings are present.
    This function is called during application initialization to restore
    the last-used settings or apply defaults.

    Inputs:
        app_instance (App): The main application instance, providing access to its config object,
                            Tkinter variables, and other necessary attributes like SCAN_BAND_RANGES.
    Process:
        1. Reads `app_instance.CONFIG_FILE` into `app_instance.config` using `configparser`.
        2. Ensures `DEFAULT_SETTINGS` section exists, creating it if not.
        3. Dynamically generates default selected bands string from `app_instance.SCAN_BAND_RANGES`.
        4. Populates `DEFAULT_SETTINGS` with predefined default values if any are missing.
        5. Ensures `LAST_USED_SETTINGS` section exists, creating it if not.
        6. **Ensures `last_window_geometry` is set in `LAST_USED_SETTINGS`:**
           If `last_window_geometry` is not found in `LAST_USED_SETTINGS`, it copies the
           `DEFAULT_WINDOW_GEOMETRY` from `DEFAULT_SETTINGS` to `LAST_USED_SETTINGS` (or a hardcoded default if missing).
        7. Calls `_`"""
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    app_instance.config.read(app_instance.CONFIG_FILE)

    if 'DEFAULT_SETTINGS' not in app_instance.config:
        app_instance.config['DEFAULT_SETTINGS'] = {}
        debug_print("Created missing 'DEFAULT_SETTINGS' section.", file=current_file, function=current_function)

    if 'LAST_USED_SETTINGS' not in app_instance.config:
        app_instance.config['LAST_USED_SETTINGS'] = {}
        debug_print("Created missing 'LAST_USED_SETTINGS' section.", file=current_file, function=current_function)

    # Define default values for all settings
    # Note: default_selected_bands is dynamically generated below
    default_values = {
        'default_rbw_step_size_hz': '1000000',
        'default_cycle_wait_time_seconds': '0.5',
        'default_maxhold_time_seconds': '3',
        'default_scan_rbw_hz': '10000',
        'default_reference_level_dbm': '-40',
        'default_freq_shift_hz': '0',
        'default_maxhold_enabled': 'True',
        'default_high_sensitivity': 'True',
        'default_preamp_on': 'True',
        'default_scan_rbw_segmentation': '1000000.0',
        'default_default_focus_width': '10000.0',
        'default_include_gov_markers': 'True',
        'default_include_tv_markers': 'True',
        'default_open_html_after_complete': 'True',
        'default_general_debug_enabled': 'False',
        'default_log_visa_commands_enabled': 'False',
        'default_window_geometry': '1400x780+100+100',
        'default_scan_directory': 'scan_data',
        'default_scan_name': 'MyScan',
        'default_num_scan_cycles': '1',
        'default_include_markers': 'False', # New default setting
    }

    # Ensure all default settings are present
    for key, value in default_values.items():
        if key not in app_instance.config['DEFAULT_SETTINGS']:
            app_instance.config['DEFAULT_SETTINGS'][key] = value
            debug_print(f"Added missing default setting: {key}={value}", file=current_file, function=current_function)

    # Dynamically generate default_selected_bands if missing
    if 'default_selected_bands' not in app_instance.config['DEFAULT_SETTINGS']:
        default_bands_str = ",".join([band["Band Name"] for band in app_instance.SCAN_BAND_RANGES])
        app_instance.config['DEFAULT_SETTINGS']['default_selected_bands'] = default_bands_str
        debug_print(f"Generated default_selected_bands: {default_bands_str}", file=current_file, function=current_function)

    # Ensure last_window_geometry is set in LAST_USED_SETTINGS
    if 'last_window_geometry' not in app_instance.config['LAST_USED_SETTINGS']:
        default_geometry = app_instance.config['DEFAULT_SETTINGS'].get('default_window_geometry', '1400x780+100+100')
        app_instance.config['LAST_USED_SETTINGS']['last_window_geometry'] = default_geometry
        debug_print(f"Set last_window_geometry to default: {default_geometry}", file=current_file, function=current_function)

    # Helper to get config value, with fallback to default, then hardcoded default
    def _get_config_value(section, key, default_key=None, hardcoded_fallback=None):
        if app_instance.config.has_option(section, key):
            return app_instance.config.get(section, key)
        elif default_key and app_instance.config.has_option('DEFAULT_SETTINGS', default_key):
            return app_instance.config.get('DEFAULT_SETTINGS', default_key)
        else:
            return hardcoded_fallback

    def _get_boolean_config(last_key, default_key):
        value_str = _get_config_value('LAST_USED_SETTINGS', last_key, default_key, 'False')
        return value_str.lower() == 'true'

    # Load settings into Tkinter variables using the mapping
    for var_name, (last_key, default_key, tk_var) in app_instance.setting_var_map.items():
        if isinstance(tk_var, tk.BooleanVar):
            tk_var.set(_get_boolean_config(last_key, default_key))
        elif isinstance(tk_var, tk.IntVar):
            try:
                value = int(_get_config_value('LAST_USED_SETTINGS', last_key, default_key, '0'))
                tk_var.set(value)
            except ValueError:
                debug_print(f"Warning: Could not convert '{_get_config_value('LAST_USED_SETTINGS', last_key, default_key, '0')}' to int for {last_key}. Using default.", file=current_file, function=current_function)
                tk_var.set(int(app_instance.config['DEFAULT_SETTINGS'].get(default_key, '0')))
        else: # StringVar
            tk_var.set(_get_config_value('LAST_USED_SETTINGS', last_key, default_key, ''))

    # Special handling for output_folder_var which maps to scan_directory_var's config keys
    app_instance.output_folder_var.set(app_instance.scan_directory_var.get())

    debug_print("Configuration loaded.", file=current_file, function=current_function)


def save_config(app_instance):
    """
    Saves the current application settings from Tkinter variables to `config.ini`.
    This function is called when the application is closing or when a scan starts,
    ensuring that the last-used settings are preserved.

    Inputs:
        app_instance (App): The main application instance, providing access to its config object
                            and Tkinter variables.
    Process:
        1. Iterates through `app_instance.setting_var_map` to save the current value of each
           Tkinter variable to the `LAST_USED_SETTINGS` section of `app_instance.config`.
        2. Specifically handles saving the currently selected frequency bands.
        3. Updates and saves the current window geometry.
        4. Writes the updated configuration to `app_instance.CONFIG_FILE`.
    Outputs: None
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if 'LAST_USED_SETTINGS' not in app_instance.config:
        app_instance.config['LAST_USED_SETTINGS'] = {}
        debug_print("Created missing 'LAST_USED_SETTINGS' section during save.", file=current_file, function=current_function)

    # Save general settings using the map
    for var_name, (last_key, default_key, tk_var) in app_instance.setting_var_map.items():
        if last_key and var_name != 'selected_bands_str_var': # Skip the string version as we handle bands separately
            app_instance.config['LAST_USED_SETTINGS'][last_key] = str(tk_var.get())
            debug_print(f"Saving '{last_key}': '{tk_var.get()}'", file=current_file, function=current_function)
    
    # Save selected bands
    selected_band_names = [item["band"]["Band Name"] for item in app_instance.band_vars if item["var"].get()]
    bands_to_save = ",".join(selected_band_names)
    app_instance.config['LAST_USED_SETTINGS']['last_selected_bands'] = bands_to_save
    debug_print(f"Saving 'last_selected_bands': '{bands_to_save}'", file=current_file, function=current_function)
    
    # Force update of window geometry before saving
    app_instance.update_idletasks() # IMPORTANT: Ensure geometry is up-to-date
    current_geometry = app_instance.winfo_geometry()
    app_instance.config['LAST_USED_SETTINGS']['last_window_geometry'] = current_geometry
    debug_print(f"Saving 'last_window_geometry': '{current_geometry}'", file=current_file, function=current_function)

    try:
        with open(app_instance.CONFIG_FILE, 'w') as configfile:
            app_instance.config.write(configfile)
        debug_print(f"Configuration saved to {app_instance.CONFIG_FILE}", file=current_file, function=current_function)
    except IOError as e:
        debug_print(f"❌ Error saving configuration to {app_instance.CONFIG_FILE}: {e}", file=current_file, function=current_function)
    except Exception as e:
        debug_print(f"❌ An unexpected error occurred while saving configuration: {e}", file=current_file, function=current_function)
