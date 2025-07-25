# src/config_manager.py
import configparser
import os
import tkinter as tk # Import tkinter for update_idletasks

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
        7. Calls `_populate_vars_from_config()` to set Tkinter variables based on loaded config.
    Outputs: None
    """
    print(f"DEBUG: Attempting to load configuration from {app_instance.CONFIG_FILE}...")
    
    # Create a ConfigParser instance
    app_instance.config = configparser.ConfigParser()

    # Read the config file
    if os.path.exists(app_instance.CONFIG_FILE):
        app_instance.config.read(app_instance.CONFIG_FILE)
        print(f"DEBUG: Config file '{app_instance.CONFIG_FILE}' read successfully.")
    else:
        print(f"DEBUG: Config file '{app_instance.CONFIG_FILE}' not found. Will create with defaults.")

    # Ensure DEFAULT_SETTINGS section exists
    if 'DEFAULT_SETTINGS' not in app_instance.config:
        app_instance.config['DEFAULT_SETTINGS'] = {}
        print("DEBUG: Created [DEFAULT_SETTINGS] section.")

    # Define default values for settings that might be missing
    default_settings_to_ensure = {
        'default_rbw_step_size_hz': '1000000',
        'default_cycle_wait_time_seconds': '0',
        'default_maxhold_time_seconds': '3',
        'default_scan_rbw_hz': '10000',
        'default_reference_level_dbm': '-40',
        'default_freq_shift_hz': '0',
        'default_maxhold_enabled': 'True',
        'default_include_gov_markers': 'True',
        'default_include_tv_markers': 'True',
        'default_open_html_after_complete': 'True',
        'default_high_sensitivity': 'True',
        'default_preamp_on': 'True',
        'default_debug_mode': 'False',
        'default_window_geometry': '1400x780+100+100', # Default geometry
        'default_scan_directory': 'scan_data',
        'default_scan_name': 'MyScan',
        'default_default_focus_width': '10000.0', # Default for new focus width
    }

    # Dynamically generate default_selected_bands from SCAN_BAND_RANGES
    default_selected_bands_str = ",".join([band["Band Name"] for band in app_instance.SCAN_BAND_RANGES])
    default_settings_to_ensure['default_selected_bands'] = default_selected_bands_str

    # Populate DEFAULT_SETTINGS with predefined defaults if they are missing
    for key, value in default_settings_to_ensure.items():
        if key not in app_instance.config['DEFAULT_SETTINGS']:
            app_instance.config['DEFAULT_SETTINGS'][key] = value
            print(f"DEBUG: Set default '{key}' to '{value}'.")

    # Ensure LAST_USED_SETTINGS section exists
    if 'LAST_USED_SETTINGS' not in app_instance.config:
        app_instance.config['LAST_USED_SETTINGS'] = {}
        print("DEBUG: Created [LAST_USED_SETTINGS] section.")

    # Ensure last_window_geometry is set in LAST_USED_SETTINGS
    if 'last_window_geometry' not in app_instance.config['LAST_USED_SETTINGS']:
        default_geometry = app_instance.config['DEFAULT_SETTINGS'].get('default_window_geometry', '1400x780+100+100')
        app_instance.config['LAST_USED_SETTINGS']['last_window_geometry'] = default_geometry
        print(f"DEBUG: Initialized 'last_window_geometry' in LAST_USED_SETTINGS to default: {default_geometry}")

    _populate_vars_from_config(app_instance)
    print("DEBUG: Configuration loaded and variables populated.")

def _populate_vars_from_config(app_instance):
    """
    Populates Tkinter variables in the app_instance with values from the loaded config.
    It prioritizes LAST_USED_SETTINGS, falling back to DEFAULT_SETTINGS if a setting
    is not found in LAST_USED_SETTINGS.

    Inputs:
        app_instance (App): The main application instance.
    Process:
        1. Iterates through `app_instance.setting_var_map`.
        2. For each setting, attempts to get its value from `LAST_USED_SETTINGS`.
        3. If not found, falls back to `DEFAULT_SETTINGS`.
        4. Converts the string value to the appropriate Python type (float, int, bool, str).
        5. Sets the corresponding Tkinter variable.
        6. Handles special cases for selected bands checkboxes.
    Outputs: None (modifies Tkinter variable states)
    """
    for var_name, (last_key, default_key, var_type) in app_instance.setting_var_map.items():
        value = None
        # Try to get from LAST_USED_SETTINGS
        if last_key and 'LAST_USED_SETTINGS' in app_instance.config:
            value = app_instance.config['LAST_USED_SETTINGS'].get(last_key)
        
        # If not found in LAST_USED_SETTINGS, try DEFAULT_SETTINGS
        if value is None and default_key and 'DEFAULT_SETTINGS' in app_instance.config:
            value = app_instance.config['DEFAULT_SETTINGS'].get(default_key)

        if value is not None:
            try:
                if var_type == bool:
                    converted_value = value.lower() == 'true'
                elif var_type == int:
                    converted_value = int(float(value)) # Handle potential float strings from config
                elif var_type == float:
                    converted_value = float(value)
                else: # str
                    converted_value = value
                
                # Special handling for last_selected_bands_str
                if var_name == 'last_selected_bands_str':
                    app_instance.last_selected_bands_str.set(converted_value)
                    # This will be used by set_band_checkboxes_from_config_logic later
                else:
                    tk_var = getattr(app_instance, var_name)
                    tk_var.set(converted_value)
                    # print(f"DEBUG: Populated {var_name} with {converted_value} (type: {var_type.__name__})")
            except ValueError as e:
                print(f"WARNING: Could not convert config value for '{var_name}' ('{value}') to type {var_type.__name__}: {e}")
        else:
            print(f"WARNING: No value found for '{var_name}' in config. Using Tkinter default.")

def save_config(app_instance):
    """
    Saves the current application settings from Tkinter variables to `config.ini`.
    This function is called upon application shutdown or when a scan starts.

    Inputs:
        app_instance (App): The main application instance.
    Process:
        1. Ensures `LAST_USED_SETTINGS` section exists in `app_instance.config`.
        2. Iterates through `app_instance.setting_var_map` and saves the current value
           of each Tkinter variable to the corresponding key in `LAST_USED_SETTINGS`.
        3. Handles special saving for selected frequency bands.
        4. Retrieves the current window geometry and saves it.
        5. Writes the updated configuration to `app_instance.CONFIG_FILE`.
    Outputs: None
    """
    print("DEBUG: save_config called.") # Debug print to show when save_config is invoked
    # Ensure LAST_USED_SETTINGS section exists
    if 'LAST_USED_SETTINGS' not in app_instance.config:
        app_instance.config['LAST_USED_SETTINGS'] = {}

    print("DEBUG: Saving configuration...")
    for var_name, (last_key, _, _) in app_instance.setting_var_map.items():
        if last_key and var_name != 'last_selected_bands_str':
            tk_var = getattr(app_instance, var_name)
            app_instance.config['LAST_USED_SETTINGS'][last_key] = str(tk_var.get())
            print(f"DEBUG: Saving '{last_key}': '{tk_var.get()}'")
    
    selected_band_names = [item["band"]["Band Name"] for item in app_instance.band_vars if item["var"].get()]
    bands_to_save = ",".join(selected_band_names)
    app_instance.config['LAST_USED_SETTINGS']['last_selected_bands'] = bands_to_save
    print(f"DEBUG: Saving 'last_selected_bands': '{bands_to_save}'")
    
    # Force update of window geometry before saving
    app_instance.update_idletasks() # IMPORTANT: Ensure geometry is up-to-date
    current_geometry = app_instance.winfo_geometry()
    app_instance.config['LAST_USED_SETTINGS']['last_window_geometry'] = current_geometry
    print(f"DEBUG: Saving 'last_window_geometry': '{current_geometry}'")

    try:
        with open(app_instance.CONFIG_FILE, 'w') as configfile:
            app_instance.config.write(configfile)
        print(f"Configuration saved successfully to {app_instance.CONFIG_FILE}")
    except IOError as e:
        print(f"❌ Error saving configuration to {app_instance.CONFIG_FILE}: {e}")
    except Exception as e:
        print(f"❌ An unexpected error occurred while saving configuration: {e}")
