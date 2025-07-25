# src/config_manager.py
import configparser
import os

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
           `DEFAULT_WINDOW_GEOMETRY` from `DEFAULT_SETTINGS` to `LAST_USED_SETTINGS`.
        7. Calls `_populate_vars_from_config()` to set Tkinter variables based on loaded config.
    Outputs: None
    """
    app_instance.config.read(app_instance.CONFIG_FILE)

    # Default settings section
    if 'DEFAULT_SETTINGS' not in app_instance.config:
        app_instance.config['DEFAULT_SETTINGS'] = {}
    
    # Dynamically generate the default selected bands string from SCAN_BAND_RANGES
    default_selected_bands_str = ",".join([band["Band Name"] for band in app_instance.SCAN_BAND_RANGES])

    # Ensure all default settings are present with their default values
    default_settings_map = {
        'DEFAULT_RBW_STEP_SIZE_HZ': '1000000',
        'DEFAULT_CYCLE_WAIT_TIME_SECONDS': '0',
        'DEFAULT_MAXHOLD_TIME_SECONDS': '3',
        'DEFAULT_SCAN_RBW_HZ': '10000',
        'DEFAULT_REFERENCE_LEVEL_DBM': '-40',
        'DEFAULT_FREQ_SHIFT_HZ': '0',
        'DEFAULT_MAXHOLD_ENABLED': 'True',
        'DEFAULT_INCLUDE_GOV_MARKERS': 'True',
        'DEFAULT_INCLUDE_TV_MARKERS': 'True',
        'DEFAULT_OPEN_HTML_AFTER_COMPLETE': 'True',
        'DEFAULT_HIGH_SENSITIVITY': 'True',
        'DEFAULT_PREAMP_ON': 'True',
        'DEFAULT_DEBUG_MODE': 'False',
        'DEFAULT_WINDOW_GEOMETRY': '1400x780+100+100',
        'DEFAULT_SCAN_DIRECTORY': 'scan_data',
        'DEFAULT_SCAN_NAME': 'MyScan',
        'DEFAULT_SELECTED_BANDS': default_selected_bands_str,
    }
    for key, default_val in default_settings_map.items():
        if key.lower() not in [k.lower() for k in app_instance.config['DEFAULT_SETTINGS'].keys()]:
            app_instance.config['DEFAULT_SETTINGS'][key] = default_val
    
    # Last used settings section
    if 'LAST_USED_SETTINGS' not in app_instance.config:
        app_instance.config['LAST_USED_SETTINGS'] = {}

    # Ensure LAST_WINDOW_GEOMETRY is present in LAST_USED_SETTINGS
    if not app_instance.config.has_option('LAST_USED_SETTINGS', 'last_window_geometry') or not app_instance.config.get('LAST_USED_SETTINGS', 'last_window_geometry'):
        app_instance.config['LAST_USED_SETTINGS']['last_window_geometry'] = app_instance.config.get('DEFAULT_SETTINGS', 'DEFAULT_WINDOW_GEOMETRY')

    # Load values into Tkinter variables, prioritizing LAST_USED_SETTINGS
    _populate_vars_from_config(app_instance)

def _populate_vars_from_config(app_instance):
    """
    Populates Tkinter variables with values from the loaded configuration.
    It prioritizes `LAST_USED_SETTINGS` and falls back to `DEFAULT_SETTINGS`
    if a value is not found or is invalid in the last used section.

    Inputs:
        app_instance (App): The main application instance, providing access to its config object
                            and Tkinter variables defined in `setting_var_map`.
    Process:
        1. Iterates through `app_instance.setting_var_map`, which defines the mapping
           between Tkinter variables and config keys.
        2. For each setting, attempts to retrieve the value from `LAST_USED_SETTINGS`.
        3. If not found or parsing fails, attempts to retrieve from `DEFAULT_SETTINGS`.
        4. Converts the retrieved string value to the appropriate Python type (bool, float, int, str).
        5. Sets the corresponding Tkinter variable's value.
        6. Handles a special case for `resource_var` if no resource is found.
    Outputs: None
    """
    for var_name, (last_key, default_key, type_converter) in app_instance.setting_var_map.items():
        tk_var = getattr(app_instance, var_name)
        value_to_set = None

        # Try to load from LAST_USED_SETTINGS first
        if app_instance.config.has_option('LAST_USED_SETTINGS', last_key) and app_instance.config.get('LAST_USED_SETTINGS', last_key):
            try:
                if type_converter == bool:
                    value_to_set = app_instance.config.getboolean('LAST_USED_SETTINGS', last_key)
                elif type_converter == float:
                    value_to_set = app_instance.config.getfloat('LAST_USED_SETTINGS', last_key)
                elif type_converter == int:
                    value_to_set = app_instance.config.getint('LAST_USED_SETTINGS', last_key)
                else:
                    value_to_set = app_instance.config.get('LAST_USED_SETTINGS', last_key)
            except ValueError as e:
                print(f"Warning: Could not parse '{last_key}' from LAST_USED_SETTINGS (value: '{app_instance.config.get('LAST_USED_SETTINGS', last_key)}'). Error: {e}")
                value_to_set = None
        
        # If not found in LAST_USED_SETTINGS or empty/invalid, try DEFAULT_SETTINGS
        if value_to_set is None and default_key and app_instance.config.has_option('DEFAULT_SETTINGS', default_key):
            try:
                if type_converter == bool:
                    value_to_set = app_instance.config.getboolean('DEFAULT_SETTINGS', default_key)
                elif type_converter == float:
                    value_to_set = app_instance.config.getfloat('DEFAULT_SETTINGS', default_key)
                elif type_converter == int:
                    value_to_set = app_instance.config.getint('DEFAULT_SETTINGS', default_key)
                else:
                    value_to_set = app_instance.config.get('DEFAULT_SETTINGS', default_key)
            except ValueError as e:
                print(f"Warning: Could not parse '{default_key}' from DEFAULT_SETTINGS (value: '{app_instance.config.get('DEFAULT_SETTINGS', default_key)}'). Error: {e}")
                value_to_set = None

        if value_to_set is not None:
            tk_var.set(value_to_set)
        else:
            if var_name == 'resource_var':
                tk_var.set("No Resources Found")

def save_config(app_instance):
    """
    Saves the current state of all configurable settings from the GUI's Tkinter
    variables to the `config.ini` file, specifically in the `LAST_USED_SETTINGS` section.
    This ensures that user preferences are preserved between application sessions.

    Inputs:
        app_instance (App): The main application instance, providing access to its config object,
                            Tkinter variables, and the current window geometry.
    Process:
        1. Ensures the `LAST_USED_SETTINGS` section exists in `app_instance.config`.
        2. Iterates through `app_instance.setting_var_map` and saves the current value
           of each Tkinter variable to its corresponding `last_used_config_key`.
        3. Special handling to save the currently selected frequency bands.
        4. Saves the current window geometry.
        5. Writes the updated configuration to `app_instance.CONFIG_FILE`.
    Outputs: None
    """
    # Ensure LAST_USED_SETTINGS section exists
    if 'LAST_USED_SETTINGS' not in app_instance.config:
        app_instance.config['LAST_USED_SETTINGS'] = {}

    for var_name, (last_key, _, _) in app_instance.setting_var_map.items():
        if last_key and var_name != 'last_selected_bands_str':
            tk_var = getattr(app_instance, var_name)
            app_instance.config['LAST_USED_SETTINGS'][last_key] = str(tk_var.get())
    
    selected_band_names = [item["band"]["Band Name"] for item in app_instance.band_vars if item["var"].get()]
    app_instance.config['LAST_USED_SETTINGS']['last_selected_bands'] = ",".join(selected_band_names)
    
    app_instance.config['LAST_USED_SETTINGS']['last_window_geometry'] = app_instance.winfo_geometry()

    with open(app_instance.CONFIG_FILE, 'w') as configfile:
        app_instance.config.write(configfile)
    print(f"Configuration saved to {app_instance.CONFIG_FILE}")
