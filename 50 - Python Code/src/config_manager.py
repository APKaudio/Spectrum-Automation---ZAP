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
    print("DEBUG: Loading configuration...")
    app_instance.config = configparser.ConfigParser()
    
    # Ensure default sections exist
    if not app_instance.config.has_section('DEFAULT_SETTINGS'):
        app_instance.config.add_section('DEFAULT_SETTINGS')
    if not app_instance.config.has_section('LAST_USED_SETTINGS'):
        app_instance.config.add_section('LAST_USED_SETTINGS')

    # Define default values for DEFAULT_SETTINGS if they don't exist
    # These are hardcoded defaults if config.ini is completely missing or malformed
    app_instance.config.setdefault('DEFAULT_SETTINGS', {})
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_rbw_step_size_hz', '1000000')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_cycle_wait_time_seconds', '0')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_maxhold_time_seconds', '3')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_scan_rbw_hz', '10000') # This is the segmentation RBW
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_reference_level_dbm', '-40')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_freq_shift_hz', '0')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_maxhold_enabled', 'True')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_include_gov_markers', 'True')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_include_tv_markers', 'True')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_open_html_after_complete', 'True')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_high_sensitivity', 'True')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_preamp_on', 'True')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_debug_mode', 'False')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_window_geometry', '1400x780+100+100')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_scan_directory', 'scan_data')
    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_scan_name', 'MyScan')
    
    # Dynamically generate default_selected_bands from SCAN_BAND_RANGES
    # Ensure SCAN_BAND_RANGES is accessible, e.g., passed to app_instance or imported
    if hasattr(app_instance, 'SCAN_BAND_RANGES'):
        default_bands_str = ",".join([band["Band Name"] for band in app_instance.SCAN_BAND_RANGES])
        app_instance.config['DEFAULT_SETTINGS'].setdefault('default_selected_bands', default_bands_str)
    else:
        app_instance.config['DEFAULT_SETTINGS'].setdefault('default_selected_bands', 'Low VHF+FM,High VHF+216,UHF 400-500') # Fallback
        print("WARNING: app_instance.SCAN_BAND_RANGES not found. Using limited default bands.")

    app_instance.config['DEFAULT_SETTINGS'].setdefault('default_default_focus_width', '10000.0') # Corrected default key

    # Read the config file
    try:
        app_instance.config.read(app_instance.CONFIG_FILE)
        print(f"DEBUG: Config file '{app_instance.CONFIG_FILE}' read successfully.")
    except Exception as e:
        print(f"WARNING: Could not read config file '{app_instance.CONFIG_FILE}'. Using default settings. Error: {e}")

    # Ensure last_window_geometry is set in LAST_USED_SETTINGS
    if 'last_window_geometry' not in app_instance.config['LAST_USED_SETTINGS']:
        app_instance.config['LAST_USED_SETTINGS']['last_window_geometry'] = app_instance.config['DEFAULT_SETTINGS'].get('default_window_geometry', '1400x780+100+100')
        print("DEBUG: Set last_window_geometry from default.")

    _populate_vars_from_config(app_instance)
    print("DEBUG: Configuration loaded.")


def _populate_vars_from_config(app_instance):
    """
    Populates Tkinter variables in the main application instance with values
    from the loaded configuration (LAST_USED_SETTINGS or DEFAULT_SETTINGS).

    Inputs:
        app_instance (App): The main application instance, containing Tkinter variables
                            and the loaded config object.
    Process:
        1. Iterates through `app_instance.setting_var_map`.
        2. For each setting, attempts to retrieve its value from `LAST_USED_SETTINGS`.
        3. If not found in `LAST_USED_SETTINGS`, falls back to `DEFAULT_SETTINGS`.
        4. Converts the retrieved string value to the appropriate Python type
           (float, bool, str) and sets the corresponding Tkinter variable.
        5. Handles `ValueError` during conversion by falling back to the default value.
        6. Handles special cases for selected frequency bands.
    Outputs: None (modifies Tkinter variables directly)
    """
    print("DEBUG: Populating Tkinter variables from config...")
    for var_name, (last_key, default_key, type_converter) in app_instance.setting_var_map.items():
        tk_var = getattr(app_instance, var_name)
        
        # Get value from LAST_USED_SETTINGS, or fallback to DEFAULT_SETTINGS
        value = app_instance.config['LAST_USED_SETTINGS'].get(last_key)
        if value is None:
            value = app_instance.config['DEFAULT_SETTINGS'].get(default_key)

        print(f"DEBUG: Processing '{var_name}' (last_key='{last_key}', default_key='{default_key}'). Retrieved value: '{value}' (Type: {type(value)})")

        if value is not None:
            try:
                if type_converter == bool:
                    tk_var.set(value.lower() == 'true')
                else:
                    tk_var.set(type_converter(value))
            except ValueError as e:
                print(f"ERROR: Error converting config value for {var_name} (key: {last_key}): '{value}' to {type_converter.__name__}. Error: {e}")
                # Fallback to default if conversion fails (already tried, but re-confirming)
                default_value = app_instance.config['DEFAULT_SETTINGS'].get(default_key)
                if default_value is not None:
                    try:
                        if type_converter == bool:
                            tk_var.set(default_value.lower() == 'true')
                        else:
                            tk_var.set(type_converter(default_value))
                        print(f"DEBUG: Set {var_name} to default: '{tk_var.get()}'")
                    except ValueError as e_default:
                        print(f"CRITICAL ERROR: Could not set default value for {var_name}: '{default_value}'. Error: {e_default}")
                else:
                    print(f"CRITICAL ERROR: No default value found for {var_name}.")
            except Exception as e:
                print(f"CRITICAL ERROR: An unexpected error occurred while setting {var_name}: {e}")
        else:
            print(f"WARNING: No value found for {var_name} (key: {last_key} or {default_key}). Tkinter variable not set.")

    # Special handling for frequency band checkboxes
    selected_bands_str = app_instance.selected_bands_str_var.get()
    selected_band_names = [name.strip() for name in selected_bands_str.split(',') if name.strip()]
    
    # Ensure app_instance.band_vars is initialized
    if not hasattr(app_instance, 'band_vars'):
        app_instance.band_vars = [] # Initialize if not present

    for band_item in app_instance.band_vars:
        band_name = band_item["band"]["Band Name"]
        band_item["var"].set(band_name in selected_band_names)
    print("DEBUG: Frequency band checkboxes updated.")


def save_config(app_instance):
    """
    Saves the current state of Tkinter variables to the `LAST_USED_SETTINGS`
    section of the `config.ini` file. This ensures that user preferences
    are preserved across application sessions.

    Inputs:
        app_instance (App): The main application instance, providing access to its config object,
                            Tkinter variables, and other necessary attributes.
    Process:
        1. Iterates through `app_instance.setting_var_map`.
        2. Retrieves the current value from each Tkinter variable.
        3. Stores the value as a string in the `LAST_USED_SETTINGS` section.
        4. Handles special saving for selected frequency bands by joining their names.
        5. Updates and saves the window geometry.
        6. Writes the updated configuration to `config.ini`.
    Outputs: None
    """
    print("DEBUG: save_config called.")
    print("DEBUG: Saving configuration...")
    for var_name, (last_key, _, _) in app_instance.setting_var_map.items():
        if last_key and var_name != 'last_selected_bands_str': # last_selected_bands_str is handled separately
            tk_var = getattr(app_instance, var_name)
            app_instance.config['LAST_USED_SETTINGS'][last_key] = str(tk_var.get())
            print(f"DEBUG: Saving '{last_key}': '{tk_var.get()}'")
    
    # Save selected bands
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
        print(f"ERROR: Could not save configuration to {app_instance.CONFIG_FILE}: {e}")
        tk.messagebox.showerror("Config Save Error", f"Could not save configuration: {e}")
    except Exception as e:
        print(f"ERROR: An unexpected error occurred during config save: {e}")
        tk.messagebox.showerror("Config Save Error", f"An unexpected error occurred during config save: {e}")

