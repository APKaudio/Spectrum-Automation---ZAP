# src/config_manager.py
import configparser
import os
import tkinter as tk # Import tkinter for update_idletasks
from utils.instrument_control import debug_print # Import debug_print

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
        7. Calls `_populate_vars_from_config()` to set Tkinter variables based on loaded config
    Outputs: None
    """
    debug_print(f"Attempting to load configuration from {app_instance.CONFIG_FILE}...")
    
    # Ensure the config file exists, create with defaults if not
    if not os.path.exists(app_instance.CONFIG_FILE):
        debug_print(f"Config file '{app_instance.CONFIG_FILE}' not found. Creating with default settings.")
        app_instance.config['DEFAULT_SETTINGS'] = {
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
            'default_window_geometry': '1400x780+100+100',
            'default_scan_directory': 'scan_data',
            'default_scan_name': 'MyScan',
            'default_scan_rbw_segmentation': '100000', # New default
            'default_default_focus_width': '10000.0' # New default
        }
        # Dynamically generate default_selected_bands
        default_bands_str = ",".join([band["Band Name"] for band in app_instance.SCAN_BAND_RANGES])
        app_instance.config['DEFAULT_SETTINGS']['default_selected_bands'] = default_bands_str
        
        app_instance.config['LAST_USED_SETTINGS'] = {} # Initialize empty last used settings
        # Copy defaults to last used if no last used exists
        for key, value in app_instance.config['DEFAULT_SETTINGS'].items():
            app_instance.config['LAST_USED_SETTINGS'][key] = value

        with open(app_instance.CONFIG_FILE, 'w') as configfile:
            app_instance.config.write(configfile)
        debug_print(f"Created default config file at {app_instance.CONFIG_FILE}")

    # Read the configuration file
    app_instance.config.read(app_instance.CONFIG_FILE)

    # Ensure sections exist
    if 'DEFAULT_SETTINGS' not in app_instance.config:
        app_instance.config['DEFAULT_SETTINGS'] = {}
    if 'LAST_USED_SETTINGS' not in app_instance.config:
        app_instance.config['LAST_USED_SETTINGS'] = {}

    # Populate default settings if missing
    defaults_to_ensure = {
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
        'default_window_geometry': '1400x780+100+100',
        'default_scan_directory': 'scan_data',
        'default_scan_name': 'MyScan',
        'default_scan_rbw_segmentation': '100000', # New default
        'default_default_focus_width': '10000.0' # New default
    }
    for key, value in defaults_to_ensure.items():
        if key not in app_instance.config['DEFAULT_SETTINGS']:
            app_instance.config['DEFAULT_SETTINGS'][key] = value

    # Ensure last_window_geometry is set in LAST_USED_SETTINGS
    if 'last_window_geometry' not in app_instance.config['LAST_USED_SETTINGS']:
        app_instance.config['LAST_USED_SETTINGS']['last_window_geometry'] = \
            app_instance.config['DEFAULT_SETTINGS'].get('default_window_geometry', '1400x780+100+100')

    # Populate Tkinter variables from LAST_USED_SETTINGS or DEFAULT_SETTINGS
    _populate_vars_from_config(app_instance)

    debug_print("Configuration loaded.")


def _populate_vars_from_config(app_instance):
    """
    Internal helper function to set Tkinter variables based on loaded config.
    Prioritizes LAST_USED_SETTINGS, falls back to DEFAULT_SETTINGS.
    """
    for var_name, (last_key, default_key, tk_var) in app_instance.setting_var_map.items():
        value = None
        if last_key and last_key in app_instance.config['LAST_USED_SETTINGS']:
            value = app_instance.config['LAST_USED_SETTINGS'][last_key]
            debug_print(f"Found '{last_key}' in LAST_USED_SETTINGS: '{value}'")
        elif default_key and default_key in app_instance.config['DEFAULT_SETTINGS']:
            value = app_instance.config['DEFAULT_SETTINGS'][default_key]
            debug_print(f"Found '{default_key}' in DEFAULT_SETTINGS: '{value}'")

        if value is not None:
            # Type conversion based on Tkinter variable type
            try:
                if isinstance(tk_var, tk.BooleanVar):
                    tk_var.set(value.lower() == 'true')
                elif isinstance(tk_var, tk.DoubleVar): # For float values
                    tk_var.set(float(value))
                elif isinstance(tk_var, tk.IntVar): # For integer values
                    tk_var.set(int(value))
                else: # Default to StringVar
                    tk_var.set(value)
                debug_print(f"Set {var_name} to '{tk_var.get()}' (from config)")
            except ValueError as e:
                debug_print(f"ERROR: Could not convert config value '{value}' for {var_name} (expected {type(tk_var).__name__}): {e}. Using default/current.")
                # If conversion fails, the variable retains its initial value or default.

    # Special handling for selected bands (list of strings)
    selected_bands_str = app_instance.config['LAST_USED_SETTINGS'].get('last_selected_bands',
                                                                        app_instance.config['DEFAULT_SETTINGS'].get('default_selected_bands', ''))
    selected_band_names = [name.strip() for name in selected_bands_str.split(',') if name.strip()]
    
    for band_item in app_instance.band_vars:
        band_name = band_item["band"]["Band Name"]
        band_item["var"].set(band_name in selected_band_names)
    debug_print(f"Set selected bands to: {selected_band_names}")

    # Update window geometry
    last_geometry = app_instance.config['LAST_USED_SETTINGS'].get('last_window_geometry')
    if last_geometry:
        try:
            app_instance.geometry(last_geometry)
            debug_print(f"Set window geometry to {last_geometry}")
        except TclError as e:
            debug_print(f"WARNING: Could not apply saved window geometry '{last_geometry}': {e}. Using default.")


def save_config(app_instance):
    """
    Saves the current application settings from Tkinter variables to `config.ini`.
    This ensures that the last-used settings are preserved across application sessions.

    Inputs:
        app_instance (App): The main application instance, providing access to its config object
                            and Tkinter variables.
    Outputs: None
    """
    debug_print("Saving configuration...")
    for var_name, (last_key, _, _) in app_instance.setting_var_map.items():
        if last_key and var_name != 'selected_bands_str_var': # Skip the string version as we handle bands separately
            tk_var = getattr(app_instance, var_name)
            app_instance.config['LAST_USED_SETTINGS'][last_key] = str(tk_var.get())
            debug_print(f"Saving '{last_key}': '{tk_var.get()}'")
    
    # Save selected bands
    selected_band_names = [item["band"]["Band Name"] for item in app_instance.band_vars if item["var"].get()]
    bands_to_save = ",".join(selected_band_names)
    app_instance.config['LAST_USED_SETTINGS']['last_selected_bands'] = bands_to_save
    debug_print(f"Saving 'last_selected_bands': '{bands_to_save}'")
    
    # Force update of window geometry before saving
    app_instance.update_idletasks() # IMPORTANT: Ensure geometry is up-to-date
    current_geometry = app_instance.winfo_geometry()
    app_instance.config['LAST_USED_SETTINGS']['last_window_geometry'] = current_geometry
    debug_print(f"Saving 'last_window_geometry': '{current_geometry}'")

    try:
        with open(app_instance.CONFIG_FILE, 'w') as configfile:
            app_instance.config.write(configfile)
        debug_print(f"Configuration saved successfully to {app_instance.CONFIG_FILE}")
    except IOError as e:
        debug_print(f"ERROR: Could not save configuration to {app_instance.CONFIG_FILE}: {e}")
        tk.messagebox.showerror("Config Save Error", f"Could not save configuration: {e}")

