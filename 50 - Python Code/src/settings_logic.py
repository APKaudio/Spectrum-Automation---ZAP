# src/settings_logic.py
import tkinter as tk
from tkinter import messagebox
import inspect
import os # Import os for path manipulation

# Import necessary functions/constants from utils
from utils.instrument_control import debug_print
from src.config_manager import load_config, save_config # Import save_config

def restore_default_settings_logic(app_instance):
    """
    Restores all application settings to their default values as defined in config.ini.
    This function updates the Tkinter variables and then saves these defaults to
    the LAST_USED_SETTINGS section, effectively resetting the user's saved preferences.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    debug_print("Attempting to restore default settings...", file=current_file, function=current_function)

    # Iterate through the setting_var_map and set each Tkinter variable
    # to its default value from the DEFAULT_SETTINGS section.
    for tk_var_name, (last_key, default_key, tk_var_instance) in app_instance.setting_var_map.items():
        if default_key: # Only restore if a default_key is defined
            default_value_str = app_instance.config.get('DEFAULT_SETTINGS', default_key, fallback=None)
            if default_value_str is not None:
                try:
                    # Convert to appropriate type based on Tkinter variable type
                    if isinstance(tk_var_instance, tk.BooleanVar):
                        # configparser reads booleans as strings "True"/"False"
                        tk_var_instance.set(default_value_str.lower() == 'true')
                    elif isinstance(tk_var_instance, tk.IntVar):
                        tk_var_instance.set(int(default_value_str))
                    elif isinstance(tk_var_instance, tk.DoubleVar): # For float values if used
                        tk_var_instance.set(float(default_value_str))
                    else: # Default to StringVar
                        tk_var_instance.set(default_value_str)
                    debug_print(f"Restored '{last_key}' to default: '{default_value_str}'", file=current_file, function=current_function)
                except ValueError as e:
                    debug_print(f"Error converting default value for '{last_key}' ('{default_value_str}'): {e}", file=current_file, function=current_function)
            else:
                debug_print(f"No default value found for '{default_key}' in config.ini.", file=current_file, function=current_function)
        else:
            debug_print(f"Skipping '{tk_var_name}': No default key defined.", file=current_file, function=current_function)

    # Special handling for selected bands: restore to default (all selected)
    default_selected_bands_str = app_instance.config.get('DEFAULT_SETTINGS', 'default_selected_bands', fallback='')
    if default_selected_bands_str:
        default_selected_band_names = [name.strip() for name in default_selected_bands_str.split(',') if name.strip()]
        for band_item in app_instance.band_vars:
            band_item["var"].set(band_item["band"]["Band Name"] in default_selected_band_names)
        debug_print(f"Restored selected bands to default: {default_selected_band_names}", file=current_file, function=current_function)
    else:
        # If no default_selected_bands in config, ensure all are selected
        for band_item in app_instance.band_vars:
            band_item["var"].set(True)
        debug_print("No default_selected_bands in config.ini. All bands set to selected.", file=current_file, function=current_function)

    # Update sliders to reflect new default values
    app_instance._set_initial_slider_positions()
    
    # Reset setting colors (e.g., remove any "changed" highlights)
    app_instance.reset_setting_colors_logic()

    # Save the current (now default) settings to LAST_USED_SETTINGS
    save_config(app_instance)

    messagebox.showinfo("Settings Restored", "All settings have been restored to their default values and saved.")
    debug_print("Default settings restored and saved.", file=current_file, function=current_function)


# Removed the duplicated reset_scan_buttons_logic from here.
# It is now solely managed in src/scan_logic.py
