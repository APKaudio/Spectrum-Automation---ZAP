# src/settings_logic.py
import tkinter as tk
from tkinter import messagebox
import os
import sys
import subprocess

from utils.instrument_control import set_debug_mode

def restore_default_settings_logic(app_instance):
    """
    Restores all configurable settings in the GUI to their default values
    as defined in the `DEFAULT_SETTINGS` section of `config.ini`.
    This provides a quick way for users to revert to a known good configuration.

    Inputs:
        app_instance (App): The main application instance, providing access to Tkinter variables and other attributes.
    Process:
        1. Prints a message indicating restoration.
        2. Iterates through `app_instance.setting_var_map`.
        3. For each setting, retrieves its default value from `app_instance.config` and converts it
           to the appropriate Python type.
        4. Sets the corresponding Tkinter variable to this default value.
        5. Handles special updates for slider widgets to reflect the new values.
        6. Resets all frequency band checkboxes to `True` (selected).
        7. Calls `update_vbw_display_logic()` to refresh the VBW based on the restored RBW.
        8. Calls `reset_setting_colors_logic()` to clear any visual indications of unapplied settings.
        9. Calls `save_config()` to save these restored defaults as the new last-used settings.
        10. Displays an informational messagebox upon completion.
    Outputs: None
    """
    print("Restoring default settings...")
    for var_name, (_, default_key, type_converter) in app_instance.setting_var_map.items():
        if default_key and app_instance.config.has_option('DEFAULT_SETTINGS', default_key):
            tk_var = getattr(app_instance, var_name)
            try:
                if type_converter == bool:
                    tk_var.set(app_instance.config.getboolean('DEFAULT_SETTINGS', default_key))
                elif type_converter == float:
                    tk_var.set(app_instance.config.getfloat('DEFAULT_SETTINGS', default_key))
                elif type_converter == int:
                    tk_var.set(app_instance.config.getint('DEFAULT_SETTINGS', default_key))
                else:
                    tk_var.set(app_instance.config.get('DEFAULT_SETTINGS', default_key))
                print(f"  Restored {default_key} to {tk_var.get()}")
            except ValueError as e:
                print(f"Error restoring default for {default_key}: {e}. Skipping.")

    app_instance.rbw_slider_index_var.set(app_instance.rbw_val_to_idx.get(int(app_instance.desired_scan_rbw_segmentation_var.get()), 0))
    app_instance.freq_shift_slider_index_var.set(app_instance.freq_shift_val_to_idx.get(int(app_instance.shift_freq_var.get()), 0))
    
    app_instance.cycle_hold_time_slider.set(app_instance.desired_max_hold_time_var.get())
    app_instance.cycle_wait_time_slider.set(app_instance.desired_cycle_wait_time_var.get())

    for item in app_instance.band_vars:
        item["var"].set(True)

    update_vbw_display_logic(app_instance)
    reset_setting_colors_logic(app_instance)
    
    from src.config_manager import save_config
    save_config(app_instance)
    messagebox.showinfo("Settings Restored", "Default settings have been restored and saved as last used.")

def update_debug_mode_global_logic(app_instance, *args):
    """
    Updates the global debug mode variable in the `instrument_control` module
    whenever the state of the debug mode checkbox in the GUI changes.
    This allows for dynamic enabling/disabling of verbose logging for VISA commands.

    Inputs:
        app_instance (App): The main application instance, providing access to Tkinter variables.
        *args: Standard Tkinter trace arguments (not used).
    Process:
        1. Retrieves the current boolean value from `app_instance.debug_mode_var`.
        2. Calls `set_debug_mode()` from `instrument_control` to update the global flag.
        3. Calls `save_config()` to immediately persist the debug mode setting.
    Outputs: None
    """
    set_debug_mode(app_instance.debug_mode_var.get())
    from src.config_manager import save_config
    save_config(app_instance)

def update_scan_rbw_from_slider_index_logic(app_instance, *args):
    """
    Updates the `desired_scan_rbw_segmentation_var` (which controls the instrument's RBW)
    based on the position of the RBW slider. This links the visual slider
    to the numerical RBW setting.

    Inputs:
        app_instance (App): The main application instance, providing access to Tkinter variables and RBW values.
        *args: Standard Tkinter trace arguments (not used).
    Process:
        1. Retrieves the current index from `app_instance.rbw_slider_index_var`.
        2. Uses the index to look up the corresponding RBW value from `app_instance.rbw_values`.
        3. Sets the `app_instance.desired_scan_rbw_segmentation_var` to this float value.
        4. Includes error handling for invalid index access.
    Outputs: None
    """
    try:
        idx = app_instance.rbw_slider_index_var.get()
        if 0 <= idx < len(app_instance.rbw_values):
            app_instance.desired_scan_rbw_segmentation_var.set(float(app_instance.rbw_values[idx]))
    except Exception as e:
        print(f"Error updating scan RBW from slider index: {e}")

def update_freq_shift_from_slider_index_logic(app_instance, *args):
    """
    Updates the `shift_freq_var` (which controls the frequency offset)
    based on the position of the Frequency Shift slider. This links the visual slider
    to the numerical frequency offset setting.

    Inputs:
        app_instance (App): The main application instance, providing access to Tkinter variables and frequency shift values.
        *args: Standard Tkinter trace arguments (not used).
    Process:
        1. Retrieves the current index from `app_instance.freq_shift_slider_index_var`.
        2. Uses the index to look up the corresponding frequency shift value from `app_instance.freq_shift_values`.
        3. Sets the `app_instance.shift_freq_var` to this float value.
        4. Includes error handling for invalid index access.
    Outputs: None
    """
    try:
        idx = app_instance.freq_shift_slider_index_var.get()
        if 0 <= idx < len(app_instance.freq_shift_values):
            app_instance.shift_freq_var.set(float(app_instance.freq_shift_values[idx]))
    except Exception as e:
        print(f"Error updating frequency shift from slider index: {e}")

def reset_setting_colors_logic(app_instance):
    """
    Resets the foreground color of all input entry widgets in the "Scan Configuration"
    section to white. This is typically called after settings are successfully applied
    to provide visual feedback that the settings are no longer "dirty" or unapplied.

    Inputs:
        app_instance (App): The main application instance, providing access to `desired_setting_entries`.
    Process:
        1. Iterates through the `app_instance.desired_setting_entries` dictionary.
        2. For each entry that is a `tk.Entry` widget, sets its `fg` (foreground)
           color to "white".
    Outputs: None
    """
    for key, entry_widget in app_instance.desired_setting_entries.items():
        if isinstance(entry_widget, tk.Entry):
            entry_widget.config(fg="white")

def set_band_checkboxes_from_config_logic(app_instance):
    """
    Sets the state (checked/unchecked) of the frequency band checkboxes
    based on the `last_selected_bands` setting loaded from `config.ini`.
    If no last selected bands are found, all checkboxes are set to `True` (selected).

    Inputs:
        app_instance (App): The main application instance, providing access to `last_selected_bands_str` and `band_vars`.
    Process:
        1. Retrieves the string of comma-separated band names from `app_instance.last_selected_bands_str`.
        2. Splits the string into a list of selected band names.
        3. Iterates through `app_instance.band_vars` (which holds each band's name and its Tkinter `BooleanVar`).
        4. If a band's name is found in the list of selected band names, its `BooleanVar` is set to `True`.
        5. Otherwise, it's set to `False`.
        6. If no `last_selected_bands` are found in config, all bands are set to `True`.
    Outputs: None
    """
    selected_bands_from_config = app_instance.last_selected_bands_str.get()
    if selected_bands_from_config:
        selected_band_names = [name.strip() for name in selected_bands_from_config.split(',') if name.strip()]
        for item in app_instance.band_vars:
            band_name = item["band"]["Band Name"]
            if band_name in selected_band_names:
                item["var"].set(True)
            else:
                item["var"].set(False)
    else:
        for item in app_instance.band_vars:
            item["var"].set(True)

def update_vbw_display_logic(app_instance):
    """
    Updates the displayed Video Bandwidth (VBW) value based on the
    current Resolution Bandwidth (RBW) setting. The VBW is typically
    set to approximately 1/3 of the RBW for spectrum analyzers.

    Inputs:
        app_instance (App): The main application instance, providing access to `desired_scan_rbw_segmentation_var` and `desired_vbw_display_var`.
    Process:
        1. Retrieves the current RBW value from `app_instance.desired_scan_rbw_segmentation_var`.
        2. Calculates VBW as RBW / 3.
        3. Sets the `app_instance.desired_vbw_display_var` to the calculated VBW (as an integer string).
        4. Handles `ValueError` if the RBW input is not a valid number.
    Outputs: None
    """
    try:
        scan_rbw_val = float(app_instance.desired_scan_rbw_segmentation_var.get())
        app_instance.desired_vbw_display_var.set(str(int(scan_rbw_val / 3)))
    except ValueError:
        app_instance.desired_vbw_display_var.set("Invalid RBW")

def open_output_folder_logic(app_instance):
    """
    Opens the specified output folder in the user's default file explorer.
    This provides a convenient way for users to access their saved scan data and plots.

    Inputs:
        app_instance (App): The main application instance, providing access to `output_folder_var`.
    Process:
        1. Retrieves the output folder path from `app_instance.output_folder_var`.
        2. Converts the path to an absolute path if it's relative.
        3. Checks if the folder exists, showing a warning if not.
        4. Uses `os.startfile` (Windows), `subprocess.run(['open', ...])` (macOS),
           or `subprocess.run(['xdg-open', ...])` (Linux) to open the folder.
        5. Prints success or error messages to the console and displays a messagebox on error.
    Outputs: None
    """
    folder_path = app_instance.output_folder_var.get()
    if not os.path.isabs(folder_path):
        folder_path = os.path.join(os.getcwd(), folder_path)

    if not os.path.exists(folder_path):
        messagebox.showwarning("Folder Not Found", f"The folder '{folder_path}' does not exist.")
        print(f"🚫 Folder not found: {folder_path}")
        return
    
    try:
        if sys.platform == "win32":
            os.startfile(folder_path)
        elif sys.platform == "darwin":
            subprocess.run(['open', folder_path])
        else:
            subprocess.run(['xdg-open', folder_path])
        print(f"✅ Opened folder: {folder_path}")
    except Exception as e:
        messagebox.showerror("Error Opening Folder", f"Failed to open folder '{folder_path}': {e}")
        print(f"❌ Error opening folder: {e}")
