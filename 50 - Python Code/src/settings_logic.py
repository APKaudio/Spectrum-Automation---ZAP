# src/settings_logic.py
import tkinter as tk
from tkinter import messagebox
import os
import sys
import subprocess

from utils.instrument_control import set_debug_mode
# from src.config_manager import save_config # Import locally to avoid circular dependency

def restore_default_settings_logic(app_instance):
    """
    Restores all configurable settings in the GUI to their default values
    as defined in the `DEFAULT_SETTINGS` section of `config.ini`.
    This provides a quick way for users to revert to a known good configuration.

    Inputs: None
    Process:
        1. Prints a message indicating restoration.
        2. Iterates through `app_instance.setting_var_map`.
        3. For each setting, retrieves its default value from `config.ini` and converts it
           to the appropriate Python type.
        4. Sets the corresponding Tkinter variable to this default value.
        5. Handles special updates for slider widgets to reflect the new values.
        6. Resets all frequency band checkboxes to `True` (selected).
        7. Calls `update_vbw_display_logic()` to refresh the VBW based on the restored RBW.
        8. Calls `reset_setting_colors_logic()` to clear any visual indications of unapplied settings.
        9. Calls `_save_config()` to save these restored defaults as the new last-used settings.
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
    
    from src.config_manager import save_config # Import locally to avoid circular dependency
    save_config(app_instance)
    messagebox.showinfo("Settings Restored", "Default settings have been restored and saved as last used.")

def update_debug_mode_global_logic(app_instance, *args):
    set_debug_mode(app_instance.debug_mode_var.get())
    from src.config_manager import save_config # Import locally
    save_config(app_instance)

def update_scan_rbw_from_slider_index_logic(app_instance, *args):
    try:
        idx = app_instance.rbw_slider_index_var.get()
        if 0 <= idx < len(app_instance.rbw_values):
            app_instance.desired_scan_rbw_segmentation_var.set(float(app_instance.rbw_values[idx]))
        app_instance.update_vbw_display() # Added to update VBW when RBW slider changes
    except Exception as e:
        print(f"Error updating scan RBW from slider index: {e}")

def update_freq_shift_from_slider_index_logic(app_instance, *args):
    try:
        idx = app_instance.freq_shift_slider_index_var.get()
        if 0 <= idx < len(app_instance.freq_shift_values):
            app_instance.shift_freq_var.set(float(app_instance.freq_shift_values[idx]))
    except Exception as e:
        print(f"Error updating frequency shift from slider index: {e}")

def reset_setting_colors_logic(app_instance):
    for key, entry_widget in app_instance.desired_setting_entries.items():
        if isinstance(entry_widget, tk.Entry):
            entry_widget.config(fg="white")

def set_band_checkboxes_from_config_logic(app_instance):
    # Corrected variable name from last_selected_bands_str to selected_bands_str_var
    selected_bands_from_config = app_instance.selected_bands_str_var.get()
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
    Outputs: None (generates an HTML file, potentially opens a browser)
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
            # Use os.startfile for Windows, which is often more reliable for opening folders
            print(f"DEBUG (Open Folder): Attempting to open '{folder_path}' using os.startfile.")
            os.startfile(folder_path)
            print(f"✅ Opened folder: {folder_path}")
        elif sys.platform == "darwin":
            command = ['open', folder_path]
            print(f"DEBUG (Open Folder): Executing command: {command}")
            result = subprocess.run(command, capture_output=True, text=True, check=False)
            if result.returncode == 0:
                print(f"✅ Opened folder: {folder_path}")
                if result.stdout:
                    print(f"DEBUG (Open Folder) stdout: {result.stdout.strip()}")
                if result.stderr:
                    print(f"DEBUG (Open Folder) stderr: {result.stderr.strip()}")
            else:
                error_msg = f"Failed to open folder '{folder_path}'. Return code: {result.returncode}"
                if result.stdout:
                    error_msg += f"\nStdout: {result.stdout.strip()}"
                if result.stderr:
                    error_msg += f"\nStderr: {result.stderr.strip()}"
                messagebox.showerror("Error Opening Folder", error_msg)
                print(f"❌ Error opening folder: {error_msg}")
        else: # Linux and other Unix-like systems
            command = ['xdg-open', folder_path]
            print(f"DEBUG (Open Folder): Executing command: {command}")
            result = subprocess.run(command, capture_output=True, text=True, check=False)
            if result.returncode == 0:
                print(f"✅ Opened folder: {folder_path}")
                if result.stdout:
                    print(f"DEBUG (Open Folder) stdout: {result.stdout.strip()}")
                if result.stderr:
                    print(f"DEBUG (Open Folder) stderr: {result.stderr.strip()}")
            else:
                error_msg = f"Failed to open folder '{folder_path}'. Return code: {result.returncode}"
                if result.stdout:
                    error_msg += f"\nStdout: {result.stdout.strip()}"
                if result.stderr:
                    error_msg += f"\nStderr: {result.stderr.strip()}"
                messagebox.showerror("Error Opening Folder", error_msg)
                print(f"❌ Error opening folder: {error_msg}")

    except FileNotFoundError:
        messagebox.showerror("Error Opening Folder", f"Command not found to open folder. Please ensure 'explorer' (Windows), 'open' (macOS), or 'xdg-open' (Linux) is in your PATH.")
        print(f"❌ Command not found error when trying to open folder: {sys.exc_info()[1]}")
    except Exception as e:
        messagebox.showerror("Error Opening Folder", f"An unexpected error occurred while trying to open folder '{folder_path}': {e}")
        print(f"❌ An unexpected error occurred opening folder: {e}")
