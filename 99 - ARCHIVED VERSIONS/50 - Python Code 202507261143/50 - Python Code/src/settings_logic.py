# src/settings_logic.py
import tkinter as tk
from tkinter import messagebox
import os
import sys
import subprocess

from utils.instrument_control import set_debug_mode, debug_print
from src.config_manager import save_config # Import save_config locally to avoid circular dependency

def restore_default_settings_logic(app_instance):
    """
    Restores all configurable settings in the GUI to their default values
    as defined in the `DEFAULT_SETTINGS` section of `config.ini`.
    This provides a quick way for users to revert to a known good configuration.

    Inputs:
        app_instance (App): The main application instance, providing access to its config object,
                            Tkinter variables, and other necessary attributes.
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
        9. Calls `_save_config()` to save the restored defaults as the new last-used settings.
    Outputs: None
    """
    debug_print("Restoring default settings...")
    
    # Restore settings from DEFAULT_SETTINGS section
    for var_name, (last_key, default_key, tk_var) in app_instance.setting_var_map.items():
        if default_key and default_key in app_instance.config['DEFAULT_SETTINGS']:
            default_value_str = app_instance.config['DEFAULT_SETTINGS'][default_key]
            try:
                if isinstance(tk_var, tk.BooleanVar):
                    tk_var.set(default_value_str.lower() == 'true')
                elif isinstance(tk_var, tk.DoubleVar):
                    tk_var.set(float(default_value_str))
                elif isinstance(tk_var, tk.IntVar):
                    tk_var.set(int(default_value_str))
                else: # StringVar
                    tk_var.set(default_value_str)
                debug_print(f"  Restored {default_key} to {tk_var.get()}")
            except ValueError as e:
                debug_print(f"  WARNING: Could not restore {default_key} with value '{default_value_str}': {e}")
        else:
            debug_print(f"  WARNING: Default setting '{default_key}' not found in config.ini. Skipping.")

    # Special handling for sliders to update their positions
    app_instance._set_initial_slider_positions()

    # Reset all frequency band checkboxes to True (selected)
    for band_item in app_instance.band_vars:
        band_item["var"].set(True)
    debug_print("  Restored all frequency band selections to default (selected).")

    # Update VBW display
    # FIX: Corrected method name from _update_vbw_display_callback to update_vbw_display_logic
    app_instance.update_vbw_display_logic() 

    # Reset setting colors to indicate they are now default/applied
    app_instance.reset_setting_colors_logic()

    # Ensure debug mode and log visa commands are updated immediately
    set_debug_mode(app_instance.general_debug_enabled_var.get())
    set_log_visa_commands_mode(app_instance.log_visa_commands_enabled_var.get())

    # Save the restored defaults as the new last-used settings
    save_config(app_instance)
    messagebox.showinfo("Settings Restored", "All settings have been restored to their default values and saved.")


def open_output_folder_logic(output_folder_path):
    """
    Opens the specified output folder in the system's file explorer.
    """
    if not os.path.exists(output_folder_path):
        messagebox.showwarning("Folder Not Found", f"The output folder '{output_folder_path}' does not exist.")
        print(f"🚫 Output folder not found: {output_folder_path}")
        return
    
    try:
        if sys.platform == "win32":
            os.startfile(output_folder_path)
            print(f"✅ Opened output folder: {output_folder_path}")
        elif sys.platform == "darwin":
            subprocess.run(['open', output_folder_path], check=True)
            print(f"✅ Opened output folder: {output_folder_path}")
        else: # Linux
            subprocess.run(['xdg-open', output_folder_path], check=True)
            print(f"✅ Opened output folder: {output_folder_path}")
    except FileNotFoundError:
        messagebox.showerror("Error Opening Folder", f"Command not found to open folder. Please ensure 'explorer' (Windows), 'open' (macOS), or 'xdg-open' (Linux) is in your PATH.")
        print(f"❌ Command not found error when trying to open output folder: {sys.exc_info()[1]}")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error Opening Folder", f"Failed to open output folder '{output_folder_path}': {e}")
        print(f"❌ Error opening output folder: {e}")
    except Exception as e:
        messagebox.showerror("Error Opening Folder", f"An unexpected error occurred while trying to open the output folder: {e}")
        print(f"❌ An unexpected error occurred while trying to open the output folder: {e}")


def open_preset_folder_logic():
    """
    Opens the default instrument preset folder (C:\PRESETS) in the system's file explorer.
    """
    preset_folder_path = "C:\\PRESETS" # Hardcoded path for instrument presets
    
    if not os.path.exists(preset_folder_path):
        messagebox.showwarning("Folder Not Found", f"The instrument preset folder '{preset_folder_path}' does not exist.")
        print(f"🚫 Instrument preset folder not found: {preset_folder_path}")
        return

    try:
        if sys.platform == "win32":
            os.startfile(preset_folder_path)
            print(f"✅ Opened instrument preset folder: {preset_folder_path}")
        elif sys.platform == "darwin":
            subprocess.run(['open', preset_folder_path], check=True)
            print(f"✅ Opened instrument preset folder: {preset_folder_path}")
        else: # Linux
            subprocess.run(['xdg-open', preset_folder_path], check=True)
            print(f"✅ Opened instrument preset folder: {preset_folder_path}")
    except FileNotFoundError:
        messagebox.showerror("Error Opening Folder", f"Command not found to open folder. Please ensure 'explorer' (Windows), 'open' (macOS), or 'xdg-open' (Linux) is in your PATH.")
        print(f"❌ Command not found error when trying to open preset folder: {sys.exc_info()[1]}")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error Opening Folder", f"Failed to open preset folder '{preset_folder_path}': {e}")
        print(f"❌ Error opening preset folder: {e}")
    except Exception as e:
        messagebox.showerror("Error Opening Folder", f"An unexpected error occurred while trying to open the preset folder: {e}")
        print(f"❌ An unexpected error occurred while trying to open the preset folder: {e}")

# Renamed from _update_vbw_display_callback to update_vbw_display_logic in main_app.py
# This function is now directly called from App instance.
# Keeping a placeholder here if it was intended to be a separate logic function,
# but the error indicates it's a method of app_instance.
# def update_vbw_display_logic(app_instance):
#    app_instance.update_vbw_display_logic()


# The following functions are now methods of the App class and are called directly.
# They are kept here for reference if they were intended to be standalone logic functions
# but the traceback suggests they are methods of the app_instance.

def update_scan_rbw_from_slider_index_logic(app_instance, val):
    # This function is now a method of App class
    app_instance.update_scan_rbw_from_slider_index_logic(val)

def update_max_hold_time_from_slider_index_logic(app_instance, val):
    # This function is now a method of App class
    app_instance.update_max_hold_time_from_slider_index_logic(val)

def update_cycle_wait_time_from_slider_index_logic(app_instance, val):
    # This function is now a method of App class
    app_instance.update_cycle_wait_time_from_slider_index_logic(val)
