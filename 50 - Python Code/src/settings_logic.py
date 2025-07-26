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
        9. Calls `_save_config()` to save these restored defaults as the new last-used settings.
        10. Displays an informational messagebox upon completion.
    Outputs: None
    """
    debug_print("Restoring default settings...")
    
    # Re-read config to ensure we have the latest default values
    app_instance.config.read(app_instance.CONFIG_FILE)
    default_settings = app_instance.config['DEFAULT_SETTINGS']

    for var_name, (last_key, default_key, tk_var) in app_instance.setting_var_map.items():
        if default_key and default_key in default_settings:
            value = default_settings[default_key]
            try:
                if isinstance(tk_var, tk.BooleanVar):
                    tk_var.set(value.lower() == 'true')
                elif isinstance(tk_var, tk.DoubleVar):
                    tk_var.set(float(value))
                elif isinstance(tk_var, tk.IntVar):
                    tk_var.set(int(value))
                else: # Default to StringVar
                    tk_var.set(value)
                debug_print(f" Restored {default_key} to {tk_var.get()}")
            except ValueError as e:
                debug_print(f"ERROR: Could not convert default value '{value}' for {default_key}: {e}")
                
    # Reset all frequency band checkboxes to True (selected)
    for band_item in app_instance.band_vars:
        band_item["var"].set(True)
    debug_print(" Restored all frequency band selections to default (selected).")

    # Manually trigger VBW update after RBW is restored
    app_instance._update_vbw_display_callback()

    # Reset slider positions to reflect restored values
    try:
        app_instance.rbw_slider_index_var.set(app_instance.rbw_val_to_idx.get(int(float(app_instance.desired_rbw_var.get())), 0))
        app_instance.max_hold_time_slider_index_var.set(app_instance.max_hold_time_val_to_idx.get(int(float(app_instance.desired_max_hold_time_var.get())), 0))
        
        current_cycle_wait = float(app_instance.desired_cycle_wait_time_var.get())
        closest_cycle_wait_idx = min(range(len(app_instance.cycle_wait_time_values)), 
                                     key=lambda i: abs(app_instance.cycle_wait_time_values[i] - current_cycle_wait))
        app_instance.cycle_wait_time_slider_index_var.set(closest_cycle_wait_idx)

        debug_print(" Sliders reset to default positions.")
    except Exception as e:
        debug_print(f"ERROR: Could not reset slider positions after restoring defaults: {e}")


    # Reset setting colors (if implemented)
    app_instance.reset_setting_colors_logic()

    # Save these restored defaults as the new last-used settings
    save_config(app_instance)

    messagebox.showinfo("Settings Restored", "All settings have been restored to their default values.")


def update_debug_mode_global_logic(app_instance, *args):
    set_debug_mode(app_instance.debug_mode_var.get())
    from src.config_manager import save_config # Import locally
    save_config(app_instance)

def update_scan_rbw_from_slider_index_logic(app_instance, *args):
    """
    Updates the `desired_rbw_var` (which controls the instrument's RBW)
    based on the position of the RBW slider. This links the visual slider
    to the numerical RBW setting.
    """
    try:
        idx = app_instance.rbw_slider_index_var.get()
        if 0 <= idx < len(app_instance.rbw_values):
            app_instance.desired_rbw_var.set(str(app_instance.rbw_values[idx]))
        app_instance._update_vbw_display_callback() # Update VBW display as RBW changed
        app_instance._update_setting_color_callback()
    except Exception as e:
        debug_print(f"Error updating RBW from slider index: {e}")

def update_max_hold_time_from_slider_index_logic(app_instance, *args):
    """
    Updates the `desired_max_hold_time_var` based on the Max Hold Time slider.
    """
    try:
        idx = app_instance.max_hold_time_slider_index_var.get()
        if 0 <= idx < len(app_instance.max_hold_time_values):
            app_instance.desired_max_hold_time_var.set(str(app_instance.max_hold_time_values[idx]))
        app_instance._update_setting_color_callback()
    except Exception as e:
        debug_print(f"Error updating Max Hold Time from slider index: {e}")

def update_cycle_wait_time_from_slider_index_logic(app_instance, *args):
    """
    Updates the `desired_cycle_wait_time_var` based on the Cycle Wait Time slider.
    """
    try:
        idx = app_instance.cycle_wait_time_slider_index_var.get()
        if 0 <= idx < len(app_instance.cycle_wait_time_values):
            app_instance.desired_cycle_wait_time_var.set(str(app_instance.cycle_wait_time_values[idx]))
        app_instance._update_setting_color_callback()
    except Exception as e:
        debug_print(f"Error updating Cycle Wait Time from slider index: {e}")


def reset_setting_colors_logic(app_instance):
    # This function needs to iterate through the actual labels in app_instance.setting_var_map
    # rather than assuming a 'desired_setting_entries' attribute.
    # The current implementation in main_app.py's _update_setting_color_callback
    # already handles setting label colors. This function should reset them to black.
    for var_name, (last_key, _, _) in app_instance.setting_var_map.items():
        if last_key:
            # This part is currently commented out in main_app.py's _update_setting_color_callback
            # to prevent errors. If specific labels need color management, they should be
            # explicitly defined and referenced here.
            pass


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
        scan_rbw_val = float(app_instance.desired_rbw_var.get()) # Use desired_rbw_var for instrument RBW
        app_instance.desired_vbw_display_var.set(str(int(scan_rbw_val * app_instance.VBW_RBW_RATIO))) # Use VBW_RBW_RATIO
    except ValueError:
        app_instance.desired_vbw_display_var.set("Invalid RBW")

def open_output_folder_logic(app_instance):
    """
    Opens the specified output folder in the user's default file explorer.
    This provides a convenient way for users to access their saved scan data and plots.

    Inputs:
        app_instance (App): The main application instance, providing access to `scan_directory_var`.
    Process:
        1. Retrieves the output folder path from `app_instance.scan_directory_var`.
        2. Converts the path to an absolute path if it's relative.
        3. Checks if the folder exists, showing a warning if not.
        4. Uses `os.startfile` (Windows), `subprocess.run(['open', ...])` (macOS),
           or `subprocess.run(['xdg-open', ...])` (Linux) to open the folder.
        5. Prints success or error messages to the console and displays a messagebox on error.
    Outputs: None (generates an HTML file, potentially opens a browser)
    """
    folder_path = app_instance.scan_directory_var.get() # Use scan_directory_var
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

def open_preset_folder_logic(app_instance):
    """
    Opens the default preset folder in the user's default file explorer.
    """
    # Assuming a default preset folder or one defined in config.ini
    # For now, let's use a placeholder path or a path from config if available.
    # If config.ini has a 'default_preset_directory' in DEFAULT_SETTINGS, use that.
    # Otherwise, you might need a hardcoded default or derive it.
    # For this example, let's assume it's C:\PRESETS as often seen in instrument contexts.
    preset_folder_path = "C:\\PRESETS" # Placeholder, adjust as needed

    # You might want to get this from config.ini if it's stored there
    # try:
    #     preset_folder_path = app_instance.config['DEFAULT_SETTINGS'].get('default_preset_directory', "C:\\PRESETS")
    # except KeyError:
    #     pass # Use the hardcoded default if section/key not found

    if not os.path.isabs(preset_folder_path):
        preset_folder_path = os.path.join(os.getcwd(), preset_folder_path)

    if not os.path.exists(preset_folder_path):
        messagebox.showwarning("Folder Not Found", f"The preset folder '{preset_folder_path}' does not exist.")
        print(f"🚫 Preset folder not found: {preset_folder_path}")
        return
    
    try:
        if sys.platform == "win32":
            os.startfile(preset_folder_path)
            print(f"✅ Opened preset folder: {preset_folder_path}")
        elif sys.platform == "darwin":
            subprocess.run(['open', preset_folder_path], check=True)
            print(f"✅ Opened preset folder: {preset_folder_path}")
        else: # Linux
            subprocess.run(['xdg-open', preset_folder_path], check=True)
            print(f"✅ Opened preset folder: {preset_folder_path}")
    except FileNotFoundError:
        messagebox.showerror("Error Opening Folder", f"Command not found to open folder. Please ensure 'explorer' (Windows), 'open' (macOS), or 'xdg-open' (Linux) is in your PATH.")
        print(f"❌ Command not found error when trying to open preset folder: {sys.exc_info()[1]}")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error Opening Folder", f"Failed to open preset folder '{preset_folder_path}': {e}")
        print(f"❌ Error opening preset folder: {e}")
    except Exception as e:
        messagebox.showerror("Error Opening Folder", f"An unexpected error occurred while trying to open preset folder '{preset_folder_path}': {e}")
        print(f"❌ An unexpected error occurred opening preset folder: {e}")

def open_report_folder_logic(app_instance):
    """
    Opens the default report output folder in the user's default file explorer.
    """
    # Assuming reports are saved in the same scan_directory for simplicity,
    # or you can define a separate report directory in config.ini.
    report_folder_path = app_instance.scan_directory_var.get() # Assuming reports go to scan data directory

    if not os.path.isabs(report_folder_path):
        report_folder_path = os.path.join(os.getcwd(), report_folder_path)

    if not os.path.exists(report_folder_path):
        messagebox.showwarning("Folder Not Found", f"The report folder '{report_folder_path}' does not exist.")
        print(f"🚫 Report folder not found: {report_folder_path}")
        return
    
    try:
        if sys.platform == "win32":
            os.startfile(report_folder_path)
            print(f"✅ Opened report folder: {report_folder_path}")
        elif sys.platform == "darwin":
            subprocess.run(['open', report_folder_path], check=True)
            print(f"✅ Opened report folder: {report_folder_path}")
        else: # Linux
            subprocess.run(['xdg-open', report_folder_path], check=True)
            print(f"✅ Opened report folder: {report_folder_path}")
    except FileNotFoundError:
        messagebox.showerror("Error Opening Folder", f"Command not found to open folder. Please ensure 'explorer' (Windows), 'open' (macOS), or 'xdg-open' (Linux) is in your PATH.")
        print(f"❌ Command not found error when trying to open report folder: {sys.exc_info()[1]}")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error Opening Folder", f"Failed to open report folder '{report_folder_path}': {e}")
        print(f"❌ Error opening report folder: {e}")
    except Exception as e:
        messagebox.showerror("Error Opening Folder", f"An unexpected error occurred while trying to open report folder '{report_folder_path}': {e}")
        print(f"❌ An unexpected error occurred opening report folder: {e}")
