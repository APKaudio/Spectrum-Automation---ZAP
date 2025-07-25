# src/instrument_logic.py
import tkinter as tk
from tkinter import messagebox
import pyvisa
import os
import sys
from utils.instrument_control import (
    set_debug_mode, list_visa_resources, connect_to_instrument,
    disconnect_instrument as control_disconnect_instrument,
    initialize_instrument, query_current_instrument_settings,
    query_device_presets as control_query_device_presets,
    load_selected_preset as control_load_selected_preset
)
from utils.frequency_bands import MHZ_TO_HZ

def populate_resources_logic(app_instance):
    """
    Populates the VISA resource dropdown menu with available instruments.

    Inputs:
        app_instance (App): The main application instance, providing access to
                            `rm`, `instrument_list`, `resource_var`, `resource_dropdown`,
                            and various button states.
    Process:
        1. Uses `list_visa_resources` to find available instruments.
        2. Updates the `resource_var` with the first found resource or "No Resources Found".
        3. Clears and repopulates the `resource_dropdown` menu.
        4. Enables/disables the connect button based on resource availability.
        5. Starts/stops the connect button blinking animation.
        6. Disables scan, disconnect, apply, and load preset buttons.
        7. Handles potential exceptions during resource listing.
    Outputs: None (modifies GUI state and prints to console)
    """
    try:
        app_instance.instrument_list = list_visa_resources(app_instance.rm)
        if app_instance.instrument_list:
            last_device = app_instance.resource_var.get()
            if not last_device or last_device not in app_instance.instrument_list:
                app_instance.resource_var.set(app_instance.instrument_list[0]) if app_instance.instrument_list else "No Resources Found"
            
            menu = app_instance.resource_dropdown["menu"]
            menu.delete(0, "end")
            for resource in app_instance.instrument_list:
                menu.add_command(label=resource, command=tk._setit(app_instance.resource_var, resource))
            
            app_instance.connect_button.config(state=tk.NORMAL)
            app_instance._start_connect_button_blink()
        else:
            app_instance.resource_var.set("No resources found")
            menu = app_instance.resource_dropdown["menu"]
            menu.delete(0, "end")
            menu.add_command(label="No resources found", command=tk._setit(app_instance.resource_var, "No resources found"))
            app_instance.connect_button.config(state=tk.DISABLED)
            app_instance._stop_connect_button_blink()
        
        app_instance.start_scan_button.config(state=tk.DISABLED)
        app_instance.disconnect_button.config(state=tk.DISABLED)
        app_instance.apply_button.config(state=tk.DISABLED)
        app_instance.load_preset_button.config(state=tk.DISABLED)
        app_instance.pause_resume_button.config(state=tk.DISABLED)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to list VISA resources: {e}")
        app_instance.resource_var.set("Error listing resources")
        reset_gui_on_disconnect_or_error(app_instance)

def connect_instrument_logic(app_instance):
    """
    Establishes a connection to the selected VISA instrument and initializes its settings.

    Inputs:
        app_instance (App): The main application instance, providing access to
                            instrument connection details, settings, and GUI elements.
    Process:
        1. Retrieves the selected VISA resource.
        2. Checks for existing connections and attempts to close them.
        3. Calls `connect_to_instrument` to establish the connection.
        4. If connected, updates the window title with the instrument model.
        5. Retrieves desired instrument settings from GUI variables.
        6. Calls `initialize_instrument` to apply settings to the device.
        7. Queries and prints current instrument settings.
        8. Enables/disables relevant GUI buttons (scan, disconnect, apply).
        9. Stops the connect button blinking.
        10. Queries device preset files (if not N9340B) and updates the preset tree.
        11. Handles various exceptions (ValueError for invalid input, general Exception for others).
        12. Calls `reset_gui_on_disconnect_or_error` on failure to reset GUI state.
    Outputs: None (modifies GUI state, prints to console, shows messageboxes)
    """
    selected_resource = app_instance.resource_var.get()
    if selected_resource == "No resources found" or "Error listing resources" in selected_resource:
        messagebox.showwarning("Connection Warning", "Please select a valid VISA resource.")
        return

    if app_instance.inst:
        try:
            control_disconnect_instrument(app_instance.inst)
            app_instance.inst = None
            print("🔌 Closed existing connection.")
        except Exception as e:
            print(f"Error closing existing connection: {e}")

    try:
        app_instance.inst, app_instance.instrument_model = connect_to_instrument(app_instance.rm, selected_resource)
        if app_instance.inst:
            app_instance.title(f"RF Spectrum Analyzer Controller - {app_instance.instrument_model} - {os.path.basename(sys.argv[0])}")

            ref_level = float(app_instance.desired_ref_level_var.get())
            high_sensitivity_on = app_instance.high_sensitivity_var.get()
            preamp_on = app_instance.desired_preamp_var.get()
            rbw_config = int(float(app_instance.desired_scan_rbw_segmentation_var.get()))
            vbw_config = int(float(app_instance.desired_vbw_display_var.get()))

            if initialize_instrument(app_instance.inst, ref_level, high_sensitivity_on, preamp_on, rbw_config, vbw_config, app_instance.instrument_model):
                print("Desired settings successfully applied to the instrument.")
                query_current_instrument_settings(app_instance.inst, MHZ_TO_HZ)
                
                app_instance.start_scan_button.config(state=tk.NORMAL)
                app_instance.stop_scan_button.config(state=tk.DISABLED)
                app_instance.pause_resume_button.config(state=tk.DISABLED)
                app_instance.disconnect_button.config(state=tk.NORMAL)
                app_instance.apply_button.config(state=tk.NORMAL)
                app_instance._stop_connect_button_blink()

                if app_instance.instrument_model != "N9340B":
                    preset_files = control_query_device_presets(app_instance.inst)
                    update_preset_tree(app_instance, preset_files)
                else:
                    print("ℹ️ Skipping device preset query for N9340B model.")
                    update_preset_tree(app_instance, [])
                    app_instance.preset_tree.insert("", "end", values=("Presets not supported for N9340B.",), tags=("disabled",))
                    app_instance.load_preset_button.config(state=tk.DISABLED)

            else:
                messagebox.showerror("Initialization Failed", "Instrument initialization with desired settings failed.")
                control_disconnect_instrument(app_instance.inst)
                app_instance.inst = None
                reset_gui_on_disconnect_or_error(app_instance)
        else:
            messagebox.showerror("Connection Failed", "Could not connect to instrument. Check console for details.")
            reset_gui_on_disconnect_or_error(app_instance)

    except ValueError as e:
        messagebox.showerror("Input Error", f"Invalid numeric value for instrument settings: {e}. Please check Reference Level, RBW, and Max Hold Time.")
        reset_gui_on_disconnect_or_error(app_instance)
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        print(f"❌ An unexpected error occurred during connection: {e}")
        reset_gui_on_disconnect_or_error(app_instance)

def disconnect_instrument_logic(app_instance):
    """
    Disconnects from the currently connected VISA instrument.

    Inputs:
        app_instance (App): The main application instance, providing access to
                            the instrument object (`inst`) and GUI elements.
    Process:
        1. Calls `control_disconnect_instrument` to close the VISA connection.
        2. Resets `app_instance.inst` and `app_instance.instrument_model` to None.
        3. Prints a disconnection message.
        4. Calls `reset_gui_on_disconnect_or_error` to revert GUI state.
        5. Updates the window title.
        6. If resources are available, starts the connect button blinking.
        7. Shows an error messagebox if disconnection fails.
    Outputs: None (modifies GUI state, prints to console, shows messageboxes)
    """
    if control_disconnect_instrument(app_instance.inst):
        app_instance.inst = None
        app_instance.instrument_model = None
        print("Disconnected.")
        reset_gui_on_disconnect_or_error(app_instance)
        app_instance.title(f"RF Spectrum Analyzer Controller - {os.path.basename(sys.argv[0])}")
        if app_instance.instrument_list and app_instance.resource_var.get() != "No resources found":
            app_instance._start_connect_button_blink()
    else:
        messagebox.showerror("Disconnect Error", "Failed to disconnect instrument.")

def apply_settings_to_device_logic(app_instance):
    """
    Applies the current settings from the GUI to the connected instrument.

    Inputs:
        app_instance (App): The main application instance, providing access to
                            the instrument object (`inst`) and desired settings.
    Process:
        1. Checks if an instrument is connected. If not, shows a warning.
        2. Retrieves desired settings (reference level, preamp, RBW, VBW) from GUI variables.
        3. Calls `initialize_instrument` to send these settings to the device.
        4. Prints success or error messages to the console.
        5. Calls `reset_setting_colors_logic` to update GUI visual feedback.
        6. Handles `ValueError` for invalid numeric inputs and general `Exception`.
    Outputs: None (modifies instrument state, prints to console, shows messageboxes)
    """
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first to apply settings.")
        return

    try:
        ref_level = float(app_instance.desired_ref_level_var.get())
        high_sensitivity_on = app_instance.high_sensitivity_var.get()
        preamp_on = app_instance.desired_preamp_var.get()
        rbw_config = int(float(app_instance.desired_scan_rbw_segmentation_var.get()))
        vbw_config = int(float(app_instance.desired_vbw_display_var.get()))

        if initialize_instrument(app_instance.inst, ref_level, high_sensitivity_on, preamp_on, rbw_config, vbw_config, app_instance.instrument_model):
            print("Desired settings successfully applied to the instrument.")
            app_instance.reset_setting_colors()
        else:
            messagebox.showerror("Apply Failed", "Failed to apply settings to the instrument. Check console for details.")
    except ValueError as e:
        messagebox.showerror("Input Error", f"Invalid numeric value for instrument settings: {e}. Please check Reference Level, RBW, and Max Hold Time.")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred while applying settings: {e}")
        print(f"❌ Error applying settings: {e}")

def update_preset_tree(app_instance, preset_files):
    """
    Updates the Treeview widget that displays available device preset files.

    Inputs:
        app_instance (App): The main application instance, providing access to `preset_tree`.
        preset_files (list): A list of strings, where each string is the name of a preset file.
    Process:
        1. Clears all existing items from the `preset_tree`.
        2. If `preset_files` is not empty, inserts each preset name as a new item.
           - Adds a "Mon" tag for blue foreground if "MON" is in the preset name.
        3. If `preset_files` is empty, inserts a "No .STA preset files found." message.
        4. Disables the "Load Selected Preset" button.
    Outputs: None (modifies GUI treeview)
    """
    for item in app_instance.preset_tree.get_children():
        app_instance.preset_tree.delete(item)
    if preset_files:
        for preset_name in sorted(preset_files):
            tags = ()
            if "MON" in preset_name.upper():
                tags = ("Mon",)
            app_instance.preset_tree.insert("", "end", values=(preset_name,), tags=tags)
    else:
        app_instance.preset_tree.insert("", "end", values=("No .STA preset files found.",))
    app_instance.load_preset_button.config(state=tk.DISABLED)

def on_preset_select(app_instance, event):
    """
    Event handler for when a preset file is selected in the preset treeview.
    Enables the "Load Selected Preset" button if a valid preset is selected and instrument is connected.

    Inputs:
        app_instance (App): The main application instance, providing access to `preset_tree`,
                            `inst`, `instrument_model`, and `load_preset_button`.
        event (tk.Event): The Tkinter event object (not directly used, but part of binding).
    Process:
        1. Retrieves the currently selected item(s) from the `preset_tree`.
        2. If a valid item is selected and an instrument is connected (and not N9340B model),
           enables the `load_preset_button`.
        3. Otherwise, disables the `load_preset_button`.
    Outputs: None (modifies GUI button state)
    """
    selected_items = app_instance.preset_tree.selection()
    if selected_items and app_instance.inst and app_instance.instrument_model != "N9340B":
        app_instance.load_preset_button.config(state=tk.NORMAL)
    else:
        app_instance.load_preset_button.config(state=tk.DISABLED)

def load_selected_preset_logic(app_instance):
    """
    Loads the currently selected preset file onto the connected instrument.

    Inputs:
        app_instance (App): The main application instance, providing access to
                            `inst`, `instrument_model`, `preset_tree`, and messageboxes.
    Process:
        1. Checks if an instrument is connected. If not, shows a warning.
        2. Checks if the instrument model is N9340B, which does not support presets.
        3. Retrieves the name of the selected preset from the `preset_tree`.
        4. Calls `control_load_selected_preset` to send the load command to the instrument.
        5. Prints success or error messages to the console and shows messageboxes.
    Outputs: None (modifies instrument state, prints to console, shows messageboxes)
    """
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
        return
    
    if app_instance.instrument_model == "N9340B":
        messagebox.showwarning("Preset Not Supported", "Loading presets is not supported for the N9340B model.")
        print("🚫 Attempted to load preset on N9340B, which is not supported.")
        return

    selected_items = app_instance.preset_tree.selection()
    if not selected_items:
        messagebox.showwarning("No Preset Selected", "Please select a preset file from the list to load.")
        return

    selected_preset_name = app_instance.preset_tree.item(selected_items[0], 'values')[0]
    
    if control_load_selected_preset(app_instance.inst, selected_preset_name, MHZ_TO_HZ):
        print(f"Preset '{selected_preset_name}' loaded successfully via instrument_control.")
    else:
        messagebox.showerror("Load Preset Failed", f"Failed to load preset: {selected_preset_name}. Check console for details.")

def reset_gui_on_disconnect_or_error(app_instance):
    """
    Resets various GUI elements to their disconnected/initial state after
    an instrument disconnect or an error.

    Inputs:
        app_instance (App): The main application instance, providing access to
                            various buttons, the preset tree, and connection state.
    Process:
        1. Disables scan, stop, pause/resume, disconnect, apply, and load preset buttons.
        2. Clears and resets the preset tree to show "No instrument connected."
        3. Enables the connect button.
        4. Starts the connect button blinking if resources are available.
    Outputs: None (modifies GUI state)
    """
    app_instance.start_scan_button.config(state=tk.DISABLED)
    app_instance.stop_scan_button.config(state=tk.DISABLED)
    app_instance.pause_resume_button.config(state=tk.DISABLED)
    app_instance.disconnect_button.config(state=tk.DISABLED)
    app_instance.apply_button.config(state=tk.DISABLED)
    app_instance.load_preset_button.config(state=tk.DISABLED)
    for item in app_instance.preset_tree.get_children():
        app_instance.preset_tree.delete(item)
    app_instance.preset_tree.insert("", "end", values=("No instrument connected.",))
    app_instance.connect_button.config(state=tk.NORMAL)
    if app_instance.instrument_list and app_instance.resource_var.get() != "No resources found":
        app_instance._start_connect_button_blink()
    else:
        app_instance._stop_connect_button_blink()

def set_focus_frequency_logic(app_instance, center_frequency_hz, device_name, span_hz):
    """
    Sets the instrument's center frequency and span to focus on a specific device.

    Inputs:
        app_instance (App): The main application instance, providing access to the instrument object (`inst`).
        center_frequency_hz (float): The desired center frequency in Hz.
        device_name (str): The name of the device being focused on (for logging/display).
        span_hz (float): The desired span (width) in Hz around the center frequency.
    Process:
        1. Checks if an instrument is connected. If not, prints a warning and returns False.
        2. Sends the `[:SENSe]:FREQuency:CENTer <freq>` command to the instrument.
        3. Sends the `[:SENSe]:FREQuency:SPAN <freq>` command to the instrument.
        4. Prints success or error messages to the console.
        5. Handles potential exceptions during VISA communication.
    Outputs:
        bool: True if commands were sent successfully, False otherwise.
    """
    if not app_instance.inst:
        print("🚫 Instrument not connected. Cannot set focus frequency.")
        return False
    
    try:
        # Set center frequency
        app_instance.inst.write(f":SENSe:FREQuency:CENTer {center_frequency_hz}")
        print(f"Sent: :SENSe:FREQuency:CENTer {center_frequency_hz} Hz")

        # Set span
        app_instance.inst.write(f":SENSe:FREQuency:SPAN {span_hz}")
        print(f"Sent: :SENSe:FREQuency:SPAN {span_hz} Hz")

        print(f"✅ Instrument focused on '{device_name}' at {center_frequency_hz / MHZ_TO_HZ:.3f} MHz with span {span_hz} Hz.")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error while setting focus frequency for '{device_name}': {e}")
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred while setting focus frequency for '{device_name}': {e}")
        return False
