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
    load_selected_preset as control_load_selected_preset,
    debug_print # Import debug_print
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
        
        # Get the last used device from the config (which is already loaded into resource_var)
        last_used_device = app_instance.resource_var.get()
        
        selected_device_set = False

        if app_instance.instrument_list:
            # Check if the last used device is in the currently found list
            if last_used_device and last_used_device in app_instance.instrument_list:
                app_instance.resource_var.set(last_used_device)
                debug_print(f"Set resource to last used: {last_used_device}")
                selected_device_set = True
            else:
                # If last used device is not found or was empty, default to the first available
                app_instance.resource_var.set(app_instance.instrument_list[0])
                debug_print(f"Last used device not found or empty. Defaulting to: {app_instance.instrument_list[0]}")
                selected_device_set = True
            
            app_instance.connect_button.config(state=tk.NORMAL)
            app_instance._start_connect_button_blink() # Start blinking when resources are found
        else:
            app_instance.resource_var.set("No Resources Found")
            app_instance.connect_button.config(state=tk.DISABLED)
            app_instance._stop_connect_button_blink() # Stop blinking if no resources
            selected_device_set = True # Indicate that a device state has been set (no resources)
        
        # Update the dropdown menu
        menu = app_instance.resource_dropdown["menu"]
        menu.delete(0, "end")
        for resource in app_instance.instrument_list:
            menu.add_command(label=resource, command=tk._setit(app_instance.resource_var, resource))
        
        # If no device was explicitly set (e.g., if instrument_list was empty and last_used_device was also empty)
        if not selected_device_set:
            app_instance.resource_var.set("No Resources Found")


        # Disable other buttons until connected
        app_instance.disconnect_button.config(state=tk.DISABLED)
        app_instance.apply_button.config(state=tk.DISABLED)
        app_instance.load_preset_button.config(state=tk.DISABLED)
        app_instance.start_scan_button.config(state=tk.DISABLED)
        app_instance.stop_scan_button.config(state=tk.DISABLED)
        app_instance.pause_resume_button.config(state=tk.DISABLED)

    except Exception as e:
        messagebox.showerror("Resource Error", f"Failed to list VISA resources: {e}")
        print(f"❌ Error listing VISA resources: {e}")

def connect_instrument_logic(app_instance):
    """
    Handles the connection process to the selected VISA instrument.

    Inputs:
        app_instance (App): The main application instance, providing access to
                            `rm`, `resource_var`, `inst`, `instrument_model`,
                            and various button states.
    Process:
        1. Retrieves the selected resource name from `resource_var`.
        2. Calls `connect_to_instrument` from `utils.instrument_control`.
        3. If connection is successful:
           - Stores the instrument object and model in `app_instance.inst` and `app_instance.instrument_model`.
           - Stops the connect button blinking.
           - Enables disconnect, apply, and scan buttons.
           - Queries and updates the preset tree.
           - Prints success message.
        4. If connection fails, resets GUI elements and prints error.
    Outputs: None (modifies GUI state and prints to console)
    """
    resource_name = app_instance.resource_var.get()
    if resource_name == "No Resources Found":
        messagebox.showwarning("No Instrument Selected", "Please select a VISA resource to connect.")
        return

    app_instance.connect_button.config(state=tk.DISABLED)
    app_instance._stop_connect_button_blink() # Stop blinking immediately on click

    inst, model = connect_to_instrument(app_instance.rm, resource_name)
    if inst:
        app_instance.inst = inst
        app_instance.instrument_model = model
        app_instance.disconnect_button.config(state=tk.NORMAL)
        app_instance.apply_button.config(state=tk.NORMAL)
        app_instance.start_scan_button.config(state=tk.NORMAL)
        
        # Query and update preset tree after successful connection
        update_preset_tree(app_instance)
        
        print(f"✅ Successfully connected to {resource_name} (Model: {model}).")
        # Do NOT initialize instrument settings here.
        # Settings will be applied when the "Apply Settings to Device" button is pressed.
    else:
        app_instance.inst = None
        app_instance.instrument_model = None
        messagebox.showerror("Connection Failed", f"Could not connect to {resource_name}.")
        print(f"❌ Failed to connect to {resource_name}.")
        reset_gui_on_disconnect_or_error(app_instance) # Reset GUI on failure

def disconnect_instrument_logic(app_instance):
    """
    Handles the disconnection process from the currently connected instrument.

    Inputs:
        app_instance (App): The main application instance, providing access to
                            `inst` and various button states.
    Process:
        1. Calls `control_disconnect_instrument` from `utils.instrument_control`.
        2. If disconnection is successful:
           - Resets `app_instance.inst` and `app_instance.instrument_model` to None.
           - Resets GUI elements to reflect disconnected state.
           - Starts connect button blinking.
           - Prints success message.
        3. If disconnection fails, prints error.
    Outputs: None (modifies GUI state and prints to console)
    """
    if app_instance.inst:
        if control_disconnect_instrument(app_instance.inst):
            app_instance.inst = None
            app_instance.instrument_model = None
            print("✅ Instrument disconnected successfully.")
            reset_gui_on_disconnect_or_error(app_instance)
        else:
            messagebox.showerror("Disconnect Error", "Failed to disconnect from instrument.")
            print("❌ Failed to disconnect instrument.")
    else:
        print("No instrument to disconnect.")
        reset_gui_on_disconnect_or_error(app_instance) # Ensure GUI is reset even if inst is already None

def apply_settings_to_device_logic(app_instance):
    """
    Applies the current settings from the GUI to the connected instrument.

    Inputs:
        app_instance (App): The main application instance, providing access to
                            `inst`, `desired_ref_level_var`, `high_sensitivity_var`,
                            `desired_preamp_var`, `desired_scan_rbw_segmentation_var`,
                            `desired_vbw_display_var`, `instrument_model`.
    Process:
        1. Checks if an instrument is connected.
        2. Retrieves current settings from Tkinter variables.
        3. Calls `initialize_instrument` from `utils.instrument_control` to apply settings.
        4. Queries and prints current instrument settings for verification.
        5. Prints status messages.
        6. Resets setting colors to default (indicating settings are applied).
    Outputs: None (modifies instrument state and prints to console)
    """
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument before applying settings.")
        return

    ref_level = float(app_instance.desired_ref_level_var.get())
    high_sensitivity_on = app_instance.high_sensitivity_var.get()
    preamp_on = app_instance.desired_preamp_var.get()
    rbw_config = int(app_instance.desired_scan_rbw_segmentation_var.get())
    vbw_config = int(float(app_instance.desired_vbw_display_var.get())) # Ensure VBW is int

    print("\nApplying settings to instrument...")
    if initialize_instrument(app_instance.inst, ref_level, high_sensitivity_on, preamp_on, rbw_config, vbw_config, app_instance.instrument_model):
        print("✅ Desired settings successfully applied to the instrument.")
        query_current_instrument_settings(app_instance.inst, MHZ_TO_HZ)
        app_instance.reset_setting_colors() # Reset colors after successful application
    else:
        messagebox.showerror("Apply Settings Failed", "Failed to apply settings to the instrument. Check console for details.")
        print("❌ Failed to apply settings to the instrument.")

def update_preset_tree(app_instance):
    """
    Updates the Treeview widget with preset files queried from the instrument.

    Inputs:
        app_instance (App): The main application instance.
    Process:
        1. Clears existing items in `preset_tree`.
        2. Calls `control_query_device_presets` to get available presets.
        3. Inserts each preset into the `preset_tree`.
        4. Handles cases where no presets are found.
    Outputs: None (modifies GUI state and prints to console)
    """
    for item in app_instance.preset_tree.get_children():
        app_instance.preset_tree.delete(item)

    if app_instance.inst:
        presets = control_query_device_presets(app_instance.inst)
        if presets:
            for preset in presets:
                app_instance.preset_tree.insert("", "end", values=(preset,), tags=("Mon",))
        else:
            app_instance.preset_tree.insert("", "end", values=("No Presets Found",), tags=("NoPresets",))
            app_instance.preset_tree.tag_configure("NoPresets", foreground="red")
        app_instance.load_preset_button.config(state=tk.DISABLED) # Disable until a preset is selected
    else:
        app_instance.preset_tree.insert("", "end", values=("Connect to instrument to view presets",), tags=("NoConnection",))
        app_instance.preset_tree.tag_configure("NoConnection", foreground="gray")
        app_instance.load_preset_button.config(state=tk.DISABLED)

def on_preset_select(app_instance, event):
    """
    Event handler for preset treeview selection. Enables the load preset button.

    Inputs:
        app_instance (App): The main application instance.
        event: The Tkinter event object.
    Process:
        1. Checks if an item is selected in the `preset_tree`.
        2. If a valid preset is selected and an instrument is connected, enables `load_preset_button`.
        3. Handles specific conditions for the N9340B model (which might not support presets in the same way).
    Outputs: None (modifies GUI state)
    """
    selected_item = app_instance.preset_tree.selection()
    if selected_item:
        item_text = app_instance.preset_tree.item(selected_item, "values")[0]
        if item_text != "No Presets Found" and item_text != "Connect to instrument to view presets" and app_instance.inst:
            # Disable for N9340B as it doesn't support presets in this manner
            if app_instance.instrument_model == "N9340B":
                app_instance.load_preset_button.config(state=tk.DISABLED)
                debug_print("Load Preset button disabled for N9340B.")
            else:
                app_instance.load_preset_button.config(state=tk.NORMAL)
                debug_print(f"Preset '{item_text}' selected. Load button enabled.")
        else:
            app_instance.load_preset_button.config(state=tk.DISABLED)
            debug_print("Load Preset button disabled (no valid selection or no instrument).")
    else:
        app_instance.load_preset_button.config(state=tk.DISABLED)
        debug_print("Load Preset button disabled (no selection).")

def load_selected_preset_logic(app_instance):
    """
    Loads the selected preset file onto the instrument.

    Inputs:
        app_instance (App): The main application instance.
    Process:
        1. Retrieves the selected preset name from the `preset_tree`.
        2. Calls `control_load_selected_preset` from `utils.instrument_control`.
        3. Prints status messages.
    Outputs: None (modifies instrument state and prints to console)
    """
    selected_item = app_instance.preset_tree.selection()
    if selected_item:
        selected_preset_name = app_instance.preset_tree.item(selected_item, "values")[0]
        if app_instance.inst:
            control_load_selected_preset(app_instance.inst, selected_preset_name, MHZ_TO_HZ)
        else:
            messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
    else:
        messagebox.showwarning("No Preset Selected", "Please select a preset to load from the list.")

def reset_gui_on_disconnect_or_error(app_instance):
    """
    Resets the GUI elements to their default disconnected state.

    Inputs:
        app_instance (App): The main application instance.
    Process:
        1. Disables scan, disconnect, apply, and load preset buttons.
        2. Enables refresh and connect buttons.
        3. Clears the preset tree.
        4. Starts the connect button blinking animation.
        5. Resets the instrument model.
    Outputs: None (modifies GUI state)
    """
    app_instance.start_scan_button.config(state=tk.DISABLED)
    app_instance.stop_scan_button.config(state=tk.DISABLED)
    app_instance.pause_resume_button.config(state=tk.DISABLED)
    app_instance.disconnect_button.config(state=tk.DISABLED)
    app_instance.apply_button.config(state=tk.DISABLED)
    app_instance.load_preset_button.config(state=tk.DISABLED)
    app_instance.connect_button.config(state=tk.NORMAL)
    app_instance.refresh_button.config(state=tk.NORMAL)
    
    # Clear preset tree
    for item in app_instance.preset_tree.get_children():
        app_instance.preset_tree.delete(item)
    app_instance.preset_tree.insert("", "end", values=("Connect to instrument to view presets",), tags=("NoConnection",))

    app_instance._start_connect_button_blink()
    app_instance.instrument_model = None # Clear instrument model on disconnect

def set_focus_frequency_logic(app_instance, center_frequency_hz, span_hz, device_name="N/A"):
    """
    Sets the instrument's center frequency and span.
    Also configures Trace 1 to Normal, Trace 2 to Max Hold, and Trace 3 to Min Hold.

    Inputs:
        app_instance (App): The main application instance.
        center_frequency_hz (float): The desired center frequency in Hz.
        span_hz (float): The desired span (width) in Hz around the center frequency.
        device_name (str): The name of the device/marker being focused on (for logging).
    Process:
        1. Checks if an instrument is connected. If not, prints a warning and returns False.
        2. Sends the `[:SENSe]:FREQuency:CENTer <freq>` command to the instrument.
        3. Sends the `[:SENSe]:FREQuency:SPAN <freq>` command to the instrument.
        4. Sets Trace 1 to 'NORM', Trace 2 to 'MAXHold', and Trace 3 to 'MINHold'.
        5. Prints success or error messages to the console.
        6. Handles potential exceptions during VISA communication.
    Outputs:
        bool: True if commands were sent successfully, False otherwise.
    """
    if not app_instance.inst:
        print("🚫 Instrument not connected. Cannot set focus frequency or trace modes.")
        return False
    
    debug_print(f"DEBUG (inst_logic): Received center_frequency_hz: {center_frequency_hz} (type: {type(center_frequency_hz)})")
    debug_print(f"DEBUG (inst_logic): Received span_hz: {span_hz} (type: {type(span_hz)})")
    debug_print(f"DEBUG (inst_logic): Received device_name: {device_name} (type: {type(device_name)})")

    try:
        # Set center frequency
        if not app_instance.inst.write(f":SENSe:FREQuency:CENTer {center_frequency_hz}"): return False
        debug_print(f"Sent: :SENSe:FREQuency:CENTer {center_frequency_hz} Hz")

        # Set span
        span_hz_float = float(span_hz)
        if not app_instance.inst.write(f":SENSe:FREQuency:SPAN {span_hz_float}"): return False
        debug_print(f"Sent: :SENSe:FREQuency:SPAN {span_hz_float} Hz")

        # Set Trace Modes as requested
        if not app_instance.inst.write(":TRAC1:MODE WRITe"): return False
        debug_print("Sent: :TRAC1:MODE NORM")

        if not app_instance.inst.write(":TRAC2:MODE MAXHold"): return False
        debug_print("Sent: :TRAC2:MODE MAXHold")

        if not app_instance.inst.write(":TRAC3:MODE MINHold"): return False
        debug_print("Sent: :TRAC3:MODE MINHold")

        print(f"✅ Instrument focused on '{device_name}' at {center_frequency_hz / MHZ_TO_HZ:.3f} MHz with span {span_hz_float} Hz. Trace modes set.")
        return True
    except ValueError as e:
        error_msg = f"❌ Error converting span_hz to float in set_focus_frequency_logic: {e}. Received value: '{span_hz}' (type: {type(span_hz)})"
        print(error_msg)
        messagebox.showerror("Type Conversion Error", error_msg)
        return False
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error while setting focus frequency or trace modes: {e}")
        messagebox.showerror("VISA Error", f"Failed to set instrument focus or trace modes: {e}")
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred while setting focus frequency or trace modes: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred while setting instrument focus or trace modes: {e}")
        return False

def set_marker_and_trace_modes_logic(app_instance, marker_frequency_hz, marker_name="Marker"):
    """
    Sets a marker on the instrument at the specified frequency and configures trace mode.

    Inputs:
        app_instance (App): The main application instance.
        marker_frequency_hz (float): The frequency in Hz where the marker should be placed.
        marker_name (str, optional): A descriptive name for the marker. Defaults to "Marker".
    Process:
        1. Checks if an instrument is connected.
        2. Sends SCPI commands to activate Marker 1, set its frequency, and enable marker peak search.
        3. Sets the trace type to Normal (or other desired mode).
        4. Prints status messages.
        5. Handles potential exceptions during VISA communication.
    Outputs:
        bool: True if commands were sent successfully, False otherwise.
    """
    if not app_instance.inst:
        print("🚫 Instrument not connected. Cannot set marker or trace modes.")
        messagebox.showwarning("Not Connected", "Please connect to an instrument to set markers.")
        return False

    try:
        # Activate Marker 1
        if not app_instance.inst.write(":CALCulate:MARKer1:STATe ON"): return False
        debug_print("Sent: :CALCulate:MARKer1:STATe ON")

        # Set Marker 1 frequency
        if not app_instance.inst.write(f":CALCulate:MARKer1:X {marker_frequency_hz}"): return False
        debug_print(f"Sent: :CALCulate:MARKer1:X {marker_frequency_hz} Hz")

        # Enable Marker Peak Search (optional, but common for markers)
        if not app_instance.inst.write(":CALCulate:MARKer1:MAXimum:PEAK"): return False
        debug_print("Sent: :CALCulate:MARKer1:MAXimum:PEAK")

        # Set trace type to Normal (or other desired mode like Clear Write)
        # This might be needed if the instrument was in Max Hold or Min Hold
        if not app_instance.inst.write(":DISPlay:WINDow:TRACe:TYPE NORM"): return False
        debug_print("Sent: :DISPlay:WINDow:TRACe:TYPE NORM")

        print(f"✅ Marker '{marker_name}' set at {marker_frequency_hz / MHZ_TO_HZ:.3f} MHz. Trace mode set to Normal.")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error while setting marker/trace modes: {e}")
        messagebox.showerror("VISA Error", f"Failed to set instrument marker/trace modes: {e}")
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred while setting marker/trace modes: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred while setting instrument marker/trace modes: {e}")
        return False