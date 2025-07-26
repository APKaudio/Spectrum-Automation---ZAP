# src/instrument_logic.py
import tkinter as tk
from tkinter import messagebox
import pyvisa
import os
import sys
import inspect # Import inspect module
from utils.instrument_control import (
    set_debug_mode, list_visa_resources, connect_to_instrument,
    disconnect_instrument as control_disconnect_instrument,
    initialize_instrument, query_current_instrument_settings,
    query_device_presets as control_query_device_presets,
    load_selected_preset as control_load_selected_preset,
    debug_print # Import debug_print
)
from utils.frequency_bands import MHZ_TO_HZ
from src.config_manager import save_config # Import save_config

def _get_float_value(tk_var, default_value, setting_name):
    """
    Safely retrieves a float value from a Tkinter StringVar.
    If the string is empty or cannot be converted, returns a default value.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    try:
        val_str = tk_var.get()
        if not val_str:
            debug_print(f"Warning: Tkinter variable for '{setting_name}' is empty. Using default: {default_value}", file=current_file, function=current_function)
            return default_value
        return float(val_str)
    except ValueError:
        debug_print(f"Error: Could not convert '{val_str}' for '{setting_name}' to float. Using default: {default_value}", file=current_file, function=current_function)
        return default_value

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
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Attempting to populate VISA resources...", file=current_file, function=current_function)
    try:
        app_instance.instrument_list = list_visa_resources(app_instance.rm)
        
        # Get the last used device from the config (which is already loaded into resource_var)
        last_used_device = app_instance.gpib_device_var.get() # Use gpib_device_var for last used
        
        selected_device_set = False

        if app_instance.instrument_list:
            # Check if the last used device is in the currently found list
            if last_used_device and last_used_device in app_instance.instrument_list:
                app_instance.resource_var.set(last_used_device)
                debug_print(f"Set resource to last used: {last_used_device}", file=current_file, function=current_function)
                selected_device_set = True
            else:
                # If last used device is not found or was empty, default to the first available
                app_instance.resource_var.set(app_instance.instrument_list[0])
                debug_print(f"Last used device not found or empty. Defaulting to: {app_instance.instrument_list[0]}", file=current_file, function=current_function)
                selected_device_set = True
            
            app_instance.connect_button.config(state=tk.NORMAL)
            app_instance._start_connect_button_blink() # Start blinking when resources are found
        else:
            app_instance.resource_var.set("No Resources Found")
            app_instance.connect_button.config(state=tk.DISABLED)
            app_instance._start_connect_button_blink() # Start blinking if no resources
            selected_device_set = True # Indicate that a device state has been set (no resources)
        
        # Update the dropdown menu
        menu = app_instance.resource_dropdown["menu"]
        menu.delete(0, "end")
        for resource in app_instance.instrument_list:
            menu.add_command(label=resource, command=tk._setit(app_instance.resource_var, resource))
        
        # Disable other buttons until connected
        app_instance.disconnect_button.config(state=tk.DISABLED)
        app_instance.apply_button.config(state=tk.DISABLED)
        app_instance.start_scan_button.config(state=tk.DISABLED)
        app_instance.stop_scan_button.config(state=tk.DISABLED)
        app_instance.pause_resume_button.config(state=tk.DISABLED)
        app_instance.query_presets_button.config(state=tk.DISABLED) # Ensure this is disabled until connected
        app_instance.plot_button.config(state=tk.DISABLED) # Disable plot button on refresh

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
           - Queries and updates the preset buttons.
           - Prints success message.
        4. If connection fails, resets GUI elements and prints error.
    Outputs: None (modifies GUI state and prints to console)
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    selected_resource = app_instance.resource_var.get()
    if selected_resource == "No Resources Found" or not selected_resource:
        messagebox.showwarning("No Resource Selected", "Please select a VISA resource to connect to.")
        return

    debug_print(f"Attempting to connect to: {selected_resource}", file=current_file, function=current_function)
    app_instance.console_output.insert(tk.END, f"Connecting to {selected_resource}...\n", "cyan")
    app_instance.console_output.see(tk.END)

    try:
        # Unpack the tuple returned by connect_to_instrument
        app_instance.inst, debug_mode_status = connect_to_instrument(app_instance.rm, selected_resource)
        
        if app_instance.inst:
            app_instance.gpib_device_var.set(selected_resource) # Update GPIB device display

            # Retrieve current desired settings from GUI variables for initialization
            # Use the helper function to safely get float values, providing reasonable defaults
            ref_level = _get_float_value(app_instance.desired_reference_level_var, -40.0, "Reference Level")
            high_sensitivity = app_instance.desired_high_sensitivity_var.get()
            preamp = app_instance.desired_preamp_on_var.get()
            rbw = _get_float_value(app_instance.desired_rbw_var, 10000.0, "RBW")
            vbw = _get_float_value(app_instance.desired_vbw_display_var, rbw / app_instance.VBW_RBW_RATIO, "VBW") # Default VBW based on RBW

            # Call initialize_instrument with all required arguments
            app_instance.instrument_model = initialize_instrument(
                app_instance.inst,
                ref_level_dbm=ref_level,
                high_sensitivity_on=high_sensitivity,
                preamp_on=preamp,
                rbw_config_val=rbw,
                vbw_config_val=vbw,
                model_match=None # Pass None for model_match as the model is returned by this function
            )
            query_current_instrument_settings(app_instance.inst, MHZ_TO_HZ) # Query and display settings

            app_instance.connect_button.config(state=tk.DISABLED)
            app_instance.disconnect_button.config(state=tk.NORMAL)
            app_instance.start_scan_button.config(state=tk.NORMAL)
            app_instance.apply_button.config(state=tk.NORMAL)
            app_instance.query_presets_button.config(state=tk.NORMAL) # Enable query presets button
            app_instance.plot_button.config(state=tk.NORMAL) # Enable plot button on connect
            app_instance._stop_connect_button_blink() # Stop blinking on successful connect
            messagebox.showinfo("Connection Successful", f"Successfully connected to {selected_resource} ({app_instance.instrument_model}).")
            debug_print(f"✅ Successfully connected to {selected_resource} ({app_instance.instrument_model}).", file=current_file, function=current_function)
        else:
            messagebox.showerror("Connection Failed", f"Could not connect to {selected_resource}.")
            debug_print(f"❌ Failed to connect to {selected_resource}.", file=current_file, function=current_function)
            app_instance._reset_gui_on_disconnect_or_error() # Reset GUI on failure
    except pyvisa.errors.VisaIOError as e:
        messagebox.showerror("VISA Error", f"VISA I/O error during connection to {selected_resource}: {e}")
        debug_print(f"❌ VISA I/O error: {e}", file=current_file, function=current_function)
        app_instance._reset_gui_on_disconnect_or_error()
    except Exception as e:
        messagebox.showerror("Connection Error", f"An unexpected error occurred during connection: {e}")
        debug_print(f"❌ Unexpected connection error: {e}", file=current_file, function=current_function)
        app_instance._reset_gui_on_disconnect_or_error()


def disconnect_instrument_logic(app_instance):
    """
    Disconnects from the currently connected VISA instrument.

    Inputs:
        app_instance (App): The main application instance, providing access to `inst`.
    Process:
        1. Calls `control_disconnect_instrument` from `instrument_control`.
        2. Resets GUI elements to a disconnected state.
        3. Displays a success or error message.
    Outputs: None (modifies GUI state and prints to console)
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "No instrument is currently connected.")
        return

    debug_print("Attempting to disconnect instrument...", file=current_file, function=current_function)
    try:
        control_disconnect_instrument(app_instance.inst)
        messagebox.showinfo("Disconnected", "Instrument disconnected successfully.")
        debug_print("✅ Instrument disconnected successfully.", file=current_file, function=current_function)
    except Exception as e:
        messagebox.showerror("Disconnect Error", f"Error disconnecting instrument: {e}")
        debug_print(f"❌ Error disconnecting instrument: {e}", file=current_file, function=current_function)
    finally:
        app_instance._reset_gui_on_disconnect_or_error() # Always reset GUI state


def apply_settings_to_device_logic(app_instance):
    """
    Applies the current settings from the GUI to the connected instrument.
    Also saves the current configuration to config.ini.

    Inputs:
        app_instance (App): The main application instance.
    Process:
        1. Checks for active instrument connection.
        2. Retrieves values from Tkinter variables.
        3. Constructs and sends SCPI commands to the instrument.
        4. Handles potential VISA errors during command transmission.
        5. Calls `save_config` to persist the current settings.
    Outputs: None
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first to apply settings.")
        return

    debug_print("Applying settings to instrument...", file=current_file, function=current_function)
    try:
        # Get values from Tkinter variables
        rbw_hz = float(app_instance.desired_rbw_var.get())
        max_hold_time_s = float(app_instance.desired_max_hold_time_var.get())
        ref_level_dbm = float(app_instance.desired_reference_level_var.get())
        freq_shift_hz = float(app_instance.desired_freq_shift_var.get())
        maxhold_enabled = app_instance.desired_maxhold_enabled_var.get()
        high_sensitivity = app_instance.desired_high_sensitivity_var.get()
        preamp_on = app_instance.desired_preamp_on_var.get()

        # Apply RBW
        if not app_instance.inst.write(f":SENSe:BANDwidth:RESolution {rbw_hz}"): return
        debug_print(f"Sent: :SENSe:BANDwidth:RESolution {rbw_hz}", file=current_file, function=current_function)

        # Apply Max Hold Time (if enabled)
        if maxhold_enabled:
            if not app_instance.inst.write(":DISPlay:WINDow:TRACe:MODE MAXH"): return # Set to Max Hold mode
            debug_print("Sent: :DISPlay:WINDow:TRACe:MODE MAXH", file=current_file, function=current_function)
            # Note: Max Hold Time might be a display setting or a sweep time setting,
            # depending on the instrument. Assuming it influences sweep time or display persistence.
            # This command might need adjustment based on the exact instrument model.
            # For now, we'll just ensure max hold mode is ON.
            # Some instruments might not have a direct "max hold time" command, but rather a "sweep time"
            # or "trace average" setting that implicitly affects max hold duration.
            # For N9340B, there isn't a direct "MAXH time" setting. It's usually a continuous max hold.
            # We will just ensure the mode is set.
        else:
            if not app_instance.inst.write(":DISPlay:WINDow:TRACe:MODE WRIT"): return # Set to Clear Write mode
            debug_print("Sent: :DISPlay:WINDow:TRACe:MODE WRIT", file=current_file, function=current_function)

        # Apply Reference Level
        if not app_instance.inst.write(f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM"): return
        debug_print(f"Sent: :DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM", file=current_file, function=current_function)

        # Apply Frequency Shift (if instrument supports it, this is a general command)
        # Note: Frequency shift is often a marker or measurement specific setting,
        # not a global instrument setting in all analyzers. This command might need
        # to be adjusted or removed if the instrument doesn't support it directly.
        # For N9340B, there isn't a direct global frequency shift.
        # This might be for a marker or a specific measurement function.
        # For now, we'll keep it as a placeholder.
        # if not app_instance.inst.write(f":FREQuency:OFFSet {freq_shift_hz}"): return
        # debug_print(f"Sent: :FREQuency:OFFSet {freq_shift_hz}")

        # Apply High Sensitivity (often related to preamp or detector mode)
        # N9340B has :INPut:ATTenuation:AUTO ON|OFF or :SENSe:POWer:RF:RANGe:AUTO ON|OFF
        # High sensitivity could mean turning off auto-attenuation or enabling preamp.
        if high_sensitivity:
            if not app_instance.inst.write(":SENSe:POWer:RF:RANGe:AUTO OFF"): return # Disable auto-range for max sensitivity
            debug_print("Sent: :SENSe:POWer:RF:RANGe:AUTO OFF", file=current_file, function=current_function)
        else:
            if not app_instance.inst.write(":SENSe:POWer:RF:RANGe:AUTO ON"): return # Enable auto-range
            debug_print("Sent: :SENSe:POWer:RF:RANGe:AUTO ON", file=current_file, function=current_function)

        # Apply Preamp On/Off
        if preamp_on:
            if not app_instance.inst.write(":INPut:GAIN:STATe ON"): return # Enable preamplifier
            debug_print("Sent: :INPut:GAIN:STATe ON", file=current_file, function=current_function)
        else:
            if not app_instance.inst.write(":INPut:GAIN:STATe OFF"): return # Disable preamplifier
            debug_print("Sent: :INPut:GAIN:STATe OFF", file=current_file, function=current_function)

        messagebox.showinfo("Settings Applied", "Settings applied to instrument successfully and configuration saved.")
        debug_print("✅ Settings applied to instrument successfully.", file=current_file, function=current_function)
        
        # Save the current configuration after applying settings
        save_config(app_instance)

    except pyvisa.errors.VisaIOError as e:
        messagebox.showerror("VISA Error", f"VISA I/O error applying settings: {e}")
        debug_print(f"❌ VISA I/O error applying settings: {e}", file=current_file, function=current_function)
    except Exception as e:
        messagebox.showerror("Apply Settings Error", f"An unexpected error occurred while applying settings: {e}")
        debug_print(f"❌ Unexpected error applying settings: {e}", file=current_file, function=current_function)


def update_preset_buttons(app_instance, buttons_frame):
    """
    Queries the device for available presets and displays them as buttons in the GUI.
    Clears existing buttons before populating new ones.

    Inputs:
        app_instance (App): The main application instance.
        buttons_frame (ttk.Frame): The frame where preset buttons should be placed.
    Process:
        1. Clears all existing widgets (buttons) from `buttons_frame`.
        2. Queries device presets using `control_query_device_presets`.
        3. For each preset, creates a `ttk.Button` with the preset name.
        4. Binds the button's command to `load_selected_preset_logic` with the preset name.
        5. Handles cases where no presets are found or an error occurs.
    Outputs: None (modifies GUI state and prints to console)
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first to query presets.")
        return
    
    # Clear existing buttons
    for widget in buttons_frame.winfo_children():
        widget.destroy()

    debug_print("Querying device presets...", file=current_file, function=current_function)
    try:
        presets = control_query_device_presets(app_instance.inst)
        if presets:
            for i, preset_name in enumerate(sorted(presets)): # Sort for consistent display
                # Create a button for each preset
                preset_button = ttk.Button(buttons_frame, text=preset_name,
                                           command=lambda name=preset_name: load_selected_preset_logic(app_instance, name),
                                           style='GreyText.TButton')
                preset_button.grid(row=i // 3, column=i % 3, padx=5, pady=5, sticky="ew") # Arrange in a grid
            buttons_frame.grid_columnconfigure(0, weight=1) # Ensure columns expand
            buttons_frame.grid_columnconfigure(1, weight=1)
            buttons_frame.grid_columnconfigure(2, weight=1)
            debug_print(f"✅ Displayed {len(presets)} presets as buttons.", file=current_file, function=current_function)
        else:
            ttk.Label(buttons_frame, text="No presets found on device.").grid(row=0, column=0, columnspan=3, padx=5, pady=5)
            debug_print("🚫 No presets found on device.", file=current_file, function=current_function)
    except Exception as e:
        messagebox.showerror("Preset Query Error", f"Failed to query device presets: {e}")
        debug_print(f"❌ Error querying device presets: {e}", file=current_file, function=current_function)


def load_selected_preset_logic(app_instance, selected_preset_name):
    """
    Loads the specified preset file onto the instrument.

    Inputs:
        app_instance (App): The main application instance.
        selected_preset_name (str): The name of the preset to load.
    Process:
        1. Checks for active instrument connection.
        2. Calls `control_load_selected_preset` from `instrument_control`.
        3. Displays a success or error message.
    Outputs: None
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first to load a preset.")
        return
    
    if app_instance.instrument_model == "N9340B":
        messagebox.showwarning("Feature Not Supported", "The connected instrument (N9340B) does not support loading presets via SCPI commands.")
        debug_print("🚫 N9340B does not support loading presets.", file=current_file, function=current_function)
        return

    debug_print(f"Attempting to load preset: {selected_preset_name}", file=current_file, function=current_function)
    try:
        if control_load_selected_preset(app_instance.inst, selected_preset_name):
            messagebox.showinfo("Preset Loaded", f"Preset '{selected_preset_name}' loaded successfully.")
            debug_print(f"✅ Preset '{selected_preset_name}' loaded successfully.", file=current_file, function=current_function)
            # Optionally, query current settings after loading preset to update GUI
            query_current_instrument_settings(app_instance.inst, MHZ_TO_HZ)
        else:
            messagebox.showerror("Preset Load Failed", f"Failed to load preset '{selected_preset_name}'. See console for details.")
            debug_print(f"❌ Failed to load preset '{selected_preset_name}'.", file=current_file, function=current_function)
    except Exception as e:
        messagebox.showerror("Preset Load Error", f"An unexpected error occurred while loading preset: {e}")
        debug_print(f"❌ Unexpected error loading preset: {e}", file=current_file, function=current_function)


def set_focus_frequency_logic(app_instance, frequency_hz):
    """
    Sets the instrument's center frequency (or marker frequency) to the specified value.

    Inputs:
        app_instance (App): The main application instance.
        frequency_hz (float): The frequency in Hz to set.
    Process:
        1. Checks for active instrument connection.
        2. Sends SCPI command to set center frequency.
        3. Displays success or error message.
    Outputs: None
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first to set frequency.")
        return

    debug_print(f"Attempting to set instrument frequency to {frequency_hz} Hz...", file=current_file, function=current_function)
    try:
        # For N9340B, setting center frequency is :SENSe:FREQuency:CENTer
        if not app_instance.inst.write(f":SENSe:FREQuency:CENTer {frequency_hz}"): return
        debug_print(f"Sent: :SENSe:FREQuency:CENTer {frequency_hz}", file=current_file, function=current_function)
        messagebox.showinfo("Frequency Set", f"Instrument center frequency set to {frequency_hz / MHZ_TO_HZ:.3f} MHz.")
        debug_print(f"✅ Instrument frequency set to {frequency_hz / MHZ_TO_HZ:.3f} MHz.", file=current_file, function=current_function)
    except pyvisa.errors.VisaIOError as e:
        messagebox.showerror("VISA Error", f"VISA I/O error setting frequency: {e}")
        debug_print(f"❌ VISA I/O error setting frequency: {e}", file=current_file, function=current_function)
    except Exception as e:
        messagebox.showerror("Frequency Set Error", f"An unexpected error occurred while setting frequency: {e}")
        debug_print(f"❌ Unexpected error setting frequency: {e}", file=current_file, function=current_function)


def set_marker_and_trace_modes_logic(app_instance, marker_frequency_hz, marker_name="Marker 1"):
    """
    Sets a marker at a specific frequency and ensures trace mode is normal.

    Inputs:
        app_instance (App): The main application instance.
        marker_frequency_hz (float): The frequency in Hz for the marker.
        marker_name (str): The name of the marker (e.g., "Marker 1").
    Process:
        1. Checks for active instrument connection.
        2. Sends SCPI commands to enable Marker 1, set its frequency, and enable peak search.
        3. Sets the trace type to Normal (Clear Write).
        4. Displays success or error message.
    Outputs: None
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first to set markers.")
        return

    debug_print(f"Attempting to set marker '{marker_name}' at {marker_frequency_hz} Hz...", file=current_file, function=current_function)
    try:
        # Enable Marker 1
        if not app_instance.inst.write(":CALCulate:MARKer1:STATe ON"): return False
        debug_print("Sent: :CALCulate:MARKer1:STATe ON", file=current_file, function=current_function)

        # Set Marker 1 frequency
        if not app_instance.inst.write(f":CALCulate:MARKer1:X {marker_frequency_hz}"): return False
        debug_print(f"Sent: :CALCulate:MARKer1:X {marker_frequency_hz} Hz", file=current_file, function=current_function)

        # Enable Marker Peak Search (optional, but common for markers)
        if not app_instance.inst.write(":CALCulate:MARKer1:MAXimum:PEAK"): return False
        debug_print("Sent: :CALCulate:MARKer1:MAXimum:PEAK", file=current_file, function=current_function)

        # Set trace type to Normal (or other desired mode like Clear Write)
        # This might be needed if the instrument was in Max Hold or Min Hold
        if not app_instance.inst.write(":DISPlay:WINDow:TRACe:TYPE NORM"): return False
        debug_print("Sent: :DISPlay:WINDow:TRACe:TYPE NORM", file=current_file, function=current_function)

        print(f"✅ Marker '{marker_name}' set at {marker_frequency_hz / MHZ_TO_HZ:.3f} MHz. Trace mode set to Normal.")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error while setting marker/trace modes: {e}")
        messagebox.showerror("VISA Error", f"Failed to set instrument marker/trace modes: {e}")
        return False
    except Exception as e:
        print(f"🚨 An unexpected error occurred while setting marker/trace modes: {e}")
        messagebox.showerror("Marker/Trace Error", f"An unexpected error occurred while setting instrument marker/trace modes: {e}")
        return False
