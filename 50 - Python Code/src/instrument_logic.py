# src/instrument_logic.py
import tkinter as tk
# from tkinter import messagebox # Removed messagebox import
import pyvisa
import os
import sys
import inspect # Import inspect module
from utils.instrument_control import (
    set_debug_mode, list_visa_resources, connect_to_instrument,
    disconnect_instrument as control_disconnect_instrument,
    initialize_instrument, # This is the correct initialize_instrument from utils
    query_current_instrument_settings,
    query_device_presets as control_query_device_presets, # Alias the imported function
    load_selected_preset as control_load_selected_preset,
    debug_print # Import debug_print
)
from utils.frequency_bands import MHZ_TO_HZ, VBW_RBW_RATIO # Import VBW_RBW_RATIO
from src.config_manager import save_config # Import save_config
import tkinter.ttk as ttk # Import ttk for themed widgets

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
            return float(default_value)
        return float(val_str)
    except ValueError:
        print(f"❌ Invalid value for {setting_name}: '{tk_var.get()}'. Using default: {default_value}")
        debug_print(f"ValueError for {setting_name}: '{tk_var.get()}'. Using default: {default_value}", file=current_file, function=current_function)
        return float(default_value)

def _get_int_value(tk_var, default_value, setting_name):
    """
    Safely retrieves an int value from a Tkinter IntVar or StringVar.
    If the string is empty or cannot be converted, returns a default value.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    try:
        val_str = tk_var.get()
        if not val_str:
            debug_print(f"Warning: Tkinter variable for '{setting_name}' is empty. Using default: {default_value}", file=current_file, function=current_function)
            return int(default_value)
        return int(default_value) # Return default if empty
    except ValueError:
        print(f"❌ Invalid value for {setting_name}: '{tk_var.get()}'. Using default: {default_value}")
        debug_print(f"ValueError for {setting_name}: '{tk_var.get()}'. Using default: {default_value}", file=current_file, function=current_function)
        return int(default_value)

def _get_bool_value(tk_var, default_value, setting_name):
    """
    Safely retrieves a boolean value from a Tkinter BooleanVar.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    try:
        return tk_var.get()
    except Exception:
        print(f"❌ Invalid value for {setting_name}: '{tk_var.get()}'. Using default: {default_value}")
        debug_print(f"Error getting boolean value for {setting_name}: '{tk_var.get()}'. Using default: {default_value}", file=current_file, function=current_function)
        return default_value


def populate_resources_logic(app_instance):
    """
    Populates the VISA resource dropdown menu with available instruments.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    app_instance.instrument_list = []
    try:
        resources = list_visa_resources(app_instance.rm)
        if not resources:
            print("🚫 No VISA resources found. Ensure instrument is connected and drivers are installed.")
            debug_print("No VISA resources found.", file=current_file, function=current_function)
            app_instance.connect_button.config(state=tk.DISABLED)
            app_instance._start_connect_button_blink()
            return
        
        app_instance.instrument_list = resources
        app_instance.resource_dropdown['menu'].delete(0, 'end') # Clear existing options
        for resource in resources:
            app_instance.resource_dropdown['menu'].add_command(label=resource, command=tk._setit(app_instance.resource_var, resource))
        
        # Try to re-select the last used device if it's still available
        last_gpib = app_instance.gpib_device_var.get()
        if last_gpib and last_gpib in resources:
            app_instance.resource_var.set(last_gpib)
            print(f"✅ Last used instrument '{last_gpib}' found.")
        elif resources:
            app_instance.resource_var.set(resources[0]) # Select the first resource by default
            print(f"✅ Found VISA resources. Selected: {resources[0]}")
        
        app_instance.connect_button.config(state=tk.NORMAL)
        app_instance._stop_connect_button_blink() # Stop blinking if resources are found
        debug_print("VISA resources populated.", file=current_file, function=current_function)

    except Exception as e:
        print(f"❌ Error listing VISA resources: {e}")
        debug_print(f"Error listing VISA resources: {e}", file=current_file, function=current_function)
        app_instance.connect_button.config(state=tk.DISABLED)
        app_instance._start_connect_button_blink()


def connect_instrument_logic(app_instance):
    """
    Connects to the selected instrument and initializes it.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    resource_name = app_instance.resource_var.get()
    if not resource_name:
        print("🚫 Please select a VISA resource first.")
        debug_print("No VISA resource selected.", file=current_file, function=current_function)
        return

    print(f"\nAttempting to connect to {resource_name}...")
    debug_print(f"Attempting to connect to {resource_name}", file=current_file, function=current_function)

    try:
        app_instance.inst, app_instance.instrument_model = connect_to_instrument(app_instance.rm, resource_name)
        if app_instance.inst:
            app_instance.gpib_device_var.set(resource_name) # Save last connected device
            save_config(app_instance) # Save config immediately after successful connection

            print(f"✅ Successfully connected to {app_instance.instrument_model} at {resource_name}")
            debug_print(f"Successfully connected to {app_instance.instrument_model} at {resource_name}", file=current_file, function=current_function)
            
            # Initialize instrument settings
            print("Initializing instrument settings...")
            debug_print("Initializing instrument settings...", file=current_file, function=current_function)
            
            # Get current settings from Tkinter variables
            scan_rbw_hz = _get_float_value(app_instance.desired_rbw_var, app_instance.config.get('DEFAULT_SETTINGS', 'default_scan_rbw_hz'), 'Scan RBW')
            ref_level_dbm = _get_float_value(app_instance.desired_reference_level_var, app_instance.config.get('DEFAULT_SETTINGS', 'default_reference_level_dbm'), 'Reference Level')
            maxhold_enabled = _get_bool_value(app_instance.desired_maxhold_enabled_var, app_instance.config.getboolean('DEFAULT_SETTINGS', 'default_maxhold_enabled'), 'Max Hold Enabled')
            high_sensitivity = _get_bool_value(app_instance.desired_high_sensitivity_var, app_instance.config.getboolean('DEFAULT_SETTINGS', 'default_high_sensitivity'), 'High Sensitivity')
            preamp_on = _get_bool_value(app_instance.desired_preamp_on_var, app_instance.config.getboolean('DEFAULT_SETTINGS', 'default_preamp_on'), 'Preamplifier ON')
            
            # Pass app_instance.instrument_model to initialize_instrument
            if initialize_instrument(app_instance.inst, ref_level_dbm, scan_rbw_hz, maxhold_enabled, high_sensitivity, preamp_on, app_instance.instrument_model):
                print("✅ Instrument initialized with current settings.")
                debug_print("Instrument initialized with current settings.", file=current_file, function=current_function)
            else:
                print("❌ Failed to initialize instrument settings.")
                debug_print("Failed to initialize instrument settings.", file=current_file, function=current_function)
                # Don't return False here, connection is still established even if init fails

            app_instance.connect_button.config(state=tk.DISABLED)
            app_instance.disconnect_button.config(state=tk.NORMAL)
            app_instance.start_scan_button.config(state=tk.NORMAL)
            app_instance.apply_button.config(state=tk.NORMAL)
            app_instance._stop_connect_button_blink() # Stop blinking on successful connection

            # Enable query presets button if instrument is connected
            if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
                app_instance.preset_files_tab.query_presets_button.config(state=tk.NORMAL)

        else:
            print("❌ Connection failed: Instrument object is None.")
            debug_print("Connection failed: Instrument object is None.", file=current_file, function=current_function)
            app_instance.connect_button.config(state=tk.NORMAL) # Re-enable connect button
            app_instance._start_connect_button_blink() # Start blinking again
            
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error during connection: {e}")
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: tk.messagebox.showerror("VISA Connection Error", f"Failed to connect to instrument: {e}"))
        debug_print(f"VISA error during connection: {e}", file=current_file, function=current_function)
        app_instance._reset_gui_on_disconnect_or_error()
    except Exception as e:
        print(f"❌ An unexpected error occurred during connection: {e}")
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: tk.messagebox.showerror("Connection Error", f"An unexpected error occurred: {e}"))
        debug_print(f"Unexpected error during connection: {e}", file=current_file, function=current_function)
        app_instance._reset_gui_on_disconnect_or_error()


def disconnect_instrument_logic(app_instance):
    """
    Disconnects from the currently connected instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if app_instance.inst:
        print("\nAttempting to disconnect instrument...")
        debug_print("Attempting to disconnect instrument...", file=current_file, function=current_function)
        try:
            control_disconnect_instrument(app_instance.inst)
            print("✅ Instrument disconnected.")
            debug_print("Instrument disconnected.", file=current_file, function=current_function)
        except Exception as e:
            print(f"❌ Error during disconnection: {e}")
            debug_print(f"Error during disconnection: {e}", file=current_file, function=current_function)
    else:
        print("🚫 No instrument is currently connected.")
        debug_print("No instrument connected to disconnect.", file=current_file, function=current_function)
    
    app_instance._reset_gui_on_disconnect_or_error()


def apply_settings_to_device_logic(app_instance):
    """
    Applies the current settings from the GUI to the connected instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not app_instance.inst:
        print("🚫 Not connected to instrument. Cannot apply settings.")
        debug_print("Not connected to instrument, cannot apply settings.", file=current_file, function=current_function)
        return

    print("\nApplying settings to device...")
    debug_print("Applying settings to device...", file=current_file, function=current_function)

    try:
        scan_rbw_hz = _get_float_value(app_instance.desired_rbw_var, app_instance.config.get('DEFAULT_SETTINGS', 'default_scan_rbw_hz'), 'Scan RBW')
        ref_level_dbm = _get_float_value(app_instance.desired_reference_level_var, app_instance.config.get('DEFAULT_SETTINGS', 'default_reference_level_dbm'), 'Reference Level')
        maxhold_enabled = _get_bool_value(app_instance.desired_maxhold_enabled_var, app_instance.config.getboolean('DEFAULT_SETTINGS', 'default_maxhold_enabled'), 'Max Hold Enabled')
        high_sensitivity = _get_bool_value(app_instance.desired_high_sensitivity_var, app_instance.config.getboolean('DEFAULT_SETTINGS', 'default_high_sensitivity'), 'High Sensitivity')
        preamp_on = _get_bool_value(app_instance.desired_preamp_on_var, app_instance.config.getboolean('DEFAULT_SETTINGS', 'default_preamp_on'), 'Preamplifier ON')

        # Set RBW
        if not app_instance.inst.write(f":SENSe:BANDwidth:RESolution {scan_rbw_hz}"): return False
        debug_print(f"Sent: :SENSe:BANDwidth:RESolution {scan_rbw_hz}", file=current_file, function=current_function)

        # Set VBW (RBW / 3)
        vbw_hz = scan_rbw_hz * VBW_RBW_RATIO
        if not app_instance.inst.write(f":SENSe:BANDwidth:VIDeo {vbw_hz}"): return False
        debug_print(f"Sent: :SENSe:BANDwidth:VIDeo {vbw_hz}", file=current_file, function=current_function)

        # Set Reference Level
        if not app_instance.inst.write(f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM"): return False
        debug_print(f"Sent: :DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM", file=current_file, function=current_function)

        # Set Max Hold
        max_hold_state = "ON" if maxhold_enabled else "OFF"
        if not app_instance.inst.write(f":DISPlay:WINDow:TRACe:MODE {max_hold_state}"): return False
        debug_print(f"Sent: :DISPlay:WINDow:TRACe:MODE {max_hold_state}", file=current_file, function=current_function)

        # Set High Sensitivity (N9340B specific)
        if app_instance.instrument_model == "N9340B":
            high_sensitivity_state = "ON" if high_sensitivity else "OFF"
            if not app_instance.inst.write(f":SENSe:POWer:RF:HSENse {high_sensitivity_state}"): return False
            debug_print(f"Sent: :SENSe:POWer:RF:HSENse {high_sensitivity_state}", file=current_file, function=current_function)
        else:
            debug_print("Instrument is not N9340B, skipping High Sensitivity setting.", file=current_file, function=current_function)

        # Set Preamplifier (N9340B specific)
        if app_instance.instrument_model == "N9340B":
            preamp_state = "ON" if preamp_on else "OFF"
            if not app_instance.inst.write(f":SENSe:POWer:RF:GAIN:STATe {preamp_state}"): return False
            debug_print(f"Sent: :SENSe:POWer:RF:GAIN:STATe {preamp_state}", file=current_file, function=current_function)
        else:
            debug_print("Instrument is not N9340B, skipping Preamplifier setting.", file=current_file, function=current_function)
        
        print("✅ Settings applied to device successfully.")
        debug_print("Settings applied to device successfully.", file=current_file, function=current_function)
        save_config(app_instance) # Save settings after applying
        app_instance.reset_setting_colors_logic() # Reset colors after successful apply

    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error applying settings: {e}")
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: tk.messagebox.showerror("VISA Error", f"Failed to apply settings to instrument: {e}"))
        debug_print(f"VISA error applying settings: {e}", file=current_file, function=current_function)
    except Exception as e:
        print(f"❌ An unexpected error occurred while applying settings: {e}")
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: tk.messagebox.showerror("Error", f"An unexpected error occurred while applying settings: {e}"))
        debug_print(f"Unexpected error applying settings: {e}", file=current_file, function=current_function)


def load_selected_preset_logic(app_instance, selected_preset_name):
    """
    Loads a selected instrument preset (.sta file) onto the connected instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not app_instance.inst:
        print("🚫 Not connected to instrument, cannot load preset.")
        debug_print("Not connected to instrument, cannot load preset.", file=current_file, function=current_function)
        return

    if not selected_preset_name:
        print("🚫 No preset selected to load.")
        debug_print("No preset selected to load.", file=current_file, function=current_function)
        return

    print(f"\nAttempting to load preset: {selected_preset_name}")
    debug_print(f"Attempting to load preset: {selected_preset_name}", file=current_file, function=current_function)

    try:
        success, center_freq, span, rbw = control_load_selected_preset(app_instance.inst, selected_preset_name)
        if success:
            print(f"✅ Preset '{selected_preset_name}' loaded successfully.")
            debug_print(f"Preset '{selected_preset_name}' loaded successfully.", file=current_file, function=current_function)
            if app_instance.preset_files_tab:
                # Update the button text with queried settings
                app_instance.preset_files_tab.update_preset_button_info(selected_preset_name, center_freq, span, rbw)
        else:
            print(f"❌ Failed to load preset '{selected_preset_name}'.")
            debug_print(f"Failed to load preset '{selected_preset_name}'.", file=current_file, function=current_function)
            # Schedule messagebox to run on the main thread
            app_instance.after(0, lambda: tk.messagebox.showerror("Preset Load Error", f"Failed to load preset '{selected_preset_name}'. See console for details."))

    except Exception as e:
        print(f"❌ An unexpected error occurred while loading preset: {e}")
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: tk.messagebox.showerror("Error", f"An unexpected error occurred while loading preset: {e}"))
        debug_print(f"Unexpected error loading preset: {e}", file=current_file, function=current_function)


def query_device_presets_logic(app_instance):
    """
    Queries the connected instrument for available preset files and updates the PresetFilesTab.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not app_instance.inst:
        print("🚫 Not connected to instrument. Cannot query presets.")
        debug_print("Not connected to instrument, cannot query presets.", file=current_file, function=current_function)
        return

    print("\nQuerying device for presets...")
    debug_print("Querying device for presets...", file=current_file, function=current_function)

    try:
        # Pass the instrument model to the control function
        presets = control_query_device_presets(app_instance.inst, app_instance.instrument_model)
        if presets is not None:
            print(f"✅ Found {len(presets)} presets on device.")
            debug_print(f"Found {len(presets)} presets on device.", file=current_file, function=current_function)
            if app_instance.preset_files_tab:
                app_instance.preset_files_tab.populate_preset_buttons(presets, source="device")
        else:
            print("🚫 No presets found on device or device does not support preset querying.")
            debug_print("No presets found on device or device does not support preset querying.", file=current_file, function=current_function)
            # Schedule messagebox to run on the main thread
            app_instance.after(0, lambda: tk.messagebox.showwarning("No Device Presets", "No presets found on device or device does not support preset querying."))

    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error querying device presets: {e}")
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: tk.messagebox.showerror("VISA Error", f"Failed to query device presets: {e}"))
        debug_print(f"VISA error querying device presets: {e}", file=current_file, function=current_function)
    except Exception as e:
        print(f"❌ An unexpected error occurred while querying device presets: {e}")
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: tk.messagebox.showerror("Error", f"An unexpected error occurred while querying device presets: {e}"))
        debug_print(f"Unexpected error querying device presets: {e}", file=current_file, function=current_function)


def set_focus_frequency_logic(app_instance, frequency_hz, span_hz):
    """
    Sets the instrument's center frequency and span.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not app_instance.inst:
        print("🚫 Not connected to instrument. Cannot set focus frequency.")
        debug_print("Not connected to instrument, cannot set focus frequency.", file=current_file, function=current_function)
        return False

    print(f"\nSetting instrument focus: Center Freq {frequency_hz / MHZ_TO_HZ:.3f} MHz, Span {span_hz / MHZ_TO_HZ:.3f} MHz")
    debug_print(f"Setting instrument focus: Center Freq {frequency_hz / MHZ_TO_HZ:.3f} MHz, Span {span_hz / MHZ_TO_HZ:.3f} MHz", file=current_file, function=current_function)

    try:
        # Set Center Frequency
        if not app_instance.inst.write(f":SENSe:FREQuency:CENTer {frequency_hz}"): return False
        debug_print(f"Sent: :SENSe:FREQuency:CENTer {frequency_hz}", file=current_file, function=current_function)

        # Set Span
        if not app_instance.inst.write(f":SENSe:FREQuency:SPAN {span_hz}"): return False
        debug_print(f"Sent: :SENSe:FREQuency:SPAN {span_hz}", file=current_file, function=current_function)

        # Set trace type to Normal (or other desired mode like Clear Write)
        # This might be needed if the instrument was in Max Hold or Min Hold
        if not app_instance.inst.write(":DISPlay:WINDow:TRACe:TYPE NORM"): return False
        debug_print("Sent: :DISPlay:WINDow:TRACe:TYPE NORM", file=current_file, function=current_function)

        print(f"✅ Instrument focus set to {frequency_hz / MHZ_TO_HZ:.3f} MHz center, {span_hz / MHZ_TO_HZ:.3f} MHz span. Trace mode set to Normal.")
        debug_print(f"Instrument focus set to {frequency_hz / MHZ_TO_HZ:.3f} MHz center, {span_hz / MHZ_TO_HZ:.3f} MHz span.", file=current_file, function=current_function)
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error while setting focus frequency/span: {e}")
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: tk.messagebox.showerror("VISA Error", f"Failed to set instrument focus frequency/span: {e}"))
        debug_print(f"VISA Error setting focus frequency/span: {e}", file=current_file, function=current_function)
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred while setting focus frequency/span: {e}")
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: tk.messagebox.showerror("Error", f"An unexpected error occurred while setting focus frequency/span: {e}"))
        debug_print(f"Unexpected error setting focus frequency/span: {e}", file=current_file, function=current_function)
        return False


def set_marker_and_trace_modes_logic(app_instance, marker_frequency_hz, marker_name):
    """
    Sets a marker on the instrument at the specified frequency and configures trace modes.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not app_instance.inst:
        print("🚫 Not connected to instrument. Cannot set marker.")
        debug_print("Not connected to instrument, cannot set marker.", file=current_file, function=current_function)
        return False

    print(f"\nSetting marker '{marker_name}' at {marker_frequency_hz / MHZ_TO_HZ:.3f} MHz...")
    debug_print(f"Setting marker '{marker_name}' at {marker_frequency_hz / MHZ_TO_HZ:.3f} MHz", file=current_file, function=current_function)

    try:
        # Activate Marker 1
        if not app_instance.inst.write(":CALCulate:MARKer1:STATe ON"): return False
        debug_print("Sent: :CALCulate:MARKer1:STATe ON", file=current_file, function=current_function)

        # Set Marker 1 to specified frequency
        if not app_instance.inst.write(f":CALCulate:MARKer1:X {marker_frequency_hz}"): return False
        debug_print(f"Sent: :CALCulate:MARKer1:X {marker_frequency_hz}", file=current_file, function=current_function)

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
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: tk.messagebox.showerror("VISA Error", f"Failed to set instrument marker or trace modes: {e}"))
        debug_print(f"VISA Error setting marker/trace modes: {e}", file=current_file, function=current_function)
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred while setting marker/trace modes: {e}")
        # Schedule messagebox to run on the main thread
        app_instance.after(0, lambda: tk.messagebox.showerror("Error", f"An unexpected error occurred while setting marker/trace modes: {e}"))
        debug_print(f"Unexpected error setting marker/trace modes: {e}", file=current_file, function=current_function)
        return False

