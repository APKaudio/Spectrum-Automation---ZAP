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
    query_current_instrument_settings, # This is the one imported from utils
    query_device_presets as control_query_device_presets, # Alias the imported function
    load_selected_preset as control_load_selected_preset,
    debug_print # Import debug_print
)
from utils.frequency_bands import MHZ_TO_HZ, VBW_RBW_RATIO # Import VBW_RBW_RATIO
from src.config_manager import save_config # Import save_config
import tkinter.ttk as ttk # Import ttk for themed widgets

def _get_float_value(tk_var, default_value, setting_name, console_print_func):
    """
    Safely retrieves a float value from a Tkinter StringVar.
    If the string is empty or cannot be converted, returns a default value.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    try:
        val_str = tk_var.get()
        if not val_str:
            debug_print(f"Warning: Tkinter variable for '{setting_name}' is empty. Using default value: {default_value}", file=current_file, function=current_function, console_print_func=console_print_func)
            return float(default_value)
        return float(val_str)
    except ValueError:
        console_print_func(f"❌ Error: Invalid value for '{setting_name}': '{tk_var.get()}'. Using default value: {default_value}")
        # Corrected f-string: changed default_file to current_file
        debug_print(f"ValueError for '{setting_name}': '{tk_var.get()}'. Using default: {default_value}", file=current_file, function=current_function, console_print_func=console_print_func)
        return float(default_value)


def populate_resources_logic(app_instance, console_print_func):
    """
    Populates the VISA resource dropdown with available instruments.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Populating VISA resources...", file=current_file, function=current_function, console_print_func=console_print_func)
    
    resources = list_visa_resources(console_print_func)
    # Corrected: app_instance is the main App, so access its resource_combobox attribute
    app_instance.resource_combobox['values'] = resources 
    if resources:
        # Try to select the last used device, otherwise select the first one
        last_device = app_instance.config.get('LAST_USED_SETTINGS', 'last_gpib_device', fallback='')
        if last_device in resources:
            app_instance.resource_combobox.set(last_device) 
        else:
            app_instance.resource_combobox.set(resources[0]) 
        debug_print(f"Found resources: {resources}", file=current_file, function=current_function, console_print_func=console_print_func)
    else:
        app_instance.resource_combobox.set("No resources found") 
        debug_print("No VISA resources found.", file=current_file, function=current_function, console_print_func=console_print_func)


def connect_instrument_logic(app_instance, selected_resource, console_print_func):
    """
    Connects to the selected instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Attempting to connect instrument...", file=current_file, function=current_function, console_print_func=console_print_func)

    if not selected_resource or selected_resource == "No resources found":
        console_print_func("⚠️ Warning: Please select a valid VISA resource.")
        debug_print("No valid VISA resource selected for connection.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    try:
        inst = connect_to_instrument(selected_resource, console_print_func) # Pass console_print_func
        if inst:
            # Initialize the instrument with default settings after connection
            console_print_func(f"Attempting to initialize instrument: {selected_resource}")
            init_success, instrument_model = initialize_instrument(inst, console_print_func) # Pass console_print_func and get model
            if not init_success:
                console_print_func(f"❌ Error: Failed to initialize instrument {selected_resource}. Disconnecting.")
                debug_print(f"Initialization failed for {selected_resource}.", file=current_file, function=current_function, console_print_func=console_print_func)
                control_disconnect_instrument(inst, console_print_func) # Pass console_print_func
                app_instance.inst = None
                app_instance.instrument_model = None
                return False
            
            app_instance.inst = inst
            app_instance.instrument_model = instrument_model
            console_print_func(f"✅ Successfully connected and initialized to {selected_resource} (Model: {instrument_model})")
            
            # Query and display current instrument settings after connection and initialization
            query_current_instrument_settings_logic(app_instance, console_print_func)

            return True
        else:
            console_print_func(f"❌ Error: Failed to connect to {selected_resource}")
            debug_print(f"Failed to connect to {selected_resource}.", file=current_file, function=current_function, console_print_func=console_print_func)
            app_instance.inst = None
            app_instance.instrument_model = None
            return False
    except Exception as e:
        console_print_func(f"❌ An unexpected error occurred during connection: {e}")
        debug_print(f"Unexpected error during connection: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        app_instance.inst = None
        app_instance.instrument_model = None
        return False


def disconnect_instrument_logic(app_instance, console_print_func):
    """
    Disconnects from the currently connected instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Attempting to disconnect instrument...", file=current_file, function=current_function, console_print_func=console_print_func)

    if app_instance.inst:
        control_disconnect_instrument(app_instance.inst, console_print_func) # Pass console_print_func
        app_instance.inst = None # Clear the instrument instance
        app_instance.instrument_model = None # Clear the instrument model
        console_print_func("✅ Instrument disconnected.")
    else:
        console_print_func("ℹ️ Info: No instrument is currently connected.")
    debug_print("Disconnect instrument logic complete.", file=current_file, function=current_function, console_print_func=console_print_func)


def apply_settings_to_device_logic(app_instance, console_print_func):
    """
    Applies the current settings from the GUI to the connected instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Applying settings to device...", file=current_file, function=current_function, console_print_func=console_print_func)

    if not app_instance.inst:
        console_print_func("⚠️ Warning: No instrument connected. Cannot apply settings.")
        debug_print("No instrument connected for applying settings.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    inst = app_instance.inst

    try:
        # Get values from Tkinter variables, using helper for safety
        rbw_step_size_hz = _get_float_value(app_instance.rbw_step_size_hz_var, 1000000.0, "RBW Step Size", console_print_func)
        cycle_wait_time_seconds = _get_float_value(app_instance.cycle_wait_time_seconds_var, 0.5, "Cycle Wait Time", console_print_func)
        maxhold_time_seconds = _get_float_value(app_instance.maxhold_time_seconds_var, 3.0, "Max Hold Time", console_print_func)
        scan_rbw_hz = _get_float_value(app_instance.scan_rbw_hz_var, 10000.0, "Scan RBW", console_print_func)
        reference_level_dbm = _get_float_value(app_instance.reference_level_dbm_var, -40.0, "Reference Level", console_print_func)
        freq_shift_hz = _get_float_value(app_instance.freq_shift_hz_var, 0.0, "Frequency Shift", console_print_func)
        scan_rbw_segmentation = _get_float_value(app_instance.scan_rbw_segmentation_var, 1000000.0, "Scan RBW Segmentation", console_print_func)
        
        num_scan_cycles = app_instance.num_scan_cycles_var.get()
        maxhold_enabled = app_instance.maxhold_enabled_var.get()
        high_sensitivity = app_instance.high_sensitivity_var.get()
        preamp_on = app_instance.preamp_on_var.get()

        console_print_func("\nApplying settings to instrument...")

        # Set Reference Level
        if not inst.write(f":POW:RF:RLEVEL {reference_level_dbm}DBM"): return False
        debug_print(f"Sent: :POW:RF:RLEVEL {reference_level_dbm}DBM", file=current_file, function=current_function, console_print_func=console_print_func)
        console_print_func(f"✅ Reference Level set to {reference_level_dbm} dBm.")

        # Set RBW
        if not inst.write(f":BAND:RES {scan_rbw_hz}"): return False
        debug_print(f"Sent: :BAND:RES {scan_rbw_hz}", file=current_file, function=current_function, console_print_func=console_print_func)
        console_print_func(f"✅ RBW set to {scan_rbw_hz} Hz.")

        # Set VBW (calculated from RBW)
        vbw_hz = scan_rbw_hz * VBW_RBW_RATIO
        if not inst.write(f":BAND:VID {vbw_hz}"): return False
        debug_print(f"Sent: :BAND:VID {vbw_hz}", file=current_file, function=current_function, console_print_func=console_print_func)
        console_print_func(f"✅ VBW set to {vbw_hz} Hz (RBW/3).")

        # Set High Sensitivity (if applicable)
        if app_instance.instrument_model != "N9340B": # N9340B does not support high sensitivity
            sensitivity_state = "ON" if high_sensitivity else "OFF"
            if not inst.write(f":SENSe:SPECTRUM:HSENS {sensitivity_state}"): return False
            debug_print(f"Sent: :SENSe:SPECTRUM:HSENS {sensitivity_state}", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"✅ High Sensitivity set to {sensitivity_state}.")
        else:
            console_print_func("ℹ️ Info: N9340B does not support High Sensitivity setting.")
            debug_print("N9340B does not support High Sensitivity.", file=current_file, function=current_function, console_print_func=console_print_func)

        # Set Preamp
        preamp_state = "ON" if preamp_on else "OFF"
        if not inst.write(f":POWer:ATTenuator:PREamp {preamp_state}"): return False
        debug_print(f"Sent: :POWer:ATTenuator:PREamp {preamp_state}", file=current_file, function=current_function, console_print_func=console_print_func)
        console_print_func(f"✅ Preamp set to {preamp_state}.")

        # Set Max Hold (if applicable)
        trace_type = "MAXH" if maxhold_enabled else "NORM" # Normal if max hold is off
        if not inst.write(f":DISPlay:WINDow:TRACe:TYPE {trace_type}"): return False
        debug_print(f"Sent: :DISPlay:WINDow:TRACe:TYPE {trace_type}", file=current_file, function=current_function, console_print_func=console_print_func)
        console_print_func(f"✅ Trace mode set to {trace_type}.")

        console_print_func("🎉 All settings applied successfully to the instrument.")
        return True

    except pyvisa.errors.VisaIOError as e:
        console_print_func(f"❌ VISA error applying settings: {e}")
        debug_print(f"VISA Error applying settings: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False
    except Exception as e:
        console_print_func(f"❌ An unexpected error occurred while applying settings: {e}")
        debug_print(f"An unexpected error occurred while applying settings: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False


def set_focus_frequency_logic(app_instance, frequency_mhz, marker_name, console_print_func):
    """
    Sets the instrument's center frequency and span to focus on a specific marker frequency.
    The span is determined by the desired_default_focus_width_var.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Setting focus frequency for {marker_name} at {frequency_mhz} MHz...", file=current_file, function=current_function, console_print_func=console_print_func)

    if not app_instance.inst:
        console_print_func("⚠️ Warning: No instrument connected. Cannot set focus frequency.")
        debug_print("No instrument connected for setting focus frequency.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    inst = app_instance.inst
    
    # Get the desired focus width from the Tkinter variable (in MHz)
    try:
        focus_width_mhz = float(app_instance.desired_default_focus_width_var.get())
    except ValueError:
        console_print_func("⚠️ Warning: Invalid value for Default Focus Width. Using 1 MHz.")
        focus_width_mhz = 1.0 # Fallback to a default value
        debug_print("Invalid focus width, defaulting to 1 MHz.", file=current_file, function=current_function, console_print_func=console_print_func)

    if focus_width_mhz <= 0:
        console_print_func("⚠️ Warning: Default Focus Width must be a positive value. Using 1 MHz.")
        focus_width_mhz = 1.0 # Fallback to a default value
        debug_print("Invalid focus width, defaulting to 1 MHz.", file=current_file, function=current_function, console_print_func=console_print_func)

    # Convert to Hz for instrument commands
    center_frequency_hz = frequency_mhz * MHZ_TO_HZ
    span_hz = focus_width_mhz * MHZ_TO_HZ

    try:
        # Set Center Frequency
        if not inst.write(f":SENSe:FREQuency:CENTer {center_frequency_hz}"): return False
        debug_print(f"Sent: :SENSe:FREQuency:CENTer {center_frequency_hz}", file=current_file, function=current_function, console_print_func=console_print_func)

        # Set Span
        if not inst.write(f":SENSe:FREQuency:SPAN {span_hz}"): return False
        debug_print(f"Sent: :SENSe:FREQuency:SPAN {span_hz}", file=current_file, function=current_function, console_print_func=console_print_func)

        console_print_func(f"✅ Instrument focus set to Center: {frequency_mhz:.3f} MHz, Span: {focus_width_mhz:.3f} MHz.")
        
        # Update current instrument settings display in GUI (if these vars exist in app_instance)
        if hasattr(app_instance, 'current_center_freq_var'):
            app_instance.after(0, lambda: app_instance.current_center_freq_var.set(f"{frequency_mhz:.3f} MHz"))
        if hasattr(app_instance, 'current_span_var'):
            app_instance.after(0, lambda: app_instance.current_span_var.set(f"{focus_width_mhz:.3f} MHz"))
        
        return True
    except pyvisa.errors.VisaIOError as e:
        console_print_func(f"❌ VISA error setting focus frequency: {e}")
        debug_print(f"VISA Error setting focus frequency: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False
    except Exception as e:
        console_print_func(f"❌ An unexpected error occurred while setting focus frequency: {e}")
        debug_print(f"An unexpected error occurred while setting focus frequency: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False


def query_current_instrument_settings_logic(app_instance, console_print_func):
    """
    Queries the current instrument settings (center frequency, span, RBW)
    and updates the GUI display.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Querying current instrument settings...", file=current_file, function=current_function, console_print_func=console_print_func)

    if not app_instance.inst:
        console_print_func("⚠️ Warning: No instrument connected. Cannot query current settings.")
        debug_print("No instrument connected for querying current settings.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    try:
        center_freq_hz, span_hz, rbw_hz = query_current_instrument_settings(app_instance.inst, console_print_func) # Pass console_print_func
        
        center_freq_mhz = center_freq_hz / MHZ_TO_HZ if center_freq_hz is not None else None
        span_mhz = span_hz / MHZ_TO_HZ if span_hz is not None else None
        rbw_khz = rbw_hz / 1000 if rbw_hz is not None else None

        # Update current instrument settings display in GUI (if these vars exist in app_instance)
        if hasattr(app_instance, 'current_center_freq_var'):
            app_instance.after(0, lambda: app_instance.current_center_freq_var.set(f"{center_freq_mhz:.3f} MHz" if center_freq_mhz is not None else "N/A"))
        if hasattr(app_instance, 'current_span_var'):
            app_instance.after(0, lambda: app_instance.current_span_var.set(f"{span_mhz:.3f} MHz" if span_mhz is not None else "N/A"))
        if hasattr(app_instance, 'current_rbw_var'):
            app_instance.after(0, lambda: app_instance.current_rbw_var.set(f"{rbw_khz:.1f} kHz" if rbw_khz is not None else "N/A"))
        
        console_print_func(f"✅ Current instrument settings: Center Freq: {center_freq_mhz:.3f} MHz, Span: {span_mhz:.3f} MHz, RBW: {rbw_khz:.1f} kHz")
        debug_print("Current instrument settings queried and displayed.", file=current_file, function=current_function, console_print_func=console_print_func)
        return True
    except pyvisa.errors.VisaIOError as e:
        console_print_func(f"❌ VISA error querying current settings: {e}")
        debug_print(f"VISA Error querying current settings: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False
    except Exception as e:
        console_print_func(f"❌ An unexpected error occurred while querying current settings: {e}")
        debug_print(f"An unexpected error occurred while querying current settings: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False


def load_selected_preset_logic(app_instance, selected_preset_name, console_print_func):
    """
    Loads a specified preset file onto the connected instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Attempting to load preset: {selected_preset_name}", file=current_file, function=current_function, console_print_func=console_print_func)

    if not app_instance.inst:
        console_print_func("⚠️ Warning: No instrument connected. Cannot load preset.")
        debug_print("No instrument connected for loading preset.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False, None, None, None

    success, center_freq, span, rbw = control_load_selected_preset(app_instance.inst, selected_preset_name, console_print_func) # Pass console_print_func
    
    if success:
        # Update current instrument settings display in GUI (if these vars exist in app_instance)
        if hasattr(app_instance, 'current_center_freq_var'):
            app_instance.after(0, lambda: app_instance.current_center_freq_var.set(f"{center_freq:.3f} MHz" if center_freq else "N/A"))
        if hasattr(app_instance, 'current_span_var'):
            app_instance.after(0, lambda: app_instance.current_span_var.set(f"{span:.3f} MHz" if span else "N/A"))
        if hasattr(app_instance, 'current_rbw_var'):
            app_instance.after(0, lambda: app_instance.current_rbw_var.set(f"{rbw / 1000:.1f} kHz" if rbw else "N/A"))
        console_print_func(f"✅ Preset '{selected_preset_name}' loaded successfully.")
    else:
        console_print_func(f"❌ Error: Failed to load preset '{selected_preset_name}'.")
    
    return success, center_freq, span, rbw


def query_device_presets_logic(app_instance, console_print_func):
    """
    Queries the connected instrument for available preset (.sta) files.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Querying device presets...", file=current_file, function=current_function, console_print_func=console_print_func)

    if not app_instance.inst:
        console_print_func("⚠️ Warning: No instrument connected. Cannot query presets.")
        debug_print("No instrument connected for querying presets.", file=current_file, function=current_function, console_print_func=console_print_func)
        return []

    presets = control_query_device_presets(app_instance.inst, console_print_func) # Pass console_print_func
    if presets:
        console_print_func(f"✅ Found {len(presets)} presets on the instrument.")
    else:
        console_print_func("ℹ️ Info: No presets found on the instrument or failed to retrieve.")
    return presets


def set_marker_and_trace_modes_logic(app_instance, marker_frequency_hz, marker_name, console_print_func):
    """
    Sets a marker at the specified frequency and ensures the trace mode is Normal.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Setting marker at {marker_frequency_hz} Hz for '{marker_name}'...", file=current_file, function=current_function, console_print_func=console_print_func)

    if not app_instance.inst:
        console_print_func("⚠️ Warning: No instrument connected. Cannot set marker or trace modes.")
        debug_print("No instrument connected for setting marker/trace modes.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    try:
        # Set marker 1 to the specified frequency
        if not app_instance.inst.write(f":CALCulate:MARKer1:X {marker_frequency_hz}"): return False
        debug_print(f"Sent: :CALCulate:MARKer1:X {marker_frequency_hz}", file=current_file, function=current_function, console_print_func=console_print_func)

        # Activate marker 1
        if not app_instance.inst.write(":CALCulate:MARKer1:STATe ON"): return False
        debug_print("Sent: :CALCulate:MARKer1:STATe ON", file=current_file, function=current_function, console_print_func=console_print_func)

        # Set marker 1 to peak (optional, but often desired for markers)
        if not app_instance.inst.write(":CALCulate:MARKer1:MAXimum:PEAK"): return False
        debug_print("Sent: :CALCulate:MARKer1:MAXimum:PEAK", file=current_file, function=current_function, console_print_func=console_print_func)

        # Set trace type to Normal (or other desired mode like Clear Write)
        # This might be needed if the instrument was in Max Hold or Min Hold
        if not app_instance.inst.write(":DISPlay:WINDow:TRACe:TYPE NORM"): return False
        debug_print(f"Sent: :DISPlay:WINDow:TRACe:TYPE NORM", file=current_file, function=current_function, console_print_func=console_print_func)

        console_print_func(f"✅ Marker '{marker_name}' set at {marker_frequency_hz / MHZ_TO_HZ:.3f} MHz. Trace mode set to Normal.")
        return True
    except pyvisa.errors.VisaIOError as e:
        console_print_func(f"❌ VISA error while setting marker/trace modes: {e}")
        debug_print(f"VISA Error setting marker/trace modes: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False
    except Exception as e:
        console_print_func(f"❌ An unexpected error occurred while setting marker/trace modes: {e}")
        debug_print(f"An unexpected error occurred while setting marker/trace modes: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

