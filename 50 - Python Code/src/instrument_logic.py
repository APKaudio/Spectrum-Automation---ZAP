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
    debug_print # Import debug_print
)
from utils.frequency_bands import MHZ_TO_HZ, VBW_RBW_RATIO # Import VBW_RBW_RATIO
from src.config_manager import save_config # Import save_config
import tkinter.ttk as ttk # Import ttk for themed widgets

# Import set_marker_and_trace_modes_logic from marker_utils
from utils.marker_utils import set_marker_and_trace_modes_logic


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
            debug_print(f"Warning: Tkinter variable for '{setting_name}' is empty. Using default: {default_value}", file=current_file, function=current_function, console_print_func=console_print_func)
            return default_value
        return float(val_str)
    except ValueError:
        console_print_func(f"❌ Error: Invalid numeric value for {setting_name}: '{val_str}'. Using default: {default_value}")
        debug_print(f"ValueError: Invalid numeric value for {setting_name}: '{val_str}'. Using default: {default_value}", file=current_file, function=current_function, console_print_func=console_print_func)
        return default_value
    except Exception as e:
        console_print_func(f"❌ An unexpected error occurred while getting {setting_name}: {e}")
        debug_print(f"Unexpected error getting {setting_name}: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return default_value

def populate_resources_logic(app_instance, console_print_func):
    """
    Populates the VISA resource dropdown with available instruments.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Populating VISA resources...", file=current_file, function=current_function, console_print_func=console_print_func)
    
    resources = list_visa_resources(console_print_func)
    app_instance.resource_names.set([]) # Clear existing options
    if resources:
        app_instance.resource_names.set(resources)
        # Attempt to set the last used resource if it's still available
        last_used_resource = app_instance.config.get('LAST_USED_SETTINGS', 'last_gpib_device', fallback='')
        if last_used_resource and last_used_resource in resources:
            app_instance.selected_resource.set(last_used_resource)
            console_print_func(f"✅ Last used resource '{last_used_resource}' found and selected.")
            debug_print(f"Last used resource '{last_used_resource}' found and selected.", file=current_file, function=current_function, console_print_func=console_print_func)
        else:
            app_instance.selected_resource.set(resources[0]) # Select the first resource by default
            console_print_func(f"✅ Resources found. Selected: {resources[0]}")
            debug_print(f"Resources found. Selected: {resources[0]}", file=current_file, function=current_function, console_print_func=console_print_func)
    else:
        app_instance.selected_resource.set("No Resources Found")
        console_print_func("❌ No VISA resources found. Ensure NI-VISA is installed and instrument is connected.")
        debug_print("No VISA resources found.", file=current_file, function=current_function, console_print_func=console_print_func)

def connect_instrument_logic(app_instance, console_print_func):
    """
    Connects to the selected VISA instrument and initializes it.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    resource_name = app_instance.selected_resource.get()
    console_print_func(f"\nAttempting to connect to {resource_name}...")
    debug_print(f"Attempting to connect to {resource_name}...", file=current_file, function=current_function, console_print_func=console_print_func)

    if resource_name == "No Resources Found" or not resource_name:
        console_print_func("⚠️ Warning: No valid VISA resource selected.")
        debug_print("No valid VISA resource selected for connection.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    try:
        inst = connect_to_instrument(resource_name, console_print_func)
        if inst:
            app_instance.inst = inst
            console_print_func(f"✅ Successfully connected to {resource_name}")
            debug_print(f"Successfully connected to {resource_name}", file=current_file, function=current_function, console_print_func=console_print_func)
            
            # Initialize instrument settings
            if initialize_instrument(app_instance.inst, console_print_func):
                console_print_func("✅ Instrument initialized with default settings.")
                debug_print("Instrument initialized with default settings.", file=current_file, function=current_function, console_print_func=console_print_func)
                
                # Query and display current instrument settings
                center_freq, span, rbw = query_current_instrument_settings(app_instance.inst, MHZ_TO_HZ, console_print_func)
                app_instance.center_freq_var.set(f"{center_freq:.3f}")
                app_instance.span_var.set(f"{span:.3f}")
                app_instance.rbw_var.set(f"{rbw:.3f}")
                console_print_func(f"Current Instrument Settings: Center Freq={center_freq:.3f} MHz, Span={span:.3f} MHz, RBW={rbw:.3f} MHz")
                
                app_instance.update_connection_status(True) # Update GUI status
                save_config(app_instance) # Save the successfully connected resource
                return True
            else:
                console_print_func("❌ Failed to initialize instrument.")
                debug_print("Failed to initialize instrument.", file=current_file, function=current_function, console_print_func=console_print_func)
                control_disconnect_instrument(app_instance.inst, console_print_func) # Disconnect if initialization fails
                app_instance.inst = None
                app_instance.update_connection_status(False)
                return False
        else:
            console_print_func(f"❌ Failed to connect to {resource_name}.")
            debug_print(f"Failed to connect to {resource_name}.", file=current_file, function=current_function, console_print_func=console_print_func)
            app_instance.update_connection_status(False)
            return False
    except Exception as e:
        console_print_func(f"❌ An error occurred during connection: {e}")
        debug_print(f"Error during connection to {resource_name}: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        app_instance.update_connection_status(False)
        return False

def disconnect_instrument_logic(app_instance, console_print_func):
    """
    Disconnects from the currently connected VISA instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    console_print_func("\nAttempting to disconnect instrument...")
    debug_print("Attempting to disconnect instrument...", file=current_file, function=current_function, console_print_func=console_print_func)

    if app_instance.inst:
        control_disconnect_instrument(app_instance.inst, console_print_func)
        app_instance.inst = None
        console_print_func("✅ Instrument disconnected.")
        debug_print("Instrument disconnected.", file=current_file, function=current_function, console_print_func=console_print_func)
    else:
        console_print_func("ℹ️ Info: No instrument to disconnect.")
        debug_print("No instrument to disconnect.", file=current_file, function=current_function, console_print_func=console_print_func)
    app_instance.update_connection_status(False) # Update GUI status

def apply_settings_logic(app_instance, console_print_func):
    """
    Applies the current settings from the GUI to the connected instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    console_print_func("\nAttempting to apply settings to instrument...")
    debug_print("Attempting to apply settings to instrument...", file=current_file, function=current_function, console_print_func=console_print_func)

    if not app_instance.inst:
        console_print_func("⚠️ Warning: No instrument connected. Cannot apply settings.")
        debug_print("No instrument connected. Cannot apply settings.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    try:
        # Get values from Tkinter variables, converting to appropriate types
        center_freq_mhz = _get_float_value(app_instance.center_freq_var, 100.0, "Center Frequency", console_print_func)
        span_mhz = _get_float_value(app_instance.span_var, 10.0, "Span", console_print_func)
        rbw_hz = _get_float_value(app_instance.rbw_var, 10000.0, "RBW", console_print_func)
        ref_level_dbm = _get_float_value(app_instance.ref_level_var, -40.0, "Reference Level", console_print_func)
        freq_shift_hz = _get_float_value(app_instance.freq_shift_var, 0.0, "Frequency Shift", console_print_func)
        max_hold_enabled = app_instance.max_hold_enabled_var.get()
        high_sensitivity = app_instance.high_sensitivity_var.get()
        preamp_on = app_instance.preamp_on_var.get() # This variable is redundant if high_sensitivity controls both.
                                                    # Keeping it for now as it's in the map.
        rbw_segmentation_hz = _get_float_value(app_instance.rbw_segmentation_var, 1_000_000.0, "RBW Segmentation", console_print_func)
        default_focus_width_mhz = _get_float_value(app_instance.desired_default_focus_width_var, 10.0, "Default Focus Width", console_print_func)

        # Convert frequencies to Hz for SCPI commands
        center_freq_hz = center_freq_mhz * MHZ_TO_HZ
        span_hz = span_mhz * MHZ_TO_HZ
        default_focus_width_hz = default_focus_width_mhz * MHZ_TO_HZ

        success = True

        # --- Apply Center Frequency ---
        debug_print(f"Querying current center frequency for comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_center_freq_str = app_instance.inst.query(":SENSe:FREQuency:CENTer?", console_print_func)
        current_center_freq_hz = float(current_center_freq_str) if current_center_freq_str else None
        
        if current_center_freq_hz is not None and abs(current_center_freq_hz - center_freq_hz) < 1: # Tolerance for float comparison
            debug_print(f"Center frequency already at {center_freq_hz} Hz. Skipping command.", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"ℹ️ Info: Center frequency already at {center_freq_mhz:.3f} MHz.")
        else:
            debug_print(f"Setting center frequency to {center_freq_hz} Hz...", file=current_file, function=current_function, console_print_func=console_print_func)
            if not app_instance.inst.write(f":SENSe:FREQuency:CENTer {center_freq_hz}"):
                success = False
                console_print_func(f"❌ Failed to set Center Frequency to {center_freq_mhz:.3f} MHz.")
                debug_print(f"Failed to set Center Frequency.", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                console_print_func(f"✅ Center Frequency set to {center_freq_mhz:.3f} MHz.")

        # --- Apply Span ---
        debug_print(f"Querying current span for comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_span_str = app_instance.inst.query(":SENSe:FREQuency:SPAN?", console_print_func)
        current_span_hz = float(current_span_str) if current_span_str else None

        if current_span_hz is not None and abs(current_span_hz - span_hz) < 1:
            debug_print(f"Span already at {span_hz} Hz. Skipping command.", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"ℹ️ Info: Span already at {span_mhz:.3f} MHz.")
        else:
            debug_print(f"Setting span to {span_hz} Hz...", file=current_file, function=current_function, console_print_func=console_print_func)
            if not app_instance.inst.write(f":SENSe:FREQuency:SPAN {span_hz}"):
                success = False
                console_print_func(f"❌ Failed to set Span to {span_mhz:.3f} MHz.")
                debug_print(f"Failed to set Span.", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                console_print_func(f"✅ Span set to {span_mhz:.3f} MHz.")

        # --- Apply RBW ---
        debug_print(f"Querying current RBW for comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_rbw_str = app_instance.inst.query(":SENSe:BANDwidth:RESolution?", console_print_func)
        current_rbw_hz = float(current_rbw_str) if current_rbw_str else None

        if current_rbw_hz is not None and abs(current_rbw_hz - rbw_hz) < 1:
            debug_print(f"RBW already at {rbw_hz} Hz. Skipping command.", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"ℹ️ Info: RBW already at {rbw_hz:.0f} Hz.")
        else:
            debug_print(f"Setting RBW to {rbw_hz} Hz...", file=current_file, function=current_function, console_print_func=console_print_func)
            if not app_instance.inst.write(f":SENSe:BANDwidth:RESolution {rbw_hz}"):
                success = False
                console_print_func(f"❌ Failed to set RBW to {rbw_hz:.0f} Hz.")
                debug_print(f"Failed to set RBW.", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                console_print_func(f"✅ RBW set to {rbw_hz:.0f} Hz.")

        # --- Apply Reference Level ---
        debug_print(f"Querying current Reference Level for comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_ref_level_str = app_instance.inst.query(":DISPlay:WINDow:TRACe:Y:RLEVel?", console_print_func)
        current_ref_level_dbm = float(current_ref_level_str) if current_ref_level_str else None

        if current_ref_level_dbm is not None and abs(current_ref_level_dbm - ref_level_dbm) < 0.1: # Compare floats
            debug_print(f"Reference Level already at {ref_level_dbm} dBm. Skipping command.", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"ℹ️ Info: Reference Level already at {ref_level_dbm:.1f} dBm.")
        else:
            debug_print(f"Setting Reference Level to {ref_level_dbm} dBm...", file=current_file, function=current_function, console_print_func=console_print_func)
            if not app_instance.inst.write(f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}"):
                success = False
                console_print_func(f"❌ Failed to set Reference Level to {ref_level_dbm:.1f} dBm.")
                debug_print(f"Failed to set Reference Level.", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                console_print_func(f"✅ Reference Level set to {ref_level_dbm:.1f} dBm.")

        # --- Apply Frequency Shift ---
        debug_print(f"Querying current Frequency Shift for comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_freq_shift_str = app_instance.inst.query(":INPut:RFSense:FREQuency:SHIFt?", console_print_func)
        current_freq_shift_hz = float(current_freq_shift_str) if current_freq_shift_str else None

        if current_freq_shift_hz is not None and abs(current_freq_shift_hz - freq_shift_hz) < 1:
            debug_print(f"Frequency Shift already at {freq_shift_hz} Hz. Skipping command.", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"ℹ️ Info: Frequency Shift already at {freq_shift_hz:.0f} Hz.")
        else:
            debug_print(f"Setting Frequency Shift to {freq_shift_hz} Hz...", file=current_file, function=current_function, console_print_func=console_print_func)
            if not app_instance.inst.write(f":INPut:RFSense:FREQuency:SHIFt {freq_shift_hz}"):
                success = False
                console_print_func(f"❌ Failed to set Frequency Shift to {freq_shift_hz:.0f} Hz.")
                debug_print(f"Failed to set Frequency Shift.", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                console_print_func(f"✅ Frequency Shift set to {freq_shift_hz:.0f} Hz.")

        # --- Apply Max Hold ---
        debug_print(f"Querying current Trace Type for Max Hold comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_trace_type_str = app_instance.inst.query(":DISPlay:WINDow:TRACe:TYPE?", console_print_func)
        
        desired_trace_type_command = ":DISPlay:WINDow:TRACe:TYPE MAXHold" if max_hold_enabled else ":DISPlay:WINDow:TRACe:TYPE NORM"
        desired_trace_type_status = "MAXH" if max_hold_enabled else "NORM"

        if current_trace_type_str and desired_trace_type_status in current_trace_type_str.upper():
            debug_print(f"Max Hold state already set to {'Enabled' if max_hold_enabled else 'Disabled'}. Skipping command.", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"ℹ️ Info: Max Hold already {'Enabled' if max_hold_enabled else 'Disabled'}.")
        else:
            debug_print(f"Setting Max Hold state to {'Enabled' if max_hold_enabled else 'Disabled'}...", file=current_file, function=current_function, console_print_func=console_print_func)
            if not app_instance.inst.write(desired_trace_type_command):
                success = False
                console_print_func(f"❌ Failed to set Max Hold to {'Enabled' if max_hold_enabled else 'Disabled'}.")
                debug_print(f"Failed to set Max Hold.", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                console_print_func(f"✅ Max Hold set to {'Enabled' if max_hold_enabled else 'Disabled'}.")

        # --- Apply High Sensitivity / Preamp ---
        # High sensitivity typically means Attenuation OFF and Preamplifier ON
        debug_print(f"Querying current Attenuation Auto state for High Sensitivity comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_atten_auto_str = app_instance.inst.query(":INPut:ATTenuation:AUTO?", console_print_func)
        debug_print(f"Querying current Preamplifier state for High Sensitivity comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_preamp_state_str = app_instance.inst.query(":INPut:GAIN:STATe?", console_print_func)

        current_high_sensitivity_state = (current_atten_auto_str and "OFF" in current_atten_auto_str.upper()) and \
                                         (current_preamp_state_str and "ON" in current_preamp_state_str.upper())

        if high_sensitivity == current_high_sensitivity_state:
            debug_print(f"High Sensitivity already set to {'Enabled' if high_sensitivity else 'Disabled'}. Skipping commands.", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"ℹ️ Info: High Sensitivity already {'Enabled' if high_sensitivity else 'Disabled'}.")
        else:
            debug_print(f"Setting High Sensitivity to {'Enabled' if high_sensitivity else 'Disabled'}...", file=current_file, function=current_function, console_print_func=console_print_func)
            if high_sensitivity:
                if not app_instance.inst.write(":INPut:ATTenuation:AUTO OFF"): success = False
                if not app_instance.inst.write(":INPut:ATTenuation 0"): success = False # Set attenuation to 0 dB
                if not app_instance.inst.write(":INPut:GAIN:STATe ON"): success = False # Turn on preamplifier
            else:
                if not app_instance.inst.write(":INPut:ATTenuation:AUTO ON"): success = False
                if not app_instance.inst.write(":INPut:GAIN:STATe OFF"): success = False
            
            if success:
                console_print_func(f"✅ High Sensitivity set to {'Enabled' if high_sensitivity else 'Disabled'}.")
            else:
                console_print_func(f"❌ Failed to set High Sensitivity to {'Enabled' if high_sensitivity else 'Disabled'}.")
                debug_print(f"Failed to set High Sensitivity.", file=current_file, function=current_function, console_print_func=console_print_func)


        # --- Apply VBW/RBW ratio (fixed to 1/3 as per frequency_bands.py) ---
        vbw_hz = rbw_hz * VBW_RBW_RATIO
        debug_print(f"Querying current VBW for comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_vbw_str = app_instance.inst.query(":SENSe:BANDwidth:VIDeo?", console_print_func)
        current_vbw_hz = float(current_vbw_str) if current_vbw_str else None

        if current_vbw_hz is not None and abs(current_vbw_hz - vbw_hz) < 1:
            debug_print(f"VBW already at {vbw_hz} Hz. Skipping command.", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"ℹ️ Info: VBW already at {vbw_hz:.0f} Hz.")
        else:
            debug_print(f"Setting VBW to {vbw_hz} Hz...", file=current_file, function=current_function, console_print_func=console_print_func)
            if not app_instance.inst.write(f":SENSe:BANDwidth:VIDeo {vbw_hz}"):
                success = False
                console_print_func(f"❌ Failed to set VBW to {vbw_hz:.0f} Hz.")
                debug_print(f"Failed to set VBW.", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                console_print_func(f"✅ VBW set to {vbw_hz:.0f} Hz.")
        
        if success:
            console_print_func("✅ All applicable settings applied successfully.")
            debug_print("All applicable settings applied successfully.", file=current_file, function=current_function, console_print_func=console_print_func)
            save_config(app_instance) # Save current settings as last used
            app_instance.reset_setting_colors_logic() # Reset colors after successful apply
            return True
        else:
            console_print_func("❌ Failed to apply all settings. Check connection and instrument status.")
            debug_print("Failed to apply all settings.", file=current_file, function=current_function, console_print_func=console_print_func)
            return False

    except pyvisa.errors.VisaIOError as e:
        console_print_func(f"❌ VISA error while applying settings: {e}")
        debug_print(f"VISA Error applying settings: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False
    except Exception as e:
        console_print_func(f"❌ An unexpected error occurred while applying settings: {e}")
        debug_print(f"An unexpected error occurred while applying settings: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

def query_current_instrument_settings_logic(app_instance, console_print_func):
    """
    Queries the current settings from the instrument and updates the GUI.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    console_print_func("\nQuerying current instrument settings...")
    debug_print("Querying current instrument settings...", file=current_file, function=current_function, console_print_func=console_print_func)

    if not app_instance.inst:
        console_print_func("⚠️ Warning: No instrument connected. Cannot query settings.")
        debug_print("No instrument connected. Cannot query settings.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    try:
        center_freq_hz, span_hz, rbw_hz = query_current_instrument_settings(app_instance.inst, MHZ_TO_HZ, console_print_func)
        
        # Update Tkinter variables
        app_instance.center_freq_var.set(f"{center_freq_hz / MHZ_TO_HZ:.3f}")
        app_instance.span_var.set(f"{span_hz / MHZ_TO_HZ:.3f}")
        app_instance.rbw_var.set(f"{rbw_hz / MHZ_TO_HZ:.3f}") # RBW is often displayed in MHz or KHz
        
        # Query and update other settings as needed
        # Example: Reference Level
        ref_level_dbm_str = app_instance.inst.query(":DISPlay:WINDow:TRACe:Y:RLEVel?", console_print_func)
        if ref_level_dbm_str:
            try:
                app_instance.ref_level_var.set(f"{float(ref_level_dbm_str):.1f}")
            except ValueError:
                debug_print(f"Could not convert queried reference level: {ref_level_dbm_str}", file=current_file, function=current_function, console_print_func=console_print_func)

        # Query and update Max Hold state
        trace_type_query = app_instance.inst.query(":DISPlay:WINDow:TRACe:TYPE?", console_print_func)
        if trace_type_query:
            app_instance.max_hold_enabled_var.set("MAXH" in trace_type_query.upper())

        # Query and update High Sensitivity / Preamp state
        atten_auto_query = app_instance.inst.query(":INPut:ATTenuation:AUTO?", console_print_func)
        gain_state_query = app_instance.inst.query(":INPut:GAIN:STATe?", console_print_func)
        if atten_auto_query and gain_state_query:
            # High sensitivity is typically attenuation off and preamp on
            app_instance.high_sensitivity_var.set("OFF" in atten_auto_query.upper() and "ON" in gain_state_query.upper())

        console_print_func("✅ Current instrument settings updated in GUI.")
        debug_print("Current instrument settings updated in GUI.", file=current_file, function=current_function, console_print_func=console_print_func)
        return True
    except pyvisa.errors.VisaIOError as e:
        console_print_func(f"❌ VISA error while querying settings: {e}")
        debug_print(f"VISA Error querying settings: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False
    except Exception as e:
        console_print_func(f"❌ An unexpected error occurred while querying settings: {e}")
        debug_print(f"An unexpected error occurred while querying settings: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False


def load_selected_preset_logic(app_instance, selected_preset_name, console_print_func):
    """
    Loads a selected preset onto the instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    # Use the control_load_selected_preset from utils.instrument_control
    success, center_freq, span, rbw = control_load_selected_preset(
        app_instance.inst, selected_preset_name, console_print_func
    )
    if success:
        # Update GUI settings based on loaded preset
        app_instance.center_freq_var.set(f"{center_freq:.3f}")
        app_instance.span_var.set(f"{span:.3f}")
        app_instance.rbw_var.set(f"{rbw:.3f}")
        console_print_func(f"GUI settings updated from preset: Center Freq={center_freq:.3f} MHz, Span={span:.3f} MHz, RBW={rbw:.3f} MHz")
        debug_print("GUI settings updated from loaded preset.", file=current_file, function=current_function, console_print_func=console_print_func)
        return True
    else:
        console_print_func(f"❌ Failed to load preset '{selected_preset_name}'.")
        debug_print(f"Failed to load preset '{selected_preset_name}'.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

def query_device_presets_logic(app_instance, console_print_func):
    """
    Queries the connected instrument for a list of preset files.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    # Use the control_query_device_presets from utils.instrument_control
    presets = control_query_device_presets(app_instance.inst, console_print_func)
    if presets is not None:
        app_instance.preset_files_tab.display_presets(presets, "device")
        console_print_func(f"✅ Queried {len(presets)} presets from device.")
        debug_print(f"Queried {len(presets)} presets from device.", file=current_file, function=current_function, console_print_func=console_print_func)
        return True
    else:
        console_print_func("❌ Failed to query presets from device.")
        debug_print("Failed to query presets from device.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False
