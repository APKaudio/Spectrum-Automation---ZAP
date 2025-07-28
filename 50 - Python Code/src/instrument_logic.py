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
    # Removed query_current_instrument_settings, as it will be replaced by individual query_safe calls
    debug_print, # Import debug_print
    query_safe # Import query_safe for individual queries
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
    
    raw_resources = list_visa_resources(console_print_func)
    debug_print(f"Raw resources found: {raw_resources}", file=current_file, function=current_function, console_print_func=console_print_func)
    
    # Sanitize each resource name before setting it to app_instance.resource_names
    sanitized_resources = []
    for resource in raw_resources:
        # Apply the same sanitization logic as in connect_instrument_logic
        sanitized_resource = resource.strip().replace("'", "").replace('"', "").rstrip(',').rstrip(')')
        sanitized_resources.append(sanitized_resource)
    debug_print(f"Sanitized resources: {sanitized_resources}", file=current_file, function=current_function, console_print_func=console_print_func)

    app_instance.resource_names.set("") # Clear existing options by setting to empty string
    debug_print("Cleared existing resource names in app_instance.resource_names.", file=current_file, function=current_function, console_print_func=console_print_func)

    if sanitized_resources:
        # Join the list into a single space-separated string for the StringVar
        app_instance.resource_names.set(" ".join(sanitized_resources))
        debug_print(f"Set app_instance.resource_names to: '{app_instance.resource_names.get()}'", file=current_file, function=current_function, console_print_func=console_print_func)

        # Attempt to set the last used resource if it's still available
        last_used_resource = app_instance.config.get('LAST_USED_SETTINGS', 'last_gpib_device', fallback='')
        debug_print(f"Last used resource from config: '{last_used_resource}'", file=current_file, function=current_function, console_print_func=console_print_func)

        if last_used_resource and last_used_resource in sanitized_resources:
            app_instance.selected_resource.set(last_used_resource)
            console_print_func(f"✅ Last used resource '{last_used_resource}' found and selected.")
            debug_print(f"Last used resource '{last_used_resource}' found and selected.", file=current_file, function=current_function, console_print_func=console_print_func)
        else:
            app_instance.selected_resource.set(sanitized_resources[0]) # Select the first resource by default
            console_print_func(f"✅ Resources found. Selected: {sanitized_resources[0]}")
            debug_print(f"Selected first resource by default: {sanitized_resources[0]}", file=current_file, function=current_function, console_print_func=console_print_func)
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

    # The resource_name should already be sanitized by populate_resources_logic,
    # but a final strip is harmless.
    sanitized_resource_name = resource_name.strip()
    
    console_print_func(f"\nAttempting to connect to {sanitized_resource_name}...")
    debug_print(f"Attempting to connect to {sanitized_resource_name}...", file=current_file, function=current_function, console_print_func=console_print_func)

    if sanitized_resource_name == "No Resources Found" or not sanitized_resource_name:
        console_print_func("⚠️ Warning: No valid VISA resource selected or resource name is empty after sanitization.")
        debug_print("No valid VISA resource selected for connection or empty after sanitization.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    try:
        inst = connect_to_instrument(sanitized_resource_name, console_print_func)
        if inst:
            app_instance.inst = inst
            console_print_func(f"✅ Successfully connected to {sanitized_resource_name}")
            debug_print(f"Successfully connected to {sanitized_resource_name}", file=current_file, function=current_function, console_print_func=console_print_func)
            
            # --- Determine Instrument Model ---
            model_match = "UNKNOWN" # Default
            try:
                idn_response = query_safe(inst, "*IDN?", console_print_func) # Use query_safe
                if idn_response:
                    # Example IDN: Agilent Technologies,N9340B,MY48060001,A.01.00
                    parts = idn_response.split(',')
                    if len(parts) > 1:
                        model_match = parts[1].strip()
                        app_instance.instrument_model = model_match # Store model in app_instance
                        console_print_func(f"✅ Detected instrument model: {model_match}")
                        debug_print(f"Detected instrument model: {model_match}", file=current_file, function=current_function, console_print_func=console_print_func)
            except Exception as idn_e:
                console_print_func(f"⚠️ Warning: Could not query instrument IDN: {idn_e}. Assuming UNKNOWN model.")
                debug_print(f"Error querying IDN: {idn_e}", file=current_file, function=current_function, console_print_func=console_print_func)

            # --- Retrieve settings from app_instance Tkinter variables for initialization ---
            # These are the user's desired initial settings, not necessarily the current instrument state.
            # Corrected variable names
            init_ref_level_dbm = _get_float_value(app_instance.reference_level_dbm_var, -40.0, "Reference Level", console_print_func)
            init_high_sensitivity_on = app_instance.high_sensitivity_var.get()
            init_preamp_on = app_instance.preamp_on_var.get()
            init_rbw_config_val = _get_float_value(app_instance.scan_rbw_hz_var, 10000.0, "Scan RBW", console_print_func)
            init_vbw_config_val = init_rbw_config_val * VBW_RBW_RATIO # Derived VBW

            # --- Initialize instrument settings ---
            if initialize_instrument(
                app_instance.inst,
                init_ref_level_dbm,
                init_high_sensitivity_on,
                init_preamp_on,
                init_rbw_config_val,
                init_vbw_config_val,
                model_match, # Pass the detected model
                console_print_func
            ):
                console_print_func("✅ Instrument initialized with default settings.")
                debug_print("Instrument initialized with default settings.", file=current_file, function=current_function, console_print_func=console_print_func)
                
                # Query and display current instrument settings in InstrumentTab
                # This will update the InstrumentTab's local Tkinter variables
                if hasattr(app_instance, 'instrument_tab'):
                    app_instance.instrument_tab._query_settings_display()
                
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
            console_print_func(f"❌ Failed to connect to {sanitized_resource_name}.")
            debug_print(f"Failed to connect to {sanitized_resource_name}.", file=current_file, function=current_function, console_print_func=console_print_func)
            app_instance.update_connection_status(False)
            return False
    except Exception as e:
        console_print_func(f"❌ An error occurred during connection: {e}")
        debug_print(f"Error during connection to {sanitized_resource_name}: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
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
        try:
            control_disconnect_instrument(app_instance.inst, console_print_func)
            app_instance.inst = None
            console_print_func("✅ Instrument disconnected.")
            debug_print("Instrument disconnected.", file=current_file, function=current_function, console_print_func=console_print_func)
        except Exception as e:
            console_print_func(f"❌ An error occurred during disconnection: {e}")
            debug_print(f"Error during disconnection: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
    else:
        console_print_func("ℹ️ Info: No instrument to disconnect.")
        debug_print("No instrument to disconnect.", file=current_file, function=current_function, console_print_func=console_print_func)
    app_instance.update_connection_status(False) # Update GUI status

def apply_settings_logic(app_instance, console_print_func):
    """
    Applies the current settings from the GUI's main application variables to the connected instrument.
    This includes Reference Level, Frequency Shift, Max Hold, High Sensitivity, Preamp, and Scan RBW/VBW.
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
        # Get values from Tkinter variables directly from app_instance
        # Corrected variable names
        ref_level_dbm = _get_float_value(app_instance.reference_level_dbm_var, -40.0, "Reference Level", console_print_func)
        freq_shift_hz = _get_float_value(app_instance.freq_shift_hz_var, 0.0, "Frequency Shift", console_print_func)
        max_hold_enabled = app_instance.maxhold_enabled_var.get() # Corrected variable name
        high_sensitivity = app_instance.high_sensitivity_var.get()
        preamp_on = app_instance.preamp_on_var.get()
        
        # Get the desired RBW from the scan config variable
        rbw_hz_to_apply = _get_float_value(app_instance.scan_rbw_hz_var, 10000.0, "Scan RBW", console_print_func)
        vbw_hz_to_apply = rbw_hz_to_apply * VBW_RBW_RATIO # Derived VBW

        success = True

              
        # --- Apply High Sensitivity / Preamp ---
        # High sensitivity typically means Attenuation OFF and Preamplifier ON
        debug_print(f"Querying current Attenuation Auto state for High Sensitivity comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_atten_auto_str = query_safe(app_instance.inst, ":INPut:ATTenuation:AUTO?", console_print_func) # Query actual state
        debug_print(f"Querying current Preamplifier state for High Sensitivity comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_preamp_state_str = query_safe(app_instance.inst, ":INPut:GAIN:STATe?", console_print_func) # Query actual state

        current_high_sensitivity_state = (current_atten_auto_str and "OFF" in current_atten_auto_str.upper()) and \
                                         (current_preamp_state_str and "ON" in current_preamp_state_str.upper())

        if high_sensitivity == current_high_sensitivity_state:
            debug_print(f"High Sensitivity already set to {'Enabled' if high_sensitivity else 'Disabled'}. Skipping commands.", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"ℹ️ Info: High Sensitivity already {'Enabled' if high_sensitivity else 'Disabled'}.")
        else:
            debug_print(f"Setting High Sensitivity to {'Enabled' if high_sensitivity else 'Disabled'}...", file=current_file, function=current_function, console_print_func=console_print_func)
            if high_sensitivity:
                if not write_safe(app_instance.inst, ":INPut:ATTenuation:AUTO OFF", console_print_func): success = False
                if not write_safe(app_instance.inst, ":INPut:ATTenuation 0", console_print_func): success = False # Set attenuation to 0 dB
                if not write_safe(app_instance.inst, ":INPut:GAIN:STATe ON", console_print_func): success = False # Turn on preamplifier
            else:
                if not write_safe(app_instance.inst, ":INPut:ATTenuation:AUTO ON", console_print_func): success = False
                if not write_safe(app_instance.inst, ":INPut:GAIN:STATe OFF", console_print_func): success = False
            
            if success:
                console_print_func(f"✅ High Sensitivity set to {'Enabled' if high_sensitivity else 'Disabled'}.")
            else:
                console_print_func(f"❌ Failed to set High Sensitivity to {'Enabled' if high_sensitivity else 'Disabled'}.")
                debug_print(f"Failed to set High Sensitivity.", file=current_file, function=current_function, console_print_func=console_print_func)


        # --- Apply RBW (from scan config) ---
        debug_print(f"Querying current RBW for comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_rbw_str = query_safe(app_instance.inst, ":SENSe:BANDwidth:RESolution?", console_print_func)
        current_rbw_hz = float(current_rbw_str) if current_rbw_str else None

        if current_rbw_hz is not None and abs(current_rbw_hz - rbw_hz_to_apply) < 1:
            debug_print(f"RBW already at {rbw_hz_to_apply} Hz. Skipping command.", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"ℹ️ Info: RBW already at {rbw_hz_to_apply:.0f} Hz.")
        else:
            debug_print(f"Setting RBW to {rbw_hz_to_apply} Hz...", file=current_file, function=current_function, console_print_func=console_print_func)
            if not write_safe(app_instance.inst, f":SENSe:BANDwidth:RESolution {rbw_hz_to_apply}", console_print_func):
                success = False
                console_print_func(f"❌ Failed to set RBW to {rbw_hz_to_apply:.0f} Hz.")
                debug_print(f"Failed to set RBW.", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                console_print_func(f"✅ RBW set to {rbw_hz_to_apply:.0f} Hz.")

        # --- Apply VBW (derived from RBW) ---
        debug_print(f"Querying current VBW for comparison...", file=current_file, function=current_function, console_print_func=console_print_func)
        current_vbw_str = query_safe(app_instance.inst, ":SENSe:BANDwidth:VIDeo?", console_print_func)
        current_vbw_hz = float(current_vbw_str) if current_vbw_str else None

        if current_vbw_hz is not None and abs(current_vbw_hz - vbw_hz_to_apply) < 1:
            debug_print(f"VBW already at {vbw_hz_to_apply} Hz. Skipping command.", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"ℹ️ Info: VBW already at {vbw_hz_to_apply:.0f} Hz.")
        else:
            debug_print(f"Setting VBW to {vbw_hz_to_apply} Hz...", file=current_file, function=current_function, console_print_func=console_print_func)
            if not write_safe(app_instance.inst, f":SENSe:BANDwidth:VIDeo {vbw_hz_to_apply}", console_print_func):
                success = False
                console_print_func(f"❌ Failed to set VBW to {vbw_hz_to_apply:.0f} Hz.")
                debug_print(f"Failed to set VBW.", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                console_print_func(f"✅ VBW set to {vbw_hz_to_apply:.0f} Hz.")
        
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
        # Use query_safe for all instrument queries
        center_freq_str = query_safe(app_instance.inst, ":SENSe:FREQuency:CENTer?", console_print_func)
        span_str = query_safe(app_instance.inst, ":SENSe:FREQuency:SPAN?", console_print_func)
        rbw_str = query_safe(app_instance.inst, ":SENSe:BANDwidth:RESolution?", console_print_func)
        
        center_freq_hz = float(center_freq_str) if center_freq_str else 0.0
        span_hz = float(span_str) if span_str else 0.0
        rbw_hz = float(rbw_str) if rbw_str else 0.0

        # Query attenuation and gain states for high sensitivity display
        atten_auto_query = query_safe(app_instance.inst, ":INPut:ATTenuation:AUTO?", console_print_func)
        gain_state_query = query_safe(app_instance.inst, ":INPut:GAIN:STATe?", console_print_func)

        # Update Tkinter variables in the InstrumentTab
        if hasattr(app_instance, 'instrument_tab'):
            app_instance.instrument_tab.current_center_freq_var.set(f"{center_freq_hz / MHZ_TO_HZ:.3f}")
            app_instance.instrument_tab.current_span_var.set(f"{span_hz / MHZ_TO_HZ:.3f}")
            app_instance.instrument_tab.current_rbw_var.set(f"{rbw_hz:.0f}") # RBW in Hz, displayed as integer
            
            if atten_auto_query and gain_state_query:
                # High sensitivity is typically attenuation off and preamp on
                app_instance.instrument_tab.current_high_sensitivity_var.set("Enabled" if ("OFF" in atten_auto_query.upper() and "ON" in gain_state_query.upper()) else "Disabled")

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
    from utils.preset_utils import load_selected_preset as control_load_selected_preset # Import here
    success, center_freq, span, rbw = control_load_selected_preset(
        app_instance.inst, selected_preset_name, console_print_func
    )
    if success:
        # Update GUI settings based on loaded preset
        # These variables are now in InstrumentTab's local vars
        if hasattr(app_instance, 'instrument_tab'):
            app_instance.instrument_tab.current_center_freq_var.set(f"{center_freq:.3f}")
            app_instance.instrument_tab.current_span_var.set(f"{span:.3f}")
            app_instance.instrument_tab.current_rbw_var.set(f"{rbw:.0f}") # RBW in Hz, displayed as integer

            # Also update the main app's variables that are tied to settings if they exist
            # This ensures consistency if these variables are used elsewhere (e.g., for saving config)
            # Corrected variable names
            app_instance.reference_level_dbm_var.set(app_instance.instrument_tab.current_ref_level_var.get())
            app_instance.freq_shift_hz_var.set(app_instance.instrument_tab.current_freq_shift_var.get())
            # For max hold and high sensitivity, you might need to query the instrument again
            # or infer from the preset if it includes these states.
            # For simplicity, we'll rely on the instrument_tab's query_settings_display to update these.
            app_instance.instrument_tab._query_settings_display()


        console_print_func(f"GUI settings updated from preset: Center Freq={center_freq:.3f} MHz, Span={span:.3f} MHz, RBW={rbw:.0f} Hz")
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
    from utils.preset_utils import query_device_presets as control_query_device_presets # Import here
    presets = control_query_device_presets(app_instance.inst, console_print_func)
    if presets is not None:
        if hasattr(app_instance, 'preset_files_tab'):
            app_instance.preset_files_tab.populate_preset_buttons(presets, "device")
        console_print_func(f"✅ Queried {len(presets)} presets from device.")
        debug_print(f"Queried {len(presets)} presets from device.", file=current_file, function=current_function, console_print_func=console_print_func)
        return presets # Return the list of presets
    else:
        console_print_func("❌ Failed to query presets from device.")
        debug_print("Failed to query presets from device.", file=current_file, function=current_function, console_print_func=console_print_func)
        return None # Return None on failure

