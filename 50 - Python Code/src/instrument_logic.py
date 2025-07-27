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
            return default_value
        return float(val_str)
    except ValueError:
        debug_print(f"Error: Could not convert '{val_str}' for '{setting_name}' to float. Using default: {default_value}", file=current_file, function=current_function)
        return default_value
    except TclError: # Catch TclError for cases where tk_var might not be initialized yet
        debug_print(f"Error: TclError for '{setting_name}'. Using default: {default_value}", file=current_file, function=current_function)
        return default_value

def _get_int_value(tk_var, default_value, setting_name):
    """
    Safely retrieves an integer value from a Tkinter StringVar.
    If the string is empty or cannot be converted, returns a default value.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    try:
        val_str = tk_var.get()
        if not val_str:
            debug_print(f"Warning: Tkinter variable for '{setting_name}' is empty. Using default: {default_value}", file=current_file, function=current_function)
            return default_value
        return int(float(val_str)) # Convert to float first to handle "10000.0"
    except ValueError:
        debug_print(f"Error: Could not convert '{val_str}' for '{setting_name}' to int. Using default: {default_value}", file=current_file, function=current_function)
        return default_value
    except TclError:
        debug_print(f"Error: TclError for '{setting_name}'. Using default: {default_value}", file=current_file, function=current_function)
        return default_value

def _get_bool_value(tk_var, default_value, setting_name):
    """
    Safely retrieves a boolean value from a Tkinter BooleanVar.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    try:
        return tk_var.get()
    except TclError:
        debug_print(f"Error: TclError for '{setting_name}'. Using default: {default_value}", file=current_file, function=current_function)
        return default_value


def populate_resources_logic(app_instance, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Populates the VISA resource dropdown list with available instruments.
    Enables the connect button if resources are found.
    """
    debug_print("Populating VISA resources...", file=file, function=function)
    if app_instance.rm is None:
        print("❌ PyVISA Resource Manager not initialized. Cannot list resources.")
        app_instance.instrument_list = []
        app_instance.resource_var.set("No resources found")
        app_instance.connect_button.config(state=tk.DISABLED) # Ensure disabled if no RM
        return

    try:
        resources = app_instance.rm.list_resources()
        debug_print(f"Found VISA resources: {resources}", file=file, function=function)
        
        app_instance.instrument_list = list(resources)
        
        if app_instance.instrument_list:
            app_instance.resource_var.set(app_instance.instrument_list[0]) # Set default to first resource
            print(f"✅ Found {len(app_instance.instrument_list)} VISA resources.")
            debug_print(f"Available VISA resources: {app_instance.instrument_list}", file=file, function=function)
            
            # --- FIX: Enable the connect button here ---
            app_instance.connect_button.config(state=tk.NORMAL) 
            app_instance._stop_connect_button_blink() # Stop blinking if resources are found and button enabled
            # --- END FIX ---

            # Update the OptionMenu with new resources
            menu = app_instance.resource_dropdown["menu"]
            menu.delete(0, "end")
            for resource in app_instance.instrument_list:
                menu.add_command(label=resource, command=tk._setit(app_instance.resource_var, resource))
            
            # If a last_gpib_device was loaded, try to select it
            last_gpib_device = app_instance.gpib_device_var.get()
            if last_gpib_device and last_gpib_device in app_instance.instrument_list:
                app_instance.resource_var.set(last_gpib_device)
                debug_print(f"Auto-selected last used device: {last_gpib_device}", file=file, function=function)
            else:
                debug_print("No last used device to auto-select or not found.", file=file, function=function)

        else:
            app_instance.resource_var.set("No resources found")
            print("🚫 No VISA resources found.")
            app_instance.connect_button.config(state=tk.DISABLED) # Disable if no resources
            debug_print("No VISA resources found. Connect button disabled.", file=file, function=function)
            app_instance._start_connect_button_blink() # Start blinking if no resources are found

    except Exception as e:
        print(f"❌ Error listing VISA resources: {e}")
        messagebox.showerror("VISA Error", f"Failed to list VISA resources: {e}")
        app_instance.instrument_list = []
        app_instance.resource_var.set("Error listing resources")
        app_instance.connect_button.config(state=tk.DISABLED) # Disable on error
        debug_print(f"Error listing VISA resources: {e}", file=file, function=function)
        app_instance._start_connect_button_blink() # Start blinking on error


def connect_instrument_logic(app_instance, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Attempts to connect to the selected VISA instrument.
    Updates GUI state based on connection success or failure.
    """
    debug_print("Attempting to connect to instrument...", file=file, function=function)
    selected_resource = app_instance.resource_var.get()
    if not selected_resource or selected_resource == "No resources found" or "Error" in selected_resource:
        messagebox.showwarning("Connection Error", "Please select a valid VISA resource.")
        debug_print("No valid VISA resource selected for connection.", file=file, function=function)
        return

    print(f"Attempting to connect to: {selected_resource}")
    app_instance.connect_button.config(state=tk.DISABLED) # Disable button while connecting
    app_instance._stop_connect_button_blink() # Stop blinking when attempting connection

    try:
        # Pass the app_instance.rm to the connect_to_instrument function
        inst, model = connect_to_instrument(app_instance.rm, selected_resource)
        if inst:
            app_instance.inst = inst
            app_instance.instrument_model = model
            app_instance.gpib_device_var.set(selected_resource) # Set GPIB device var
            print(f"✅ Successfully connected to {selected_resource} (Model: {model}).")
            
            # --- FIX: Retrieve and pass missing arguments to initialize_instrument ---
            high_sensitivity_on = _get_bool_value(app_instance.desired_high_sensitivity_var, True, "desired_high_sensitivity_var")
            preamp_on = _get_bool_value(app_instance.desired_preamp_on_var, True, "desired_preamp_on_var")
            rbw_config_val = _get_int_value(app_instance.desired_rbw_var, 10000, "desired_rbw_var")
            # Calculate vbw_config_val based on rbw_config_val and VBW_RBW_RATIO
            vbw_config_val = rbw_config_val * VBW_RBW_RATIO
            model_match = app_instance.instrument_model # This is already available
            # Get reference level from app_instance
            ref_level_dbm = _get_float_value(app_instance.desired_reference_level_var, -40, "desired_reference_level_var")

            if initialize_instrument(
                app_instance.inst,
                ref_level_dbm, # Added ref_level_dbm
                high_sensitivity_on,
                preamp_on,
                rbw_config_val,
                vbw_config_val,
                model_match # Pass the model to initialize_instrument
            ):
            # --- END FIX ---
                print("✅ Instrument initialized successfully.")
                query_current_instrument_settings(app_instance.inst, MHZ_TO_HZ) # Query and display settings
                app_instance.disconnect_button.config(state=tk.NORMAL)
                app_instance.start_scan_button.config(state=tk.NORMAL)
                app_instance.apply_button.config(state=tk.NORMAL)
                app_instance.plot_button.config(state=tk.NORMAL) # Enable plot button on connect
                
                # Enable query presets button if not N9340B
                if app_instance.instrument_model != "N9340B":
                    if hasattr(app_instance, 'preset_files_tab') and hasattr(app_instance.preset_files_tab, 'query_presets_button'):
                        app_instance.preset_files_tab.query_presets_button.config(state=tk.NORMAL)
                
                save_config(app_instance) # Save last connected device
            else:
                print("❌ Failed to initialize instrument after connection.")
                messagebox.showerror("Connection Error", "Failed to initialize instrument. Disconnecting.")
                control_disconnect_instrument(app_instance.inst) # Disconnect if initialization fails
                app_instance._reset_gui_on_disconnect_or_error() # Reset GUI state
        else:
            print("❌ Failed to connect to instrument.")
            messagebox.showerror("Connection Error", "Could not establish connection to the instrument.")
            app_instance._reset_gui_on_disconnect_or_error() # Reset GUI state
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA I/O Error: {e}")
        messagebox.showerror("VISA Error", f"Failed to connect to instrument: {e}")
        app_instance._reset_gui_on_disconnect_or_error() # Reset GUI state
    except Exception as e:
        print(f"❌ An unexpected error occurred during connection: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        app_instance._reset_gui_on_disconnect_or_error() # Reset GUI state


def disconnect_instrument_logic(app_instance, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Disconnects from the currently connected instrument.
    Resets GUI state to reflect disconnection.
    """
    debug_print("Attempting to disconnect instrument...", file=file, function=function)
    if app_instance.inst:
        print(f"Disconnecting from {app_instance.gpib_device_var.get()}...")
        if control_disconnect_instrument(app_instance.inst):
            print("✅ Instrument disconnected.")
        else:
            print("⚠️ Failed to gracefully disconnect instrument. It might still be connected.")
        app_instance._reset_gui_on_disconnect_or_error() # Reset GUI state
    else:
        messagebox.showwarning("Not Connected", "No instrument is currently connected.")
        debug_print("Attempted to disconnect, but no instrument was connected.", file=file, function=function)


def apply_settings_to_device_logic(app_instance, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Applies the current settings from the GUI to the connected instrument.
    """
    debug_print("Applying settings to device...", file=file, function=function)
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
        debug_print("Cannot apply settings: Instrument not connected.", file=file, function=function)
        return False

    print("\n--- Applying Settings to Instrument ---")
    
    # Get values from Tkinter variables, using helper functions for safety
    rbw_hz = _get_int_value(app_instance.desired_rbw_var, 10000, "desired_rbw_var")
    max_hold_time_seconds = _get_int_value(app_instance.desired_max_hold_time_var, 3, "desired_max_hold_time_var")
    cycle_wait_time_seconds = _get_float_value(app_instance.desired_cycle_wait_time_var, 0.5, "desired_cycle_wait_time_var")
    reference_level_dbm = _get_float_value(app_instance.desired_reference_level_var, -40, "desired_reference_level_var")
    freq_shift_hz = _get_float_value(app_instance.desired_freq_shift_var, 0, "desired_freq_shift_var")
    maxhold_enabled = _get_bool_value(app_instance.desired_maxhold_enabled_var, True, "desired_maxhold_enabled_var")
    high_sensitivity = _get_bool_value(app_instance.desired_high_sensitivity_var, True, "desired_high_sensitivity_var")
    preamp_on = _get_bool_value(app_instance.desired_preamp_on_var, True, "desired_preamp_on_var")
    scan_rbw_segmentation = _get_float_value(app_instance.desired_scan_rbw_segmentation_var, 1000000.0, "desired_scan_rbw_segmentation_var")
    default_focus_width = _get_float_value(app_instance.desired_default_focus_width_var, 10000.0, "desired_default_focus_width_var")

    try:
        inst = app_instance.inst

        # Set RBW
        if not inst.write(f":SENSe:BANDwidth:RESolution {rbw_hz}"): return False
        debug_print(f"Sent: :SENSe:BANDwidth:RESolution {rbw_hz}", file=file, function=function)

        # Set VBW (calculated from RBW)
        vbw_hz = rbw_hz * app_instance.VBW_RBW_RATIO
        if not inst.write(f":SENSe:BANDwidth:VIDeo {vbw_hz}"): return False
        debug_print(f"Sent: :SENSe:BANDwidth:VIDeo {vbw_hz}", file=file, function=function)

        # Set Reference Level
        if not inst.write(f":DISPlay:WINDow:TRACe:Y:RLEVel {reference_level_dbm}DBM"): return False
        debug_print(f"Sent: :DISPlay:WINDow:TRACe:Y:RLEVel {reference_level_dbm}DBM", file=file, function=function)

        # Set Max Hold (if enabled)
        if maxhold_enabled:
            if not inst.write(":DISPlay:WINDow:TRACe:TYPE MAXH"): return False
            debug_print("Sent: :DISPlay:WINDow:TRACe:TYPE MAXH", file=file, function=function)
        else:
            if not inst.write(":DISPlay:WINDow:TRACe:TYPE NORM"): return False # Or CLEARW
            debug_print("Sent: :DISPlay:WINDow:TRACe:TYPE NORM", file=file, function=function)

        # Set Preamp (if applicable and enabled)
        if app_instance.instrument_model != "N9340B": # N9340B does not have a controllable preamp
            preamp_state = "ON" if preamp_on else "OFF"
            if not inst.write(f":SENSe:POWer:RF:GAIN:STATe {preamp_state}"): return False
            debug_print(f"Sent: :SENSe:POWer:RF:GAIN:STATe {preamp_state}", file=file, function=function)
        else:
            debug_print("Preamp setting skipped for N9340B model.", file=file, function=function)


        # High Sensitivity (specific to N9342CN, or other models that support it)
        if app_instance.instrument_model == "N9342CN":
            high_sens_state = "ON" if high_sensitivity else "OFF"
            if not inst.write(f":SENSe:POWer:RF:HSENse {high_sens_state}"): return False
            debug_print(f"Sent: :SENSe:POWer:RF:HSENse {high_sens_state}", file=file, function=function)
        else:
            debug_print("High Sensitivity setting skipped for non-N9342CN model.", file=file, function=function)

        # Query and display current instrument settings after applying
        print("\n--- Current Instrument Settings After Apply ---")
        query_current_instrument_settings(inst, MHZ_TO_HZ)
        print("---------------------------------------------")
        
        print("✅ Settings applied successfully.")
        messagebox.showinfo("Settings Applied", "Instrument settings applied successfully!")
        app_instance.reset_setting_colors_logic() # Reset colors after successful application
        return True

    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error applying settings: {e}")
        messagebox.showerror("VISA Error", f"Failed to apply settings: {e}")
        debug_print(f"VISA Error applying settings: {e}", file=file, function=function)
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred while applying settings: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred while applying settings: {e}")
        debug_print(f"An unexpected error occurred while applying settings: {e}", file=file, function=function)
        return False


def load_selected_preset_logic(app_instance, selected_preset_name, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Loads a selected preset onto the connected instrument.
    """
    debug_print(f"Loading selected preset: {selected_preset_name}", file=file, function=function)
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
        debug_print("Cannot load preset: Instrument not connected.", file=file, function=function)
        return False

    # Call the utility function from instrument_control
    success = control_load_selected_preset(app_instance.inst, selected_preset_name)
    if success:
        print(f"✅ Preset '{selected_preset_name}' loaded successfully.")
        # After loading a preset, it's good practice to re-query settings
        query_current_instrument_settings(app_instance.inst, MHZ_TO_HZ)
    else:
        print(f"❌ Failed to load preset '{selected_preset_name}'.")
        # Error message already handled by control_load_selected_preset

    return success


def query_device_presets_logic(app_instance, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Queries the connected instrument for available presets and updates the GUI.
    """
    debug_print("Querying device presets...", file=file, function=function)
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
        debug_print("Cannot query presets: Instrument not connected.", file=file, function=function)
        return

    if app_instance.instrument_model == "N9340B":
        messagebox.showinfo("Preset Feature", "The N9340B model does not support querying device presets via SCPI.")
        print("ℹ️ N9340B does not support querying device presets.")
        # Clear any existing preset buttons if the model doesn't support it
        if hasattr(app_instance, 'preset_files_tab'):
            app_instance.preset_files_tab.clear_preset_buttons()
        return

    print("Querying device presets. This may take a moment...")
    try:
        presets = control_query_device_presets(app_instance.inst)
        if presets is not None:
            print(f"✅ Found {len(presets)} presets on device.")
            debug_print(f"Device presets found: {presets}", file=file, function=function)
            # Update the PresetFilesTab with the new list of presets
            if hasattr(app_instance, 'preset_files_tab'):
                app_instance.preset_files_tab.populate_preset_buttons(presets)
            else:
                debug_print("PresetFilesTab not found on app_instance.", file=file, function=function)
        else:
            print("🚫 No presets found or error during query.")
            messagebox.showwarning("Preset Query", "No presets found on the device or an error occurred during query.")
            # Clear any existing preset buttons
            if hasattr(app_instance, 'preset_files_tab'):
                app_instance.preset_files_tab.clear_preset_buttons()

    except Exception as e:
        print(f"❌ An error occurred while querying presets: {e}")
        messagebox.showerror("Preset Query Error", f"An unexpected error occurred while querying presets: {e}")
        debug_print(f"Error querying presets: {e}", file=file, function=function)
        # Clear any existing preset buttons on error
        if hasattr(app_instance, 'preset_files_tab'):
            app_instance.preset_files_tab.clear_preset_buttons()


def set_focus_frequency_logic(app_instance, frequency_hz, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Sets the instrument's center frequency to a specified value (in Hz).
    """
    debug_print(f"Setting focus frequency to {frequency_hz} Hz...", file=file, function=function)
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
        debug_print("Cannot set focus frequency: Instrument not connected.", file=file, function=function)
        return False

    try:
        # Set center frequency
        if not app_instance.inst.write(f":SENSe:FREQuency:CENTer {frequency_hz}"): return False
        debug_print(f"Sent: :SENSe:FREQuency:CENTer {frequency_hz}", file=file, function=function)

        # Set span around the focus frequency
        focus_width = _get_float_value(app_instance.desired_default_focus_width_var, 10000000.0, "desired_default_focus_width_var")
        if not app_instance.inst.write(f":SENSe:FREQuency:SPAN {focus_width}"): return False
        debug_print(f"Sent: :SENSe:FREQuency:SPAN {focus_width}", file=file, function=function)

        print(f"✅ Instrument center frequency set to {frequency_hz / MHZ_TO_HZ:.3f} MHz with span {focus_width} Hz.")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error while setting focus frequency: {e}")
        messagebox.showerror("VISA Error", f"Failed to set instrument focus frequency: {e}")
        debug_print(f"VISA Error setting focus frequency: {e}", file=file, function=function)
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred while setting focus frequency: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred while setting focus frequency: {e}")
        debug_print(f"An unexpected error occurred while setting focus frequency: {e}", file=file, function=function)
        return False


def set_marker_and_trace_modes_logic(app_instance, marker_frequency_hz, marker_name, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Sets a marker at a specified frequency and configures trace modes.
    """
    debug_print(f"Setting marker '{marker_name}' at {marker_frequency_hz} Hz and trace modes...", file=file, function=function)
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
        debug_print("Cannot set marker: Instrument not connected.", file=file, function=function)
        return False

    try:
        # Activate Marker 1
        if not app_instance.inst.write(":CALCulate:MARKer1:STATe ON"): return False
        debug_print("Sent: :CALCulate:MARKer1:STATe ON", file=file, function=function)

        # Set Marker 1 frequency
        if not app_instance.inst.write(f":CALCulate:MARKer1:X {marker_frequency_hz}"): return False
        debug_print(f"Sent: :CALCulate:MARKer1:X {marker_frequency_hz}", file=file, function=function)

        # Enable Marker Peak Search (optional, but common for markers)
        if not app_instance.inst.write(":CALCulate:MARKer1:MAXimum:PEAK"): return False
        debug_print("Sent: :CALCulate:MARKer1:MAXimum:PEAK", file=file, function=function)

        # Set trace type to Normal (or other desired mode like Clear Write)
        # This might be needed if the instrument was in Max Hold or Min Hold
        if not app_instance.inst.write(":DISPlay:WINDow:TRACe:TYPE NORM"): return False
        debug_print("Sent: :DISPlay:WINDow:TRACe:TYPE NORM", file=file, function=function)

        print(f"✅ Marker '{marker_name}' set at {marker_frequency_hz / MHZ_TO_HZ:.3f} MHz. Trace mode set to Normal.")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error while setting marker/trace modes: {e}")
        messagebox.showerror("VISA Error", f"Failed to set instrument marker or trace modes: {e}")
        debug_print(f"VISA Error setting marker/trace modes: {e}", file=file, function=function)
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred while setting marker/trace modes: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred while setting marker/trace modes: {e}")
        debug_print(f"An unexpected error occurred while setting marker/trace modes: {e}", file=file, function=function)
        return False
