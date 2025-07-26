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
    query_device_presets as control_query_device_presets,
    load_selected_preset as control_load_selected_preset,
    debug_print # Import debug_print
)
from utils.frequency_bands import MHZ_TO_HZ
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

def populate_resources_logic(app_instance, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Populates the VISA resource dropdown with available instruments.
    """
    debug_print("Populating VISA resources...", file=file, function=function)
    if app_instance.rm:
        try:
            app_instance.instrument_list = list_visa_resources(app_instance.rm)
            if app_instance.instrument_list:
                # This line will be overridden by main_app.py if last_gpib_device is found
                app_instance.resource_var.set(app_instance.instrument_list[0]) 
                print(f"✅ Found {len(app_instance.instrument_list)} VISA resources.")
            else:
                app_instance.resource_var.set("No resources found")
                print("🚫 No VISA resources found.")
            # Update the OptionMenu with new list
            menu = app_instance.resource_dropdown["menu"]
            menu.delete(0, "end")
            for resource in app_instance.instrument_list:
                menu.add_command(label=resource, command=tk._setit(app_instance.resource_var, resource))
            app_instance.connect_button.config(state=tk.NORMAL if app_instance.instrument_list else tk.DISABLED)
        except Exception as e:
            print(f"❌ Error listing VISA resources: {e}")
            messagebox.showerror("VISA Error", f"Failed to list VISA resources: {e}")
            app_instance.resource_var.set("Error listing resources")
            app_instance.instrument_list = []
            app_instance.connect_button.config(state=tk.DISABLED)
    else:
        print("🚫 PyVISA Resource Manager not initialized. Cannot list resources.")
        app_instance.resource_var.set("RM not initialized")
        app_instance.connect_button.config(state=tk.DISABLED)


def connect_instrument_logic(app_instance, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Connects to the selected instrument and initializes it.
    """
    debug_print("Attempting to connect to instrument...", file=file, function=function)
    selected_resource = app_instance.resource_var.get()
    if not selected_resource or selected_resource == "No resources found" or selected_resource == "RM not initialized":
        messagebox.showwarning("No Resource Selected", "Please select a valid VISA resource.")
        return

    try:
        # UNPACK THE TUPLE HERE!
        instrument_obj, instrument_model_str = connect_to_instrument(app_instance.rm, selected_resource)
        
        if instrument_obj: # Check if connection was successful
            app_instance.inst = instrument_obj # Store the actual instrument object
            app_instance.instrument_model = instrument_model_str # Store the model string

            # Retrieve values for initialize_instrument from app_instance's Tkinter variables
            ref_level_dbm = _get_float_value(app_instance.desired_reference_level_var, -40.0, "Reference Level")
            high_sensitivity_on = app_instance.desired_high_sensitivity_var.get()
            preamp_on = app_instance.desired_preamp_on_var.get()
            rbw_config_val = _get_float_value(app_instance.desired_rbw_var, 10000.0, "RBW")
            vbw_config_val = rbw_config_val / app_instance.VBW_RBW_RATIO # Calculate VBW based on RBW
            
            # Pass the actual instrument object to initialize_instrument
            initialization_successful = initialize_instrument(
                app_instance.inst, # Pass the actual instrument object here
                ref_level_dbm=ref_level_dbm,
                high_sensitivity_on=high_sensitivity_on,
                preamp_on=preamp_on,
                rbw_config_val=rbw_config_val,
                vbw_config_val=vbw_config_val,
                model_match=app_instance.instrument_model # Pass current model, or it will be updated by initialize_instrument
            )

            if initialization_successful: # Check if initialization was successful
                print(f"✅ Successfully connected to {app_instance.instrument_model} at {selected_resource}")
                app_instance.gpib_device_var.set(selected_resource) # Display connected device
                app_instance.connect_button.config(state=tk.DISABLED)
                app_instance.disconnect_button.config(state=tk.NORMAL)
                app_instance.start_scan_button.config(state=tk.NORMAL)
                app_instance.apply_button.config(state=tk.NORMAL)
                app_instance.query_presets_button.config(state=tk.NORMAL)
                app_instance.plot_button.config(state=tk.NORMAL) # Enable plot button on connect
                app_instance._stop_connect_button_blink() # Stop blinking on successful connection
                save_config(app_instance) # Save the last connected device
                
                # Apply initial settings from GUI to the instrument after connection
                apply_settings_to_device_logic(app_instance)

            else:
                print("❌ Failed to initialize instrument after connection.")
                messagebox.showerror("Initialization Error", "Failed to initialize instrument after connection.")
                # Ensure disconnect is called with the actual instrument object
                control_disconnect_instrument(app_instance.inst)
                app_instance.inst = None # Clear the instrument reference
                app_instance._reset_gui_on_disconnect_or_error()
        else:
            print("❌ Failed to connect to instrument.")
            messagebox.showerror("Connection Error", "Failed to connect to instrument.")
            app_instance._reset_gui_on_disconnect_or_error()
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error during connection: {e}")
        messagebox.showerror("VISA Error", f"Failed to connect to instrument: {e}")
        app_instance._reset_gui_on_disconnect_or_error()
    except Exception as e:
        print(f"❌ An unexpected error occurred during connection: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        app_instance._reset_gui_on_disconnect_or_error()


def disconnect_instrument_logic(app_instance, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Disconnects from the current instrument.
    """
    debug_print("Attempting to disconnect from instrument...", file=file, function=function)
    if app_instance.inst:
        try:
            control_disconnect_instrument(app_instance.inst)
            app_instance.inst = None
            app_instance.instrument_model = "Unknown"
            app_instance.gpib_device_var.set("") # Clear displayed device
            print("✅ Instrument disconnected.")
            app_instance._reset_gui_on_disconnect_or_error()
        except Exception as e:
            print(f"❌ Error during disconnection: {e}")
            messagebox.showerror("Disconnection Error", f"An error occurred during disconnection: {e}")
            app_instance._reset_gui_on_disconnect_or_error()
    else:
        print("ℹ️ No instrument to disconnect.")
        messagebox.showinfo("Info", "No instrument is currently connected.")


def apply_settings_to_device_logic(app_instance, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Applies the current GUI settings to the connected instrument.
    """
    debug_print("Applying settings to device...", file=file, function=function)
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
        return

    try:
        # Retrieve values, using _get_float_value for numerical settings
        rbw_hz = _get_float_value(app_instance.desired_rbw_var, 10000.0, "RBW")
        ref_level_dbm = _get_float_value(app_instance.desired_reference_level_var, -40.0, "Reference Level")
        freq_shift_hz = _get_float_value(app_instance.desired_freq_shift_var, 0.0, "Frequency Shift")
        maxhold_enabled = app_instance.desired_maxhold_enabled_var.get()
        high_sensitivity = app_instance.desired_high_sensitivity_var.get()
        preamp_on = app_instance.desired_preamp_on_var.get()

        # Calculate VBW (VBW = RBW / 3 as per common practice)
        vbw_hz = rbw_hz / app_instance.VBW_RBW_RATIO
        app_instance.desired_vbw_display_var.set(f"{vbw_hz:.0f}") # Update VBW display

        print("\n--- Applying Instrument Settings ---")
        # Set RBW
        if not app_instance.inst.write(f":SENSe:BANDwidth:RESolution {rbw_hz}"): return
        debug_print(f"Sent: :SENSe:BANDwidth:RESolution {rbw_hz}", file=file, function=function)
        # Set VBW
        if not app_instance.inst.write(f":SENSe:BANDwidth:VIDeo {vbw_hz}"): return
        debug_print(f"Sent: :SENSe:BANDwidth:VIDeo {vbw_hz}", file=file, function=function)
        # Set Reference Level
        if not app_instance.inst.write(f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM"): return
        debug_print(f"Sent: :DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM", file=file, function=function)
        # Set Frequency Shift (if instrument supports it, otherwise this command might error)
        # Check if freq_shift_hz is non-zero before sending command
        if freq_shift_hz != 0.0:
            if not app_instance.inst.write(f":SENSe:FREQuency:RF:SHIFt {freq_shift_hz}"): return
            debug_print(f"Sent: :SENSe:FREQuency:RF:SHIFt {freq_shift_hz}", file=file, function=function)
        
        # Set Max Hold (Trace Type)
        if maxhold_enabled:
            if not app_instance.inst.write(":DISPlay:WINDow:TRACe:TYPE MAXH"): return
            debug_print("Sent: :DISPlay:WINDow:TRACe:TYPE MAXH", file=file, function=function)
        else:
            if not app_instance.inst.write(":DISPlay:WINDow:TRACe:TYPE NORM"): return
            debug_print("Sent: :DISPlay:WINDow:TRACe:TYPE NORM", file=file, function=function)

        # Set High Sensitivity (assuming this maps to something like noise floor extension or specific mode)
        # This command is highly instrument-specific. For N9340B, it might be related to preamp or detector.
        # Placeholder: If a specific command for "high sensitivity" exists, it would go here.
        # For now, we'll just print a debug message if it's enabled.
        if high_sensitivity:
            debug_print("High Sensitivity mode is enabled (instrument command not implemented).", file=file, function=function)
        
        # Set Preamp On/Off (if instrument supports it)
        if app_instance.instrument_model != "N9340B": # N9340B does not have a controllable preamp
            if preamp_on:
                if not app_instance.inst.write(":SENSe:POWer:RF:GAIN:STATe ON"): return
                debug_print("Sent: :SENSe:POWer:RF:GAIN:STATe ON (Preamp On)", file=file, function=function)
            else:
                if not app_instance.inst.write(":SENSe:POWer:RF:GAIN:STATe OFF"): return
                debug_print("Sent: :SENSe:POWer:RF:GAIN:STATe OFF (Preamp Off)", file=file, function=function)
        else:
            debug_print("Preamp control not available for N9340B. Skipping command.", file=file, function=function)


        print("✅ Settings applied to instrument.")
        app_instance.reset_setting_colors_logic() # Reset colors after successful application
        save_config(app_instance) # Save current settings as last used
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error applying settings: {e}")
        messagebox.showerror("VISA Error", f"Failed to apply settings to instrument: {e}")
    except Exception as e:
        print(f"❌ An unexpected error occurred while applying settings: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred while applying settings: {e}")


def update_preset_buttons(app_instance, parent_frame, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Queries device presets and populates the GUI with buttons for each preset.
    """
    debug_print("Updating preset buttons...", file=file, function=function)
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first to query presets.")
        return

    # Clear existing buttons
    for widget in parent_frame.winfo_children():
        widget.destroy()

    try:
        presets = control_query_device_presets(app_instance.inst)
        if presets:
            print(f"✅ Found {len(presets)} presets.")
            row_idx = 0
            col_idx = 0
            for preset_name in sorted(presets):
                # Create a button for each preset
                btn = ttk.Button(parent_frame, text=preset_name, style='GreyText.TButton',
                                 command=lambda name=preset_name: load_selected_preset_logic(app_instance, name))
                btn.grid(row=row_idx, column=col_idx, padx=5, pady=5, sticky="ew")
                
                col_idx += 1
                if col_idx >= 3: # 3 buttons per row
                    col_idx = 0
                    row_idx += 1
            parent_frame.grid_columnconfigure(0, weight=1)
            parent_frame.grid_columnconfigure(1, weight=1)
            parent_frame.grid_columnconfigure(2, weight=1)
        else:
            ttk.Label(parent_frame, text="No presets found on device.", background="#333333", foreground="white").pack(padx=10, pady=10)
            print("ℹ️ No presets found on device.")
    except Exception as e:
        messagebox.showerror("Preset Query Error", f"Failed to query device presets: {e}")
        print(f"❌ Error querying presets: {e}")


def load_selected_preset_logic(app_instance, selected_preset_name, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Loads the selected preset onto the instrument.
    """
    debug_print(f"Loading preset: {selected_preset_name}", file=file, function=function)
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
        return

    try:
        if control_load_selected_preset(app_instance.inst, selected_preset_name, MHZ_TO_HZ): # Pass MHZ_TO_HZ
            print(f"✅ Preset '{selected_preset_name}' loaded successfully.")
            app_instance.reset_setting_colors_logic() # Reset colors after loading preset
        else:
            messagebox.showerror("Preset Load Error", f"Failed to load preset '{selected_preset_name}'. Check console for details.")
    except Exception as e:
        messagebox.showerror("Preset Load Error", f"An unexpected error occurred while loading preset: {e}")
        print(f"❌ Error loading preset: {e}")


def set_focus_frequency_logic(app_instance, frequency_hz, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Sets the instrument's center frequency.
    """
    current_file = __file__
    current_function = inspect.currentframe().f_code.co_name
    debug_print(f"Setting focus frequency to {frequency_hz} Hz...", file=current_file, function=current_function)
    if not app_instance.inst:
        debug_print("Cannot set focus frequency: Instrument not connected.", file=current_file, function=current_function)
        messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
        return False

    try:
        # Set the center frequency
        if not app_instance.inst.write(f":SENSe:FREQuency:CENTer {frequency_hz}"): return False
        debug_print(f"Sent: :SENSe:FREQuency:CENTer {frequency_hz} Hz", file=current_file, function=current_function)
        
        # Removed the messagebox.showinfo here as requested
        print(f"✅ Instrument center frequency set to {frequency_hz / MHZ_TO_HZ:.3f} MHz.")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error while setting focus frequency: {e}")
        messagebox.showerror("VISA Error", f"Failed to set instrument center frequency: {e}")
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred while setting focus frequency: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        return False


def set_marker_and_trace_modes_logic(app_instance, marker_frequency_hz, marker_name, file=__file__, function=inspect.currentframe().f_code.co_name):
    """
    Sets a marker at the specified frequency and ensures trace mode is normal.
    """
    current_file = __file__
    current_function = inspect.currentframe().f_code.co_name
    debug_print(f"Setting marker '{marker_name}' at {marker_frequency_hz} Hz and trace modes...", file=current_file, function=current_function)
    if not app_instance.inst:
        debug_print("Cannot set marker/trace modes: Instrument not connected.", file=current_file, function=current_function)
        messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
        return False

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
        print(f"❌ An unexpected error occurred while setting marker/trace modes: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        return False

