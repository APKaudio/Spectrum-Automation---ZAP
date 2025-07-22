# instrument_control.py

import pyvisa
import time
import re
import tkinter as tk # Used for messagebox in this module for direct instrument errors

# Global variable for debug mode, controlled by the GUI
_debug_mode_enabled = False

def set_debug_mode(enabled):
    """Sets the global debug mode variable for this module."""
    global _debug_mode_enabled
    _debug_mode_enabled = enabled
    print(f"Instrument Control Debug Mode: {'Enabled' if _debug_mode_enabled else 'Disabled'}")

def query_safe(inst, command):
    """Safely queries the instrument, handling VISA errors."""
    try:
        response = inst.query(command).strip()
        if _debug_mode_enabled: # Conditional print
            print(f"Query: '{command}' -> Response: '{response}'")
        return response
    except pyvisa.errors.VisaIOError as e:
        print(f"VISA Error during query '{command}': {e}")
        return None
    except Exception as e:
        print(f"Error parsing response for '{command}': {e}")
        return None

def write_safe(inst, command):
    """Safely writes to the instrument, handling VISA errors."""
    try:
        inst.write(command)
        if _debug_mode_enabled: # Conditional print
            print(f"Write: '{command}'")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"VISA Error during write '{command}': {e}")
        return False

def list_visa_resources(resource_manager):
    """
    Lists available VISA resources.
    Args:
        resource_manager (pyvisa.ResourceManager): The PyVISA resource manager instance.
    Returns:
        list: A list of available VISA resource strings.
    """
    try:
        resources = resource_manager.list_resources()
        print(f"Found VISA resources: {resources}")
        return list(resources)
    except Exception as e:
        print(f"Error listing VISA resources: {e}")
        return []

def connect_to_instrument(resource_manager, selected_resource):
    """
    Establishes connection to the selected instrument.
    Args:
        resource_manager (pyvisa.ResourceManager): The PyVISA resource manager instance.
        selected_resource (str): The VISA resource string to connect to.
    Returns:
        tuple: (pyvisa.resources.Resource, str) - The connected instrument object and its model string.
               Returns (None, None) on failure.
    """
    inst = None
    instrument_model = None
    try:
        inst = resource_manager.open_resource(selected_resource)
        inst.timeout = 5000 # 5 seconds timeout
        inst.read_termination = '\n' # N9340B typically terminates with newline
        inst.write_termination = '\n'
        
        # Query instrument ID
        instrument_id = query_safe(inst, "*IDN?")
        if instrument_id:
            print(f"✅ Connected to: {instrument_id.strip()}")
            model_match = re.search(r'(N9342CN|N9340B)', instrument_id) 
            instrument_model = model_match.group(0) if model_match else "Unknown Model"
            print(f"Detected Instrument Model: {instrument_model}")
            return inst, instrument_model
        else:
            print("🚫 Could not query instrument ID. Check connection or address.")
            inst.close()
            return None, None
    except pyvisa.errors.VisaIOError as e:
        print(f"🛑 VISA Error connecting to {selected_resource}: {e}")
        if inst:
            inst.close()
        return None, None
    except Exception as e:
        print(f"🚨 An unexpected error occurred during connection: {e}")
        if inst:
            inst.close()
        return None, None

def disconnect_instrument(inst):
    """
    Closes the connection to the instrument.
    Args:
        inst (pyvisa.resources.Resource): The instrument object to disconnect.
    Returns:
        bool: True if disconnected successfully, False otherwise.
    """
    if inst:
        try:
            inst.close()
            print("🔌 Instrument disconnected.")
            return True
        except Exception as e:
            print(f"Error disconnecting instrument: {e}")
            return False
    return True # Already disconnected or no instrument to begin with

def initialize_instrument(inst, ref_level_dbm, high_sensitivity_on, preamp_on, rbw_config_val, vbw_config_val, model_match):
    """Initializes the spectrum analyzer with basic settings."""
    print("✨ Initializing instrument with desired settings...")
    try:
        # Reset the instrument to a known state using *RST first
        write_safe(inst, "*RST")
        query_safe(inst, "*OPC?") # Wait for operation to complete
        time.sleep(1) # Give it a moment after reset

        # Set reference level
        write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}")
        print(f"✅ Set reference level to {ref_level_dbm} dBm.")

        # Set preamplifier
        if preamp_on:
            write_safe(inst, ":POWer:ATTenuation:AUTO ON")
            write_safe(inst, ":POWer:GAIN ON")
            print("✅ Preamplifier ON.")
            write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM")
            print(f"✅ Set reference level to {ref_level_dbm} dBm.")
        else:
            write_safe(inst, ":POWer:GAIN OFF")
            print("✅ Preamplifier OFF.")

        # Set high sensitivity (preamplifier)
        if high_sensitivity_on:
            write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel -50")
            write_safe(inst, ":POWer:ATTenuation 0")
            write_safe(inst, ":POWer:GAIN 1")
            write_safe(inst, ":POWer:HSENsitive ON")
            print("✅ High sensitivity turned ON.")
        else:
            write_safe(inst, ":POWer:HSENsitive OFF")
            write_safe(inst, ":POWer:ATTenuation 10")
            write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM")
            print(f"✅ Set reference level to {ref_level_dbm} dBm.")
            print("✅ High sensitivity turned OFF.")
        
        write_safe(inst, ":TRAC1:MODE WRITe")
        print(f"✅ Trace 1 sent to write")

        write_safe(inst, ":TRAC2:MODE MAXHold")
        print(f"✅ Trace 2 sent to MAX HOLD")

        write_safe(inst, ":TRAC3:MODE MINHold")
        print(f"✅ Trace 3 sent to Min Hold")
        
        # Display scale is always LOGarithmic
        write_safe(inst, ":DISPlay:WINDow:TRACe:Y:SCALe:SPACing LOGarithmic")
        print("✅ Display scale set to LOGarithmic (always).")
        
        write_safe(inst, ":SENSe:BANDwidth:VIDeo:AUTO ON")
        write_safe(inst, ":SENSe:SWEep:TIME:AUTO ON")
        print("✅ Sweep time set to AUTO.")

        if model_match == "N9340B":
            write_safe(inst, ":TRACe:FORMat:DATA ASCii") 
        else:
            write_safe(inst, ":FORMat:DATA ASCii") 
            print("✅ Set trace data format to ASCII for data transfer.")
      
        print("🎉 Instrument initialized successfully with desired settings.")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"🛑 Failed to initialize instrument with desired settings: {e}")
        return False
    except Exception as e:
        print(f"🚨 An unexpected error occurred during instrument initialization: {e}")
        return False

def query_current_instrument_settings(inst, MHZ_TO_HZ):
    """
    Queries the instrument for its current settings and prints them to the console.
    Args:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        MHZ_TO_HZ (int): Conversion factor from MHz to Hz.
    """
    if not inst:
        print("Not connected to instrument, cannot query settings.")
        return

    print("\nQuerying current instrument settings from device (console log only)...")
    try:
        print(f"  Reference Level (dBm): {query_safe(inst, ':DISPlay:WINDow:TRACe:Y:RLEVel?')}")
        print(f"  Preamplifier (High Sensitivity ON/OFF): {query_safe(inst, ':POWer:GAIN?')}")
        
        rbw_hz = query_safe(inst, ':SENSe:BANDwidth:RESolution?')
        if rbw_hz:
            try:
                print(f"  RBW: {float(rbw_hz) / MHZ_TO_HZ:.3f} MHz")
            except ValueError:
                print(f"  RBW: {rbw_hz} (could not convert to MHz)")

        print(f"  VBW (Hz): {query_safe(inst, ':SENSe:BANDwidth:VIDeo?')}")
        print(f"  Sweep Time Auto (ON/OFF): {query_safe(inst, ':SENSe:SWEep:TIME:AUTO?')}")
        
        start_freq_hz = query_safe(inst, ':SENSe:FREQuency:STARt?')
        if start_freq_hz:
            try:
                print(f"  Start Freq: {float(start_freq_hz) / MHZ_TO_HZ:.3f} MHz")
            except ValueError:
                print(f"  Start Freq: {start_freq_hz} (could not convert to MHz)")
        
        stop_freq_hz = query_safe(inst, ':SENSe:FREQuency:STOP?')
        if stop_freq_hz:
            try:
                print(f"  Stop Freq: {float(stop_freq_hz) / MHZ_TO_HZ:.3f} MHz")
            except ValueError:
                print(f"  Stop Freq: {stop_freq_hz} (could not convert to MHz)")

        print("✅ Current instrument settings queried successfully.")

    except Exception as e:
        print(f"🛑 Error querying instrument settings: {e}")

def query_device_presets(inst):
    """
    Queries the connected instrument for a list of preset files in "C:\\PRESETS\\".
    Args:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
    Returns:
        list: A list of .STA preset file names.
    """
    if not inst:
        print("Not connected to instrument, cannot query device presets.")
        return []

    print("\nQuerying device preset files from C:\\PRESETS\\...")
    preset_files = []
    try:
        response = query_safe(inst, ':MMEMory:CATalog? "C:\\\\PRESETS\\\\"')

        if response is None:
            print("🚫 No response received for preset catalog query.")
            return []

        parts = response.split(',')
        if len(parts) < 3:
            print(f"🚫 Unexpected response format for preset catalog: {response}")
            return []

        # The actual item listings start after the first 3 parts
        for i in range(3, len(parts), 4):
            if i + 3 < len(parts):
                name = parts[i].strip()
                item_type = parts[i+1].strip()
                if item_type.upper() == "STA" and name.upper().endswith(".STA"):
                    preset_files.append(name)
            else:
                print(f"Warning: Incomplete item entry found at index {i} in preset catalog response.")
                break

        if preset_files:
            print(f"✅ Found {len(preset_files)} '.STA' preset files.")
        else:
            print("🚫 No '.STA' preset files found in C:\\PRESETS\\.")
        return sorted(preset_files) # Return sorted list
    except pyvisa.errors.VisaIOError as e:
        print(f"🛑 VISA Error querying device presets: {e}")
        return []
    except Exception as e:
        print(f"❌ An unexpected error occurred while querying presets: {e}")
        return []

def load_selected_preset(inst, selected_preset_name, MHZ_TO_HZ):
    """
    Loads the selected preset file onto the instrument.
    Args:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        selected_preset_name (str): The name of the preset file to load.
        MHZ_TO_HZ (int): Conversion factor from MHz to Hz.
    Returns:
        bool: True if preset loaded successfully, False otherwise.
    """
    if not inst:
        print("Not connected to instrument, cannot load preset.")
        return False

    preset_path = f"C:\\\\PRESETS\\\\{selected_preset_name}"
    command = f':MMEMory:LOAD STA,"{preset_path}"'
    
    print(f"\nAttempting to load preset: {selected_preset_name}")
    try:
        if write_safe(inst, command):
            print(f"✅ Preset '{selected_preset_name}' loaded successfully.")
            
            # Query and display current instrument settings after loading preset
            print("\n--- Current Instrument Settings after Preset Load ---")
            query_current_instrument_settings(inst, MHZ_TO_HZ) # Use the dedicated query function
            print("--------------------------------------------------")
            return True
        else:
            print(f"❌ Failed to load preset '{selected_preset_name}'.")
            return False
    except Exception as e:
        print(f"❌ An unexpected error occurred while loading preset: {e}")
        return False
