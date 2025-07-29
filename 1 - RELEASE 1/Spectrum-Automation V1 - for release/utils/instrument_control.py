# instrument_control.py
#
# This module provides a high-level interface for controlling the RF spectrum analyzer
# via PyVISA. It abstracts the low-level SCPI commands, offering functions for
# connecting, disconnecting, initializing, and querying instrument settings.
# It also includes robust error handling for VISA communication.
#
# Author: Anthony Peter Kuzub
# Blog: www.Like.audio (Contributor to this project)
#
# Professional services for customizing and tailoring this software to your specific
# application can be negotiated. There is no change to use, modify, or fork this software.
#
# Build Log: https://like.audio/category/software/spectrum-scanner/
# Source Code: https://github.com/APKaudio/
#
import pyvisa
import time
import re
import tkinter as tk # Used for messagebox in this module for direct instrument errors

# Global variable for debug mode, controlled by the GUI
_debug_mode_enabled = False

def set_debug_mode(enabled):
    """
    Sets the global debug mode variable for this module.
    When debug mode is enabled, `query_safe` and `write_safe` functions
    will print the SCPI commands sent and responses received to the console.

    Inputs:
        enabled (bool): True to enable debug mode, False to disable.
    Process:
        1. Updates the `_debug_mode_enabled` global variable.
        2. Prints a confirmation message to the console.
    Outputs:
        None. (Side effect: Modifies a global variable and prints to console.)
    """
    global _debug_mode_enabled
    _debug_mode_enabled = enabled
    print(f"Instrument Control Debug Mode: {'Enabled' if _debug_mode_enabled else 'Disabled'}")

def query_safe(inst, command):
    """
    Safely queries the instrument, handling PyVISA errors.
    This function sends a query command to the instrument and returns its response.
    It includes error handling for `VisaIOError` (communication issues) and other exceptions.

    Inputs:
        inst (pyvisa.resources.Resource): The connected PyVISA instrument object.
        command (str): The SCPI query command to send to the instrument (e.g., "*IDN?").
    Process:
        1. Attempts to execute `inst.query(command)`.
        2. Strips whitespace from the response.
        3. If `_debug_mode_enabled` is True, prints the query and its response.
        4. Catches `pyvisa.errors.VisaIOError` for communication specific issues.
        5. Catches general `Exception` for other parsing or unexpected errors.
    Outputs:
        str or None: The stripped response string from the instrument if successful;
                     None if an error occurs.
    """
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
    """
    Safely writes a command to the instrument, handling PyVISA errors.
    This function sends a command to the instrument without expecting a direct response.
    It includes error handling for `VisaIOError` and other exceptions.

    Inputs:
        inst (pyvisa.resources.Resource): The connected PyVISA instrument object.
        command (str): The SCPI command to write to the instrument (e.g., ":SENSE:FREQ:START 1000000").
    Process:
        1. Attempts to execute `inst.write(command)`.
        2. If `_debug_mode_enabled` is True, prints the command being written.
        3. Catches `pyvisa.errors.VisaIOError` for communication specific issues.
        4. Catches general `Exception` for other unexpected errors.
    Outputs:
        bool: True if the command was written successfully; False if an error occurs.
    """
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
    Lists available VISA resources (connected instruments).
    This function uses the PyVISA resource manager to discover all connected
    VISA-compliant devices.

    Inputs:
        resource_manager (pyvisa.ResourceManager): The PyVISA resource manager instance.
    Process:
        1. Calls `resource_manager.list_resources()` to get a tuple of resource strings.
        2. Converts the tuple to a list.
        3. Prints the found resources to the console.
        4. Catches any `Exception` during the listing process.
    Outputs:
        list: A list of available VISA resource strings (e.g., 'USB0::0x0957::0xFFEF::SG05300002::0::INSTR').
              Returns an empty list on failure.
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
    This function attempts to open a connection to a specified VISA resource,
    sets communication timeouts and termination characters, and queries the
    instrument's identification to confirm a successful connection.

    Inputs:
        resource_manager (pyvisa.ResourceManager): The PyVISA resource manager instance.
        selected_resource (str): The VISA resource string to connect to (e.g., 'USB0::0x0957::0xFFEF::SG05300002::0::INSTR').
    Process:
        1. Attempts to open the resource using `resource_manager.open_resource()`.
        2. Sets `inst.timeout` to 5 seconds and `inst.read_termination`/`inst.write_termination` to newline.
        3. Queries the instrument's ID using `*IDN?`.
        4. Extracts the instrument model (N9342CN or N9340B) from the IDN string using regex.
        5. Prints connection status and detected model.
        6. Handles `pyvisa.errors.VisaIOError` (connection specific issues) and general `Exception`.
        7. Ensures the instrument connection is closed on failure.
    Outputs:
        tuple: `(pyvisa.resources.Resource, str)` - The connected instrument object and its model string.
               Returns `(None, None)` on failure to connect or identify.
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

    Inputs:
        inst (pyvisa.resources.Resource): The instrument object to disconnect.
    Process:
        1. If `inst` is not None, attempts to close the connection using `inst.close()`.
        2. Prints a disconnection message.
        3. Catches any `Exception` during the disconnection process.
    Outputs:
        bool: True if disconnected successfully or if no instrument was connected; False on error.
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
    """
    Initializes the spectrum analyzer with basic settings such as reference level,
    preamplifier state, high sensitivity mode, and trace configurations.
    This function sets up the instrument for a scan.

    Inputs:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        ref_level_dbm (float): The desired reference level in dBm.
        high_sensitivity_on (bool): True to enable high sensitivity mode, False otherwise.
        preamp_on (bool): True to turn the preamplifier ON, False otherwise.
        rbw_config_val (float): The Resolution Bandwidth (RBW) value to configure on the instrument in Hz.
                                (Note: This parameter is currently not directly used for setting RBW in this function
                                but is passed for consistency with the GUI's intent. RBW for scanning is set in `scan_bands`.)
        vbw_config_val (float): The Video Bandwidth (VBW) value to configure on the instrument in Hz.
                                (Note: This parameter is currently not directly used for setting VBW in this function
                                but is passed for consistency with the GUI's intent. VBW for scanning is set in `scan_bands`.)
        model_match (str): The detected model of the instrument (e.g., "N9340B", "N9342CN").
                           Used for model-specific SCPI commands.
    Process:
        1. **Reset**: Sends `*RST` to reset the instrument to a known state, then waits for operation completion.
        2. **Reference Level**: Sets the display reference level.
        3. **Preamplifier/High Sensitivity**: Configures the preamplifier and high sensitivity mode based on `preamp_on`
           and `high_sensitivity_on` flags. This involves setting attenuation and gain.
        4. **Trace Modes**: Configures Trace 1 to 'WRITe', Trace 2 to 'MAXHold', and Trace 3 to 'MINHold'.
        5. **Display Scale**: Sets the Y-axis display scale to 'LOGarithmic'.
        6. **Sweep Time**: Sets sweep time to 'AUTO'.
        7. **Data Format**: Sets the trace data format to 'ASCII' for data transfer, with a model-specific command.
        8. Prints status messages for each configuration step.
        9. Handles `pyvisa.errors.VisaIOError` and general `Exception`.
    Outputs:
        bool: True if initialization is successful; False on failure.
    """
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
    Queries the instrument for its current settings (Reference Level, Preamplifier,
    RBW, VBW, Sweep Time, Start/Stop Frequency) and prints them to the console.
    This provides a snapshot of the instrument's state.

    Inputs:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        MHZ_TO_HZ (int): Conversion factor from MHz to Hz, used for displaying frequencies.
    Process:
        1. Checks if `inst` is connected.
        2. Queries various settings using `query_safe` (e.g., `:DISPlay:WINDow:TRACe:Y:RLEVel?`).
        3. Prints the queried values to the console, converting frequencies to MHz for readability.
        4. Handles `ValueError` if frequency strings cannot be converted to float.
        5. Catches general `Exception` for any other errors during querying.
    Outputs:
        None. (Side effect: Prints instrument settings to console.)
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
    Queries the connected instrument for a list of preset files stored in its
    internal "C:\\PRESETS\\" directory. This allows the GUI to display available
    presets for loading.

    Inputs:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
    Process:
        1. Checks if `inst` is connected.
        2. Sends the SCPI command `:MMEMory:CATalog? "C:\\\\PRESETS\\\\"` to list directory contents.
        3. Parses the comma-separated response string to extract file names and types.
        4. Filters for files with the "STA" type (state files) and ending with ".STA".
        5. Sorts the found preset names alphabetically.
        6. Prints the number of found presets or a message if none are found.
        7. Handles `pyvisa.errors.VisaIOError` and general `Exception` during the query/parsing.
    Outputs:
        list: A sorted list of `.STA` preset file names (e.g., `['MY_PRESET.STA', 'DEFAULT.STA']`).
              Returns an empty list on failure or if no presets are found.
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
    This function sends the SCPI command to instruct the spectrum analyzer
    to load a previously saved state file (`.STA`). After loading, it
    queries and prints the instrument's current settings to confirm the change.

    Inputs:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        selected_preset_name (str): The name of the preset file to load (e.g., "MY_PRESET.STA").
        MHZ_TO_HZ (int): Conversion factor from MHz to Hz, passed to `query_current_instrument_settings`.
    Process:
        1. Checks if `inst` is connected.
        2. Constructs the full path to the preset file (e.g., `C:\\PRESETS\\MY_PRESET.STA`).
        3. Sends the SCPI command `:MMEMory:LOAD STA,"{preset_path}"` using `write_safe`.
        4. If loading is successful, calls `query_current_instrument_settings` to display
           the instrument's new configuration.
        5. Prints status messages.
        6. Handles general `Exception` during the loading process.
    Outputs:
        bool: True if the preset is loaded successfully; False otherwise.
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
