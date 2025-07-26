# utils/instrument_control.py
#
# This module provides low-level functions for communicating with the spectrum analyzer
# via PyVISA. It includes functions for safely writing commands, querying data,
# connecting/disconnecting, initializing instrument settings, and managing device presets.
# This module is designed to abstract the direct VISA communication details from the
# higher-level application logic.
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
from tkinter import messagebox # Corrected import: directly import messagebox

# Global variable for debug mode, controlled by GUI checkbox
DEBUG_MODE = False

def set_debug_mode(mode):
    """
    Sets the global debug mode flag. When debug mode is True,
    VISA commands and responses are printed to the console.

    Inputs:
        mode (bool): True to enable debug mode, False to disable.
    Process:
        1. Updates the global `DEBUG_MODE` variable.
        2. Prints the new debug mode status.
    Outputs: None
    """
    global DEBUG_MODE
    DEBUG_MODE = mode
    debug_print(f"Debug Mode set to: {DEBUG_MODE}") # Changed to debug_print

def debug_print(message):
    """
    Prints a debug message to the console only if DEBUG_MODE is enabled.

    Inputs:
        message (str): The message string to print.
    Outputs: None
    """
    if DEBUG_MODE:
        print(f"DEBUG: {message}")

def log_visa_command(direction, command_or_response):
    """
    Logs VISA commands sent to and responses received from the instrument
    if debug mode is enabled.

    Inputs:
        direction (str): "SENT" for commands sent, "RECV" for responses received.
        command_or_response (str): The actual SCPI command string or instrument response.
    Process:
        1. If `DEBUG_MODE` is True, prints the direction and the command/response.
    Outputs: None
    """
    if DEBUG_MODE:
        print(f"VISA {direction}: {command_or_response.strip()}")

def query_safe(inst, command, delay=0.1):
    """
    Safely queries the instrument and returns the response.
    Includes error handling and debug logging.

    Inputs:
        inst (pyvisa.resources.Resource): The VISA instrument object.
        command (str): The SCPI query command to send.
        delay (float, optional): Time in seconds to wait after sending the command. Defaults to 0.1.
    Process:
        1. Calls `log_visa_command` to log the sent command.
        2. Attempts to query the instrument using `inst.query()`.
        3. Calls `log_visa_command` to log the received response.
        4. Introduces a `time.sleep` for stability.
        5. Handles `pyvisa.errors.VisaIOError` and general `Exception`.
        6. Displays an error messagebox on failure.
    Outputs:
        str: The instrument's response if successful, an empty string otherwise.
    """
    if not inst:
        debug_print("Not connected to instrument, cannot query.") # Changed to debug_print
        return ""
    try:
        log_visa_command("SENT", command)
        response = inst.query(command).strip()
        log_visa_command("RECV", response)
        time.sleep(delay) # Small delay for stability
        return response
    except pyvisa.errors.VisaIOError as e:
        messagebox.showerror("VISA IO Error", f"Error querying instrument: {e}\nCommand: {command}")
        print(f"❌ VISA IO Error during query: {e}")
        return ""
    except Exception as e:
        messagebox.showerror("Instrument Error", f"An unexpected error occurred during query: {e}\nCommand: {command}")
        print(f"❌ Unexpected error during query: {e}")
        return ""

def write_safe(inst, command, delay=0.1):
    """
    Safely writes a command to the instrument.
    Includes error handling and debug logging.

    Inputs:
        inst (pyvisa.resources.Resource): The VISA instrument object.
        command (str): The SCPI command to send.
        delay (float, optional): Time in seconds to wait after sending the command. Defaults to 0.1.
    Process:
        1. Calls `log_visa_command` to log the sent command.
        2. Attempts to write to the instrument using `inst.write()`.
        3. Introduces a `time.sleep` for stability.
        4. Handles `pyvisa.errors.VisaIOError` and general `Exception`.
        5. Displays an error messagebox on failure.
    Outputs:
        bool: True if the command was written successfully, False otherwise.
    """
    if not inst:
        debug_print("Not connected to instrument, cannot write.") # Changed to debug_print
        return False
    try:
        log_visa_command("SENT", command)
        inst.write(command)
        time.sleep(delay) # Small delay for stability
        return True
    except pyvisa.errors.VisaIOError as e:
        messagebox.showerror("VISA IO Error", f"Error writing to instrument: {e}\nCommand: {command}")
        print(f"❌ VISA IO Error during write: {e}")
        return False
    except Exception as e:
        messagebox.showerror("Instrument Error", f"An unexpected error occurred during write: {e}\nCommand: {command}")
        print(f"❌ Unexpected error during write: {e}")
        return False

def list_visa_resources(rm):
    """
    Lists available VISA resources (instruments).

    Inputs:
        rm (pyvisa.ResourceManager): The PyVISA resource manager instance.
    Process:
        1. Attempts to list resources using `rm.list_resources()`.
        2. Handles `pyvisa.errors.VisaIOError` if no backend is found or other VISA issues.
        3. Displays an error messagebox on failure.
    Outputs:
        list: A list of available VISA resource strings.
    """
    try:
        resources = rm.list_resources()
        debug_print(f"Found VISA resources: {resources}") # Changed to debug_print
        return list(resources)
    except pyvisa.errors.VisaIOError as e:
        messagebox.showerror("VISA Error", f"Could not list VISA resources. Is NI-VISA or Keysight VISA installed?\nError: {e}")
        print(f"❌ VISA Error listing resources: {e}")
        return []
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred while listing resources: {e}")
        print(f"❌ Unexpected error listing resources: {e}")
        return []

def connect_to_instrument(rm, resource_name):
    """
    Connects to a specific VISA instrument.

    Inputs:
        rm (pyvisa.ResourceManager): The PyVISA resource manager instance.
        resource_name (str): The full VISA resource string of the instrument.
    Process:
        1. Attempts to open the instrument resource.
        2. Queries the instrument's identification string (`*IDN?`).
        3. Extracts the instrument model from the IDN string.
        4. Handles `pyvisa.errors.VisaIOError` and general `Exception`.
        5. Displays an error messagebox on failure.
    Outputs:
        tuple: (instrument_object, instrument_model_string) if successful, (None, None) otherwise.
    """
    inst = None
    instrument_model = None
    try:
        inst = rm.open_resource(resource_name)
        inst.timeout = 5000 # Set a timeout (5 seconds)
        print(f"✅ Connected to {resource_name}")
        
        # Query instrument identification
        idn = query_safe(inst, "*IDN?")
        if idn:
            parts = idn.split(',')
            if len(parts) > 1:
                instrument_model = parts[1].strip()
                debug_print(f"Instrument Model: {instrument_model}") # Changed to debug_print
            else:
                instrument_model = "Unknown Model"
                debug_print(f"Instrument IDN: {idn} (Model not parsed)") # Changed to debug_print
        else:
            instrument_model = "Unknown Model"
            debug_print("Could not query instrument IDN.") # Changed to debug_print
        
        return inst, instrument_model
    except pyvisa.errors.VisaIOError as e:
        messagebox.showerror("Connection Error", f"Could not connect to {resource_name}: {e}")
        print(f"❌ Connection Error: {e}")
        if inst:
            inst.close()
        return None, None
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred during connection: {e}")
        print(f"❌ Unexpected error during connection: {e}")
        if inst:
            inst.close()
        return None, None

def disconnect_instrument(inst):
    """
    Disconnects from the given VISA instrument.

    Inputs:
        inst (pyvisa.resources.Resource): The VISA instrument object to disconnect.
    Process:
        1. Checks if the instrument object is valid.
        2. Attempts to close the instrument connection.
        3. Handles `pyvisa.errors.VisaIOError` and general `Exception`.
        4. Displays an error messagebox on failure.
    Outputs:
        bool: True if disconnected successfully, False otherwise.
    """
    if inst:
        try:
            inst.close()
            print("✅ Instrument disconnected.")
            return True
        except pyvisa.errors.VisaIOError as e:
            messagebox.showerror("Disconnect Error", f"Error disconnecting instrument: {e}")
            print(f"❌ Disconnect Error: {e}")
            return False
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred during disconnection: {e}")
            print(f"❌ Unexpected error during disconnection: {e}")
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
        if not write_safe(inst, "*RST"): return False
        if not query_safe(inst, "*OPC?"): return False # Wait for operation to complete
        time.sleep(1) # Give it a moment after reset

        # Set reference level
        if not write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}"): return False
        print(f"✅ Set reference level to {ref_level_dbm} dBm.")

        # Set preamplifier
        if preamp_on:
            if not write_safe(inst, ":POWer:ATTenuation:AUTO ON"): return False
            if not write_safe(inst, ":POWer:GAIN ON"): return False
            print("✅ Preamplifier ON.")
            # Note: The original code re-set RLEVel here, preserving that behavior
            if not write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM"): return False
            print(f"✅ Set reference level to {ref_level_dbm} dBm.")
        else:
            if not write_safe(inst, ":POWer:GAIN OFF"): return False
            print("✅ Preamplifier OFF.")

        # Set high sensitivity (preamplifier)
        if high_sensitivity_on:
            # Note: The original code set RLEVel to -50 here, preserving that behavior
            if not write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel -50"): return False
            if not write_safe(inst, ":POWer:ATTenuation 0"): return False
            if not write_safe(inst, ":POWer:GAIN 1"): return False
            if not write_safe(inst, ":POWer:HSENsitive ON"): return False
            print("✅ High sensitivity turned ON.")
        else:
            if not write_safe(inst, ":POWer:HSENsitive OFF"): return False
            if not write_safe(inst, ":POWer:ATTenuation 10"): return False
            # Note: The original code re-set RLEVel here, preserving that behavior
            if not write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM"): return False
            print(f"✅ Set reference level to {ref_level_dbm} dBm.")
            print("✅ High sensitivity turned OFF.")
        
        # Configure Trace Modes
        if not write_safe(inst, ":TRAC1:MODE WRITe"): return False
        print(f"✅ Trace 1 sent to write")

        if not write_safe(inst, ":TRAC2:MODE MAXHold"): return False
        print(f"✅ Trace 2 sent to MAX HOLD")

        if not write_safe(inst, ":TRAC3:MODE MINHold"): return False
        print(f"✅ Trace 3 sent to Min Hold")
        
        # Display scale is always LOGarithmic
        if not write_safe(inst, ":DISPlay:WINDow:TRACe:Y:SCALe:SPACing LOGarithmic"): return False
        print("✅ Display scale set to LOGarithmic (always).")
        
        # Set VBW and Sweep Time to AUTO
        if not write_safe(inst, ":SENSe:BANDwidth:VIDeo:AUTO ON"): return False
        if not write_safe(inst, ":SENSe:SWEep:TIME:AUTO ON"): return False
        print("✅ VBW and Sweep time set to AUTO.")

        # Set trace data format
        if model_match == "N9340B":
            if not write_safe(inst, ":TRACe:FORMat:DATA ASCii"): return False
        else:
            if not write_safe(inst, ":FORMat:DATA ASCii"): return False
        print("✅ Set trace data format to ASCII for data transfer.")
      
        print("🎉 Instrument initialized successfully with desired settings.")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"🛑 Failed to initialize instrument with desired settings: {e}")
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred during instrument initialization: {e}")
        messagebox.showerror("Initialization Error", f"An unexpected error occurred during initialization: {e}")
        return False

def query_current_instrument_settings(inst, MHZ_TO_HZ):
    """
    Queries and prints the current key settings of the connected instrument.

    Inputs:
        inst (pyvisa.resources.Resource): The VISA instrument object.
        MHZ_TO_HZ (int): Conversion factor from MHz to Hz.
    Process:
        1. Queries various settings (center frequency, span, RBW, VBW, Ref Level).
        2. Prints the queried settings to the console.
        3. Handles errors during querying.
    Outputs: None
    """
    if not inst:
        print("Not connected to instrument, cannot query settings.")
        return

    print("\n--- Current Instrument Settings ---")
    try:
        center_freq_hz = query_safe(inst, ":SENSe:FREQuency:CENTer?")
        if center_freq_hz: print(f"Center Frequency: {float(center_freq_hz) / MHZ_TO_HZ:.3f} MHz")

        span_hz = query_safe(inst, ":SENSe:FREQuency:SPAN?")
        if span_hz: print(f"Span: {float(span_hz)} Hz")

        rbw = query_safe(inst, ":SENSe:BANDwidth:RESolution?")
        if rbw: print(f"Resolution Bandwidth (RBW): {float(rbw)} Hz")

        vbw = query_safe(inst, ":SENSe:BANDwidth:VIDeo?")
        if vbw: print(f"Video Bandwidth (VBW): {float(vbw)} Hz")

        ref_level = query_safe(inst, ":DISPlay:WINDow:TRACe:Y:RLEVel?")
        if ref_level: print(f"Reference Level: {float(ref_level):.2f} dBm")

        # Removed the following queries as they were causing timeouts:
        # preamp_state = query_safe(inst, ":INPut:GAIN:STATe?")
        # if preamp_state: print(f"Preamplifier State: {'ON' if int(float(preamp_state)) == 1 else 'OFF'}")

        # attenuator_auto = query_safe(inst, ":INPut:ATTenuator:AUTO?")
        # if attenuator_auto: print(f"Attenuator Auto: {'ON' if int(float(attenuator_auto)) == 1 else 'OFF'}")

        # attenuator_value = query_safe(inst, ":INPut:ATTenuator?")
        # if attenuator_value: print(f"Attenuator Value: {float(attenuator_value):.1f} dB")

    except Exception as e:
        print(f"❌ Error querying current instrument settings: {e}")
    print("-----------------------------------")

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
                name = parts[i].strip().strip('"') # Strip quotes
                item_type = parts[i+1].strip().strip('"') # Strip quotes
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
        messagebox.showerror("VISA Error", f"Error querying device presets: {e}")
        return []
    except Exception as e:
        print(f"❌ An unexpected error occurred while querying presets: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred while querying presets: {e}")
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
            
            # --- Removed explicit delays and OPC query after preset load as requested ---
            # The *OPC? after *RST in initialize_instrument should be sufficient for initial setup.
            # If further stability issues arise, consider adding specific delays where needed.
            # --- End removed section ---

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
        messagebox.showerror("Preset Load Error", f"An unexpected error occurred while loading preset: {e}")
        return False