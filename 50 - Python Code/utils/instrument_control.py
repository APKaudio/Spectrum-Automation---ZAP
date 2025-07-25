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

def initialize_instrument(inst, ref_level, high_sensitivity_on, preamp_on, rbw_config, vbw_config, instrument_model):
    """
    Initializes the instrument with basic settings (reset, reference level, preamp, RBW, VBW).

    Inputs:
        inst (pyvisa.resources.Resource): The VISA instrument object.
        ref_level (float): The desired reference level in dBm.
        high_sensitivity_on (bool): True to enable high sensitivity (preamp).
        preamp_on (bool): True to enable the preamplifier.
        rbw_config (int): Resolution Bandwidth in Hz.
        vbw_config (int): Video Bandwidth in Hz.
        instrument_model (str): The model of the instrument (e.g., "N9340B").
    Process:
        1. Resets the instrument to its default state (`*RST`).
        2. Sets the display to normal mode (`:DISPlay:ENABle ON`).
        3. Sets the reference level (`:DISPlay:WINDow:TRACe:Y:RLEVel`).
        4. Configures high sensitivity/preamp based on `instrument_model`.
        5. Sets RBW and VBW.
        6. Handles errors during instrument configuration.
    Outputs:
        bool: True if initialization is successful, False otherwise.
    """
    if not inst:
        print("Not connected to instrument, cannot initialize.")
        return False
    try:
        # Reset instrument
        if not write_safe(inst, "*RST"): return False
        print("Instrument reset.")

        # Set display on
        if not write_safe(inst, ":DISPlay:ENABle ON"): return False
        print("Display enabled.")

        # Set reference level
        if not write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level}DBM"): return False
        print(f"Reference Level set to {ref_level} dBm.")

        # Configure high sensitivity/preamp based on instrument model
        if instrument_model == "N9340B":
            # N9340B uses :INPut:ATTenuator:AUTO for high sensitivity
            # and :INPut:GAIN:STATe for preamp
            if high_sensitivity_on:
                if not write_safe(inst, ":INPut:ATTenuator:AUTO OFF"): return False
                if not write_safe(inst, ":INPut:ATTenuator 0DB"): return False # 0 dB attenuation for high sens
                print("N9340B High Sensitivity (Attenuator 0dB) ON.")
            else:
                if not write_safe(inst, ":INPut:ATTenuator:AUTO ON"): return False
                print("N9340B High Sensitivity (Attenuator Auto) OFF.")

            if preamp_on:
                if not write_safe(inst, ":INPut:GAIN:STATe ON"): return False
                print("N9340B Preamp ON.")
            else:
                if not write_safe(inst, ":INPut:GAIN:STATe OFF"): return False
                print("N9340B Preamp OFF.")
        else:
            # Generic commands for other instruments (adjust as needed)
            if high_sensitivity_on:
                if not write_safe(inst, ":INPut:ATTenuator:AUTO OFF"): return False
                if not write_safe(inst, ":INPut:ATTenuator 0DB"): return False
                print("Generic High Sensitivity (Attenuator 0dB) ON.")
            else:
                if not write_safe(inst, ":INPut:ATTenuator:AUTO ON"): return False
                print("Generic High Sensitivity (Attenuator Auto) OFF.")

            if preamp_on:
                if not write_safe(inst, ":INPut:GAIN:STATe ON"): return False
                print("Generic Preamp ON.")
            else:
                if not write_safe(inst, ":INPut:GAIN:STATe OFF"): return False
                print("Generic Preamp OFF.")

        # Set RBW and VBW
        if not write_safe(inst, f":SENSe:BANDwidth:RESolution {rbw_config}"): return False
        print(f"RBW set to {rbw_config} Hz.")
        if not write_safe(inst, f":SENSe:BANDwidth:VIDeo {vbw_config}"): return False
        print(f"VBW set to {vbw_config} Hz.")

        return True
    except Exception as e:
        messagebox.showerror("Initialization Error", f"Failed to initialize instrument settings: {e}")
        print(f"❌ Initialization Error: {e}")
        return False

def query_current_instrument_settings(inst, MHZ_TO_HZ):
    """
    Queries and prints the current key settings of the connected instrument.

    Inputs:
        inst (pyvisa.resources.Resource): The VISA instrument object.
        MHZ_TO_HZ (int): Conversion factor from MHz to Hz.
    Process:
        1. Queries various settings (center frequency, span, RBW, VBW, Ref Level, Preamp, Attenuator).
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

        preamp_state = query_safe(inst, ":INPut:GAIN:STATe?")
        if preamp_state: print(f"Preamplifier State: {'ON' if int(preamp_state) == 1 else 'OFF'}")

        attenuator_auto = query_safe(inst, ":INPut:ATTenuator:AUTO?")
        if attenuator_auto: print(f"Attenuator Auto: {'ON' if int(attenuator_auto) == 1 else 'OFF'}")

        attenuator_value = query_safe(inst, ":INPut:ATTenuator?")
        if attenuator_value: print(f"Attenuator Value: {float(attenuator_value):.1f} dB")

    except Exception as e:
        print(f"❌ Error querying current instrument settings: {e}")
    print("-----------------------------------")

def query_device_presets(inst):
    """
    Queries the connected instrument for available preset (.STA) files.
    This function is specific to instruments that support `:MMEMory:CATalog:STATe?`.

    Inputs:
        inst (pyvisa.resources.Resource): The VISA instrument object.
    Process:
        1. Sends the SCPI command `:MMEMory:CATalog:STATe?` to list preset files.
        2. Parses the response to extract individual preset file names.
        3. Prints the list of found presets.
        4. Handles `pyvisa.errors.VisaIOError` and general `Exception`.
    Outputs:
        list: A list of preset file names (strings), or an empty list if none found or error.
    """
    if not inst:
        print("Not connected to instrument, cannot query presets.")
        return []
    
    print("\nAttempting to query device presets...")
    try:
        # Query preset files (e.g., "D:\PRESET\MY_PRESET.STA")
        response = query_safe(inst, ":MMEMory:CATalog:STATe?")
        
        if response:
            # Response format example: '"C:\PRESETS\DEFAULT.STA","C:\PRESETS\MY_PRESET.STA"'
            # We need to split by comma and strip quotes, and only take the basename
            preset_paths = [path.strip().strip('"') for path in response.split(',')]
            preset_names = [p.split('\\')[-1] for p in preset_paths if p.lower().endswith('.sta')]
            print(f"Found {len(preset_names)} preset files: {preset_names}")
            return preset_names
        else:
            print("No .STA preset files found or empty response.")
            return []
    except Exception as e:
        print(f"❌ An unexpected error occurred while querying presets: {e}")
        return []

def load_selected_preset(inst, selected_preset_name, MHZ_TO_HZ):
    """
    Loads a specified preset file onto the connected instrument.

    Inputs:
        inst (pyvisa.resources.Resource): The VISA instrument object.
        selected_preset_name (str): The name of the preset file to load (e.g., "MY_PRESET.STA").
        MHZ_TO_HZ (int): Conversion factor from MHz to Hz, used for querying settings after load.
    Process:
        1. Constructs the full path for the preset (assumes C:\PRESETS\ on instrument).
        2. Sends the SCPI command `:MMEMory:LOAD STA,"{preset_path}"` using `write_safe`.
        3. If loading is successful, calls `query_current_instrument_settings` to display
           the instrument's new configuration.
        4. Prints status messages.
        5. Handles general `Exception` during the loading process.
    Outputs:
        bool: True if the preset is loaded successfully; False otherwise.
    """
    if not inst:
        print("Not connected to instrument, cannot load preset.")
        return False

    # Assuming presets are in C:\PRESETS on the instrument
    preset_path = f"C:\\PRESETS\\{selected_preset_name}"
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
