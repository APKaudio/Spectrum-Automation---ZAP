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
import tkinter as tk # For messagebox, used in _debug_mode_enabled and within scan_bands

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
    print(f"Debug mode set to: {DEBUG_MODE}")

def debug_print(message):
    """
    Prints a debug message to the console if DEBUG_MODE is enabled.

    Inputs:
        message (str): The message to print.
    Outputs: None
    """
    if DEBUG_MODE:
        print(f"DEBUG: {message}")

def write_safe(inst, command):
    """
    Safely writes a SCPI command to the instrument.
    Includes error handling and debug printing.

    Inputs:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        command (str): The SCPI command string to write.
    Process:
        1. Prints the command if debug mode is enabled.
        2. Attempts to write the command to the instrument.
        3. Catches `pyvisa.errors.VisaIOError` and prints an error message.
    Outputs:
        bool: True if the command was written successfully, False otherwise.
    """
    debug_print(f"Sending: {command}")
    try:
        inst.write(command)
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"🚫 VISA Write Error for command '{command}': {e}")
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred while writing command '{command}': {e}")
        return False

def query_safe(inst, command):
    """
    Safely queries the instrument for data.
    Includes error handling and debug printing.

    Inputs:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        command (str): The SCPI query command string.
    Process:
        1. Prints the command if debug mode is enabled.
        2. Attempts to query the instrument and read the response.
        3. Prints the response if debug mode is enabled.
        4. Catches `pyvisa.errors.VisaIOError` and prints an error message.
    Outputs:
        str or None: The response string from the instrument, or None on failure.
    """
    debug_print(f"Querying: {command}")
    try:
        response = inst.query(command).strip()
        debug_print(f"Received: {response}")
        return response
    except pyvisa.errors.VisaIOError as e:
        print(f"🚫 VISA Query Error for command '{command}': {e}")
        return None
    except Exception as e:
        print(f"❌ An unexpected error occurred while querying command '{command}': {e}")
        return None

def list_visa_resources(rm):
    """
    Lists available VISA resources (instruments) on the system.

    Inputs:
        rm (pyvisa.ResourceManager): The PyVISA resource manager instance.
    Process:
        1. Attempts to list resources using `rm.list_resources()`.
        2. Handles `pyvisa.errors.VisaIOError` if no VISA backend is found.
        3. Prints the found resources.
    Outputs:
        list: A list of available VISA resource strings.
    """
    try:
        resources = rm.list_resources()
        debug_print(f"Found VISA resources: {resources}")
        return list(resources)
    except pyvisa.errors.VisaIOError as e:
        print(f"🚫 VISA Error: No VISA backend found or error listing resources. Is NI-VISA/Keysight VISA installed? {e}")
        tk.messagebox.showerror("VISA Error", f"No VISA backend found or error listing resources. Is NI-VISA/Keysight VISA installed?\n\nError: {e}")
        return []
    except Exception as e:
        print(f"❌ An unexpected error occurred while listing VISA resources: {e}")
        tk.messagebox.showerror("Error", f"An unexpected error occurred while listing VISA resources: {e}")
        return []

def connect_to_instrument(rm, resource_name):
    """
    Establishes a connection to the specified VISA instrument.

    Inputs:
        rm (pyvisa.ResourceManager): The PyVISA resource manager instance.
        resource_name (str): The VISA address of the instrument to connect to.
    Process:
        1. Attempts to open the instrument resource.
        2. Sets a timeout for communication.
        3. Queries the instrument's identification string (`*IDN?`).
        4. Extracts the instrument model from the IDN string.
        5. Handles `pyvisa.errors.VisaIOError` and other exceptions during connection.
    Outputs:
        tuple: (instrument_object, model_string) if successful, (None, None) otherwise.
    """
    inst = None
    model = None
    try:
        inst = rm.open_resource(resource_name)
        inst.timeout = 5000 # Set a timeout of 5 seconds
        idn = query_safe(inst, "*IDN?")
        if idn:
            parts = idn.split(',')
            if len(parts) >= 2:
                model = parts[1].strip()
                print(f"Connected to: {idn}")
                return inst, model
            else:
                print(f"Warning: Could not parse model from IDN: {idn}")
                return inst, "Unknown Model"
        else:
            print(f"🚫 Failed to get IDN from {resource_name}.")
            inst.close()
            return None, None
    except pyvisa.errors.VisaIOError as e:
        print(f"🚫 VISA Connection Error to {resource_name}: {e}")
        tk.messagebox.showerror("Connection Error", f"Failed to connect to {resource_name}:\n{e}")
        if inst: inst.close()
        return None, None
    except Exception as e:
        print(f"❌ An unexpected error occurred during connection to {resource_name}: {e}")
        tk.messagebox.showerror("Error", f"An unexpected error occurred during connection to {resource_name}:\n{e}")
        if inst: inst.close()
        return None, None

def disconnect_instrument(inst):
    """
    Closes the connection to the instrument.

    Inputs:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
    Process:
        1. Attempts to close the instrument connection.
        2. Handles `pyvisa.errors.VisaIOError` and other exceptions.
    Outputs:
        bool: True if disconnected successfully, False otherwise.
    """
    if inst:
        try:
            inst.close()
            return True
        except pyvisa.errors.VisaIOError as e:
            print(f"🚫 VISA Disconnect Error: {e}")
            tk.messagebox.showerror("Disconnect Error", f"Failed to disconnect from instrument:\n{e}")
            return False
        except Exception as e:
            print(f"❌ An unexpected error occurred during disconnection: {e}")
            tk.messagebox.showerror("Error", f"An unexpected error occurred during disconnection:\n{e}")
            return False
    return True # Already disconnected or no instrument to begin with

def initialize_instrument(inst, ref_level, high_sensitivity_on, preamp_on, rbw_config, vbw_config, instrument_model):
    """
    Initializes the spectrum analyzer with specified settings.

    Inputs:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        ref_level (float): The reference level in dBm.
        high_sensitivity_on (bool): True to enable high sensitivity mode.
        preamp_on (bool): True to turn on the preamplifier.
        rbw_config (int): The Resolution Bandwidth in Hz.
        vbw_config (int): The Video Bandwidth in Hz.
        instrument_model (str): The model string of the connected instrument.
    Process:
        1. Sets the instrument to preset state.
        2. Configures display, trace, and detector modes.
        3. Sets reference level, attenuation, preamp, and RBW/VBW.
        4. Handles model-specific commands (e.g., N9340B).
        5. Prints status messages.
        6. Handles `Exception` during initialization.
    Outputs:
        bool: True if initialization is successful, False otherwise.
    """
    print("\nInitializing instrument settings...")
    try:
        # Reset to preset state
        if not write_safe(inst, "*RST"): return False
        if not write_safe(inst, ":SYSTem:PRESet"): return False # Agilent-specific preset command

        # Configure Display and Trace
        if not write_safe(inst, ":DISPlay:ENABle ON"): return False
        if not write_safe(inst, ":TRACe:TYPE NORM"): return False # Set trace to normal mode (Clear Write)
        if not write_safe(inst, ":DETector:FUNCtion AVERage"): return False # Set detector to Average

        # Reference Level
        if not write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level}DBM"): return False

        # Attenuation (Auto or specific) - For N9340B, typically auto-managed or specific commands
        if instrument_model == "N9340B":
            # N9340B might not have direct ATTenuation commands like higher-end models.
            # It often manages it automatically based on Ref Level.
            # If specific attenuation control is needed, refer to N9340B programming manual.
            pass
        else:
            # For other instruments, you might set attenuation based on ref_level
            # For simplicity, let's assume auto for now or a fixed value if needed.
            # if not write_safe(inst, ":POWer:ATTenuation:AUTO ON"): return False
            pass

        # Preamplifier
        if preamp_on:
            if not write_safe(inst, ":INPut:GAIN:STATe ON"): return False # Common command for preamp
            print("Preamplifier: ON")
        else:
            if not write_safe(inst, ":INPut:GAIN:STATe OFF"): return False
            print("Preamplifier: OFF")

        # High Sensitivity Mode (if applicable, often related to preamp/attenuation)
        if high_sensitivity_on:
            # This is often handled by preamp or specific instrument modes.
            # For N9340B, it's typically tied to preamp and input range.
            print("High Sensitivity Mode: Enabled (via preamp/settings)")
        else:
            print("High Sensitivity Mode: Disabled")

        # RBW and VBW
        if not write_safe(inst, f":SENSe:BANDwidth:RESolution {rbw_config}"): return False
        if not write_safe(inst, f":SENSe:BANDwidth:VIDeo {vbw_config}"): return False
        
        print(f"RBW set to: {rbw_config} Hz")
        print(f"VBW set to: {vbw_config} Hz")

        print("✅ Instrument initialization complete.")
        return True
    except Exception as e:
        print(f"❌ Error initializing instrument: {e}")
        tk.messagebox.showerror("Initialization Error", f"Failed to initialize instrument: {e}")
        return False

def query_current_instrument_settings(inst, MHZ_TO_HZ):
    """
    Queries and prints the current key settings of the instrument.

    Inputs:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        MHZ_TO_HZ (int): Conversion factor from MHz to Hz.
    Process:
        1. Queries various instrument settings (frequency, span, RBW, VBW, Ref Level, Attenuation, Preamplifier).
        2. Prints the queried settings to the console.
        3. Handles `Exception` during queries.
    Outputs: None
    """
    print("\n--- Current Instrument Settings ---")
    try:
        center_freq = query_safe(inst, ":SENSe:FREQuency:CENTer?")
        span = query_safe(inst, ":SENSe:FREQuency:SPAN?")
        rbw = query_safe(inst, ":SENSe:BANDwidth:RESolution?")
        vbw = query_safe(inst, ":SENSe:BANDwidth:VIDeo?")
        ref_level = query_safe(inst, ":DISPlay:WINDow:TRACe:Y:RLEVel?")
        attenuation = query_safe(inst, ":POWer:ATTenuation?") # May not be supported by all models
        preamp_state = query_safe(inst, ":INPut:GAIN:STATe?") # Common command for preamp state

        print(f"Center Frequency: {float(center_freq) / MHZ_TO_HZ:.3f} MHz")
        print(f"Span: {float(span) / MHZ_TO_HZ:.3f} MHz")
        print(f"RBW: {float(rbw)} Hz")
        print(f"VBW: {float(vbw)} Hz")
        print(f"Reference Level: {float(ref_level):.2f} dBm")
        print(f"Attenuation: {float(attenuation):.2f} dB" if attenuation else "Attenuation: N/A or Auto")
        print(f"Preamplifier: {'ON' if preamp_state and int(float(preamp_state)) == 1 else 'OFF'}")

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
        2. Sends the SCPI command `:MMEMory:CATalog? "C:\\PRESETS\\"` to list directory contents.
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

        debug_print(f"Raw response from :MMEMory:CATalog?: '{response}'")

        if response is None:
            print("🚫 No response received for preset catalog query.")
            return []

        parts = response.split(',')
        debug_print(f"Parsed response parts: {parts}")

        # The actual item listings start after the first 3 parts (dir_count, file_count, total_size)
        # Each file entry is typically 4 parts: name, type, size, date/time
        if len(parts) < 3: # Minimum expected parts: directory count, file count, total size
            print(f"🚫 Unexpected response format for preset catalog: Not enough initial parts. Response: '{response}'")
            return []

        # Check if there are actual file entries
        if len(parts) >= 4: # At least one file entry (name, type, size, date)
            # Iterate starting from the 4th part (index 3) with a step of 4
            for i in range(3, len(parts), 4):
                if i + 3 < len(parts): # Ensure there are enough parts for a full entry (name, type, size, date)
                    name = parts[i].strip().strip('"') # Remove quotes if present
                    item_type = parts[i+1].strip().strip('"')
                    # parts[i+2] is size, parts[i+3] is date/time - not used for filtering

                    debug_print(f"  Found entry: Name='{name}', Type='{item_type}'")

                    if item_type.upper() == "STA" and name.upper().endswith(".STA"):
                        preset_files.append(name)
                else:
                    print(f"Warning: Incomplete item entry found at index {i} in preset catalog response. Remaining parts: {parts[i:]}")
                    break # Stop processing if an incomplete entry is found

        if preset_files:
            print(f"✅ Found {len(preset_files)} '.STA' preset files.")
        else:
            print("🚫 No '.STA' preset files found in C:\\PRESETS\\ or response was empty/malformed.")
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

    # Ensure backslashes are correctly escaped for the SCPI command
    # The SCPI command needs the path with single backslashes in the string literal,
    # but Python string itself needs double backslashes for literal backslashes.
    # The f-string will correctly interpret \\ as \
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
