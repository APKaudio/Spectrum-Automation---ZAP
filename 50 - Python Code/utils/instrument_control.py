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
# from tkinter import messagebox # Corrected import: directly import messagebox - REMOVED
import inspect # Import inspect module
import os # Import os module to fix NameError
from datetime import datetime # Import datetime for timestamp

# Global variable for debug mode, controlled by GUI checkbox
DEBUG_MODE = False
LOG_VISA_COMMANDS = False # New global variable for VISA command logging

def set_debug_mode(mode):
    """
    Sets the global debug mode flag. When debug mode is True,
    general debug messages are printed to the console.

    Inputs:
        mode (bool): True to enable debug mode, False to disable.
    Process:
        1. Updates the global DEBUG_MODE variable.
    Outputs: None
    """
    global DEBUG_MODE
    DEBUG_MODE = mode
    # print(f"Debug Mode set to: {DEBUG_MODE}") # This can cause recursion if debug_print uses print

def set_log_visa_commands_mode(mode):
    """
    Sets the global VISA command logging flag. When enabled,
    all VISA commands sent and received are logged.

    Inputs:
        mode (bool): True to enable logging, False to disable.
    Process:
        1. Updates the global LOG_VISA_COMMANDS variable.
    Outputs: None
    """
    global LOG_VISA_COMMANDS
    LOG_VISA_COMMANDS = mode
    # print(f"VISA Command Logging set to: {LOG_VISA_COMMANDS}") # This can cause recursion if debug_print uses print


def debug_print(message, file=None, function=None):
    """
    Prints a debug message to stdout if DEBUG_MODE is True.
    Includes timestamp, originating file, and function for better traceability.

    Inputs:
        message (str): The debug message to print.
        file (str, optional): The __file__ of the calling module.
        function (str, optional): The name of the calling function.
    Process:
        1. Checks if DEBUG_MODE is enabled.
        2. Formats the message with timestamp, file, and function info.
        3. Prints the formatted message to stdout.
    Outputs: None
    """
    if DEBUG_MODE:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3] # Milliseconds
        file_info = os.path.basename(file) if file else "N/A"
        function_info = function if function else "N/A"
        print(f"[DEBUG] [{timestamp}] [{file_info}:{function_info}] {message}")


def log_visa_command(direction, command_or_response, file=None, function=None):
    """
    Logs VISA commands sent to or responses received from the instrument if LOG_VISA_COMMANDS is True.
    Includes direction (SENT/RECV), timestamp, originating file, and function.

    Inputs:
        direction (str): "SENT" or "RECV" to indicate command direction.
        command_or_response (str): The actual VISA command string or response.
        file (str, optional): The __file__ of the calling module.
        function (str, optional): The name of the calling function.
    Process:
        1. Checks if LOG_VISA_COMMANDS is enabled.
        2. Formats the log message with direction, timestamp, file, and function info.
        3. Prints the formatted message to stdout.
    Outputs: None
    """
    if LOG_VISA_COMMANDS:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3] # Milliseconds
        file_info = os.path.basename(file) if file else "N/A"
        function_info = function if function else "N/A"
        print(f"[VISA_LOG] [{direction}] [{timestamp}] [{file_info}:{function_info}] {command_or_response.strip()}")


def query_safe(inst, command, delay=0.1):
    """
    Safely sends a query command to the instrument and returns the response.
    Includes error handling and logging.

    Inputs:
        inst (pyvisa.resources.MessageBasedResource): The PyVISA instrument object.
        command (str): The SCPI query command to send.
        delay (float): Optional delay in seconds after sending the command.
    Process:
        1. Logs the sent command.
        2. Sends the query command to the instrument.
        3. Waits for a specified delay.
        4. Reads the response from the instrument.
        5. Logs the received response.
        6. Returns the stripped response string.
    Outputs:
        str: The stripped response from the instrument, or None if an error occurs.
    Raises:
        pyvisa.errors.VisaIOError: If a VISA communication error occurs.
        Exception: For other unexpected errors.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    try:
        log_visa_command("SENT", command, file=current_file, function=current_function)
        response = inst.query(command)
        time.sleep(delay) # Small delay for stability
        log_visa_command("RECV", response, file=current_file, function=current_function)
        return response.strip()
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA IO Error during query '{command}': {e}")
        debug_print(f"VISA IO Error during query '{command}': {e}", file=current_file, function=current_function)
        # messagebox.showerror("VISA Error", f"Failed to query instrument: {e}") # Removed messagebox
        raise # Re-raise to allow higher-level error handling
    except Exception as e:
        print(f"❌ An unexpected error occurred during query '{command}': {e}")
        debug_print(f"Unexpected error during query '{command}': {e}", file=current_file, function=current_function)
        # messagebox.showerror("Error", f"An unexpected error occurred during query: {e}") # Removed messagebox
        raise # Re-raise to allow higher-level error handling


def write_safe(inst, command, delay=0.1):
    """
    Safely sends a write command to the instrument. Includes error handling and logging.

    Inputs:
        inst (pyvisa.resources.MessageBasedResource): The PyVISA instrument object.
        command (str): The SCPI write command to send.
        delay (float): Optional delay in seconds after sending the command.
    Outputs:
        bool: True if the command was written successfully, False otherwise.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    try:
        log_visa_command("SENT", command, file=current_file, function=current_function)
        inst.write(command)
        time.sleep(delay) # Small delay for stability
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA IO Error during write '{command}': {e}")
        debug_print(f"VISA IO Error during write '{command}': {e}", file=current_file, function=current_function)
        # messagebox.showerror("VISA Error", f"Failed to write to instrument: {e}") # Removed messagebox
        return False
    except Exception as e:
        print(f"❌ An unexpected error occurred during write '{command}': {e}")
        debug_print(f"Unexpected error during write '{command}': {e}", file=current_file, function=current_function)
        # messagebox.showerror("Error", f"An unexpected error occurred during write: {e}") # Removed messagebox
        return False


def list_visa_resources(rm):
    """
    Lists all available VISA resources.

    Inputs:
        rm (pyvisa.ResourceManager): The PyVISA Resource Manager instance.
    Outputs:
        list: A list of available VISA resource strings.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    try:
        resources = rm.list_resources()
        debug_print(f"Found VISA resources: {resources}", file=current_file, function=current_function)
        return list(resources)
    except Exception as e:
        print(f"❌ Error listing VISA resources: {e}")
        debug_print(f"Error listing VISA resources: {e}", file=current_file, function=current_function)
        # messagebox.showerror("VISA Error", f"Failed to list VISA resources: {e}") # Removed messagebox
        return []


def connect_to_instrument(rm, resource_name):
    """
    Connects to a specified VISA instrument and queries its identification.

    Inputs:
        rm (pyvisa.ResourceManager): The PyVISA Resource Manager instance.
        resource_name (str): The VISA resource string (e.g., 'TCPIP0::192.168.1.100::inst0::INSTR').
    Outputs:
        tuple: (instrument_object, instrument_model_string) on success, (None, None) on failure.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    inst = None
    instrument_model = None
    try:
        inst = rm.open_resource(resource_name)
        inst.timeout = 5000 # Set a timeout (milliseconds)
        inst.read_termination = '\n' # Set termination character
        inst.write_termination = '\n'
        
        # Query instrument identification
        idn = query_safe(inst, "*IDN?")
        if idn:
            print(f"Connected to: {idn}")
            debug_print(f"Connected to: {idn}", file=current_file, function=current_function)
            # Extract model from IDN string (example for Keysight/Agilent)
            parts = idn.split(',')
            if len(parts) > 1:
                instrument_model = parts[1].strip() # Model is usually the second part
                print(f"Instrument Model: {instrument_model}")
                debug_print(f"Instrument Model: {instrument_model}", file=current_file, function=current_function)
            else:
                instrument_model = "Unknown Model"
                print("Could not determine instrument model from IDN.")
                debug_print("Could not determine instrument model from IDN.", file=current_file, function=current_function)
            return inst, instrument_model
        else:
            print("❌ Failed to query instrument IDN.")
            debug_print("Failed to query instrument IDN.", file=current_file, function=current_function)
            inst.close()
            return None, None
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA IO Error connecting to {resource_name}: {e}")
        debug_print(f"VISA IO Error connecting to {resource_name}: {e}", file=current_file, function=current_function)
        # messagebox.showerror("VISA Connection Error", f"Failed to connect to {resource_name}: {e}") # Removed messagebox
        if inst: inst.close()
        return None, None
    except Exception as e:
        print(f"❌ An unexpected error occurred connecting to {resource_name}: {e}")
        debug_print(f"Unexpected error connecting to {resource_name}: {e}", file=current_file, function=current_function)
        # messagebox.showerror("Connection Error", f"An unexpected error occurred: {e}") # Removed messagebox
        if inst: inst.close()
        return None, None


def disconnect_instrument(inst):
    """
    Closes the connection to the instrument.

    Inputs:
        inst (pyvisa.resources.MessageBasedResource): The PyVISA instrument object.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    if inst:
        try:
            inst.close()
            debug_print("Instrument connection closed.", file=current_file, function=current_function)
        except Exception as e:
            print(f"❌ Error closing instrument connection: {e}")
            debug_print(f"Error closing instrument connection: {e}", file=current_file, function=current_function)
            # messagebox.showerror("Disconnection Error", f"Error closing instrument connection: {e}") # Removed messagebox


def initialize_instrument(inst, ref_level_dbm, rbw_hz, maxhold_enabled, high_sensitivity, preamp_on, instrument_model):
    """
    Initializes the spectrum analyzer with basic settings.

    Inputs:
        inst (pyvisa.resources.MessageBasedResource): The PyVISA instrument object.
        ref_level_dbm (float): Reference level in dBm.
        rbw_hz (float): Resolution Bandwidth in Hz.
        maxhold_enabled (bool): True to enable Max Hold, False otherwise.
        high_sensitivity (bool): True to enable high sensitivity (N9340B specific).
        preamp_on (bool): True to turn preamplifier on (N9340B specific).
        instrument_model (str): The detected model of the instrument.
    Outputs:
        bool: True if initialization was successful, False otherwise.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    try:
        # Reset instrument to known state
        if not write_safe(inst, "*RST"): return False
        debug_print("Sent: *RST", file=current_file, function=current_function)

        # Set Reference Level
        if not write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM"): return False
        debug_print(f"Sent: :DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM", file=current_file, function=current_function)

        # Set RBW
        if not write_safe(inst, f":SENSe:BANDwidth:RESolution {rbw_hz}"): return False
        debug_print(f"Sent: :SENSe:BANDwidth:RESolution {rbw_hz}", file=current_file, function=current_function)

        # Set VBW (RBW / 3)
        vbw_hz = rbw_hz / 3
        if not write_safe(inst, f":SENSe:BANDwidth:VIDeo {vbw_hz}"): return False
        debug_print(f"Sent: :SENSe:BANDwidth:VIDeo {vbw_hz}", file=current_file, function=current_function)

        # Set Max Hold
        max_hold_state = "ON" if maxhold_enabled else "OFF"
        if not write_safe(inst, f":DISPlay:WINDow:TRACe:MODE {max_hold_state}"): return False
        debug_print(f"Sent: :DISPlay:WINDow:TRACe:MODE {max_hold_state}", file=current_file, function=current_function)

        # Set High Sensitivity (N9340B specific)
        if instrument_model == "N9340B": # Use the passed instrument_model
            high_sensitivity_state = "ON" if high_sensitivity else "OFF"
            if not write_safe(inst, f":SENSe:POWer:RF:HSENse {high_sensitivity_state}"): return False
            debug_print(f"Sent: :SENSe:POWer:RF:HSENse {high_sensitivity_state}", file=current_file, function=current_function)
        else:
            debug_print("Instrument is not N9340B, skipping High Sensitivity setting during init.", file=current_file, function=current_function)

        # Set Preamplifier (N9340B specific)
        if instrument_model == "N9340B": # Use the passed instrument_model
            preamp_state = "ON" if preamp_on else "OFF"
            if not write_safe(inst, f":SENSe:POWer:RF:GAIN:STATe {preamp_state}"): return False
            debug_print(f"Sent: :SENSe:POWer:RF:GAIN:STATe {preamp_state}", file=current_file, function=current_function)
        else:
            debug_print("Instrument is not N9340B, skipping Preamplifier setting during init.", file=current_file, function=current_function)

        # Set Trace Type to Clear Write
        if not write_safe(inst, ":DISPlay:WINDow:TRACe:TYPE CWRite"): return False
        debug_print("Sent: :DISPlay:WINDow:TRACe:TYPE CWRite", file=current_file, function=current_function)

        # Set Sweep Time to Auto
        if not write_safe(inst, ":SENSe:SWEep:TIME:AUTO ON"): return False
        debug_print("Sent: :SENSe:SWEep:TIME:AUTO ON", file=current_file, function=current_function)

        # Set data format to ASCII
        if not write_safe(inst, ":FORMat:TRACe:DATA ASCii"): return False
        debug_print("Sent: :FORMat:TRACe:DATA ASCii", file=current_file, function=current_function)

        return True
    except Exception as e:
        print(f"❌ Error during instrument initialization: {e}")
        debug_print(f"Error during instrument initialization: {e}", file=current_file, function=current_function)
        # messagebox.showerror("Initialization Error", f"Failed to initialize instrument: {e}") # Removed messagebox
        return False


def query_current_instrument_settings(inst, MHZ_TO_HZ):
    """
    Queries the instrument for its current center frequency, span, and RBW.

    Inputs:
        inst (pyvisa.resources.MessageBasedResource): The PyVISA instrument object.
        MHZ_TO_HZ (int): Conversion factor from MHz to Hz.
    Outputs:
        tuple: (center_freq_hz, span_hz, rbw_hz) or (None, None, None) on error.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    center_freq, span, rbw = None, None, None
    try:
        center_freq_str = query_safe(inst, ":SENSe:FREQuency:CENTer?")
        span_str = query_safe(inst, ":SENSe:FREQuency:SPAN?")
        rbw_str = query_safe(inst, ":SENSe:BANDwidth:RESolution?")

        center_freq = float(center_freq_str) if center_freq_str else None
        span = float(span_str) if span_str else None
        rbw = float(rbw_str) if rbw_str else None

        if center_freq is not None and span is not None and rbw is not None:
            print(f"Current Instrument Settings: Center Freq={center_freq / MHZ_TO_HZ:.3f} MHz, Span={span / MHZ_TO_HZ:.3f} MHz, RBW={rbw:.0f} Hz")
            debug_print(f"Current Instrument Settings: Center Freq={center_freq / MHZ_TO_HZ:.3f} MHz, Span={span / MHZ_TO_HZ:.3f} MHz, RBW={rbw:.0f} Hz", file=current_file, function=current_function)
        else:
            print("❌ Could not query all current instrument settings.")
            debug_print("Could not query all current instrument settings.", file=current_file, function=current_function)
        return center_freq, span, rbw
    except Exception as e:
        print(f"❌ Error querying current instrument settings: {e}")
        debug_print(f"Error querying current instrument settings: {e}", file=current_file, function=current_function)
        return None, None, None


def query_device_presets(inst, instrument_model):
    """
    Queries the instrument for available preset files (.sta).
    This is specific to certain instrument models (e.g., N9342CN).

    Inputs:
        inst (pyvisa.resources.MessageBasedResource): The PyVISA instrument object.
        instrument_model (str): The detected model of the instrument.
    Outputs:
        list: A list of preset filenames (e.g., ['PRESET1.STA', 'MONITOR.STA']), or None if not supported/error.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if instrument_model not in ["N9342CN", "N9340B"]: # Add other models if they support this
        print(f"ℹ️ Instrument model {instrument_model} does not support direct preset querying.")
        debug_print(f"Instrument model {instrument_model} does not support direct preset querying.", file=current_file, function=current_function)
        return None

    presets = []
    try:
        # Query directory listing of C:\PRESETS
        # Command might vary by instrument. This is a common one for Agilent/Keysight.
        dir_listing = query_safe(inst, ':MMEMory:CATalog? "C:\\PRESETS\\*.STA"')
        debug_print(f"Preset directory listing raw response: {dir_listing}", file=current_file, function=current_function)

        if dir_listing:
            # Parse the response. It might be a comma-separated list of quoted strings.
            # Example: '"C:\PRESETS\PRESET1.STA","C:\PRESETS\MONITOR.STA"'
            # We need to extract just the filename.
            # Use regex to find quoted strings and extract the filename part
            # This regex looks for "C:\PRESETS\<filename>" and captures <filename>
            matches = re.findall(r'"C:\\\\PRESETS\\\\([a-zA-Z0-9_.-]+\.STA)"', dir_listing, re.IGNORECASE)
            presets = [match.upper() for match in matches] # Convert to uppercase for consistency
            debug_print(f"Parsed device presets: {presets}", file=current_file, function=current_function)
        return presets
    except pyvisa.errors.VisaIOError as e:
        print(f"❌ VISA error querying device presets: {e}")
        debug_print(f"VISA error querying device presets: {e}", file=current_file, function=current_function)
        return None
    except Exception as e:
        print(f"❌ An unexpected error occurred while querying device presets: {e}")
        debug_print(f"Unexpected error querying device presets: {e}", file=current_file, function=current_function)
        return None


def load_selected_preset(inst, selected_preset_name):
    """
    Loads a specified preset file (.sta) onto the instrument.

    Inputs:
        inst (pyvisa.resources.MessageBasedResource): The PyVISA instrument object.
        selected_preset_name (str): The name of the preset file (e.g., 'MYPRESET.STA').
    Outputs:
        tuple: (bool, center_freq_hz, span_hz, rbw_hz) indicating success and queried values or None.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not inst:
        debug_print("Not connected to instrument, cannot load preset.", file=current_file, function=current_function)
        return False, None, None, None

    # Ensure the path is correctly formatted for the instrument.
    # Instruments typically expect backslashes and quotes escaped.
    # Example: C:\PRESETS\MYPRESET.STA -> "C:\\PRESETS\\MYPRESET.STA"
    preset_path = f"C:\\\\PRESETS\\\\{selected_preset_name}"
    command = f':MMEMory:LOAD STA,"{preset_path}"'
    
    print(f"\nAttempting to load preset: {selected_preset_name}")
    try:
        if write_safe(inst, command):
            print(f"✅ Preset '{selected_preset_name}' loaded successfully.")
            
            # Import MHZ_TO_HZ locally for this function to avoid circular dependency
            from utils.frequency_bands import MHZ_TO_HZ 
            
            # Query and display current instrument settings after loading preset
            center_freq, span, rbw = query_current_instrument_settings(inst, MHZ_TO_HZ)
            return True, center_freq, span, rbw
        else:
            debug_print(f"Failed to load preset '{selected_preset_name}'.", file=current_file, function=current_function)
            return False, None, None, None
    except Exception as e:
        debug_print(f"An unexpected error occurred while loading preset: {e}", file=current_file, function=current_function)
        # messagebox.showerror("Preset Load Error", f"An unexpected error occurred while loading preset: {e}") # Removed messagebox
        return False, None, None, None

