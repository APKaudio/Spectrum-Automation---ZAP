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

def set_log_visa_commands_mode(mode):
    """
    Sets the global VISA command logging flag. When True,
    all VISA commands sent and received are printed.

    Inputs:
        mode (bool): True to enable logging, False to disable.
    Process:
        1. Updates the global LOG_VISA_COMMANDS variable.
    Outputs: None
    """
    global LOG_VISA_COMMANDS
    LOG_VISA_COMMANDS = mode

def debug_print(message, file=None, function=None, console_print_func=None):
    """
    Prints a debug message if DEBUG_MODE is enabled.
    Includes file and function context for better traceability.
    """
    if DEBUG_MODE:
        timestamp = datetime.now().strftime("%M.%S")
        prefix = ""
        if file:
            prefix += f"[{os.path.basename(file)}"
            if function:
                prefix += f":{function}] "
            else:
                prefix += "] "
        elif function:
            prefix += f"[{function}] "
        
        full_message = f"🚫🐛 [{timestamp}] {prefix}{message}"
        if console_print_func:
            console_print_func(full_message)
        else:
            print(full_message) # Fallback to standard print


def log_visa_command(command, direction="SENT", console_print_func=None):
    """
    Logs VISA commands sent to or received from the instrument if LOG_VISA_COMMANDS is enabled.
    """
    if LOG_VISA_COMMANDS:
        timestamp = datetime.now().strftime("%M.%S")
        log_message = f"💳🌲 [{timestamp}] {direction}: {command.strip()}"
        if console_print_func:
            console_print_func(log_message)
        else:
            print(log_message)


def list_visa_resources(console_print_func=None):
    """
    Lists available VISA resources (instruments).
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Listing VISA resources...", file=current_file, function=current_function, console_print_func=console_print_func)
    try:
        rm = pyvisa.ResourceManager()
        resources = rm.list_resources()
        debug_print(f"Found VISA resources: {resources}", file=current_file, function=current_function, console_print_func=console_print_func)
        return list(resources)
    except Exception as e:
        error_msg = f"❌ Error listing VISA resources: {e}"
        if console_print_func:
            console_print_func(error_msg)
        debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
        return []


def connect_to_instrument(resource_name, console_print_func=None):
    """
    Establishes a connection to a VISA instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Connecting to instrument: {resource_name}", file=current_file, function=current_function, console_print_func=console_print_func)
    try:
        rm = pyvisa.ResourceManager()
        inst = rm.open_resource(resource_name)
        inst.timeout = 5000 # Set a timeout (milliseconds)
        inst.read_termination = '\n' # Set read termination character
        inst.write_termination = '\n' # Set write termination character
        inst.query_delay = 0.1 # Small delay between write and read for query
        if console_print_func:
            console_print_func(f"✅ Successfully connected to {resource_name}")
        debug_print(f"Connection successful to {resource_name}", file=current_file, function=current_function, console_print_func=console_print_func)
        return inst
    except pyvisa.errors.VisaIOError as e:
        error_msg = f"❌ VISA error connecting to {resource_name}: {e}"
        if console_print_func:
            console_print_func(error_msg)
        debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
        return None
    except Exception as e:
        error_msg = f"❌ An unexpected error occurred while connecting to {resource_name}: {e}"
        if console_print_func:
            console_print_func(error_msg)
        debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
        return None


def disconnect_instrument(inst, console_print_func=None):
    """
    Closes the connection to a VISA instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Disconnecting instrument...", file=current_file, function=current_function, console_print_func=console_print_func)
    if inst:
        try:
            inst.close()
            if console_print_func:
                console_print_func("✅ Instrument disconnected.")
            debug_print("Instrument connection closed.", file=current_file, function=current_function, console_print_func=console_print_func)
            return True
        except pyvisa.errors.VisaIOError as e:
            error_msg = f"❌ VISA error disconnecting instrument: {e}"
            if console_print_func:
                console_print_func(error_msg)
            debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
            return False
        except Exception as e:
            error_msg = f"❌ An unexpected error occurred while disconnecting instrument: {e}"
            if console_print_func:
                console_print_func(error_msg)
            debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
            return False
    return False


def write_safe(inst, command, console_print_func=None):
    """
    Safely writes a command to the instrument.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    log_visa_command(command, "SENT", console_print_func)
    try:
        inst.write(command)
        debug_print(f"Command sent: {command.strip()}", file=current_file, function=current_function, console_print_func=console_print_func)
        return True
    except pyvisa.errors.VisaIOError as e:
        error_msg = f"❌ VISA error sending command '{command.strip()}': {e}"
        if console_print_func:
            console_print_func(error_msg)
        debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
        return False
    except Exception as e:
        error_msg = f"❌ An unexpected error occurred while sending command '{command.strip()}': {e}"
        if console_print_func:
            console_print_func(error_msg)
        debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
        return False


def query_safe(inst, command, console_print_func=None):
    """
    Safely queries the instrument and returns the response.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    log_visa_command(command, "SENT", console_print_func)
    try:
        response = inst.query(command).strip()
        log_visa_command(response, "RECEIVED", console_print_func)
        debug_print(f"Query '{command.strip()}' response: {response}", file=current_file, function=current_function, console_print_func=console_print_func)
        return response
    except pyvisa.errors.VisaIOError as e:
        error_msg = f"❌ VISA error querying '{command.strip()}': {e}"
        if console_print_func:
            console_print_func(error_msg)
        debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
        return None
    except Exception as e:
        error_msg = f"❌ An unexpected error occurred while querying '{command.strip()}': {e}"
        if console_print_func:
            console_print_func(error_msg)
        debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
        return None


def initialize_instrument(inst, console_print_func=None):
    """
    Initializes the instrument to a known state.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Initializing instrument...", file=current_file, function=current_function, console_print_func=console_print_func)
    try:
        if not write_safe(inst, "*RST", console_print_func): return False # Reset instrument
        if not write_safe(inst, "*CLS", console_print_func): return False # Clear status
        if not write_safe(inst, ":SYSTem:DISPlay:UPDate ON", console_print_func): return False # Enable display updates
        if not write_safe(inst, ":FORMat:TRACe:DATA ASCii", console_print_func): return False # ASCII format for trace data
        if not write_safe(inst, ":DISPlay:WINDow:TRACe:TYPE NORM", console_print_func): return False # Normal trace mode
        if not write_safe(inst, ":POW:ATT:AUTO ON", console_print_func): return False # Auto attenuator
        if not write_safe(inst, ":FREQ:CENT 1GHZ", console_print_func): return False # Default center frequency
        if not write_safe(inst, ":FREQ:SPAN 1GHZ", console_print_func): return False # Default span
        if not write_safe(inst, ":BAND:RES 10000", console_print_func): return False # Default RBW
        if not write_safe(inst, ":BAND:VID 3000", console_print_func): return False # Default VBW (RBW/3)
        
        
        debug_print("Instrument initialization commands sent.", file=current_file, function=current_function, console_print_func=console_print_func)
        return True
    except Exception as e:
        error_msg = f"❌ Error during instrument initialization: {e}"
        if console_print_func:
            console_print_func(error_msg)
        debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
        return False


def query_current_instrument_settings(inst, MHZ_TO_HZ, console_print_func=None):
    """
    Queries the instrument for its current center frequency, span, and RBW.
    Returns values in MHz for frequency/span and Hz for RBW.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Querying current instrument settings...", file=current_file, function=current_function, console_print_func=console_print_func)

    center_freq_hz = None
    span_hz = None
    rbw_hz = None

    try:
        center_freq_str = query_safe(inst, ":SENSe:FREQuency:CENTer?", console_print_func)
        if center_freq_str is not None:
            center_freq_hz = float(center_freq_str)
        
        span_str = query_safe(inst, ":SENSe:FREQuency:SPAN?", console_print_func)
        if span_str is not None:
            span_hz = float(span_str)
        
        rbw_str = query_safe(inst, ":BAND:RES?", console_print_func)
        if rbw_str is not None:
            rbw_hz = float(rbw_str)

        debug_print(f"Queried settings: Center Freq={center_freq_hz} Hz, Span={span_hz} Hz, RBW={rbw_hz} Hz", file=current_file, function=current_function, console_print_func=console_print_func)

        return (center_freq_hz / MHZ_TO_HZ) if center_freq_hz is not None else None, \
               (span_hz / MHZ_TO_HZ) if span_hz is not None else None, \
               rbw_hz # RBW is returned in Hz

    except ValueError as e:
        error_msg = f"❌ Error parsing instrument query response: {e}"
        if console_print_func:
            console_print_func(error_msg)
        debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None, None
    except Exception as e:
        error_msg = f"❌ An unexpected error occurred while querying instrument settings: {e}"
        if console_print_func:
            console_print_func(error_msg)
        debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None, None


def query_device_presets(inst, console_print_func=None):
    """
    Queries the instrument for a list of available preset (.sta) files.
    This command is specific to certain Keysight/Agilent instruments.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Querying device presets from instrument...", file=current_file, function=current_function, console_print_func=console_print_func)
    
    presets = []
    try:
        # Query the directory of .sta files. The exact command might vary by instrument.
        # This is a common command for Keysight/Agilent.
        response = query_safe(inst, ":MMEMory:CATalog? \"C:\\PRESETS\\*.STA\"", console_print_func)
        
        if response:
            # The response typically looks like: '"C:\PRESETS\PRESET1.STA","C:\PRESETS\PRESET2.STA"'
            # or '1,"C:\PRESETS\",1,"PRESET1.STA",1,"PRESET2.STA"' depending on instrument firmware.
            # We need to parse this to get just the filenames.
            
            # Remove quotes and split by comma
            cleaned_response = response.strip().strip('"')
            
            # Attempt to parse based on common formats
            if "," in cleaned_response:
                parts = cleaned_response.split(',')
                for part in parts:
                    # Remove any leading/trailing quotes and backslashes
                    filename = part.strip().strip('"').replace('\\', '/')
                    if filename.lower().endswith(".sta"):
                        # Extract just the filename from the path
                        base_filename = os.path.basename(filename)
                        presets.append(base_filename)
            elif cleaned_response.lower().endswith(".sta"):
                # Single preset returned
                base_filename = os.path.basename(cleaned_response.replace('\\', '/'))
                presets.append(base_filename)
            
            debug_print(f"Parsed device presets: {presets}", file=current_file, function=current_function, console_print_func=console_print_func)
            return presets
        else:
            debug_print("No response received for preset catalog query.", file=current_file, function=current_function, console_print_func=console_print_func)
            return []
    except Exception as e:
        error_msg = f"❌ Error querying device presets: {e}"
        if console_print_func:
            console_print_func(error_msg)
        debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
        return []


def load_selected_preset(inst, selected_preset_name, console_print_func=None):
    """
    Loads a specified preset file onto the connected instrument.
    Returns True on success, False on failure, and the instrument's
    current center frequency, span, and RBW after loading.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Loading preset '{selected_preset_name}' to instrument.", file=current_file, function=current_function, console_print_func=console_print_func)

    if not inst:
        if console_print_func:
            console_print_func("⚠️ Warning: No instrument connected. Cannot load preset.")
        debug_print("No instrument connected for loading preset.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False, None, None, None

    # Ensure the path is correctly formatted for the instrument.
    # Instruments typically expect backslashes and quotes escaped.
    # Example: C:\PRESETS\MYPRESET.STA -> "C:\\PRESETS\\MYPRESET.STA"
    preset_path = f"C:\\\\PRESETS\\\\{selected_preset_name}"
    command = f':MMEMory:LOAD STA,\"{preset_path}\"'
    
    if console_print_func:
        console_print_func(f"\nAttempting to load preset: {selected_preset_name}")
    try:
        if write_safe(inst, command, console_print_func):
            if console_print_func:
                console_print_func(f"✅ Preset '{selected_preset_name}' loaded successfully.")
            
            # Import MHZ_TO_HZ locally for this function to avoid circular dependency
            from utils.frequency_bands import MHZ_TO_HZ 
            
            # Query and display current instrument settings after loading preset
            center_freq, span, rbw = query_current_instrument_settings(inst, MHZ_TO_HZ, console_print_func)
            return True, center_freq, span, rbw
        else:
            debug_print(f"Failed to load preset '{selected_preset_name}'.", file=current_file, function=current_function, console_print_func=console_print_func)
            return False, None, None, None
    except Exception as e:
        debug_print(f"An unexpected error occurred while loading preset: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        if console_print_func:
            console_print_func(f"❌ Error loading preset '{selected_preset_name}': {e}")
        return False, None, None, None