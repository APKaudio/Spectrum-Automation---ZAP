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
        log_message = f"�🌲 [{timestamp}] {direction}: {command.strip()}"
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
    Returns an empty string if an error occurs or no response, to prevent NoneType errors.
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
        return "" # Return empty string on error
    except Exception as e:
        error_msg = f"❌ An unexpected error occurred while querying '{command.strip()}': {e}"
        if console_print_func:
            console_print_func(error_msg)
        debug_print(error_msg, file=current_file, function=current_function, console_print_func=console_print_func)
        return "" # Return empty string on error


def initialize_instrument(inst, ref_level_dbm, high_sensitivity_on, preamp_on, rbw_config_val, vbw_config_val, model_match, console_print_func=None):
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
        console_print_func (function, optional): Function to use for console output.
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
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    if console_print_func:
        console_print_func("✨ Initializing instrument with desired settings.")
    debug_print("Initializing instrument with desired settings...", file=current_file, function=current_function, console_print_func=console_print_func)
    try:
        # Reset the instrument to a known state using *RST first
        if not write_safe(inst, "*RST", console_print_func):
            debug_print("Failed to send *RST.", file=current_file, function=current_function, console_print_func=console_print_func)
            return False
        time.sleep(0.5) # Small delay after reset to allow instrument to process
        if not query_safe(inst, "*OPC?", console_print_func):
            debug_print("Failed to query *OPC? after *RST (timeout likely).", file=current_file, function=current_function, console_print_func=console_print_func)
            return False # Wait for operation to complete
        time.sleep(1) # Give it a moment after reset and OPC

        
                # Set preamplifier
        if preamp_on:
            if not write_safe(inst, ":POWer:ATTenuation:AUTO ON", console_print_func):
                debug_print("Failed to set :POWer:ATTenuation:AUTO ON.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
            if not write_safe(inst, ":POWer:GAIN ON", console_print_func):
                debug_print("Failed to set :POWer:GAIN ON.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
            if console_print_func:
                console_print_func("✅ Preamplifier ON.")
            debug_print("Preamplifier ON.", file=current_file, function=current_function, console_print_func=console_print_func)
            # Note: The original code re-set RLEVel here, preserving that behavior
            if not write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM", console_print_func):
                debug_print(f"Failed to re-set reference level to {ref_level_dbm} dBm after preamp config.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
            if console_print_func:
                console_print_func(f"✅ Set reference level to {ref_level_dbm} dBm.")
            debug_print(f"Re-set reference level to {ref_level_dbm} dBm after preamp config.", file=current_file, function=current_function, console_print_func=console_print_func)
        else:
            if not write_safe(inst, ":POWer:GAIN OFF", console_print_func):
                debug_print("Failed to set :POWer:GAIN OFF.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
            if console_print_func:
                console_print_func("✅ Preamplifier OFF.")
            debug_print("Preamplifier OFF.", file=current_file, function=current_function, console_print_func=console_print_func)

        # Set high sensitivity (preamplifier)
        if high_sensitivity_on:
            # Note: The original code set RLEVel to -50 here, preserving that behavior
            if not write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel -50D", console_print_func):
                debug_print("Failed to set reference level to -50 dBm for high sensitivity.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
            if not write_safe(inst, ":POWer:ATTenuation 0", console_print_func):
                debug_print("Failed to set :POWer:ATTenuation 0.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
            if not write_safe(inst, ":POWer:GAIN 1", console_print_func):
                debug_print("Failed to set :POWer:GAIN 1.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
            if not write_safe(inst, ":POWer:HSENsitive ON", console_print_func):
                debug_print("Failed to set :POWer:HSENsitive ON.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
            if console_print_func:
                console_print_func("✅ High sensitivity turned ON.")
            debug_print("High sensitivity turned ON.", file=current_file, function=current_function, console_print_func=console_print_func)
        else:
            if not write_safe(inst, ":POWer:HSENsitive OFF", console_print_func):
                debug_print("Failed to set :POWer:HSENsitive OFF.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
            if not write_safe(inst, ":POWer:ATTenuation 10", console_print_func):
                debug_print("Failed to set :POWer:ATTenuation 10.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
            # Note: The original code re-set RLEVel here, preserving that behavior
            if not write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM", console_print_func):
                debug_print(f"Failed to re-set reference level to {ref_level_dbm} dBm after high sensitivity config.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
            if console_print_func:
                console_print_func(f"✅ Set reference level to {ref_level_dbm} dBm.")
                console_print_func("✅ High sensitivity turned OFF.")
            debug_print("High sensitivity turned OFF.", file=current_file, function=current_function, console_print_func=console_print_func)
        
        # Configure Trace Modes
        if not write_safe(inst, ":TRAC1:MODE WRITe", console_print_func):
            debug_print("Failed to set :TRAC1:MODE WRITe.", file=current_file, function=current_function, console_print_func=console_print_func)
            return False
        if console_print_func:
            console_print_func(f"✅ Trace 1 sent to write")
        debug_print("Trace 1 set to WRITE.", file=current_file, function=current_function, console_print_func=console_print_func)

        if not write_safe(inst, ":TRAC2:MODE MAXHold", console_print_func):
            debug_print("Failed to set :TRAC2:MODE MAXHold.", file=current_file, function=current_function, console_print_func=console_print_func)
            return False
        if console_print_func:
            console_print_func(f"✅ Trace 2 sent to MAX HOLD")
        debug_print("Trace 2 set to MAX HOLD.", file=current_file, function=current_function, console_print_func=console_print_func)

        if not write_safe(inst, ":TRAC3:MODE MINHold", console_print_func):
            debug_print("Failed to set :TRAC3:MODE MINHold.", file=current_file, function=current_function, console_print_func=console_print_func)
            return False
        if console_print_func:
            console_print_func(f"✅ Trace 3 sent to Min Hold")
        debug_print("Trace 3 set to MIN HOLD.", file=current_file, function=current_function, console_print_func=console_print_func)
        
        # Display scale is always LOGarithmic
        if not write_safe(inst, ":DISPlay:WINDow:TRACe:Y:SCALe:SPACing LOGarithmic", console_print_func):
            debug_print("Failed to set :DISPlay:WINDow:TRACe:Y:SCALe:SPACing LOGarithmic.", file=current_file, function=current_function, console_print_func=console_print_func)
            return False
        if console_print_func:
            console_print_func("✅ Display scale set to LOGarithmic (always).")
        debug_print("Display scale set to LOGarithmic.", file=current_file, function=current_function, console_print_func=console_print_func)
        
        # Set VBW and Sweep Time to AUTO
        if not write_safe(inst, ":SENSe:BANDwidth:VIDeo:AUTO ON", console_print_func):
            debug_print("Failed to set :SENSe:BANDwidth:VIDeo:AUTO ON.", file=current_file, function=current_function, console_print_func=console_print_func)
            return False
        if not write_safe(inst, ":SENSe:SWEep:TIME:AUTO ON", console_print_func):
            debug_print("Failed to set :SENSe:SWEep:TIME:AUTO ON.", file=current_file, function=current_function, console_print_func=console_print_func)
            return False
        if console_print_func:
            console_print_func("✅ VBW and Sweep time set to AUTO.")
        debug_print("VBW and Sweep time set to AUTO.", file=current_file, function=current_function, console_print_func=console_print_func)

        # Set trace data format
        if model_match == "N9340B":
            if not write_safe(inst, ":TRACe:FORMat:DATA ASCii", console_print_func):
                debug_print("Failed to set :TRACe:FORMat:DATA ASCii for N9340B.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
        else:
            if not write_safe(inst, ":FORMat:DATA ASCii", console_print_func):
                debug_print("Failed to set :FORMat:DATA ASCii for non-N9340B.", file=current_file, function=current_function, console_print_func=console_print_func)
                return False
        if console_print_func:
            console_print_func("✅ Set trace data format to ASCII for data transfer.")
        debug_print("Trace data format set to ASCII.", file=current_file, function=current_function, console_print_func=console_print_func)
       
        if console_print_func:
            console_print_func("🎉 Instrument initialized successfully with desired settings.")
        debug_print("Instrument initialized successfully.", file=current_file, function=current_function, console_print_func=console_print_func)
        return True
    except pyvisa.errors.VisaIOError as e:
        if console_print_func:
            console_print_func(f"🛑 Failed to initialize instrument with desired settings: {e}")
        debug_print(f"VISA Error during instrument initialization: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False
    except Exception as e:
        if console_print_func:
            console_print_func(f"❌ An unexpected error occurred during instrument initialization: {e}")
        debug_print(f"Unexpected error during instrument initialization: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False



def _query_settings_display(self):
        """
        Queries the current settings from the instrument and updates the display variables.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Querying current instrument settings for display...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if not self.app_instance.inst:
            self.console_print_func("⚠️ Warning: No instrument connected. Cannot query settings for display.")
            debug_print("No instrument connected. Cannot query settings for display.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            self._clear_settings_display()
            return False

        try:
            # Query Center Frequency
            center_freq_str = query_safe(self.app_instance.inst, ":SENSe:FREQuency:CENTer?", self.console_print_func)
            self.current_center_freq_var.set(f"{float(center_freq_str) / MHZ_TO_HZ:.3f}" if center_freq_str else "N/A")

            # Query Span
            span_str = query_safe(self.app_instance.inst, ":SENSe:FREQuency:SPAN?", self.console_print_func)
            self.current_span_var.set(f"{float(span_str) / MHZ_TO_HZ:.3f}" if span_str else "N/A")

            # Query RBW
            rbw_str = query_safe(self.app_instance.inst, ":SENSe:BANDwidth:RESolution?", self.console_print_func)
            self.current_rbw_var.set(f"{float(rbw_str):.0f}" if rbw_str else "N/A")

            # Query Reference Level
            ref_level_str = query_safe(self.app_instance.inst, ":DISPlay:WINDow:TRACe:Y:RLEVel?", self.console_print_func)
            self.current_ref_level_var.set(f"{float(ref_level_str):.1f}" if ref_level_str else "N/A")

            # Query High Sensitivity / Preamp state
            atten_auto_query = query_safe(self.app_instance.inst, ":INPut:ATTenuation:AUTO?", self.console_print_func)
            gain_state_query = query_safe(self.app_instance.inst, ":INPut:GAIN:STATe?", self.console_print_func)
            high_sensitivity_status = "Enabled" if (atten_auto_query and "OFF" in atten_auto_query.upper() and \
                                                   gain_state_query and "ON" in gain_state_query.upper()) else "Disabled"
            self.current_high_sensitivity_var.set(high_sensitivity_status)

            self.console_print_func("✅ Current instrument settings displayed.")
            debug_print("Current instrument settings displayed.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return True
        except Exception as e:
            self.console_print_func(f"❌ Error querying instrument settings for display: {e}")
            debug_print(f"Error querying instrument settings for display: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            self._clear_settings_display()
            return False