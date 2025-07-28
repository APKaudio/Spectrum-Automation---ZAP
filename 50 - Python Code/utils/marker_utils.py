# utils/marker_utils.py
#
# This module provides utility functions specifically for managing marker-related
# logic, such as defining span options and sending span-related SCPI commands
# to the instrument. This helps in keeping the MarkersDisplayTab cleaner and
# more focused on UI management.
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
import inspect
from utils.instrument_control import debug_print, write_safe, query_safe # Import query_safe
from utils.frequency_bands import MHZ_TO_HZ # Import MHZ_TO_HZ for conversion


# Define the span options with descriptive names and values in Hz
# The values are in Hz, but the display text will be in MHz
SPAN_OPTIONS = {
    "Ultra Wide": 10_000_000, # 10 MHz
    "Wide": 5_000_000,     # 5 MHz
    "Normal": 1_000_000,   # 1 MHz
    "Tight": 750_000,      # 0.750 MHz (750 KHz)
    "Microscope": 500_000  # 0.500 MHz (500 KHz)
}

def set_span_logic(inst, span_hz, center_freq_hz=None, live_mode=False, max_hold_mode=False, min_hold_mode=False, console_print_func=None):
    """
    Sets the instrument's span, optionally its center frequency, and configures trace modes.
    Directly sends commands without querying current state to ensure desired settings are applied.

    Inputs:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        span_hz (float): The desired span in Hz.
        center_freq_hz (float, optional): The desired center frequency in Hz. If None,
                                          the center frequency is not changed.
        live_mode (bool): True to set TRAC1 to WRITE.
        max_hold_mode (bool): True to set TRAC2 to MAXHold.
        min_hold_mode (bool): True to set TRAC3 to MINHold.
        console_print_func (function, optional): Function to use for console output.
    Returns:
        bool: True if commands were sent successfully, False otherwise.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Initiating set_span_logic: Desired Span={span_hz} Hz, Desired Center Freq={center_freq_hz} Hz, Live={live_mode}, MaxHold={max_hold_mode}, MinHold={min_hold_mode}", file=current_file, function=current_function, console_print_func=console_print_func)

    if not inst:
        if console_print_func:
            console_print_func("⚠️ Warning: No instrument connected. Cannot set span/frequency/trace modes.")
        debug_print("No instrument connected. Cannot set span/frequency/trace modes.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    overall_success = True

    try:
        # --- Set Center Frequency (if provided) ---
        if center_freq_hz is not None:
            formatted_freq_hz = int(center_freq_hz) if center_freq_hz == int(center_freq_hz) else center_freq_hz
            debug_print(f"Sending center frequency command: :SENSe:FREQuency:CENTer {formatted_freq_hz}", file=current_file, function=current_function, console_print_func=console_print_func)
            if not write_safe(inst, f":SENSe:FREQuency:CENTer {formatted_freq_hz}", console_print_func):
                if console_print_func:
                    console_print_func(f"❌ Error: Failed to send center frequency command: :SENSe:FREQuency:CENTer {formatted_freq_hz}")
                debug_print(f"Failed to send center frequency command.", file=current_file, function=current_function, console_print_func=console_print_func)
                overall_success = False
            else:
                if console_print_func:
                    display_freq_mhz = int(center_freq_hz / MHZ_TO_HZ) if (center_freq_hz / MHZ_TO_HZ) == int(center_freq_hz / MHZ_TO_HZ) else f"{center_freq_hz / MHZ_TO_HZ:.3f}"
                    console_print_func(f"✅ Instrument center frequency set to {display_freq_mhz} MHz.")

        # --- Set Span ---
        debug_print(f"Sending span command: :SENSe:FREQuency:SPAN {span_hz}", file=current_file, function=current_function, console_print_func=console_print_func)
        if not write_safe(inst, f":SENSe:FREQuency:SPAN {span_hz}", console_print_func):
            if console_print_func:
                console_print_func(f"❌ Error: Failed to send span command: :SENSe:FREQuency:SPAN {span_hz}")
            debug_print(f"Failed to send span command.", file=current_file, function=current_function, console_print_func=console_print_func)
            overall_success = False
        else:
            if console_print_func:
                console_print_func(f"✅ Instrument span set to {span_hz / MHZ_TO_HZ:.3f} MHz.")

        # --- Set Trace Modes ---
        # Directly set trace modes without querying current state
        trace_commands_to_send = []
        
        # TRAC1: Live Mode (WRITe) or BLANK
        desired_trac1_mode = "WRITe" if live_mode else "BLANK"
        trace_commands_to_send.append(f":TRAC1:MODE {desired_trac1_mode}")
        debug_print(f"Adding TRAC1 command: :TRAC1:MODE {desired_trac1_mode}", file=current_file, function=current_function, console_print_func=console_print_func)

        # TRAC2: Max Hold or BLANK
        desired_trac2_mode = "MAXHold" if max_hold_mode else "BLANK"
        trace_commands_to_send.append(f":TRAC2:MODE {desired_trac2_mode}")
        debug_print(f"Adding TRAC2 command: :TRAC2:MODE {desired_trac2_mode}", file=current_file, function=current_function, console_print_func=console_print_func)

        # TRAC3: Min Hold or BLANK
        desired_trac3_mode = "MINHold" if min_hold_mode else "BLANK"
        trace_commands_to_send.append(f":TRAC3:MODE {desired_trac3_mode}")
        debug_print(f"Adding TRAC3 command: :TRAC3:MODE {desired_trac3_mode}", file=current_file, function=current_function, console_print_func=console_print_func)
        
        # TRAC4: Always BLANK as per requirement
        desired_trac4_mode = "BLANK"
        trace_commands_to_send.append(f":TRAC4:MODE {desired_trac4_mode}")
        debug_print(f"Adding TRAC4 command: :TRAC4:MODE {desired_trac4_mode}", file=current_file, function=current_function, console_print_func=console_print_func)


        if trace_commands_to_send:
            combined_trace_command = ";".join(trace_commands_to_send)
            debug_print(f"Sending combined trace mode command: {combined_trace_command}", file=current_file, function=current_function, console_print_func=console_print_func)
            if not write_safe(inst, combined_trace_command, console_print_func):
                if console_print_func:
                    console_print_func(f"❌ Error: Failed to send trace mode commands.")
                debug_print(f"Failed to send trace mode commands.", file=current_file, function=current_function, console_print_func=console_print_func)
                overall_success = False
            else:
                if console_print_func:
                    console_print_func(f"✅ Trace modes updated: TRAC1:{'WRITE' if live_mode else 'BLANK'}, TRAC2:{'MAXHOLD' if max_hold_mode else 'BLANK'}, TRAC3:{'MINHOLD' if min_hold_mode else 'BLANK'}, TRAC4:BLANK.")
        else:
            debug_print("All trace modes already configured as desired. No trace commands sent.", file=current_file, function=current_function, console_print_func=console_print_func)
            if console_print_func:
                console_print_func("ℹ️ Info: All trace modes already configured as desired.")
        
        return overall_success

    except pyvisa.errors.VisaIOError as e:
        if console_print_func:
            console_print_func(f"❌ VISA error while setting span/frequency/trace modes: {e}")
        debug_print(f"VISA Error setting span/frequency/trace modes in marker_utils: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False
    except Exception as e:
        if console_print_func:
            console_print_func(f"❌ An unexpected error occurred while setting span/frequency/trace modes: {e}")
        debug_print(f"An unexpected error occurred while setting span/frequency/trace modes in marker_utils: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False


def set_marker_and_trace_modes_logic(app_instance, marker_frequency_hz, marker_name, console_print_func):
    """
    Sets a marker at the specified frequency.
    This function is kept separate from set_span_logic as it specifically handles marker setup.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Setting marker at {marker_frequency_hz} Hz for '{marker_name}'...", file=current_file, function=current_function, console_print_func=console_print_func)

    if not app_instance.inst:
        console_print_func("⚠️ Warning: No instrument connected. Cannot set marker.")
        debug_print("No instrument connected for setting marker.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    try:
        # Set marker 1 to the specified frequency
        # Format frequency to integer if it's a whole number, otherwise keep decimal for precision
        formatted_marker_freq_hz = int(marker_frequency_hz) if marker_frequency_hz == int(marker_frequency_hz) else marker_frequency_hz
        if not write_safe(app_instance.inst, f":CALCulate:MARKer1:X {formatted_marker_freq_hz}", console_print_func): return False
        debug_print(f"Sent: :CALCulate:MARKer1:X {formatted_marker_freq_hz}", file=current_file, function=current_function, console_print_func=console_print_func)

        # Activate marker 1 (always turn on when setting)
        if not write_safe(app_instance.inst, ":CALCulate:MARKer1:STATe ON", console_print_func): return False
        debug_print("Sent: :CALCulate:MARKer1:STATe ON", file=current_file, function=current_function, console_print_func=console_print_func)

        # Set marker 1 to peak (always set to peak when setting a marker)
        if not write_safe(app_instance.inst, ":CALCulate:MARKer1:MAXimum:PEAK", console_print_func): return False
        debug_print("Sent: :CALCulate:MARKer1:MAXimum:PEAK", file=current_file, function=current_function, console_print_func=console_print_func)

        # Display without decimal if it's a whole number, otherwise with .3f
        display_marker_freq_mhz = int(marker_frequency_hz / MHZ_TO_HZ) if (marker_frequency_hz / MHZ_TO_HZ) == int(marker_frequency_hz / MHZ_TO_HZ) else f"{marker_frequency_hz / MHZ_TO_HZ:.3f}"
        console_print_func(f"✅ Marker '{marker_name}' set at {display_marker_freq_mhz} MHz.")
        return True
    except pyvisa.errors.VisaIOError as e:
        console_print_func(f"❌ VISA error while setting marker: {e}")
        debug_print(f"VISA Error setting marker: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False
    except Exception as e:
        console_print_func(f"❌ An unexpected error occurred while setting marker: {e}")
        debug_print(f"An unexpected error occurred while setting marker: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

