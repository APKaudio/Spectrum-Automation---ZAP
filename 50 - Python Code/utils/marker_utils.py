# utils/marker_utils.py
import inspect
from utils.instrument_control import debug_print, write_safe, query_safe
from utils.frequency_bands import MHZ_TO_HZ

# Constants for Span Options (used in MarkersDisplayTab)
SPAN_OPTIONS = {
    "Full Span": 100 * MHZ_TO_HZ, # This would typically be a special value for full span
    "Normal": 10 * MHZ_TO_HZ, # Example: 10 MHz
    "Zoom 1": 1 * MHZ_TO_HZ,  # Example: 1 MHz
    "Zoom 2": 100 * 1000,    # Example: 100 KHz
    "Zoom 3": 10 * 1000,     # Example: 10 KHz
}

def set_span_logic(inst, span_hz, center_freq_hz, live_mode, max_hold_mode, min_hold_mode, console_print_func):
    """
    Sets the instrument's span, center frequency (if provided), and trace modes.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Applying span: {span_hz} Hz, center_freq: {center_freq_hz} Hz, Live: {live_mode}, MaxHold: {max_hold_mode}, MinHold: {min_hold_mode}", file=current_file, function=current_function, console_print_func=console_print_func)

    if not inst:
        console_print_func("⚠️ Warning: Instrument not connected. Cannot set span or trace mode.")
        debug_print("Instrument not connected for set_span_logic.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    success = True

    # Set Center Frequency if provided
    if center_freq_hz is not None:
        if not write_safe(inst, f":SENSe:FREQuency:CENTer {center_freq_hz}", console_print_func):
            success = False
            console_print_func(f"❌ Failed to set center frequency to {center_freq_hz / MHZ_TO_HZ:.3f} MHz.")

    # Set Span
    if span_hz == 0.0: # Special case for "Full Span"
        if not write_safe(inst, ":SENSe:FREQuency:SPAN MAX", console_print_func): # Use MAX for full span
            success = False
            console_print_func("❌ Failed to set Full Span.")
    else:
        if not write_safe(inst, f":SENSe:FREQuency:SPAN {span_hz}", console_print_func):
            success = False
            console_print_func(f"❌ Failed to set span to {span_hz / MHZ_TO_HZ:.3f} MHz.")

    # Apply Trace Modes
    # Ensure only the selected mode is active, and others are blanked.
    if live_mode:
        if not write_safe(inst, ":TRAC1:MODE WRITe", console_print_func): success = False
        console_print_func("✅ Trace 1 set to Live (WRITe).")
    else:
        if not write_safe(inst, ":TRAC1:MODE BLANK", console_print_func): success = False
        console_print_func("ℹ️ Trace 1 set to BLANK.")

    if max_hold_mode:
        if not write_safe(inst, ":TRAC2:MODE MAXHold", console_print_func): success = False
        console_print_func("✅ Trace 2 set to Max Hold.")
    else:
        if not write_safe(inst, ":TRAC2:MODE BLANK", console_print_func): success = False
        console_print_func("ℹ️ Trace 2 set to BLANK.")

    if min_hold_mode:
        if not write_safe(inst, ":TRAC3:MODE MINHold", console_print_func): success = False
        console_print_func("✅ Trace 3 set to Min Hold.")
    else:
        if not write_safe(inst, ":TRAC3:MODE BLANK", console_print_func): success = False
        console_print_func("ℹ️ Trace 3 set to BLANK.")


    if success:
        console_print_func("✅ Span and trace mode applied.")
        debug_print("Span and trace mode applied successfully.", file=current_file, function=current_function, console_print_func=console_print_func)
    else:
        console_print_func("❌ Failed to apply span or trace mode.")
        debug_print("Failed to apply span or trace mode.", file=current_file, function=current_function, console_print_func=console_print_func)
    return success


def set_marker_and_trace_modes_logic(app_instance, frequency_hz, marker_name, console_print_func):
    """
    Sets a marker on the instrument to the specified frequency and enables it.
    Also ensures trace mode is set appropriately (e.g., Clear Write).
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Setting marker to {frequency_hz} Hz for '{marker_name}'...", file=current_file, function=current_function, console_print_func=console_print_func)

    inst = app_instance.inst
    if not inst:
        console_print_func("⚠️ Warning: Instrument not connected. Cannot set marker.")
        debug_print("Instrument not connected for set_marker_and_trace_modes_logic.", file=current_file, function=current_function, console_print_func=console_print_func)
        return False

    success = True

    try:
        # Ensure Marker 1 is ON
        if not write_safe(inst, ":CALCulate:MARKer1:STATe ON", console_print_func): success = False
        # Set Marker 1 to the specified frequency
        if not write_safe(inst, f":CALCulate:MARKer1:X {frequency_hz}", console_print_func): success = False
        
        # Query the Y value of the marker
        marker_y_value_str = query_safe(inst, ":CALCulate:MARKer1:Y?", console_print_func)
        marker_y_value = float(marker_y_value_str) if marker_y_value_str else "N/A"

        if success:
            console_print_func(f"✅ Marker set to {frequency_hz / MHZ_TO_HZ:.3f} MHz (Amplitude: {marker_y_value} dBm).")
            debug_print(f"Marker 1 set to {frequency_hz} Hz, amplitude {marker_y_value}.", file=current_file, function=current_function, console_print_func=console_print_func)
        else:
            console_print_func(f"❌ Failed to set marker for {marker_name}.")
            debug_print(f"Failed to set marker for {marker_name}.", file=current_file, function=current_function, console_print_func=console_print_func)
        return success
    except Exception as e:
        console_print_func(f"❌ Error setting marker: {e}")
        debug_print(f"Error setting marker: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        return False
