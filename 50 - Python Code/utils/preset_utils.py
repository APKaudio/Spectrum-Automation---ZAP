# utils/preset_utils.py
#
# This module provides utility functions for interacting with instrument presets,
# including querying available presets from the device and loading selected presets.
# It abstracts the low-level SCPI commands for preset management.
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
import inspect
import os
from datetime import datetime

# Import necessary functions from instrument_control and frequency_bands
from utils.instrument_control import debug_print, query_safe, write_safe, query_current_instrument_settings # Re-added query_current_instrument_settings
from ref.frequency_bands import MHZ_TO_HZ

def query_device_presets(inst, console_print_func=None):
    """
    Queries the connected instrument for a list of preset files stored in its
    internal "C:\\PRESETS\\" directory. This allows the GUI to display available
    presets for loading.

    Inputs:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        console_print_func (function, optional): Function to use for console output.
    Process:
        1. Checks if `inst` is connected.
        2. Sends the SCPI command `:MMEMory:CATalog? "C:\\\\PRESETS\\\\"` to list directory contents.
        3. Parses the comma-separated response string to extract file names and types.
        4. Filters for files with the "STA" type (state files) and ending with ".STA".
        5. Sorts the found preset names alphabetically.
        6. Prints the number of found presets or a message if none are found.
        7. Handles `pyvisa.errors.VisaIOError` and general `Exception`.
    Outputs:
        list: A sorted list of `.STA` preset file names (e.g., `['MY_PRESET.STA', 'DEFAULT.STA']`).
              Returns an empty list on failure or if no presets are found.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not inst:
        debug_print("Not connected to instrument, cannot query device presets.", file=current_file, function=current_function, console_print_func=console_print_func)
        if console_print_func:
            console_print_func("⚠️ Warning: No instrument connected. Cannot query presets.")
        return []

    if console_print_func:
        console_print_func(f"Instrument instance passed to query_device_presets (inside preset_utils): {inst}")
    debug_print(f"Instrument instance (inside preset_utils): {inst}", file=current_file, function=current_function, console_print_func=console_print_func)

    debug_print("Querying device preset files from C:\\PRESETS\\...", file=current_file, function=current_function, console_print_func=console_print_func)
    preset_files = []
    response = None # Initialize response to None

    try:
        # First attempt with specific path
        console_print_func(f"Attempting instrument query for presets: ':MMEMory:CATalog? \"C:\\\\PRESETS\\\\\"'")
        debug_print(f"Attempting query: ':MMEMory:CATalog? \"C:\\\\PRESETS\\\\\"'", file=current_file, function=current_function, console_print_func=console_print_func)
        response = query_safe(inst, ':MMEMory:CATalog? "C:\\\\PRESETS\\\\"', console_print_func)
        console_print_func(f"Raw response from instrument (with path): {response!r}")
        debug_print(f"Response from instrument (with path): {response!r}", file=current_file, function=current_function, console_print_func=console_print_func)

        if response is None or "Error" in (response or ""): # Check for None or explicit error in response
            console_print_func("Specific preset path query failed or returned error. Trying generic catalog query.")
            debug_print("Specific preset path query failed. Trying generic catalog.", file=current_file, function=current_function, console_print_func=console_print_func)
            response = query_safe(inst, ':MMEMory:CATalog?', console_print_func) # Try generic catalog
            console_print_func(f"Raw response from instrument (generic): {response!r}")
            debug_print(f"Response from instrument (generic): {response!r}", file=current_file, function=current_function, console_print_func=console_print_func)


        if response is None:
            debug_print("No response received for preset catalog query after all attempts.", file=current_file, function=current_function, console_print_func=console_print_func)
            if console_print_func:
                console_print_func("ℹ️ Info: No response received for preset catalog query.")
            return []

        parts = response.split(',')
        # The first three parts are typically header information (e.g., "C:", "DIR", "SIZE")
        # The actual item listings start after the first 3 parts, in groups of 4 (name, type, size, date)
        if len(parts) < 3: # Minimum expected parts for a valid response
            debug_print(f"Unexpected response format for preset catalog: {response}", file=current_file, function=current_function, console_print_func=console_print_func)
            if console_print_func:
                console_print_func(f"❌ Error: Unexpected response format for preset catalog: {response}")
            return []

        # Iterate through the parts, assuming chunks of 4 for each file/directory entry
        for i in range(3, len(parts), 4):
            if i + 3 < len(parts): # Ensure there are enough parts for a complete entry
                name = parts[i].strip().strip('"') # Strip quotes from name
                item_type = parts[i+1].strip().strip('"') # Strip quotes from type
                # We are interested in files with type "STA" (state files) and ending with ".STA"
                if item_type.upper() == "STA" and name.upper().endswith(".STA"):
                    preset_files.append(name)
            else:
                debug_print(f"Warning: Incomplete item entry found at index {i} in preset catalog response.", file=current_file, function=current_function, console_print_func=console_print_func)
                break # Stop if an incomplete entry is found

        if preset_files:
            debug_print(f"Found {len(preset_files)} '.STA' preset files.", file=current_file, function=current_function, console_print_func=console_print_func)
            if console_print_func:
                console_print_func(f"✅ Found {len(preset_files)} '.STA' preset files on the instrument.")
        else:
            debug_print("No '.STA' preset files found in C:\\PRESETS\\ or generic catalog.", file=current_file, function=current_function, console_print_func=console_print_func)
            if console_print_func:
                console_print_func("ℹ️ Info: No '.STA' preset files found on the instrument.")
        return sorted(preset_files) # Return sorted list
    except pyvisa.errors.VisaIOError as e:
        debug_print(f"🛑 VISA Error querying device presets: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        if console_print_func:
            console_print_func(f"🛑 VISA Error querying device presets: {e}")
        return []
    except Exception as e:
        debug_print(f"❌ An unexpected error occurred while querying presets: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        if console_print_func:
            console_print_func(f"❌ An unexpected error occurred while querying presets: {e}")
        return []

def load_selected_preset(inst, selected_preset_name, console_print_func=None):
    """
    Loads the selected preset file onto the instrument.
    This function sends the SCPI command to instruct the spectrum analyzer
    to load a previously saved state file (`.STA`). After loading, it
    queries and prints the instrument's current settings to confirm the change.

    Inputs:
        inst (pyvisa.resources.Resource): The connected VISA instrument object.
        selected_preset_name (str): The name of the preset file to load (e.g., "MY_PRESET.STA").
        console_print_func (function, optional): Function to use for console output.
    Process:
        1. Checks if `inst` is connected.
        2. Constructs the full path to the preset file (e.g., `C:\\PRESETS\\MY_PRESET.STA`).
        3. Sends the SCPI command `:MMEMory:LOAD STA,"{preset_path}"` using `write_safe`.
        4. If loading is successful, calls `query_current_instrument_settings` to display
           the instrument's new configuration and returns its values.
        5. Prints status messages.
        6. Handles general `Exception` during the loading process.
    Outputs:
        tuple: (bool, center_freq_hz, span_hz, rbw_hz). True if the preset is loaded successfully;
               False otherwise. center_freq_hz, span_hz, and rbw_hz are the queried values or None.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if not inst:
        debug_print("Not connected to instrument, cannot load preset.", file=current_file, function=current_function, console_print_func=console_print_func)
        if console_print_func:
            console_print_func("⚠️ Warning: No instrument connected. Cannot load preset.")
        return False, None, None, None

    preset_path = f"C:\\\\PRESETS\\\\{selected_preset_name}"
    command = f':MMEMory:LOAD STA,"{preset_path}"'
    
    if console_print_func:
        console_print_func(f"\nAttempting to load preset: {selected_preset_name}")
    debug_print(f"Attempting to load preset: {selected_preset_name}", file=current_file, function=current_function, console_print_func=console_print_func)

    try:
        if write_safe(inst, command, console_print_func):
            if console_print_func:
                console_print_func(f"✅ Preset '{selected_preset_name}' loaded successfully.")
            debug_print(f"Preset '{selected_preset_name}' loaded successfully.", file=current_file, function=current_function, console_print_func=console_print_func)
            
            # Query and display current instrument settings after loading preset
            # MHZ_TO_HZ is imported at the module level in preset_utils.py
            center_freq, span, rbw = query_current_instrument_settings(inst, MHZ_TO_HZ, console_print_func)
            return True, center_freq, span, rbw
        else:
            debug_print(f"Failed to load preset '{selected_preset_name}'.", file=current_file, function=current_function, console_print_func=console_print_func)
            if console_print_func:
                console_print_func(f"❌ Error: Failed to load preset '{selected_preset_name}'.")
            return False, None, None, None
    except Exception as e:
        debug_print(f"An unexpected error occurred while loading preset: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        if console_print_func:
            console_print_func(f"❌ An unexpected error occurred while loading preset: {e}")
        return False, None, None, None


def query_device_presets_logic(app_instance, console_print_func):
    """
    Queries the connected instrument for a list of preset files.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Calling query_device_presets_logic...", file=current_file, function=current_function, console_print_func=console_print_func)
    
    # Add debug print for app_instance.inst before calling query_device_presets
    console_print_func(f"Instrument instance in query_device_presets_logic: {app_instance.inst}")
    debug_print(f"Instrument instance in query_device_presets_logic: {app_instance.inst}", file=current_file, function=current_function, console_print_func=console_print_func)

    # Directly call the query_device_presets function within this module
    presets = query_device_presets(app_instance.inst, console_print_func)
    
    if presets is not None:
        if hasattr(app_instance, 'preset_files_tab'):
            app_instance.preset_files_tab.populate_preset_buttons(presets, "device")
        console_print_func(f"✅ Queried {len(presets)} presets from device.")
        debug_print(f"Queried {len(presets)} presets from device.", file=current_file, function=current_function, console_print_func=console_print_func)
        return presets # Return the list of presets
    else:
        console_print_func("❌ Failed to query presets from device.")
        debug_print("Failed to query presets from device.", file=current_file, function=current_function, console_print_func=console_print_func)
        return None # Return None on failure
