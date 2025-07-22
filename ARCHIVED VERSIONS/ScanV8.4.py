import tkinter as tk
from tkinter import messagebox
import pyvisa
import time
import argparse
import struct
import numpy as np
import os
import csv
from datetime import datetime
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

import sys
import subprocess

# Define constants for better readability and easier modification
MHZ_TO_HZ = 1_000_000 # Conversion factor from MHz to Hz

# Updated wait time variable and its usage for the continuous loop
DEFAULT_RBW_STEP_SIZE_HZ = 10000 # 10 kHz RBW resolution desired per data point
DEFAULT_CYCLE_WAIT_TIME_SECONDS = 30 # 5 minutes wait (300 seconds) between full scan cycles
DEFAULT_MAXHOLD_TIME_SECONDS = 5 # Default max hold time for the new argument

# Define the frequency bands to *SCAN* (User's specified bands for instrument operation)
# This list will be used by the scan_bands function.
SCAN_BAND_RANGES = [
    {"Band Name": "Low VHF+FM", "Start MHz": 50.000, "Stop MHz": 110.000},
    {"Band Name": "High VHF+216", "Start MHz": 170.000, "Stop MHz": 220.000},
    {"Band Name": "UHF -1", "Start MHz": 400.000, "Stop MHz": 700.000},
    {"Band Name": "UHF -2", "Start MHz": 700.000, "Stop MHz": 900.000},
    {"Band Name": "900 ISM-STL", "Start MHz": 900.000, "Stop MHz": 970.000},
    {"Band Name": "AFTRCC-1", "Start MHz": 1430.000, "Stop MHz": 1540.000},
    {"Band Name": "DECT-ALL", "Start MHz": 1880.000, "Stop MHz": 2000.000},
    {"Band Name": "2 GHz Cams", "Start MHz": 2000.000, "Stop MHz": 2390.000},
]


frequency_TV_Channel_bands_full_list = [
    [54, 60,  'TV-CH2 - VHF-L'],
    [60, 66,  'TV-CH3 - VHF-L'],
    [66, 72,  'TV-CH4 - VHF-L'],
    [76, 82,  'TV-CH5 - VHF-L'],
    [82, 88,  'TV-CH6 - VHF-L'],
    [174, 180, 'TV-CH7 - VHF-H'],
    [180, 186, 'TV-CH8 - VHF-H'],
    [186, 192, 'TV-CH9 - VHF-H'],
    [192, 198, 'TV-CH10 - VHF-H'],
    [198, 204, 'TV-CH11 - VHF-H'],
    [204, 210, 'TV-CH12 - VHF-H'],
    [210, 216, 'TV-CH13 - VHF-H'],
    [470, 476, 'TV-CH14 - UHF'],
    [476, 482, 'TV-CH15 - UHF'],
    [482, 488, 'TV-CH16 - UHF'],
    [488, 494, 'TV-CH17 - UHF'],
    [494, 500, 'TV-CH18 - UHF'],
    [500, 506, 'TV-CH19 - UHF'],
    [506, 512, 'TV-CH20 - UHF'],
    [512, 518, 'TV-CH21 - UHF'],
    [518, 524, 'TV-CH22 - UHF'],
    [524, 530, 'TV-CH23 - UHF'],
    [530, 536, 'TV-CH24 - UHF'],
    [536, 542, 'TV-CH25 - UHF'],
    [542, 548, 'TV-CH26 - UHF'],
    [548, 554, 'TV-CH27 - UHF'],
    [554, 560, 'TV-CH28 - UHF'],
    [560, 566, 'TV-CH29 - UHF'],
    [566, 572, 'TV-CH30 - UHF'],
    [572, 578, 'TV-CH31 - UHF'],
    [578, 584, 'TV-CH32 - UHF'],
    [584, 590, 'TV-CH33 - UHF'],
    [590, 596, 'TV-CH34 - UHF'],
    [596, 602, 'TV-CH35 - UHF'],
    [602, 608, 'TV-CH36 - UHF'],
    [608, 614, 'TV-CH37 - UHF'],
    [614, 620, 'TV-CH38 - UHF'],
    [620, 626, 'TV-CH39 - UHF'],
    [626, 632, 'TV-CH40 - UHF'],
    [632, 638, 'TV-CH41 - UHF'],
    [638, 644, 'TV-CH42 - UHF'],
    [644, 650, 'TV-CH43 - UHF'],
    [650, 656, 'TV-CH44 - UHF'],
    [656, 662, 'TV-CH45 - UHF'],
    [662, 668, 'TV-CH46 - UHF'],
    [668, 674, 'TV-CH47 - UHF'],
    [674, 680, 'TV-CH48 - UHF'],
    [680, 686, 'TV-CH49 - UHF'],
    [686, 692, 'TV-CH50 - UHF'],
    [692, 698, 'TV-CH51 - UHF'],
]

# This list will be dynamically created for plotting purposes only, from frequency_bands_full_list
TV_PLOT_BAND_MARKERS = []
for band_info in frequency_TV_Channel_bands_full_list:
    TV_PLOT_BAND_MARKERS.append({
        "Start MHz": band_info[0],
        "Stop MHz": band_info[1],
        "Band Name": band_info[2].strip() # Use strip() to remove leading/trailing spaces
    })



# Declare the comprehensive frequency_bands array (for PLOTTING MARKERS ONLY)
# This array will be used to define the PLOT_BAND_MARKERS for plotting.
gov_frequency_bands_full_list = [
    [50.0, 54.0, 'AMATEUR'],
    [54.0, 72.0, 'BROADCASTING'],
    [72.0, 73.0, 'FIXED MOBILE'],
    [73.0, 74.6, 'RADIO ASTRONOMY'],
    [74.6, 74.8, 'FIXED MOBILE'],
    [74.8, 75.2, 'AERONAUTICAL RADIONAVIGATION'],
    [75.2, 76.0, 'FIXED MOBILE'],
    [76.0, 108.0, 'BROADCASTING'],
    [108.0, 117.975, 'AERONAUTICAL RADIONAVIGATION'],
    [117.975, 137.0, 'AERONAUTICAL MOBILE (R)'],
    [137.0, 138.0, 'METEOROLOGICAL-SATELLITE (space-to-Earth) MOBILE-SATELLITE (space-to-Earth)'],
    [138.0, 144.0, 'FIXED LAND MOBILE Space research (space-to-Earth)'],
    [144.0, 146.0, 'AMATEUR AMATEUR-SATELLITE'],
    [146.0, 148.0, 'AMATEUR'],
    [148.0, 149.9, 'FIXED LAND MOBILE MOBILE-SATELLITE (Earth-to-space)'],
    [149.9, 150.05, 'MOBILE-SATELLITE (Earth-to-space)'],
    [150.05, 156.4875, 'MOBILE Fixed'],
    [156.4875, 156.5625, 'MARITIME MOBILE (distress and calling via DSC )'],
    [156.5625, 156.7625, 'MOBILE Fixed'],
    [156.7625, 156.7875, 'MARITIME MOBILE MOBILE-SATELLITE (Earth-to-space)'],
    [156.7875, 156.8125, 'MARITIME MOBILE (distress and calling)'],
    [156.8125, 156.8375, 'MARITIME MOBILE MOBILE-SATELLITE (Earth-to-space)'],
    [156.8375, 157.1875, 'MOBILE Fixed'],
    [157.1875, 157.3375, 'MOBILE Fixed Maritime mobile-satellite'],
    [157.3375, 161.7875, 'MOBILE Fixed'],
    [161.7875, 161.9375, 'MOBILE Fixed Maritime mobile-satellite'],
    [161.9375, 161.9625, 'MOBILE Fixed Maritime mobile-satellite (Earth-to-space)'],
    [161.9625, 161.9875, 'AERONAUTICAL MOBILE (OR ) MARITIME MOBILE MOBILE-SATELLITE (Earth-to-space)'],
    [161.9875, 162.0125, 'MOBILE Fixed Maritime mobile-satellite (Earth-to-space)'],
    [162.0125, 162.0375, 'AERONAUTICAL MOBILE (OR ) MARITIME MOBILE MOBILE-SATELLITE (Earth-to-space)'],
    [162.0375, 174.0, 'MOBILE Fixed'],
    [174.0, 216.0, 'BROADCASTING'],
    [216.0, 219.0, 'FIXED MARITIME MOBILE LAND MOBILE'],
    [219.0, 220.0, 'FIXED MARITIME MOBILE LAND MOBILE Amateur'],
    [220.0, 222.0, 'FIXED MOBILE Amateur'],
    [222.0, 225.0, 'AMATEUR'],
    [225.0, 312.0, 'FIXED MOBILE'],
    [312.0, 315.0, 'FIXED MOBILE Mobile-satellite (Earth-to-space)'],
    [315.0, 328.6, 'FIXED MOBILE'],
    [328.6, 335.4, 'AERONAUTICAL RADIONAVIGATION'],
    [335.4, 387.0, 'FIXED MOBILE'],
    [387.0, 390.0, 'FIXED MOBILE Mobile-satellite (space-to-Earth)'],
    [390.0, 399.9, 'FIXED MOBILE'],
    [399.9, 400.05, 'MOBILE-SATELLITE (Earth-to-space)'],
    [400.05, 400.15, 'STANDARD FREQUENCY AND TIME SIGNAL-SATELLITE (400.1 MHz)'],
    [400.15, 401.0, 'METEOROLOGICAL AIDS METEOROLOGICAL-SATELLITE (space-to-Earth)'],
    [401.0, 402.0, 'METEOROLOGICAL AIDS SPACE OPERATION (space-to-Earth)'],
    [402.0, 403.0, 'METEOROLOGICAL AIDS EARTH EXPLORATION-SATELLITE'],
    [403.0, 406.0, 'METEOROLOGICAL AIDS Fixed Mobile except aeronautical mobile'],
    [406.0, 406.1, 'MOBILE-SATELLITE (Earth-to-space)'],
    [406.1, 410.0, 'MOBILE except aeronautical mobile RADIO ASTRONOMY Fixed'],
    [410.0, 414.0, 'MOBILE except aeronautical mobile SPACE RESEARCH (space-to-space) Fixed'],
    [414.0, 415.0, 'FIXED SPACE RESEARCH (space-to-space) Mobile except aeronautical mobile'],
    [415.0, 419.0, 'MOBILE except aeronautical mobile SPACE RESEARCH (space-to-space) Fixed'],
    [419.0, 420.0, 'FIXED SPACE RESEARCH (space-to-space) Mobile except aeronautical mobile'],
    [420.0, 430.0, 'MOBILE except aeronautical mobile Fixed'],
    [430.0, 432.0, 'RADIOLOCATION Amateur'],
    [432.0, 438.0, 'RADIOLOCATION Amateur Earth Exploration-Satellite (active)'],
    [438.0, 450.0, 'RADIOLOCATION Amateur'],
    [450.0, 455.0, 'MOBILE Fixed'],
    [455.0, 456.0, 'FIXED MOBILE MOBILE-SATELLITE (Earth-to-space)'],
    [456.0, 459.0, 'MOBILE Fixed'],
    [459.0, 460.0, 'FIXED MOBILE MOBILE-SATELLITE (Earth-to-space)'],
    [460.0, 470.0, 'MOBILE Fixed'],
    [470.0, 608.0, 'BROADCASTING'],
    [608.0, 614.0, 'RADIO ASTRONOMY Mobile-satellite except aeronautical mobile-satellite (Earth-to-space)'],
    [614.0, 698.0, 'FIXED MOBILE BROADCASTING'],
    [698.0, 806.0, 'FIXED MOBILE BROADCASTING'],
    [806.0, 890.0, 'MOBILE Fixed'],
    [890.0, 902.0, 'FIXED MOBILE except aeronautical mobile Radiolocation'],
    [902.0, 928.0, 'FIXED RADIOLOCATION Amateur Mobile except aeronautical mobile'],
    [928.0, 929.0, 'FIXED MOBILE except aeronautical mobile Radiolocation'],
    [929.0, 932.0, 'MOBILE except aeronautical mobile Fixed Radiolocation'],
    [932.0, 932.5, 'FIXED MOBILE except aeronautical mobile Radiolocation'],
    [932.5, 935.0, 'FIXED Mobile except aeronautical mobile Radiolocation'],
    [935.0, 941.0, 'MOBILE except aeronautical mobile Fixed Radiolocation'],
    [941.0, 941.5, 'FIXED MOBILE except aeronautical mobile Radiolocation'],
    [941.5, 942.0, 'FIXED Mobile except aeronautical mobile Radiolocation'],
    [942.0, 944.0, 'FIXED Mobile'],
    [944.0, 952.0, 'FIXED MOBILE'],
    [952.0, 956.0, 'FIXED MOBILE'],
    [956.0, 960.0, 'FIXED Mobile'],
    [960.0, 1164.0, 'AERONAUTICAL MOBILE (R) AERONAUTICAL RADIONAVIGATION'],
    [1164.0, 1215.0, 'AERONAUTICAL RADIONAVIGATION RADIONAVIGATION-SATELLITE (space-to-Earth) (space-to-space)'],
    [1215.0, 1240.0, 'EARTH EXPLORATION-SATELLITE (active) RADIOLOCATION RADIONAVIGATION-SATELLITE '],
    [1240.0, 1300.0, 'EARTH EXPLORATION-SATELLITE (active) RADIOLOCATION RADIONAVIGATION-SATELLITE'],
    [1300.0, 1350.0, 'RADIOLOCATION AERONAUTICAL RADIONAVIGATION RADIONAVIGATION-SATELLITE (Earth-to-space)'],
    [1350.0, 1390.0, 'FIXED MOBILE RADIOLOCATION'],
    [1390.0, 1400.0, 'FIXED MOBILE'],
    [1400.0, 1427.0, 'EARTH EXPLORATION-SATELLITE (passive) RADIO ASTRONOMY SPACE RESEARCH (passive)'],
    [1427.0, 1429.0, 'SPACE OPERATION (Earth-to-space) FIXED'],
    [1429.0, 1452.0, 'FIXED MOBILE'],
    [1452.0, 1492.0, 'FIXED MOBILE BROADCASTING'],
    [1492.0, 1525.0, 'FIXED MOBILE'],
    [1525.0, 1530.0, 'MOBILE-SATELLITE (space-to-Earth) Earth Exploration-Satellite Space operation (space-to-Earth)'],
    [1530.0, 1535.0, 'MOBILE-SATELLITE (space-to-Earth) Earth Exploration-Satellite'],
    [1535.0, 1559.0, 'MOBILE-SATELLITE (space-to-Earth)'],
    [1559.0, 1610.0, 'AERONAUTICAL RADIONAVIGATION RADIONAVIGATION-SATELLITE (space-to-Earth) (space-to-space)'],
    [1610.0, 1610.6, 'MOBILE-SATELLITE (Earth-to-space) AERONAUTICAL RADIONAVIGATION'],
    [1610.6, 1613.8, 'MOBILE-SATELLITE (Earth-to-space) RADIO ASTRONOMY AERONAUTICAL RADIONAVIGATION'],
    [1613.8, 1621.35, 'MOBILE-SATELLITE (Earth-to-space) AERONAUTICAL RADIONAVIGATION Mobile-satellite (space-to-Earth)'],
    [1621.35, 1626.5, 'MARITIME MOBILE-SATELLITE (space-to-Earth)'],
    [1626.5, 1660.0, 'MOBILE-SATELLITE (Earth-to-space)'],
    [1660.0, 1660.5, 'MOBILE-SATELLITE (Earth-to-space) RADIO ASTRONOMY'],
    [1660.5, 1668.0, 'RADIO ASTRONOMY SPACE RESEARCH (passive) Fixed'],
    [1668.0, 1668.4, 'RADIO ASTRONOMY SPACE RESEARCH (passive) Fixed'],
    [1668.4, 1670.0, 'METEOROLOGICAL AIDS FIXED RADIO ASTRONOMY'],
    [1670.0, 1675.0, 'METEOROLOGICAL AIDS FIXED METEOROLOGICAL-SATELLITE (space-to-Earth) MOBILE except aeronautical mobile'],
    [1675.0, 1700.0, 'METEOROLOGICAL AIDS METEOROLOGICAL-SATELLITE (space-to-Earth)'],
    [1700.0, 1710.0, 'FIXED METEOROLOGICAL-SATELLITE (space-to-Earth)'],
    [1710.0, 1755.0, 'FIXED MOBILE'],
    [1755.0, 1780.0, 'FIXED MOBILE'],
    [1780.0, 1850.0, 'FIXED Mobile'],
    [1850.0, 2000.0, 'FIXED MOBILE'],
    [2000.0, 2020.0, 'MOBILE MOBILE-SATELLITE (Earth-to-space)'],
    [2020.0, 2025.0, 'FIXED MOBILE'],
    [2025.0, 2110.0, 'EARTH EXPLORATION-SATELLITE (Earth-to-space) (space-to-space)'],
    [2110.0, 2120.0, 'FIXED MOBILE SPACE RESEARCH (deep space) (Earth-to-space)'],
    [2120.0, 2180.0, 'FIXED MOBILE'],
    [2180.0, 2200.0, 'MOBILE MOBILE-SATELLITE (space-to-Earth)'],
    [2200.0, 2290.0, 'EARTH EXPLORATION-SATELLITE (space-to-Earth) (space-to-space)'],
    [2290.0, 2300.0, 'FIXED SPACE RESEARCH (deep space) (Earth-to-space) Mobile'],
    [2300.0, 2450.0, 'FIXED MOBILE RADIOLOCATION Amateur'],
    [2450.0, 2483.5, 'FIXED MOBILE RADIOLOCATION'],
    [2483.5, 2500.0, 'FIXED MOBILE-SATELLITE (space-to-Earth) RADIOLOCATION RADIODETERMINATION-SATELLITE (space-to-Earth)'],
    [2500.0, 2596.0, 'FIXED MOBILE except aeronautical mobile'],
    [2596.0, 2655.0, 'BROADCASTING FIXED MOBILE except aeronautical mobile'],
    [2655.0, 2686.0, 'BROADCASTING FIXED MOBILE except aeronautical mobile Earth Exploration-Satellite (passive) Radio astronomy Space research (passive)'],
    [2686.0, 2690.0, 'FIXED MOBILE except aeronautical mobile Earth Exploration-Satellite (passive) Radio astronomy Space research (passive)'],
    [2690.0, 2700.0, 'EARTH EXPLORATION-SATELLITE (passive) RADIO ASTRONOMY SPACE RESEARCH (passive)'],
    [2700.0, 2900.0, 'AERONAUTICAL RADIONAVIGATION Radiolocation'],
    [2900.0, 3100.0, 'RADIOLOCATION RADIONAVIGATION'],
    [3100.0, 3300.0, 'RADIOLOCATION Earth Exploration-Satellite (active) Space research (active)'],
    [3300.0, 3450.0, 'RADIOLOCATION Amateur'],
]

# This list will be dynamically created for plotting purposes only, from frequency_bands_full_list
GOV_PLOT_BAND_MARKERS = []
for band_info in gov_frequency_bands_full_list:
    GOV_PLOT_BAND_MARKERS.append({
        "Start MHz": band_info[0],
        "Stop MHz": band_info[1],
        "Band Name": band_info[2].strip() # Use strip() to remove leading/trailing spaces
    })




def query_safe(inst, command):
    """
    Safely queries the instrument, handling PyVISA errors.
    Args:
        inst (pyvisa.resources.Resource): The PyVISA instrument object.
        command (str): The SCPI command to query.
    Returns:
        str: The response from the instrument, stripped of whitespace,
             or "[Not Supported or Timeout]" if an error occurs.
    """
    try:
        return inst.query(command).strip()
    except pyvisa.VisaIOError:
        return "[Not Supported or Timeout]"

def write_safe(inst, command):
    """
    Safely writes a command to the instrument, handling PyVISA errors.
    Args:
        inst (pyvisa.resources.Resource): The PyVISA instrument object.
        command (str): The SCPI command to write.
    Returns:
        bool: True if the write was successful, False otherwise.
    """
    try:
        inst.write(command)
        return True
    except pyvisa.VisaIOError as e:
        print(f"Error writing command '{command}': {e}")
        return False

# The setup_arguments function is no longer needed for command-line parsing,
# as parameters will be taken from the GUI.
# def setup_arguments():
#     """
#     Parses command-line arguments for the spectrum analyzer sweep.
#     Allows customization of filename prefix, frequency range, step size, and user.
#     """
#     parser = argparse.ArgumentParser(description="Spectrum Analyzer Sweep and CSV Export")
#     # ... (arguments) ...
#     return parser.parse_args()

def initialize_instrument(visa_address):
    """
    Establishes a VISA connection to the instrument and performs initial configuration.
    This function now includes a retry mechanism to ensure the instrument is ready.

    Args:
        visa_address (str): The VISA address of the spectrum analyzer.

    Returns:
        pyvisa.resources.Resource: The instrument object if connection is successful, else None.
    """
    rm = pyvisa.ResourceManager()
    inst = None
    max_retries = 5
    retry_delay = 10  # seconds
    restart_wait_time = 23 # seconds to wait after a power reset

    for attempt in range(max_retries):
        try:
            print(f"Attempting to connect to instrument at {visa_address} (Attempt {attempt + 1}/{max_retries})...")
            inst = rm.open_resource(visa_address)
            inst.timeout = 30000 # Set timeout to 30 seconds for queries and data transfer

            # --- MODIFIED LOGIC HERE ---
            # This block handles the initial restart command if a connection is established.
            try:
                if inst:
                    #write_safe(inst, ":SYSTem:POWer:RESet")
                    print("⏳ Sent instrument restart command. Waiting for it to come back online...") # Moved emoji
                    #inst.close() # Close the current connection as the instrument will reboot
                    #time.sleep(restart_wait_time) # WAIT FOR THE RESTART TO COMPLETE
                    # After the wait, the next iteration of the loop will attempt a fresh connection
                else:
                    print("🛠️ Instrument not initialized yet, skipping direct restart command. Will attempt full initialization.") # Moved emoji
            except Exception as e:
                print(f"⚠️ Could not send restart command (likely connection already down or instrument rebooting): {e}. Proceeding to re-initialize.") # Moved emoji
                if inst: # Ensure resource is closed if it was briefly opened
                    inst.close()
                time.sleep(retry_delay) # Wait before the next attempt

            # After potentially sending a restart, or if it wasn't needed/failed,
            # we now try to establish a *fresh* connection and query for IDN.
            print(f"🔄 Re-attempting full connection and IDN query after potential restart (Attempt {attempt + 1}/{max_retries})...") # Moved emoji
            inst = rm.open_resource(visa_address) # Re-open the resource
            inst.timeout = 30000 # Reset timeout to 30 seconds

            # Try to query *IDN? to confirm connection
            idn_response = query_safe(inst, '*IDN?')
            if "[Not Supported or Timeout]" not in idn_response and idn_response:
                print(f"🎉 Successfully connected to: {idn_response}\n") # Moved emoji

                # Clear and reset the instrument
                write_safe(inst, "*CLS")
                write_safe(inst, "*RST")
                query_safe(inst, "*OPC?") # Wait for operations to complete
                print("✅ Instrument cleared and reset.") # Moved emoji

                print("📏 Set RBW to 1 kHz ") # Moved emoji
                inst.write(":SENSE:BAND:RES 1KHZ") # Set RBW to 1 kHz
                print("📺 Set BW to 1 kHz ") # Moved emoji
                inst.write(":SENSE:BAND:VID 1KHZ") # Set VBW to 1 kHz 

                # Configure preamplifier for high sensitivity
                write_safe(inst, ":SENS:POW:GAIN ON") # Equivalent to ':POWer:GAIN 1' for most Keysight instruments
                print("📡 Preamplifier turned ON for high sensitivity.") # Moved emoji

                # Configure display and marker settings
                print("📊 Display set to logarithmic scale") # Moved emoji
                write_safe(inst, ":DISP:WIND:TRAC:Y:SCAL LOG") # Logarithmic scale (dBm)
                print("📉 Display set to reference level -30 dBm.") # Moved emoji
                write_safe(inst, ":DISP:WIND:TRAC:Y:SCAL:RLEVel -30DBM") # Reference level (corrected from :DISP:WIND:TRAC:Y:RLEV)

                # Turn on all six markers and set them to position mode (neutral state)
                # They will be used by the peak table functionality per segment.
                # Note: N9340B typically supports 4 markers. Attempting 6, but instrument might ignore extra.
            #    for i in range(1, 7): # Markers 1 through 6
            #        write_safe(inst, f":CALC:MARK{i}:STATe ON")
            #        write_safe(f":CALC:MARK{i}:MODE POS")
             #       print(f"Marker {i} enabled in position mode (for peak table display).")
                print("⏸️ Trace 1 set to max hold") # Moved emoji
                write_safe(inst, ":TRAC1:MODE MAXHold")
                
                return inst
            else:
                print(f"😔 Failed to get IDN response: {idn_response}. Retrying...") # Moved emoji
                if inst: # Close the resource before retrying
                    inst.close()
                time.sleep(retry_delay)

        except pyvisa.VisaIOError as e:
            print(f"❌ VISA Error: Could not connect to or communicate with the instrument at {visa_address}: {e}") # Moved emoji
            print(f"⏳ Retrying in {retry_delay} seconds...") # Moved emoji
            if inst: # Ensure resource is closed if opened but failed
                inst.close()
            time.sleep(retry_delay)
        except Exception as e:
            print(f"💥 An unexpected error occurred during instrument initialization (Attempt {attempt + 1}): {e}") # Moved emoji
            print(f"⏳ Retrying in {retry_delay} seconds...") # Moved emoji
            if inst:
                inst.close()
            time.sleep(retry_delay)

    print(f"😞 Failed to initialize instrument after {max_retries} attempts. Please check connection and instrument status.") # Moved emoji
    return None

def scan_bands(inst, csv_writer, max_hold_time, rbw, selected_bands, last_scanned_band_index=0):
    """
    Iterates through predefined frequency bands, sets the start/stop frequencies,
    reduces RBW to 10000 Hz, and triggers a sweep for each band.
    It extracts trace data, writes it directly to the provided CSV writer,
    and also returns it for further processing (plotting).
    This function now dynamically segments bands to maintain a consistent
    effective resolution bandwidth per trace point.
    It also displays the time of day for each band scanned.
    
    Added last_scanned_band_index to resume from where it left off.

    Args:
        inst (pyvisa.resources.Resource): The PyVISA instrument object.
        csv_writer (csv.writer): The CSV writer object to write data to.
        max_hold_time (float): Duration in seconds for which MAX Hold should be active.
                                If > 0, MAX Hold mode is enabled for the scan.
        rbw (float): Resolution Bandwidth for segmenting bands.
        selected_bands (list): A list of band dictionaries to scan.
        last_scanned_band_index (int): Index of the band to start scanning from.
                                        Used for resuming scans after an error.
    Returns:
        tuple: (list: all_scan_data, int: last_successful_band_index)
                all_scan_data: A list of dictionaries, where each dictionary represents a data point
                               with 'Band Name', 'Frequency (Hz)', and 'Level (dBm)'.
                last_successful_band_index: The index of the last band that was fully
                                            or partially scanned successfully.
    """
    all_scan_data = [] # To store all data points across all bands for plotting
    last_successful_band_index = last_scanned_band_index

    print("\n--- 📡 Starting Band Scan ---") # Moved emoji

    # Set trace format back to ASCII
    write_safe(inst, ":FORMat:DATA ASCii") # Set data format to ASCII
    print("💾 Set trace data format to ASCII for data transfer.") # Moved emoji

    # Set read_termination to carriage return and line feed for ASCII data.
    # The N9340B manual (page 249) specifies that numeric data is returned with
    # a line feed after each value, and the message terminates with LF, so setting
    # termination to '\n' (LF) is appropriate.
    inst.read_termination = '\n' 
    inst.encoding = 'ascii' # ASCII data should be decoded as ASCII


    # *** Use selected_bands for scanning the instrument ***
    # Iterate through bands starting from last_scanned_band_index
    for i in range(last_scanned_band_index, len(selected_bands)):
        band = selected_bands[i]
        band_name = band["Band Name"]
        band_start_freq_hz = band["Start MHz"] * MHZ_TO_HZ
        band_stop_freq_hz = band["Stop MHz"] * MHZ_TO_HZ

        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"\n📈 [{current_time}] Processing Band: {band_name} (Total Range: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {band_stop_freq_hz/MHZ_TO_HZ:.3f} MHz)") # Moved emoji

        # Query the actual number of sweep points from the instrument for *this band*
        actual_sweep_points = 461 # Fallback to a common default if query fails
        
        # Calculate the optimal span for each segment to achieve desired RBW per point for *this band*
        # We want (Segment Span / (Actual Points - 1)) = Desired RBW
        # So, Segment Span = Desired RBW * (Actual Points - 1)
        optimal_segment_span_hz = rbw * (actual_sweep_points - 1)
        print(f"🎯 Optimal segment span to achieve {DEFAULT_RBW_STEP_SIZE_HZ/1000:.0f} kHz effective RBW per point for {band_name}: {optimal_segment_span_hz / MHZ_TO_HZ:.3f} MHz.")


        # Calculate total number of segments for the current band
        full_band_span_hz = band_stop_freq_hz - band_start_freq_hz
        if full_band_span_hz <= 0:
            total_segments_in_band = 1 # A single point or zero span, still one "segment" to process
        else:
            total_segments_in_band = int(np.ceil(full_band_span_hz / optimal_segment_span_hz))
            # Ensure at least one segment if the band has any span
            if total_segments_in_band == 0:
                total_segments_in_band = 1

        # --- RE-ADDED: "Wake Up" and Initial Band Configuration ---
        # Force a narrow span at the band's start frequency to ensure it "wakes up" and tunes
        write_safe(inst, f":SENS:FREQ:CENT {band_start_freq_hz}")
        write_safe(inst, ":SENS:FREQ:SPAN 1000HZ") # Set to 1 kHz span
        query_safe(inst, "*OPC?") # Wait for operation to complete
        time.sleep(0.2) # Small delay for the instrument to settle after narrow span

        # Now, explicitly set the START and STOP for the *first segment* of this new band
        # This will be the first segment's actual range.
        first_segment_stop_freq_hz = min(band_start_freq_hz + optimal_segment_span_hz, band_stop_freq_hz)
        write_safe(inst, f":SENS:FREQ:STAR {band_start_freq_hz}")
        write_safe(inst, f":SENS:FREQ:STOP {first_segment_stop_freq_hz}")
        query_safe(inst, "*OPC?") # Wait for operation to complete
        time.sleep(0.2) # Small delay for the instrument to settle
        print(f"🚀 Instrument forced to initial segment range for new band: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {first_segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz.") # Moved emoji
        # --- END RE-ADDED BLOCK ---

        current_segment_start_freq_hz = band_start_freq_hz
        segment_counter = 0

        while current_segment_start_freq_hz < band_stop_freq_hz:
            segment_counter += 1
            segment_stop_freq_hz = min(current_segment_start_freq_hz + optimal_segment_span_hz, band_stop_freq_hz)
            actual_segment_span_hz = segment_stop_freq_hz - current_segment_start_freq_hz

            if actual_segment_span_hz <= 0: # Avoid infinite loop if start == stop or negative span
                break

            # If the last segment is very small, it might result in less than 2 points,
            # which could cause issues with frequency step calculation.
            # We ensure minimum span if points are fixed.
            if actual_sweep_points > 1 and actual_segment_span_hz < (DEFAULT_RBW_STEP_SIZE_HZ * (actual_sweep_points - 1)):
                if segment_stop_freq_hz == band_stop_freq_hz: # This is the very last segment
                    pass # Allow smaller span for last segment, it will have fewer effective points but covers the end
                else:
                    # For intermediate segments, if the calculated span is too small, skip to next full segment
                    current_segment_start_freq_hz += optimal_segment_span_hz
                    continue # Skip this tiny segment, move to next potential full segment

            # Set instrument frequency range for the current segment
            write_safe(inst, f":SENS:FREQ:STAR {current_segment_start_freq_hz}")
            write_safe(inst, f":SENS:FREQ:STOP {segment_stop_freq_hz}")
            
            # Add a small delay after setting frequencies to allow instrument to configure
            time.sleep(0.1)

            query_safe(inst, "*OPC?") # Wait for the sweep to completed
            time.sleep(0.5) # Add a small delay for data processing within the instrument

            # Add settling time for max hold values to show up, if max hold is enabled
            if max_hold_time > 0:
               
                
                # Initial print without overwrite to start the line
               
            
                for sec_wait in range(int(max_hold_time), 0, -1):
                    # For numbers greater than 10, just show the number
                    display_text = f"⏳{sec_wait}"
                    sys.stdout.write(display_text) # \r to overwrite line
                    sys.stdout.flush()
                    time.sleep(1)
                sys.stdout.write("✅") # Clear the line and add newline
                sys.stdout.flush()
            
            #write_safe(inst, f":CALCulate:MARKer:ALL")
            
            query_safe(inst, "*OPC?") # Wait for the sweep to complete
            
            # Calculate progress for the emoji bar - Using more compatible ASCII characters
            progress_percentage = (segment_counter / total_segments_in_band)
            bar_length = 20 # Total number of characters in the bar
            filled_length = int(round(bar_length * progress_percentage))
            # Using '█' (U+2588 Full Block) and '-' (Hyphen) for better compatibility
            progressbar = '█' * filled_length + '-' * (bar_length - filled_length)

            # Combined print statement as per user request
            print(f"{progressbar}🔍 Span:📊{actual_segment_span_hz/MHZ_TO_HZ:.3f} MHz--📈{current_segment_start_freq_hz/MHZ_TO_HZ:.3f} MHz to 📉{segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz   ✅{segment_counter} of {total_segments_in_band}.")

            # Read and process trace data
            trace_data = []
            try:
                # Query trace data as ASCII
                raw_data_string = inst.query(":TRAC1:DATA?")
                # Split the string by comma and convert each part to float
                trace_data = [float(val.strip()) for val in raw_data_string.split(',')]

                num_trace_points_actual = len(trace_data)

                if num_trace_points_actual > 1:
                    freq_step_per_point_actual = actual_segment_span_hz / (num_trace_points_actual - 1)
                elif num_trace_points_actual == 1:
                    freq_step_per_point_actual = 0 # Single point, no step
                else:
                    freq_step_per_point_actual = 0 # No points
                    print("🚫 No trace data received for this segment.") # Moved emoji
                    current_segment_start_freq_hz = segment_stop_freq_hz
                    continue # Move to the next segment if no data

                # Loop to append data to all_scan_data and write to CSV
                for j, amp_value in enumerate(trace_data):
                    current_freq_for_point_hz = current_segment_start_freq_hz + (j * freq_step_per_point_actual)

                    # Append to list for plotting later
                    all_scan_data.append({
                        "Frequency (MHz)": current_freq_for_point_hz / MHZ_TO_HZ, # Store in MHz for DataFrame consistency
                        "Level (dBm)": amp_value,
                        "Band Name": band_name,

                    })

                    # Write directly to CSV file with desired order and units
                    csv_writer.writerow([
                        f"{current_freq_for_point_hz / MHZ_TO_HZ:.2f}",  # Frequency in MHz
                        f"{amp_value:.2f}"#,                                # Level in dBm
                        #band_name,                                         # Band Name

                    ])
                
                # Update last_successful_band_index after successfully processing a band
                last_successful_band_index = i

            except pyvisa.VisaIOError as e:
                print(f"🛑 Error reading trace data (PyVISA IO Error): {e}") # Moved emoji
                print(f"🐛 Raw data string potentially causing error: {e}") # PyVISA error will contain details
                # Restore original read_termination and encoding before re-raising or handling
                # No need to restore if we always set it to ASCII before the loop and it's consistent
                raise # Re-raise the exception to be caught by the main loop for recovery
            except ValueError as e:
                print(f"🚫 Error processing ASCII data (ValueError - cannot convert/unpack): {e}") # Moved emoji
                print(f"🐞 Raw data string for parsing: {raw_data_string[:100]}...") # Print a snippet of the problematic data
            except Exception as e:
                print(f"🚨 An unexpected error occurred during trace processing: {e}") # Moved emoji

            current_segment_start_freq_hz = segment_stop_freq_hz # Move to the start of the next segment

    # Restore original read_termination and encoding - these are now set consistently at the start of scan_bands
    # and do not need to be "restored" to a previous state at the end of the function.
    print("\n--- 🎉 Band Scan Complete! ---") # Moved emoji
    return all_scan_data, last_successful_band_index # Return the collected data and last successful index


def plot_spectrum_data(df: pd.DataFrame, output_html_filename: str, plot_title: str, include_gov_markers: bool, include_tv_markers: bool):
    """
    Generates an interactive Plotly Express line plot from the spectrum analyzer data.
    The plot is saved as an HTML file.

    Args:
        df (pd.DataFrame): DataFrame containing 'Frequency (MHz)', 'Level (dBm)',
                           and 'Band Name' columns.
        output_html_filename (str): The name of the HTML file to save the plot to.
        plot_title (str): The title for the Plotly chart.
        include_gov_markers (bool): Whether to include Government frequency band markers.
        include_tv_markers (bool): Whether to include TV channel band markers.
    """
    print(f"\n--- 📊 Generating Interactive Plot: {output_html_filename} ---") # Moved emoji

    # Create Interactive Plot with Plotly Express
    fig = px.line(df,
                  x="Frequency (MHz)",
                  y="Level (dBm)",
                  color="Band Name",  # Differentiate lines by band name
                  title=plot_title, # Use the dynamic plot_title here
                  labels={"Frequency (MHz)": "Frequency (MHz)", "Level (dBm)": "Amplitude (dBm)"},
                  hover_data={"Frequency (MHz)": ':.2f', "Level (dBm)": ':.2f', "Band Name": True}
                 )

    # Determine the full Y-axis range
    y_min_data = df['Level (dBm)'].min()
    y_max_data = df['Level (dBm)'].max()

    # Set fixed Y-axis range for the plot and for the band rectangles
    y_range_min = -100  # A reasonable lower bound for spectrum data
    y_range_max = 0     # As per your desired reference level max

    # Adjust y_range_min if data goes lower than -100 dBm
    if y_min_data < y_range_min:
        y_range_min = y_min_data - 10  # Provide some padding below the lowest point

    # Set Y-axis (Amplitude) Maximum to 0 dBm and apply range
    fig.update_yaxes(range=[y_range_min, y_range_max],
                     title="Amplitude (dBm)",
                     showgrid=True, gridwidth=1)

    # --- Add TV Band Markers ---
    if include_tv_markers:
        # Define colors for the TV band markers and text
        tv_marker_line_color = "rgba(255, 255, 0, 0.7)"  # Bright yellow, semi-transparent
        tv_marker_text_color = "yellow"
        tv_band_fill_color = "rgba(255, 255, 0, 0.05)"   # Very light yellow, highly transparent fill

        for band in TV_PLOT_BAND_MARKERS:
            # Add a shaded rectangle to represent the frequency band allocation
            fig.add_shape(
                type="rect",
                x0=band["Start MHz"],
                y0=y_range_min,  # Span full Y-axis range
                x1=band["Stop MHz"],
                y1=y_range_max,  # Span full Y-axis range
                line=dict(
                    color=tv_marker_line_color,
                    width=0.3,
                    dash="dot",
                ),
                fillcolor=tv_band_fill_color,
                layer="below",
            )

            # Add text markers using go.Scatter with mode='text'
            x_center = (band["Start MHz"] + band["Stop MHz"]) / 2
            y_text_position = y_range_max - (y_range_max - y_range_min) * 0.05

            fig.add_trace(go.Scatter(
                x=[x_center],
                y=[y_text_position],
                mode='text',
                text=[f"{band['Band Name']}<br>{band['Start MHz']:.1f}-{band['Stop MHz']:.1f} MHz"],
                textfont=dict(
                    size=8,
                    color=tv_marker_text_color
                ),
                showlegend=False,
                hoverinfo='text',
                name=f"Band Label: {band['Band Name']}"
            ))

    # --- Add Government Band Markers ---
    if include_gov_markers:
        # Define colors for the Government band markers and text
        gov_marker_line_color = "rgba(255, 0, 0, 0.9)"  # Red, semi-transparent
        gov_marker_text_color = "red"
        gov_band_fill_color = "rgba(255, 0, 0, 0.1)"    # Very light red, highly transparent fill

        # Define the four y-offsets for staggering
        y_offset_level_1 = 0.20
        y_offset_level_2 = 0.25
        y_offset_level_3 = 0.30
        y_offset_level_4 = 0.35
        y_offset_levels = [y_offset_level_1, y_offset_level_2, y_offset_level_3, y_offset_level_4]

        for i, band in enumerate(GOV_PLOT_BAND_MARKERS):
            # Add a shaded rectangle to represent the frequency band allocation
            fig.add_shape(
                type="rect",
                x0=band["Start MHz"],
                y0=y_range_min,
                x1=band["Stop MHz"],
                y1=y_range_max,
                line=dict(
                    color=gov_marker_line_color,
                    width=0.3,
                    dash="dot",
                ),
                fillcolor=gov_band_fill_color,
                layer="below",
            )

            # Add text markers using go.Scatter with mode='text'
            x_center = (band["Start MHz"] + band["Stop MHz"]) / 2

            # Determine the Y position based on staggering using modulo for 4 levels
            current_y_offset = y_offset_levels[i % len(y_offset_levels)]
            y_text_position = y_range_max - (y_range_max - y_range_min) * current_y_offset

            fig.add_trace(go.Scatter(
                x=[x_center],
                y=[y_text_position],
                mode='text',
                text=[f"{band['Band Name']}<br>{band['Start MHz']:.1f}-{band['Stop MHz']:.1f} MHz"],
                textfont=dict(
                    size=8,
                    color=gov_marker_text_color
                ),
                showlegend=False,
                hoverinfo='text',
                name=f"Band Label: {band['Band Name']}"
            ))

    # Set X-axis (Frequency) to Logarithmic Scale
    fig.update_xaxes(type="log",
                     title="Frequency (MHz)",
                     showgrid=True, gridwidth=1,
                     tickformat=None)

    # Apply Dark Mode Theme
    fig.update_layout(template="plotly_dark")

    # Save the plot as an HTML file
    fig.write_html(output_html_filename, auto_open=True)
    print(f"🖼️ --- Plotly Express Interactive Plot Generated and saved to {output_html_filename} ---") # Moved emoji


def wait_with_interrupt(wait_time_seconds):
    """
    Provides a timed delay with user interruption capability.
    Allows skipping the wait, quitting the program, or resuming.

    Args:
        wait_time_seconds (int): The total time to wait in seconds.
    """
    print("\n" + "="*50)
    print(f"⏰ Next full scan cycle in {wait_time_seconds // 60} minutes and {wait_time_seconds % 60} seconds.") # Moved emoji
    print("🛑 Press Ctrl+C at any time during the countdown to interact.") # Moved emoji
    print("="*50)

    seconds_remaining = int(wait_time_seconds) # Ensure integer for countdown
    skip_wait = False

    while seconds_remaining > 0:
        minutes = seconds_remaining // 60
        seconds = seconds_remaining % 60
        sys.stdout.write(f"\rTime until next scan: {minutes:02d}:{seconds:02d} ")
        sys.stdout.flush()

        try:
            time.sleep(1)
        except KeyboardInterrupt:
            sys.stdout.write("\n")
            sys.stdout.flush()
            choice = input("🤔 Countdown interrupted. (S)kip wait, (Q)uit program, or (R)esume countdown? ").strip().lower() # Moved emoji
            if choice == 's':
                skip_wait = True
                print("⏩ Skipping remaining wait time. Starting next scan shortly...") # Moved emoji
                break
            elif choice == 'q':
                print("👋 Exiting program.") # Moved emoji
                sys.exit(0)
            else:
                print("⏯️ Resuming countdown...") # Moved emoji
        seconds_remaining -= 1

    if not skip_wait:
        sys.stdout.write("\r" + " "*50 + "\r") # Clear the last countdown line
        sys.stdout.write("▶️ Starting next scan now!\n") # Moved emoji, Print a full message after countdown
        sys.stdout.flush()


def check_and_install_dependencies():
    """
    Checks if required Python modules are installed and offers to install them.
    If modules are missing and the user agrees, it attempts to install them via pip.
    """
    required_modules = {
        "pyvisa": "pyvisa",
        "numpy": "numpy",
        "pandas": "pandas",
        "plotly": "plotly",
    }

    missing_modules = []
    for module_name, pip_name in required_modules.items():
        try:
            __import__(module_name)
        except ImportError:
            missing_modules.append(pip_name)

    if missing_modules:
        print("\n--- ⚠️ Missing Dependencies ---") # Moved emoji
        print("The following Python modules are required but not installed:")
        for module in missing_modules:
            print(f"📦 - {module}") # Moved emoji
        print("----------------------------")

        install_choice = input("Do you want to attempt to install them now? (y/n): ").strip().lower()
        if install_choice == 'y':
            print("⬇️ Attempting to install missing modules...") # Moved emoji
            for module in missing_modules:
                try:
                    print(f"⚙️ Installing {module}...") # Moved emoji
                    # Use sys.executable to ensure pip associated with current Python interpreter is used
                    subprocess.check_call([sys.executable, "-m", "pip", "install", module])
                    print(f"✅ Successfully installed {module}.") # Moved emoji
                except subprocess.CalledProcessError as e:
                    print(f"❌ Error installing {module}: {e}") # Moved emoji
                    print(f"🧑‍💻 Please try installing it manually by running: pip install {module}") # Moved emoji
                    sys.exit(1) # Exit if an installation fails
                except Exception as e:
                    print(f"💥 An unexpected error occurred during installation of {module}: {e}") # Moved emoji
                    print(f"🧑‍💻 Please try installing it manually by running: pip install {module}") # Moved emoji
                    sys.exit(1)
            print("👍 All required modules installed successfully (or already present).") # Moved emoji
        else:
            print("😔 Installation declined. Please install the missing modules manually to run the script.") # Moved emoji
            for module in missing_modules:
                print(f"pip install {module}")
            sys.exit(1) # Exit if installation is declined and modules are missing
    else:
        print("\n✨ All required Python modules are already installed.") # Moved emoji

class ScanApp:
    def __init__(self, master):
        self.master = master
        master.title("Spectrum Analyzer Scan Configuration")
        master.geometry("600x750") # Adjusted size to accommodate more elements

        # Variables to hold input values
        self.scan_name_var = tk.StringVar(value="My_Spectrum_Scan")
        self.rbw_step_size_var = tk.StringVar(value=str(DEFAULT_RBW_STEP_SIZE_HZ))
        self.max_hold_time_var = tk.StringVar(value=str(DEFAULT_MAXHOLD_TIME_SECONDS))
        self.cycle_wait_time_var = tk.StringVar(value=str(DEFAULT_CYCLE_WAIT_TIME_SECONDS))

        # Variables for band checkboxes
        self.band_vars = []
        for band in SCAN_BAND_RANGES:
            # Default to all bands selected
            var = tk.BooleanVar(value=True) 
            self.band_vars.append((band, var))

        # Variables for plot markers
        self.include_gov_markers_var = tk.BooleanVar(value=True)
        self.include_tv_markers_var = tk.BooleanVar(value=True)

        self.create_widgets()

    def create_widgets(self):
        # Frame for the top "Start Scan" button
        top_button_frame = tk.Frame(self.master)
        top_button_frame.pack(pady=(10, 5))
        tk.Button(top_button_frame, text="Start Scan", command=self.start_scan).pack()

        # Main frame for left and right columns
        main_layout_frame = tk.Frame(self.master)
        main_layout_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Left column for text fields
        left_panel_frame = tk.Frame(main_layout_frame)
        left_panel_frame.pack(side=tk.LEFT, padx=(0, 10), anchor="nw") # Anchor North-West for top-left alignment

        # --- Scan Name ---
        tk.Label(left_panel_frame, text="Scan Series Name:").pack(pady=(10, 2), anchor="w")
        self.scan_name_entry = tk.Entry(left_panel_frame, textvariable=self.scan_name_var, width=30)
        self.scan_name_entry.pack(pady=2, anchor="w")

        # --- RBW Step Size ---
        tk.Label(left_panel_frame, text="RBW Step Size (Hz):").pack(pady=(10, 2), anchor="w")
        self.rbw_step_size_entry = tk.Entry(left_panel_frame, textvariable=self.rbw_step_size_var, width=20)
        self.rbw_step_size_entry.pack(pady=2, anchor="w")

        # --- Max Hold Time ---
        tk.Label(left_panel_frame, text="Max Hold Time (seconds):").pack(pady=(10, 2), anchor="w")
        self.max_hold_time_entry = tk.Entry(left_panel_frame, textvariable=self.max_hold_time_var, width=20)
        self.max_hold_time_entry.pack(pady=2, anchor="w")

        # --- Cycle Wait Time ---
        tk.Label(left_panel_frame, text="Cycle Wait Time (seconds):").pack(pady=(10, 2), anchor="w")
        self.cycle_wait_time_entry = tk.Entry(left_panel_frame, textvariable=self.cycle_wait_time_var, width=20)
        self.cycle_wait_time_entry.pack(pady=2, anchor="w")

        # Right column for the Quit button
        right_panel_frame = tk.Frame(main_layout_frame)
        right_panel_frame.pack(side=tk.RIGHT, padx=(10, 0), anchor="ne") # Anchor North-East for top-right alignment

        # --- Quit Button ---
        # Added some padding to align it visually with the top of the left column's content
        tk.Button(right_panel_frame, text="Quit", command=self.quit_app).pack(pady=(10, 2), anchor="e") 


        # --- Frequency Band Selection ---
        band_frame = tk.LabelFrame(self.master, text="Select Frequency Bands to Scan")
        band_frame.pack(pady=10, padx=10, fill="both", expand=True) # Use fill and expand for better layout

        # Use a canvas and scrollbar if many bands are present
        # Set a fixed height for the canvas to shrink the visible area
        canvas = tk.Canvas(band_frame, height=150) # Adjusted height to show fewer items at once
        scrollbar = tk.Scrollbar(band_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        for band, var in self.band_vars:
            tk.Checkbutton(scrollable_frame, text=f"{band['Band Name']} ({band['Start MHz']:.0f}-{band['Stop MHz']:.0f} MHz)", variable=var).pack(anchor="w")
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")


        # --- Plot Marker Options ---
        plot_options_frame = tk.LabelFrame(self.master, text="Plot Options")
        plot_options_frame.pack(pady=10, padx=10, fill="x")
        tk.Checkbutton(plot_options_frame, text="Include Government Band Markers", variable=self.include_gov_markers_var).pack(anchor="w")
        tk.Checkbutton(plot_options_frame, text="Include TV Channel Markers", variable=self.include_tv_markers_var).pack(anchor="w")

        # The original button_frame is no longer needed as buttons are moved
        # button_frame = tk.Frame(self.master)
        # button_frame.pack(pady=10)

        # tk.Button(button_frame, text="Start Scan", command=self.start_scan).pack(side=tk.LEFT, padx=10)
        # tk.Button(button_frame, text="Quit", command=self.quit_app).pack(side=tk.RIGHT, padx=10)

    def start_scan(self):
        try:
            scan_name = self.scan_name_var.get().strip()
            if not scan_name:
                messagebox.showerror("Input Error", "Scan Series Name cannot be empty.")
                return

            rbw = float(self.rbw_step_size_var.get())
            max_hold_time = float(self.max_hold_time_var.get())
            cycle_wait_time = float(self.cycle_wait_time_var.get())

            if rbw <= 0:
                messagebox.showerror("Input Error", "RBW Step Size must be a positive number.")
                return
            if max_hold_time < 0:
                messagebox.showerror("Input Error", "Max Hold Time cannot be negative.")
                return
            if cycle_wait_time < 0:
                messagebox.showerror("Input Error", "Cycle Wait Time cannot be negative.")
                return

            selected_bands = [band for band, var in self.band_vars if var.get()]
            if not selected_bands:
                messagebox.showerror("Input Error", "Please select at least one frequency band to scan.")
                return

            include_gov = self.include_gov_markers_var.get()
            include_tv = self.include_tv_markers_var.get()

            # Close the GUI window before starting the scan
            self.master.destroy()

            # Call the main scan logic with the collected parameters
            run_spectrum_scan_logic(
                scan_name=scan_name,
                rbw_step_size=rbw,
                max_hold_time=max_hold_time,
                cycle_wait_time=cycle_wait_time,
                selected_bands=selected_bands,
                include_gov_markers=include_gov,
                include_tv_markers=include_tv
            )

        except ValueError:
            messagebox.showerror("Input Error", "Please ensure all numerical inputs are valid numbers.")
        except Exception as e:
            messagebox.showerror("An Error Occurred", f"An unexpected error occurred: {e}")

    def quit_app(self):
        self.master.destroy()
        sys.exit(0) # Ensure the program exits properly

# Rename the original main function to a new name to be called by the GUI
def run_spectrum_scan_logic(scan_name, rbw_step_size, max_hold_time, cycle_wait_time, selected_bands, include_gov_markers, include_tv_markers):
    """
    Main function to connect to the N9340B Spectrum Analyzer,
    run initial setup, perform band scans, and then read final configuration.
    This now runs in a continuous loop with an interruptible wait time.
    Parameters are passed from the GUI.
    """
    # check_and_install_dependencies() # This should ideally be run once at the very start of the script

    # Determine instrument address dynamically by listing resources and picking the first one
    rm = pyvisa.ResourceManager()
    available_resources = rm.list_resources()

    instrument_address = None
    if available_resources:
        # Assuming only one VISA instrument is connected, take the first one
        instrument_address = available_resources[0]
        print(f"🔌 Automatically selected instrument: {instrument_address}")
    else:
        print("🚫 Error: No VISA instruments found. Please ensure your instrument is connected and drivers are installed.")
        messagebox.showerror("Instrument Error", "No VISA instruments found. Please ensure your instrument is connected and drivers are installed.")
        return


    inst = None
    scan_cycle_count = 0
    # Keep track of the last successfully scanned band index for recovery
    last_successful_band_index = 0

    try:
        while True: # This loop makes the program repeat indefinitely
            scan_cycle_count += 1
            print(f"\n--- 🔄 Starting Scan Cycle #{scan_cycle_count} ---")
            
            # This flag controls if we should skip to plotting/waiting or re-run the scan
            restart_scan_from_beginning_of_band = False 

            # Initialize/re-initialize the instrument. This function now handles retries.
            inst = initialize_instrument(instrument_address)

            if inst is None:
                print(f"⏳ Instrument initialization failed in cycle #{scan_cycle_count}. Waiting and retrying...")
                wait_with_interrupt(cycle_wait_time) # Use the GUI provided wait time
                continue # Skip to the next cycle attempt

            # File & directory setup for CSV and HTML plot
            # Base directory for all N9340 scans
            base_scan_dir = os.path.join(os.getcwd(), "N9340 Scans")
            
            # Create subfolder name based on the sanitized name
            name_folder = scan_name.replace(" ", "_") # Use the name from GUI, sanitize
            
            # Construct the full path to the specific scan subfolder
            scan_dir = os.path.join(base_scan_dir, name_folder)
            
            # Create the directory and any necessary parent directories
            os.makedirs(scan_dir, exist_ok=True)
            print(f"📁 Data will be saved in: {scan_dir}")

            timestamp = datetime.now().strftime("%Y%m%d@%H%M")
            csv_filename = os.path.join(scan_dir, f"{name_folder}_{timestamp}.csv")
            html_plot_filename = os.path.join(scan_dir, f"{name_folder}_{timestamp}.html")

            all_scan_data_current_cycle = [] # Collect data for the current cycle's plot

            try:
                print(f"📝 Opening CSV file for writing: {csv_filename}")
                with open(csv_filename, mode='w', newline='') as csvfile:
                    csv_writer = csv.writer(csvfile)

                    # Pass the selected_bands list from GUI
                    all_scan_data_current_cycle, last_successful_band_index = \
                        scan_bands(inst, csv_writer, max_hold_time, rbw_step_size, selected_bands, last_successful_band_index)
                
                # If scan_bands completes without raising an error, reset last_successful_band_index
                # for the next full cycle.
                last_successful_band_index = 0 

                if not all_scan_data_current_cycle:
                    print("📉 No scan data collected in this cycle. Skipping plotting.")
                else:
                    # Convert collected data to pandas DataFrame
                    df = pd.DataFrame(all_scan_data_current_cycle)
                    # Call the plotting function with GUI parameters
                    plot_spectrum_data(df, html_plot_filename, scan_name, include_gov_markers, include_tv_markers)

            except pyvisa.VisaIOError as e:
                print(f"🚨 !!! CRITICAL VISA I/O ERROR during scan: {e} !!!")
                print("🩹 Attempting to close instrument, re-initialize, and resume scan from last successful point.")
                if inst:
                    try:
                        inst.close()
                        print("🔌 Instrument connection closed.")
                    except Exception as close_e:
                        print(f"💥 Error closing instrument connection: {close_e}")
                inst = None # Ensure inst is None so initialize_instrument will try to open a new one
                
                print("🔄 Will attempt to re-initialize and continue scan in the next cycle.")
                time.sleep(5) # Short delay before re-attempting connection
                continue # Immediately go to the next cycle to try and reconnect/resume

            except Exception as e:
                print(f"🛑 An unexpected error occurred during scan cycle #{scan_cycle_count}: {e}")
                print("😴 Proceeding to wait period.")

            # CALLING THE INTERRUPTIBLE WAIT FUNCTION
            wait_with_interrupt(cycle_wait_time) # Uses the GUI provided wait time

    except KeyboardInterrupt:
        print("\n👋 Program interrupted by user (Ctrl+C) outside of wait period. Exiting.")
    except Exception as e:
        print(f"🚨 An unexpected critical error occurred in the main loop: {e}")
    finally:
        if inst and inst.session: # Check if inst object exists and has an active session
            inst.close()
            print("\n🔌 Connection to N9340B closed.")

# The actual entry point of the script
if __name__ == '__main__':
    # Check dependencies first
    check_and_install_dependencies()
    
    # Then launch the GUI
    root = tk.Tk()
    app = ScanApp(root)
    root.mainloop()
