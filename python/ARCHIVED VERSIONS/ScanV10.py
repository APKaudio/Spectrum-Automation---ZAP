import tkinter as tk
from tkinter import messagebox, scrolledtext
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
import threading # Import threading for running scan in a separate thread

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


def initialize_instrument(inst, clear_reset, preamplifier_on, display_log, ref_level, max_hold_on, rbw_config_val):
    """
    Performs initial configuration of the instrument based on GUI settings.
    RBW and VBW settings are now handled in scan_bands.

    Args:
        inst (pyvisa.resources.Resource): The PyVISA instrument object.
        clear_reset (bool): Whether to clear and reset the instrument.
        preamplifier_on (bool): Whether to turn on the preamplifier.
        display_log (bool): Whether to set display to logarithmic scale.
        ref_level (float): Reference level in dBm.
        max_hold_on (bool): Whether to set trace 1 to max hold.
        rbw_config_val (str): Resolution Bandwidth value from GUI (e.g., "1KHZ").
    Returns:
        bool: True if initialization is successful, False otherwise.
    """
    try:
        inst.timeout = 30000 # Set timeout to 30 seconds for queries and data transfer

        if clear_reset:
            write_safe(inst, "*CLS")
            write_safe(inst, "*RST")
            

            #query_safe(inst, "*OPC?") # Wait for operations to complete
            inst.clear() # Flush buffer after OPC query
            print("✅ Instrument cleared and reset.")

        if preamplifier_on:
            write_safe(inst, ":SENS:POW:GAIN ON")
            print("📡 Preamplifier turned ON for high sensitivity.")
        else:
            write_safe(inst, ":SENS:POW:GAIN OFF")
            print("📡 Preamplifier turned OFF.")

        if display_log:
            write_safe(inst, ":DISP:WIND:TRAC:Y:SCAL LOG")
            print("📊 Display set to logarithmic scale")
        else:
            write_safe(inst, ":DISP:WIND:TRAC:Y:SCAL LIN")
            print("📊 Display set to linear scale")

        write_safe(inst, f":DISP:WIND:TRAC:Y:SCAL:RLEVel {ref_level}DBM")
        print(f"📉 Display set to reference level {ref_level} dBm.")

        if max_hold_on:
            write_safe(inst, ":TRAC1:MODE MAXHold")
            print("⏸️ Trace 1 set to max hold")
        else:
            write_safe(inst, ":TRAC1:MODE WRITe") # Or NORMal, depending on desired default
            print("▶️ Trace 1 set to normal/write mode")

        print("⚙️ Setting traces 2, 3, and 4 to BLANK mode...")
        write_safe(inst, ":TRAC2:MODE BLANK")
        write_safe(inst, ":TRAC3:MODE BLANK")
        write_safe(inst, ":TRAC4:MODE BLANK")

        # Configure markers 1-6
        for marker_num in range(1, 6): # Changed range to include 6 markers
            write_safe(inst, f":CALC:MARK{marker_num}:STAT ON") # Enable marker
            write_safe(inst, f":CALC:MARK{marker_num}:MODE NORMal") # Set to normal mode
            # Set marker bandwidth resolution to the configured RBW
            write_safe(inst, f":CALC:MARK{marker_num}:BWID:RES {rbw_config_val}")
        print(f"✅ Markers 1-6 enabled and set to {rbw_config_val} bandwidth.") # Updated print message
        
        
        return True
    except pyvisa.VisaIOError as e:
        print(f"❌ VISA Error during instrument configuration: {e}")
        return False
    except Exception as e:
        print(f"💥 An unexpected error occurred during instrument configuration: {e}")
        return False


def scan_bands(inst, csv_writer, max_hold_time, rbw, selected_bands, last_scanned_band_index=0, rbw_config_val="1000", vbw_config_val="1000"):
    """
    Iterates through predefined frequency bands, sets the start/stop frequencies,
    and triggers a sweep for each band. It collects data by moving Marker 1
    across the segment's frequency range, writing data to the CSV writer,
    and returning it for plotting.
    This function now dynamically segments bands to maintain a consistent
    effective resolution bandwidth per trace point.
    It also displays the time of day for each band scanned.
    
    Added last_scanned_band_index to resume from where it left off.

    Args:
        inst (pyvisa.resources.Resource): The PyVISA instrument object.
        csv_writer (csv.writer): The CSV writer object to write data to.
        max_hold_time (float): Duration in seconds for which MAX Hold should be active.
                                If > 0, MAX Hold mode is enabled for the scan.
        rbw (float): Resolution Bandwidth for segmenting bands. This will now be used as the step size for marker data collection.
        selected_bands (list): A list of band dictionaries to scan.
        last_scanned_band_index (int): Index of the band to start scanning from.
                                        Used for resuming scans after an error.
        rbw_config_val (str): Resolution Bandwidth value from GUI (e.g., "1KHZ").
        vbw_config_val (str): Video Bandwidth value from GUI (e.g., "1KHZ").
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

    print("💾 Using marker-based data collection by sweeping Marker 1.")

    # *** Use selected_bands for scanning the instrument ***
    # Iterate through bands starting from last_scanned_band_index
    for i in range(last_scanned_band_index, len(selected_bands)):
        band = selected_bands[i]
        band_name = band["Band Name"]
        band_start_freq_hz = band["Start MHz"] * MHZ_TO_HZ
        band_stop_freq_hz = band["Stop MHz"] * MHZ_TO_HZ

        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"\n📈 [{current_time}] Processing Band: {band_name} (Total Range: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {band_stop_freq_hz/MHZ_TO_HZ:.3f} MHz)") # Moved emoji

        # --- RE-ADDED: "Wake Up" and Initial Band Configuration ---

      #  initialize_instrument(inst, 1, 1, 1, -30, 1, 1000)

        # Force a narrow span at the band's start frequency to ensure it "wakes up" and tunes
        write_safe(inst, f":SENS:FREQ:CENT {band_start_freq_hz}")
        write_safe(inst, ":SENS:FREQ:SPAN 1000HZ") # Set to 1 kHz span
        print("\n--- ✅ Clear and Reset for next scan ---") # Moved emoji
        
        # Query the actual number of sweep points from the instrument for *this band*
        # The :SENSe:SWEep:POINts? command is not a real command, hardcoding to 462.
        actual_sweep_points = 401 # Changed to 401 points for 10 divisions * 40 points/div + 1
        print(f"📊 Using fixed {actual_sweep_points} sweep points per trace for {band_name}.")


        # Calculate the optimal span for each segment to achieve desired RBW per point for *this band*
        # We want (Segment Span / (Actual Points - 1)) = Desired RBW
        # So, Segment Span = Desired RBW * (Actual Points - 1)
        # Note: 'rbw' from GUI is the desired step size for marker collection.
        # The segment span should be large enough to cover multiple marker steps.
        # For 401 points, 100 divisions, so 40 points per division.
        # If we want 4 markers per segment at 25%, 50%, 75% of the *display*,
        # and the display has 401 points, the span should accommodate this.
        # Let's use the full span of the band for setting the instrument span for each segment.
        # The marker stepping will then happen within this span.
        
        # Calculate total number of segments for the current band
        full_band_span_hz = band_stop_freq_hz - band_start_freq_hz
        if full_band_span_hz <= 0:
            total_segments_in_band = 1 # A single point or zero span, still one "segment" to process
            optimal_segment_span_hz = 1 # Minimum span for a single point
        else:
            # The number of points per segment is fixed (401).
            # The step size for marker collection is 'rbw' (from GUI, e.g., 10kHz).
            # So, the effective span of one "sweep" (where we collect 4 markers) is 4 * rbw.
            # We need to divide the full band span by this effective step size to get total steps.
            # However, the user's request implies that the *instrument* is set to a span,
            # and then markers are moved *within that span*.
            # Let's assume the "segment" here is the full band, and we step markers within it.
            # Or, if we still want to segment, the segment span should be large enough for multiple marker steps.
            
            # Let's simplify: Set the instrument span to the full band span.
            # Then, step Marker 1 through this span with 'rbw' as the step.
            # The 25%, 50%, 75% markers will be relative to Marker 1's position.
            
            # If we are still using "segments" that are smaller than the full band,
            # then optimal_segment_span_hz should be related to the number of points we want to collect
            # within that segment using Marker 1's steps.
            
            # Re-interpreting the "400 points wide" and "4.6 MHz BW" comment:
            # If the instrument has 401 points displayed, and the span is 4.6 MHz,
            # then each point represents 4.6 MHz / 400 = 11.5 kHz.
            # The user wants to set markers at 25%, 50%, 75% of the *screen* (400 divisions).
            # This means the marker positions are relative to the current *instrument span*.

            # Let's define `points_per_division = 40` (401 points, 10 divisions implies 40 points/div)
            # The marker positions are at 100, 200, 300, 400. points from the start of the sweep.
            
            # The previous logic of `optimal_segment_span_hz` was tied to `actual_sweep_points`.
            # Let's keep the segment logic, but adjust how markers are placed.
            
            # For marker-based data, the 'optimal_segment_span_hz' should be the span
            # that we set the instrument to for each sub-sweep.
            # If we want 401 points effectively, and each point is 'rbw' apart,
            # then the span should be (401 - 1) * rbw.
            optimal_segment_span_hz = rbw * (actual_sweep_points - 1)
            total_segments_in_band = int(np.ceil(full_band_span_hz / optimal_segment_span_hz))
            if total_segments_in_band == 0:
                total_segments_in_band = 1
        print(f"🎯 Optimal segment span for instrument setting: {optimal_segment_span_hz / MHZ_TO_HZ:.3f} MHz.")

        # Initialize a temporary list to hold all data points for the current band
        current_band_data = []

        # Now, explicitly set the START and STOP for the *first segment* of this new band
        current_segment_start_freq_hz = band_start_freq_hz # Initialize for the loop
        segment_counter = 0
        while current_segment_start_freq_hz < band_stop_freq_hz:
            segment_counter += 1
            segment_stop_freq_hz = min(current_segment_start_freq_hz + optimal_segment_span_hz, band_stop_freq_hz)
            actual_segment_span_hz = segment_stop_freq_hz - current_segment_start_freq_hz

            # Initialize current_marker_base_freq for this segment.
            # This ensures it's always defined before the inner while loop's condition is checked.
            current_marker_base_freq = current_segment_start_freq_hz

            if actual_segment_span_hz <= 0:
                # Avoid infinite loop if start == stop or negative span
                print(f"⚠️ Skipping segment due to zero or negative span: {current_segment_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz")
                break # Exit this segment loop, move to next band or end scan

            # Set instrument frequency range for the current segment
            write_safe(inst, f":SENS:FREQ:STAR {current_segment_start_freq_hz}")
            write_safe(inst, f":SENS:FREQ:STOP {segment_stop_freq_hz}")

            # Add a small delay after setting frequencies to allow instrument to configure
            time.sleep(0.1)
            query_safe(inst, "*OPC?") # Wait for the sweep to completed
            inst.clear() # Flush buffer after OPC query
            time.sleep(0.5) # Add a small delay for data processing within the instrument

            # Add settling time for max hold values to show up, if max hold is enabled
            if max_hold_time > 0:
                for sec_wait in range(int(max_hold_time), 0, -1):
                    display_text = f"⏳{sec_wait}"
                    sys.stdout.write(display_text) # \r to overwrite line
                    sys.stdout.flush()
                    time.sleep(1)
                sys.stdout.write("✅") # Clear the line and add newline
                sys.stdout.flush()

            #write_safe(inst, f":INITiate:CONTinuous 0")
            #print("\n--- ✅ Continuous OFF SCAN FROZEN ON SCREEN ---") # Moved emoji
            query_safe(inst, "*OPC?") # Wait for the sweep to complete
            inst.clear() # Flush buffer after OPC query

            # Calculate progress for the emoji bar - Using more compatible ASCII characters
            progress_percentage = (segment_counter / total_segments_in_band)
            bar_length = 20 # Total number of characters in the bar
            filled_length = int(round(bar_length * progress_percentage))
            # Using '█' (U+2588 Full Block) and '-' (Hyphen) for better compatibility
            progressbar = '█' * filled_length + '-' * (bar_length - filled_length)

            # Combined print statement as per user request
            print(f"{progressbar} 🔍 📈{current_segment_start_freq_hz/MHZ_TO_HZ:.3f} MHz to 📉{segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz ✅{segment_counter} of {total_segments_in_band}.")

            # Read and process data using markers
            # The instrument has 401 points (0 to 400).
            # Quarter points are at 100, 200, 300, 400.
            # Marker positions are relative to the *current segment's start frequency*.

            # Calculate the frequency step per display point for the current segment
            if (actual_sweep_points - 1) > 0:
                freq_step_per_display_point = actual_segment_span_hz / (actual_sweep_points - 1)
            else:
                freq_step_per_display_point = 0 # Should not happen with actual_sweep_points = 401

            marker_offsets_percentage = [0.0, 0.2, 0.4, 0.6, 0.8] # Relative to current segment span
            
            # Ensure rbw is not zero to prevent infinite loop
            if rbw <= 0:
                print("🚫 RBW step size is zero or negative. Cannot collect marker data.")
                current_segment_start_freq_hz = segment_stop_freq_hz # Skip to next segment
                continue

            # Loop to collect data points using Marker 1 as the reference
            # The loop condition is now based on Marker 4's potential position
            while (current_marker_base_freq + (marker_offsets_percentage[4] * actual_segment_span_hz)) <= segment_stop_freq_hz:
                try:
                    marker_data_points_temp = [] # Initialize an empty list to store results

                    # --- Construct the concatenated command strings ---
                    set_commands = []
                    query_commands = []
                    for marker_idx in range(5):
                        marker_num = marker_idx + 1
                        # Calculate the frequency for the current marker (same as before)
                        marker_freq_hz = current_marker_base_freq + (marker_offsets_percentage[marker_idx] * actual_segment_span_hz)
                        marker_freq_hz = min(marker_freq_hz, segment_stop_freq_hz) # Cap at segment stop frequency
                        
                        # Append the set command
                        set_commands.append(f":CALC:MARK{marker_num}:X {marker_freq_hz}HZ")
                        
                        # Append the query command for amplitude. Note the leading colon for absolute path.
                        # We want to query all Y values in a single string.
                        query_commands.append(f":CALC:MARK{marker_num}:Y?")

                    # Join the commands with a semicolon
                    full_set_command = ";".join(set_commands)
                    full_query_command = ";".join(query_commands)

                    # --- Execute the commands ---
                    # 1. Send all marker frequency set commands in one go
                    write_safe(inst, full_set_command)
                    
                    # 2. Query all marker amplitudes in one go
                    # The instrument is expected to return the values separated by semicolons.
                    # E.g., "-10.123;-12.456;-8.901;-15.789;-20.500"
                    amp_values_str = query_safe(inst, full_query_command)
                    amp_values_str_list = amp_values_str.split(';')

                    # Now, iterate through the received amplitudes and the original marker calculations
                    for marker_idx in range(5):
                        marker_num = marker_idx + 1 # Re-declare or use if in scope
                        # Get the amplitude value for the current marker
                        amp_value = float(amp_values_str_list[marker_idx])

                        # Re-calculate marker_freq_hz for data point consistency (important!)
                        # This ensures the frequency stored matches what was *commanded* to the instrument,
                        # as the query for Y? doesn't return the X value.
                        marker_freq_hz = current_marker_base_freq + (marker_offsets_percentage[marker_idx] * actual_segment_span_hz)
                        marker_freq_hz = min(marker_freq_hz, segment_stop_freq_hz)

                        marker_data_points_temp.append({
                            "Frequency (MHz)": marker_freq_hz / MHZ_TO_HZ,
                            "Level (dBm)": amp_value,
                            "Band Name": band_name
                        })
                        csv_writer.writerow([band_name, marker_freq_hz / MHZ_TO_HZ, amp_value])
                        all_scan_data.append(marker_data_points_temp[-1]) # Add the last appended data point

                    # Move to the next set of marker positions, effectively stepping Marker 1
                    # Increment current_marker_base_freq by rbw (the desired step size)
                    current_marker_base_freq += rbw

                except Exception as e:
                    print(f"⚠️ Error collecting marker data in band {band_name}, segment {segment_counter}: {e}. Attempting to continue...")
                    break # Break from inner while loop, move to next segment or band

            current_segment_start_freq_hz = segment_stop_freq_hz # Move to the start of the next segment

        last_successful_band_index = i # Update last successful band index
    
    print("\n--- 🏁 Band Scan Complete ---") # Moved emoji
    return all_scan_data, last_successful_band_index


# --- Plotting Functions ---
def plot_data(data, csv_file_path):
    if not data:
        print("No data to plot.")
        return

    df = pd.DataFrame(data)
    
    # Ensure 'Frequency (MHz)' is numeric for sorting and plotting
    df['Frequency (MHz)'] = pd.to_numeric(df['Frequency (MHz)'])
    df = df.sort_values(by='Frequency (MHz)').reset_index(drop=True)

    # Convert 'Level (dBm)' to numeric, coercing errors to NaN
    df['Level (dBm)'] = pd.to_numeric(df['Level (dBm)'], errors='coerce')
    # Drop rows where 'Level (dBm)' became NaN
    df.dropna(subset=['Level (dBm)'], inplace=True)

    fig = px.line(df, x="Frequency (MHz)", y="Level (dBm)", 
                  title="RF Spectrum Scan",
                  labels={"Frequency (MHz)": "Frequency (MHz)", "Level (dBm)": "Power Level (dBm)"},
                  hover_data={"Frequency (MHz)": ':.3f', "Level (dBm)": ':.2f', "Band Name": True},
                  line_shape='linear') # Use 'linear' for straight lines between points

    # Add band markers for GOV_PLOT_BAND_MARKERS
    for band in GOV_PLOT_BAND_MARKERS:
        fig.add_vrect(x0=band["Start MHz"], x1=band["Stop MHz"],
                      annotation_text=band["Band Name"], 
                      annotation_position="top left",
                      fillcolor="LightSalmon", opacity=0.2, line_width=0)

    # Add TV channel band markers (if desired)
    for band in TV_PLOT_BAND_MARKERS:
        fig.add_vrect(x0=band["Start MHz"], x1=band["Stop MHz"],
                      annotation_text=band["Band Name"], 
                      annotation_position="top left",
                      fillcolor="LightGreen", opacity=0.2, line_width=0)

    fig.update_layout(
        xaxis_title="Frequency (MHz)",
        yaxis_title="Power Level (dBm)",
        hovermode="x unified",
        template="plotly_dark", # Use a dark theme for better visibility
        # Add tooltips for band markers
        annotations=[
            dict(
                x=(band["Start MHz"] + band["Stop MHz"]) / 2,
                y=1.0, # Position at top of y-axis
                xref="x", yref="paper",
                text=band["Band Name"],
                showarrow=False,
                font=dict(size=8, color="black"),
                align="center",
                opacity=0.7
            ) for band in GOV_PLOT_BAND_MARKERS + TV_PLOT_BAND_MARKERS # Combine lists
        ]
    )

    # Ensure the plot is saved to an HTML file
    plot_output_dir = "plots"
    os.makedirs(plot_output_dir, exist_ok=True)
    plot_file_name = datetime.now().strftime("spectrum_scan_%Y%m%d_%H%M%S.html")
    plot_output_path = os.path.join(plot_output_dir, plot_file_name)
    fig.write_html(plot_output_path, auto_open=False)
    print(f"📊 Plot saved to {plot_output_path}")
    return plot_output_path

# --- Dependency Check and Install ---
def check_and_install_dependencies():
    required_packages = ['pyvisa', 'pandas', 'plotly']
    missing_packages = []
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)

    if missing_packages:
        print(f"Missing packages: {', '.join(missing_packages)}. Attempting to install...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing_packages)
            print("Dependencies installed successfully.")
        except Exception as e:
            print(f"Failed to install dependencies: {e}. Please install them manually using 'pip install <package_name>'.")
            sys.exit(1) # Exit if essential dependencies cannot be installed

# --- Console Redirector Class ---
class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag
        self.buffer = "" # A buffer to store partial lines

    def write(self, str_to_write):
        self.buffer += str_to_write
        # If a newline is present, process the buffer
        if '\n' in self.buffer:
            parts = self.buffer.split('\n', 1)
            line = parts[0]
            self.buffer = parts[1] if len(parts) > 1 else ""

            self.widget.configure(state="normal")
            self.widget.insert(tk.END, line + "\n", self.tag)
            self.widget.see(tk.END) # Auto-scroll to the end
            self.widget.configure(state="disabled")
        else:
            # If no newline, just update the current line (for progress bars)
            # This is a bit more complex for a Text widget as it doesn't naturally overwrite lines.
            # For simplicity, we'll append. Real progress bar updates in a Text widget are tricky.
            # For now, let's just ensure it's written and then the next newline will push it.
            # If the user requires true overwriting, a more advanced solution would be needed.
            self.widget.configure(state="normal")
            # To simulate overwriting for non-newline outputs (like '⏳X'), we could try deleting the last line
            # This is a basic attempt and might not be perfect for all cases.
            if str_to_write.strip() and not str_to_write.endswith('\n'):
                # Try to delete the last line if it's not empty and doesn't end with a newline
                # This is a heuristic for progress bar updates
                current_last_line = self.widget.get("end-2c linestart", "end-1c")
                if current_last_line.strip(): # If the last line actually has content
                    try:
                        self.widget.delete("end-2c linestart", "end-1c")
                    except:
                        pass # Ignore errors if line is already empty or doesn't exist

            self.widget.insert(tk.END, str_to_write, self.tag)
            self.widget.see(tk.END)
            self.widget.configure(state="disabled")

    def flush(self):
        # This is often needed for file-like objects, but for a Text widget, write handles flushing
        pass

# --- Tkinter GUI Application ---
class ScanApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RF Spectrum Scanner")
        
        # Set initial window size: 1000 pixels wide, height will adjust based on content
        # We can set a default height or let Tkinter calculate it. Let's aim for a reasonable height.
        self.root.geometry("1000x700") 

        self.rm = pyvisa.ResourceManager()
        self.inst = None
        self.scan_thread = None # To hold the scanning thread
        self.scanning = False # Flag to control scanning process

        # Create two main frames: one for the original GUI, one for the console output
        self.main_frame = tk.Frame(root)
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.console_frame = tk.Frame(root, width=300, bg="black")
        self.console_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=False, padx=10, pady=10)
        self.console_frame.pack_propagate(False) # Prevent console_frame from shrinking

        # Console output Text widget
        self.console_output = scrolledtext.ScrolledText(self.console_frame, wrap=tk.WORD, bg="black", fg="white", font=("Consolas", 10))
        self.console_output.pack(expand=True, fill=tk.BOTH)
        self.console_output.configure(state="disabled") # Make it read-only
        
        # Redirect stdout and stderr to the console_output widget
        sys.stdout = TextRedirector(self.console_output, "stdout")
        sys.stderr = TextRedirector(self.console_output, "stderr")

        print("--- RF Spectrum Scanner GUI Initialized ---")

        # --- GUI Elements for main_frame (existing controls) ---
        self.create_widgets(self.main_frame)
        self.populate_resources()


    def create_widgets(self, parent_frame):
        # Resource selection
        resource_frame = tk.LabelFrame(parent_frame, text="Instrument Connection", padx=10, pady=10)
        resource_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(resource_frame, text="VISA Resource:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.resource_var = tk.StringVar(self.root)
        self.resource_dropdown = tk.OptionMenu(resource_frame, self.resource_var, "No Resources Found")
        self.resource_dropdown.grid(row=0, column=1, sticky=tk.EW, pady=2)
        self.refresh_button = tk.Button(resource_frame, text="Refresh", command=self.populate_resources)
        self.refresh_button.grid(row=0, column=2, padx=5, pady=2)
        self.connect_button = tk.Button(resource_frame, text="Connect", command=self.connect_instrument)
        self.connect_button.grid(row=0, column=3, padx=5, pady=2)

        # Instrument Settings
        settings_frame = tk.LabelFrame(parent_frame, text="Scan Settings", padx=10, pady=10)
        settings_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(settings_frame, text="Reference Level (dBm):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.ref_level_entry = tk.Entry(settings_frame)
        self.ref_level_entry.insert(0, "-30")
        self.ref_level_entry.grid(row=0, column=1, sticky=tk.EW, pady=2)

        self.preamp_var = tk.BooleanVar(self.root, value=True)
        tk.Checkbutton(settings_frame, text="Preamplifier ON", variable=self.preamp_var).grid(row=1, column=0, sticky=tk.W, pady=2)

        self.log_scale_var = tk.BooleanVar(self.root, value=True)
        tk.Checkbutton(settings_frame, text="Logarithmic Scale", variable=self.log_scale_var).grid(row=2, column=0, sticky=tk.W, pady=2)
        
        self.max_hold_var = tk.BooleanVar(self.root, value=False) # Changed default to False as per discussion
        tk.Checkbutton(settings_frame, text="Max Hold Trace 1", variable=self.max_hold_var).grid(row=3, column=0, sticky=tk.W, pady=2)
        
        tk.Label(settings_frame, text="Max Hold Time (s):").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.max_hold_time_entry = tk.Entry(settings_frame)
        self.max_hold_time_entry.insert(0, str(DEFAULT_MAXHOLD_TIME_SECONDS))
        self.max_hold_time_entry.grid(row=4, column=1, sticky=tk.EW, pady=2)

        tk.Label(settings_frame, text="RBW (Hz):").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.rbw_entry = tk.Entry(settings_frame)
        self.rbw_entry.insert(0, str(DEFAULT_RBW_STEP_SIZE_HZ))
        self.rbw_entry.grid(row=5, column=1, sticky=tk.EW, pady=2)

        tk.Label(settings_frame, text="Scan Cycle Wait Time (s):").grid(row=6, column=0, sticky=tk.W, pady=2)
        self.cycle_wait_time_entry = tk.Entry(settings_frame)
        self.cycle_wait_time_entry.insert(0, str(DEFAULT_CYCLE_WAIT_TIME_SECONDS))
        self.cycle_wait_time_entry.grid(row=6, column=1, sticky=tk.EW, pady=2)

        # Output folder and filename
        output_frame = tk.LabelFrame(parent_frame, text="Output Settings", padx=10, pady=10)
        output_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(output_frame, text="Output Folder:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.output_folder_entry = tk.Entry(output_frame)
        self.output_folder_entry.insert(0, "scan_data")
        self.output_folder_entry.grid(row=0, column=1, sticky=tk.EW, pady=2)

        # Band Selection
        band_selection_frame = tk.LabelFrame(parent_frame, text="Frequency Band Selection", padx=10, pady=10)
        band_selection_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        self.band_checkboxes = []
        self.band_vars = []

        # Use a canvas with a scrollbar for the band selection
        band_canvas = tk.Canvas(band_selection_frame)
        band_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        band_scrollbar = tk.Scrollbar(band_selection_frame, orient="vertical", command=band_canvas.yview)
        band_scrollbar.pack(side=tk.RIGHT, fill="y")

        band_canvas.configure(yscrollcommand=band_scrollbar.set)
        band_canvas.bind('<Configure>', lambda e: band_canvas.configure(scrollregion = band_canvas.bbox("all")))

        self.inner_band_frame = tk.Frame(band_canvas)
        band_canvas.create_window((0, 0), window=self.inner_band_frame, anchor="nw")

        for i, band in enumerate(SCAN_BAND_RANGES):
            var = tk.BooleanVar(self.root, value=True) # All selected by default
            cb = tk.Checkbutton(self.inner_band_frame, text=f"{band['Band Name']} ({band['Start MHz']:.3f}-{band['Stop MHz']:.3f} MHz)", variable=var)
            cb.grid(row=i, column=0, sticky=tk.W, padx=5, pady=1)
            self.band_checkboxes.append(cb)
            self.band_vars.append(var)

        # Action Buttons (Start/Stop)
        action_button_frame = tk.Frame(parent_frame, padx=10, pady=10)
        action_button_frame.pack(pady=10, padx=10, fill=tk.X)

        self.start_scan_button = tk.Button(action_button_frame, text="Start Scan", command=self.start_scan, height=2, bg="green", fg="white")
        self.start_scan_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.stop_scan_button = tk.Button(action_button_frame, text="Stop Scan", command=self.stop_scan, height=2, bg="red", fg="white", state=tk.DISABLED)
        self.stop_scan_button.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)
        
        self.plot_button = tk.Button(action_button_frame, text="Generate Plot (Last Scan)", command=self.generate_plot, height=2)
        self.plot_button.pack(pady=10)
        self.plot_button.pack_forget() # Hide initially, show after scan

        self.last_scan_data = [] # To store data for plotting
        self.last_csv_file_path = "" # To store CSV path for plotting

    def populate_resources(self):
        try:
            resources = self.rm.list_resources()
            if resources:
                self.resource_var.set(resources[0])
                menu = self.resource_dropdown["menu"]
                menu.delete(0, "end")
                for resource in resources:
                    menu.add_command(label=resource, command=tk._setit(self.resource_var, resource))
                print("✅ VISA resources refreshed.")
            else:
                self.resource_var.set("No Resources Found")
                menu = self.resource_dropdown["menu"]
                menu.delete(0, "end")
                menu.add_command(label="No Resources Found", command=tk._setit(self.resource_var, "No Resources Found"))
                print("❌ No VISA resources found.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to list VISA resources: {e}")
            print(f"❌ Error listing VISA resources: {e}")

    def connect_instrument(self):
        resource_name = self.resource_var.get()
        if resource_name == "No Resources Found":
            messagebox.showwarning("Connection Warning", "Please select a valid VISA resource.")
            print("⚠️ Cannot connect: No VISA resources selected.")
            return

        try:
            if self.inst:
                self.inst.close()
                print("🔌 Closed existing instrument connection.")

            self.inst = self.rm.open_resource(resource_name)
            
            # Retrieve and print instrument ID
            try:
                identity = query_safe(self.inst, "*IDN?")
                print(f"✅ Connected to: {identity}")
            except Exception:
                print("✅ Connected to instrument (IDN query failed or unsupported).")

            # Initial configuration
            clear_reset = True # Always clear and reset on initial connection
            preamplifier_on = self.preamp_var.get()
            display_log = self.log_scale_var.get()
            ref_level = float(self.ref_level_entry.get())
            max_hold_on = self.max_hold_var.get()
            # RBW config for initialization, this is a string like "1000HZ"
            rbw_config_val = f"{int(float(self.rbw_entry.get()))}HZ" 

            if initialize_instrument(self.inst, clear_reset, preamplifier_on, display_log, ref_level, max_hold_on, rbw_config_val):
                messagebox.showinfo("Connection Status", "Instrument connected and configured successfully!")
                self.start_scan_button.config(state=tk.NORMAL)
            else:
                messagebox.showerror("Connection Error", "Failed to configure instrument.")
                self.inst.close()
                self.inst = None
                self.start_scan_button.config(state=tk.DISABLED)

        except pyvisa.VisaIOError as e:
            messagebox.showerror("Connection Error", f"Could not connect to {resource_name}: {e}")
            print(f"❌ VISA Connection Error: {e}")
            self.inst = None
            self.start_scan_button.config(state=tk.DISABLED)
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numeric values for settings.")
            print("❌ Input Error: Invalid numeric value for settings.")
            self.start_scan_button.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
            print(f"💥 Unexpected error during connection: {e}")
            self.inst = None
            self.start_scan_button.config(state=tk.DISABLED)

    def start_scan(self):
        if self.inst is None:
            messagebox.showwarning("Scan Warning", "Please connect to an instrument first.")
            print("⚠️ Scan cannot start: No instrument connected.")
            return
        
        if self.scanning:
            messagebox.showinfo("Scan Status", "Scan is already running.")
            print("ℹ️ Scan is already running.")
            return

        self.scanning = True
        self.start_scan_button.config(state=tk.DISABLED)
        self.stop_scan_button.config(state=tk.NORMAL)
        self.plot_button.pack_forget() # Hide plot button when scan starts
        self.last_scan_data = [] # Clear previous scan data
        self.last_csv_file_path = ""

        self.scan_thread = threading.Thread(target=self._run_scan_loop)
        self.scan_thread.daemon = True # Allow the program to exit even if thread is running
        self.scan_thread.start()
        print("▶️ Scan initiated in a separate thread.")

    def stop_scan(self):
        if not self.scanning:
            messagebox.showinfo("Scan Status", "Scan is not currently running.")
            print("ℹ️ Scan is not currently running.")
            return

        self.scanning = False
        print("🛑 Stop command issued. Waiting for current cycle to complete...")
        messagebox.showinfo("Scan Control", "Scan stop command sent. It will stop after the current band or wait period completes.")
        self.start_scan_button.config(state=tk.NORMAL)
        self.stop_scan_button.config(state=tk.DISABLED)

    def _run_scan_loop(self):
        # This method will contain the main scan loop logic
        # It needs to be robust against instrument disconnection and allow stopping
        
        output_folder = self.output_folder_entry.get()
        os.makedirs(output_folder, exist_ok=True)
        current_datetime_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_file_name = os.path.join(output_folder, f"spectrum_data_{current_datetime_str}.csv")
        self.last_csv_file_path = csv_file_name

        selected_bands = [
            SCAN_BAND_RANGES[i] for i, var in enumerate(self.band_vars) if var.get()
        ]
        if not selected_bands:
            print("🚫 No frequency bands selected for scanning. Please select at least one band.")
            messagebox.showwarning("Scan Warning", "No frequency bands selected for scanning. Please select at least one band.")
            self.scanning = False
            self.root.after(0, self._reset_buttons_after_scan) # Reset buttons on GUI thread
            return

        cycle_wait_time = float(self.cycle_wait_time_entry.get())
        max_hold_time = float(self.max_hold_time_entry.get())
        rbw = float(self.rbw_entry.get())

        scan_cycle_count = 0
        last_scanned_band_index = 0
        
        try:
            with open(csv_file_name, 'w', newline='') as csv_file:
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow(['Band Name', 'Frequency (MHz)', 'Level (dBm)'])
                print(f"💾 Data will be saved to: {csv_file_name}")

                while self.scanning:
                    scan_cycle_count += 1
                    print(f"\n--- 🚀 Starting Scan Cycle #{scan_cycle_count} ---")
                    if self.inst is None:
                        print("🚫 Instrument not connected. Attempting to reconnect...")
                        # Try to reconnect using the last selected resource
                        try:
                            self.inst = self.rm.open_resource(self.resource_var.get())
                            identity = query_safe(self.inst, "*IDN?")
                            print(f"✅ Reconnected to: {identity}")
                            # Re-initialize instrument after reconnection
                            clear_reset = True
                            preamplifier_on = self.preamp_var.get()
                            display_log = self.log_scale_var.get()
                            ref_level = float(self.ref_level_entry.get())
                            max_hold_on = self.max_hold_var.get()
                            rbw_config_val = f"{int(rbw)}HZ"
                            if not initialize_instrument(self.inst, clear_reset, preamplifier_on, display_log, ref_level, max_hold_on, rbw_config_val):
                                print("❌ Failed to re-initialize instrument. Stopping scan.")
                                self.scanning = False
                                break
                        except Exception as e:
                            print(f"❌ Failed to reconnect to instrument: {e}. Stopping scan.")
                            messagebox.showerror("Scan Error", f"Lost connection to instrument and failed to reconnect: {e}. Stopping scan.")
                            self.scanning = False
                            break

                    try:
                        # Pass rbw and vbw as strings for instrument configuration
                        rbw_config_str = f"{int(rbw)}HZ"
                        vbw_config_str = f"{int(rbw/3)}HZ" # VBW typically 1/3 of RBW, adjust as needed

                        current_scan_data, last_scanned_band_index = scan_bands(
                            self.inst, csv_writer, max_hold_time, rbw, selected_bands,
                            last_scanned_band_index, rbw_config_str, vbw_config_str
                        )
                        self.last_scan_data.extend(current_scan_data) # Accumulate data

                        if not self.scanning: # Check if stop button was pressed during scan_bands
                            print("🛑 Scan stopped by user command.")
                            break

                    except pyvisa.VisaIOError as e:
                        print(f"❌ VISA Error during scan cycle #{scan_cycle_count}: {e}")
                        messagebox.showwarning("Scan Warning", f"Instrument communication error: {e}. Attempting to re-initialize and continue.")
                        if self.inst:
                            try:
                                self.inst.close()
                                print("🔌 Closed instrument connection due to error.")
                            except Exception as close_e:
                                print(f"💥 Error closing instrument connection: {close_e}")
                        self.inst = None # Ensure inst is None so initialize_instrument will try to open a new one
                        
                        print("🔄 Will attempt to re-initialize and continue scan in the next cycle.")
                        time.sleep(5) # Short delay before re-attempting connection
                        continue # Immediately go to the next cycle to try and reconnect/resume

                    except Exception as e:
                        print(f"🛑 An unexpected error occurred during scan cycle #{scan_cycle_count}: {e}")
                        print("😴 Proceeding to wait period.")
                        # Even if an error occurs, we still want to wait if scanning is still true
                        # and then potentially continue or stop based on the user's intent.
                        if not self.scanning:
                            break # If an error occurred AND stop was pressed, break

                    # CALLING THE INTERRUPTIBLE WAIT FUNCTION
                    # This function needs to be aware of the self.scanning flag
                    self._wait_with_interrupt(cycle_wait_time) # Uses the GUI provided wait time
                    
                    if not self.scanning: # Check again after wait period
                        print("🛑 Scan stopped by user command after wait period.")
                        break

            print("\n👋 Program finished.")

        except Exception as e:
            print(f"🚨 An unexpected critical error occurred in the main loop: {e}")
            messagebox.showerror("Critical Error", f"A critical error occurred in the main scan loop: {e}")
        finally:
            if self.inst and self.inst.session: # Check if inst object exists and has an active session
                self.inst.close()
                print("\n🔌 Connection to N9340B closed.")
            self.root.after(0, self._reset_buttons_after_scan) # Ensure buttons are reset on GUI thread

    def _wait_with_interrupt(self, seconds):
        start_time = time.time()
        print(f"Waiting for {seconds} seconds (interruptible)...")
        while time.time() - start_time < seconds:
            if not self.scanning: # Check the flag frequently
                print("\nWait interrupted by stop command.")
                return
            time.sleep(0.1) # Check every 100ms
        print("\nWait complete.")

    def _reset_buttons_after_scan(self):
        self.start_scan_button.config(state=tk.NORMAL)
        self.stop_scan_button.config(state=tk.DISABLED)
        if self.last_scan_data: # Only show plot button if there's data
            self.plot_button.pack(pady=10) # Show plot button after scan finishes

    def generate_plot(self):
        if not self.last_scan_data:
            messagebox.showwarning("Plot Warning", "No scan data available to plot. Please run a scan first.")
            print("🚫 No data to plot.")
            return
        
        print("Generating plot...")
        try:
            plot_path = plot_data(self.last_scan_data, self.last_csv_file_path)
            messagebox.showinfo("Plot Generated", f"Plot saved to {plot_path}")
            print(f"✅ Plot generation complete: {plot_path}")
            # Optionally open the plot in a web browser
            # import webbrowser
            # webbrowser.open(plot_path)
        except Exception as e:
            messagebox.showerror("Plot Error", f"Failed to generate plot: {e}")
            print(f"❌ Error generating plot: {e}")


# The actual entry point of the script
if __name__ == '__main__':
    # Check dependencies first
    check_and_install_dependencies()
    
    # Then launch the GUI
    root = tk.Tk()
    app = ScanApp(root)
    root.mainloop()
