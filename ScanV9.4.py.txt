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
            query_safe(inst, "*OPC?") # Wait for operations to complete
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


def scan_bands(inst, csv_writer, max_hold_time, rbw, selected_bands, last_scanned_band_index=0, rbw_config_val="1KHZ", vbw_config_val="1KHZ"):
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

        initialize_instrument(inst, 1, 1, 1, -30, 1, 1000)

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
            # The marker positions are at 100, 200, 300, 400 points from the start of the sweep.
            
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
                            "Band Name": band_name,
                        })


                    # Sort marker_data_points_temp by frequency before appending
                    marker_data_points_temp.sort(key=lambda x: x["Frequency (MHz)"])

                    # Add the sorted marker data points to the current_band_data
                    current_band_data.extend(marker_data_points_temp)

                    # Advance current_marker_base_freq by the step size
                    current_marker_base_freq += rbw # Advance by the desired resolution bandwidth (step size)

                except pyvisa.VisaIOError as e:
                    print(f"❌ VISA Error during marker data collection: {e}")
                    # Decide if this error is critical enough to stop the whole scan
                    # For now, we'll just break out of the current marker loop and try the next segment/band
                    break
                except ValueError as e:
                    print(f"❌ Data parsing error (e.g., non-float amplitude received): {e}. Raw: '{amp_values_str}'")
                    # Break out of the current marker loop and try the next segment/band
                    break
                except IndexError as e:
                    print(f"❌ Not enough marker amplitude values received: {e}. Raw: '{amp_values_str}'")
                    # Break out of the current marker loop and try the next segment/band
                    break
                except Exception as e:
                    print(f"💥 An unexpected error occurred during marker data collection: {e}")
                    break # Exit current marker loop, try next segment/band
    
            # Advance to the next segment's start frequency
            current_segment_start_freq_hz = segment_stop_freq_hz
            # Update the last successfully scanned band index after each segment completes
            last_successful_band_index = i
        
        # After processing all segments for the current band, sort current_band_data
        current_band_data.sort(key=lambda x: x["Frequency (MHz)"])

        # Write the sorted current_band_data to CSV and add to all_scan_data
        for data_point in current_band_data:
            all_scan_data.append(data_point) # Add to overall list for plotting
            csv_writer.writerow([
                f"{data_point['Frequency (MHz)']:.2f}",
                f"{data_point['Level (dBm)']:.2f}",
            ])
        print(f"✅ Band '{band_name}' data collected, sorted, and written to CSV.") # Add confirmation
    
    print("\n--- ✅ Band Scan Complete ---") # Moved emoji
    print("\n--- ✅ Clear and Reset for next scan ---") # Moved emoji
    write_safe(inst, "*CLS")
    write_safe(inst, "*RST")
    query_safe(inst, "*OPC?") # Wait for operations to complete
            


    return all_scan_data, last_successful_band_index




def plot_spectrum_data(df: pd.DataFrame, output_html_filename: str, plot_title: str, include_gov_markers: bool, include_tv_markers: bool, open_html_after_complete: bool):
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
        open_html_after_complete (bool): Whether to automatically open the generated HTML file.
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

    # Dynamically set X-axis range based on actual data
    x_min_data = df['Frequency (MHz)'].min()
    x_max_data = df['Frequency (MHz)'].max()

    # Set X-axis (Frequency) to Logarithmic Scale and apply dynamic range
    fig.update_xaxes(type="log",
                     title="Frequency (MHz)",
                     showgrid=True, gridwidth=1,
                     tickformat=None,
                     range=[np.log10(x_min_data), np.log10(x_max_data)]) # Apply dynamic range here

    # --- Add TV Band Markers ---
    if include_tv_markers:
        # Define colors for the TV band markers and text
        tv_marker_line_color = "rgba(255, 255, 0, 0.7)"  # Bright yellow, semi-transparent
        tv_marker_text_color = "yellow"
        tv_band_fill_color = "rgba(255, 255, 0, 0.05)"   # Very light yellow, highly transparent fill

        for band in TV_PLOT_BAND_MARKERS:
            # Only add markers if they are within the actual scanned frequency range
            if band["Start MHz"] < x_max_data and band["Stop MHz"] > x_min_data:
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
            # Only add markers if they are within the actual scanned frequency range
            if band["Start MHz"] < x_max_data and band["Stop MHz"] > x_min_data:
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

    # Apply Dark Mode Theme
    fig.update_layout(template="plotly_dark")

    # Save the plot as an HTML file
    fig.write_html(output_html_filename, auto_open=open_html_after_complete)
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
        master.geometry("700x700") # Adjusted width to 700 pixels, height to 700 pixels

        # Variables to hold input values
        self.scan_name_var = tk.StringVar(value="My_Spectrum_Scan")
        self.rbw_step_size_var = tk.StringVar(value=str(DEFAULT_RBW_STEP_SIZE_HZ))
        self.max_hold_time_var = tk.StringVar(value=str(DEFAULT_MAXHOLD_TIME_SECONDS)) # Changed to StringVar for OptionMenu
        self.cycle_wait_time_var = tk.StringVar(value=str(DEFAULT_CYCLE_WAIT_TIME_SECONDS))
        self.visa_address_var = tk.StringVar(value="Not Connected") # To display the VISA address
        self.instrument_instance = None # To hold the instrument object

        # New variables for instrument configuration controls
        self.clear_reset_var = tk.BooleanVar(value=True)
        self.rbw_config_var = tk.StringVar(value="1KHZ")
        self.vbw_config_var = tk.StringVar(value="1KHZ")
        self.preamplifier_var = tk.BooleanVar(value=True)
        self.display_log_var = tk.BooleanVar(value=True)
        self.ref_level_var = tk.StringVar(value="-30") # Default to -30 dBm
        self.max_hold_var = tk.BooleanVar(value=True) # Using BooleanVar for single radio button behavior

        # Variables for band checkboxes
        self.band_vars = []
        for band in SCAN_BAND_RANGES:
            # Default to all bands selected
            var = tk.BooleanVar(value=True) 
            self.band_vars.append((band, var))

        # Variables for plot markers
        self.include_gov_markers_var = tk.BooleanVar(value=True)
        self.include_tv_markers_var = tk.BooleanVar(value=True)
        self.open_html_after_complete_var = tk.BooleanVar(value=True) # New variable for "Open HTML after complete"

        self.create_widgets()
        self.connect_and_display_visa() # Attempt connection on startup

    def create_widgets(self):
        # Frame for VISA address display
        visa_info_frame = tk.Frame(self.master)
        visa_info_frame.pack(pady=(10, 5))
        tk.Label(visa_info_frame, text="Connected Instrument:").pack(side=tk.LEFT)
        self.visa_address_label = tk.Label(visa_info_frame, textvariable=self.visa_address_var, fg="blue")
        self.visa_address_label.pack(side=tk.LEFT)

        # Frame for buttons (Start Scan, System Restart, Reset, Quit)
        button_row_frame = tk.Frame(self.master)
        button_row_frame.pack(pady=(5, 10))

        # Start Scan button - initially disabled and grey
        self.start_scan_button = tk.Button(button_row_frame, text="Start Scan", command=self.start_scan, state=tk.DISABLED, bg="green", fg="white") # Set fg to white
        self.start_scan_button.pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_row_frame, text="System Restart", command=self.system_restart).pack(side=tk.LEFT, padx=5)
        tk.Button(button_row_frame, text="Reset Instrument", command=self.reset_instrument).pack(side=tk.LEFT, padx=5)
        tk.Button(button_row_frame, text="Quit", command=self.quit_app).pack(side=tk.LEFT, padx=5)


        # Main frame for left and right columns (rest of the inputs)
        main_layout_frame = tk.Frame(self.master)
        main_layout_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Left column for scan parameters
        scan_params_frame = tk.LabelFrame(main_layout_frame, text="Scan Parameters")
        scan_params_frame.pack(side=tk.LEFT, padx=(0, 10), anchor="nw", fill="y") # Anchor North-West for top-left alignment

        # --- Scan Name ---
        tk.Label(scan_params_frame, text="Scan Series Name:").pack(pady=(10, 2), anchor="w")
        self.scan_name_entry = tk.Entry(scan_params_frame, textvariable=self.scan_name_var, width=30)
        self.scan_name_entry.pack(pady=2, anchor="w")

        # --- RBW Step Size (for scan_bands function) ---
        tk.Label(scan_params_frame, text="RBW Step Size (Hz) for Scan:").pack(pady=(10, 2), anchor="w")
        self.rbw_step_size_entry = tk.Entry(scan_params_frame, textvariable=self.rbw_step_size_var, width=20)
        self.rbw_step_size_entry.pack(pady=2, anchor="w")

        # --- Max Hold Time (Dropdown) ---
        tk.Label(scan_params_frame, text="Max Hold Time (seconds):").pack(pady=(10, 2), anchor="w")
        max_hold_options = [str(i) for i in range(11)] # Values 0 to 10
        self.max_hold_time_var.set(str(DEFAULT_MAXHOLD_TIME_SECONDS)) # Set default value
        self.max_hold_time_menu = tk.OptionMenu(scan_params_frame, self.max_hold_time_var, *max_hold_options)
        self.max_hold_time_menu.pack(pady=2, anchor="w")
        self.max_hold_time_menu.config(width=10) # Adjust width of the dropdown

        # --- Cycle Wait Time ---
        tk.Label(scan_params_frame, text="Cycle Wait Time (seconds):").pack(pady=(10, 2), anchor="w")
        self.cycle_wait_time_entry = tk.Entry(scan_params_frame, textvariable=self.cycle_wait_time_var, width=20)
        self.cycle_wait_time_entry.pack(pady=2, anchor="w")

        # --- Instrument Configuration Options (Middle Column) ---
        config_frame = tk.LabelFrame(main_layout_frame, text="Instrument Initial Configuration")
        config_frame.pack(side=tk.LEFT, padx=(0, 10), anchor="nw", fill="y") # Anchor North-West for top-left alignment

        tk.Checkbutton(config_frame, text="Clear and Reset Instrument (*CLS; *RST)", variable=self.clear_reset_var).pack(anchor="w")

        tk.Label(config_frame, text="Set RBW (e.g., 1KHZ, 10KHZ):").pack(pady=(5, 2), anchor="w")
        self.rbw_config_entry = tk.Entry(config_frame, textvariable=self.rbw_config_var, width=15)
        self.rbw_config_entry.pack(pady=2, anchor="w")

        tk.Label(config_frame, text="Set VBW (e.g., 1KHZ, 10KHZ):").pack(pady=(5, 2), anchor="w")
        self.vbw_config_entry = tk.Entry(config_frame, textvariable=self.vbw_config_var, width=15)
        self.vbw_config_entry.pack(pady=2, anchor="w")

        tk.Checkbutton(config_frame, text="Preamplifier ON for high sensitivity", variable=self.preamplifier_var).pack(anchor="w")
        tk.Checkbutton(config_frame, text="Display set to logarithmic scale", variable=self.display_log_var).pack(anchor="w")

        tk.Label(config_frame, text="Display Reference Level (dBm):").pack(pady=(5, 2), anchor="w")
        # Validate input to be numeric only for reference level
        vcmd = (self.master.register(self.validate_numeric_input), '%P')
        self.ref_level_entry = tk.Entry(config_frame, textvariable=self.ref_level_var, width=10, validate="key", validatecommand=vcmd)
        self.ref_level_entry.pack(pady=2, anchor="w")

        # Max Hold Radio Button (acting as a checkbox for a single option)
        max_hold_frame = tk.Frame(config_frame)
        max_hold_frame.pack(anchor="w", pady=(5,0))
        tk.Radiobutton(max_hold_frame, text="Trace 1 set to max hold", variable=self.max_hold_var, value=True).pack(side=tk.LEFT)


        # --- Plot Marker Options (Rightmost Column) ---
        plot_options_frame = tk.LabelFrame(main_layout_frame, text="Plot Options")
        plot_options_frame.pack(side=tk.LEFT, padx=(0, 10), anchor="nw", fill="y") # Placed next to config_frame

        tk.Checkbutton(plot_options_frame, text="Include Government Band Markers", variable=self.include_gov_markers_var).pack(anchor="w")
        tk.Checkbutton(plot_options_frame, text="Include TV Channel Markers", variable=self.include_tv_markers_var).pack(anchor="w")
        tk.Checkbutton(plot_options_frame, text="Open HTML after complete", variable=self.open_html_after_complete_var).pack(anchor="w") # New checkbox

        # --- Frequency Band Selection ---
        # This frame will now be below the three side-by-side frames
        band_frame = tk.LabelFrame(self.master, text="Select Frequency Bands to Scan")
        band_frame.pack(pady=10, padx=10, fill="both", expand=True) 

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
    
    def validate_numeric_input(self, P):
        """Validates that the input is a valid number (integer or float), including negative signs."""
        if P.strip() == "": # Allow empty string for temporary input clearing
            return True
        try:
            float(P)
            return True
        except ValueError:
            return False

    def connect_and_display_visa(self):
        """Attempts to connect to the VISA instrument and display its address."""
        rm = pyvisa.ResourceManager()
        available_resources = rm.list_resources()

        if available_resources:
            instrument_address = available_resources[0]
            self.visa_address_var.set(instrument_address)
            print(f"🔌 Automatically selected instrument: {instrument_address}")
            
            # Attempt to open the resource (without full initialization) to check connectivity
            try:
                temp_inst = rm.open_resource(instrument_address)
                # Now, perform the full initialization with GUI settings
                if initialize_instrument(
                    temp_inst,
                    self.clear_reset_var.get(),
                    self.preamplifier_var.get(),
                    self.display_log_var.get(),
                    float(self.ref_level_var.get()), # Convert to float
                    self.max_hold_var.get(),
                    self.rbw_config_var.get() # Pass rbw_config_val here
                ):
                    self.instrument_instance = temp_inst # Store the initialized instance
                    self.visa_address_label.config(fg="green") # Change color to green on successful connection
                    self.start_scan_button.config(state=tk.NORMAL, bg="green") # Enable and color green
                else:
                    # Initialization failed, close the temporary instance
                    temp_inst.close()
                    self.instrument_instance = None
                    self.visa_address_label.config(fg="red") # Change color to red on failed connection
                    self.start_scan_button.config(state=tk.DISABLED, bg="lightgray") # Disable and color grey
                    messagebox.showerror("Initialization Error", "Failed to initialize instrument with current settings. Check console for details.")

            except pyvisa.VisaIOError as e:
                self.instrument_instance = None
                self.visa_address_label.config(fg="red")
                self.start_scan_button.config(state=tk.DISABLED, bg="lightgray")
                print(f"❌ VISA Error: Could not connect to or communicate with the instrument at {instrument_address}: {e}")
                messagebox.showerror("Connection Error", f"Could not connect to instrument: {e}. Check console for details.")
            except ValueError as e:
                self.instrument_instance = None
                self.visa_address_label.config(fg="red")
                self.start_scan_button.config(state=tk.DISABLED, bg="lightgray")
                messagebox.showerror("Configuration Error", f"Invalid numeric input for instrument settings: {e}. Please check RBW/VBW/Ref Level formats.")
            except Exception as e:
                self.instrument_instance = None
                self.visa_address_label.config(fg="red")
                self.start_scan_button.config(state=tk.DISABLED, bg="lightgray")
                messagebox.showerror("Unexpected Error", f"An unexpected error occurred during connection/initialization: {e}. Check console for details.")

        else:
            self.visa_address_var.set("No Instrument Found")
            self.visa_address_label.config(fg="red")
            self.start_scan_button.config(state=tk.DISABLED, bg="lightgray") # Disable and color grey
            print("🚫 Error: No VISA instruments found. Please ensure your instrument is connected and drivers are installed.")

    def system_restart(self): # Renamed from system_power_off
        """Sends a system restart command to the connected instrument."""
        if self.instrument_instance:
            try:
                write_safe(self.instrument_instance, ":SYSTem:POWer:RESet") # Changed command
                messagebox.showinfo("Instrument Control", "Sent System Restart command. Instrument will reboot.")
                print("Sent :SYSTem:POWer:RESet command. Instrument will reboot.")
                # After sending restart, assume connection might be lost, so reset instance
                self.instrument_instance.close()
                self.instrument_instance = None
                self.visa_address_var.set("Restarting...")
                self.visa_address_label.config(fg="orange") # Indicate restarting state
                self.start_scan_button.config(state=tk.DISABLED, bg="lightgray") # Disable during restart
                # Optionally, re-attempt connection after a delay
                self.master.after(25000, self.connect_and_display_visa) # Wait 25 seconds then try to reconnect
            except Exception as e:
                messagebox.showerror("Instrument Error", f"Failed to send Restart command: {e}")
                print(f"Error sending :SYSTem:POWer:RESet: {e}")
        else:
            messagebox.showwarning("No Instrument", "No instrument connected to send restart command.")

    def reset_instrument(self):
        """Sends a reset command to the connected instrument."""
        if self.instrument_instance:
            try:
                write_safe(self.instrument_instance, "*RST")
                messagebox.showinfo("Instrument Control", "Sent *RST (Reset) command.")
                print("Sent *RST command.")
            except Exception as e:
                messagebox.showerror("Instrument Error", f"Failed to send Reset command: {e}")
                print(f"Error sending *RST: {e}")
        else:
            messagebox.showwarning("No Instrument", "No instrument connected to send reset command.")

    def start_scan(self):
        try:
            # Ensure instrument is connected before starting scan
            if not self.instrument_instance:
                messagebox.showerror("Connection Error", "No instrument connected. Please ensure the device is connected and try again.")
                return

            scan_name = self.scan_name_var.get().strip()
            if not scan_name:
                messagebox.showerror("Input Error", "Scan Series Name cannot be empty.")
                return

            rbw = float(self.rbw_step_size_var.get())
            max_hold_time = float(self.max_hold_time_var.get())
            cycle_wait_time = float(self.cycle_wait_time_var.get())

            # Get the RBW and VBW configuration values from the GUI
            rbw_config = self.rbw_config_var.get()
            vbw_config = self.vbw_config_var.get()


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
            open_html = self.open_html_after_complete_var.get() # Get the new checkbox value

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
                include_tv_markers=include_tv,
                rbw_config_val=rbw_config, # Pass new parameters
                vbw_config_val=vbw_config,  # Pass new parameters
                open_html_after_complete=open_html # Pass the new parameter
            )

        except ValueError:
            messagebox.showerror("Input Error", "Please ensure all numerical inputs are valid numbers.")
        except Exception as e:
            messagebox.showerror("An Error Occurred", f"An unexpected error occurred: {e}")

    def quit_app(self):
        # Ensure the instrument connection is closed before quitting
        if self.instrument_instance:
            try:
                self.instrument_instance.close()
                print("🔌 Instrument connection closed before quitting.")
            except Exception as e:
                print(f"Error closing instrument on quit: {e}")
        self.master.destroy()
        sys.exit(0) # Ensure the program exits properly

# The main instrument initialization function, now accepting GUI parameters
def run_spectrum_scan_logic(scan_name, rbw_step_size, max_hold_time, cycle_wait_time, selected_bands, include_gov_markers, include_tv_markers, rbw_config_val, vbw_config_val, open_html_after_complete):
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
            try:
                # Re-open the resource for the main scan loop if it was closed or not passed
                if inst is None:
                    inst = rm.open_resource(instrument_address)
                    inst.timeout = 30000 # Reset timeout
                    print(f"🔄 Re-opened instrument connection for scan cycle #{scan_cycle_count}.")
                
                # The RBW and VBW settings are now applied in scan_bands,
                # so initialize_instrument doesn't need them for setting the instrument.
                # However, other settings (preamp, display, ref level, max hold) are still applied here.
                # The parameters rbw_config_val and vbw_config_val are passed to initialize_instrument
                # but are effectively ignored there for setting the instrument.
                if not initialize_instrument(
                    inst,
                    True, # Always clear and reset on re-init in the loop for consistency
                    True, # Assuming preamplifier on for re-init in loop
                    True, # Assuming display log for re-init in loop
                    -30,  # Assuming -30 dBm for re-init in loop
                    True, # Assuming max hold for re-init in loop
                    rbw_config_val # Pass rbw_config_val here
                ):
                    print(f"⏳ Instrument re-initialization failed in cycle #{scan_cycle_count}. Waiting and retrying...")
                    wait_with_interrupt(cycle_wait_time)
                    continue # Skip to the next cycle attempt

            except pyvisa.VisaIOError as e:
                print(f"❌ VISA Error during re-connection in main loop: {e}")
                inst = None
                print(f"⏳ Instrument re-connection failed in cycle #{scan_cycle_count}. Waiting and retrying...")
                wait_with_interrupt(cycle_wait_time)
                continue # Skip to the next cycle attempt
            except Exception as e:
                print(f"💥 An unexpected error occurred during instrument setup in main loop: {e}")
                inst = None
                print(f"⏳ Instrument setup failed in cycle #{scan_cycle_count}. Waiting and retrying...")
                wait_with_interrupt(cycle_wait_time)
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

                    # Pass the selected_bands list and RBW/VBW config values from GUI
                    all_scan_data_current_cycle, last_successful_band_index = \
                        scan_bands(inst, csv_writer, max_hold_time, rbw_step_size, selected_bands, last_successful_band_index, rbw_config_val, vbw_config_val)
                
                # If scan_bands completes without raising an error, reset last_successful_band_index
                # for the next full cycle.
                last_successful_band_index = 0 

                if not all_scan_data_current_cycle:
                    print("📉 No scan data collected in this cycle. Skipping plotting.")
                else:
                    # Convert collected data to pandas DataFrame
                    df = pd.DataFrame(all_scan_data_current_cycle)
                    # Call the plotting function with GUI parameters
                    plot_spectrum_data(df, html_plot_filename, scan_name, include_gov_markers, include_tv_markers, open_html_after_complete)

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
