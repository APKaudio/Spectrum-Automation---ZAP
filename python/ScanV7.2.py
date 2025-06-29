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
import sys # Import sys for stdout and exit
import subprocess # NEW: Import subprocess to run pip commands

# Define constants for better readability and easier modification
MHZ_TO_HZ = 1_000_000 # Conversion factor from MHz to Hz

# Updated wait time variable and its usage for the continuous loop
DEFAULT_RBW_STEP_SIZE_HZ = 10000 # 10 kHz RBW resolution desired per data point
DEFAULT_CYCLE_WAIT_TIME_SECONDS = 300 # 5 minutes wait (300 seconds) between full scan cycles
DEFAULT_MAXHOLD_TIME_SECONDS = 3 # Default max hold time for the new argument


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

# Declare the comprehensive frequency_bands array (for PLOTTING MARKERS ONLY)
# This array will be used to define the PLOT_BAND_MARKERS for plotting.
frequency_bands_full_list = [
 
    [50.0, 54.0, 'AMATEURend '],
    [54.0, 72.0, 'BROADCASTINGend '],
    [72.0, 73.0, 'FIXEDend  MOBILEend '],
    [73.0, 74.6, 'RADIO ASTRONOMYend '],
    [74.6, 74.8, 'FIXEDend  MOBILEend '],
    [74.8, 75.2, 'AERONAUTICAL RADIONAVIGATIONend  5.180'],
    [75.2, 76.0, 'FIXEDend  MOBILEend '],
    [76.0, 108.0, 'BROADCASTINGend '],
    [108.0, 117.975, 'AERONAUTICAL RADIONAVIGATIONend  5.197A'],
    [117.975, 137.0, 'AERONAUTICAL MOBILEend  (R) 5.111 5.200'],
    [137.0, 138.0, 'METEOROLOGICAL-SATELLITEend  (space-to-Earth) MOBILE-SATELLITEend  (space-to-Earth) 5.208A 5.208B 5.209 SPACE OPERATIONend  (space-to-Earth) 5.203C 5.209A SPACE RESEARCHend  (space-to-Earth) 5.208'],
    [138.0, 144.0, 'FIXEDend  LAND MOBILEend  Space research (space-to-Earth)'],
    [144.0, 146.0, 'AMATEURend  AMATEUR-SATELLITEend '],
    [146.0, 148.0, 'AMATEURend '],
    [148.0, 149.9, 'FIXEDend  LAND MOBILEend  MOBILE-SATELLITEend  (Earth-to-space) 5.209 C26 5.218 5.218A 5.219'],
    [149.9, 150.05, 'MOBILE-SATELLITEend  (Earth-to-space) 5.209 5.220'],
    [150.05, 156.4875, 'MOBILEend  Fixed 5.226'],
    [156.4875, 156.5625, 'MARITIME MOBILEend  (distress and calling via DSCend ) 5.111 5.226 C32'],
    [156.5625, 156.7625, 'MOBILEend  Fixed 5.226'],
    [156.7625, 156.7875, 'MARITIME MOBILEend  MOBILE-SATELLITEend  (Earth-to-space) 5.111 5.226 5.228'],
    [156.7875, 156.8125, 'MARITIME MOBILEend  (distress and calling) 5.111 5.226'],
    [156.8125, 156.8375, 'MARITIME MOBILEend  MOBILE-SATELLITEend  (Earth-to-space) 5.111 5.226 5.228'],
    [156.8375, 157.1875, 'MOBILEend  Fixed 5.226'],
    [157.1875, 157.3375, 'MOBILEend  Fixed Maritime mobile-satellite 5.208A 5.208B 5.228AB 5.228AC 5.226'],
    [157.3375, 161.7875, 'MOBILEend  Fixed 5.226'],
    [161.7875, 161.9375, 'MOBILEend  Fixed Maritime mobile-satellite 5.208A 5.208B 5.228AB 5.228AC 5.226'],
    [161.9375, 161.9625, 'MOBILEend  Fixed Maritime mobile-satellite (Earth-to-space) 5.228AA 5.226'],
    [161.9625, 161.9875, 'AERONAUTICAL MOBILEend  (ORend ) MARITIME MOBILEend  MOBILE-SATELLITEend  (Earth-to-space) 5.228C 5.228D C53'],
    [161.9875, 162.0125, 'MOBILEend  Fixed Maritime mobile-satellite (Earth-to-space) 5.228AA 5.226'],
    [162.0125, 162.0375, 'AERONAUTICAL MOBILEend  (ORend ) MARITIME MOBILEend  MOBILE-SATELLITEend  (Earth-to-space) 5.228C 5.228D C53'],
    [162.0375, 174.0, 'MOBILEend  Fixed 5.226'],
    [174.0, 216.0, 'BROADCASTINGend '],
    [216.0, 219.0, 'FIXEDend  MARITIME MOBILEend  LAND MOBILEend  5.242'],
    [219.0, 220.0, 'FIXEDend  MARITIME MOBILEend  LAND MOBILEend  5.242 Amateur C11'],
    [220.0, 222.0, 'FIXEDend  MOBILEend  Amateur C11'],
    [222.0, 225.0, 'AMATEURend '],
    [225.0, 312.0, 'FIXEDend  MOBILEend  5.111 5.254 5.256 C5'],
    [312.0, 315.0, 'FIXEDend  MOBILEend  Mobile-satellite (Earth-to-space) 5.254 5.255 C5'],
    [315.0, 328.6, 'FIXEDend  MOBILEend  5.254 C5'],
    [328.6, 335.4, 'AERONAUTICAL RADIONAVIGATIONend  5.258'],
    [335.4, 387.0, 'FIXEDend  MOBILEend  5.254 C5'],
    [387.0, 390.0, 'FIXEDend  MOBILEend  Mobile-satellite (space-to-Earth) 5.208A 5.208B 5.254 5.255 C5'],
    [390.0, 399.9, 'FIXEDend  MOBILEend  5.254 C5'],
    [399.9, 400.05, 'MOBILE-SATELLITEend  (Earth-to-space) 5.209 5.220 5.260A 5.260B C19'],
    [400.05, 400.15, 'STANDARD FREQUENCY AND TIME SIGNAL-SATELLITEend  (400.1 MHz) 5.261'],
    [400.15, 401.0, 'METEOROLOGICAL AIDSend  METEOROLOGICAL-SATELLITEend  (space-to-Earth) MOBILE-SATELLITEend  (space-to-Earth) 5.208A 5.208B 5.209 SPACE OPERATIONend  (space-to-Earth) 5.203C 5.209A SPACE RESEARCHend  (space-to-Earth) 5.208'],
    [401.0, 402.0, 'METEOROLOGICAL AIDSend  SPACE OPERATIONend  (space-to-Earth) EARTH EXPLORATION-SATELLITEend  (Earth-to-space) METEOROLOGICAL-SATELLITEend  (Earth-to-space) Fixed Mobile except aeronautical mobile 5.264A 5.264B'],
    [402.0, 403.0, 'METEOROLOGICAL AIDSend  EARTH EXPLORATION-SATELLITEend  (Earth-to-space) METEOROLOGICAL-SATELLITEend  (Earth-to-space) Fixed Mobile except aeronautical mobile 5.264A 5.264B'],
    [403.0, 406.0, 'METEOROLOGICAL AIDSend  Fixed Mobile except aeronautical mobile 5.265'],
    [406.0, 406.1, 'MOBILE-SATELLITEend  (Earth-to-space) 5.265 5.266 5.267'],
    [406.1, 410.0, 'MOBILEend  except aeronautical mobile RADIO ASTRONOMYend  Fixed 5.149 5.265'],
    [410.0, 414.0, 'MOBILEend  except aeronautical mobile SPACE RESEARCHend  (space-to-space) 5.268 Fixed'],
    [414.0, 415.0, 'FIXEDend  SPACE RESEARCHend  (space-to-space) 5.268 Mobile except aeronautical mobile'],
    [415.0, 419.0, 'MOBILEend  except aeronautical mobile SPACE RESEARCHend  (space-to-space) 5.268 Fixed'],
    [419.0, 420.0, 'FIXEDend  SPACE RESEARCHend  (space-to-space) 5.268 Mobile except aeronautical mobile'],
    [420.0, 430.0, 'MOBILEend  except aeronautical mobile Fixed C10'],
    [430.0, 432.0, 'RADIOLOCATIONend  Amateur'],
    [432.0, 438.0, 'RADIOLOCATIONend  Amateur Earth Exploration-Satellite (active) 5.279A 5.282'],
    [438.0, 450.0, 'RADIOLOCATIONend  5.285 Amateur 5.284 5.286'],
    [450.0, 455.0, 'MOBILEend  5.286AAend  C23 Fixed 5.209 5.286 5.286A 5.286B 5.286C 5.286D C26A C26B'],
    [455.0, 456.0, 'FIXEDend  MOBILEend  5.286AAend  C23 MOBILE-SATELLITEend  (Earth-to-space) 5.209 5.286A 5.286B 5.286C C26A C26B'],
    [456.0, 459.0, 'MOBILEend  5.286AAend  5.287 C23 Fixed'],
    [459.0, 460.0, 'FIXEDend  MOBILEend  5.286AAend  C23 MOBILE-SATELLITEend  (Earth-to-space) 5.209 5.286A 5.286B 5.286C C26A C26B'],
    [460.0, 470.0, 'MOBILEend  5.286AAend  5.287 C23 Fixed 5.289'],
    [470.0, 608.0, 'BROADCASTINGend  5.293 5.295 5.297 C24 C24A'],
    [608.0, 614.0, 'RADIO ASTRONOMYend  Mobile-satellite except aeronautical mobile-satellite (Earth-to-space)'],
    [614.0, 698.0, 'FIXEDend  MOBILEend  BROADCASTINGend  5.293 5.308A C24 C24A'],
    [698.0, 806.0, 'FIXEDend  MOBILEend  5.317A C7 BROADCASTINGend  5.293'],
    [806.0, 890.0, 'MOBILEend  5.317A C7 Fixed 5.317 5.318'],
    [890.0, 902.0, 'FIXEDend  MOBILEend  except aeronautical mobile 5.317A C7 Radiolocation C5A 5.318'],
    [902.0, 928.0, 'FIXEDend  RADIOLOCATIONend  C5A Amateur Mobile except aeronautical mobile 5.150'],
    [928.0, 929.0, 'FIXEDend  MOBILEend  except aeronautical mobile 5.317A C7 Radiolocation C5A'],
    [929.0, 932.0, 'MOBILEend  except aeronautical mobile 5.317A C7 Fixed Radiolocation C5A'],
    [932.0, 932.5, 'FIXEDend  MOBILEend  except aeronautical mobile 5.317A C7 Radiolocation C5A'],
    [932.5, 935.0, 'FIXEDend  Mobile except aeronautical mobile 5.317A C7 Radiolocation C5A'],
    [935.0, 941.0, 'MOBILEend  except aeronautical mobile 5.317A C7 Fixed Radiolocation C5A'],
    [941.0, 941.5, 'FIXEDend  MOBILEend  except aeronautical mobile 5.317A C7 Radiolocation C5A'],
    [941.5, 942.0, 'FIXEDend  Mobile except aeronautical mobile 5.317A C7 Radiolocation C5A'],
    [942.0, 944.0, 'FIXEDend  Mobile 5.317A C7'],
    [944.0, 952.0, 'FIXEDend  MOBILEend  5.317A C7'],
    [952.0, 956.0, 'FIXEDend  MOBILEend  5.317A C7'],
    [956.0, 960.0, 'FIXEDend  Mobile 5.317A C7'],
    [960.0, 1164.0, 'AERONAUTICAL MOBILEend  (R) 5.327A AERONAUTICAL RADIONAVIGATIONend  5.328 5.328AA'],
    [1164.0, 1215.0, 'AERONAUTICAL RADIONAVIGATIONend  5.328 RADIONAVIGATION-SATELLITEend  (space-to-Earth) (space-to-space)  5.328B 5.328A'],
    [1215.0, 1240.0, 'EARTH EXPLORATION-SATELLITEend  (active) RADIOLOCATIONend  RADIONAVIGATION-SATELLITEend  (space-to-Earth) (space-to-space)  5.328B 5.329 5.329A SPACE RESEARCHend  (active) 5.332'],
    [1240.0, 1300.0, 'EARTH EXPLORATION-SATELLITEend  (active) RADIOLOCATIONend  RADIONAVIGATION-SATELLITEend  (space-to-Earth) (space-to-space)  5.328B 5.329 5.329A SPACE RESEARCHend  (active) Amateur 5.282 5.331 5.332 5.335 5.335A'],
    [1300.0, 1350.0, 'RADIOLOCATIONend  AERONAUTICAL RADIONAVIGATIONend  5.337 RADIONAVIGATION-SATELLITEend  (Earth-to-space) 5.149 5.337A'],
    [1350.0, 1390.0, 'FIXEDend  MOBILEend  RADIOLOCATIONend  5.149 5.338A 5.339 C5 C27'],
    [1390.0, 1400.0, 'FIXEDend  MOBILEend  5.149 5.339 C27B'],
    [1400.0, 1427.0, 'EARTH EXPLORATION-SATELLITEend  (passive) RADIO ASTRONOMYend  SPACE RESEARCHend  (passive) 5.340 5.341'],
    [1427.0, 1429.0, 'SPACE OPERATIONend  (Earth-to-space) FIXEDend  5.338A 5.341'],
    [1429.0, 1452.0, 'FIXEDend  MOBILEend  5.338A 5.341'],
    [1452.0, 1492.0, 'FIXEDend  MOBILEend  5.343 BROADCASTINGend  5.341 5.345'],
    [1492.0, 1525.0, 'FIXEDend  MOBILEend  5.341'],
    [1525.0, 1530.0, 'MOBILE-SATELLITEend  (space-to-Earth) 5.208B 5.351A Earth Exploration-Satellite Space operation (space-to-Earth) 5.341 5.351 5.354'],
    [1530.0, 1535.0, 'MOBILE-SATELLITEend  (space-to-Earth) 5.208B 5.351A 5.353A Earth Exploration-Satellite 5.341 5.351 5.354'],
    [1535.0, 1559.0, 'MOBILE-SATELLITEend  (space-to-Earth) 5.208B 5.351A 5.341 5.351 5.353A 5.354 5.356 5.357 5.357A'],
    [1559.0, 1610.0, 'AERONAUTICAL RADIONAVIGATIONend  RADIONAVIGATION-SATELLITEend  (space-to-Earth) (space-to-space)  5.208B 5.328B 5.329A 5.341'],
    [1610.0, 1610.6, 'MOBILE-SATELLITEend  (Earth-to-space) 5.351A AERONAUTICAL RADIONAVIGATIONend  5.341 5.364 5.366 5.367 5.368 5.372'],
    [1610.6, 1613.8, 'MOBILE-SATELLITEend  (Earth-to-space) 5.351A RADIO ASTRONOMYend  AERONAUTICAL RADIONAVIGATIONend  5.149 5.341 5.364 5.366 5.367 5.368 5.372'],
    [1613.8, 1621.35, 'MOBILE-SATELLITEend  (Earth-to-space) 5.351A AERONAUTICAL RADIONAVIGATIONend  Mobile-satellite (space-to-Earth) 5.208B 5.341 5.364 5.365 5.366 5.367 5.368 5.372'],
    [1621.35, 1626.5, 'MARITIME MOBILE-SATELLITEend  (space-to-Earth) 5.373 5.373A MOBILE-SATELLITE (Earth-to-space)end  5.351A AERONAUTICAL RADIONAVIGATIONend  Mobile-satellite (space-to-Earth) except maritime mobile-satellite (space-to-Earth) 5.208B 5.341 5.364 5.365 5.366 5.367 5.368 5.372'],
    [1626.5, 1660.0, 'MOBILE-SATELLITEend  (Earth-to-space) 5.351A 5.341 5.351 5.353A 5.354 5.357A 5.374 5.375 5.376'],
    [1660.0, 1660.5, 'MOBILE-SATELLITEend  (Earth-to-space) 5.351A RADIO ASTRONOMYend  5.149 5.341 5.351 5.354 5.376A'],
    [1660.5, 1668.0, 'RADIO ASTRONOMYend  SPACE RESEARCHend  (passive) Fixed 5.149 5.341 5.379A'],
    [1668.0, 1668.4, 'RADIO ASTRONOMYend  SPACE RESEARCHend  (passive) Fixed 5.149 5.341 5.379A'],
    [1668.4, 1670.0, 'METEOROLOGICAL AIDSend  FIXEDend  RADIO ASTRONOMYend  5.149 5.341 5.379D 5.379E'],
    [1670.0, 1675.0, 'METEOROLOGICAL AIDSend  FIXEDend  METEOROLOGICAL-SATELLITEend  (space-to-Earth) MOBILEend  except aeronautical mobile 5.341 5.379D 5.379E'],
    [1675.0, 1700.0, 'METEOROLOGICAL AIDSend  METEOROLOGICAL-SATELLITEend  (space-to-Earth) 5.289 5.341'],
    [1700.0, 1710.0, 'FIXEDend  METEOROLOGICAL-SATELLITEend  (space-to-Earth) 5.289 5.341'],
    [1710.0, 1755.0, 'FIXEDend  MOBILEend  5.384A 5.149 5.341 5.385 5.386'],
    [1755.0, 1780.0, 'FIXEDend  MOBILE 5.384A 5.386'],
    [1780.0, 1850.0, 'FIXEDend  Mobile 5.384A C5 5.386'],
    [1850.0, 2000.0, 'FIXEDend  MOBILEend  5.384A 5.388A 5.388 5.389B C35'],
    [2000.0, 2020.0, 'MOBILEend  MOBILE-SATELLITEend  (Earth-to-space) 5.351A 5.388 5.389A 5.389C 5.389E C36'],
    [2020.0, 2025.0, 'FIXEDend  MOBILEend  5.388 C37'],
    [2025.0, 2110.0, 'EARTH EXPLORATION-SATELLITEend  (Earth-to-space) (space-to-space) FIXEDend  SPACE OPERATIONend  (Earth-to-space) (space-to-space) SPACE RESEARCHend  (Earth-to-space) (space-to-space) Mobile 5.391 C5 5.392'],
    [2110.0, 2120.0, 'FIXEDend  MOBILEend  5.388A SPACE RESEARCHend  (deep space) (Earth-to-space) 5.388'],
    [2120.0, 2180.0, 'FIXEDend  MOBILEend  5.388A 5.388'],
    [2180.0, 2200.0, 'MOBILEend  MOBILE-SATELLITEend  (space-to-Earth) 5.351A 5.388 5.389A C36'],
    [2200.0, 2290.0, 'EARTH EXPLORATION-SATELLITEend  (space-to-Earth) (space-to-space) FIXEDend  SPACE OPERATIONend  (space-to-Earth) (space-to-space) SPACE RESEARCHend  (space-to-Earth) (space-to-space) Mobile 5.391 C5 5.392'],
    [2290.0, 2300.0, 'FIXEDend  SPACE RESEARCHend  (deep space) (Earth-to-space) Mobile C5'],
    [2300.0, 2450.0, 'FIXEDend  MOBILEend  5.384A 5.394 C34 RADIOLOCATIONend  Amateur 5.150 5.282 5.393 C12 C13 C13A C17'],
    [2450.0, 2483.5, 'FIXEDend  MOBILEend  RADIOLOCATIONend  5.150'],
    [2483.5, 2500.0, 'FIXEDend  C38 MOBILE-SATELLITEend  (space-to-Earth) 5.351A RADIOLOCATIONend  RADIODETERMINATION-SATELLITEend  (space-to-Earth) 5.398 5.150 5.402'],
    [2500.0, 2596.0, 'FIXEDend  MOBILEend  except aeronautical mobile 5.384A 5.416'],
    [2596.0, 2655.0, 'BROADCASTINGend  FIXEDend  MOBILEend  except aeronautical mobile 5.384A 5.339 5.416'],
    [2655.0, 2686.0, 'BROADCASTINGend  FIXEDend  MOBILEend  except aeronautical mobile 5.384A Earth Exploration-Satellite (passive) Radio astronomy Space research (passive) 5.149 5.416'],
    [2686.0, 2690.0, 'FIXEDend  MOBILEend  except aeronautical mobile 5.384A Earth Exploration-Satellite (passive) Radio astronomy Space research (passive) 5.149'],
    [2690.0, 2700.0, 'EARTH EXPLORATION-SATELLITEend  (passive) RADIO ASTRONOMYend  SPACE RESEARCHend  (passive) 5.340'],
    [2700.0, 2900.0, 'AERONAUTICAL RADIONAVIGATIONend  5.337 Radiolocation 5.423 5.424 C14 C54'],
    [2900.0, 3100.0, 'RADIOLOCATIONend  5.424A RADIONAVIGATIONend  5.426 5.425 5.427'],
    [3100.0, 3300.0, 'RADIOLOCATIONend  Earth Exploration-Satellite (active) Space research (active) 5.149'],
    [3300.0, 3450.0, 'RADIOLOCATIONend  5.433 C5 Amateur 5.149 5.282'],
]

# This list will be dynamically created for plotting purposes only, from frequency_bands_full_list
PLOT_BAND_MARKERS = []
for band_info in frequency_bands_full_list:
    PLOT_BAND_MARKERS.append({
        "Start MHz": band_info[0],
        "Stop MHz": band_info[1],
        "Band Name": band_info[2].strip() # Use strip() to remove leading/trailing spaces
    })



# Define VISA addresses for different users/instruments
VISA_ADDRESSES = {
    'apk': 'USB0::0x0957::0xFFEF::CN03480580::0::INSTR',
    'zap': 'USB1::0x0957::0xFFEF::SG05300002::0::INSTR' # Original user provided 'USB1', keeping it.
}

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

def setup_arguments():
    """
    Parses command-line arguments for the spectrum analyzer sweep.
    Allows customization of filename prefix, frequency range, step size, and user.
    """
    parser = argparse.ArgumentParser(description="Spectrum Analyzer Sweep and CSV Export")

    parser.add_argument('--name', type=str, default="Scan ",
                        help='Prefix for the output CSV filename')
    parser.add_argument('--start', type=float, default=None,
                        help='Start frequency in Hz (overrides default bands if provided with --endFreq)')
    parser.add_argument('--end', type=float, default=None,
                        help='End frequency in Hz (overrides default bands if provided with --startFreq)')
    parser.add_argument('--rbw', type=float, default=DEFAULT_RBW_STEP_SIZE_HZ, # Updated to new variable name
                        help='Step size in Hz')
    parser.add_argument('--user', type=str, choices=['apk', 'zap'], default='zap',
                        help='Specify who is running the program: "apk" or "zap". Default is "zap".')
    parser.add_argument('--hold', type=float, default=DEFAULT_MAXHOLD_TIME_SECONDS,
                        help='Duration in seconds for which MAX Hold should be active during scans. Set to 0 to disable. (Note: Instrument\'s MAX Hold is typically a continuous mode; this value serves as a flag for enablement during the entire scan duration).')
    # ADDED: Argument for cycle wait time
    parser.add_argument('--wait', type=float, default=DEFAULT_CYCLE_WAIT_TIME_SECONDS,
                        help=f'Wait time in seconds between full scan cycles. Default is {DEFAULT_CYCLE_WAIT_TIME_SECONDS} seconds.')

    return parser.parse_args()

def initialize_instrument(visa_address):
    """
    Establishes a VISA connection to the instrument and performs initial configuration.

    Args:
        visa_address (str): The VISA address of the spectrum analyzer.

    Returns:
        pyvisa.resources.Resource: The instrument object if connection is successful, else None.
    """
    rm = pyvisa.ResourceManager()
    inst = None
    try:
        inst = rm.open_resource(visa_address)
        inst.timeout = 5000 # Set timeout to 5 seconds for queries
        print(f"\nSuccessfully connected to: {query_safe(inst, '*IDN?')}\n")

        # Clear and reset the instrument
        write_safe(inst, "*CLS")
        write_safe(inst, "*RST")
        query_safe(inst, "*OPC?") # Wait for operations to complete
        print("Instrument cleared and reset.")

        # Configure preamplifier for high sensitivity
        write_safe(inst, ":SENS:POW:GAIN ON") # Equivalent to ':POWer:GAIN 1' for most Keysight instruments
        print("Preamplifier turned ON for high sensitivity.")

        # Configure display and marker settings
        write_safe(inst, ":DISP:WIND:TRAC:Y:SCAL LOG") # Logarithmic scale (dBm)
        write_safe(inst, ":DISP:WIND:TRAC:Y:SCAL:RLEVel -30DBM") # Reference level (corrected from :DISP:WIND:TRAC:Y:RLEV)

        # Turn on all six markers and set them to position mode (neutral state)
        # They will be used by the peak table functionality per segment.
        # Note: N9340B typically supports 4 markers. Attempting 6, but instrument might ignore extra.
        for i in range(1, 7): # Markers 1 through 6
            write_safe(inst, f":CALC:MARK{i}:STATe ON")
            write_safe(inst, f":CALC:MARK{i}:MODE POS")
            print(f"Marker {i} enabled in position mode (for peak table display).")

        print("Display set to logarithmic scale, reference level -30 dBm, and Markers 1-6 enabled.")

        return inst
    except pyvisa.VisaIOError as e:
        print(f"VISA Error: Could not connect to or communicate with the instrument at {visa_address}: {e}")
        print("Please ensure the instrument is connected, powered on, and the VISA address is correct.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred during instrument initialization: {e}")
        return None

def scan_bands(inst, csv_writer, max_hold_time, rbw):
    """
    Iterates through predefined frequency bands, sets the start/stop frequencies,
    reduces RBW to 10000 Hz, and triggers a sweep for each band.
    It extracts trace data, writes it directly to the provided CSV writer,
    and also returns it for further processing (plotting).
    This function now dynamically segments bands to maintain a consistent
    effective resolution bandwidth per trace point.

    Args:
        inst (pyvisa.resources.Resource): The PyVISA instrument object.
        csv_writer (csv.writer): The CSV writer object to write data to.
        max_hold_time (float): Duration in seconds for which MAX Hold should be active.
                                If > 0, MAX Hold mode is enabled for the scan.
    Returns:
        list: A list of dictionaries, where each dictionary represents a data point
              with 'Band Name', 'Frequency (Hz)', and 'Level (dBm)'.
    """
    all_scan_data = [] # To store all data points across all bands for plotting

    print("\n--- Starting Band Scan ---")

    # Configure MAX Hold mode based on max_hold_time argument
    if max_hold_time > 0:
        write_safe(inst, ":TRAC1:MODE MAXHold")
        print("Trace 1 set to MAX Hold mode for the duration of the scan.")
    else:
        # Ensure normal/clear write mode if max hold is not requested
        write_safe(inst, ":TRAC1:MODE WRITe")
        print("Trace 1 set to normal WRITE mode (MAX Hold disabled).")


    try:
        actual_sweep_points = int(float(query_safe(inst, ":SENS:SWE:POINts?")))
        print(f"Actual sweep points set by instrument: {actual_sweep_points}.")
    except ValueError:
        print(f"  Could not parse actual sweep points from instrument response: '{query_safe(inst, ":SENS:SWE:POINts?")}'. Defaulting to 401.")
        actual_sweep_points = 401 # Fallback to a common default if query fails

    if actual_sweep_points <= 1: # Ensure we have at least 2 points to calculate a span
        print(f"Warning: Instrument returned {actual_sweep_points} sweep points. Cannot effectively segment bands. Skipping scan.")
        return []

    # Calculate the optimal span for each segment to achieve desired RBW per point
    # We want (Segment Span / (Actual Points - 1)) = Desired RBW
    # So, Segment Span = Desired RBW * (Actual Points - 1)
    optimal_segment_span_hz = rbw * (actual_sweep_points - 1)
    print(f"Optimal segment span to achieve {DEFAULT_RBW_STEP_SIZE_HZ/1000:.0f} kHz effective RBW per point: {optimal_segment_span_hz / MHZ_TO_HZ:.3f} MHz.")


    # Set trace format to REAL (binary)
    write_safe(inst, ":TRAC:FORMat REAL")
    print("Set trace data format to REAL (binary) for efficient data transfer.")

    # Save current read_termination and encoding
    original_read_termination = inst.read_termination
    original_encoding = inst.encoding

    # Set read_termination to empty string for raw binary data reads.
    # This is crucial to prevent truncation if binary data contains bytes
    # that could be interpreted as termination characters.
    inst.read_termination = ''
    # Use latin-1 or iso-8859-1 for raw byte decoding if necessary,
    # though for binary, decoding should happen after struct.unpack
    inst.encoding = 'latin-1'

    # *** Use SCAN_BAND_RANGES for scanning the instrument ***
    for band in SCAN_BAND_RANGES:
        band_name = band["Band Name"]
        band_start_freq_hz = band["Start MHz"] * MHZ_TO_HZ
        band_stop_freq_hz = band["Stop MHz"] * MHZ_TO_HZ

        print(f"\nProcessing Band: {band_name} (Total Range: {band_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {band_stop_freq_hz/MHZ_TO_HZ:.3f} MHz)")

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

            print(f"  Scanning Segment {segment_counter}: {current_segment_start_freq_hz/MHZ_TO_HZ:.3f} MHz to {segment_stop_freq_hz/MHZ_TO_HZ:.3f} MHz (Span: {actual_segment_span_hz/MHZ_TO_HZ:.3f} MHz)")

            # Set instrument frequency range for the current segment
            write_safe(inst, f":SENS:FREQ:STAR {current_segment_start_freq_hz}")
            write_safe(inst, f":SENS:FREQ:STOP {segment_stop_freq_hz}")
            # Explicitly set span to ensure the instrument uses it for the sweep
            write_safe(inst, f":SENS:FREQ:SPAN {actual_segment_span_hz}")

            time.sleep(0.1) # Small delay to ensure commands are processed

            # Add settling time for max hold values to show up, if max hold is enabled
            if max_hold_time > 0:
                print(f"Waiting {max_hold_time} seconds for MAX hold to settle...", end='')
                for i in range(int(max_hold_time), 0, -1):
                    print(f"\rWaiting {i} seconds for MAX hold to settle...     ", end='') # \r to overwrite line
                    time.sleep(1)
                print("\rMAX hold settle time complete.                              ") # Clear the line after countdown

            query_safe(inst, "*OPC?") # Wait for the sweep to complete
            print(f"  Sweep completed for segment {segment_counter}.")

            # Read and process trace data
            trace_data = []
            raw_bytes = b''
            try:
                inst.write(":TRAC1:DATA?")
                raw_bytes = inst.read_raw()

                # Assuming raw binary float32 values without IEEE 488.2 header
                num_values_received = len(raw_bytes) // 4
                if num_values_received == 0:
                    print("    No trace data bytes received for this segment.")
                    current_segment_start_freq_hz = segment_stop_freq_hz
                    continue # Move to the next segment if no data

                trace_data = list(struct.unpack('<' + 'f' * num_values_received, raw_bytes))

                # Use the actual number of points received for frequency calculation for this segment
                num_trace_points_actual = len(trace_data)

                if num_trace_points_actual > 1:
                    freq_step_per_point_actual = actual_segment_span_hz / (num_trace_points_actual - 1)
                elif num_trace_points_actual == 1:
                    freq_step_per_point_actual = 0 # Single point, no step
                else:
                    freq_step_per_point_actual = 0 # No points

                # Loop to append data to all_scan_data and write to CSV
                for i, amp_value in enumerate(trace_data):
                    current_freq_for_point_hz = current_segment_start_freq_hz + (i * freq_step_per_point_actual)

                    # Append to list for plotting later
                    all_scan_data.append({
                        "Frequency (MHz)": current_freq_for_point_hz / MHZ_TO_HZ, # Store in MHz for DataFrame consistency
                        "Level (dBm)": amp_value,
                        "Band Name": band_name,

                    })

                    # Write directly to CSV file with desired order and units
                    csv_writer.writerow([
                        f"{current_freq_for_point_hz / MHZ_TO_HZ:.2f}",  # Frequency in MHz
                        f"{amp_value:.2f}",                                # Level in dBm
                        band_name,                                         # Band Name

                    ])

            except pyvisa.VisaIOError as e:
                print(f"    Error reading trace data (PyVISA IO Error): {e}")
                print(f"    Raw bytes potentially causing error: {raw_bytes[:100]}...")
            except ValueError as e:
                print(f"    Error processing binary data (ValueError): {e}")
                print(f"    Raw bytes potentially causing error: {raw_bytes[:100]}...")
            except struct.error as e:
                print(f"    Error unpacking binary data (Struct Error - incorrect format/length): {e}")
                print(f"    Raw bytes for struct unpack: {raw_bytes[:100]}...")
            except Exception as e:
                print(f"    An unexpected error occurred during trace processing: {e}")

            current_segment_start_freq_hz = segment_stop_freq_hz # Move to the start of the next segment

    # Restore original read_termination and encoding
    inst.read_termination = original_read_termination
    inst.encoding = original_encoding
    print("\n--- Band Scan Complete ---")
    return all_scan_data # Return the collected data


def plot_spectrum_data(df, output_html_filename):
    """
    Generates an interactive Plotly Express line plot from the spectrum analyzer data.
    The plot is saved as an HTML file.
    """
    print(f"\n--- Generating Interactive Plot: {output_html_filename} ---")

    # The DataFrame already contains 'Frequency (MHz)'
    # Create Interactive Plot with Plotly Express
    fig = px.line(df,
                  x="Frequency (MHz)",
                  y="Level (dBm)",
                  color="Band Name", # Differentiate lines by band name
                  title="Spectrum Analyzer Scans: Amplitude vs Frequency",
                  labels={"Frequency (MHz)": "Frequency (MHz)", "Level (dBm)": "Amplitude (dBm)"},
                  # Added Band Name to hover_data explicitly.
                  hover_data={"Frequency (MHz)": ':.2f', "Level (dBm)": ':.2f', "Band Name": True}
                 )

    # Determine the full Y-axis range *before* adding shapes,
    # ensuring it's what the plot will actually use.
    # It's better to calculate this once and use consistently.
    y_min_data = df['Level (dBm)'].min()
    y_max_data = df['Level (dBm)'].max()

    # Set fixed Y-axis range for the plot and for the band rectangles
    y_range_min = -100 # A reasonable lower bound for spectrum data
    y_range_max = 0    # As per your desired reference level max

    # Adjust y_range_min if data goes lower than -100 dBm
    if y_min_data < y_range_min:
        y_range_min = y_min_data - 10 # Provide some padding below the lowest point

    # Set Y-axis (Amplitude) Maximum to 0 dBm and apply range
    fig.update_yaxes(range=[y_range_min, y_range_max],
                     title="Amplitude (dBm)",
                     showgrid=True, gridwidth=1)

    # Define distinct colors for the markers to improve visibility
    marker_line_color = "rgba(255, 255, 0, 0.7)" # Bright yellow, semi-transparent
    marker_text_color = "yellow"

    # *** Modify this loop to use PLOT_BAND_MARKERS and plot only the start frequency ***
    for band in PLOT_BAND_MARKERS:
        # Add a vertical line at the start frequency of the band
        fig.add_shape(
            type="line",
            x0=band["Start MHz"],
            y0=y_range_min, # Span full Y-axis range
            x1=band["Start MHz"],
            y1=y_range_max, # Span full Y-axis range
            line=dict(
                color=marker_line_color,
                width=0.15, # Slightly thicker line
                dash="dash", # Use a dashed line
            ),
            layer="below", # Draw below the trace lines
        )
        # Add text markers for band names at the start frequency
        fig.add_annotation(
            x=band["Start MHz"] / 1000, # Position at the start of the band
            # Calculate Y position relative to the overall plot Y-range, 10% down from the top
            y=-80,
            text=band["Band Name"],
            showarrow=False,
            font=dict(
                size=9,
                color=marker_text_color # Use bright yellow for text
            ),
            # Keep text rotated for better readability with many markers
            textangle=-90
        )


    # Set X-axis (Frequency) to Logarithmic Scale
    fig.update_xaxes(type="log",
                     title="Frequency (MHz)",
                     showgrid=True, gridwidth=1,
                     tickformat = None) # Setting tickformat to None or "" often helps prevent scientific notation

    # Apply Dark Mode Theme
    fig.update_layout(template="plotly_dark")

    fig.write_html(output_html_filename, auto_open=True)
    print(f"\n--- Plotly Express Interactive Plot Generated and saved to {output_html_filename} ---")


def wait_with_interrupt(wait_time_seconds):
    """
    Provides a timed delay with user interruption capability.
    Allows skipping the wait, quitting the program, or resuming.

    Args:
        wait_time_seconds (int): The total time to wait in seconds.
    """
    print("\n" + "="*50)
    print(f"Next full scan cycle in {wait_time_seconds // 60} minutes and {wait_time_seconds % 60} seconds.")
    print("Press Ctrl+C at any time during the countdown to interact.")
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
            choice = input("Countdown interrupted. (S)kip wait, (Q)uit program, or (R)esume countdown? ").strip().lower()
            if choice == 's':
                skip_wait = True
                print("Skipping remaining wait time. Starting next scan shortly...")
                break
            elif choice == 'q':
                print("Exiting program.")
                sys.exit(0)
            else:
                print("Resuming countdown...")
        seconds_remaining -= 1

    if not skip_wait:
        sys.stdout.write("\r" + " "*50 + "\r") # Clear the last countdown line
        sys.stdout.flush()
        print("Starting next scan now!")

    print("\n" + "="*50 + "\n")


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
        print("\n--- Missing Dependencies ---")
        print("The following Python modules are required but not installed:")
        for module in missing_modules:
            print(f"- {module}")
        print("----------------------------")

        install_choice = input("Do you want to attempt to install them now? (y/n): ").strip().lower()
        if install_choice == 'y':
            print("Attempting to install missing modules...")
            for module in missing_modules:
                try:
                    print(f"Installing {module}...")
                    # Use sys.executable to ensure pip associated with current Python interpreter is used
                    subprocess.check_call([sys.executable, "-m", "pip", "install", module])
                    print(f"Successfully installed {module}.")
                except subprocess.CalledProcessError as e:
                    print(f"Error installing {module}: {e}")
                    print(f"Please try installing it manually by running: pip install {module}")
                    sys.exit(1) # Exit if an installation fails
                except Exception as e:
                    print(f"An unexpected error occurred during installation of {module}: {e}")
                    print(f"Please try installing it manually by running: pip install {module}")
                    sys.exit(1)
            print("All required modules installed successfully (or already present).")
        else:
            print("Installation declined. Please install the missing modules manually to run the script.")
            for module in missing_modules:
                print(f"pip install {module}")
            sys.exit(1) # Exit if installation is declined and modules are missing
    else:
        print("\nAll required Python modules are already installed.")


def main():
    """
    Main function to connect to the N9340B Spectrum Analyzer,
    run initial setup, perform band scans, and then read final configuration.
    This now runs in a continuous loop with an interruptible wait time.
    """
    # NEW: Check and install dependencies before proceeding
    check_and_install_dependencies()

    args = setup_arguments() # Parse command-line arguments

    # Determine instrument address based on user argument
    instrument_address = VISA_ADDRESSES.get(args.user)

    if not instrument_address:
        print(f"Error: No VISA address configured for user '{args.user}'.")
        print("Available devices:", pyvisa.ResourceManager().list_resources())
        return

    # Initialize the instrument outside the loop, as it should stay connected
    # for continuous operation unless a connection error occurs.
    inst = initialize_instrument(instrument_address)

    if inst is None:
        print("Instrument initialization failed. Exiting.")
        return

    # Main continuous scan loop
    try:
        scan_cycle_count = 0
        while True: # This loop makes the program repeat indefinitely
            scan_cycle_count += 1
            print(f"\n--- Starting Scan Cycle #{scan_cycle_count} ---")

            # File & directory setup for CSV and HTML plot
            scan_dir = os.path.join(os.getcwd(), "N9340 Scans")
            os.makedirs(scan_dir, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H-%M-%S")
            csv_filename = os.path.join(scan_dir, f"{args.name.strip()}_{timestamp}.csv")
            html_plot_filename = os.path.join(scan_dir, f"{args.name.strip()}_{timestamp}.html")

            try:
                print(f"Opening CSV file for writing: {csv_filename}")
                with open(csv_filename, mode='w', newline='') as csvfile:
                    csv_writer = csv.writer(csvfile)
                    # Write header row to CSV with desired order and units
                    csv_writer.writerow(["Frequency (MHz)", "Level (dBm)", "Band Name"])

                    # After successful initialization, proceed with scanning bands
                    # scan_bands now takes the csv_writer and still returns the collected data for plotting
                    all_scan_data = scan_bands(inst, csv_writer, args.hold, args.rbw)

                if not all_scan_data:
                    print("No scan data collected in this cycle. Skipping plotting.")
                else:
                    # Convert collected data to pandas DataFrame
                    df = pd.DataFrame(all_scan_data)
                    # Call the plotting function
                    plot_spectrum_data(df, html_plot_filename)

            except Exception as e:
                print(f"An error occurred during scan cycle #{scan_cycle_count}: {e}")
                # For now, it will log and continue to the wait period.

            # CALLING THE INTERRUPTIBLE WAIT FUNCTION
            wait_with_interrupt(args.wait) # Uses the wait time from command-line arguments

    except KeyboardInterrupt:
        print("\nProgram interrupted by user (Ctrl+C) outside of wait period. Exiting.")
    except Exception as e:
        print(f"An unexpected critical error occurred in the main loop: {e}")
    finally:
        if inst and inst.session: # Check if inst object exists and has an active session
            inst.close()
            print("\nConnection to N9340B closed.")

if __name__ == '__main__':
    main()
