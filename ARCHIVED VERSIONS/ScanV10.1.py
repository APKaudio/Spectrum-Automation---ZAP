# main_app.py
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
import re # Import re module for regex in connect_instrument

# Import the frequency band definitions from the new file
# Assuming frequency_bands.py exists in the same directory
try:
    from frequency_bands import (
        MHZ_TO_HZ,
        SCAN_BAND_RANGES,
        TV_PLOT_BAND_MARKERS,
        GOV_PLOT_BAND_MARKERS
    )
except ImportError:
    print("Error: frequency_bands.py not found. Please ensure it's in the same directory.")
    # Define dummy values to prevent errors if file is missing
    MHZ_TO_HZ = 1_000_000
    SCAN_BAND_RANGES = [
        {"Band Name": "Dummy Band 1", "Start MHz": 100, "Stop MHz": 200},
        {"Band Name": "Dummy Band 2", "Start MHz": 400, "Stop MHz": 500}
    ]
    TV_PLOT_BAND_MARKERS = []
    GOV_PLOT_BAND_MARKERS = []


# Updated wait time variable and its usage for the continuous loop
DEFAULT_RBW_STEP_SIZE_HZ = 10000 # 10 kHz RBW resolution desired per data point
DEFAULT_CYCLE_WAIT_TIME_SECONDS = 30 # 30 seconds wait between full scan cycles
DEFAULT_MAXHOLD_TIME_SECONDS = 5 # Default max hold time for the new argument

# --- Utility Functions ---

def check_and_install_dependencies():
    """Checks for necessary libraries and installs them if missing."""
    dependencies = ['pyvisa', 'pandas', 'plotly', 'numpy']
    installed = True
    for dep in dependencies:
        try:
            __import__(dep)
        except ImportError:
            print(f"Installing missing dependency: {dep}")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", dep])
                print(f"Successfully installed {dep}.")
            except Exception as e:
                print(f"Failed to install {dep}. Please install it manually. Error: {e}")
                installed = False
    return installed

def query_safe(inst, command):
    """Safely queries the instrument, handling VISA errors."""
    try:
        response = inst.query(command).strip()
        print(f"Query: '{command}' -> Response: '{response}'") # Log query and response
        return response
    except pyvisa.errors.VisaIOError as e:
        print(f"VISA Error during query '{command}': {e}")
        return None
    except Exception as e:
        print(f"Error parsing response for '{command}': {e}")
        return None

def write_safe(inst, command):
    """Safely writes to the instrument, handling VISA errors."""
    try:
        inst.write(command)
        print(f"Write: '{command}'") # Log write command
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"VISA Error during write '{command}': {e}")
        return False

# --- Instrument Control Functions ---

def initialize_instrument(inst, ref_level_dbm, preamp_on, display_log_scale, rbw_config_val, vbw_config_val, instrument_model):
    """Initializes the Keysight N9340B spectrum analyzer with basic settings."""
    print("✨ Initializing instrument...")
    try:
        write_safe(inst, "*RST") # Reset to known state
        query_safe(inst, "*OPC?") # Wait for operation complete
        time.sleep(1) # Give it a moment

        write_safe(inst, "SYSTem:DISPlay:UPDate ON") # Ensure display updates
        query_safe(inst, "*OPC?")

        # Set reference level
        write_safe(inst, f":DISPlay:WINDow:TRACe:Y:RLEVel {ref_level_dbm}DBM")
        query_safe(inst, "*OPC?")
        print(f"✅ Set reference level to {ref_level_dbm} dBm.")

        # Set preamplifier
        if preamp_on:
            write_safe(inst, ":SENSe:POWer:RF:GAIN:STATe ON") # Corrected SCPI command for preamp
            print("✅ Preamplifier ON.")
        else:
            write_safe(inst, ":SENSe:POWer:RF:GAIN:STATe OFF") # Corrected SCPI command for preamp
            print("✅ Preamplifier OFF.")
        query_safe(inst, "*OPC?")

        # Set display scale
        if display_log_scale:
            write_safe(inst, ":DISPlay:WINDow:TRACe:Y:SCALe:SPACing LOGarithmic")
            print("✅ Display scale set to LOG.")
        else:
            write_safe(inst, ":DISPlay:WINDow:TRACe:Y:SCALe:SPACing LINear")
            print("✅ Display scale set to LIN.")
        query_safe(inst, "*OPC?")

        # Set RBW and VBW
        # Ensure RBW and VBW auto coupling is off if manual values are set
        write_safe(inst, ":SENSe:BANDwidth:RESolution:AUTO OFF")
        write_safe(inst, f":SENSe:BANDwidth:RESolution {rbw_config_val}HZ")
        query_safe(inst, "*OPC?")

        write_safe(inst, ":SENSe:BANDwidth:VIDeo:AUTO OFF")
        write_safe(inst, f":SENSe:BANDwidth:VIDeo {vbw_config_val}HZ")
        query_safe(inst, "*OPC?")
        print(f"✅ Set RBW to {rbw_config_val} Hz, VBW to {vbw_config_val} Hz.")
        
        # Set sweep time to auto
        write_safe(inst, ":SENSe:SWEep:TIME:AUTO ON")
        query_safe(inst, "*OPC?")
        print("✅ Sweep time set to AUTO.")

        # Configure trace 1 to clear write
        write_safe(inst, ":TRACe1:MODE WRITe") # Corrected SCPI command for trace mode
        query_safe(inst, "*OPC?")
        print("✅ Trace mode set to CLEAR/WRITE.")
        
        # Configure Markers (N9340B typically supports 6 markers)
        for i in range(1, 7): # Enable and configure markers 1 to 6
            write_safe(inst, f":CALCulate:MARKer{i}:STATe ON")
            query_safe(inst, "*OPC?")
            write_safe(inst, f":CALCulate:MARKer{i}:MODE POSition") # Set to normal position mode
            query_safe(inst, "*OPC?")
            # Marker RBW is not a direct setting for individual markers in N9340B,
            # it's usually tied to the main RBW. Removed this line.
        print("✅ Markers 1-6 enabled and configured to position mode.")

        print("🎉 Instrument initialized successfully.")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"🛑 Failed to initialize instrument: {e}")
        return False

def scan_bands(app_instance, bands_to_scan, rbw, vbw, max_hold_time):
    """
    Scans selected frequency bands using the :TRACe:DATA? TRACE1 command.
    Collects frequency and power data.
    """
    app_instance.all_scanned_data = [] # Clear previous scan data
    app_instance.last_scan_data = []
    
    # Assuming N9340B has fixed 401 sweep points
    fixed_sweep_points = 401 

    current_scan_progress = 0
    total_segments = len(bands_to_scan)
    print(f"Starting scan of {total_segments} bands...")

    for i, band in enumerate(bands_to_scan):
        if not app_instance.scanning: # Allow stopping scan from GUI
            print("\nScan interrupted by user.")
            break

        band_name = band["Band Name"]
        start_freq_hz = band["Start MHz"] * MHZ_TO_HZ
        stop_freq_hz = band["Stop MHz"] * MHZ_TO_HZ
        
        print(f"\nScanning band: {band_name} ({band['Start MHz']:.3f} MHz - {band['Stop MHz']:.3f} MHz)")

        try:
            # Set instrument to a known state for this segment
            write_safe(app_instance.inst, ":TRACe1:MODE WRITe") # Clear/Write trace
            query_safe(app_instance.inst, "*OPC?")
            
            # Set frequency range for the current segment
            write_safe(app_instance.inst, f":SENSe:FREQuency:STARt {start_freq_hz}")
            write_safe(app_instance.inst, f":SENSe:FREQuency:STOP {stop_freq_hz}")
            write_safe(app_instance.inst, f":SENSe:SWEep:POINts {fixed_sweep_points}") # Ensure instrument is set to 401 points
            query_safe(app_instance.inst, "*OPC?")
            
            # Read back sweep points just to confirm (though it should be 401)
            sweep_points_actual = int(query_safe(app_instance.inst, ":SENSe:SWEep:POINTS?")) # Corrected query
            if sweep_points_actual != fixed_sweep_points:
                print(f"⚠️ Warning: Instrument sweep points are {sweep_points_actual}, expected {fixed_sweep_points}.")


            # Implement Max Hold logic
            current_trace_amplitudes = None
            if max_hold_time > 0:
                print(f"Applying max hold for {max_hold_time} seconds on this segment...")
                max_hold_end_time = time.time() + max_hold_time
                while time.time() < max_hold_end_time:
                    if not app_instance.scanning: # Allow stopping scan from GUI during max hold
                        print("\nMax hold interrupted by user.")
                        break
                    
                    write_safe(app_instance.inst, ":INITiate:IMMediate") # Trigger a sweep
                    query_safe(app_instance.inst, "*OPC?") # Wait for operation complete
                    
                    # Get current trace data
                    current_sweep_data_str = query_safe(app_instance.inst, ":TRACe1:DATA?") # Query trace 1 data
                    if current_sweep_data_str is None:
                        print("Error: Could not retrieve trace data for max hold. Skipping.")
                        break

                    try:
                        sweep_amplitudes_array = np.array([float(x) for x in current_sweep_data_str.split(',')])
                    except ValueError as e:
                        print(f"Error parsing trace data during max hold: {e}. Data: {current_sweep_data_str[:100]}...")
                        break

                    if current_trace_amplitudes is None:
                        current_trace_amplitudes = sweep_amplitudes_array
                    else:
                        current_trace_amplitudes = np.maximum(current_trace_amplitudes, sweep_amplitudes_array)
                    
                    time.sleep(0.1) # Small delay between sweeps for max hold
            else:
                # Perform a single sweep if no max hold
                write_safe(app_instance.inst, ":INITiate:IMMediate")
                query_safe(app_instance.inst, "*OPC?") # Wait for operation complete
                trace_data_str = query_safe(app_instance.inst, ":TRACe1:DATA?") # Query trace 1 data
                if trace_data_str is None:
                    print("Error: Could not retrieve trace data. Skipping band.")
                    continue
                try:
                    current_trace_amplitudes = np.array([float(x) for x in trace_data_str.split(',')])
                except ValueError as e:
                    print(f"Error parsing trace data: {e}. Data: {trace_data_str[:100]}...")
                    continue


            # Generate frequency points corresponding to the trace
            if current_trace_amplitudes is not None and len(current_trace_amplitudes) == fixed_sweep_points: # Use fixed_sweep_points
                freq_points = np.linspace(start_freq_hz, stop_freq_hz, fixed_sweep_points)
                
                # Append data to the global list
                for j in range(fixed_sweep_points):
                    app_instance.all_scanned_data.append((freq_points[j], current_trace_amplitudes[j]))
                print(f"✅ Collected {len(current_trace_amplitudes)} points for {band_name}.")
            else:
                print(f"❌ Data mismatch for {band_name}: Expected {fixed_sweep_points} points, got {len(current_trace_amplitudes) if current_trace_amplitudes is not None else 'None'}.")
            

        except pyvisa.errors.VisaIOError as e:
            print(f"🛑 VISA Error during scan of {band_name}: {e}")
            messagebox.showerror("VISA Error", f"Disconnected from instrument during scan: {e}")
            app_instance.inst = None # Indicate disconnection
            break # Exit the scan loop
        except Exception as e:
            print(f"🛑 An unexpected error occurred during scan of {band_name}: {e}")
            messagebox.showerror("Scan Error", f"An unexpected error occurred during scan: {e}")
            break # Exit the scan loop
        
        current_scan_progress += 1
        app_instance.update_progress_label(f"Scanning... {current_scan_progress}/{total_segments} bands complete")
    
    app_instance.last_scan_data = app_instance.all_scanned_data
    if app_instance.last_scan_data:
        print(f"\nScan complete. Total data points collected: {len(app_instance.last_scan_data)}")
    else:
        print("\nScan completed with no data collected.")
    return app_instance.all_scanned_data

def plot_data(scanned_data, csv_file_path):
    """
    Generates an interactive Plotly HTML plot from scanned data,
    including overlays for TV and Government frequency bands.
    """
    if not scanned_data:
        print("No data to plot.")
        return None

    df = pd.DataFrame(scanned_data, columns=['Frequency_Hz', 'Power_dBm'])
    df['Frequency_MHz'] = df['Frequency_Hz'] / MHZ_TO_HZ

    fig = px.line(df, x='Frequency_MHz', y='Power_dBm', 
                  title=f'RF Spectrum Scan ({os.path.basename(csv_file_path)})',
                  labels={'Frequency_MHz': 'Frequency (MHz)', 'Power_dBm': 'Power (dBm)'},
                  line_shape='linear') # 'linear' connects points with straight lines

    # Add vertical rectangles for TV bands
    for band in TV_PLOT_BAND_MARKERS:
        fig.add_shape(type="rect",
                      xref="x", yref="paper",
                      x0=band["Start MHz"], y0=0, x1=band["Stop MHz"], y1=1,
                      line=dict(color="RoyalBlue", width=1),
                      fillcolor="RoyalBlue",
                      opacity=0.2,
                      layer="below",
                      name=band["Band Name"], # For hover info if added later
                      hovertemplate=f"<b>TV Band: {band['Band Name']}</b><br>Range: {band['Start MHz']}-{band['Stop MHz']} MHz<extra></extra>"
                     )
        # Add text label for TV bands
        fig.add_annotation(
            x=(band["Start MHz"] + band["Stop MHz"]) / 2,
            y=1.00, # Position at the top of the plot
            xref="x",
            yref="paper",
            text=f"TV {band['Band Name'].split(' ')[-1]}",
            showarrow=False,
            font=dict(color="RoyalBlue", size=8),
            textangle=-90,
            yanchor="top",
        )

    # Add vertical rectangles for Government bands
    for band in GOV_PLOT_BAND_MARKERS:
        fig.add_shape(type="rect",
                      xref="x", yref="paper",
                      x0=band["Start MHz"], y0=0, x1=band["Stop MHz"], y1=1,
                      line=dict(color="Red", width=1),
                      fillcolor="Red",
                      opacity=0.1,
                      layer="below",
                      name=band["Band Name"], # For hover info if added later
                      hovertemplate=f"<b>Gov Band: {band['Band Name']}</b><br>Range: {band['Start MHz']}-{band['Stop MHz']} MHz<extra></extra>"
                     )
        # Add text label for Gov bands
        fig.add_annotation(
            x=(band["Start MHz"] + band["Stop MHz"]) / 2,
            y=0.95, # Position slightly below TV labels
            xref="x",
            yref="paper",
            text=band["Band Name"],
            showarrow=False,
            font=dict(color="Red", size=7),
            textangle=-90,
            yanchor="top",
        )


    fig.update_layout(hovermode="x unified") # Show all traces on hover

    plot_filename = f"spectrum_scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    plot_path = os.path.join(os.getcwd(), plot_filename)
    fig.write_html(plot_path)
    return plot_path

# --- GUI Classes ---

class TextRedirector(object):
    """A class to redirect stdout/stderr to a Tkinter scrolled text widget."""
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag
        self.last_char_was_cr = False

    def write(self, str_val):
        self.widget.config(state=tk.NORMAL) # Enable editing
        
        # Handle carriage returns for in-place updates (like progress bars)
        if '\r' in str_val:
            parts = str_val.split('\r')
            for i, part in enumerate(parts):
                if self.last_char_was_cr and i == 0:
                    # If last char was CR, delete current line and then write
                    self.widget.delete("end-1c linestart", "end-1c")
                self.widget.insert(tk.END, part, (self.tag,))
                self.widget.see(tk.END)
                if i < len(parts) - 1: # If not the last part, it means there was a \r
                    self.last_char_was_cr = True
                else:
                    self.last_char_was_cr = False
        else:
            self.widget.insert(tk.END, str_val, (self.tag,))
            self.widget.see(tk.END)
            self.last_char_was_cr = False

        self.widget.config(state=tk.DISABLED) # Disable editing
        self.widget.update_idletasks() # Force update

    def flush(self):
        # This is typically called after write; ensures content is displayed
        # For a ScrolledText widget with update_idletasks, this might not need explicit action.
        pass

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RF Spectrum Analyzer Controller")
        self.geometry("1200x700") 

        self.rm = pyvisa.ResourceManager()
        self.instrument_list = []
        self.inst = None
        self.scanning = False # Flag to control scanning thread
        self.last_scan_data = None # Store the last scan data for plotting
        self.last_csv_file_path = None # Store path of last saved CSV

        # Initialize Tkinter variables for Device Configuration
        self.current_ref_level_var = tk.StringVar(self, value="N/A")
        self.current_preamp_var = tk.BooleanVar(self, value=False)
        self.current_log_scale_var = tk.BooleanVar(self, value=False)
        self.current_rbw_var = tk.StringVar(self, value="N/A")
        self.current_vbw_var = tk.StringVar(self, value="N/A")
        self.current_sweep_time_auto_var = tk.BooleanVar(self, value=False)
        self.current_start_freq_var = tk.StringVar(self, value="N/A")
        self.current_stop_freq_var = tk.StringVar(self, value="N/A")
        self.current_sweep_points_var = tk.StringVar(self, value="N/A") # For fixed sweep points

        # Initialize Tkinter variables for Scan Configuration (these are user inputs)
        self.desired_ref_level_var = tk.StringVar(self, value="-30")
        self.desired_preamp_var = tk.BooleanVar(self, value=True)
        self.desired_log_scale_var = tk.BooleanVar(self, value=True)
        self.desired_max_hold_var = tk.BooleanVar(self, value=False)
        self.desired_max_hold_time_var = tk.StringVar(self, value=str(DEFAULT_MAXHOLD_TIME_SECONDS))
        self.desired_rbw_var = tk.StringVar(self, value=str(DEFAULT_RBW_STEP_SIZE_HZ))
        self.desired_vbw_display_var = tk.StringVar(self, value=str(int(DEFAULT_RBW_STEP_SIZE_HZ / 3))) # Display only
        self.desired_cycle_wait_time_var = tk.StringVar(self, value=str(DEFAULT_CYCLE_WAIT_TIME_SECONDS))
        self.output_folder_var = tk.StringVar(self, value="scan_data")

        self.resource_var = tk.StringVar(self) # For VISA resource dropdown

        # Create two main frames: one for the GUI, one for the console output
        self.main_frame = tk.Frame(self)
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.console_frame = tk.Frame(self, width=700, bg="black")
        self.console_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.console_frame.pack_propagate(False)

        self.console_control_frame = tk.Frame(self.console_frame, bg="black")
        self.console_control_frame.pack(side=tk.TOP, fill=tk.X, pady=5)

        self.start_scan_button = tk.Button(self.console_control_frame, text="Start Scan", command=self.start_scan_thread, state=tk.DISABLED, bg="green", fg="white", height=2)
        self.start_scan_button.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill=tk.X)

        self.stop_scan_button = tk.Button(self.console_control_frame, text="Stop Scan", command=self.stop_scan, state=tk.DISABLED, bg="red", fg="white", height=2)
        self.stop_scan_button.pack(side=tk.RIGHT, padx=5, pady=5, expand=True, fill=tk.X)

        self.console_output = scrolledtext.ScrolledText(self.console_frame, wrap=tk.WORD, bg="black", fg="white", font=("Consolas", 10))
        self.console_output.pack(expand=True, fill=tk.BOTH)
        self.console_output.configure(state="disabled")
        
        sys.stdout = TextRedirector(self.console_output, "stdout")
        sys.stderr = TextRedirector(self.console_output, "stderr")

        print("--- RF Spectrum Scanner GUI Initialized ---")

        self.create_widgets()
        self.populate_resources()
        self.actual_sweep_points = 401 # Fixed sweep points for N9340B

    def create_widgets(self):
        # Resource selection
        resource_frame = tk.LabelFrame(self.main_frame, text="Instrument Connection", padx=10, pady=10)
        resource_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(resource_frame, text="VISA Resource:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.resource_dropdown = tk.OptionMenu(resource_frame, self.resource_var, "No Resources Found")
        self.resource_dropdown.grid(row=0, column=1, sticky=tk.EW, pady=2)
        self.refresh_button = tk.Button(resource_frame, text="Refresh", command=self.populate_resources)
        self.refresh_button.grid(row=0, column=2, padx=5, pady=2)
        self.connect_button = tk.Button(resource_frame, text="Connect", command=self.connect_instrument)
        self.connect_button.grid(row=0, column=3, padx=5, pady=2)

        # Device Configuration Section (Queried from Instrument)
        device_config_frame = tk.LabelFrame(self.main_frame, text="Current Device Configuration (Queried)", padx=10, pady=10)
        device_config_frame.pack(pady=10, padx=10, fill=tk.X)
        device_config_frame.columnconfigure(1, weight=1) # Allow second column to expand

        row = 0
        tk.Label(device_config_frame, text="Reference Level (dBm):").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Entry(device_config_frame, textvariable=self.current_ref_level_var, state='readonly').grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        tk.Label(device_config_frame, text="Preamplifier ON:").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Checkbutton(device_config_frame, variable=self.current_preamp_var, state='disabled').grid(row=row, column=1, sticky=tk.W, pady=2)
        row += 1

        tk.Label(device_config_frame, text="Logarithmic Scale:").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Checkbutton(device_config_frame, variable=self.current_log_scale_var, state='disabled').grid(row=row, column=1, sticky=tk.W, pady=2)
        row += 1
        
        tk.Label(device_config_frame, text="Resolution Bandwidth (Hz):").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Entry(device_config_frame, textvariable=self.current_rbw_var, state='readonly').grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        tk.Label(device_config_frame, text="Video Bandwidth (Hz):").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Entry(device_config_frame, textvariable=self.current_vbw_var, state='readonly').grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        tk.Label(device_config_frame, text="Sweep Time Auto:").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Checkbutton(device_config_frame, variable=self.current_sweep_time_auto_var, state='disabled').grid(row=row, column=1, sticky=tk.W, pady=2)
        row += 1

        tk.Label(device_config_frame, text="Start Frequency (Hz):").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Entry(device_config_frame, textvariable=self.current_start_freq_var, state='readonly').grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        tk.Label(device_config_frame, text="Stop Frequency (Hz):").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Entry(device_config_frame, textvariable=self.current_stop_freq_var, state='readonly').grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        tk.Label(device_config_frame, text="Sweep Points:").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Entry(device_config_frame, textvariable=self.current_sweep_points_var, state='readonly').grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        # Scan Configuration Section (User Input)
        scan_config_frame = tk.LabelFrame(self.main_frame, text="Scan Configuration (User Input)", padx=10, pady=10)
        scan_config_frame.pack(pady=10, padx=10, fill=tk.X)
        scan_config_frame.columnconfigure(1, weight=1) # Allow second column to expand

        row = 0
        tk.Label(scan_config_frame, text="Desired Reference Level (dBm):").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Entry(scan_config_frame, textvariable=self.desired_ref_level_var).grid(row=row, column=1, sticky=tk.EW, pady=2)
        self.desired_ref_level_var.trace_add("write", self.on_setting_change)
        row += 1

        tk.Checkbutton(scan_config_frame, text="Desired Preamplifier ON", variable=self.desired_preamp_var, command=self.on_setting_change).grid(row=row, column=0, sticky=tk.W, pady=2, columnspan=2)
        row += 1

        tk.Checkbutton(scan_config_frame, text="Desired Logarithmic Scale", variable=self.desired_log_scale_var, command=self.on_setting_change).grid(row=row, column=0, sticky=tk.W, pady=2, columnspan=2)
        row += 1
        
        tk.Label(scan_config_frame, text="Desired RBW (Hz):").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Entry(scan_config_frame, textvariable=self.desired_rbw_var).grid(row=row, column=1, sticky=tk.EW, pady=2)
        self.desired_rbw_var.trace_add("write", self.update_vbw_display_and_reinitialize)
        row += 1

        tk.Label(scan_config_frame, text="Desired VBW (Hz) (RBW/3):").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Entry(scan_config_frame, textvariable=self.desired_vbw_display_var, state='readonly').grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        tk.Checkbutton(scan_config_frame, text="Max Hold Trace 1", variable=self.desired_max_hold_var).grid(row=row, column=0, sticky=tk.W, pady=2, columnspan=2)
        row += 1
        
        tk.Label(scan_config_frame, text="Max Hold Time (s):").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Entry(scan_config_frame, textvariable=self.desired_max_hold_time_var).grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        tk.Label(scan_config_frame, text="Scan Cycle Wait Time (s):").grid(row=row, column=0, sticky=tk.W, pady=2)
        tk.Entry(scan_config_frame, textvariable=self.desired_cycle_wait_time_var).grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        # Output folder and filename
        output_frame = tk.LabelFrame(scan_config_frame, text="Output Settings", padx=10, pady=10)
        output_frame.grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=10) # Place inside scan_config_frame
        output_frame.columnconfigure(1, weight=1)

        tk.Label(output_frame, text="Output Folder:").grid(row=0, column=0, sticky=tk.W, pady=2)
        tk.Entry(output_frame, textvariable=self.output_folder_var).grid(row=0, column=1, sticky=tk.EW, pady=2)
        row += 1

        # Band Selection
        band_selection_frame = tk.LabelFrame(self.main_frame, text="Frequency Band Selection", padx=10, pady=10)
        band_selection_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        self.band_checkboxes = []
        self.band_vars = []

        band_canvas = tk.Canvas(band_selection_frame)
        band_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        band_scrollbar = tk.Scrollbar(band_selection_frame, orient="vertical", command=band_canvas.yview)
        band_scrollbar.pack(side=tk.RIGHT, fill="y")

        band_canvas.configure(yscrollcommand=band_scrollbar.set)
        band_canvas.bind('<Configure>', lambda e: band_canvas.configure(scrollregion = band_canvas.bbox("all")))

        self.inner_band_frame = tk.Frame(band_canvas)
        band_canvas.create_window((0, 0), window=self.inner_band_frame, anchor="nw")

        for i, band in enumerate(SCAN_BAND_RANGES):
            var = tk.BooleanVar(self)
            chk = tk.Checkbutton(self.inner_band_frame, text=f"{band['Band Name']} ({band['Start MHz']:.3f}-{band['Stop MHz']:.3f} MHz)", variable=var)
            chk.grid(row=i, column=0, sticky=tk.W, padx=5, pady=1)
            var.set(True) # Default all bands to selected
            self.band_checkboxes.append(chk)
            self.band_vars.append({"band": band, "var": var}) # Store band info with var
        
        self.progress_label = tk.Label(self.main_frame, text="Ready.")
        self.progress_label.pack(pady=5)

        self.plot_button = tk.Button(self.main_frame, text="Generate Plot", command=self.generate_plot)
        self.plot_button.pack(pady=10)
        self.plot_button.pack_forget()

    def update_vbw_display_and_reinitialize(self, *args):
        """Updates VBW display and triggers instrument reinitialization."""
        self.update_vbw_display()
        self.on_setting_change()

    def on_setting_change(self, *args):
        """Reinitializes the instrument when certain settings are changed."""
        if self.inst: # Only reinitialize if already connected
            print("Settings changed. Reinitializing instrument...")
            # Call connect_instrument to re-run initialization with current settings
            # This will also trigger query_instrument_settings after initialization
            self.connect_instrument() 

    def update_vbw_display(self):
        try:
            rbw_val = float(self.desired_rbw_var.get())
            vbw_val = rbw_val / 3.0 # VBW is typically RBW/3
            self.desired_vbw_display_var.set(f"{int(vbw_val)}")
        except ValueError:
            self.desired_vbw_display_var.set("Invalid RBW")

    def populate_resources(self):
        try:
            self.instrument_list = self.rm.list_resources()
            if self.instrument_list:
                self.resource_var.set(self.instrument_list[0])
                menu = self.resource_dropdown["menu"]
                menu.delete(0, "end")
                for resource in self.instrument_list:
                    menu.add_command(label=resource, command=tk._setit(self.resource_var, resource))
            else:
                self.resource_var.set("No resources found")
                menu = self.resource_dropdown["menu"]
                menu.delete(0, "end")
                menu.add_command(label="No resources found", command=tk._setit(self.resource_var, "No resources found"))
            self.start_scan_button.config(state=tk.DISABLED) # Disable until connected
        except Exception as e:
            messagebox.showerror("Error", f"Failed to list VISA resources: {e}")
            self.resource_var.set("Error listing resources")
            self.start_scan_button.config(state=tk.DISABLED)

    def connect_instrument(self):
        selected_resource = self.resource_var.get()
        if selected_resource == "No resources found" or "Error listing resources" in selected_resource:
            messagebox.showwarning("Connection Warning", "Please select a valid VISA resource.")
            return

        if self.inst:
            try:
                self.inst.close()
                print("🔌 Closed existing connection.")
            except Exception as e:
                print(f"Error closing existing connection: {e}")

        try:
            self.inst = self.rm.open_resource(selected_resource)
            self.inst.timeout = 5000 # 5 seconds timeout
            self.inst.read_termination = '\n' # N9340B typically terminates with newline
            self.inst.write_termination = '\n'
            
            # Query instrument ID
            instrument_id = query_safe(self.inst, "*IDN?")
            if instrument_id:
                print(f"✅ Connected to: {instrument_id.strip()}")
                model_match = re.search(r'N9340B|N9342C|N9343C|N9344C', instrument_id)
                instrument_model = model_match.group(0) if model_match else "Unknown Model"
                messagebox.showinfo("Connection Successful", f"Connected to: {instrument_id.strip()}")
                
                # Initialize instrument settings using the values from the GUI widgets
                ref_level = float(self.desired_ref_level_var.get())
                preamp_on = self.desired_preamp_var.get()
                display_log = self.desired_log_scale_var.get()
                rbw_config = int(float(self.desired_rbw_var.get()))
                vbw_config = int(rbw_config / 3)

                if initialize_instrument(self.inst, ref_level, preamp_on, display_log, rbw_config, vbw_config, instrument_model):
                    self.start_scan_button.config(state=tk.NORMAL)
                    self.stop_scan_button.config(state=tk.DISABLED)
                    self.query_instrument_settings() # Query and display current settings after initialization
                else:
                    messagebox.showerror("Initialization Failed", "Instrument initialization failed.")
                    self.inst.close()
                    self.inst = None
                    self.start_scan_button.config(state=tk.DISABLED)

            else:
                messagebox.showerror("Connection Failed", "Could not query instrument ID. Check connection or address.")
                if self.inst:
                    self.inst.close()
                self.inst = None
                self.start_scan_button.config(state=tk.DISABLED)
        except pyvisa.errors.VisaIOError as e:
            messagebox.showerror("VISA Error", f"Failed to connect to {selected_resource}: {e}")
            self.inst = None
            self.start_scan_button.config(state=tk.DISABLED)
        except ValueError:
            messagebox.showerror("Input Error", "Invalid numeric value for instrument settings. Please check Reference Level, RBW, and Max Hold Time.")
            self.start_scan_button.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
            self.inst = None
            self.start_scan_button.config(state=tk.DISABLED)

    def query_instrument_settings(self):
        """Queries the instrument for its current settings and updates the GUI."""
        if not self.inst:
            print("Not connected to instrument, cannot query settings.")
            return

        print("\nQuerying current instrument settings...")
        try:
            # Reference Level
            ref_level_str = query_safe(self.inst, ":DISPlay:WINDow:TRACe:Y:RLEVel?")
            if ref_level_str is not None:
                self.current_ref_level_var.set(f"{float(ref_level_str):.2f} dBm")
            else:
                self.current_ref_level_var.set("N/A")

            # Preamplifier State
            preamp_state_str = query_safe(self.inst, ":SENSe:POWer:RF:GAIN:STATe?")
            if preamp_state_str is not None:
                self.current_preamp_var.set(preamp_state_str.strip() == "1")
            else:
                self.current_preamp_var.set(False)

            # Display Scale
            display_scale_str = query_safe(self.inst, ":DISPlay:WINDow:TRACe:Y:SCALe:SPACing?")
            if display_scale_str is not None:
                self.current_log_scale_var.set(display_scale_str.strip().upper() == "LOGARITHMIC")
            else:
                self.current_log_scale_var.set(False)

            # RBW
            rbw_str = query_safe(self.inst, ":SENSe:BANDwidth:RESolution?")
            if rbw_str is not None:
                self.current_rbw_var.set(f"{float(rbw_str):.0f} Hz")
            else:
                self.current_rbw_var.set("N/A")

            # VBW
            vbw_str = query_safe(self.inst, ":SENSe:BANDwidth:VIDeo?")
            if vbw_str is not None:
                self.current_vbw_var.set(f"{float(vbw_str):.0f} Hz")
            else:
                self.current_vbw_var.set("N/A")

            # Sweep Time Auto
            sweep_time_auto_str = query_safe(self.inst, ":SENSe:SWEep:TIME:AUTO?")
            if sweep_time_auto_str is not None:
                self.current_sweep_time_auto_var.set(sweep_time_auto_str.strip() == "1")
            else:
                self.current_sweep_time_auto_var.set(False)

            # Start Frequency
            start_freq_str = query_safe(self.inst, ":SENSe:FREQuency:STARt?")
            if start_freq_str is not None:
                self.current_start_freq_var.set(f"{float(start_freq_str):.0f} Hz")
            else:
                self.current_start_freq_var.set("N/A")

            # Stop Frequency
            stop_freq_str = query_safe(self.inst, ":SENSe:FREQuency:STOP?")
            if stop_freq_str is not None:
                self.current_stop_freq_var.set(f"{float(stop_freq_str):.0f} Hz")
            else:
                self.current_stop_freq_var.set("N/A")

            # Sweep Points (N9340B fixed at 401)
            sweep_points_str = query_safe(self.inst, ":SENSe:SWEep:POINTS?")
            if sweep_points_str is not None:
                self.current_sweep_points_var.set(f"{int(float(sweep_points_str))}")
            else:
                self.current_sweep_points_var.set("N/A")

            print("✅ Current instrument settings updated in GUI.")

        except Exception as e:
            print(f"🛑 Error querying instrument settings: {e}")
            messagebox.showerror("Query Error", f"Failed to query instrument settings: {e}")
            # Reset all current settings display to N/A or default false
            self.current_ref_level_var.set("N/A")
            self.current_preamp_var.set(False)
            self.current_log_scale_var.set(False)
            self.current_rbw_var.set("N/A")
            self.current_vbw_var.set("N/A")
            self.current_sweep_time_auto_var.set(False)
            self.current_start_freq_var.set("N/A")
            self.current_stop_freq_var.set("N/A")
            self.current_sweep_points_var.set("N/A")


    def start_scan_thread(self):
        if not self.inst:
            messagebox.showwarning("Scan Warning", "Not connected to an instrument. Please connect first.")
            return

        # Get selected bands
        bands_to_scan = [band_info["band"] for band_info in self.band_vars if band_info["var"].get()]
        if not bands_to_scan:
            messagebox.showwarning("Scan Warning", "No frequency bands selected. Please select at least one band to scan.")
            return

        try:
            # Get desired settings from user input fields
            ref_level = float(self.desired_ref_level_var.get())
            preamp_on = self.desired_preamp_var.get()
            display_log = self.desired_log_scale_var.get()
            rbw_val = float(self.desired_rbw_var.get())
            max_hold_time = float(self.desired_max_hold_time_var.get())
            cycle_wait_time = float(self.desired_cycle_wait_time_var.get())

            if rbw_val <= 0 or max_hold_time < 0 or cycle_wait_time < 0:
                raise ValueError("RBW must be positive, Max Hold Time and Cycle Wait Time must be non-negative.")
            vbw_val = rbw_val / 3.0 # VBW is RBW/3
        except ValueError as e:
            messagebox.showerror("Input Error", f"Invalid input for settings: {e}")
            return

        self.scanning = True
        self.start_scan_button.config(state=tk.DISABLED)
        self.stop_scan_button.config(state=tk.NORMAL)
        self.plot_button.pack_forget() # Hide plot button during scan
        self.update_progress_label("Starting scan...")

        # Run scan in a separate thread
        scan_thread = threading.Thread(target=self._run_scan_loop, args=(bands_to_scan, rbw_val, vbw_val, max_hold_time, cycle_wait_time))
        scan_thread.daemon = True # Allow the thread to exit with the main program
        scan_thread.start()

    def _run_scan_loop(self, bands_to_scan, rbw_val, vbw_val, max_hold_time, cycle_wait_time):
        """Inner function to run the continuous scan in a separate thread."""
        scan_cycle_count = 0
        while self.scanning: # Loop indefinitely until self.scanning is False
            scan_cycle_count += 1
            self.update_progress_label(f"Scan cycle #{scan_cycle_count} starting...")
            print(f"\n--- Starting Scan Cycle #{scan_cycle_count} ---")

            if self.inst is None:
                print("Attempting to reconnect to instrument...")
                try:
                    # Attempt to reconnect using the last selected resource
                    selected_resource = self.resource_var.get()
                    self.inst = self.rm.open_resource(selected_resource)
                    self.inst.timeout = 5000
                    self.inst.read_termination = '\n'
                    self.inst.write_termination = '\n'
                    instrument_id = query_safe(self.inst, "*IDN?")
                    if instrument_id:
                        print(f"✅ Reconnected to: {instrument_id.strip()}")
                        ref_level = float(self.desired_ref_level_var.get())
                        preamp_on = self.desired_preamp_var.get()
                        display_log = self.desired_log_scale_var.get()
                        rbw_config = int(rbw_val)
                        vbw_config = int(vbw_val)
                        # Re-initialize instrument with current settings
                        initialize_instrument(self.inst, ref_level, preamp_on, display_log, rbw_config, vbw_config, "N9340B") # Model is hardcoded for re-init
                        self.query_instrument_settings() # Query and display current settings after re-initialization
                    else:
                        raise ValueError("Could not query instrument ID after reconnect.")
                except Exception as e:
                    print(f"🛑 Reconnection failed: {e}. Will try again next cycle.")
                    self.inst = None
                    time.sleep(5) # Delay before next reconnection attempt
                    continue # Skip current scan cycle if reconnection fails

            try:
                scanned_data = scan_bands(self, bands_to_scan, rbw_val, vbw_val, max_hold_time)

                if scanned_data:
                    # Save data to CSV
                    output_folder = self.output_folder_var.get()
                    os.makedirs(output_folder, exist_ok=True)
                    csv_filename = f"spectrum_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                    self.last_csv_file_path = os.path.join(output_folder, csv_filename)
                    with open(self.last_csv_file_path, 'w', newline='') as f:
                        writer = csv.writer(f)
                        writer.writerow(['Frequency_Hz', 'Power_dBm'])
                        writer.writerows(scanned_data)
                    print(f"✅ Scan data saved to: {self.last_csv_file_path}")
                else:
                    print("🚫 No data collected in this scan cycle.")

            except pyvisa.errors.VisaIOError as e:
                print(f"🛑 VISA Error in scan cycle #{scan_cycle_count}: {e}")
                messagebox.showerror("VISA Error", f"Lost connection to instrument: {e}")
                self.inst = None # Mark as disconnected
                print("🔄 Will attempt to re-initialize and continue scan in the next cycle.")
                time.sleep(5) # Short delay before re-attempting connection
                continue # Immediately go to the next cycle to try and reconnect/resume

            except Exception as e:
                print(f"🛑 An unexpected error occurred during scan cycle #{scan_cycle_count}: {e}")
                print("😴 Proceeding to wait period.")

            # If the scan completed (either fully or after an error that didn't stop the loop)
            # Wait for the cycle_wait_time, allowing for stop interruption
            self._wait_with_interrupt(cycle_wait_time)
            
            # If the scan was stopped during the wait, the loop condition will catch it
            if not self.scanning:
                print("\nScan loop terminated.")
                break # Exit the while loop if scan was stopped

        self._reset_buttons_after_scan()
        self.update_progress_label("Scan stopped or completed all cycles.")
        print("--- Scan thread finished. ---")


    def _wait_with_interrupt(self, wait_time_seconds):
        """Waits for a specified time but allows interruption by `self.scanning` flag."""
        start_time = time.time()
        end_time = start_time + wait_time_seconds
        
        while time.time() < end_time and self.scanning:
            remaining_time = int(end_time - time.time())
            if remaining_time >= 0:
                sys.stdout.write(f"\rWaiting for next cycle... {remaining_time} seconds remaining.   ")
                sys.stdout.flush()
            time.sleep(0.5) # Check every 500ms
        print("\nWait complete.")

    def stop_scan(self):
        self.scanning = False
        messagebox.showinfo("Scan Control", "Scan stop requested. Finishing current operation...")
        self.start_scan_button.config(state=tk.NORMAL)
        self.stop_scan_button.config(state=tk.DISABLED)
        # The _run_scan_loop thread will naturally exit after its current operation/wait.

    def update_progress_label(self, message):
        self.progress_label.config(text=message)
        self.update_idletasks() # Ensure GUI updates

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
    # Ensure dependencies are installed before running the app
    if check_and_install_dependencies():
        app = App()
        app.mainloop()
    else:
        print("Critical dependencies missing. Please install them to run the application.")
        messagebox.showerror("Dependency Error", "Some required Python packages are missing. Please install them manually and try again.")
