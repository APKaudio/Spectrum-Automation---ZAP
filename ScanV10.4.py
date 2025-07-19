# Updated ScanV10.4.py
import tkinter as tk
from tkinter import messagebox, scrolledtext
import pyvisa
import time
import argparse
import struct # Not strictly needed for ASCII, but kept for consistency with original imports
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
DEFAULT_RBW_STEP_SIZE_HZ = 1000000 # 10 kHz RBW resolution desired per data point
DEFAULT_CYCLE_WAIT_TIME_SECONDS = 0 # 30 seconds wait between full scan cycles
DEFAULT_MAXHOLD_TIME_SECONDS = 3 # Default max hold time for the new argument

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

def initialize_instrument(inst, ref_level_dbm, preamp_on, display_log_scale, rbw_config_val, vbw_config_val):
    """Initializes the Keysight N9340B spectrum analyzer with basic settings."""
    print("✨ Initializing instrument with desired settings...")
    try:
        # Reset is now only called when connecting, or if explicitly desired
        # It's assumed the instrument is in a controlled state via connection/disconnection.

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
        

        # Set display scale
        if display_log_scale:
            write_safe(inst, ":DISPlay:WINDow:TRACe:Y:SCALe:SPACing LOGarithmic")
            print("✅ Display scale set to LOG.")
        else:
            write_safe(inst, ":DISPlay:WINDow:TRACe:Y:SCALe:SPACing LINear")
            print("✅ Display scale set to LIN.")
        

        # Set RBW and VBW
        write_safe(inst, ":SENSe:BANDwidth:RESolution:AUTO OFF")
        write_safe(inst, f":SENSe:BANDwidth:RESolution {rbw_config_val}")
        

        write_safe(inst, ":SENSe:BANDwidth:VIDeo:AUTO OFF")
        write_safe(inst, f":SENSe:BANDwidth:VIDeo {vbw_config_val}")
        
        print(f"✅ Set RBW to {rbw_config_val} Hz, VBW to {vbw_config_val} Hz.")
        
        # Set sweep time to auto
        write_safe(inst, ":SENSe:SWEep:TIME:AUTO ON")
        
        print("✅ Sweep time set to AUTO.")

        # Configure trace 1 to clear write
        write_safe(inst, ":TRACe1:MODE WRITe") # Corrected SCPI command for trace mode
        print("✅ Trace mode set to CLEAR/WRITE.")
        
        # Explicitly set data format to ASCII for :TRACe:DATA? query
        write_safe(inst, ":FORMat:DATA ASCii") 
        print("✅ Set trace data format to ASCII for data transfer.")

        # Configure Markers (N9340B typically supports 6 markers)
        for i in range(1, 7): # Enable and configure markers 1 to 6
            write_safe(inst, f":CALCulate:MARKer{i}:STATe ON")
        
            write_safe(inst, f":CALCulate:MARKer{i}:MODE POSition") # Set to normal position mode
        
        print("✅ Markers 1-6 enabled and configured to position mode.")

        print("🎉 Instrument initialized successfully with desired settings.")
        return True
    except pyvisa.errors.VisaIOError as e:
        print(f"🛑 Failed to initialize instrument with desired settings: {e}")
        return False

def scan_bands(app_instance, csv_writer, bands_to_scan, rbw, vbw, max_hold_time):
    """
    Scans selected frequency bands using the :TRACe:DATA? TRACE1 command.
    Collects frequency and power data and writes to CSV.
    The number of sweep points is determined by the length of the returned trace data.
    """
    app_instance.all_scanned_data = [] # Clear previous scan data
    app_instance.last_scan_data = []
    
    # The number of sweep points will now be determined dynamically from the instrument's response
    # No longer using a fixed_sweep_points constant here.

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
                    current_sweep_data_str = query_safe(app_instance.inst, ":TRACe:DATA? TRACe1") # Query trace 1 data
                    if current_sweep_data_str is None:
                        print("Error: Could not retrieve trace data for max hold. Skipping.")
                        break

                    try:
                        # Ensure string is not empty before splitting
                        if current_sweep_data_str:
                            sweep_amplitudes_array = np.array([float(x) for x in current_sweep_data_str.split(',')])
                        else:
                            sweep_amplitudes_array = np.array([]) # Empty array if no data
                    except ValueError as e:
                        print(f"Error parsing trace data during max hold: {e}. Raw data (first 200 chars): '{current_sweep_data_str[:200]}'")
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
                trace_data_str = query_safe(app_instance.inst, "TRACe:DATA? TRACe1") # Query trace 1 data
                if trace_data_str is None:
                    print("Error: Could not retrieve trace data. Skipping band.")
                    continue
                try:
                    # Ensure string is not empty before splitting
                    if trace_data_str:
                        current_trace_amplitudes = np.array([float(x) for x in trace_data_str.split(',')])
                    else:
                        current_trace_amplitudes = np.array([]) # Empty array if no data
                except ValueError as e:
                    print(f"Error parsing trace data: {e}. Raw data (first 200 chars): '{trace_data_str[:200]}'")
                    continue

            # Determine actual sweep points from the returned data array
            actual_sweep_points_returned = len(current_trace_amplitudes)

            # Generate frequency points corresponding to the trace using actual_sweep_points_returned
            if current_trace_amplitudes is not None and actual_sweep_points_returned > 0:
                freq_points = np.linspace(start_freq_hz, stop_freq_hz, actual_sweep_points_returned)
                
                # Append data to the global list and write to CSV
                for j in range(actual_sweep_points_returned):
                    # Append for plotting later
                    app_instance.all_scanned_data.append((freq_points[j], current_trace_amplitudes[j]))
                    
                    # Write directly to CSV file
                    csv_writer.writerow([
                        f"{freq_points[j] / MHZ_TO_HZ:.2f}",  # Frequency in MHz
                        f"{current_trace_amplitudes[j]:.2f}"  # Level in dBm
                    ])
                print(f"✅ Collected {actual_sweep_points_returned} points for {band_name}.")
            else:
                print(f"❌ No data or invalid data received for {band_name}.")
            

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
        self.current_csv_file = None # To hold the open CSV file object

        # Dictionary to hold Entry widgets for coloring
        self.desired_setting_entries = {}

        # Initialize Tkinter variables for Device Configuration (Queried Values)
        self.current_ref_level_var = tk.StringVar(self, value="N/A")
        self.current_preamp_var = tk.BooleanVar(self, value=False)
        self.current_log_scale_var = tk.BooleanVar(self, value=False)
        self.current_rbw_var = tk.StringVar(self, value="N/A")
        self.current_vbw_var = tk.StringVar(self, value="N/A")
        self.current_sweep_time_auto_var = tk.BooleanVar(self, value=False)
        self.current_start_freq_var = tk.StringVar(self, value="N/A")
        self.current_stop_freq_var = tk.StringVar(self, value="N/A")
        # Removed self.current_sweep_points_var as it will be determined dynamically
        
        # Initialize Tkinter variables for Scan Configuration (User Input, can be pushed)
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

        # Using lambda for button commands to ensure correct binding
        self.start_scan_button = tk.Button(self.console_control_frame, text="Start Scan", command=lambda: self.start_scan_thread(), state=tk.DISABLED, bg="green", fg="white", height=2)
        self.start_scan_button.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill=tk.X)

        self.stop_scan_button = tk.Button(self.console_control_frame, text="Stop Scan", command=lambda: self.stop_scan(), state=tk.DISABLED, bg="red", fg="white", height=2)
        self.stop_scan_button.pack(side=tk.RIGHT, padx=5, pady=5, expand=True, fill=tk.X)

        self.console_output = scrolledtext.ScrolledText(self.console_frame, wrap=tk.WORD, bg="black", fg="white", font=("Consolas", 10))
        self.console_output.pack(expand=True, fill=tk.BOTH)
        self.console_output.configure(state="disabled")

        sys.stdout = TextRedirector(self.console_output, "stdout")
        sys.stderr = TextRedirector(self.console_output, "stderr")

        print("--- RF Spectrum Scanner GUI Initialized ---")
        self.create_widgets()
        # Defer populate_resources call to prevent AttributeError during __init__
        self.after(0, self.populate_resources)
        # self.actual_sweep_points = 401 # No longer a fixed constant, determined dynamically

    def create_widgets(self):
        # Resource selection
        resource_frame = tk.LabelFrame(self.main_frame, text="Instrument Connection", padx=10, pady=10)
        resource_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(resource_frame, text="VISA Resource:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.resource_dropdown = tk.OptionMenu(resource_frame, self.resource_var, "No Resources Found")
        self.resource_dropdown.grid(row=0, column=1, sticky=tk.EW, pady=2)

        # Using lambda for button commands to ensure correct binding
        self.refresh_button = tk.Button(resource_frame, text="Refresh", command=lambda: self.populate_resources())
        self.refresh_button.grid(row=0, column=2, padx=5, pady=2)

        self.connect_button = tk.Button(resource_frame, text="Connect", command=lambda: self.connect_instrument())
        self.connect_button.grid(row=0, column=3, padx=5, pady=2)
        
        self.disconnect_button = tk.Button(resource_frame, text="Disconnect", command=lambda: self.disconnect_instrument(), state=tk.DISABLED)
        self.disconnect_button.grid(row=0, column=4, padx=5, pady=2)

        # Device Configuration (Queried Values)
        current_settings_frame = tk.LabelFrame(self.main_frame, text="Current Device Configuration (Read from Device)", padx=10, pady=10)
        current_settings_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(current_settings_frame, text="Reference Level (dBm):").grid(row=0, column=0, sticky=tk.W, pady=2)
        tk.Label(current_settings_frame, textvariable=self.current_ref_level_var).grid(row=0, column=1, sticky=tk.W, pady=2)

        tk.Label(current_settings_frame, text="Preamplifier (ON/OFF):").grid(row=1, column=0, sticky=tk.W, pady=2)
        tk.Checkbutton(current_settings_frame, variable=self.current_preamp_var, state=tk.DISABLED).grid(row=1, column=1, sticky=tk.W, pady=2)

        tk.Label(current_settings_frame, text="Display Scale (LOG/LIN):").grid(row=2, column=0, sticky=tk.W, pady=2)
        tk.Checkbutton(current_settings_frame, variable=self.current_log_scale_var, state=tk.DISABLED).grid(row=2, column=1, sticky=tk.W, pady=2)

        tk.Label(current_settings_frame, text="RBW (Hz):").grid(row=3, column=0, sticky=tk.W, pady=2)
        tk.Label(current_settings_frame, textvariable=self.current_rbw_var).grid(row=3, column=1, sticky=tk.W, pady=2)
        
        tk.Label(current_settings_frame, text="VBW (Hz):").grid(row=4, column=0, sticky=tk.W, pady=2)
        tk.Label(current_settings_frame, textvariable=self.current_vbw_var).grid(row=4, column=1, sticky=tk.W, pady=2)

        tk.Label(current_settings_frame, text="Sweep Time Auto (ON/OFF):").grid(row=5, column=0, sticky=tk.W, pady=2)
        tk.Checkbutton(current_settings_frame, variable=self.current_sweep_time_auto_var, state=tk.DISABLED).grid(row=5, column=1, sticky=tk.W, pady=2)

        tk.Label(current_settings_frame, text="Start Freq (Hz):").grid(row=6, column=0, sticky=tk.W, pady=2)
        tk.Label(current_settings_frame, textvariable=self.current_start_freq_var).grid(row=6, column=1, sticky=tk.W, pady=2)

        tk.Label(current_settings_frame, text="Stop Freq (Hz):").grid(row=7, column=0, sticky=tk.W, pady=2)
        tk.Label(current_settings_frame, textvariable=self.current_stop_freq_var).grid(row=7, column=1, sticky=tk.W, pady=2)


        # Scan Configuration (User Input)
        scan_settings_frame = tk.LabelFrame(self.main_frame, text="Scan Configuration (Push to Device)", padx=10, pady=10)
        scan_settings_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(scan_settings_frame, text="Reference Level (dBm):").grid(row=0, column=0, sticky=tk.W, pady=2)
        entry_ref_level = tk.Entry(scan_settings_frame, textvariable=self.desired_ref_level_var)
        entry_ref_level.grid(row=0, column=1, sticky=tk.EW, pady=2)
        self.desired_setting_entries["ref_level"] = entry_ref_level

        tk.Label(scan_settings_frame, text="Preamplifier (ON/OFF):").grid(row=1, column=0, sticky=tk.W, pady=2)
        check_preamp = tk.Checkbutton(scan_settings_frame, variable=self.desired_preamp_var)
        check_preamp.grid(row=1, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["preamp"] = check_preamp # Associate for coloring

        tk.Label(scan_settings_frame, text="Display Scale (LOG/LIN):").grid(row=2, column=0, sticky=tk.W, pady=2)
        check_log_scale = tk.Checkbutton(scan_settings_frame, variable=self.desired_log_scale_var)
        check_log_scale.grid(row=2, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["log_scale"] = check_log_scale

        tk.Label(scan_settings_frame, text="RBW (Hz):").grid(row=3, column=0, sticky=tk.W, pady=2)
        entry_rbw = tk.Entry(scan_settings_frame, textvariable=self.desired_rbw_var)
        entry_rbw.grid(row=3, column=1, sticky=tk.EW, pady=2)
        self.desired_setting_entries["rbw"] = entry_rbw

        tk.Label(scan_settings_frame, text="VBW (Hz):").grid(row=4, column=0, sticky=tk.W, pady=2)
        entry_vbw = tk.Entry(scan_settings_frame, textvariable=self.desired_vbw_display_var)
        entry_vbw.grid(row=4, column=1, sticky=tk.EW, pady=2)
        self.desired_setting_entries["vbw"] = entry_vbw

        tk.Label(scan_settings_frame, text="Max Hold Enabled:").grid(row=5, column=0, sticky=tk.W, pady=2)
        check_max_hold = tk.Checkbutton(scan_settings_frame, variable=self.desired_max_hold_var)
        check_max_hold.grid(row=5, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["max_hold_enabled"] = check_max_hold
        
        tk.Label(scan_settings_frame, text="Max Hold Time (s):").grid(row=6, column=0, sticky=tk.W, pady=2)
        entry_max_hold_time = tk.Entry(scan_settings_frame, textvariable=self.desired_max_hold_time_var)
        entry_max_hold_time.grid(row=6, column=1, sticky=tk.EW, pady=2)
        self.desired_setting_entries["max_hold_time"] = entry_max_hold_time

        tk.Label(scan_settings_frame, text="Cycle Wait Time (s):").grid(row=7, column=0, sticky=tk.W, pady=2)
        entry_cycle_wait = tk.Entry(scan_settings_frame, textvariable=self.desired_cycle_wait_time_var)
        entry_cycle_wait.grid(row=7, column=1, sticky=tk.EW, pady=2)
        self.desired_setting_entries["cycle_wait_time"] = entry_cycle_wait
        
        # Output Folder with Open Folder Button
        output_folder_frame = tk.Frame(scan_settings_frame)
        output_folder_frame.grid(row=8, column=0, columnspan=2, sticky=tk.EW, pady=2)
        tk.Label(output_folder_frame, text="Output Folder:").pack(side=tk.LEFT, padx=(0, 5))
        entry_output_folder = tk.Entry(output_folder_frame, textvariable=self.output_folder_var)
        entry_output_folder.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.desired_setting_entries["output_folder"] = entry_output_folder
        open_folder_button = tk.Button(output_folder_frame, text="Open Folder", command=lambda: self.open_output_folder())
        open_folder_button.pack(side=tk.RIGHT, padx=(5, 0))


        self.apply_button = tk.Button(scan_settings_frame, text="Apply Settings to Device", command=lambda: self.apply_settings_to_device(), state=tk.DISABLED)
        self.apply_button.grid(row=9, column=0, columnspan=2, pady=10, sticky=tk.EW)

        # Band Selection - Moved to appear after scan settings
        band_selection_frame = tk.LabelFrame(self.main_frame, text="Frequency Band Selection", padx=10, pady=10)
        # Crucial change: added expand=True here to ensure it takes vertical space
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

        # Plot button (initially hidden) - This button is in the console_control_frame, not main_frame
        self.plot_button = tk.Button(self.console_control_frame, text="Generate Plot", command=lambda: self.generate_plot(), state=tk.DISABLED, bg="blue", fg="white", height=2)
        self.plot_button.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill=tk.X) # Initially packed, then hidden/shown
        self.plot_button.pack_forget() # Ensure it's hidden by default

        # Configure columns to expand
        for i in range(5): # Adjust if more columns are added
            resource_frame.grid_columnconfigure(i, weight=1)
        for i in range(2): # For settings frames
            current_settings_frame.grid_columnconfigure(i, weight=1)
            scan_settings_frame.grid_columnconfigure(i, weight=1)

        # Update VBW display initially
        self.update_vbw_display()

    def update_vbw_display(self):
        """Updates the VBW display based on the current RBW setting."""
        try:
            rbw_val = float(self.desired_rbw_var.get())
            self.desired_vbw_display_var.set(str(int(rbw_val / 3)))
        except ValueError:
            self.desired_vbw_display_var.set("Invalid RBW")


    def open_output_folder(self):
        """Opens the specified output folder in the file explorer."""
        folder_path = self.output_folder_var.get()
        if not os.path.exists(folder_path):
            messagebox.showwarning("Folder Not Found", f"The folder '{folder_path}' does not exist.")
            print(f"🚫 Folder not found: {folder_path}")
            return
        
        try:
            if sys.platform == "win32":
                os.startfile(folder_path)
            elif sys.platform == "darwin": # macOS
                subprocess.run(['open', folder_path])
            else: # linux variants
                subprocess.run(['xdg-open', folder_path])
            print(f"✅ Opened folder: {folder_path}")
        except Exception as e:
            messagebox.showerror("Error Opening Folder", f"Failed to open folder '{folder_path}': {e}")
            print(f"❌ Error opening folder: {e}")

    def connect_instrument(self):
        """Establishes connection to the selected instrument and queries its settings."""
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
                
                # Immediately query current settings and display them without pushing any config
                self.query_instrument_settings()
                
                # Now, push the desired settings from the GUI to the instrument (initial configuration)
                ref_level = float(self.desired_ref_level_var.get())
                preamp_on = self.desired_preamp_var.get()
                display_log = self.desired_log_scale_var.get()
                rbw_config = int(float(self.desired_rbw_var.get()))
                vbw_config = int(rbw_config / 3)

                # Reset the instrument to a known state using *RST first during connection
                write_safe(self.inst, "*RST")
                query_safe(self.inst, "*OPC?")
                time.sleep(1) # Give it a moment after reset

                if initialize_instrument(self.inst, ref_level, preamp_on, display_log, rbw_config, vbw_config):
                    self.start_scan_button.config(state=tk.NORMAL)
                    self.stop_scan_button.config(state=tk.DISABLED)
                    self.disconnect_button.config(state=tk.NORMAL)
                    self.reset_setting_colors() # Settings pushed, so revert colors to black
                    self.query_instrument_settings() # Re-query to confirm pushed settings
                else:
                    messagebox.showerror("Initialization Failed", "Instrument initialization with desired settings failed.")
                    self.inst.close()
                    self.inst = None
                    self.start_scan_button.config(state=tk.DISABLED)
                    self.disconnect_button.config(state=tk.DISABLED)

            else:
                messagebox.showerror("Connection Failed", "Could not query instrument ID. Check connection or address.")
                if self.inst:
                    self.inst.close()
                self.inst = None
                self.start_scan_button.config(state=tk.DISABLED)
                self.disconnect_button.config(state=tk.DISABLED)
        except pyvisa.errors.VisaIOError as e:
            messagebox.showerror("VISA Error", f"Failed to connect to {selected_resource}: {e}")
            self.inst = None
            self.start_scan_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.DISABLED)
        except ValueError as e:
            messagebox.showerror("Input Error", f"Invalid numeric value for instrument settings: {e}. Please check Reference Level, RBW, and Max Hold Time.")
            self.start_scan_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
            self.inst = None
            self.start_scan_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.DISABLED)

    def disconnect_instrument(self):
        """Closes the connection to the instrument."""
        if self.inst:
            try:
                self.inst.close()
                self.inst = None
                print("🔌 Instrument disconnected.")
                messagebox.showinfo("Disconnected", "Instrument disconnected successfully.")
                self.start_scan_button.config(state=tk.DISABLED)
                self.stop_scan_button.config(state=tk.DISABLED)
                self.disconnect_button.config(state=tk.DISABLED)
                self.update_progress_label("Disconnected.")
                # Clear current device configuration display
                self.current_ref_level_var.set("N/A")
                self.current_preamp_var.set(False)
                self.current_log_scale_var.set(False)
                self.current_rbw_var.set("N/A")
                self.current_vbw_var.set("N/A")
                self.current_sweep_time_auto_var.set(False)
                self.current_start_freq_var.set("N/A")
                self.current_stop_freq_var.set("N/A")
                self.current_sweep_points_var.set("N/A")
                self.plot_button.pack_forget() # Hide plot button on disconnect
            except Exception as e:
                messagebox.showerror("Disconnect Error", f"Error disconnecting instrument: {e}")
                print(f"Error disconnecting: {e}")
        else:
            messagebox.showwarning("Disconnect Warning", "No instrument is currently connected.")
    def reset_setting_colors(self):
        """Resets the text color of all desired setting entries to black."""
        for key, entry_widget in self.desired_setting_entries.items():
            if isinstance(entry_widget, tk.Entry):
                entry_widget.config(fg="black")
            # For Checkbuttons, there's no direct foreground change for the text
            # unless a custom widget is implemented.

    def query_instrument_settings(self):
        """Queries the instrument for its current settings and updates the GUI."""
        if not self.inst:
            print("Not connected to instrument, cannot query settings.")
            return

        print("\nQuerying current instrument settings...")
        try:
            # Reference Lvel
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
        """Starts the scanning process in a separate thread."""
        if not self.inst:
            messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
            return
        
        if self.scanning:
            messagebox.showwarning("Scan in Progress", "A scan is already running.")
            return

        # Disable buttons during scan
        self.start_scan_button.config(state=tk.DISABLED)
        self.stop_scan_button.config(state=tk.NORMAL)
        self.connect_button.config(state=tk.DISABLED)
        self.disconnect_button.config(state=tk.DISABLED)
        self.apply_button.config(state=tk.DISABLED)
        self.plot_button.pack_forget() # Hide plot button during scan

        self.scanning = True
        print("\nStarting spectrum scan...")
        
        # Get scan configuration from GUI
        max_hold_enabled = self.desired_max_hold_var.get()
        max_hold_time = float(self.desired_max_hold_time_var.get()) if max_hold_enabled else 0
        
        # Ensure output directory exists
        output_folder = self.output_folder_var.get()
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            print(f"Created output directory: {output_folder}")

        # Prepare CSV file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_filename = f"scan_data_{timestamp}.csv"
        self.last_csv_file_path = os.path.join(output_folder, csv_filename)

        try:
            # Open CSV file and store the file object for the thread
            self.current_csv_file = open(self.last_csv_file_path, 'w', newline='')
            csv_writer = csv.writer(self.current_csv_file)
            # Write CSV header
            csv_writer.writerow(["Frequency (MHz)", "Level (dBm)"])
            print(f"CSV file created: {self.last_csv_file_path}")

            # Get selected bands
            selected_bands = [item["band"] for item in self.band_vars if item["var"].get()]
            if not selected_bands:
                messagebox.showwarning("No Bands Selected", "Please select at least one frequency band to scan.")
                print("🚫 No bands selected for scan.")
                self.stop_scan() # Reset GUI state
                return

            scan_thread = threading.Thread(target=self._run_scan, args=(csv_writer, max_hold_time, selected_bands))
            scan_thread.daemon = True # Allow program to exit even if thread is running
            scan_thread.start()
        except Exception as e:
            messagebox.showerror("File Error", f"Could not create CSV file: {e}")
            print(f"❌ Error creating CSV file: {e}")
            self.stop_scan() # Reset GUI state

    def _run_scan(self, csv_writer, max_hold_time, selected_bands):
        """Internal method to run the scan logic, called by the thread."""
        try:
            # Re-read desired settings for the scan, in case they were changed after apply
            rbw_val = float(self.desired_rbw_var.get())
            vbw_val = float(self.desired_vbw_display_var.get()) # Use the displayed VBW value

            # Passed csv_writer to scan_bands
            # Pass selected_bands to scan_bands
            scanned_data = scan_bands(self, csv_writer, selected_bands, rbw_val, vbw_val, max_hold_time) 
            self.last_scan_data = scanned_data
            
            if not self.scanning: # If scan was stopped by user
                print("\nScan process finished (interrupted).")
            else:
                print("\nScan process finished.")
            
        except Exception as e:
            messagebox.showerror("Scan Error", f"An error occurred during scanning: {e}")
            print(f"❌ Scan thread encountered an error: {e}")
        finally:
            self.scanning = False
            if self.current_csv_file:
                self.current_csv_file.close() # Ensure the CSV file is closed
                print("CSV file closed.")
            self.after(100, self.reset_scan_buttons) # Use after to update GUI from main thread


    def populate_resources(self):
        """Populates the VISA resource dropdown."""
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
            self.disconnect_button.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to list VISA resources: {e}")
            self.resource_var.set("Error listing resources")
            self.start_scan_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.DISABLED)


    def stop_scan(self):
        """Stops the ongoing scan."""
        self.scanning = False
        print("\nAttempting to stop scan... Please wait for current sweep to finish.")
        self.stop_scan_button.config(state=tk.DISABLED) # Disable stop button immediately

    def reset_scan_buttons(self):
        """Resets the state of scan-related buttons after a scan completes or stops."""
        self.start_scan_button.config(state=tk.NORMAL)
        if self.inst: # Only enable disconnect/apply if connected
            self.disconnect_button.config(state=tk.NORMAL)
            self.apply_button.config(state=tk.NORMAL)
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
