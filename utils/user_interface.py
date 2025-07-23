# user_interface.py

import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, TclError
from tkinter import ttk
import time
import os
import sys
import subprocess
import threading
import re
from datetime import datetime
import pandas as pd
import csv
import pyvisa
import configparser

# Import constants from frequency_bands.py
try:
    # Changed to absolute import, relying on main_app.py adding project root to sys.path
    from utils.frequency_bands import (
        MHZ_TO_HZ,
        SCAN_BAND_RANGES,
        TV_PLOT_BAND_MARKERS,
        GOV_PLOT_BAND_MARKERS
    )
except ImportError:
    print("Error: frequency_bands.py not found. Please ensure it's in the same directory or in the 'utils' folder.")
    MHZ_TO_HZ = 1_000_000
    SCAN_BAND_RANGES = [
        {"Band Name": "Dummy Band 1", "Start MHz": 100, "Stop MHz": 200},
        {"Band Name": "Dummy Band 2", "Start MHz": 400, "Stop MHz": 500}
    ]
    TV_PLOT_BAND_MARKERS = []
    GOV_PLOT_BAND_MARKERS = []

# Import scanning, plotting, CSV, and instrument control utilities
# Changed to absolute imports, relying on main_app.py adding project root to sys.path
from utils.scan_instrument import scan_bands
from utils.plotting_utils import plot_single_scan_data, _open_plot_in_browser
from utils.averaging_utils import generate_current_cycle_average_csv_and_plot, generate_historical_average_plot
from utils.csv_utils import write_scan_data_to_csv
from utils.instrument_control import (
    set_debug_mode,
    list_visa_resources,
    connect_to_instrument,
    disconnect_instrument as control_disconnect_instrument,
    initialize_instrument,
    query_current_instrument_settings,
    query_device_presets as control_query_device_presets,
    load_selected_preset as control_load_selected_preset
)

# Define the config file name, ensuring it's in the same directory as the script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, 'config.ini')

class TextRedirector(object):
    """A class to redirect stdout/stderr to a Tkinter scrolled text widget."""
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag
        self.last_char_was_cr = False

    def write(self, str_val):
        self.widget.config(state=tk.NORMAL)
        
        if '\r' in str_val:
            parts = str_val.split('\r')
            for i, part in enumerate(parts):
                if self.last_char_was_cr and i == 0:
                    self.widget.delete("end-1c linestart", "end-1c")
                self.widget.insert(tk.END, part, (self.tag,))
                self.widget.see(tk.END)
                if i < len(parts) - 1:
                    self.last_char_was_cr = True
                else:
                    self.last_char_was_cr = False
        else:
            self.widget.insert(tk.END, str_val, (self.tag,))
            self.widget.see(tk.END)
            self.last_char_was_cr = False

        self.widget.config(state=tk.DISABLED)
        self.widget.update_idletasks()

    def flush(self):
        pass

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.configure(bg="black") 
        self.title(f"RF Spectrum Analyzer Controller - {os.path.basename(sys.argv[0])}")
        
        self.rm = pyvisa.ResourceManager()
        self.instrument_list = []
        self.inst = None
        self.scanning = False
        self.paused = False
        self.last_scan_data = None # This will still store the in-memory data for potential other uses
        self.collected_scans_dataframes = []
        self.instrument_model = None

        self.scan_cycle_count = 0
        self.current_freq_offset = 0

        self.desired_setting_entries = {}

        # Removed: New variables for session-wide CSV
        # self.session_csv_file_path = None
        # self.session_csv_headers_written = False

        # Initialize Tkinter variables
        self.desired_ref_level_var = tk.StringVar(self)
        self.desired_preamp_var = tk.BooleanVar(self)
        self.high_sensitivity_var = tk.BooleanVar(self)
        self.desired_max_hold_var = tk.BooleanVar(self)
        self.desired_max_hold_time_var = tk.DoubleVar(self)
        self.desired_rbw_var = tk.StringVar(self) # This is for the RBW display, not scan_rbw_segmentation
        self.desired_vbw_display_var = tk.StringVar(self)
        self.desired_cycle_wait_time_var = tk.DoubleVar(self)
        self.output_folder_var = tk.StringVar(self)
        self.scan_name_var = tk.StringVar(self)
        self.resource_var = tk.StringVar(self)
        self.include_gov_markers_var = tk.BooleanVar(self)
        self.include_tv_markers_var = tk.BooleanVar(self)
        self.open_html_after_complete_var = tk.BooleanVar(self)
        self.desired_scan_rbw_segmentation_var = tk.DoubleVar(self)
        self.shift_freq_var = tk.DoubleVar(self)
        self.debug_mode_var = tk.BooleanVar(self)
        self.last_selected_bands_str = tk.StringVar(self) # New variable for selected bands

        # Initialize band_checkboxes and band_vars here
        # This ensures they exist before _load_config tries to set band states
        self.band_checkboxes = []
        self.band_vars = []

        # Map Tkinter variables to their corresponding config keys for last used settings
        # (tk_var_name_string, last_used_config_key, default_config_key, type_converter)
        self.setting_var_map = {
            'desired_ref_level_var': ('last_reference_level_dbm', 'default_reference_level_dbm', float),
            'desired_preamp_var': ('last_preamp_on', 'default_preamp_on', bool),
            'high_sensitivity_var': ('last_high_sensitivity', 'default_high_sensitivity', bool),
            'desired_max_hold_var': ('last_maxhold_enabled', 'default_maxhold_enabled', bool),
            'desired_max_hold_time_var': ('last_maxhold_time_seconds', 'default_maxhold_time_seconds', float),
            'desired_rbw_var': ('last_rbw_step_size_hz', 'default_rbw_step_size_hz', int), # This is for the RBW display
            'desired_cycle_wait_time_var': ('last_cycle_wait_time_seconds', 'default_cycle_wait_time_seconds', float),
            'output_folder_var': ('last_scan_directory', 'default_scan_directory', str),
            'scan_name_var': ('last_scan_name', 'default_scan_name', str),
            'resource_var': ('last_gpib_device', '', str), # No direct default for this, populated by populate_resources
            'include_gov_markers_var': ('last_include_gov_markers', 'default_include_gov_markers', bool),
            'include_tv_markers_var': ('last_include_tv_markers', 'default_include_tv_markers', bool),
            'open_html_after_complete_var': ('last_open_html_after_complete', 'default_open_html_after_complete', bool),
            'desired_scan_rbw_segmentation_var': ('last_scan_rbw_hz', 'default_scan_rbw_hz', float),
            'shift_freq_var': ('last_freq_shift_hz', 'default_freq_shift_hz', float),
            'debug_mode_var': ('last_debug_mode', 'default_debug_mode', bool),
            'last_selected_bands_str': ('last_selected_bands', 'default_selected_bands', str), # New entry for selected bands
        }

        # Load configuration at startup
        self.config = configparser.ConfigParser()
        self._load_config() # This now only loads/sets defaults in self.config, no saving yet

        
        
        
        # Set window geometry from config, or fallback to default
        initial_geometry = self.config.get('LAST_USED_SETTINGS', 'LAST_WINDOW_GEOMETRY')
        if not initial_geometry:
            initial_geometry = self.config.get('DEFAULT_SETTINGS', 'DEFAULT_WINDOW_GEOMETRY')
        self.geometry(initial_geometry)

        self.rbw_values = [5000, 10000, 25000, 50000, 100000]
        self.rbw_val_to_idx = {val: i for i, val in enumerate(self.rbw_values)}
        # Initialize slider index based on the loaded desired_scan_rbw_segmentation_var
        self.rbw_slider_index_var = tk.IntVar(self, value=self.rbw_val_to_idx.get(int(self.desired_scan_rbw_segmentation_var.get()), 0))
        self.rbw_slider_index_var.trace_add("write", self._update_scan_rbw_from_slider_index)

        self.freq_shift_values = [0, 500, 1000, 5000, 10000]
        self.freq_shift_val_to_idx = {val: i for i, val in enumerate(self.freq_shift_values)}
        # Initialize slider index based on the loaded shift_freq_var
        self.freq_shift_slider_index_var = tk.IntVar(self, value=self.freq_shift_val_to_idx.get(int(self.shift_freq_var.get()), 0))
        self.freq_shift_slider_index_var.trace_add("write", self._update_freq_shift_from_slider_index)

        self.debug_mode_var.trace_add("write", self._update_debug_mode_global)

        self.main_frame = tk.Frame(self, bg="black")
        self.main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.console_frame = tk.Frame(self, width=700, bg="black")
        self.console_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.console_frame.pack_propagate(False)

        self.console_control_frame = tk.Frame(self.console_frame, bg="black")
        self.console_control_frame.pack(side=tk.TOP, fill=tk.X, pady=5)

        self.start_scan_button = tk.Button(self.console_control_frame, text="Start Scan", command=lambda: self.start_scan_thread(), state=tk.DISABLED, bg="green", fg="white", height=2)
        self.start_scan_button.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill=tk.X)

        self.pause_resume_button = tk.Button(self.console_control_frame, text="Pause Scan", command=self.toggle_pause_scan, state=tk.DISABLED, bg="orange", fg="white", height=2)
        self.pause_resume_button.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill=tk.X)

        self.stop_scan_button = tk.Button(self.console_control_frame, text="Stop Scan", command=lambda: self.stop_scan(), state=tk.DISABLED, bg="red", fg="white", height=2)
        self.stop_scan_button.pack(side=tk.RIGHT, padx=5, pady=5, expand=True, fill=tk.X)

        self.console_output = scrolledtext.ScrolledText(self.console_frame, wrap=tk.WORD, bg="black", fg="white", font=("Consolas", 10))
        self.console_output.pack(expand=True, fill=tk.BOTH)
        self.console_output.configure(state="disabled")

        sys.stdout = TextRedirector(self.console_output, "stdout")
        sys.stderr = TextRedirector(self.console_output, "stderr")

        print_art()

        print("--- RF Spectrum Scanner GUI Initialized ---")
        self.create_widgets()
        self.after(0, self.populate_resources)

        

        # Bind the closing protocol to save config
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _load_config(self):
        """Loads configuration from config.ini, setting defaults if file or sections are missing."""
        self.config.read(CONFIG_FILE)

        # Default settings section
        if 'DEFAULT_SETTINGS' not in self.config:
            self.config['DEFAULT_SETTINGS'] = {}
        
        # Dynamically generate the default selected bands string from SCAN_BAND_RANGES
        default_selected_bands_str = ",".join([band["Band Name"] for band in SCAN_BAND_RANGES])

        # Ensure all default settings are present with their default values
        default_settings_map = {
            'DEFAULT_RBW_STEP_SIZE_HZ': '1000000',
            'DEFAULT_CYCLE_WAIT_TIME_SECONDS': '0',
            'DEFAULT_MAXHOLD_TIME_SECONDS': '3',
            'DEFAULT_SCAN_RBW_HZ': '10000',
            'DEFAULT_REFERENCE_LEVEL_DBM': '-40',
            'DEFAULT_FREQ_SHIFT_HZ': '0',
            'DEFAULT_MAXHOLD_ENABLED': 'True',
            'DEFAULT_INCLUDE_GOV_MARKERS': 'True',
            'DEFAULT_INCLUDE_TV_MARKERS': 'True',
            'DEFAULT_OPEN_HTML_AFTER_COMPLETE': 'True',
            'DEFAULT_HIGH_SENSITIVITY': 'True',
            'DEFAULT_PREAMP_ON': 'True',
            'DEFAULT_DEBUG_MODE': 'False',
            'DEFAULT_WINDOW_GEOMETRY': '1400x780+100+100',
            'DEFAULT_SCAN_DIRECTORY': 'scan_data', # Added default
            'DEFAULT_SCAN_NAME': 'MyScan', # Added default
            'DEFAULT_SELECTED_BANDS': default_selected_bands_str, # Dynamically set default selected bands
        }
        for key, default_val in default_settings_map.items():
            if key.lower() not in [k.lower() for k in self.config['DEFAULT_SETTINGS'].keys()]: # Case-insensitive check
                self.config['DEFAULT_SETTINGS'][key] = default_val
        
        # Last used settings section
        if 'LAST_USED_SETTINGS' not in self.config:
            self.config['LAST_USED_SETTINGS'] = {}

        # Load values into Tkinter variables, prioritizing LAST_USED_SETTINGS
        self._populate_vars_from_config()


    def _populate_vars_from_config(self):
        """Populates Tkinter variables from config, prioritizing LAST_USED_SETTINGS, then DEFAULT_SETTINGS."""
        for var_name, (last_key, default_key, type_converter) in self.setting_var_map.items():
            tk_var = getattr(self, var_name)
            value_to_set = None

            # Try to load from LAST_USED_SETTINGS first
            if self.config.has_option('LAST_USED_SETTINGS', last_key) and self.config.get('LAST_USED_SETTINGS', last_key):
                try:
                    if type_converter == bool:
                        value_to_set = self.config.getboolean('LAST_USED_SETTINGS', last_key)
                    elif type_converter == float:
                        value_to_set = self.config.getfloat('LAST_USED_SETTINGS', last_key)
                    elif type_converter == int:
                        value_to_set = self.config.getint('LAST_USED_SETTINGS', last_key)
                    else: # str
                        value_to_set = self.config.get('LAST_USED_SETTINGS', last_key)
                except ValueError as e:
                    print(f"Warning: Could not parse '{last_key}' from LAST_USED_SETTINGS (value: '{self.config.get('LAST_USED_SETTINGS', last_key)}'). Error: {e}")
                    value_to_set = None # Fallback to default
            
            # If not found in LAST_USED_SETTINGS or empty/invalid, try DEFAULT_SETTINGS
            if value_to_set is None and default_key and self.config.has_option('DEFAULT_SETTINGS', default_key):
                try:
                    if type_converter == bool:
                        value_to_set = self.config.getboolean('DEFAULT_SETTINGS', default_key)
                    elif type_converter == float:
                        value_to_set = self.config.getfloat('DEFAULT_SETTINGS', default_key)
                    elif type_converter == int:
                        value_to_set = self.config.getint('DEFAULT_SETTINGS', default_key)
                    else: # str
                        value_to_set = self.config.get('DEFAULT_SETTINGS', default_key)
                except ValueError as e:
                    print(f"Warning: Could not parse '{default_key}' from DEFAULT_SETTINGS (value: '{self.config.get('DEFAULT_SETTINGS', default_key)}'). Error: {e}")
                    value_to_set = None # Set to None if default is also invalid

            if value_to_set is not None:
                tk_var.set(value_to_set)
            else:
                # Fallback for resource_var if no last_gpib_device and no default_gpib_device
                if var_name == 'resource_var':
                    tk_var.set("No Resources Found")
                # For other variables, if no valid setting found, they will retain their default Tkinter variable value (e.g., 0 for int/float, False for bool, "" for string)

        # _set_band_checkboxes_from_config() is now called AFTER create_widgets()
        # This ensures self.band_vars is populated.


    def _save_config(self):
        """Saves current settings to config.ini into LAST_USED_SETTINGS."""
        # Ensure LAST_USED_SETTINGS section exists
        if 'LAST_USED_SETTINGS' not in self.config:
            self.config['LAST_USED_SETTINGS'] = {}

        for var_name, (last_key, _, _) in self.setting_var_map.items():
            if last_key and var_name != 'last_selected_bands_str': # Exclude last_selected_bands_str from generic save
                tk_var = getattr(self, var_name)
                self.config['LAST_USED_SETTINGS'][last_key] = str(tk_var.get())
        
        # Save selected bands separately
        selected_band_names = [item["band"]["Band Name"] for item in self.band_vars if item["var"].get()]
        self.config['LAST_USED_SETTINGS']['last_selected_bands'] = ",".join(selected_band_names)
        
        # Save current window geometry
        self.config['LAST_USED_SETTINGS']['last_window_geometry'] = self.winfo_geometry()

        with open(CONFIG_FILE, 'w') as configfile:
            self.config.write(configfile)
        print(f"Configuration saved to {CONFIG_FILE}")

    def on_closing(self):
        """Handler for window closing event to save config."""
        self._save_config()
        self.destroy()

    def _update_debug_mode_global(self, *args):
        """Updates the global debug mode variable when the checkbox state changes."""
        set_debug_mode(self.debug_mode_var.get())
        # Also update the config setting immediately
        # This specific debug mode update is handled by the generic _save_config now.
        # self.config['DEFAULT_SETTINGS']['DEFAULT_DEBUG_MODE'] = str(self.debug_mode_var.get())
        self._save_config()


    def _update_scan_rbw_from_slider_index(self, *args):
        """Updates scan RBW from slider index."""
        try:
            idx = self.rbw_slider_index_var.get()
            if 0 <= idx < len(self.rbw_values):
                self.desired_scan_rbw_segmentation_var.set(float(self.rbw_values[idx]))
        except Exception as e:
            print(f"Error updating scan RBW from slider index: {e}")

    def _update_freq_shift_from_slider_index(self, *args):
        """Updates frequency shift from slider index."""
        try:
            idx = self.freq_shift_slider_index_var.get()
            if 0 <= idx < len(self.freq_shift_values):
                self.shift_freq_var.set(float(self.freq_shift_values[idx]))
        except Exception as e:
            print(f"Error updating frequency shift from slider index: {e}")

    def create_widgets(self):
        style = ttk.Style(self)
        style.theme_use('clam')

        style.configure("Treeview", 
                        background="grey", 
                        foreground="white", 
                        fieldbackground="grey",
                        bordercolor="black",
                        lightcolor="grey",
                        darkcolor="grey")
        style.map("Treeview", 
                  background=[("selected", "blue")], 
                  foreground=[("selected", "white")])
        
        style.configure("Treeview.Heading", 
                        background="darkgrey", 
                        foreground="white",
                        font=('TkDefaultFont', 10, 'bold'))

        style.configure("Vertical.TScrollbar", 
                        background="darkgrey", 
                        troughcolor="black", 
                        bordercolor="black",
                        arrowcolor="white")
        style.map("Vertical.TScrollbar",
                  background=[('active', 'gray')])

        resource_frame = tk.LabelFrame(self.main_frame, text="Instrument Connection", padx=10, pady=10, bg="black", fg="white")
        resource_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(resource_frame, text="VISA Resource:", bg="black", fg="white").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.resource_dropdown = tk.OptionMenu(resource_frame, self.resource_var, "No Resources Found")
        self.resource_dropdown.config(bg="grey", fg="white", highlightbackground="grey", highlightcolor="grey", activebackground="darkgrey", activeforeground="white")
        self.resource_dropdown.grid(row=0, column=1, sticky=tk.EW, pady=2)
        self.resource_dropdown["menu"].config(bg="grey", fg="white", activebackground="darkgrey", activeforeground="white")

        self.refresh_button = tk.Button(resource_frame, text="Refresh", command=lambda: self.populate_resources(), bg="darkgrey", fg="white")
        self.refresh_button.grid(row=0, column=2, padx=5, pady=2)

        self.connect_button = tk.Button(resource_frame, text="Connect", command=lambda: self.connect_instrument(), bg="darkgrey", fg="white")
        self.connect_button.grid(row=0, column=3, padx=5, pady=2)
        
        self.disconnect_button = tk.Button(resource_frame, text="Disconnect", command=lambda: self.disconnect_instrument(), state=tk.DISABLED, bg="darkgrey", fg="white")
        self.disconnect_button.grid(row=0, column=4, padx=5, pady=2)

        scan_settings_frame = tk.LabelFrame(self.main_frame, text="Scan Configuration (Push to Device)", padx=10, pady=10, bg="black", fg="white")
        scan_settings_frame.pack(pady=10, padx=10, fill=tk.X)

        # New: Restore Default Settings Button
        restore_button = tk.Button(scan_settings_frame, text="Restore Default Settings", command=self.restore_default_settings, bg="darkgrey", fg="white")
        restore_button.grid(row=0, column=0, columnspan=2, pady=5, sticky=tk.EW)

        row_idx = 1 # Start subsequent widgets from row 1
        tk.Label(scan_settings_frame, text="Reference Level (dBm):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        entry_ref_level = tk.Entry(scan_settings_frame, textvariable=self.desired_ref_level_var, bg="grey", fg="white", insertbackground="white")
        entry_ref_level.grid(row=row_idx, column=1, sticky=tk.EW, pady=2)
        self.desired_setting_entries["ref_level"] = entry_ref_level

        row_idx += 1
        tk.Label(scan_settings_frame, text="High Sensitivity (Preamplifier):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        check_high_sensitivity = tk.Checkbutton(scan_settings_frame, variable=self.high_sensitivity_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white")
        check_high_sensitivity.grid(row=row_idx, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["high_sensitivity"] = check_high_sensitivity

        row_idx += 1
        tk.Label(scan_settings_frame, text="Preamplifier (ON/OFF):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        check_preamp = tk.Checkbutton(scan_settings_frame, variable=self.desired_preamp_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white")
        check_preamp.grid(row=row_idx, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["preamp"] = check_preamp

        row_idx += 1
        tk.Label(scan_settings_frame, text="Scan RBW (Hz):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        
        rbw_input_frame = tk.Frame(scan_settings_frame, bg="black")
        rbw_input_frame.grid(row=row_idx, column=1, sticky=tk.EW, pady=2)
        rbw_input_frame.grid_columnconfigure(0, weight=1)
        rbw_input_frame.grid_columnconfigure(1, weight=2)

        entry_scan_rbw_segmentation = tk.Entry(rbw_input_frame, textvariable=self.desired_scan_rbw_segmentation_var, bg="grey", fg="white", insertbackground="white")
        entry_scan_rbw_segmentation.grid(row=0, column=0, sticky=tk.EW, padx=(0, 5))
        self.desired_setting_entries["scan_rbw_segmentation"] = entry_scan_rbw_segmentation
        
        def update_rbw_from_entry(*args):
            try:
                val = float(self.desired_scan_rbw_segmentation_var.get())
                if val in self.rbw_val_to_idx:
                    self.rbw_slider_index_var.set(self.rbw_val_to_idx[val])
                else:
                    closest_val = min(self.rbw_values, key=lambda x: abs(x - val))
                    self.rbw_slider_index_var.set(self.rbw_val_to_idx[closest_val])
            except ValueError:
                pass

        self.rbw_slider = tk.Scale(rbw_input_frame, 
                                   variable=self.rbw_slider_index_var,
                                   from_=0, to=len(self.rbw_values) - 1,
                                   orient=tk.HORIZONTAL, showvalue=0,
                                   resolution=1,
                                   bg="black", fg="white", troughcolor="grey", highlightbackground="black",
                                   length=200)
        self.rbw_slider.grid(row=0, column=1, sticky=tk.EW)
        
        self.desired_scan_rbw_segmentation_var.trace_add("write", update_rbw_from_entry)

        row_idx += 1
        tk.Label(scan_settings_frame, text="Frequency Shift (Hz, per cycle):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        
        freq_shift_input_frame = tk.Frame(scan_settings_frame, bg="black")
        freq_shift_input_frame.grid(row=row_idx, column=1, sticky=tk.EW, pady=2)
        freq_shift_input_frame.grid_columnconfigure(0, weight=1)
        freq_shift_input_frame.grid_columnconfigure(1, weight=2)

        entry_freq_shift = tk.Entry(freq_shift_input_frame, textvariable=self.shift_freq_var, bg="grey", fg="white", insertbackground="white")
        entry_freq_shift.grid(row=0, column=0, sticky=tk.EW, padx=(0, 5))
        self.desired_setting_entries["shift_freq"] = entry_freq_shift
        
        def update_freq_shift_from_var(*args):
            try:
                val = float(self.shift_freq_var.get())
                if val in self.freq_shift_val_to_idx:
                    self.freq_shift_slider_index_var.set(self.freq_shift_val_to_idx[val])
                else:
                    closest_val = min(self.freq_shift_values, key=lambda x: abs(x - val))
                    self.freq_shift_slider_index_var.set(self.freq_shift_val_to_idx[closest_val])
            except ValueError:
                pass

        self.freq_shift_slider = tk.Scale(freq_shift_input_frame, 
                                          variable=self.freq_shift_slider_index_var,
                                          from_=0, to=len(self.freq_shift_values) - 1,
                                          orient=tk.HORIZONTAL, showvalue=0,
                                          resolution=1,
                                          bg="black", fg="white", troughcolor="grey", highlightbackground="black")
        self.freq_shift_slider.grid(row=0, column=1, sticky=tk.EW)
        
        self.shift_freq_var.trace_add("write", update_freq_shift_from_var)

        row_idx += 1
        tk.Label(scan_settings_frame, text="Cycle Hold Time (s):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        
        self.cycle_hold_time_slider = tk.Scale(scan_settings_frame, variable=self.desired_max_hold_time_var,
                                               from_=0.5, to=10.0,
                                               orient=tk.HORIZONTAL, showvalue=1, resolution=0.5,
                                               bg="black", fg="white", troughcolor="grey", highlightbackground="black")
        self.cycle_hold_time_slider.grid(row=row_idx, column=1, sticky=tk.EW, pady=2)
        # self.cycle_hold_time_slider.set(self.config.getfloat('DEFAULT_SETTINGS', 'DEFAULT_MAXHOLD_TIME_SECONDS')) # This is now handled by _populate_vars_from_config

        row_idx += 1
        tk.Label(scan_settings_frame, text="Cycle Wait Time (s):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        
        self.cycle_wait_time_slider = tk.Scale(scan_settings_frame, variable=self.desired_cycle_wait_time_var,
                                               from_=0, to=600,
                                               orient=tk.HORIZONTAL, showvalue=1, resolution=1,
                                               bg="black", fg="white", troughcolor="grey", highlightbackground="black")
        self.cycle_wait_time_slider.grid(row=row_idx, column=1, sticky=tk.EW, pady=2)
        # self.cycle_wait_time_slider.set(self.config.getfloat('DEFAULT_SETTINGS', 'DEFAULT_CYCLE_WAIT_TIME_SECONDS')) # This is now handled by _populate_vars_from_config
        
        row_idx += 1
        tk.Label(scan_settings_frame, text="Scan Name:", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        entry_scan_name = tk.Entry(scan_settings_frame, textvariable=self.scan_name_var, bg="grey", fg="white", insertbackground="white")
        entry_scan_name.grid(row=row_idx, column=1, sticky=tk.EW, pady=2)
        self.desired_setting_entries["scan_name"] = entry_scan_name

        row_idx += 1
        output_folder_frame = tk.Frame(scan_settings_frame, bg="black")
        output_folder_frame.grid(row=row_idx, column=0, columnspan=2, sticky=tk.EW, pady=2)
        
        tk.Label(output_folder_frame, text=f"Output Folder ({os.getcwd()}{os.sep}):", bg="black", fg="white").pack(side=tk.LEFT, padx=(0, 5))
        
        entry_output_folder = tk.Entry(output_folder_frame, textvariable=self.output_folder_var, bg="grey", fg="white", insertbackground="white")
        entry_output_folder.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.desired_setting_entries["output_folder"] = entry_output_folder
        open_folder_button = tk.Button(output_folder_frame, text="Open Folder", command=lambda: self.open_output_folder(), bg="darkgrey", fg="white")
        open_folder_button.pack(side=tk.RIGHT, padx=(5, 0))

        row_idx += 1
        tk.Label(scan_settings_frame, text="Include TV Band Markers:", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        chk_tv_markers = tk.Checkbutton(scan_settings_frame, variable=self.include_tv_markers_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white")
        chk_tv_markers.grid(row=row_idx, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["include_tv_markers"] = chk_tv_markers

        row_idx += 1
        tk.Label(scan_settings_frame, text="Include Government Band Markers:", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        chk_gov_markers = tk.Checkbutton(scan_settings_frame, variable=self.include_gov_markers_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white")
        chk_gov_markers.grid(row=row_idx, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["include_gov_markers"] = chk_gov_markers

        row_idx += 1
        tk.Label(scan_settings_frame, text="Open HTML Plot After Each Cycle:", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        chk_open_html = tk.Checkbutton(scan_settings_frame, variable=self.open_html_after_complete_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white")
        chk_open_html.grid(row=row_idx, column=1, sticky=tk.W, pady=2)
        self.desired_setting_entries["open_html_after_complete"] = chk_open_html

        row_idx += 1
        button_row_frame = tk.Frame(scan_settings_frame, bg="black")
        button_row_frame.grid(row=row_idx, column=0, columnspan=2, pady=10, sticky=tk.EW)
        button_row_frame.grid_columnconfigure(0, weight=1)
        button_row_frame.grid_columnconfigure(1, weight=1)

        self.apply_button = tk.Button(button_row_frame, text="Apply Settings to Device", command=lambda: self.apply_settings_to_device(), state=tk.DISABLED, bg="darkgrey", fg="white")
        self.apply_button.grid(row=0, column=0, padx=5, sticky=tk.EW)

        self.plot_button = tk.Button(button_row_frame, text="Generate Plot (Average)", command=lambda: self.generate_average_plot(), state=tk.NORMAL, bg="blue", fg="white")
        self.plot_button.grid(row=0, column=1, padx=5, sticky=tk.EW)

        bands_and_presets_frame = tk.Frame(self.main_frame, bg="black")
        bands_and_presets_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        bands_and_presets_frame.grid_columnconfigure(0, weight=1)
        bands_and_presets_frame.grid_columnconfigure(1, weight=1)
        bands_and_presets_frame.grid_rowconfigure(0, weight=1)

        band_selection_frame = tk.LabelFrame(bands_and_presets_frame, text="Frequency Band Selection", padx=10, pady=10, bg="black", fg="white")
        band_selection_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # self.band_checkboxes = [] # Moved to __init__
        # self.band_vars = [] # Moved to __init__

        band_canvas = tk.Canvas(band_selection_frame, bg="black", highlightbackground="black")
        band_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        band_scrollbar = ttk.Scrollbar(band_selection_frame, orient="vertical", command=band_canvas.yview)
        band_scrollbar.pack(side=tk.RIGHT, fill="y")

        band_canvas.configure(yscrollcommand=band_scrollbar.set)
        band_canvas.bind('<Configure>', lambda e: band_canvas.configure(scrollregion = band_canvas.bbox("all")))

        self.inner_band_frame = tk.Frame(band_canvas, bg="black")
        band_canvas.create_window((0, 0), window=self.inner_band_frame, anchor="nw")

        for i, band in enumerate(SCAN_BAND_RANGES):
            var = tk.BooleanVar(self)
            chk = tk.Checkbutton(self.inner_band_frame, text=f"{band['Band Name']} ({band['Start MHz']:.3f}-{band['Stop MHz']:.3f} MHz)", variable=var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white")
            chk.grid(row=i, column=0, sticky=tk.W, padx=5, pady=1)
            # Initial state will be set by _set_band_checkboxes_from_config after all are created
            self.band_checkboxes.append(chk)
            self.band_vars.append({"band": band, "var": var})
        
        # After creating all checkboxes, set their initial state based on loaded config
        # This call is moved here to ensure self.band_vars is populated.
        self._set_band_checkboxes_from_config()


        preset_files_frame = tk.LabelFrame(bands_and_presets_frame, text="Device Preset Files (C:\\PRESETS\\)", padx=10, pady=10, bg="black", fg="white")
        preset_files_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

        self.load_preset_button = tk.Button(preset_files_frame, text="Load Selected Preset", command=self.load_selected_preset, state=tk.DISABLED, bg="darkgrey", fg="white")
        self.load_preset_button.pack(pady=5)

        self.preset_tree = ttk.Treeview(preset_files_frame, columns=("Name",), show="headings", selectmode="browse")
        self.preset_tree.heading("Name", text="Preset File Name")
        self.preset_tree.column("Name", width=200, anchor="w")
        self.preset_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.preset_tree.tag_configure("Mon", foreground="blue")

        preset_scrollbar = ttk.Scrollbar(preset_files_frame, orient="vertical", command=self.preset_tree.yview)
        preset_scrollbar.pack(side=tk.RIGHT, fill="y")
        self.preset_tree.configure(yscrollcommand=preset_scrollbar.set)

        self.preset_tree.bind("<<TreeviewSelect>>", self._on_preset_select)

        # Removed progress_label as it's being replaced by console output
        # self.progress_label = tk.Label(self.main_frame, text="Ready.", bg="black", fg="white")
        # self.progress_label.pack(pady=5)

        debug_frame = tk.Frame(self.main_frame, bg="black")
        debug_frame.pack(pady=10, padx=10, fill=tk.X)
        tk.Checkbutton(debug_frame, text="Enable Debug Mode (Log VISA Commands)", variable=self.debug_mode_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white").pack(anchor=tk.W)

        for i in range(5):
            resource_frame.grid_columnconfigure(i, weight=1)
        scan_settings_frame.grid_columnconfigure(0, weight=1)
        scan_settings_frame.grid_columnconfigure(1, weight=1)

        self.update_vbw_display()

    # Removed update_progress_label method
    # def update_progress_label(self, message):
    #     """Updates the progress label on the GUI."""
    #     self.progress_label.config(text=message)

    def _set_band_checkboxes_from_config(self):
        """Sets the state of band checkboxes based on the loaded config."""
        selected_bands_from_config = self.last_selected_bands_str.get()
        if selected_bands_from_config:
            selected_band_names = [name.strip() for name in selected_bands_from_config.split(',') if name.strip()]
            for item in self.band_vars:
                band_name = item["band"]["Band Name"]
                if band_name in selected_band_names:
                    item["var"].set(True)
                else:
                    item["var"].set(False)
        else:
            # If no last selected bands, default to all true (or specific default)
            for item in self.band_vars:
                item["var"].set(True) # Default to all selected if no config found

    def update_vbw_display(self):
        """Updates the VBW display based on the current RBW setting."""
        try:
            scan_rbw_val = float(self.desired_scan_rbw_segmentation_var.get())
            self.desired_vbw_display_var.set(str(int(scan_rbw_val / 3)))
        except ValueError:
            self.desired_vbw_display_var.set("Invalid RBW")

    def open_output_folder(self):
        """Opens the specified output folder in the file explorer."""
        folder_path = self.output_folder_var.get()
        if not os.path.isabs(folder_path):
            folder_path = os.path.join(os.getcwd(), folder_path)

        if not os.path.exists(folder_path):
            messagebox.showwarning("Folder Not Found", f"The folder '{folder_path}' does not exist.")
            print(f"🚫 Folder not found: {folder_path}")
            return
        
        try:
            if sys.platform == "win32":
                os.startfile(folder_path)
            elif sys.platform == "darwin":
                subprocess.run(['open', folder_path])
            else:
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
                control_disconnect_instrument(self.inst)
                self.inst = None
                print("🔌 Closed existing connection.")
            except Exception as e:
                print(f"Error closing existing connection: {e}")

        try:
            self.inst, self.instrument_model = connect_to_instrument(self.rm, selected_resource)
            if self.inst:
                self.title(f"RF Spectrum Analyzer Controller - {self.instrument_model} - {os.path.basename(sys.argv[0])}")

                ref_level = float(self.desired_ref_level_var.get())
                high_sensitivity_on = self.high_sensitivity_var.get()
                preamp_on = self.desired_preamp_var.get()
                rbw_config = int(float(self.desired_scan_rbw_segmentation_var.get()))
                vbw_config = int(float(self.desired_vbw_display_var.get()))

                if initialize_instrument(self.inst, ref_level, high_sensitivity_on, preamp_on, rbw_config, vbw_config, self.instrument_model):
                    print("Desired settings successfully applied to the instrument.")
                    query_current_instrument_settings(self.inst, MHZ_TO_HZ)
                    
                    self.start_scan_button.config(state=tk.NORMAL)
                    self.stop_scan_button.config(state=tk.DISABLED)
                    self.pause_resume_button.config(state=tk.DISABLED)
                    self.disconnect_button.config(state=tk.NORMAL)
                    self.apply_button.config(state=tk.NORMAL)
                    self.reset_setting_colors()

                    if self.instrument_model != "N9340B":
                        preset_files = control_query_device_presets(self.inst)
                        self._update_preset_tree(preset_files)
                    else:
                        print("ℹ️ Skipping device preset query for N9340B model.")
                        self._update_preset_tree([])
                        self.preset_tree.insert("", "end", values=("Presets not supported for N9340B.",), tags=("disabled",))
                        self.load_preset_button.config(state=tk.DISABLED)

                else:
                    messagebox.showerror("Initialization Failed", "Instrument initialization with desired settings failed.")
                    control_disconnect_instrument(self.inst)
                    self.inst = None
                    self._reset_gui_on_disconnect_or_error()
            else:
                messagebox.showerror("Connection Failed", "Could not connect to instrument. Check console for details.")
                self._reset_gui_on_disconnect_or_error()

        except ValueError as e:
            messagebox.showerror("Input Error", f"Invalid numeric value for instrument settings: {e}. Please check Reference Level, RBW, and Max Hold Time.")
            self._reset_gui_on_disconnect_or_error()
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
            print(f"❌ An unexpected error occurred during connection: {e}")
            self._reset_gui_on_disconnect_or_error()

            
    def disconnect_instrument(self):
        """Closes the connection to the instrument."""
        if control_disconnect_instrument(self.inst):
            self.inst = None
            self.instrument_model = None
            print("Disconnected.")
            self._reset_gui_on_disconnect_or_error()
            self.title(f"RF Spectrum Analyzer Controller - {os.path.basename(sys.argv[0])}")
        else:
            messagebox.showerror("Disconnect Error", "Failed to disconnect instrument.")
            
    def apply_settings_to_device(self):
        """Applies the desired settings from the GUI to the connected instrument."""
        if not self.inst:
            messagebox.showwarning("Not Connected", "Please connect to an instrument first to apply settings.")
            return

        try:
            ref_level = float(self.desired_ref_level_var.get())
            high_sensitivity_on = self.high_sensitivity_var.get()
            preamp_on = self.desired_preamp_var.get()
            rbw_config = int(float(self.desired_scan_rbw_segmentation_var.get()))
            vbw_config = int(float(self.desired_vbw_display_var.get()))

            if initialize_instrument(self.inst, ref_level, high_sensitivity_on, preamp_on, rbw_config, vbw_config, self.instrument_model):
                print("Desired settings successfully applied to the instrument.")
                self.reset_setting_colors()
            else:
                messagebox.showerror("Apply Failed", "Failed to apply settings to the instrument. Check console for details.")
        except ValueError as e:
            messagebox.showerror("Input Error", f"Invalid numeric value for instrument settings: {e}. Please check Reference Level, RBW, and Max Hold Time.")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred while applying settings: {e}")
            print(f"❌ Error applying settings: {e}")

    def reset_setting_colors(self):
        """Resets the text color of all desired setting entries to black."""
        for key, entry_widget in self.desired_setting_entries.items():
            if isinstance(entry_widget, tk.Entry):
                entry_widget.config(fg="white")

    def _update_preset_tree(self, preset_files):
        """Helper to update the preset Treeview."""
        for item in self.preset_tree.get_children():
            self.preset_tree.delete(item)
        if preset_files:
            for preset_name in sorted(preset_files):
                tags = ()
                if "MON" in preset_name.upper():
                    tags = ("Mon",)
                self.preset_tree.insert("", "end", values=(preset_name,), tags=tags)
        else:
            self.preset_tree.insert("", "end", values=("No .STA preset files found.",))
        self.load_preset_button.config(state=tk.DISABLED)

    def _on_preset_select(self, event):
        """Enables the Load Preset button if a preset is selected."""
        selected_items = self.preset_tree.selection()
        if selected_items and self.inst and self.instrument_model != "N9340B":
            self.load_preset_button.config(state=tk.NORMAL)
        else:
            self.load_preset_button.config(state=tk.DISABLED)

    def load_selected_preset(self):
        """
        Loads the selected preset file onto the instrument using instrument_control.
        """
        if not self.inst:
            messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
            return
        
        if self.instrument_model == "N9340B":
            messagebox.showwarning("Preset Not Supported", "Loading presets is not supported for the N9340B model.")
            print("🚫 Attempted to load preset on N9340B, which is not supported.")
            return

        selected_items = self.preset_tree.selection()
        if not selected_items:
            messagebox.showwarning("No Preset Selected", "Please select a preset file from the list to load.")
            return

        selected_preset_name = self.preset_tree.item(selected_items[0], 'values')[0]
        
        if control_load_selected_preset(self.inst, selected_preset_name, MHZ_TO_HZ):
            print(f"Preset '{selected_preset_name}' loaded successfully via instrument_control.")
        else:
            messagebox.showerror("Load Preset Failed", f"Failed to load preset: {selected_preset_name}. Check console for details.")

    def restore_default_settings(self):
        """Restores all configurable settings to their default values from config.ini."""
        print("Restoring default settings...")
        for var_name, (_, default_key, type_converter) in self.setting_var_map.items():
            if default_key and self.config.has_option('DEFAULT_SETTINGS', default_key):
                tk_var = getattr(self, var_name)
                try:
                    if type_converter == bool:
                        tk_var.set(self.config.getboolean('DEFAULT_SETTINGS', default_key))
                    elif type_converter == float:
                        tk_var.set(self.config.getfloat('DEFAULT_SETTINGS', default_key))
                    elif type_converter == int:
                        tk_var.set(self.config.getint('DEFAULT_SETTINGS', default_key))
                    else: # str
                        tk_var.set(self.config.get('DEFAULT_SETTINGS', default_key))
                    print(f"  Restored {default_key} to {tk_var.get()}")
                except ValueError as e:
                    print(f"Error restoring default for {default_key}: {e}. Skipping.")

        # Special handling for sliders that are tied to index variables, as their values are derived
        self.rbw_slider_index_var.set(self.rbw_val_to_idx.get(int(self.desired_scan_rbw_segmentation_var.get()), 0))
        self.freq_shift_slider_index_var.set(self.freq_shift_val_to_idx.get(int(self.shift_freq_var.get()), 0))
        
        # Update sliders that are directly bound to DoubleVars
        self.cycle_hold_time_slider.set(self.desired_max_hold_time_var.get())
        self.cycle_wait_time_slider.set(self.desired_cycle_wait_time_var.get())

        # Restore band selection to default (all selected)
        for item in self.band_vars:
            item["var"].set(True)

        self.update_vbw_display() # Update VBW display after RBW change
        self.reset_setting_colors() # Reset colors if any were marked
        self._save_config() # Save the restored defaults as the new last used settings
        messagebox.showinfo("Settings Restored", "Default settings have been restored and saved as last used.")


    def start_scan_thread(self):
        """Starts the scanning process in a separate thread."""
        if not self.inst:
            messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
            return
        
        if self.scanning:
            messagebox.showwarning("Scan in Progress", "A scan is already running.")
            return

        # Save configuration when scan starts - this will save current GUI settings as LAST_USED
        self._save_config()

        self.start_scan_button.config(state=tk.DISABLED)
        self.stop_scan_button.config(state=tk.NORMAL)
        self.pause_resume_button.config(state=tk.NORMAL)
        self.connect_button.config(state=tk.DISABLED)
        self.disconnect_button.config(state=tk.DISABLED)
        self.apply_button.config(state=tk.DISABLED)
        self.load_preset_button.config(state=tk.DISABLED)

        self.scanning = True
        self.paused = False
        self.pause_resume_button.config(text="Pause Scan")

        print("\nStarting continuous spectrum scan...")
        
        max_hold_enabled = self.desired_max_hold_var.get()
        max_hold_time = float(self.desired_max_hold_time_var.get()) if max_hold_enabled else 0
        
        scan_rbw_segmentation = float(self.desired_scan_rbw_segmentation_var.get())
        freq_shift_value = float(self.shift_freq_var.get()) 

        rbw_config_val = scan_rbw_segmentation
        vbw_config_val = int(rbw_config_val / 3)

        selected_bands = [item["band"] for item in self.band_vars if item["var"].get()]
        if not selected_bands:
            messagebox.showwarning("No Bands Selected", "Please select at least one frequency band to scan.")
            print("🚫 No bands selected for scan.")
            self.stop_scan()
            return

        base_output_dir = self.output_folder_var.get()
        if not os.path.exists(base_output_dir):
            os.makedirs(base_output_dir)
            print(f"Created base output directory: {base_output_dir}")

        self.scan_cycle_count = 0
        self.current_freq_offset = 0

        scan_thread = threading.Thread(target=self._run_scan, 
                                       args=(selected_bands, 
                                             scan_rbw_segmentation, freq_shift_value, 
                                             rbw_config_val, vbw_config_val, max_hold_time))
        scan_thread.daemon = True
        scan_thread.start()

    def toggle_pause_scan(self):
        """Toggles the paused state of the scan."""
        if self.scanning:
            self.paused = not self.paused
            if self.paused:
                self.pause_resume_button.config(text="Resume Scan", bg="blue")
                print("Scan Paused. Click Resume to continue.")
                print("Scan paused.")
            else:
                self.pause_resume_button.config(text="Pause Scan", bg="orange")
                print("Scan Resumed.")
                print("Scan resumed.")
        else:
            messagebox.showwarning("Scan Not Active", "No scan is currently running to pause or resume.")

    def _run_scan(self, selected_bands, scan_rbw_segmentation, freq_shift_value, rbw_config_val, vbw_config_val, max_hold_time):
        """Internal method to run the scan logic, called by the thread."""
        try:
            while self.scanning:
                while self.paused:
                    print("Scan Paused. Click Resume to continue.")
                    time.sleep(0.5)
                    if not self.scanning:
                        print("\nScan process finished (interrupted).")
                        print("Scan interrupted by user.")
                        break

                if not self.scanning:
                    print("\nScan process finished (interrupted).")
                    print("Scan interrupted by user.")
                    break

                print(f"\n--- Starting Scan Cycle {self.scan_cycle_count + 1} ---")
                print(f"Current Frequency Offset: {self.current_freq_offset} Hz (Applied to all band frequencies)")
                print(f"Scan RBW: {scan_rbw_segmentation} Hz (Constant)")

                scan_name = self.scan_name_var.get()
                if not scan_name:
                    scan_name = "UnnamedScan"
                
                rbw_str = f"RBW{int(scan_rbw_segmentation/1000):04d}K"
                max_hold_time_val = float(self.desired_max_hold_time_var.get()) if self.desired_max_hold_var.get() else 0
                hold_str = f"HOLD{int(max_hold_time_val):02d}"
                offset_str = f"Offset{int(self.current_freq_offset)}"

                datetime_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                # The individual HTML plot filename is still generated here
                html_plot_path_for_single_scan = os.path.join(self.output_folder_var.get(), f"{scan_name}_{rbw_str}_{hold_str}_{offset_str}_{datetime_str}.html")

                try:
                    # scan_bands now returns the path to the CSV file
                    scanned_data, last_successful_band_index, current_scan_csv_path = scan_bands(
                        self, self.inst, selected_bands, 
                        scan_rbw_segmentation, rbw_config_val, 
                        vbw_config_val, max_hold_time, self.current_freq_offset
                    ) 
                    
                    if not self.scanning:
                        print("\nScan process finished (interrupted after band scan).")
                        print("Scan interrupted by user.")
                        if scanned_data:
                            plot_suffix = f"{scan_name}_{rbw_str}_{hold_str}_{offset_str}_{datetime_str}_INTERRUPTED"
                            html_plot_path_for_single_scan_interrupted = os.path.join(self.output_folder_var.get(), f"{plot_suffix}.html")
                            # Pass the CSV path for plotting
                            self.after(0, self.generate_single_scan_plot_and_open_wrapper, current_scan_csv_path, plot_suffix, html_plot_path_for_single_scan_interrupted, False)
                        break

                    if scanned_data:
                        df_scan = pd.DataFrame(scanned_data, columns=['Frequency_Hz', 'Power_dBm'])
                        df_scan['Frequency_MHz'] = df_scan['Frequency_Hz'] / MHZ_TO_HZ
                        self.collected_scans_dataframes.append(df_scan[['Frequency_MHz', 'Power_dBm']].copy())
                        print(f"✅ Stored scan data for averaging. Total scans collected: {len(self.collected_scans_dataframes)}")

                    self.last_scan_data = scanned_data # Still keep in-memory for other potential uses
                    
                    print(f"Cycle scan finished. Plot will be generated from CSV: {current_scan_csv_path}")

                    plot_suffix = f"{scan_name}_{rbw_str}_{hold_str}_{offset_str}_{datetime_str}"
                    # Pass the CSV path for plotting
                    self.after(0, self.generate_single_scan_plot_and_open_wrapper, current_scan_csv_path, plot_suffix, html_plot_path_for_single_scan, self.open_html_after_complete_var.get()) 
                    
                    self.scan_cycle_count += 1
                    self.current_freq_offset += freq_shift_value

                    if self.scan_cycle_count >= 10:
                        print(f"🎉 {self.scan_cycle_count} scan cycles completed. Resetting frequency offset to 0 Hz.")
                        self.current_freq_offset = 0
                        self.scan_cycle_count = 0

                    if not self.scanning:
                        print("\nScan process finished (interrupted).")
                        print("Scan interrupted by user.")
                        break
                    
                    wait_time = float(self.desired_cycle_wait_time_var.get())
                    if wait_time > 0:
                        print(f"Waiting {wait_time} seconds for next cycle...")
                        print(f"Waiting {wait_time} seconds before next scan cycle...")
                        for _ in range(int(wait_time * 10)):
                            while self.paused:
                                print("Scan Paused. Click Resume to continue.")
                                time.sleep(0.1)
                                if not self.scanning:
                                    print("\nScan process finished (interrupted during pause in wait).")
                                    print("Scan interrupted during pause in wait.")
                                    break
                            
                            if not self.scanning:
                                print("\nScan process finished (interrupted during wait).")
                                print("Scan interrupted during wait.")
                                break
                            time.sleep(0.1)

                except Exception as e:
                    self.after(0, messagebox.showerror, "Scan Cycle Error", f"An error occurred during scan cycle: {e}")
                    print(f"❌ Scan cycle encountered an error: {e}")
                    print(f"Scan cycle error: {e}")
                    self.scanning = False
                    break

            print("\nContinuous scan process terminated.")
            print("Continuous scan terminated.")
            
        except Exception as e:
            self.after(0, messagebox.showerror, "Scan Thread Error", f"An unexpected error occurred in main scan thread: {e}")
            print(f"❌ Main scan thread encountered an error: {e}")
            print(f"Main scan thread error: {e}")
        finally:
            self.scanning = False
            self.paused = False
            self.after(100, self.reset_scan_buttons)

    def populate_resources(self):
        """Populates the VISA resource dropdown."""
        try:
            self.instrument_list = list_visa_resources(self.rm)
            if self.instrument_list:
                # Set resource_var to the last used device from config, if found.
                # If last_device is blank or not in the current list, it defaults to the first available.
                last_device = self.resource_var.get() # Get value already loaded by _populate_vars_from_config
                if not last_device or last_device not in self.instrument_list:
                    self.resource_var.set(self.instrument_list[0]) if self.instrument_list else "No Resources Found"
                
                menu = self.resource_dropdown["menu"]
                menu.delete(0, "end")
                for resource in self.instrument_list:
                    menu.add_command(label=resource, command=tk._setit(self.resource_var, resource))
            else:
                self.resource_var.set("No resources found")
                menu = self.resource_dropdown["menu"]
                menu.delete(0, "end")
                menu.add_command(label="No resources found", command=tk._setit(self.resource_var, "No resources found"))
            self.start_scan_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.DISABLED)
            self.apply_button.config(state=tk.DISABLED)
            self.connect_button.config(state=tk.NORMAL)
            self.load_preset_button.config(state=tk.DISABLED)
            self.pause_resume_button.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to list VISA resources: {e}")
            self.resource_var.set("Error listing resources")
            self._reset_gui_on_disconnect_or_error()

    def stop_scan(self):
        """Stops the ongoing scan."""
        self.scanning = False
        self.paused = False
        print("\nAttempting to stop scan... Please wait for current sweep to finish.")
        self.stop_scan_button.config(state=tk.DISABLED)
        self.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED)

    def reset_scan_buttons(self):
        """Resets the state of scan-related buttons after a scan completes or stops."""
        self.start_scan_button.config(state=tk.NORMAL)
        if self.inst:
            self.disconnect_button.config(state=tk.NORMAL)
            self.apply_button.config(state=tk.NORMAL)
            if self.preset_tree.selection() and self.instrument_model != "N9340B":
                self.load_preset_button.config(state=tk.NORMAL)
        self.stop_scan_button.config(state=tk.DISABLED)
        self.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED)

    def _reset_gui_on_disconnect_or_error(self):
        """Helper to reset GUI elements to a disconnected state."""
        self.start_scan_button.config(state=tk.DISABLED)
        self.stop_scan_button.config(state=tk.DISABLED)
        self.pause_resume_button.config(state=tk.DISABLED)
        self.disconnect_button.config(state=tk.DISABLED)
        self.apply_button.config(state=tk.DISABLED)
        self.load_preset_button.config(state=tk.DISABLED)
        for item in self.preset_tree.get_children():
            self.preset_tree.delete(item)
        self.preset_tree.insert("", "end", values=("No instrument connected.",))
        self.connect_button.config(state=tk.NORMAL)

    def generate_single_scan_plot_and_open_wrapper(self, csv_file_path, plot_title_suffix, output_html_path, auto_open_browser=True):
        """
        Wrapper to call plot_single_scan_data and handle saving/opening.
        Reads data directly from the provided CSV file path.
        """
        if not os.path.exists(csv_file_path):
            print(f"🚫 CSV file not found for plotting: {csv_file_path}")
            messagebox.showerror("Plot Error", f"CSV file not found: {csv_file_path}. Cannot generate plot.")
            return

        print(f"Generating single scan plot from CSV: {csv_file_path}...")
        try:
            # Read data from the CSV file
            # Assuming no header and columns are Frequency_MHz, Power_dBm
            df = pd.read_csv(csv_file_path, header=None, names=["Frequency_MHz", "Power_dBm"])
            
            # Convert Frequency_MHz back to Frequency_Hz for plot_single_scan_data
            # which expects Frequency_Hz as its first column.
            scanned_data_from_csv = list(zip(df['Frequency_MHz'] * MHZ_TO_HZ, df['Power_dBm']))

            fig = plot_single_scan_data(
                scanned_data_from_csv, 
                plot_title_suffix,
                include_tv_markers=self.include_tv_markers_var.get(),
                include_gov_markers=self.include_gov_markers_var.get()
            )
            
            if fig:
                fig.write_html(output_html_path)

                print(f"✅ Single scan plot generation complete: {output_html_path}")
                if auto_open_browser:
                    _open_plot_in_browser(output_html_path)
            else:
                print("🚫 Plotly figure was not generated for single scan data.")
        except Exception as e:
            messagebox.showerror("Single Plot Error", f"Failed to generate single scan plot from CSV '{csv_file_path}': {e}")
            print(f"❌ Error generating single scan plot from CSV: {e}")

    def generate_average_plot(self):
        """
        Generates an average, median, and range plot from ALL relevant CSV files
        found in the current output folder base. This is triggered by the button.
        This plot also includes all individual historical scans as overlay layers.
        """
        if self.scanning and not self.paused:
            messagebox.showwarning("Plotting Error", "Cannot generate historical average plot while a scan is in progress. Please pause or stop the scan first.")
            return

        generate_historical_average_plot(
            self.scan_name_var,
            self.output_folder_var,
            self.open_html_after_complete_var,
            self.include_tv_markers_var,
            self.include_gov_markers_var
        )

    def _update_console_line(self, text_to_display, overwrite=False):
        """
        Helper function to update the console output safely from any thread,
        handling line overwriting.
        """
        self.console_output.config(state=tk.NORMAL)
        if overwrite:
            try:
                self.console_output.delete("end-1c linestart", "end-1c")
            except TclError:
                pass
        self.console_output.insert(tk.END, text_to_display)
        self.console_output.see(tk.END)
        self.console_output.config(state=tk.DISABLED)
        self.console_output.update_idletasks()


def print_art():

    print("                                                                                               ")
    print("                                               $              $$$$$                     $$ $$$$")
    print("                                               $$$            $$   $$$$$$               $$  $$ ")
    print("                                               $$$$           $$         $$$$$          $$ $$  ")
    print("                                  $$           $$ $$          $$             $$$$$      $$$$   ")
    print("                       $$$$$$$$$$$$            $$  $$$        $                  $$$    $$$$   ")
    print("             $$$$$$$$$        $$$              $$   $$$      $$                    $$   $$$    ")
    print("   $$$$$$$$$               $$$                 $$     $$     $$                $$$$     $$     ")
    print("                         $$$                   $$$$$$$$$$$   $$          $$$$$$         $$     ")
    print("                       $$$                     $$       $$$  $  $$$$$$$$                $      ")
    print("                     $$$                       $$         $$$$$                                ")
    print("                   $$$                         $$           $$                        $ $$     ")
    print("                 $$$                $$$$$$$                 $$                        $$$      ")
    print("              $$$$            $$$$$$                        $$                  $$$$$$$$$$     ")
    print("            $$$        $$$$$$$                                   $$$$$$$$$$$$$$                ")
    print("          $$$   $$$$$$$                             $$$$$$$$$$$                                ")
    print("        $$$$$$$$                  $$$$$$$$             $$$$                                    ")
    print("      $$$              $$$$$$$$$$  $$$$              $$$$$$$$                                  ")
    print("           $$$$$$$$$$$          $$$         $$$$$$$$                                           ")
    print(" $$$$$$$$$$                  $$$    $$$$$$$$                                                   ")
    print("                         $$$$$$$$$$                                                            ")
    print("                      $$$$$                        ")
    print("    ")
    print("    ")
    print("    ")
    print("Software created for  https://zimbelaudio.com/ike-zimbel/    ")
    print("A Colaboration betweeen Ike Zimbel and Anthony P. Kuzub")
    print("    ")
    print("    ")
