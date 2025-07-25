# main_app.py
#
# This is the main entry point for the RF Spectrum Analyzer Controller application.
# It handles initial setup, checks for and installs necessary Python dependencies,
# and then launches the main graphical user interface (GUI).
# This file ensures that the application environment is ready before starting the UI.
#
# Author: Anthony Peter Kuzub
# Blog: www.Like.audio (Contributor to this project)
#
# Professional services for customizing and tailoring this software to your specific
# application can be negotiated. There is no change to use, modify, or fork this software.
#
# Build Log: https://like.audio/category/software/spectrum-scanner/
# Source Code: https://github.com/APKaudio/

import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, TclError, ttk
import os
import sys
import threading
import time # For time.sleep in _run_scan logic if it remains here or is moved
import pyvisa
import configparser
import pandas as pd # If pandas is still used directly in App, otherwise move import
import csv # Added for loading existing markers.csv
import subprocess # For BeautifulSoup installation check

# BeautifulSoup Installation Check (moved to main_app.py as requested)
try:
    from bs4 import BeautifulSoup
except ImportError:
    print("BeautifulSoup4 not found. Attempting to install...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "beautifulsoup4"])
        from bs4 import BeautifulSoup
        print("BeautifulSoup4 installed successfully.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Installation Error", f"Error installing BeautifulSoup4: {e}\\nPlease install it manually by running: pip install beautifulsoup4")
        sys.exit(1)
    except Exception as e:
        messagebox.showerror("Installation Error", f"An unexpected error occurred during BeautifulSoup4 installation: {e}")
        sys.exit(1)


# Import from new modules
from src.gui_elements import TextRedirector, print_art
from src.config_manager import load_config, save_config
from src.instrument_logic import ( # Moved this import before src.tabs
    connect_instrument_logic, disconnect_instrument_logic,
    populate_resources_logic, apply_settings_to_device_logic,
    load_selected_preset_logic, update_preset_tree, on_preset_select,
    reset_gui_on_disconnect_or_error, set_focus_frequency_logic # Ensure set_focus_frequency_logic is imported
)
from src.tabs import MarkersDisplayTab, ReportConverterTab # Now imported after instrument_logic
from src.scan_logic import start_scan_thread_logic, toggle_pause_scan_logic, run_scan_logic, stop_scan_logic, reset_scan_buttons_logic
from src.plot_logic import generate_single_scan_plot_and_open_wrapper_logic, generate_average_plot_logic
from src.settings_logic import (
    restore_default_settings_logic, update_debug_mode_global_logic,
    update_scan_rbw_from_slider_index_logic, update_freq_shift_from_slider_index_logic,
    reset_setting_colors_logic, set_band_checkboxes_from_config_logic,
    update_vbw_display_logic, open_output_folder_logic
)

# Import constants from frequency_bands.py (or move to a shared constants.py if used widely)
from utils.frequency_bands import (
    MHZ_TO_HZ,
    SCAN_BAND_RANGES,
    TV_PLOT_BAND_MARKERS,
    GOV_PLOT_BAND_MARKERS
)

# Define the config file name, ensuring it's in the same directory as the script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, 'config.ini')


class App(tk.Tk):
    """
    The main application class for the RF Spectrum Analyzer Controller.
    It inherits from `tk.Tk` and sets up the entire GUI, manages application state,
    handles user interactions, and orchestrates calls to backend functions
    for instrument control, data acquisition, and plotting.
    """
    def __init__(self):
        """
        Initializes the main application window and its components.

        Inputs: None
        Process:
            1. Calls the parent `tk.Tk` constructor.
            2. Configures the main window's appearance (background, title, geometry).
            3. Initializes PyVISA ResourceManager and related instrument state variables.
            4. Initializes Tkinter `StringVar`, `BooleanVar`, and `DoubleVar` objects
               to hold the values of various GUI settings.
            5. Sets up a mapping (`setting_var_map`) between Tkinter variables and
               their corresponding keys in the `config.ini` file for persistent storage.
            6. Makes `CONFIG_FILE` and `SCAN_BAND_RANGES` accessible to other modules.
            7. Loads configuration settings from `config.ini` using `load_config()`.
            8. Sets window geometry based on loaded config or default.
            9. Initializes sliders for RBW and Frequency Shift, binding them to update
               the relevant Tkinter variables.
            10. Binds debug mode variable to update global setting.
            11. Creates main frames for GUI layout.
            12. Sets up console output area and redirects `sys.stdout` and `sys.stderr`.
            13. Prints ASCII art and a welcome message to the console.
            14. Creates main GUI widgets by calling `create_widgets()`.
            15. Populates available VISA resources using `populate_resources_logic()`.
            16. Calls `_check_and_load_markers_csv()` to automatically load existing markers data.
            17. Binds the window closing protocol to `on_closing` for saving settings.
        Outputs: None
        """
        super().__init__()
        self.configure(bg="black") 
        self.title(f"RF Spectrum Analyzer Controller - {os.path.basename(sys.argv[0])}")
        
        self.rm = pyvisa.ResourceManager()
        self.instrument_list = []
        self.inst = None
        self.scanning = False
        self.paused = False
        self.last_scan_data = None
        self.collected_scans_dataframes = []
        self.instrument_model = None

        self.scan_cycle_count = 0
        self.current_freq_offset = 0

        self.desired_setting_entries = {}

        self.blink_id = None
        self.blink_on = False

        self.desired_ref_level_var = tk.StringVar(self)
        self.desired_preamp_var = tk.BooleanVar(self)
        self.high_sensitivity_var = tk.BooleanVar(self)
        self.desired_max_hold_var = tk.BooleanVar(self)
        self.desired_max_hold_time_var = tk.DoubleVar(self)
        self.desired_rbw_var = tk.StringVar(self)
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
        self.last_selected_bands_str = tk.StringVar(self)
        self.default_focus_width_var = tk.DoubleVar(self, value=10000.0) # New: Default focus width variable

        self.band_checkboxes = []
        self.band_vars = []

        self.setting_var_map = {
            'desired_ref_level_var': ('last_reference_level_dbm', 'default_reference_level_dbm', float),
            'desired_preamp_var': ('last_preamp_on', 'default_preamp_on', bool),
            'high_sensitivity_var': ('last_high_sensitivity', 'default_high_sensitivity', bool),
            'desired_max_hold_var': ('last_maxhold_enabled', 'default_maxhold_enabled', bool),
            'desired_max_hold_time_var': ('last_maxhold_time_seconds', 'default_maxhold_time_seconds', float),
            'desired_rbw_var': ('last_rbw_step_size_hz', 'default_rbw_step_size_hz', int),
            'desired_cycle_wait_time_var': ('last_cycle_wait_time_seconds', 'default_cycle_wait_time_seconds', float),
            'output_folder_var': ('last_scan_directory', 'default_scan_directory', str),
            'scan_name_var': ('last_scan_name', 'default_scan_name', str),
            'resource_var': ('last_gpib_device', '', str),
            'include_gov_markers_var': ('last_include_gov_markers', 'default_include_gov_markers', bool),
            'include_tv_markers_var': ('last_include_tv_markers', 'default_include_tv_markers', bool),
            'open_html_after_complete_var': ('last_open_html_after_complete', 'default_open_html_after_complete', bool),
            'desired_scan_rbw_segmentation_var': ('last_scan_rbw_hz', 'default_scan_rbw_hz', float),
            'shift_freq_var': ('last_freq_shift_hz', 'default_freq_shift_hz', float),
            'debug_mode_var': ('last_debug_mode', 'default_debug_mode', bool),
            'last_selected_bands_str': ('last_selected_bands', 'default_selected_bands', str),
            'default_focus_width_var': ('last_default_focus_width', 'default_default_focus_width', float), # New entry for focus width
        }

        self.CONFIG_FILE = CONFIG_FILE # Make CONFIG_FILE accessible to config_manager
        self.SCAN_BAND_RANGES = SCAN_BAND_RANGES # Make SCAN_BAND_RANGES accessible to config_manager

        self.config = configparser.ConfigParser()
        load_config(self)

        initial_geometry = self.config.get('LAST_USED_SETTINGS', 'last_window_geometry') # Corrected key name
        if not initial_geometry:
            initial_geometry = self.config.get('DEFAULT_SETTINGS', 'DEFAULT_WINDOW_GEOMETRY')
        self.geometry(initial_geometry)

        self.rbw_values = [5000, 10000, 25000, 50000, 100000]
        self.rbw_val_to_idx = {val: i for i, val in enumerate(self.rbw_values)}
        self.rbw_slider_index_var = tk.IntVar(self, value=self.rbw_val_to_idx.get(int(self.desired_scan_rbw_segmentation_var.get()), 0))
        self.rbw_slider_index_var.trace_add("write", self._update_scan_rbw_from_slider_index)

        self.freq_shift_values = [0, 500, 1000, 5000, 10000]
        self.freq_shift_val_to_idx = {val: i for i, val in enumerate(self.freq_shift_values)}
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

        print("--- RF Spectrum Analyzer Controller - GUI Initialized ---")
        self.create_widgets()
        self.after(0, populate_resources_logic, self)
        self.after(100, self._check_and_load_markers_csv) # Auto-load markers.csv on startup

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """
        Handler for the window closing event. This function is called when
        the user attempts to close the application window.

        Inputs: None
        Process:
            1. Calls `save_config()` to persist current settings.
            2. Destroys the Tkinter root window, effectively closing the application.
        Outputs: None
        """
        save_config(self)
        self.destroy()

    def _update_debug_mode_global(self, *args):
        """
        Wrapper for `update_debug_mode_global_logic`.

        Inputs:
            *args: Standard Tkinter trace arguments (not used).
        Process:
            1. Calls `update_debug_mode_global_logic` from `src.settings_logic`, passing `self`.
        Outputs: None
        """
        update_debug_mode_global_logic(self, *args)

    def _update_scan_rbw_from_slider_index(self, *args):
        """
        Wrapper for `update_scan_rbw_from_slider_index_logic`.

        Inputs:
            *args: Standard Tkinter trace arguments (not used).
        Process:
            1. Calls `update_scan_rbw_from_slider_index_logic` from `src.settings_logic`, passing `self`.
        Outputs: None
        """
        update_scan_rbw_from_slider_index_logic(self, *args)

    def _update_freq_shift_from_slider_index(self, *args):
        """
        Wrapper for `update_freq_shift_from_slider_index_logic`.

        Inputs:
            *args: Standard Tkinter trace arguments (not used).
        Process:
            1. Calls `update_freq_shift_from_slider_index_logic` from `src.settings_logic`, passing `self`.
        Outputs: None
        """
        update_freq_shift_from_slider_index_logic(self, *args)

    def create_widgets(self):
        """
        Constructs all the graphical user interface elements of the application.
        This includes frames, labels, entry fields, buttons, checkboxes, sliders,
        and treeviews, arranging them using Tkinter's `pack` and `grid` layout managers.

        Inputs: None
        Process:
            1. Configures `ttk.Style` for consistent widget appearance.
            2. Creates `resource_frame` for instrument connection controls:
               - VISA Resource dropdown (`resource_dropdown`)
               - Refresh, Connect, Disconnect buttons (commands linked to `instrument_logic`).
            3. Creates `scan_settings_frame` for instrument configuration:
               - Restore Default Settings button (command linked to `settings_logic`).
               - Entry fields and checkboxes for Reference Level, High Sensitivity, Preamplifier.
               - Sliders and entry fields for Scan RBW and Frequency Shift.
               - Sliders for Cycle Hold Time and Cycle Wait Time.
               - Entry fields for Scan Name and Output Folder, with an "Open Folder" button (command linked to `settings_logic`).
               - Checkboxes for including TV and Government band markers in plots.
               - Checkbox for auto-opening HTML plots.
               - Buttons for "Apply Settings to Device" (command linked to `instrument_logic`) and "Generate Plot (Average)" (command linked to `plot_logic`).
            4. Creates `ttk.Notebook` for tabbed interface:
               - "Frequency Band Selection" tab with a scrollable canvas for frequency band checkboxes.
               - "Device Preset Files" tab with a "Load Selected Preset" button (command linked to `instrument_logic`) and a `ttk.Treeview`
                 to display device preset files (binds to `instrument_logic.on_preset_select`).
               - "Report Converter" tab with the `ReportConverterTab` from `src.tabs`.
            5. Creates a `debug_frame` with a checkbox to enable/disable debug mode.
            6. Configures grid weights for responsive layout.
            7. Calls `update_vbw_display_logic()` to initialize the VBW display.
        Outputs: None
        """
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

        style.configure("TNotebook", background="black", borderwidth=0)
        style.configure("TNotebook.Tab", background="darkgrey", foreground="white",
                        lightcolor="grey", darkcolor="grey", borderwidth=0)
        style.map("TNotebook.Tab", background=[("selected", "grey")],
                                   foreground=[("selected", "white")])


        resource_frame = tk.LabelFrame(self.main_frame, text="Instrument Connection", padx=10, pady=10, bg="black", fg="white")
        resource_frame.pack(pady=10, padx=10, fill=tk.X)

        tk.Label(resource_frame, text="VISA Resource:", bg="black", fg="white").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.resource_dropdown = tk.OptionMenu(resource_frame, self.resource_var, "No Resources Found")
        self.resource_dropdown.config(bg="grey", fg="white", highlightbackground="grey", highlightcolor="grey", activebackground="darkgrey", activeforeground="white")
        self.resource_dropdown.grid(row=0, column=1, sticky=tk.EW, pady=2)
        self.resource_dropdown["menu"].config(bg="grey", fg="white", activebackground="darkgrey", activeforeground="white")

        self.refresh_button = tk.Button(resource_frame, text="Refresh", command=lambda: populate_resources_logic(self), bg="darkgrey", fg="white")
        self.refresh_button.grid(row=0, column=2, padx=5, pady=2)

        self.connect_button = tk.Button(resource_frame, text="Connect", command=lambda: connect_instrument_logic(self), state=tk.DISABLED, bg="darkgrey", fg="white")
        self.connect_button.grid(row=0, column=3, padx=5, pady=2)
        
        self.disconnect_button = tk.Button(resource_frame, text="Disconnect", command=lambda: disconnect_instrument_logic(self), state=tk.DISABLED, bg="darkgrey", fg="white")
        self.disconnect_button.grid(row=0, column=4, padx=5, pady=2)

        scan_settings_frame = tk.LabelFrame(self.main_frame, text="Scan Configuration (Push to Device)", padx=10, pady=10, bg="black", fg="white")
        scan_settings_frame.pack(pady=10, padx=10, fill=tk.X)

        restore_button = tk.Button(scan_settings_frame, text="Restore Default Settings", command=lambda: restore_default_settings_logic(self), bg="darkgrey", fg="white")
        restore_button.grid(row=0, column=0, columnspan=2, pady=5, sticky=tk.EW)

        row_idx = 1
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

        row_idx += 1
        tk.Label(scan_settings_frame, text="Cycle Wait Time (s):", bg="black", fg="white").grid(row=row_idx, column=0, sticky=tk.W, pady=2)
        
        self.cycle_wait_time_slider = tk.Scale(scan_settings_frame, variable=self.desired_cycle_wait_time_var,
                                               from_=0, to=600,
                                               orient=tk.HORIZONTAL, showvalue=1, resolution=1,
                                               bg="black", fg="white", troughcolor="grey", highlightbackground="black")
        self.cycle_wait_time_slider.grid(row=row_idx, column=1, sticky=tk.EW, pady=2)
        
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
        open_folder_button = tk.Button(output_folder_frame, text="Open Folder", command=lambda: open_output_folder_logic(self), bg="darkgrey", fg="white")
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

        self.apply_button = tk.Button(button_row_frame, text="Apply Settings to Device", command=lambda: apply_settings_to_device_logic(self), state=tk.DISABLED, bg="darkgrey", fg="white")
        self.apply_button.grid(row=0, column=0, padx=5, sticky=tk.EW)

        self.plot_button = tk.Button(button_row_frame, text="Generate Plot (Average)", command=lambda: generate_average_plot_logic(self), state=tk.NORMAL, bg="blue", fg="white")
        self.plot_button.grid(row=0, column=1, padx=5, sticky=tk.EW)

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        band_selection_tab = tk.Frame(self.notebook, bg="black")
        self.notebook.add(band_selection_tab, text="Frequency Band Selection")

        band_canvas = tk.Canvas(band_selection_tab, bg="black", highlightbackground="black")
        band_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        band_scrollbar = ttk.Scrollbar(band_selection_tab, orient="vertical", command=band_canvas.yview)
        band_scrollbar.pack(side=tk.RIGHT, fill="y")

        band_canvas.configure(yscrollcommand=band_scrollbar.set)
        band_canvas.bind('<Configure>', lambda e: band_canvas.configure(scrollregion = band_canvas.bbox("all")))

        self.inner_band_frame = tk.Frame(band_canvas, bg="black")
        band_canvas.create_window((0, 0), window=self.inner_band_frame, anchor="nw")

        for i, band in enumerate(SCAN_BAND_RANGES):
            var = tk.BooleanVar(self)
            chk = tk.Checkbutton(self.inner_band_frame, text=f"{band['Band Name']} ({band['Start MHz']:.3f}-{band['Stop MHz']:.3f} MHz)", variable=var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white")
            chk.grid(row=i, column=0, sticky=tk.W, padx=5, pady=1)
            self.band_checkboxes.append(chk)
            self.band_vars.append({"band": band, "var": var})
        
        set_band_checkboxes_from_config_logic(self)

        preset_files_tab = tk.Frame(self.notebook, bg="black")
        self.notebook.add(preset_files_tab, text="Device Preset Files")

        self.load_preset_button = tk.Button(preset_files_tab, text="Load Selected Preset", command=lambda: load_selected_preset_logic(self), state=tk.DISABLED, bg="darkgrey", fg="white")
        self.load_preset_button.pack(pady=5)

        self.preset_tree = ttk.Treeview(preset_files_tab, columns=("Name",), show="headings", selectmode="browse")
        self.preset_tree.heading("Name", text="Preset File Name")
        self.preset_tree.column("Name", width=200, anchor="w")
        self.preset_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.preset_tree.tag_configure("Mon", foreground="blue")

        preset_scrollbar = ttk.Scrollbar(preset_files_tab, orient="vertical", command=self.preset_tree.yview)
        preset_scrollbar.pack(side=tk.RIGHT, fill="y")
        self.preset_tree.configure(yscrollcommand=preset_scrollbar.set)

        self.preset_tree.bind("<<TreeviewSelect>>", lambda event: on_preset_select(self, event))
        
        # Pass self (the App instance) to ReportConverterTab
        report_converter_tab = ReportConverterTab(self.notebook, app_instance=self, bg="black")
        self.notebook.add(report_converter_tab, text="Report Converter")

        debug_frame = tk.Frame(self.main_frame, bg="black")
        debug_frame.pack(pady=10, padx=10, fill=tk.X)
        tk.Checkbutton(debug_frame, text="Enable Debug Mode (Log VISA Commands)", variable=self.debug_mode_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white").pack(anchor=tk.W)

        for i in range(5):
            resource_frame.grid_columnconfigure(i, weight=1)
        scan_settings_frame.grid_columnconfigure(0, weight=1)
        scan_settings_frame.grid_columnconfigure(1, weight=1)

        update_vbw_display_logic(self)

    def add_markers_tab(self, headers, rows):
        """
        Adds a new 'Markers Display' tab to the notebook and populates it
        with the extracted marker data. This tab will display the zones,
        groups, and devices in a structured way.

        Inputs:
            headers (list): A list of column headers for the marker data.
            rows (list): A list of dictionaries, where each dictionary represents
                         a row of marker data with keys matching the headers.
        Process:
            1. Checks if a "Markers Display" tab already exists and removes it to ensure a fresh display.
            2. Creates a new `MarkersDisplayTab` instance, passing the extracted
               `headers` and `rows`.
            3. Adds this new tab to the `self.notebook` with the text "Markers Display".
            4. Selects the newly created tab to bring it into view.
        Outputs: None
        """
        for i, tab_id in enumerate(self.notebook.tabs()):
            tab_text = self.notebook.tab(tab_id, "text")
            if tab_text == "Markers Display":
                self.notebook.forget(tab_id)
                print("Existing 'Markers Display' tab removed.")
                break
        # Pass self (the App instance) to MarkersDisplayTab
        markers_display_tab = MarkersDisplayTab(self.notebook, headers=headers, rows=rows, app_instance=self, bg="black")
        self.notebook.add(markers_display_tab, text="Markers Display")
        self.notebook.select(markers_display_tab)

    def _check_and_load_markers_csv(self):
        """
        Checks for the existence of 'MARKERS.CSV' in the default output directory
        and, if found, automatically loads its content and creates the "Markers Display" tab.

        Inputs: None
        Process:
            1. Constructs the full path to `MARKERS.CSV` using `self.output_folder_var.get()`.
            2. Checks if the file exists using `os.path.exists()`.
            3. If the file exists:
               - Initializes empty lists for `headers` and `rows`.
               - Opens and reads the CSV file using `csv.DictReader` to get headers and rows.
               - Calls `self.add_markers_tab()` to create and populate the tab.
               - Prints a success message to the console.
            4. If the file does not exist, prints an informational message.
            5. Includes error handling for file reading.
        Outputs: None (may create a new tab and print to console)
        """
        markers_csv_path = os.path.join(self.output_folder_var.get(), 'MARKERS.CSV')
        
        if os.path.exists(markers_csv_path):
            print(f"Found existing MARKERS.CSV at: {markers_csv_path}. Attempting to load...")
            headers = []
            rows = []
            try:
                with open(markers_csv_path, 'r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    headers = reader.fieldnames
                    for row in reader:
                        rows.append(row)
                
                if headers and rows:
                    self.add_markers_tab(headers, rows)
                    print("✅ MARKERS.CSV loaded successfully and 'Markers Display' tab created.")
                else:
                    print("ℹ️ MARKERS.CSV found but appears empty or malformed. Skipping tab creation.")
            except Exception as e:
                print(f"❌ Error loading MARKERS.CSV: {e}")
                messagebox.showerror("Error Loading Markers", f"Failed to load MARKERS.CSV: {e}")
        else:
            print("ℹ️ No MARKERS.CSV found at startup. 'Markers Display' tab will not be automatically created.")


    def open_output_folder(self):
        """
        Wrapper for `open_output_folder_logic`.

        Inputs: None
        Process:
            1. Calls `open_output_folder_logic` from `src.settings_logic`, passing `self`.
        Outputs: None
        """
        open_output_folder_logic(self)

    def connect_instrument(self):
        """
        Wrapper for `connect_instrument_logic`.

        Inputs: None
        Process:
            1. Calls `connect_instrument_logic` from `src.instrument_logic`, passing `self`.
        Outputs: None
        """
        connect_instrument_logic(self)

    def disconnect_instrument(self):
        """
        Wrapper for `disconnect_instrument_logic`.

        Inputs: None
            None
        Process:
            1. Calls `disconnect_instrument_logic` from `src.instrument_logic`, passing `self`.
        Outputs: None
        """
        disconnect_instrument_logic(self)

    def apply_settings_to_device(self):
        """
        Wrapper for `apply_settings_to_device_logic`.

        Inputs: None
        Process:
            1. Calls `apply_settings_to_device_logic` from `src.instrument_logic`, passing `self`.
        Outputs: None
        """
        apply_settings_to_device_logic(self)

    def reset_setting_colors(self):
        """
        Wrapper for `reset_setting_colors_logic`.

        Inputs: None
        Process:
            1. Calls `reset_setting_colors_logic` from `src.settings_logic`, passing `self`.
        Outputs: None
        """
        reset_setting_colors_logic(self)

    def load_selected_preset(self):
        """
        Wrapper for `load_selected_preset_logic`.

        Inputs: None
        Process:
            1. Calls `load_selected_preset_logic` from `src.instrument_logic`, passing `self`.
        Outputs: None
        """
        load_selected_preset_logic(self)

    def restore_default_settings(self):
        """
        Wrapper for `restore_default_settings_logic`.

        Inputs: None
        Process:
            1. Calls `restore_default_settings_logic` from `src.settings_logic`, passing `self`.
        Outputs: None
        """
        restore_default_settings_logic(self)

    def start_scan_thread(self):
        """
        Wrapper for `start_scan_thread_logic`.

        Inputs: None
        Process:
            1. Calls `start_scan_thread_logic` from `src.scan_logic`, passing `self`.
        Outputs: None
        """
        start_scan_thread_logic(self)

    def toggle_pause_scan(self):
        """
        Wrapper for `toggle_pause_scan_logic`.

        Inputs: None
        Process:
            1. Calls `toggle_pause_scan_logic` from `src.scan_logic`, passing `self`.
        Outputs: None
        """
        toggle_pause_scan_logic(self)

    def _run_scan(self, selected_bands, scan_rbw_segmentation, freq_shift_value, rbw_config_val, vbw_config_val, max_hold_time):
        """
        Wrapper for `run_scan_logic`. This method is designed to be run in a separate thread.

        Inputs:
            selected_bands (list): A list of dictionaries, each representing a frequency band to scan.
            scan_rbw_segmentation (float): The Resolution Bandwidth (RBW) to use for scan segments.
            freq_shift_value (float): The frequency offset (in Hz) to apply per scan cycle.
            rbw_config_val (float): RBW value to configure on the instrument.
            vbw_config_val (float): VBW value to configure on the instrument.
            max_hold_time (float): Duration in seconds for which MAX Hold should be active.
        Process:
            1. Calls `run_scan_logic` from `src.scan_logic`, passing `self` and all scan parameters.
        Outputs: None
        """
        run_scan_logic(self, selected_bands, scan_rbw_segmentation, freq_shift_value, rbw_config_val, vbw_config_val, max_hold_time)

    def stop_scan(self):
        """
        Wrapper for `stop_scan_logic`.

        Inputs: None
        Process:
            1. Calls `stop_scan_logic` from `src.scan_logic`, passing `self`.
        Outputs: None
        """
        stop_scan_logic(self)

    def reset_scan_buttons(self):
        """
        Wrapper for `reset_scan_buttons_logic`.

        Inputs: None
        Process:
            1. Calls `reset_scan_buttons_logic` from `src.scan_logic`, passing `self`.
        Outputs: None
        """
        reset_scan_buttons_logic(self)

    def _reset_gui_on_disconnect_or_error(self):
        """
        Wrapper for `reset_gui_on_disconnect_or_error` from `src.instrument_logic`.

        Inputs: None
        Process:
            1. Calls `reset_gui_on_disconnect_or_error` from `src.instrument_logic`, passing `self`.
        Outputs: None
        """
        reset_gui_on_disconnect_or_error(self)

    def generate_single_scan_plot_and_open_wrapper(self, csv_file_path, plot_title_suffix, output_html_path, auto_open_browser=True):
        """
        Wrapper for `generate_single_scan_plot_and_open_wrapper_logic`.

        Inputs:
            csv_file_path (str): The full path to the CSV file containing the scan data.
            plot_title_suffix (str): A string to append to the plot's main title.
            output_html_path (str): The full path where the generated HTML plot should be saved.
            auto_open_browser (bool, optional): If True, the generated HTML plot will be
                                                automatically opened in the default web browser. Defaults to True.
        Process:
            1. Calls `generate_single_scan_plot_and_open_wrapper_logic` from `src.plot_logic`,
               passing `self` and all plot parameters.
        Outputs: None
        """
        generate_single_scan_plot_and_open_wrapper_logic(self, csv_file_path, plot_title_suffix, output_html_path, auto_open_browser)

    def generate_average_plot(self):
        """
        Wrapper for `generate_average_plot_logic`.

        Inputs: None
        Process:
            1. Calls `generate_average_plot_logic` from `src.plot_logic`, passing `self`.
        Outputs: None
        """
        generate_average_plot_logic(self)

    def _update_console_line(self, text_to_display, overwrite=False):
        """
        Helper function to safely update the console output widget from any thread.
        It supports overwriting the current line, which is useful for displaying
        dynamic progress updates (e.g., progress bars or countdowns).

        Inputs:
            text_to_display (str): The text string to insert into the console.
            overwrite (bool, optional): If True, the current last line in the console
                                        will be deleted before inserting `text_to_display`,
                                        creating an overwrite effect. Defaults to False.
        Process:
            1. Sets the console widget state to `tk.NORMAL` to allow editing.
            2. If `overwrite` is True, attempts to delete the last line.
            3. Inserts `text_to_display` at the end of the console.
            4. Scrolls the console to the end to show the latest text.
            5. Sets the console widget state back to `tk.DISABLED` to prevent user editing.
            6. Updates Tkinter idle tasks to ensure immediate display.
        Outputs: None
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

    def _start_connect_button_blink(self):
        """
        Initiates a blinking visual effect for the "Connect" button.
        This serves as a visual cue to the user that a connection is possible
        but not yet established.

        Inputs: None
        Process:
            1. Checks if the blinking effect is already active (`self.blink_id is None`).
            2. Sets `self.blink_on` to `True`.
            3. Calls `_blink_connect_button()` to start the actual blinking loop.
        Outputs: None
        """
        if self.blink_id is None:
            self.blink_on = True
            self._blink_connect_button()

    def _stop_connect_button_blink(self):
        """
        Stops the blinking effect on the "Connect" button and resets its color
        to its default state.

        Inputs: None
        Process:
            1. If `self.blink_id` is not None (meaning blinking is active),
               cancels the scheduled `after` event.
            2. Resets `self.blink_id` to `None`.
            3. Resets the button's background and foreground colors to their defaults.
            4. Sets `self.blink_on` to `False`.
        Outputs: None
        """
        if self.blink_id is not None:
            self.after_cancel(self.blink_id)
            self.blink_id = None
        self.connect_button.config(bg="darkgrey", fg="white")
        self.blink_on = False

    def _blink_connect_button(self):
        """
        Toggles the background and foreground colors of the "Connect" button
        at a set interval (500ms) to create the blinking animation.

        Inputs:
            None
        Process:
            1. If `self.blink_on` is `True`:
               - Retrieves the current background color of the button.
               - Toggles the colors between "darkgrey" (default) and "lightblue" (highlight).
               - Schedules itself to be called again after 500 milliseconds using `self.after()`.
        Outputs: None
        """
        if self.blink_on:
            current_bg = self.connect_button.cget("bg")
            if current_bg == "darkgrey":
                self.connect_button.config(bg="lightblue", fg="black")
            else:
                self.connect_button.config(bg="darkgrey", fg="white")
            self.blink_id = self.after(500, self._blink_connect_button)

if __name__ == "__main__":
    app = App()
    app.mainloop()
