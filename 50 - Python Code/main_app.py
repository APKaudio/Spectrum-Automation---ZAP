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
#
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
import inspect # Import inspect module


# Import local modules
from src.config_manager import load_config, save_config
from src.gui_elements import TextRedirector, print_art # Corrected import from print_logo to print_art
from src.instrument_logic import (
    populate_resources_logic, connect_instrument_logic, disconnect_instrument_logic,
    apply_settings_to_device_logic, load_selected_preset_logic, query_device_presets_logic,
    set_focus_frequency_logic, set_marker_and_trace_modes_logic
)
from src.scan_logic import (
    start_scan_thread_logic, run_scan_logic,
    stop_scan_logic, reset_scan_buttons_logic, pause_resume_scan_logic
)
from src.settings_logic import (
    restore_default_settings_logic, open_output_folder_logic,
    open_preset_folder_logic, # open_report_folder_logic, # Removed as it's redundant with open_output_folder_logic
    update_scan_rbw_from_slider_index_logic,
    update_max_hold_time_from_slider_index_logic,
    update_cycle_wait_time_from_slider_index_logic
)
from src.plot_logic import generate_single_scan_plot_and_open_wrapper_logic, generate_average_plot_logic
from utils.frequency_bands import SCAN_BAND_RANGES, MHZ_TO_HZ, VBW_RBW_RATIO
from utils.instrument_control import set_debug_mode, debug_print, set_log_visa_commands_mode

from src.report_converter_tab import ReportConverterTab
from src.instrument_preset_tab import PresetFilesTab


from src.marker_logic import MarkersDisplayTab


# --- Dependency Check and Installation ---
# This block ensures that all necessary Python packages are installed before the
# application attempts to import and use them. This is critical for user experience,
# as it prevents ModuleNotFoundError and guides the user if manual installation is needed.
REQUIRED_PACKAGES = {
    'pyvisa': 'pyvisa',
    'pandas': 'pandas',
    'beautifulsoup4': 'bs4', # FIX: Map 'beautifulsoup4' to its import name 'bs4'
    'lxml': 'lxml',
    'pdfplumber': 'pdfplumber',
    'plotly': 'plotly',
    'numpy': 'numpy'
}

def check_and_install_dependencies():
    """
    Checks if required Python packages are installed and attempts to install them if not.
    """
    print("Checking for required Python packages...")
    packages_to_install = []
    for package_name, import_name in REQUIRED_PACKAGES.items():
        try:
            __import__(import_name) # FIX: Use the actual import name for checking
        except ImportError:
            packages_to_install.append(package_name) # Use the package name for pip install

    if packages_to_install:
        print(f"Missing packages: {', '.join(packages_to_install)}. Attempting to install...")
        try:
            # Use pip to install missing packages
            subprocess.check_call([sys.executable, "-m", "pip", "install", *packages_to_install])
            print("✅ All missing packages installed successfully.")
        except subprocess.CalledProcessError as e:
            messagebox.showerror(
                "Installation Error",
                f"Failed to install one or more required packages: {e}\n"
                "Please install them manually by running:\n"
                f"pip install {' '.join(packages_to_install)}"
            )
            sys.exit(1)
        except Exception as e:
            messagebox.showerror("Installation Error", f"An unexpected error occurred during package installation: {e}")
            sys.exit(1)
    else:
        print("✅ All required packages are already installed.")

# Run dependency check at application start
check_and_install_dependencies()


class ScanConfigurationTab(ttk.Frame):
    """
    A Tkinter Frame that serves as a tab for scan configuration,
    arranging instrument connection, scan control, settings, and frequency bands
    in a 2x2 grid.
    """
    def __init__(self, master=None, app_instance=None, **kwargs):
        super().__init__(master, **kwargs)
        self.app_instance = app_instance

        # Configure grid for 2x2 layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Top Left: Instrument Connection Frame (changed from Top Right)
        instrument_frame = ttk.LabelFrame(self, text="Instrument Connection")
        instrument_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew") # Changed column to 0
        instrument_frame.grid_columnconfigure(0, weight=1) # Make dropdown expandable

        ttk.Label(instrument_frame, text="VISA Resource:").grid(row=0, column=0, padx=5, pady=2, sticky="w")

        self.app_instance.resource_dropdown = ttk.OptionMenu(instrument_frame, self.app_instance.resource_var, "", *self.app_instance.instrument_list)
        self.app_instance.resource_dropdown.grid(row=1, column=0, padx=5, pady=2, sticky="ew")

        self.app_instance.connect_button = ttk.Button(instrument_frame, text="Connect", command=lambda: connect_instrument_logic(self.app_instance), state=tk.DISABLED, style='GreyText.TButton')
        self.app_instance.connect_button.grid(row=2, column=0, padx=5, pady=2, sticky="ew")

        self.app_instance.disconnect_button = ttk.Button(instrument_frame, text="Disconnect", command=lambda: disconnect_instrument_logic(self.app_instance), state=tk.DISABLED, style='GreyText.TButton')
        self.app_instance.disconnect_button.grid(row=3, column=0, padx=5, pady=2, sticky="ew")

        self.app_instance.apply_button = ttk.Button(instrument_frame, text="Apply Settings to Device", command=lambda: apply_settings_to_device_logic(self.app_instance), state=tk.DISABLED, style='Accent.TButton')
        self.app_instance.apply_button.grid(row=4, column=0, padx=5, pady=2, sticky="ew")

        # Note: self.app_instance.preset_files_tab will be available after App._create_widgets
        # It's safer to set the command for this button later, or ensure preset_files_tab is initialized
        # before this button is used. For now, assuming it's okay due to App's init order.
        self.app_instance.load_preset_button = ttk.Button(instrument_frame, text="Load Selected Preset", command=lambda: load_selected_preset_logic(self.app_instance, self.app_instance.preset_files_tab.get_selected_preset()), state=tk.DISABLED, style='Accent.TButton')
        self.app_instance.load_preset_button.grid(row=5, column=0, padx=5, pady=2, sticky="ew")


        # Top Right: Scan Control Frame (changed from Top Left)
        scan_control_frame = ttk.LabelFrame(self, text="Scan Control")
        scan_control_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew") # Changed column to 1
        scan_control_frame.grid_columnconfigure(0, weight=1)

        self.app_instance.start_scan_button = ttk.Button(scan_control_frame, text="Start Scan", command=lambda: start_scan_thread_logic(self.app_instance), state=tk.DISABLED, style='Green.TButton')
        self.app_instance.start_scan_button.grid(row=0, column=0, padx=5, pady=2, sticky="ew")

        self.app_instance.stop_scan_button = ttk.Button(scan_control_frame, text="Stop Scan", command=lambda: stop_scan_logic(self.app_instance), state=tk.NORMAL, style='Red.TButton') # Changed to NORMAL for testing
        self.app_instance.stop_scan_button.grid(row=1, column=0, padx=5, pady=2, sticky="ew")

        self.app_instance.pause_resume_button = ttk.Button(scan_control_frame, text="Pause Scan", command=lambda: pause_resume_scan_logic(self.app_instance), state=tk.NORMAL, style='Orange.TButton') # Changed to NORMAL for testing
        self.app_instance.pause_resume_button.grid(row=2, column=0, padx=5, pady=2, sticky="ew")
        self.app_instance.pause_blink_id = None # For blinking effect

        self.app_instance.plot_button = ttk.Button(scan_control_frame, text="Generate Average Plot", command=lambda: generate_average_plot_logic(self.app_instance), state=tk.DISABLED, style='Blue.TButton')
        self.app_instance.plot_button.grid(row=3, column=0, padx=5, pady=2, sticky="ew")

        # Scan Name and Directory
        ttk.Label(scan_control_frame, text="Scan Name:").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.app_instance.scan_name_entry = ttk.Entry(scan_control_frame, textvariable=self.app_instance.scan_name_var)
        self.app_instance.scan_name_entry.grid(row=5, column=0, padx=5, pady=2, sticky="ew")

        ttk.Label(scan_control_frame, text="Output Directory:").grid(row=6, column=0, padx=5, pady=2, sticky="w")
        output_dir_frame = ttk.Frame(scan_control_frame)
        output_dir_frame.grid(row=7, column=0, padx=5, pady=2, sticky="ew")
        output_dir_frame.grid_columnconfigure(0, weight=1)
        self.app_instance.output_folder_entry = ttk.Entry(output_dir_frame, textvariable=self.app_instance.output_folder_var)
        self.app_instance.output_folder_entry.grid(row=0, column=0, sticky="ew")
        ttk.Button(output_dir_frame, text="Browse", command=self.app_instance._browse_output_folder).grid(row=0, column=1, sticky="e")

        # FIX: Added "Open Output Folder" button to Scan Control frame
        ttk.Button(scan_control_frame, text="Open Output Folder", command=lambda: open_output_folder_logic(self.app_instance.output_folder_var.get())).grid(row=8, column=0, padx=5, pady=2, sticky="ew")


        # Bottom Left: Settings Frame
        settings_frame = ttk.LabelFrame(self, text="Settings")
        settings_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        settings_frame.grid_columnconfigure(0, weight=1)

        # Sliders for RBW, Max Hold Time, Cycle Wait Time
        ttk.Label(settings_frame, text="Scan RBW (Hz):").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.app_instance.rbw_slider = ttk.Scale(settings_frame, from_=0, to=len(self.app_instance._get_rbw_options()) - 1,
                                    orient="horizontal", command=self.app_instance.update_scan_rbw_from_slider_index_logic)
        self.app_instance.rbw_slider.grid(row=1, column=0, padx=5, pady=2, sticky="ew")
        self.app_instance.rbw_value_label = ttk.Label(settings_frame, textvariable=self.app_instance.desired_rbw_var)
        self.app_instance.rbw_value_label.grid(row=2, column=0, padx=5, pady=2, sticky="w")

        self.app_instance.vbw_value_label = ttk.Label(settings_frame, text="VBW: N/A")
        self.app_instance.vbw_value_label.grid(row=3, column=0, padx=5, pady=2, sticky="w")

        ttk.Label(settings_frame, text="Max Hold Time (s):").grid(row=4, column=0, padx=5, pady=2, sticky="w")
        self.app_instance.max_hold_time_slider = ttk.Scale(settings_frame, from_=0, to=len(self.app_instance._get_max_hold_time_options()) - 1,
                                              orient="horizontal", command=self.app_instance.update_max_hold_time_from_slider_index_logic)
        self.app_instance.max_hold_time_slider.grid(row=5, column=0, padx=5, pady=2, sticky="ew")
        self.app_instance.max_hold_time_label = ttk.Label(settings_frame, textvariable=self.app_instance.desired_max_hold_time_var)
        self.app_instance.max_hold_time_label.grid(row=6, column=0, padx=5, pady=2, sticky="w")

        ttk.Label(settings_frame, text="Cycle Wait Time (s):").grid(row=7, column=0, padx=5, pady=2, sticky="w")
        self.app_instance.cycle_wait_time_slider = ttk.Scale(settings_frame, from_=0, to=len(self.app_instance._get_cycle_wait_time_options()) - 1,
                                                orient="horizontal", command=self.app_instance.update_cycle_wait_time_from_slider_index_logic)
        self.app_instance.cycle_wait_time_slider.grid(row=8, column=0, padx=5, pady=2, sticky="ew")
        self.app_instance.cycle_wait_time_label = ttk.Label(settings_frame, textvariable=self.app_instance.desired_cycle_wait_time_var)
        self.app_instance.cycle_wait_time_label.grid(row=9, column=0, padx=5, pady=2, sticky="w")

        # Other settings (checkboxes, entries)
        ttk.Checkbutton(settings_frame, text="Max Hold Enabled", variable=self.app_instance.desired_maxhold_enabled_var,
                        command=lambda: self.app_instance._on_setting_change(self.app_instance.desired_maxhold_enabled_var, 'last_maxhold_enabled')).grid(row=10, column=0, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(settings_frame, text="High Sensitivity", variable=self.app_instance.desired_high_sensitivity_var,
                        command=lambda: self.app_instance._on_setting_change(self.app_instance.desired_high_sensitivity_var, 'last_high_sensitivity')).grid(row=11, column=0, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(settings_frame, text="Preamplifier ON", variable=self.app_instance.desired_preamp_on_var,
                        command=lambda: self.app_instance._on_setting_change(self.app_instance.desired_preamp_on_var, 'last_preamp_on')).grid(row=12, column=0, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(settings_frame, text="Include Gov Markers", variable=self.app_instance.include_gov_markers_var,
                        command=lambda: self.app_instance._on_setting_change(self.app_instance.include_gov_markers_var, 'last_include_gov_markers')).grid(row=13, column=0, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(settings_frame, text="Include TV Markers", variable=self.app_instance.include_tv_markers_var,
                        command=lambda: self.app_instance._on_setting_change(self.app_instance.include_tv_markers_var, 'last_include_tv_markers')).grid(row=14, column=0, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(settings_frame, text="Open HTML After Complete", variable=self.app_instance.open_html_after_complete_var,
                        command=lambda: self.app_instance._on_setting_change(self.app_instance.open_html_after_complete_var, 'last_open_html_after_complete')).grid(row=15, column=0, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(settings_frame, text="General Debug Enabled", variable=self.app_instance.general_debug_enabled_var,
                        command=self.app_instance._toggle_general_debug).grid(row=16, column=0, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(settings_frame, text="Log VISA Commands", variable=self.app_instance.log_visa_commands_enabled_var,
                        command=self.app_instance._toggle_log_visa_commands).grid(row=17, column=0, padx=5, pady=2, sticky="w")

        # Reference Level Entry
        ttk.Label(settings_frame, text="Reference Level (dBm):").grid(row=18, column=0, padx=5, pady=2, sticky="w")
        self.app_instance.reference_level_entry = ttk.Entry(settings_frame, textvariable=self.app_instance.desired_reference_level_var)
        self.app_instance.reference_level_entry.grid(row=19, column=0, padx=5, pady=2, sticky="ew")
        self.app_instance.reference_level_entry.bind("<FocusOut>", lambda e: self.app_instance._on_setting_change(self.app_instance.desired_reference_level_var, 'last_reference_level_dbm'))

        # Frequency Shift Entry
        ttk.Label(settings_frame, text="Frequency Shift (Hz):").grid(row=20, column=0, padx=5, pady=2, sticky="w")
        self.app_instance.freq_shift_entry = ttk.Entry(settings_frame, textvariable=self.app_instance.desired_freq_shift_var)
        self.app_instance.freq_shift_entry.grid(row=21, column=0, padx=5, pady=2, sticky="ew")
        self.app_instance.freq_shift_entry.bind("<FocusOut>", lambda e: self.app_instance._on_setting_change(self.app_instance.desired_freq_shift_var, 'last_freq_shift_hz'))

        # Scan RBW Segmentation Entry
        ttk.Label(settings_frame, text="Scan RBW Segmentation (Hz):").grid(row=22, column=0, padx=5, pady=2, sticky="w")
        self.app_instance.scan_rbw_segmentation_entry = ttk.Entry(settings_frame, textvariable=self.app_instance.desired_scan_rbw_segmentation_var)
        self.app_instance.scan_rbw_segmentation_entry.grid(row=23, column=0, padx=5, pady=2, sticky="ew")
        self.app_instance.scan_rbw_segmentation_entry.bind("<FocusOut>", lambda e: self.app_instance._on_setting_change(self.app_instance.desired_scan_rbw_segmentation_var, 'last_scan_rbw_segmentation'))

        # Default Focus Width Entry
        ttk.Label(settings_frame, text="Default Focus Width (Hz):").grid(row=24, column=0, padx=5, pady=2, sticky="w")
        self.app_instance.default_focus_width_entry = ttk.Entry(settings_frame, textvariable=self.app_instance.desired_default_focus_width_var)
        self.app_instance.default_focus_width_entry.grid(row=25, column=0, padx=5, pady=2, sticky="ew")
        self.app_instance.default_focus_width_entry.bind("<FocusOut>", lambda e: self.app_instance._on_setting_change(self.app_instance.desired_default_focus_width_var, 'last_default_focus_width'))


        # Restore Defaults Button
        ttk.Button(settings_frame, text="Restore Default Settings", command=lambda: restore_default_settings_logic(self.app_instance)).grid(row=26, column=0, padx=5, pady=10, sticky="ew")

        # Removed redundant "Open Report Folder" button
        # ttk.Button(settings_frame, text="Open Report Folder", command=lambda: open_report_folder_logic(self.app_instance.output_folder_var.get())).grid(row=29, column=0, padx=5, pady=2, sticky="ew")

        # Open Preset Folder Button
        ttk.Button(settings_frame, text="Open Instrument Preset Folder", command=lambda: open_preset_folder_logic()).grid(row=27, column=0, padx=5, pady=2, sticky="ew")


        # Bottom Right: Frequency Band Selection Frame
        band_selection_frame = ttk.LabelFrame(self, text="Frequency Bands to Scan")
        band_selection_frame.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
        band_selection_frame.grid_columnconfigure(0, weight=1)

        # Use a canvas with a scrollbar for the band checkboxes
        band_canvas = tk.Canvas(band_selection_frame, borderwidth=0, highlightthickness=0, bg='#333333')
        band_canvas.grid(row=0, column=0, sticky="nsew")
        band_selection_frame.grid_rowconfigure(0, weight=1)

        band_scrollbar = ttk.Scrollbar(band_selection_frame, orient="vertical", command=band_canvas.yview)
        band_scrollbar.grid(row=0, column=1, sticky="ns")
        band_canvas.config(yscrollcommand=band_scrollbar.set)

        self.app_instance.inner_band_frame = ttk.Frame(band_canvas, style='Dark.TFrame')
        band_canvas.create_window((0, 0), window=self.app_instance.inner_band_frame, anchor="nw")

        self.app_instance.inner_band_frame.bind("<Configure>", lambda e: band_canvas.config(scrollregion=band_canvas.bbox("all")))
        band_canvas.bind('<Enter>', self.app_instance._bind_band_mouse_wheel)
        band_canvas.bind('<Leave>', self.app_instance._unbind_band_mouse_wheel)

        # Populate band checkboxes
        for i, band_item in enumerate(self.app_instance.band_vars):
            band_name = band_item["band"]["Band Name"]
            start_freq = band_item["band"]["Start MHz"]
            stop_freq = band_item["band"]["Stop MHz"]
            # Display band name with start and stop frequencies
            display_text = f"{band_name} ({start_freq:.3f} - {stop_freq:.3f} MHz)"
            var = band_item["var"]
            chk = ttk.Checkbutton(self.app_instance.inner_band_frame, text=display_text, variable=var,
                                  command=self.app_instance._on_band_checkbox_change)
            chk.grid(row=i, column=0, sticky="w", padx=5, pady=1)


class App(tk.Tk):
    """
    The main application class for the RF Spectrum Analyzer Controller.
    This class initializes the GUI, manages instrument connection,
    scan parameters, and data visualization.
    """

    # FIX: Define CONFIG_FILE using os.path.join to ensure it's next to main_app.py
    script_dir = os.path.dirname(__file__)
    CONFIG_FILE = os.path.join(script_dir, 'config.ini')

    def __init__(self):
        """
        Initializes the main application window and components.
        """
        super().__init__()
        self.title("RF Spectrum Analyzer Controller")
        
        # --- FIX: Handle icon loading gracefully ---
        try:
            # Construct the absolute path to the icon
            icon_path = os.path.join(App.script_dir, 'assets', 'app_icon.ico')
            if os.path.exists(icon_path):
                self.iconbitmap(default=icon_path) # Set application icon using absolute path
            else:
                print(f"⚠️ Warning: Icon file not found at {icon_path}. Skipping icon setting.")
        except TclError as e:
            print(f"❌ Error setting application icon: {e}. Ensure 'assets/app_icon.ico' is a valid .ico file and path is correct.")
        except Exception as e:
            print(f"❌ An unexpected error occurred while setting icon: {e}")
        # --- END FIX ---

        self.protocol("WM_DELETE_WINDOW", self._on_closing) # Handle window close event

        # --- Configuration Management ---
        self.config = configparser.ConfigParser()
        self.SCAN_BAND_RANGES = SCAN_BAND_RANGES # Make SCAN_BAND_RANGES accessible
        self.MHZ_TO_HZ = MHZ_TO_HZ # Make MHZ_TO_HZ accessible
        self.VBW_RBW_RATIO = VBW_RBW_RATIO # Make VBW_RBW_RATIO accessible

        # Tkinter variables for settings (initialized with defaults from config)
        self.setting_var_map = {} # To map config keys to Tkinter vars for easy saving/loading
        self._initialize_setting_vars() # Initialize Tkinter variables

        # Load configuration (this will populate Tkinter vars if last_used settings exist)
        load_config(self)

        # --- Instrument Connection Variables ---
        self.rm = None # PyVISA Resource Manager
        self.inst = None # Connected instrument instance
        self.instrument_model = None # Stores the detected instrument model (e.g., N9342CN)
        self.instrument_list = [] # List of available VISA resources
        self.resource_var = tk.StringVar(self) # Currently selected VISA resource
        self.gpib_device_var = tk.StringVar(self) # Stores the last connected GPIB device

        # --- Scan Control Variables ---
        self.scanning = False # Flag to indicate if a scan is in progress
        self.paused = False # Flag to indicate if a scan is paused
        self.stop_event = threading.Event() # Event to signal scan termination
        self.pause_event = threading.Event() # Event to signal scan pause/resume

        # --- Scan Data Storage ---
        self.collected_scans_dataframes = [] # List to store pandas DataFrames for each scan cycle
        self.current_scan_cycle_count = 0 # Counter for scan cycles

        # --- Create GUI Widgets ---
        self._create_widgets()

        # FIX: Set resource_var from last_gpib_device_var AFTER _create_widgets
        # and BEFORE populate_resources_logic to ensure dropdown is correctly pre-selected.
        last_gpib = self.gpib_device_var.get()
        if last_gpib:
            self.resource_var.set(last_gpib)
            debug_print(f"Set initial VISA resource to: {last_gpib}", file=__file__, function=inspect.currentframe().f_code.co_name)


        # --- Initialize PyVISA Resource Manager ---
        try:
            self.rm = pyvisa.ResourceManager()
            populate_resources_logic(self) # Populate resources on startup (dropdown options)
        except Exception as e:
            messagebox.showerror("PyVISA Error", f"Failed to initialize PyVISA Resource Manager: {e}\n"
                                                  "Please ensure PyVISA and a VISA backend (like NI-VISA or Keysight VISA) are installed correctly.")
            print(f"❌ PyVISA Resource Manager initialization failed: {e}")
            self.connect_button.config(state=tk.DISABLED)
            # Initialize blink_id to None if RM fails to init
            self.connect_button_blink_id = None
            self._start_connect_button_blink() # Start blinking if RM fails to init

        # Apply last used window geometry
        self._apply_last_window_geometry()

        # Update VBW display based on initial RBW
        self.update_vbw_display_logic()

        # Set initial colors for settings
        self.reset_setting_colors_logic()

        # Set up console redirection after widgets are created
        self._redirect_console_output()
        print_art() # Print the ASCII art logo to the console

        # Select the last used bands
        self._load_last_selected_bands()

        # FIX: Ensure debug mode is set based on loaded config
        set_debug_mode(self.general_debug_enabled_var.get())
        set_log_visa_commands_mode(self.log_visa_commands_enabled_var.get())
        debug_print(f"Initial Debug Mode: {self.general_debug_enabled_var.get()}", file=__file__, function=inspect.currentframe().f_code.co_name)
        debug_print(f"Initial VISA Log Mode: {self.log_visa_commands_enabled_var.get()}", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _initialize_setting_vars(self):
        """
        Initializes Tkinter StringVar/BooleanVar objects for all configurable settings.
        Maps them to keys for easy access and saving/loading.
        """
        # Instrument Scan Parameters
        self.desired_rbw_var = tk.StringVar(self)
        # --- FIX: Update setting_var_map structure to include (last_key, default_key, tk_var) tuple ---
        self.setting_var_map['desired_rbw_var'] = ('last_scan_rbw_hz', 'default_scan_rbw_hz', self.desired_rbw_var)

        self.desired_cycle_wait_time_var = tk.StringVar(self)
        self.setting_var_map['desired_cycle_wait_time_var'] = ('last_cycle_wait_time_seconds', 'default_cycle_wait_time_seconds', self.desired_cycle_wait_time_var)

        self.desired_max_hold_time_var = tk.StringVar(self)
        self.setting_var_map['desired_max_hold_time_var'] = ('last_maxhold_time_seconds', 'default_maxhold_time_seconds', self.desired_max_hold_time_var)

        self.desired_reference_level_var = tk.StringVar(self)
        self.setting_var_map['desired_reference_level_var'] = ('last_reference_level_dbm', 'default_reference_level_dbm', self.desired_reference_level_var)

        self.desired_freq_shift_var = tk.StringVar(self)
        self.setting_var_map['desired_freq_shift_var'] = ('last_freq_shift_hz', 'default_freq_shift_hz', self.desired_freq_shift_var)

        self.desired_maxhold_enabled_var = tk.BooleanVar(self)
        self.setting_var_map['desired_maxhold_enabled_var'] = ('last_maxhold_enabled', 'default_maxhold_enabled', self.desired_maxhold_enabled_var)

        self.desired_high_sensitivity_var = tk.BooleanVar(self)
        self.setting_var_map['desired_high_sensitivity_var'] = ('last_high_sensitivity', 'default_high_sensitivity', self.desired_high_sensitivity_var)

        self.desired_preamp_on_var = tk.BooleanVar(self)
        self.setting_var_map['desired_preamp_on_var'] = ('last_preamp_on', 'default_preamp_on', self.desired_preamp_on_var)

        self.desired_scan_rbw_segmentation_var = tk.StringVar(self)
        self.setting_var_map['desired_scan_rbw_segmentation_var'] = ('last_scan_rbw_segmentation', 'default_scan_rbw_segmentation', self.desired_scan_rbw_segmentation_var)

        self.desired_default_focus_width_var = tk.StringVar(self)
        self.setting_var_map['desired_default_focus_width_var'] = ('last_default_focus_width', 'default_default_focus_width', self.desired_default_focus_width_var)

        # Plotting and Reporting
        self.include_gov_markers_var = tk.BooleanVar(self)
        self.setting_var_map['include_gov_markers_var'] = ('last_include_gov_markers', 'default_include_gov_markers', self.include_gov_markers_var)

        self.include_tv_markers_var = tk.BooleanVar(self)
        self.setting_var_map['include_tv_markers_var'] = ('last_include_tv_markers', 'default_include_tv_markers', self.include_tv_markers_var)

        self.open_html_after_complete_var = tk.BooleanVar(self)
        self.setting_var_map['open_html_after_complete_var'] = ('last_open_html_after_complete', 'default_open_html_after_complete', self.open_html_after_complete_var)

        # General Application Settings
        self.general_debug_enabled_var = tk.BooleanVar(self)
        self.setting_var_map['general_debug_enabled_var'] = ('last_general_debug_enabled', 'default_general_debug_enabled', self.general_debug_enabled_var)

        self.log_visa_commands_enabled_var = tk.BooleanVar(self)
        self.setting_var_map['log_visa_commands_enabled_var'] = ('last_log_visa_commands_enabled', 'default_log_visa_commands_enabled', self.log_visa_commands_enabled_var)

        self.scan_directory_var = tk.StringVar(self)
        self.setting_var_map['scan_directory_var'] = ('last_scan_directory', 'default_scan_directory', self.scan_directory_var)

        self.scan_name_var = tk.StringVar(self)
        self.setting_var_map['scan_name_var'] = ('last_scan_name', 'default_scan_name', self.scan_name_var)

        self.output_folder_var = tk.StringVar(self) # For output folder path
        # Map to the same config key, but ensure it's a separate entry in the map if needed
        self.setting_var_map['output_folder_var'] = ('last_scan_directory', 'default_scan_directory', self.output_folder_var)

        self.gpib_device_var = tk.StringVar(self) # Last connected GPIB device
        self.setting_var_map['gpib_device_var'] = ('last_gpib_device', None, self.gpib_device_var) # No default_key for gpib_device

        # --- END FIX ---

        # Special handling for selected bands (list of dicts)
        self.selected_bands_str_var = tk.StringVar(self) # Stores comma-separated string of selected bands
        # This one is handled separately in load_config/save_config, not directly in setting_var_map

        # Initialize band selection variables
        self.band_vars = []
        for band in self.SCAN_BAND_RANGES:
            var = tk.BooleanVar(self, value=True) # All bands initially selected
            self.band_vars.append({"band": band, "var": var})

    def _load_last_selected_bands(self):
        """
        Loads the last selected bands from config.ini and updates the checkboxes.
        """
        last_selected_bands_str = self.config.get('LAST_USED_SETTINGS', 'last_selected_bands', fallback='')
        if last_selected_bands_str:
            selected_band_names = [name.strip() for name in last_selected_bands_str.split(',') if name.strip()]
            for band_item in self.band_vars:
                band_item["var"].set(band_item["band"]["Band Name"] in selected_band_names)
            debug_print(f"Loaded last selected bands: {selected_band_names}")
        else:
            debug_print("No last selected bands found in config.ini. All bands remain selected by default.")

    def _apply_last_window_geometry(self):
        """
        Applies the last used window geometry from config.ini.
        """
        last_geometry = self.config.get('LAST_USED_SETTINGS', 'last_window_geometry', fallback=None)
        if last_geometry:
            try:
                self.geometry(last_geometry)
                debug_print(f"Applied last window geometry: {last_geometry}")
            except TclError as e:
                debug_print(f"Error applying last window geometry '{last_geometry}': {e}. Using default.", file=__file__, function=inspect.currentframe().f_code.co_name)
                # Fallback to default if there's an issue with the saved geometry
                default_geometry = self.config.get('DEFAULT_SETTINGS', 'default_window_geometry', fallback='1400x780+100+100')
                self.geometry(default_geometry)
        else:
            debug_print("No last window geometry found. Using default.", file=__file__, function=inspect.currentframe().f_code.co_name)
            default_geometry = self.config.get('DEFAULT_SETTINGS', 'default_window_geometry', fallback='1400x780+100+100')
            self.geometry(default_geometry)


    def _create_widgets(self):
        """
        Creates and arranges all GUI widgets.
        """
        # --- Main Frame (to hold everything) ---
        self.main_frame = ttk.Frame(self, padding="10 10 10 10")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Configure main_frame grid for responsiveness: 1 row, 2 columns
        self.main_frame.grid_rowconfigure(0, weight=1) # Main content row
        self.main_frame.grid_columnconfigure(0, weight=1) # Left half for Notebook
        self.main_frame.grid_columnconfigure(1, weight=1) # Right half for Console


        # --- Main Notebook (Tabs for Scan Configuration, Markers, Report Converter, etc.) ---
        self.notebook = ttk.Notebook(self.main_frame)
        # Place notebook in the left column of the main_frame
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=0)
        
        # Add Scan Configuration Tab as the first tab
        self.scan_config_tab = ScanConfigurationTab(self.notebook, app_instance=self)
        self.notebook.add(self.scan_config_tab, text="Scan Configuration")

        # Markers Display Tab
        self.markers_display_tab = MarkersDisplayTab(self.notebook, app_instance=self)
        self.notebook.add(self.markers_display_tab, text="Markers Display")

        # Report Converter Tab
        self.report_converter_frame = ReportConverterTab(self.notebook, app_instance=self)
        self.notebook.add(self.report_converter_frame, text="Report Converter")

        # Preset Files Tab
        self.preset_files_tab = PresetFilesTab(self.notebook, app_instance=self)
        self.notebook.add(self.preset_files_tab, text="Instrument Presets")
        
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)


        # --- Console Output Area ---
        # Create console_text directly and place it in the right column of the main_frame
        self.console_text = scrolledtext.ScrolledText(self.main_frame, wrap="word", height=10, bg="#2b2b2b", fg="#cccccc", insertbackground="white")
        self.console_text.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=0)

        # Set initial slider positions based on loaded config values
        # This call needs to be after all widgets (including those in ScanConfigurationTab) are created.
        # It's currently at the end of App.__init__, which is correct.
        # self._set_initial_slider_positions() # This is called in __init__ after _create_widgets()


    def _set_initial_slider_positions(self):
        """
        Sets the initial positions of the sliders based on the loaded configuration.
        This ensures the GUI reflects the `last_used_settings` on startup.
        """
        # RBW Slider
        rbw_options = self._get_rbw_options()
        current_rbw = float(self.desired_rbw_var.get())
        try:
            # Find the index of the current RBW value in the options list
            index = rbw_options.index(current_rbw)
            self.rbw_slider.set(index)
        except ValueError:
            # If current RBW is not exactly in options, find closest or default
            closest_index = min(range(len(rbw_options)), key=lambda i: abs(rbw_options[i] - current_rbw))
            self.rbw_slider.set(closest_index)
            self.desired_rbw_var.set(str(rbw_options[closest_index])) # Update var to match slider

        # Max Hold Time Slider
        max_hold_options = self._get_max_hold_time_options()
        current_max_hold_time = float(self.desired_max_hold_time_var.get())
        try:
            index = max_hold_options.index(current_max_hold_time)
            self.max_hold_time_slider.set(index)
        except ValueError:
            closest_index = min(range(len(max_hold_options)), key=lambda i: abs(max_hold_options[i] - current_max_hold_time))
            self.max_hold_time_slider.set(closest_index)
            self.desired_max_hold_time_var.set(str(max_hold_options[closest_index]))

        # Cycle Wait Time Slider
        cycle_wait_options = self._get_cycle_wait_time_options()
        current_cycle_wait_time = float(self.desired_cycle_wait_time_var.get())
        try:
            index = cycle_wait_options.index(current_cycle_wait_time)
            self.cycle_wait_time_slider.set(index)
        except ValueError:
            closest_index = min(range(len(cycle_wait_options)), key=lambda i: abs(cycle_wait_options[i] - current_cycle_wait_time))
            self.cycle_wait_time_slider.set(closest_index)
            self.desired_cycle_wait_time_var.set(str(cycle_wait_options[closest_index]))


    def _get_rbw_options(self):
        """Returns a list of valid RBW options for the slider."""
        return [10, 30, 100, 300, 1000, 3000, 10000, 30000, 100000, 300000, 1000000, 3000000, 5000000]

    def _get_max_hold_time_options(self):
        """Returns a list of valid Max Hold Time options for the slider."""
        return [1, 2, 3, 5, 10, 15, 20, 30, 45, 60, 90, 120, 180, 240, 300, 600, 900, 1200, 1800, 2400, 3000, 3600] # Up to 1 hour

    def _get_cycle_wait_time_options(self):
        """Returns a list of valid Cycle Wait Time options for the slider."""
        return [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0]

    def update_scan_rbw_from_slider_index_logic(self, val):
        """Updates the RBW StringVar based on slider position."""
        index = int(float(val))
        options = self._get_rbw_options()
        if 0 <= index < len(options):
            self.desired_rbw_var.set(str(options[index]))
            self._on_setting_change(self.desired_rbw_var, 'last_scan_rbw_hz')
            self.update_vbw_display_logic() # Update VBW when RBW changes
        else:
            debug_print(f"RBW slider index {index} out of bounds.", file=__file__, function=inspect.currentframe().f_code.co_name)

    def update_max_hold_time_from_slider_index_logic(self, val):
        """Updates the Max Hold Time StringVar based on slider position."""
        index = int(float(val))
        options = self._get_max_hold_time_options()
        if 0 <= index < len(options):
            self.desired_max_hold_time_var.set(str(options[index]))
            self._on_setting_change(self.desired_max_hold_time_var, 'last_maxhold_time_seconds')
        else:
            debug_print(f"Max Hold Time slider index {index} out of bounds.", file=__file__, function=inspect.currentframe().f_code.co_name)

    def update_cycle_wait_time_from_slider_index_logic(self, val):
        """Updates the Cycle Wait Time StringVar based on slider position."""
        index = int(float(val))
        options = self._get_cycle_wait_time_options()
        if 0 <= index < len(options):
            self.desired_cycle_wait_time_var.set(str(options[index]))
            self._on_setting_change(self.desired_cycle_wait_time_var, 'last_cycle_wait_time_seconds')
        else:
            debug_print(f"Cycle Wait Time slider index {index} out of bounds.", file=__file__, function=inspect.currentframe().f_code.co_name)

    def update_vbw_display_logic(self, *args):
        """
        Updates the displayed VBW value based on the current RBW and VBW_RBW_RATIO.
        Called when RBW slider changes or on initial load.
        """
        try:
            current_rbw = float(self.desired_rbw_var.get())
            calculated_vbw = current_rbw * self.VBW_RBW_RATIO
            self.vbw_value_label.config(text=f"VBW: {calculated_vbw:.0f} Hz")
        except ValueError:
            self.vbw_value_label.config(text="VBW: Invalid RBW")
            debug_print("Could not calculate VBW: Invalid RBW value.", file=__file__, function=inspect.currentframe().f_code.co_name)

    def _on_setting_change(self, tk_var, config_key):
        """
        Callback for when a setting Tkinter variable changes.
        Marks the setting as 'changed' by changing its background color.
        """
        # Get the widget associated with the Tkinter variable
        widget = None
        for key, var in self.setting_var_map.items():
            # The value of var is now a tuple (last_key, default_key, tk_var_instance)
            # We need to compare tk_var_instance with the passed tk_var
            if len(var) == 3 and var[2] == tk_var: # Check if it's the correct tuple and matches tk_var
                # This is a bit indirect, but we need to find the actual widget
                # that uses this tk_var. For entries, it's straightforward.
                # For checkboxes, it's the checkbox itself.
                # This might require more robust mapping if widgets are not directly named.
                # For simplicity, let's assume direct mapping for now or rely on default styling.
                pass # We'll rely on global style changes for now

        # We'll rely on the reset_setting_colors_logic to clear colors on apply/restore
        # For now, just print a debug message
        debug_print(f"Setting '{config_key}' changed to '{tk_var.get()}'.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def reset_setting_colors_logic(self):
        """
        Resets the background color of all setting input widgets to default,
        indicating that settings are either applied or at their default state.
        This function iterates through all relevant widgets and sets their style.
        """
        # Get the default style for Entry and Checkbutton
        default_entry_style_name = 'TEntry'
        default_checkbutton_style_name = 'TCheckbutton'
        default_scale_style_name = 'TScale'

        # Apply default style to Entry widgets
        self.reference_level_entry.config(style=default_entry_style_name)
        self.freq_shift_entry.config(style=default_entry_style_name)
        self.scan_name_entry.config(style=default_entry_style_name)
        self.output_folder_entry.config(style=default_entry_style_name)
        self.scan_rbw_segmentation_entry.config(style=default_entry_style_name)
        self.default_focus_width_entry.config(style=default_entry_style_name)

        # Apply default style to Checkbutton widgets
        # This is more complex as we don't have direct references to all of them
        # A more robust solution would involve storing references to all checkbuttons
        # in a list or dictionary during their creation.
        # For now, we'll assume a global style applies or skip individual checkbuttons.
        # If you need individual styling, you'd store them like:
        # self.maxhold_chk = ttk.Checkbutton(...)
        # self.maxhold_chk.config(style=default_checkbutton_style_name)
        # For now, rely on global theme.

        # Apply default style to Scale widgets
        self.rbw_slider.config(style=default_scale_style_name)
        self.max_hold_time_slider.config(style=default_scale_style_name)
        self.cycle_wait_time_slider.config(style=default_scale_style_name)
        
        debug_print("Setting colors reset to default.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _toggle_general_debug(self):
        """Toggles the global debug mode and updates the setting."""
        new_state = self.general_debug_enabled_var.get()
        set_debug_mode(new_state)
        # The _on_setting_change expects a tk_var and a config_key.
        # It's better to pass the actual tk_var instance directly.
        # Also, the config_key should be the 'last_key' string, not the tk_var instance.
        self._on_setting_change(self.general_debug_enabled_var, 'last_general_debug_enabled')

    def _toggle_log_visa_commands(self):
        """Toggles the global VISA command logging mode and updates the setting."""
        new_state = self.log_visa_commands_enabled_var.get()
        set_log_visa_commands_mode(new_state)
        self._on_setting_change(self.log_visa_commands_enabled_var, 'last_log_visa_commands_enabled')


    def _browse_output_folder(self):
        """
        Opens a directory dialog to select the output folder for scan data.
        """
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder_var.set(folder_selected)
            debug_print(f"Output folder set to: {folder_selected}", file=__file__, function=inspect.currentframe().f_code.co_name)
            self._on_setting_change(self.output_folder_var, 'last_scan_directory')


    def _redirect_console_output(self):
        """
        Redirects stdout and stderr to the console_text widget.
        """
        sys.stdout = TextRedirector(self.console_text, "stdout")
        sys.stderr = TextRedirector(self.console_text, "stderr")
        print("Console output redirected to GUI.")

    def _update_console_line(self, message):
        """
        Updates the console with a new message, always on a new line.
        This function no longer handles overwriting; TextRedirector does that if \r is present.
        """
        if self.console_text:
            # Simply write the message. TextRedirector's write method will handle
            # appending it and adding a newline.
            sys.stdout.write(message)


    def _start_connect_button_blink(self):
        """Starts the blinking effect for the connect button."""
        # Initialize connect_button_blink_id if it's not set
        if not hasattr(self, 'connect_button_blink_id'):
            self.connect_button_blink_id = None

        if self.connect_button_blink_id is None:
            # Directly set the style to start blinking, then schedule the toggle
            self.connect_button.config(style='Red.TButton') # Start with red
            self.connect_button_blink_id = self.after(500, self._blink_connect_button)

    def _stop_connect_button_blink(self):
        """Stops the blinking effect for the connect button."""
        if hasattr(self, 'connect_button_blink_id') and self.connect_button_blink_id:
            self.after_cancel(self.connect_button_blink_id)
            self.connect_button_blink_id = None
            # Ensure button is in its normal state/color
            self.connect_button.config(style='GreyText.TButton') # Reset to default style

    def _blink_connect_button(self):
        """Toggles the connect button color for blinking."""
        current_style = self.connect_button.cget("style")
        if current_style == 'Red.TButton':
            self.connect_button.config(style='GreyText.TButton')
        else:
            self.connect_button.config(style='Red.TButton')
        self.connect_button_blink_id = self.after(500, self._blink_connect_button)


    def _toggle_pause_button_color(self):
        """Toggles the pause/resume button color for blinking."""
        current_style = self.pause_resume_button.cget("style")
        if current_style == 'Red.TButton':
            self.pause_resume_button.config(style='Orange.TButton')
        else:
            self.pause_resume_button.config(style='Red.TButton')


    def _start_pause_button_blink(self):
        """Starts the blinking effect for the pause/resume button."""
        if self.pause_blink_id is None:
            self._toggle_pause_button_color() # Initial toggle
            self.pause_blink_id = self.after(500, self._blink_pause_button)

    def _stop_pause_button_blink(self):
        """Stops the blinking effect for the pause/resume button."""
        if self.pause_blink_id:
            self.after_cancel(self.pause_blink_id)
            self.pause_blink_id = None
            self.pause_resume_button.config(style='Orange.TButton') # Reset to default style

    def _blink_pause_button(self):
        """Toggles the pause/resume button color for blinking."""
        current_style = self.pause_resume_button.cget("style")
        if current_style == 'Red.TButton':
            self.pause_resume_button.config(style='Orange.TButton')
        else:
            self.pause_resume_button.config(style='Red.TButton')
        self.pause_blink_id = self.after(500, self._blink_pause_button)


    def _on_tab_changed(self, event):
        """
        Callback for when the notebook tab changes.
        Triggers the _on_tab_selected method of the newly selected tab if it exists.
        """
        selected_tab_id = self.notebook.select()
        selected_tab_widget = self.nametowidget(selected_tab_id)
        
        # Call the _on_tab_selected method if the tab has one
        if hasattr(selected_tab_widget, '_on_tab_selected'):
            selected_tab_widget._on_tab_selected(event)

        # Update the state of the load_preset_button based on selected tab
        if hasattr(self, 'preset_files_tab') and selected_tab_widget == self.preset_files_tab:
            if self.inst and self.preset_files_tab.get_selected_preset():
                self.load_preset_button.config(state=tk.NORMAL)
            else:
                self.load_preset_button.config(state=tk.DISABLED)
        else:
            # Ensure load_preset_button is disabled if not on the preset tab or preset_files_tab is not yet initialized
            if hasattr(self, 'load_preset_button'):
                self.load_preset_button.config(state=tk.DISABLED)


    def _bind_band_mouse_wheel(self, event):
        """Binds mouse wheel events for the band selection canvas."""
        event.widget.bind_all("<MouseWheel>", self._on_band_mouse_wheel)
        event.widget.bind_all("<Button-4>", self._on_band_mouse_wheel) # For Linux
        event.widget.bind_all("<Button-5>", self._on_band_mouse_wheel) # For Linux

    def _unbind_band_mouse_wheel(self, event):
        """Unbinds mouse wheel events for the band selection canvas."""
        event.widget.unbind_all("<MouseWheel>")
        event.widget.unbind_all("<Button-4>")
        event.widget.unbind_all("<Button-5>")

    def _on_band_mouse_wheel(self, event):
        """Handles mouse wheel scrolling for the band selection canvas."""
        if sys.platform == "darwin":
            event.widget.yview_scroll(-1 * int(event.delta), "units")
        elif event.num == 4: # Linux scroll up
            event.widget.yview_scroll(-1, "units")
        elif event.num == 5: # Linux scroll down
            event.widget.yview_scroll(1, "units")
        else: # Windows
            event.widget.yview_scroll(-1 * int(event.delta/120), "units")

    def _on_band_checkbox_change(self):
        """
        Callback when a band selection checkbox changes.
        Currently, just logs the change.
        """
        selected_bands = [item["band"]["Band Name"] for item in self.band_vars if item["var"].get()]
        debug_print(f"Selected bands updated: {selected_bands}", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _reset_gui_on_disconnect_or_error(self):
        """
        Resets the GUI state after an instrument disconnection or connection error.
        """
        debug_print("Resetting GUI on disconnect or error...", file=__file__, function=inspect.currentframe().f_code.co_name)
        self.inst = None
        self.instrument_model = None
        self.scanning = False
        self.paused = False
        self.stop_event.clear()
        self.pause_event.clear()

        self.connect_button.config(state=tk.NORMAL)
        self.disconnect_button.config(state=tk.DISABLED)
        self.start_scan_button.config(state=tk.DISABLED)
        self.stop_scan_button.config(state=tk.DISABLED)
        self.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED)
        self.apply_button.config(state=tk.DISABLED)
        self.plot_button.config(state=tk.DISABLED) # Disable plot button on disconnect
        self.load_preset_button.config(state=tk.DISABLED) # Disable load preset button

        # Disable query presets button
        if hasattr(self, 'preset_files_tab') and hasattr(self.preset_files_tab, 'query_presets_button'):
            self.preset_files_tab.query_presets_button.config(state=tk.DISABLED)

        # Initialize connect_button_blink_id if it's not set
        if not hasattr(self, 'connect_button_blink_id'):
            self.connect_button_blink_id = None
        self._start_connect_button_blink()
        self._stop_pause_button_blink() # Ensure pause button stops blinking
        self.reset_setting_colors_logic() # Reset setting colors

        # Clear collected scan data
        self.collected_scans_dataframes = []
        self.current_scan_cycle_count = 0


    def _on_closing(self):
        """
        Handles the window closing event. Disconnects instrument and saves config.
        """
        print("Application closing...")
        if self.scanning:
            print("Stopping active scan before closing...")
            stop_scan_logic(self) # Ensure scan is stopped gracefully
            # Give it a moment to stop
            time.sleep(1) # Small delay to allow thread to terminate

        if self.inst:
            print("Disconnecting instrument before closing...")
            disconnect_instrument_logic(self)

        save_config(self) # Save current settings before exit
        print("Configuration saved. Goodbye!")
        self.destroy() # Destroy the Tkinter window


if __name__ == "__main__":
    # Apply a custom theme
    style = ttk.Style()
    style.theme_use('clam') # 'clam', 'alt', 'default', 'classic'

    # Configure custom styles
    style.configure('TFrame', background='#333333')
    style.configure('TLabel', background='#333333', foreground='white')
    style.configure('TLabelFrame', background='#333333', foreground='white', bordercolor='#555555')
    style.configure('TEntry', fieldbackground='#444444', foreground='white', bordercolor='#555555')
    style.map('TEntry', fieldbackground=[('focus', '#555555')])
    style.configure('TButton', background='#555555', foreground='white', font=('Inter', 10, 'bold'), borderwidth=1, relief="raised")
    style.map('TButton',
              background=[('active', '!disabled', '#777777'), ('pressed', '#333333')],
              foreground=[('disabled', '#888888')])

    # Specific button styles
    style.configure('Green.TButton', background='#4CAF50', foreground='white')
    style.map('Green.TButton', background=[('active', '!disabled', '#66BB6A'), ('pressed', '#388E3C')])

    style.configure('Red.TButton', background='#F44336', foreground='white')
    style.map('Red.TButton', background=[('active', '!disabled', '#E57373'), ('pressed', '#D32F2F')])

    style.configure('Orange.TButton', background='#FF9800', foreground='white')
    style.map('Orange.TButton', background=[('active', '!disabled', '#FFB74D'), ('pressed', '#F57C00')])

    style.configure('Blue.TButton', background='#2196F3', foreground='white')
    style.map('Blue.TButton', background=[('active', '!disabled', '#64B5F6'), ('pressed', '#1976D2')])

    style.configure('Accent.TButton', background='#00BCD4', foreground='white') # Cyan-like accent
    style.map('Accent.TButton', background=[('active', '!disabled', '#4DD0E1'), ('pressed', '#0097A7')])

    style.configure('GreyText.TButton', background='#607D8B', foreground='white') # Greyish for general buttons
    style.map('GreyText.TButton', background=[('active', '!disabled', '#90A4AE'), ('pressed', '#455A64')])

    # Style for canvas background within themed frames
    style.configure('Dark.TFrame', background='#333333')

    # MarkersDisplayTab related styles
    style.configure("Markers.TFrame", background="#000000") # Dark background for the main frame
    style.configure("Markers.TLabel", background="#000000", foreground="white")
    style.configure("Markers.Treeview.Heading", background="#3a3a3a", foreground="white")
    style.configure("Markers.Treeview", background="#4a4a4a", foreground="white", fieldbackground="#4a4a4a")
    style.map("Markers.Treeview", background=[("selected", "#0078D7")], foreground=[("selected", "white")])
    
    style.configure("Markers.TButton", 
                    background="#555555", # Darker grey for span buttons (default unselected)
                    foreground="white",   # White text
                    font=("Helvetica", 14, "bold"), # Larger font
                    padding=[15, 15, 15, 15]) # More padding
    style.map("Markers.TButton", 
              background=[('active', '#777777')]) # Lighter grey on active for unselected

    style.configure("Markers.Inner.Treeview",
                    background="#333333", # Darker grey
                    foreground="white",
                    fieldbackground="#333333", # Darker grey
                    bordercolor="black",
                    lightcolor="#333333", # Darker grey
                    darkcolor="#333333") # Darker grey
    style.map("Markers.Inner.Treeview",
              background=[("selected", "#0078D7")], # Reverted to blue highlight for treeview
              foreground=[("selected", "white")])
    
    # Configure the base TLabelFrame style and its label part
    style.configure("TLabelFrame", background="#333333", foreground="white")
    style.configure("TLabelFrame.Label", background="#333333", foreground="white")
    
    # The problematic line: This was creating a *new* local style object
    # style = ttk.Style() # REMOVED THIS REDUNDANT LINE

    # Add the new style for selected preset buttons (already correct)
    style.configure("LargePreset.TButton",
                    background="#555555", # Darker grey for buttons
                    foreground="white",
                    font=("Helvetica", 40, "bold"), # Set font size to 40
                    padding=[30, 15, 30, 15]) # Adjust padding as needed
    style.map("LargePreset.TButton",
            background=[('active', '#777777')]) # Lighter grey on active

    style.configure("SelectedPreset.TButton",
                    background="#2196F3", # A nice blue color
                    foreground="white",
                    font=("Helvetica", 40, "bold"), # Keep the 40-point font
                    padding=[30, 15, 30, 15])
    style.map("SelectedPreset.TButton",
              background=[('active', '#64B5F6'), ('pressed', '#1976D2')])
    
    # Corrected style for selected span buttons
    style.configure("SelectedSpan.TButton",
                        background="#F4902C", # Orange background
                        foreground="black",
                        font=("Helvetica", 14, "bold"), # Or whatever font you prefer
                        padding=[15, 15, 15, 15])
    style.map("SelectedSpan.TButton",
                background=[('active', '#FFB050')]) # Lighter orange on active

    app = App()
    app.mainloop()
