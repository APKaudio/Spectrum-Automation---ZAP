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
    apply_settings_logic,
    # Removed set_focus_frequency_logic as it's no longer in instrument_logic.py
    query_current_instrument_settings_logic, query_device_presets_logic,
    load_selected_preset_logic, set_marker_and_trace_modes_logic
)
from src.scan_logic import start_scan_thread_logic, stop_scan_logic, pause_scan_logic, resume_scan_logic, _update_button_states_on_connection
from src.settings_logic import restore_default_settings_logic
from src.instrument_preset_tab import PresetFilesTab # Import the new tab class
from src.marker_logic import MarkersDisplayTab # Import the MarkersDisplayTab
from src.plotting_tab import PlottingTab # Import the PlottingTab
from src.report_converter_tab import ReportConverterTab # Import the ReportConverterTab
from src.scan_tab import ScanTab # Import the new ScanTab

# Import constants from frequency_bands.py
from utils.frequency_bands import SCAN_BAND_RANGES, MHZ_TO_HZ, VBW_RBW_RATIO
from utils.instrument_control import set_debug_mode, set_log_visa_commands_mode, debug_print # Import debug_print

# Global flag for debug mode (can be controlled by GUI checkbox)
# DEBUG_MODE = False # Now controlled by instrument_control.py's global variable


class App(tk.Tk):
    """
    Main application class for the RF Spectrum Analyzer Controller.
    This class inherits from Tkinter's Tk class to create the main window
    and manage the overall application flow, including GUI setup, instrument
    communication, and data processing.
    """
    CONFIG_FILE = 'config.ini'
    DEFAULT_WINDOW_GEOMETRY = "1400x780+100+100" # Default size and position

    def __init__(self):
        """
        Initializes the main application window and sets up core components.
        """
        super().__init__()
        self.title("RF Spectrum Analyzer Controller")
        self.protocol("WM_DELETE_WINDOW", self._on_closing) # Handle window close event
        
        self.config = configparser.ConfigParser()
        self.config.read(self.CONFIG_FILE) # Read config early

        # Ensure necessary Python packages are installed
        self._check_and_install_dependencies()

        # Initialize instrument and connection variables
        self.rm = None # Resource Manager
        self.inst = None # Instrument instance
        self.instrument_model = None # To store the instrument model (e.g., N9340B)

        # Store collected scan dataframes for plotting and analysis
        self.collected_scans_dataframes = []

        # Scan state variables
        self.scanning = False
        self.scan_thread = None
        self.stop_scan_event = threading.Event()
        self.pause_scan_event = threading.Event()

        # Define SCAN_BAND_RANGES from frequency_bands.py
        self.SCAN_BAND_RANGES = SCAN_BAND_RANGES
        self.MHZ_TO_HZ = MHZ_TO_HZ
        self.VBW_RBW_RATIO = VBW_RBW_RATIO

        # Tkinter variables for settings (linked to config.ini)
        self._setup_tkinter_vars()

        # Load configuration and apply geometry
        load_config(self) # Pass self to load_config
        self._apply_saved_geometry() # Apply geometry after loading config

        # Set initial debug mode based on config
        set_debug_mode(self.general_debug_enabled_var.get())
        set_log_visa_commands_mode(self.log_visa_commands_enabled_var.get())

        # Set up GUI elements
        self._create_widgets()
        self._setup_styles() # Apply custom styles
        self._redirect_stdout_to_console() # Redirect print statements to GUI console

        # Initial population of resources and button states
        self._populate_resources()
        # Corrected: Call _update_button_states_on_connection as a function, passing self
        _update_button_states_on_connection(self)
        
        # Print the ASCII art logo to the console
        print_art()


    def _check_and_install_dependencies(self):
        """
        Checks for necessary Python packages (pyvisa, pandas, beautifulsoup4, pdfplumber)
        and attempts to install them if missing.
        """
        dependencies = {
            "pyvisa": "pyvisa",
            "pandas": "pandas",
            "bs4": "beautifulsoup4", # For BeautifulSoup
            "pdfplumber": "pdfplumber"
        }
        missing_dependencies = []

        for import_name, package_name in dependencies.items():
            try:
                __import__(import_name)
            except ImportError:
                missing_dependencies.append(package_name)

        if missing_dependencies:
            missing_str = ", ".join(missing_dependencies)
            response = messagebox.askyesno(
                "Missing Dependencies",
                f"The following Python packages are missing: {missing_str}.\n"
                "Do you want to install them now? This may take a few moments."
            )
            if response:
                try:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", *missing_dependencies])
                    messagebox.showinfo("Installation Complete", "Required packages installed successfully. Please restart the application.")
                    sys.exit(0) # Exit to restart application
                except Exception as e:
                    messagebox.showerror("Installation Failed", f"Failed to install packages: {e}\nPlease install them manually using 'pip install {missing_str}'")
                    sys.exit(1) # Exit if user chooses not to install
            else:
                messagebox.showwarning("Dependencies Missing", "Application may not function correctly without required packages.")
                sys.exit(1) # Exit if user chooses not to install


    def _setup_tkinter_vars(self):
        """
        Initializes Tkinter variables for all application settings,
        mapping them to their corresponding keys in config.ini.
        """
        # Session Description Variables (Remain in App)
        self.scan_name_var = tk.StringVar(self, value="MyScan")
        self.output_folder_var = tk.StringVar(self, value="scan_data")

        # Debugging variables (Remain in App, but GUI elements moved)
        self.general_debug_enabled_var = tk.BooleanVar(self, value=False)
        self.log_visa_commands_enabled_var = tk.BooleanVar(self, value=False)

        # Plotting variables (Remain in App, but GUI elements moved)
        self.include_gov_markers_var = tk.BooleanVar(self, value=True)
        self.include_tv_markers_var = tk.BooleanVar(self, value=True)
        self.include_markers_var = tk.BooleanVar(self, value=True) # For custom markers from CSV
        self.open_html_after_complete_var = tk.BooleanVar(self, value=True)

        # Scan Configuration Variables (Moved to ScanTab, but vars remain in App for global access)
        self.rbw_step_size_hz_var = tk.StringVar(self, value="1000000")
        self.cycle_wait_time_seconds_var = tk.StringVar(self, value="0.5")
        self.maxhold_time_seconds_var = tk.StringVar(self, value="3")
        self.scan_rbw_hz_var = tk.StringVar(self, value="10000")
        self.reference_level_dbm_var = tk.StringVar(self, value="-40")
        self.freq_shift_hz_var = tk.StringVar(self, value="0")
        self.maxhold_enabled_var = tk.BooleanVar(self, value=True)
        self.high_sensitivity_var = tk.BooleanVar(self, value=True)
        self.preamp_on_var = tk.BooleanVar(self, value=True)
        self.scan_rbw_segmentation_var = tk.StringVar(self, value="1000000.0")
        self.desired_default_focus_width_var = tk.StringVar(self, value="10000.0") # In MHz
        self.num_scan_cycles_var = tk.IntVar(self, value=1)


        # Map Tkinter variables to config.ini keys for easy loading/saving
        self.setting_var_map = {
            'rbw_step_size_hz_var': ('last_rbw_step_size_hz', 'default_rbw_step_size_hz', self.rbw_step_size_hz_var),
            'cycle_wait_time_seconds_var': ('last_cycle_wait_time_seconds', 'default_cycle_wait_time_seconds', self.cycle_wait_time_seconds_var),
            'maxhold_time_seconds_var': ('last_maxhold_time_seconds', 'default_maxhold_time_seconds', self.maxhold_time_seconds_var),
            'scan_rbw_hz_var': ('last_scan_rbw_hz', 'default_scan_rbw_hz', self.scan_rbw_hz_var),
            'reference_level_dbm_var': ('last_reference_level_dbm', 'default_reference_level_dbm', self.reference_level_dbm_var),
            'freq_shift_hz_var': ('last_freq_shift_hz', 'default_freq_shift_hz', self.freq_shift_hz_var),
            'maxhold_enabled_var': ('last_maxhold_enabled', 'default_maxhold_enabled', self.maxhold_enabled_var),
            'high_sensitivity_var': ('last_high_sensitivity', 'default_high_sensitivity', self.high_sensitivity_var),
            'preamp_on_var': ('last_preamp_on', 'default_preamp_on', self.preamp_on_var),
            'scan_rbw_segmentation_var': ('last_scan_rbw_segmentation', 'default_scan_rbw_segmentation', self.scan_rbw_segmentation_var),
            'desired_default_focus_width_var': ('last_default_focus_width', 'default_default_focus_width', self.desired_default_focus_width_var),
            'include_gov_markers_var': ('last_include_gov_markers', 'default_include_gov_markers', self.include_gov_markers_var),
            'include_tv_markers_var': ('last_include_tv_markers', 'default_include_tv_markers', self.include_tv_markers_var),
            'include_markers_var': ('last_include_markers', 'default_include_markers', self.include_markers_var),
            'open_html_after_complete_var': ('last_open_html_after_complete', 'default_open_html_after_complete', self.open_html_after_complete_var),
            'general_debug_enabled_var': ('last_general_debug_enabled', 'default_general_debug_enabled', self.general_debug_enabled_var),
            'log_visa_commands_enabled_var': ('last_log_visa_commands_enabled', 'default_log_visa_commands_enabled', self.log_visa_commands_enabled_var),
            'scan_name_var': ('last_scan_name', 'default_scan_name', self.scan_name_var),
            'output_folder_var': ('last_scan_directory', 'default_scan_directory', self.output_folder_var),
            'num_scan_cycles_var': ('last_num_scan_cycles', 'default_num_scan_cycles', self.num_scan_cycles_var)
        }

        # Tkinter variables for band selection checkboxes
        self.band_vars = []
        for band in self.SCAN_BAND_RANGES:
            # Each item in band_vars will be a dict: {"band": {...}, "var": tk.BooleanVar}
            self.band_vars.append({"band": band, "var": tk.BooleanVar(self, value=False)})

        # Link debug mode Tkinter var to the global setter function
        self.general_debug_enabled_var.trace_add("write", lambda *args: set_debug_mode(self.general_debug_enabled_var.get()))
        self.log_visa_commands_enabled_var.trace_add("write", lambda *args: set_log_visa_commands_mode(self.log_visa_commands_enabled_var.get()))


    def _apply_saved_geometry(self):
        """
        Applies the window geometry saved in config.ini, or uses a default.
        """
        saved_geometry = self.config.get('LAST_USED_SETTINGS', 'last_window_geometry', fallback=self.DEFAULT_WINDOW_GEOMETRY)
        try:
            self.geometry(saved_geometry)
            debug_print(f"Applied saved geometry: {saved_geometry}", file=__file__, function=inspect.currentframe().f_code.co_name)
        except TclError as e:
            debug_print(f"Invalid saved geometry '{saved_geometry}': {e}. Using default.", file=__file__, function=inspect.currentframe().f_code.co_name)
            self.geometry(self.DEFAULT_WINDOW_GEOMETRY)


    def _create_widgets(self):
        """
        Creates and arranges all GUI widgets in the main application window.
        """
        debug_print("Creating main application widgets...", file=__file__, function=inspect.currentframe().f_code.co_name)
        
        # Configure grid for the main window - 50/50 split, single row
        self.grid_columnconfigure(0, weight=1) # Left column
        self.grid_columnconfigure(1, weight=1) # Right column
        self.grid_rowconfigure(0, weight=1) # Main row for both columns (expands)

        # --- Left Column: Notebook for Settings and Tabs ---
        self.left_notebook = ttk.Notebook(self)
        self.left_notebook.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        # Create frames for the new tabs
        main_settings_tab_frame = ttk.Frame(self.left_notebook, style='Dark.TFrame')
        
        # Add the new frames as tabs to the left notebook
        self.left_notebook.add(main_settings_tab_frame, text="Main Settings")
        
        # Configure grid for main_settings_tab_frame
        main_settings_tab_frame.grid_columnconfigure(0, weight=1)
        main_settings_tab_frame.grid_rowconfigure(0, weight=0) # Session Description
        main_settings_tab_frame.grid_rowconfigure(1, weight=1) # Instrument Connection (expands)


        # --- Session Description Frame (Moved into Main Settings tab) ---
        session_description_frame = ttk.LabelFrame(main_settings_tab_frame, text="Session Description", style='Dark.TLabelframe')
        session_description_frame.grid(row=0, column=0, padx=5, pady=5, sticky="new")
        session_description_frame.grid_columnconfigure(1, weight=1) # Allow entry fields to expand

        ttk.Label(session_description_frame, text="Scan Name:", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.scan_name_entry = ttk.Entry(session_description_frame, textvariable=self.scan_name_var, style='TEntry')
        self.scan_name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(session_description_frame, text="Output Directory:", style='TLabel').grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.output_folder_entry = ttk.Entry(session_description_frame, textvariable=self.output_folder_var, style='TEntry')
        self.output_folder_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.browse_output_button = ttk.Button(session_description_frame, text="Browse", command=self._browse_output_folder, style='Blue.TButton')
        self.browse_output_button.grid(row=1, column=2, padx=5, pady=5)


        # --- Instrument Connection Frame (Moved into Main Settings tab) ---
        instrument_frame = ttk.LabelFrame(main_settings_tab_frame, text="Instrument Connection", style='Dark.TLabelframe')
        instrument_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew") # Now expands vertically
        instrument_frame.grid_columnconfigure(1, weight=1) # Allow combobox to expand

        ttk.Label(instrument_frame, text="VISA Resource:", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.resource_combobox = ttk.Combobox(instrument_frame, state="readonly", style='TCombobox')
        self.resource_combobox.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.resource_combobox.bind("<<ComboboxSelected>>", self._on_resource_selected)

        self.refresh_button = ttk.Button(instrument_frame, text="Refresh", command=self._populate_resources, style='Blue.TButton')
        self.refresh_button.grid(row=0, column=2, padx=5, pady=5)

        self.connect_button = ttk.Button(instrument_frame, text="Connect", command=self._connect_instrument, style='Green.TButton')
        self.connect_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        self.disconnect_button = ttk.Button(instrument_frame, text="Disconnect", command=self._disconnect_instrument, state=tk.DISABLED, style='Red.TButton')
        self.disconnect_button.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.apply_button = ttk.Button(instrument_frame, text="Apply Settings to Instrument", command=self._apply_instrument_settings, state=tk.DISABLED, style='Orange.TButton')
        self.apply_button.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="ew")

        # Removed the load_preset_button from here, as it's now handled within PresetFilesTab
        # self.load_preset_button = ttk.Button(instrument_frame, text="Load Selected Preset", command=self._load_selected_preset, state=tk.DISABLED, style='Purple.TButton')
        # self.load_preset_button.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky="ew")

        # Debugging checkboxes (Moved here from Scan Configuration)
        ttk.Checkbutton(instrument_frame, text="Enable General Debug", variable=self.general_debug_enabled_var, style='TCheckbutton').grid(row=4, column=0, columnspan=3, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(instrument_frame, text="Log VISA Commands", variable=self.log_visa_commands_enabled_var, style='TCheckbutton').grid(row=5, column=0, columnspan=3, sticky="w", padx=5, pady=2)


        # --- ScanTab (New tab for Scan Configuration and Bands to Scan) ---
        self.scan_tab = ScanTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.left_notebook.add(self.scan_tab, text="Scan Configuration")


        # --- Existing Notebook (Tabs) for Instrument Presets, Markers, Plotting, Report Converter ---
        # These tabs are now added to the same left_notebook
        self.preset_files_tab = PresetFilesTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.left_notebook.add(self.preset_files_tab, text="Instrument Presets")

        self.markers_display_tab = MarkersDisplayTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.left_notebook.add(self.markers_display_tab, text="Markers Display")

        self.plotting_tab = PlottingTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.left_notebook.add(self.plotting_tab, text="Plotting")

        self.report_converter_tab = ReportConverterTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.left_notebook.add(self.report_converter_tab, text="Report Converter")

        # Bind tab selection event to update specific tab contents for the left notebook
        self.left_notebook.bind("<<NotebookTabChanged>>", self._on_tab_change)


        # --- Right Column Container Frame ---
        # This frame will hold the Scan Control and Application Console
        right_column_container = ttk.Frame(self, style='Dark.TFrame')
        right_column_container.grid(row=0, column=1, padx=5, pady=5, sticky="nsew") # Place in main window's grid
        right_column_container.grid_columnconfigure(0, weight=1) # Single column within this container
        right_column_container.grid_rowconfigure(0, weight=0) # Scan Control row (fixed height)
        right_column_container.grid_rowconfigure(1, weight=1) # Application Console row (expands)


        # --- Scan Control Buttons Frame (Moved into right_column_container) ---
        scan_control_frame = ttk.LabelFrame(right_column_container, text="Scan Control", style='Dark.TLabelframe')
        scan_control_frame.grid(row=0, column=0, padx=5, pady=5, sticky="new") # Top of right column container
        scan_control_frame.grid_columnconfigure(0, weight=1)
        scan_control_frame.grid_columnconfigure(1, weight=1)
        scan_control_frame.grid_columnconfigure(2, weight=1)

        self.start_scan_button = ttk.Button(scan_control_frame, text="Start Scan", command=self._start_scan, style='Green.TButton')
        self.start_scan_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.pause_scan_button = ttk.Button(scan_control_frame, text="Pause Scan", command=self._pause_scan, state=tk.DISABLED, style='Orange.TButton')
        self.pause_scan_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.stop_scan_button = ttk.Button(scan_control_frame, text="Stop Scan", command=self._stop_scan, state=tk.DISABLED, style='Red.TButton')
        self.stop_scan_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")


        # --- Application Console Frame (Moved into right_column_container) ---
        console_frame = ttk.LabelFrame(right_column_container, text="Application Console", style='Dark.TLabelframe')
        console_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew") # Below scan control in right column container
        console_frame.grid_rowconfigure(0, weight=1)
        console_frame.grid_columnconfigure(0, weight=1)

        # Removed fixed height to allow grid weight to control height
        self.console_text = scrolledtext.ScrolledText(console_frame, wrap="word", bg="#2b2b2b", fg="#cccccc", insertbackground="white")
        self.console_text.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.console_text.config(state=tk.DISABLED) # Make it read-only
        
        debug_print("Main application widgets created.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _setup_styles(self):
        """
        Configures and applies custom ttk styles for a modern dark theme.
        """
        debug_print("Setting up ttk styles...", file=__file__, function=inspect.currentframe().f_code.co_name)
        style = ttk.Style(self)

        # Overall theme
        style.theme_use('clam') # 'clam', 'alt', 'default', 'classic'

        # General background and foreground colors
        BG_DARK = "#1e1e1e" # Very dark grey
        FG_LIGHT = "#cccccc" # Light grey
        ACCENT_BLUE = "#007bff" # Bootstrap primary blue
        ACCENT_GREEN = "#28a745" # Bootstrap success green
        ACCENT_RED = "#dc3545" # Bootstrap danger red
        ACCENT_ORANGE = "#ffc107" # Bootstrap warning orange (for pause/apply)
        ACCENT_PURPLE = "#6f42c1" # Bootstrap purple

        style.configure('.', background=BG_DARK, foreground=FG_LIGHT, font=('Helvetica', 10))
        style.configure('TFrame', background=BG_DARK)
        style.configure('TLabel', background=BG_DARK, foreground=FG_LIGHT)
        style.configure('TEntry', fieldbackground="#3b3b3b", foreground="#ffffff", borderwidth=1, relief="flat")
        style.map('TEntry', fieldbackground=[('focus', '#4a4a4a')]) # Slightly lighter on focus
        style.configure('TCombobox', fieldbackground="#3b3b3b", foreground="#ffffff", selectbackground=ACCENT_BLUE, selectforeground="white")
        style.map('TCombobox', fieldbackground=[('readonly', '#3b3b3b')], arrowcolor=[('!disabled', FG_LIGHT)])

        # Buttons
        style.configure('TButton',
                        background="#4a4a4a", # Darker grey for buttons
                        foreground="white",
                        font=('Helvetica', 10, 'bold'),
                        borderwidth=0,
                        focusthickness=3,
                        focuscolor=ACCENT_BLUE,
                        padding=5)
        style.map('TButton',
                background=[('active', '#606060'), ('disabled', '#303030')],
                foreground=[('disabled', '#808080')])

        # Specific button styles
        style.configure('Blue.TButton', background=ACCENT_BLUE, foreground="white")
        style.map('Blue.TButton', background=[('active', '#0056b3'), ('disabled', '#004085')])

        style.configure('Green.TButton', background=ACCENT_GREEN, foreground="white")
        style.map('Green.TButton', background=[('active', '#218838'), ('disabled', '#1e7e34')])

        style.configure('Red.TButton', background=ACCENT_RED, foreground="white")
        style.map('Red.TButton', background=[('active', '#c82333'), ('disabled', '#bd2130')])

        style.configure('Orange.TButton', background=ACCENT_ORANGE, foreground="#333333") # Dark text for contrast
        style.map('Orange.TButton', background=[('active', '#e0a800'), ('disabled', '#d39e00')])

        style.configure('Purple.TButton', background=ACCENT_PURPLE, foreground="white")
        style.map('Purple.TButton', background=[('active', '#5a2d9e'), ('disabled', '#4d2482')])


        # Checkbuttons
        style.configure('TCheckbutton', background=BG_DARK, foreground=FG_LIGHT, indicatorcolor="#4a4a4a")
        style.map('TCheckbutton',
                background=[('active', BG_DARK)], # Keep background same on active
                foreground=[('disabled', '#808080')],
                indicatorcolor=[('selected', ACCENT_BLUE)]) # Blue checkmark when selected

        # LabelFrame
        style.configure('TLabelFrame', background=BG_DARK, foreground=FG_LIGHT, borderwidth=1, relief="solid")
        style.configure('TLabelFrame.Label', background=BG_DARK, foreground=FG_LIGHT, font=('Helvetica', 10, 'bold'))
        style.configure('Dark.TLabelframe', background="#1e1e1e", foreground="#cccccc") # Consistent dark background and light text
        style.configure('Dark.TLabelframe.Label', background="#1e1e1e", foreground="#cccccc") # Ensure label matches
        style.configure('Dark.TFrame', background="#1e1e1e") # For inner frames

        # Notebook (Tabs)
        style.configure('TNotebook', background=BG_DARK, borderwidth=0)
        style.configure('TNotebook.Tab', background="#3b3b3b", foreground=FG_LIGHT, padding=[10, 5])
        style.map('TNotebook.Tab',
                background=[('selected', ACCENT_BLUE), ('active', '#4a4a4a')],
                foreground=[('selected', 'white')],
                expand=[('selected', [1, 1, 1, 0])]) # Expand selected tab slightly

        # Treeview (for MarkersDisplayTab)
        style.configure('Treeview',
                        background="#3b3b3b", # Darker grey for treeview background
                        foreground="#ffffff", # White text
                        fieldbackground="#3b3b3b", # Same as background
                        rowheight=25) # Adjust row height for better spacing
        style.map('Treeview',
                background=[('selected', ACCENT_BLUE)], # Blue highlight for selected item
                foreground=[('selected', 'white')]) # White text on selected item
        
        style.configure('Treeview.Heading',
                        background="#4a4a4a", # Darker grey for headings
                        foreground="white",
                        font=('Helvetica', 10, "bold"))
        style.map('Treeview.Heading',
                background=[('active', '#606060')]) # Lighter grey on active heading

        # Markers Tab Specific Styles
        style.configure("Markers.TFrame", background="#1e1e1e") # Main frame background
        style.configure("Markers.TLabel", background="#1e1e1e", foreground="#cccccc")
        style.configure("Markers.TButton",
                        background="#6a5acd", # Medium purple
                        foreground="white",
                        font=("Helvetica", 10, "bold"),
                        padding=[10, 5])
        style.map("Markers.TButton",
                background=[('active', '#8a7ad0')]) # Lighter purple on active

        style.configure("SelectedSpan.TButton", # Style for selected span button
                        background="#ff6347", # Tomato red/orange
                        foreground="white",
                        font=("Helvetica", 10, "bold"),
                        padding=[10, 5])
        style.map("SelectedSpan.TButton",
                background=[('active', '#e05038')]) # Darker red/orange on active

        style.configure("Markers.Inner.Treeview",
                    background="#2b2b2b", # Darker grey
                    foreground="#cccccc", # Light grey text
                    fieldbackground="#2b2b2b") # Darker grey
        style.map("Markers.Inner.Treeview",
              background=[("selected", "#0056b3")], # Darker blue highlight for treeview
              foreground=[("selected", "white")])
    
        # Configure the base TLabelFrame style and its label part
        style.configure("TLabelFrame", background="#1e1e1e", foreground="#cccccc") # Consistent dark background and light text
        style.configure("TLabelFrame.Label", background="#1e1e1e", foreground="#cccccc") # Ensure label matches
        
        # Update font size for preset buttons to 14pt (smaller than 20pt)
        style.configure("LargePreset.TButton",
                        background="#4a4a4a", # Darker grey for buttons
                        foreground="white",
                        font=("Helvetica", 14, "bold"), # Set font size to 14
                        padding=[30, 15, 30, 15]) # Adjust padding as needed
        style.map("LargePreset.TButton",
                background=[('active', '#606060')]) # Lighter grey on active

        style.configure("SelectedPreset.TButton",
                        background="#007bff", # A nice blue color (consistent with Blue.TButton)
                        foreground="white",
                        font=("Helvetica", 14, "bold"), # Keep the 14-point font
                        padding=[30, 15, 30, 15])
        style.map("SelectedPreset.TButton",
                background=[('active', '#0056b3')]) # Darker blue on active

        debug_print("ttk styles set up.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _redirect_stdout_to_console(self):
        """
        Redirects standard output and error streams to the GUI's scrolled text widget.
        """
        debug_print("Redirecting stdout/stderr to GUI console...", file=__file__, function=inspect.currentframe().f_code.co_name)
        sys.stdout = TextRedirector(self.console_text, "stdout")
        sys.stderr = TextRedirector(self.console_text, "stderr")
        print("Application console initialized.") # This will now print to the GUI console


    def _print_to_gui_console(self, message):
        """
        A helper function to print messages to the GUI console from any thread.
        Uses after() to ensure thread safety.
        """
        self.after(0, lambda: self._update_console_text(message))

    def _update_console_text(self, message):
        """
        Appends a message to the scrolled text widget.
        """
        self.console_text.config(state=tk.NORMAL)
        self.console_text.insert(tk.END, message + "\n")
        self.console_text.see(tk.END) # Scroll to the end
        self.console_text.config(state=tk.DISABLED)


    def _populate_resources(self):
        """
        Populates the VISA resource combobox with available instruments.
        """
        debug_print("Populating VISA resources...", file=__file__, function=inspect.currentframe().f_code.co_name)
        # Corrected the argument passed to populate_resources_logic
        populate_resources_logic(self, self._print_to_gui_console)


    def _on_resource_selected(self, event):
        """
        Callback when a VISA resource is selected from the combobox.
        """
        selected_resource = self.resource_combobox.get()
        self._print_to_gui_console(f"Selected resource: {selected_resource}")
        debug_print(f"Resource selected: {selected_resource}", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _connect_instrument(self):
        """
        Attempts to connect to the selected VISA instrument.
        """
        debug_print("Attempting to connect instrument...", file=__file__, function=inspect.currentframe().f_code.co_name)
        selected_resource = self.resource_combobox.get()
        
        # Pass self (app_instance) to the logic function
        connect_instrument_logic(self, selected_resource, self._print_to_gui_console)
        # Update button states after connection attempt
        _update_button_states_on_connection(self)


    def _disconnect_instrument(self):
        """
        Disconnects from the currently connected VISA instrument.
        """
        debug_print("Attempting to disconnect instrument...", file=__file__, function=inspect.currentframe().f_code.co_name)
        # Pass self (app_instance) to the logic function
        disconnect_instrument_logic(self, self._print_to_gui_console)
        # Update button states after disconnection
        _update_button_states_on_connection(self)


    def _apply_instrument_settings(self):
        """
        Applies the current settings from the GUI to the connected instrument.
        """
        debug_print("Applying instrument settings...", file=__file__, function=inspect.currentframe().f_code.co_name)
        # Pass self (app_instance) to the logic function
        apply_settings_logic(self, self._print_to_gui_console)


    def _load_selected_preset(self):
        """
        Loads the currently selected preset file onto the instrument.
        This function is called when the "Load Selected Preset" button is clicked.
        It delegates the actual loading logic to instrument_preset_tab.
        """
        debug_print("Loading selected preset...", file=__file__, function=inspect.currentframe().f_code.co_name)
        if hasattr(self, 'preset_files_tab'):
            # Call the _load_selected_preset method on the preset_files_tab instance
            self.preset_files_tab._load_selected_preset()
        else:
            self._print_to_gui_console("⚠️ Warning: Preset Files tab not initialized.")
            debug_print("PresetFilesTab not initialized.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _start_scan(self):
        """
        Initiates the scan process in a separate thread.
        """
        debug_print("Starting scan...", file=__file__, function=inspect.currentframe().f_code.co_name)
        # Pass self (app_instance) to the logic function
        start_scan_thread_logic(self)


    def _pause_scan(self):
        """
        Pauses the active scan.
        """
        debug_print("Pausing scan...", file=__file__, function=inspect.currentframe().f_code.co_name)
        # Pass self (app_instance) to the logic function
        pause_scan_logic(self)


    def _stop_scan(self):
        """
        Stops the active scan.
        """
        debug_print("Stopping scan...", file=__file__, function=inspect.currentframe().f_code.co_name)
        # Pass self (app_instance) to the logic function
        stop_scan_logic(self)


    def _browse_output_folder(self):
        """
        Opens a file dialog to select the output folder for scan data.
        """
        debug_print("Browsing output folder...", file=__file__, function=inspect.currentframe().f_code.co_name)
        folder_selected = filedialog.askdirectory(initialdir=self.output_folder_var.get())
        if folder_selected:
            self.output_folder_var.set(folder_selected)
            self._print_to_gui_console(f"Output folder set to: {folder_selected}")
            debug_print(f"Output folder set to: {folder_selected}", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _restore_default_settings(self):
        """
        Restores all settings to their default values as defined in config.ini.
        This function is now primarily handled by the ScanTab's button,
        but this method remains for consistency if needed elsewhere.
        """
        debug_print("Restoring default settings (from App)...", file=__file__, function=inspect.currentframe().f_code.co_name)
        restore_default_settings_logic(self)


    def reset_setting_colors_logic(self):
        """
        Resets the background color of all setting entry widgets to default.
        This function is called after settings are applied or restored to remove
        any visual indication of unsaved changes.
        """
        debug_print("Resetting setting colors...", file=__file__, function=inspect.currentframe().f_code.co_name)
        # Iterate through all Tkinter variables in the map
        for tk_var_name, (_, _, tk_var_instance) in self.setting_var_map.items():
            # Find the corresponding widget. This is a generic approach;
            # a more robust solution might involve storing widget references.
            # For now, we'll assume entry widgets are the primary ones needing color reset.
            if isinstance(tk_var_instance, (tk.StringVar, tk.IntVar, tk.DoubleVar)):
                # This is a bit hacky, but tries to find the entry widget associated
                # with the variable. A better design would be to store widget references.
                # For now, we'll just print a debug message if we can't find it.
                # This function is mainly for visual feedback, so it's not critical
                # if it doesn't find every single widget.
                try:
                    # Attempt to find the entry widget by iterating through children
                    # This is not guaranteed to work for all layouts
                    for widget in self.winfo_children():
                        if isinstance(widget, (ttk.LabelFrame, tk.Frame, ttk.Notebook)):
                            for child in widget.winfo_children():
                                if isinstance(child, ttk.Frame): # Check frames within notebooks/labelframes
                                    for grand_child in child.winfo_children():
                                        if isinstance(grand_child, ttk.Entry) and grand_child.cget('textvariable') == str(tk_var_instance):
                                            grand_child.config(style='TEntry') # Reset to default style
                                            debug_print(f"Reset color for {tk_var_name}", file=__file__, function=inspect.currentframe().f_code.co_name)
                                            break
                                if isinstance(child, ttk.Entry) and child.cget('textvariable') == str(tk_var_instance):
                                    child.config(style='TEntry') # Reset to default style
                                    debug_print(f"Reset color for {tk_var_name}", file=__file__, function=inspect.currentframe().f_code.co_name)
                                    break
                except Exception as e:
                    debug_print(f"Could not reset color for {tk_var_name}: {e}", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _on_closing(self):
        """
        Handles the window closing event. Saves configuration before exiting.
        """
        debug_print("Application closing. Performing cleanup...", file=__file__, function=inspect.currentframe().f_code.co_name)
        save_config(self) # Save current settings to config.ini
        self.destroy() # Close the application


    def _on_tab_change(self, event):
        """
        Handles tab change events in the Notebook.
        Calls _on_tab_selected for the newly selected tab if available.
        """
        # Get the currently selected tab's widget from the left_notebook
        selected_tab_id = self.left_notebook.select()
        selected_tab_widget = self.left_notebook.nametowidget(selected_tab_id)
        
        # Check if the selected tab has an _on_tab_selected method and call it
        if hasattr(selected_tab_widget, '_on_tab_selected'):
            selected_tab_widget._on_tab_selected(event)
            debug_print(f"Tab changed to {selected_tab_widget.winfo_class()}. Calling _on_tab_selected.", file=__file__, function=inspect.currentframe().f_code.co_name)
        else:
            debug_print(f"Tab changed to {selected_tab_widget.winfo_class()}. No _on_tab_selected method found.", file=__file__, function=inspect.currentframe().f_code.co_name)


if __name__ == "__main__":
    app = App()
    app.mainloop()

