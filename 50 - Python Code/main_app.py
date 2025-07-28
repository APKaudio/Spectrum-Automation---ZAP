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
    apply_settings_logic, query_current_instrument_settings_logic,
    load_selected_preset_logic, query_device_presets_logic
)
from src.scan_logic import update_connection_status_logic # Only update_connection_status_logic remains here
from src.settings_logic import restore_default_settings_logic
from src.plotting_tab import PlottingTab # Import PlottingTab
from src.marker_tab import MarkersDisplayTab # Import the MarkersDisplayTab
from src.report_converter_tab import ReportConverterTab # Import ReportConverterTab
from src.instrument_tab import InstrumentTab # Import the new InstrumentTab
from src.scan_control import ScanControlTab # Import the new ScanControlTab

# Import debug_print from utils
from utils.instrument_control import set_debug_mode, log_visa_command, debug_print


class App(tk.Tk):
    """
    Main application class for the RF Spectrum Analyzer Controller.
    Manages the GUI, instrument connection, scanning, plotting, and reporting.
    """
    def __init__(self):
        """
        Initializes the main application window and its components.
        """
        super().__init__()
        self.title("SPANALIZER") # Changed window title here
        # self.iconbitmap(self.resource_path("icons/spectrum_analyzer_icon.ico")) # Set window icon
        
        self.CONFIG_FILE = self.resource_path('config.ini')
        self.config = configparser.ConfigParser()
        
        # Initialize Tkinter variables and setting_var_map BEFORE loading configuration
        # Tkinter variables for settings (using StringVar for text entry, BooleanVar for checkboxes)
        self.selected_resource = tk.StringVar(self)
        self.resource_names = tk.StringVar(self) # Holds list of available resources
        # Removed center_freq_var, span_var, rbw_var as they are now displayed in InstrumentTab
        self.ref_level_var = tk.StringVar(self)
        self.freq_shift_var = tk.StringVar(self)
        self.max_hold_enabled_var = tk.BooleanVar(self)
        self.high_sensitivity_var = tk.BooleanVar(self)
        self.preamp_on_var = tk.BooleanVar(self) # Redundant with high_sensitivity_var, but kept for clarity if needed
        self.rbw_segmentation_var = tk.StringVar(self)
        self.desired_default_focus_width_var = tk.StringVar(self)
        self.num_scan_cycles_var = tk.StringVar(self)
        self.cycle_wait_time_var = tk.StringVar(self)
        self.maxhold_time_var = tk.StringVar(self)
        self.output_folder_var = tk.StringVar(self)
        self.scan_name_var = tk.StringVar(self)
        self.open_html_after_complete_var = tk.BooleanVar(self)
        self.general_debug_enabled_var = tk.BooleanVar(self)
        self.log_visa_commands_enabled_var = tk.BooleanVar(self)
        self.include_tv_markers_var = tk.BooleanVar(self)
        self.include_gov_markers_var = tk.BooleanVar(self)
        self.include_markers_var = tk.BooleanVar(self) # For general markers from CSV/report
        # Add the missing rbw_step_size_hz_var
        self.rbw_step_size_hz_var = tk.StringVar(self)
        # Add the newly identified missing variables (renamed to match config.ini)
        self.cycle_wait_time_seconds_var = tk.StringVar(self)
        self.maxhold_time_seconds_var = tk.StringVar(self)
        self.scan_rbw_hz_var = tk.StringVar(self)
        self.reference_level_dbm_var = tk.StringVar(self)
        self.freq_shift_hz_var = tk.StringVar(self)


        # Map Tkinter variables to their config keys for easy loading/saving
        self.setting_var_map = {
            # Removed center_freq_var, span_var, rbw_var from map
            'ref_level_var': ('last_reference_level_dbm', 'default_reference_level_dbm', self.ref_level_var), # This is for display
            'freq_shift_var': ('last_freq_shift_hz', 'default_freq_shift_hz', self.freq_shift_var), # This is for display
            'max_hold_enabled_var': ('last_maxhold_enabled', 'default_maxhold_enabled', self.max_hold_enabled_var),
            'high_sensitivity_var': ('last_high_sensitivity', 'default_high_sensitivity', self.high_sensitivity_var),
            'preamp_on_var': ('last_preamp_on', 'default_preamp_on', self.preamp_on_var),
            'rbw_segmentation_var': ('last_scan_rbw_segmentation', 'default_scan_rbw_segmentation', self.rbw_segmentation_var),
            'desired_default_focus_width_var': ('last_default_focus_width', 'default_default_focus_width', self.desired_default_focus_width_var),
            'num_scan_cycles_var': ('last_num_scan_cycles', 'default_num_scan_cycles', self.num_scan_cycles_var),
            'cycle_wait_time_var': ('last_cycle_wait_time_seconds', 'default_cycle_wait_time_seconds', self.cycle_wait_time_var),
            'maxhold_time_var': ('last_maxhold_time_seconds', 'default_maxhold_time_seconds', self.maxhold_time_var),
            'output_folder_var': ('last_scan_directory', 'default_scan_directory', self.output_folder_var),
            'scan_name_var': ('last_scan_name', 'default_scan_name', self.scan_name_var),
            'open_html_after_complete_var': ('last_open_html_after_complete', 'default_open_html_after_complete', self.open_html_after_complete_var),
            'general_debug_enabled_var': ('last_general_debug_enabled', 'default_general_debug_enabled', self.general_debug_enabled_var),
            'log_visa_commands_enabled_var': ('last_log_visa_commands_enabled', 'default_log_visa_commands_enabled', self.log_visa_commands_enabled_var),
            'include_tv_markers_var': ('last_include_tv_markers', 'default_include_tv_markers', self.include_tv_markers_var),
            'include_gov_markers_var': ('last_include_gov_markers', 'default_include_gov_markers', self.include_gov_markers_var),
            'include_markers_var': ('last_include_markers', 'default_include_markers', self.include_markers_var),
            # Add rbw_step_size_hz_var to the map
            'rbw_step_size_hz_var': ('last_rbw_step_size_hz', 'default_rbw_step_size_hz', self.rbw_step_size_hz_var),
            # Add the newly identified missing variables to the map
            'cycle_wait_time_seconds_var': ('last_cycle_wait_time_seconds', 'default_cycle_wait_time_seconds', self.cycle_wait_time_seconds_var),
            'maxhold_time_seconds_var': ('last_maxhold_time_seconds', 'default_maxhold_time_seconds', self.maxhold_time_seconds_var),
            'scan_rbw_hz_var': ('last_scan_rbw_hz', 'default_scan_rbw_hz', self.scan_rbw_hz_var),
            'reference_level_dbm_var': ('last_reference_level_dbm', 'default_reference_level_dbm', self.reference_level_dbm_var),
            'freq_shift_hz_var': ('last_freq_shift_hz', 'default_freq_shift_hz', self.freq_shift_hz_var),
        }

        self.load_configuration() # Load config.ini at startup, now that setting_var_map exists

        self._apply_saved_geometry() # Apply window geometry from config

        self.inst = None # Placeholder for the VISA instrument object
        # self.scan_thread = None # Moved to ScanControlTab
        # self.is_scanning = False # Moved to ScanControlTab
        self.collected_scans_dataframes = [] # To store pandas DataFrames of collected scans
        self.last_scan_markers = [] # To store markers extracted from the last scan/report

        # Band selection variables (initialized after config load)
        from utils.frequency_bands import SCAN_BAND_RANGES # Import here to avoid circular dependency
        self.SCAN_BAND_RANGES = SCAN_BAND_RANGES # Store for easy access
        self.band_vars = []
        self._initialize_band_vars() # Initialize Tkinter BooleanVars for each band

        # Apply last selected bands from config
        self._apply_last_selected_bands()

        self._set_initial_slider_positions() # Set initial slider positions based on loaded values
        self._setup_styles() # Setup ttk styles
        self._create_widgets() # Create GUI elements
        self._redirect_console() # Redirect console output to the GUI text widget

        # Bind the window closing protocol
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Initial update of connection status
        # This will now be handled by the InstrumentTab's _on_tab_selected when it's initialized
        # and ScanControlTab's update_scan_button_states.
        # self.update_connection_status(False) # Removed from here

        # Bind tab change event for notebook
        self.left_notebook.bind("<<NotebookTabChanged>>", self._on_tab_change)
        
        # Print the cool ASCII art logo
        self.after(100, print_art) # Schedule after a short delay to ensure console is ready

        # Initial population of resources (moved to InstrumentTab's _on_tab_selected or a dedicated call)
        # The InstrumentTab will handle its own initial population when it's created or selected.

        # Set initial debug mode based on config (now handled by InstrumentTab's checkboxes)
        # set_debug_mode(self.general_debug_enabled_var.get()) # Removed
        # log_visa_command(self.log_visa_commands_enabled_var.get()) # Removed
        # self.general_debug_enabled_var.trace_add("write", lambda *args: set_debug_mode(self.general_debug_enabled_var.get())) # Removed
        # self.log_visa_commands_enabled_var.trace_add("write", lambda *args: log_visa_command(self.log_visa_commands_enabled_var.get())) # Removed

        debug_print("Application initialized.", file=__file__, function=inspect.currentframe().f_code.co_name)

    def resource_path(self, relative_path):
        """
        Get the absolute path to a resource, works for dev and for PyInstaller.
        """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        
        return os.path.join(base_path, relative_path)

    def load_configuration(self):
        """
        Loads configuration from config.ini using config_manager.
        """
        load_config(self)
        debug_print(f"Configuration loaded from {self.CONFIG_FILE}", file=__file__, function=inspect.currentframe().f_code.co_name)

    def load_last_used_settings(self):
        """
        Loads the last used settings from the config object into Tkinter variables.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Loading last used settings...", file=current_file, function=current_function)
        for tk_var_name, (last_key, default_key, tk_var_instance) in self.setting_var_map.items():
            if last_key: # Only load if a last_key is defined
                value_str = self.config.get('LAST_USED_SETTINGS', last_key, fallback=None)
                if value_str is not None:
                    try:
                        # Convert string to appropriate type for Tkinter variable
                        if isinstance(tk_var_instance, tk.BooleanVar):
                            tk_var_instance.set(value_str.lower() == 'true')
                        elif isinstance(tk_var_instance, tk.DoubleVar): # Handle DoubleVar if used
                            tk_var_instance.set(float(value_str))
                        else: # Default to StringVar
                            tk_var_instance.set(value_str)
                        debug_print(f"Loaded '{last_key}': '{value_str}'", file=current_file, function=current_function)
                    except ValueError as e:
                        debug_print(f"Error loading '{last_key}' with value '{value_str}': {e}", file=current_file, function=current_function)
                else:
                    debug_print(f"No last used setting found for '{last_key}'.", file=current_file, function=current_function)
        
        # Handle output_folder_var specifically if default is relative
        if not self.output_folder_var.get():
            default_output_folder = self.config.get('DEFAULT_SETTINGS', 'default_scan_directory', fallback='scan_data')
            self.output_folder_var.set(os.path.abspath(default_output_folder)) # Set absolute path
            debug_print(f"Set default output folder to absolute path: {self.output_folder_var.get()}", file=current_file, function=current_function)

    def _apply_saved_geometry(self):
        """
        Applies the last saved window geometry or a default.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        geometry = self.config.get('LAST_USED_SETTINGS', 'last_window_geometry', fallback=None)
        if geometry:
            try:
                self.geometry(geometry)
                debug_print(f"Applied saved window geometry: {geometry}", file=current_file, function=current_function)
            except TclError as e:
                debug_print(f"Error applying saved geometry '{geometry}': {e}. Using default.", file=current_file, function=current_function)
                # Fallback to default if saved geometry is invalid
                default_geometry = self.config.get('DEFAULT_SETTINGS', 'default_window_geometry', fallback='1400x780+100+100')
                self.geometry(default_geometry)
        else:
            default_geometry = self.config.get('DEFAULT_SETTINGS', 'default_window_geometry', fallback='1400x780+100+100')
            self.geometry(default_geometry)
            debug_print(f"No saved geometry found. Applied default: {default_geometry}", file=current_file, function=current_function)


    def _initialize_band_vars(self):
        """
        Initializes Tkinter BooleanVars for each scan band.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Initializing band selection variables...", file=current_file, function=current_function)
        from utils.frequency_bands import SCAN_BAND_RANGES # Ensure this is imported here for safety
        for band in SCAN_BAND_RANGES:
            var = tk.BooleanVar(self, value=True) # Default to selected
            # Pass the entire band_item dictionary to the callback
            band_item = {"band": band, "var": var}
            self.band_vars.append(band_item)
            var.trace_add("write", lambda *args, b_item=band_item: self._on_band_selection_change(b_item))

    def _apply_last_selected_bands(self):
        """
        Applies the last selected bands from config.ini to the checkboxes.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Applying last selected bands...", file=current_file, function=current_function)
        last_selected_bands_str = self.config.get('LAST_USED_SETTINGS', 'last_selected_bands', fallback='')
        if last_selected_bands_str:
            last_selected_band_names = [name.strip() for name in last_selected_bands_str.split(',') if name.strip()]
            for band_item in self.band_vars:
                band_item["var"].set(band_item["band"]["Band Name"] in last_selected_band_names)
            debug_print(f"Restored selected bands: {last_selected_band_names}", file=current_file, function=current_function)
        else:
            # If no last_selected_bands, ensure all are selected by default
            for band_item in self.band_vars:
                band_item["var"].set(True)
            debug_print("No last_selected_bands in config.ini. All bands set to selected.", file=current_file, function=current_function)


    def _on_band_selection_change(self, band_item):
        """
        Callback when a band selection checkbox changes.
        Now receives the full band_item dictionary.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        # Access 'band' and 'var' keys from the received band_item dictionary
        debug_print(f"Band selection changed for {band_item['band']['Band Name']}: {band_item['var'].get()}", file=current_file, function=current_function)
        save_config(self) # Save config immediately on change

    def _set_initial_slider_positions(self):
        """
        Sets the initial positions of the sliders based on loaded config values.
        This must be called after Tkinter variables are loaded.
        """
        # This function might be called by individual tabs if they manage their own sliders.
        # For a centralized approach, ensure it's called after all config values are loaded.
        pass # The actual slider update logic will be in the respective tabs (e.g., ScanTab)

    def _setup_styles(self):
        """
        Configures the ttk styles for a dark theme.
        """
        s = ttk.Style()

        # Set a theme that allows for more customization
        s.theme_use('clam') # 'clam' is often a good base for dark themes

        # General dark theme for all widgets
        s.configure('.', background='#1e1e1e', foreground='#cccccc', font=('Inter', 10)) # Default font
        s.configure('TLabel', background='#1e1e1e', foreground='#cccccc')
        s.configure('TLabelframe', background='#1e1e1e', foreground='#cccccc', bordercolor='#444444')
        s.configure('TLabelframe.Label', background='#1e1e1e', foreground='#cccccc', font=('Inter', 10, 'bold'))
        s.configure('TEntry', fieldbackground='#333333', foreground='#ffffff', bordercolor='#555555')
        s.configure('TCombobox', fieldbackground='#333333', foreground='#ffffff', selectbackground='#007bff', selectforeground='white')
        s.map('TCombobox', fieldbackground=[('readonly', '#333333')]) # Keep dark background for readonly

        # Notebook (Tabs) styles
        s.configure('TNotebook', background='#1e1e1e', borderwidth=0)
        s.configure('TNotebook.Tab', background='#3a3a3a', foreground='#cccccc', padding=[10, 5], font=('Inter', 10, 'bold'))
        s.map('TNotebook.Tab', background=[('selected', '#007bff'), ('active', '#555555')], # Blue for selected, darker grey for active
                               foreground=[('selected', 'white'), ('active', 'white')])

        # Frame styles
        s.configure('Dark.TFrame', background='#1e1e1e')
        s.configure('Dark.TLabelframe', background='#1e1e1e', foreground='#cccccc', bordercolor='#444444')
        s.configure('Dark.TLabelframe.Label', background='#1e1e1e', foreground='#cccccc')

        # Button styles
        s.configure('TButton',
                    background='#007bff', # Default blue for general buttons
                    foreground='white',
                    font=('Inter', 10, 'bold'),
                    borderwidth=1,
                    relief="raised",
                    focusthickness=3,
                    focuscolor='#007bff')
        s.map('TButton',
              background=[('active', '#0056b3')], # Darker blue on hover
              foreground=[('disabled', '#888888')])

        # Custom styles for Scan Control buttons
        s.configure('Green.TButton',
                    background='#28a745', # Green
                    foreground='white',
                    font=('Inter', 10, 'bold'),
                    borderwidth=1,
                    relief="raised",
                    focusthickness=3,
                    focuscolor='#218838')
        s.map('Green.TButton',
              background=[('active', '#218838')],
              foreground=[('disabled', '#888888')])

        s.configure('Orange.TButton',
                    background='#ffc107', # Orange
                    foreground='black', # Black text for contrast
                    font=('Inter', 10, 'bold'),
                    borderwidth=1,
                    relief="raised",
                    focusthickness=3,
                    focuscolor='#e0a800')
        s.map('Orange.TButton',
              background=[('active', '#e0a800')],
              foreground=[('disabled', '#888888')])

        s.configure('Red.TButton',
                    background='#dc3545', # Red
                    foreground='white',
                    font=('Inter', 10, 'bold'),
                    borderwidth=1,
                    relief="raised",
                    focusthickness=3,
                    focuscolor='#c82333')
        s.map('Red.TButton',
              background=[('active', '#c82333')],
              foreground=[('disabled', '#888888')])


        # Styles for Markers Tab specific elements
        s.configure('Markers.TFrame', background='#1e1e1e')
        s.configure('Markers.Inner.Treeview', background='#2b2b2b', foreground='#cccccc', fieldbackground='#2b2b2b')
        s.map('Markers.Inner.Treeview', background=[('selected', '#555555')], foreground=[('selected', 'white')])
        s.configure('Markers.TLabel', background='#1e1e1e', foreground='#cccccc') # For labels within Markers tab

        # Style for Span and Trace Mode buttons (unselected state: neutral/grey)
        s.configure('Markers.TButton',
                    background='#3a3a3a', # A neutral, "great" grey/dark blue
                    foreground='white',
                    font=('Inter', 10, 'bold'),
                    borderwidth=1,
                    relief="raised",
                    focusthickness=3,
                    focuscolor='#007bff') # A subtle blue focus
        s.map('Markers.TButton',
              background=[('active', '#505050')], # Darker grey on hover
              foreground=[('disabled', '#888888')])

        # Style for selected Span and Trace Mode buttons (orange)
        s.configure('SelectedSpan.TButton',
                    background='#FF8C00', # DarkOrange
                    foreground='white',
                    font=('Inter', 10, 'bold'),
                    borderwidth=1,
                    relief="sunken", # Sunken to show it's pressed
                    focusthickness=3,
                    focuscolor='#007bff')
        s.map('SelectedSpan.TButton',
              background=[('active', '#FF6347')]) # Tomato on hover

        # Style for unselected Device buttons (neutral/grey)
        s.configure('DeviceButton.TButton',
                    background='#4a4a4a', # A slightly lighter neutral grey for device buttons
                    foreground='white',
                    font=('Inter', 9, 'bold'),
                    borderwidth=1,
                    relief="raised",
                    focusthickness=3,
                    focuscolor='#007bff')
        s.map('DeviceButton.TButton',
              background=[('active', '#606060')], # Darker grey on hover
              foreground=[('disabled', '#888888')])

        # Style for selected Device buttons (orange)
        s.configure('SelectedDevice.TButton',
                    background='#FF8C00', # DarkOrange
                    foreground='white',
                    font=('Inter', 9, 'bold'),
                    borderwidth=1,
                    relief="sunken", # Sunken to show it's pressed
                    focusthickness=3,
                    focuscolor='#007bff')
        s.map('SelectedDevice.TButton',
              background=[('active', '#FF6347')]) # Tomato on hover

        # Checkbutton styles
        s.configure('TCheckbutton',
                    background='#1e1e1e',
                    foreground='#cccccc',
                    font=('Inter', 10))
        s.map('TCheckbutton',
              background=[('active', '#1e1e1e')], # Keep background consistent on hover
              foreground=[('disabled', '#888838')])

        # Slider (Scale) styles
        s.configure('Horizontal.TScale',
                    background='#1e1e1e',
                    troughcolor='#333333',
                    sliderrelief='flat',
                    sliderthickness=20,
                    borderwidth=0)
        s.map('Horizontal.TScale',
              background=[('active', '#0056b3')]) # Darker blue on slider hover

        # Entry (Text input) styles - already covered by TEntry, but can be more specific
        s.configure('TEntry',
                    fieldbackground='#333333',
                    foreground='#ffffff',
                    insertcolor='white', # Cursor color
                    bordercolor='#555555',
                    relief='flat') # Flat border for modern look

        # ScrolledText (Console) style
        # Note: ScrolledText is a combination of Text and Scrollbar, so some styling
        # might need to be applied directly to the Text widget part.
        # The TextRedirector handles the background/foreground for the text itself.
        # This style applies to the frame around the ScrolledText.
        s.configure('TScrolledtext',
                    background='#1e1e1e',
                    borderwidth=0,
                    relief='flat')


    def _create_widgets(self):
        """
        Creates and arranges the main GUI widgets, including notebooks and tabs.
        """
        debug_print("Creating main application widgets...", file=__file__, function=inspect.currentframe().f_code.co_name)

        # Main frame to hold everything, packed into the root window
        main_frame = ttk.Frame(self, padding="10 10 10 10", style='Dark.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Configure main_frame grid
        # Row 0: Scan Control Tab (takes minimal height)
        # Row 1: Bottom content (Notebook + Console), expands vertically
        main_frame.grid_rowconfigure(0, weight=0) # Scan control row - fixed height
        main_frame.grid_rowconfigure(1, weight=1) # Bottom content row - expands
        main_frame.grid_columnconfigure(0, weight=1) # Single column for the overall layout

        # Initialize ScanControlTab first as it's a direct child of main_frame
        self.scan_control_tab = ScanControlTab(main_frame, app_instance=self, console_print_func=self._print_to_gui_console)
        self.scan_control_tab.grid(row=0, column=0, sticky="ew", padx=5, pady=5) # Place at top, span across

        # Frame for the bottom content (Notebook on left, Console on right)
        bottom_content_frame = ttk.Frame(main_frame, style='Dark.TFrame')
        bottom_content_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5) # Place below scan control, span across
        
        # Configure bottom_content_frame columns for 50/50 width split
        bottom_content_frame.grid_columnconfigure(0, weight=1) # Left notebook column
        bottom_content_frame.grid_columnconfigure(1, weight=1) # Right console column
        bottom_content_frame.grid_rowconfigure(0, weight=1) # Only one row for this frame, expands vertically

        # Left Notebook for Tabs
        self.left_notebook = ttk.Notebook(bottom_content_frame)
        self.left_notebook.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # Initialize and add tabs to the left notebook
        from src.scan_tab import ScanTab
        from src.instrument_preset_tab import PresetFilesTab

        self.instrument_tab = InstrumentTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.scan_tab = ScanTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.plotting_tab = PlottingTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.markers_tab = MarkersDisplayTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.report_converter_tab = ReportConverterTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.preset_files_tab = PresetFilesTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)

        # Add tabs to the left notebook in the desired order
        self.left_notebook.add(self.instrument_tab, text="Instrument")
        self.left_notebook.add(self.scan_tab, text="Scan Config")
        self.left_notebook.add(self.plotting_tab, text="Plotting")
        self.left_notebook.add(self.markers_tab, text="Markers")
        self.left_notebook.add(self.report_converter_tab, text="Report Converter")
        self.left_notebook.add(self.preset_files_tab, text="Presets")

        # Right Console Output
        console_frame = ttk.LabelFrame(bottom_content_frame, text="Console Output", padding="5 5 5 5", style='Dark.TLabelframe')
        console_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5) # Place in right column of bottom_content_frame
        console_frame.grid_rowconfigure(0, weight=1)
        console_frame.grid_columnconfigure(0, weight=1)

        self.console_text = scrolledtext.ScrolledText(console_frame, wrap=tk.WORD, bg="#000000", fg="#00FF00",
                                                     font=("Consolas", 9), insertbackground="white",
                                                     borderwidth=0, relief="flat", highlightthickness=0)
        self.console_text.grid(row=0, column=0, sticky="nsew")
        self.console_text.tag_config("stderr", foreground="#FF4500") # OrangeRed for errors
        self.console_text.tag_config("stdout", foreground="#00FF00") # LimeGreen for stdout
        self.console_text.tag_config("info", foreground="#1E90FF") # DodgerBlue for info
        self.console_text.tag_config("warning", foreground="#FFD700") # Gold for warnings
        self.console_text.tag_config("error", foreground="#FF0000") # Red for critical errors
        self.console_text.tag_config("debug", foreground="#8A2BE2") # BlueViolet for debug

        # Add a tag to control line spacing (as per TextRedirector)
        self.console_text.tag_config("line_spacing", spacing1=0, spacing3=0)


    def _redirect_console(self):
        """
        Redirects stdout and stderr to the console_text widget.
        """
        sys.stdout = TextRedirector(self.console_text, "stdout")
        sys.stderr = TextRedirector(self.console_text, "stderr")
        debug_print("Console output redirected to GUI.", file=__file__, function=inspect.currentframe().f_code.co_name)

    def _print_to_gui_console(self, message, tag="stdout"):
        """
        Prints a message to the GUI console text widget.
        This function is thread-safe.
        """
        self.after(0, self._update_console_line, message, tag)

    def _update_console_line(self, message, tag="stdout"):
        """
        Helper function to update the console text widget.
        Designed to be called via self.after() for thread safety.
        """
        # Ensure the widget exists and is not destroyed
        if not self.console_text.winfo_exists():
            return

        try:
            # Insert the message at the end
            self.console_text.insert(tk.END, message + "\n", (tag, "line_spacing"))
            # Scroll to the end
            self.console_text.see(tk.END)
        except TclError as e:
            # This can happen if the widget is destroyed between winfo_exists() and insert()
            # or if there's a problem with the tag configuration.
            print(f"Error updating console text widget: {e}")
            debug_print(f"Error updating console text widget: {e}", file=__file__, function=inspect.currentframe().f_code.co_name)


    def update_connection_status(self, is_connected):
        """
        Updates the GUI elements based on the instrument connection status.
        This function is the central dispatcher for updating button states across tabs.
        """
        debug_print(f"Main App: update_connection_status called. Connected: {is_connected}", file=__file__, function=inspect.currentframe().f_code.co_name)

        # Update InstrumentTab buttons and other general status
        if hasattr(self, 'instrument_tab'):
            # Pass the app_instance itself for access to its properties like inst, scan_control_tab
            update_connection_status_logic(
                self, # Pass app_instance
                is_connected,
                self._print_to_gui_console,
                # No need to pass individual buttons here, scan_logic will get them from app_instance
            )
        else:
            debug_print("InstrumentTab not yet initialized when update_connection_status called.", file=__file__, function=inspect.currentframe().f_code.co_name)

        # The scan_control_tab's update_scan_button_states is now called by update_connection_status_logic
        # so no explicit call needed here. This prevents the recursion.


    def reset_setting_colors_logic(self):
        """
        Resets the background color of all setting entry widgets to default.
        This is called after settings are successfully applied.
        """
        # This function should be implemented in the main App class if it manages
        # the Entry widgets directly, or in the ScanTab if it controls its own.
        # For now, it's a placeholder.
        debug_print("Resetting setting colors (placeholder).", file=__file__, function=inspect.currentframe().f_code.co_name)


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
    # Check for required libraries and install if missing
    # This block should ideally be run outside the GUI loop or in a splash screen
    # For simplicity, keeping it here for now.
    required_packages = ['pyvisa', 'pandas', 'plotly', 'beautifulsoup4', 'pdfplumber', 'numpy']
    missing_packages = []
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)

    if missing_packages:
        print(f"Missing packages: {', '.join(missing_packages)}. Attempting to install...")
        try:
            # Use pip to install missing packages
            subprocess.check_call([sys.executable, "-m", "pip", "install", *missing_packages])
            print("✅ Successfully installed missing packages.")
        except subprocess.CalledProcessError as e:
            print(f"❌ Error installing packages: {e}")
            print("Please install them manually using 'pip install <package_name>'")
            sys.exit(1) # Exit if essential packages cannot be installed

    # Check for NI-VISA installation (Windows specific check for now)
    if sys.platform == "win32":
        try:
            rm = pyvisa.ResourceManager()
            # Attempt to list resources to confirm NI-VISA is working
            _ = rm.list_resources()
            print("✅ NI-VISA detected and working.")
        except Exception as e:
            print(f"❌ NI-VISA not detected or not working correctly: {e}")
            print("Please ensure NI-VISA is installed and configured properly.")
            # Do not exit, allow the app to run but connection will fail.
    else:
        print("ℹ️ Info: NI-VISA check skipped (non-Windows OS).")

    app = App()
    app.mainloop()
