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
        messagebox.showerror("Installation Error", f"Error installing BeautifulSoup4: {e}\nPlease install it manually by running: pip install beautifulsoup4")
        sys.exit(1)
    except Exception as e:
        messagebox.showerror("Installation Error", f"An unexpected error occurred during BeautifulSoup4 installation: {e}")
        sys.exit(1)

# Import local modules
from src.config_manager import load_config, save_config # Import save_config
from src.gui_elements import TextRedirector, print_art # Corrected import from print_logo to print_art
from src.instrument_logic import (
    populate_resources_logic, connect_instrument_logic, disconnect_instrument_logic,
    apply_settings_to_device_logic, update_preset_buttons, load_selected_preset_logic, # Changed update_preset_tree to update_preset_buttons
    set_focus_frequency_logic, set_marker_and_trace_modes_logic
)
from src.scan_logic import (
    start_scan_thread_logic, toggle_pause_scan_logic, run_scan_logic,
    stop_scan_logic, reset_scan_buttons_logic, pause_resume_scan_logic
)
from src.settings_logic import (
    restore_default_settings_logic, open_output_folder_logic,
    open_preset_folder_logic, open_report_folder_logic,
    update_scan_rbw_from_slider_index_logic, # Import slider update logic
    update_max_hold_time_from_slider_index_logic, # Import slider update logic
    update_cycle_wait_time_from_slider_index_logic # Import slider update logic
)
from src.plot_logic import generate_single_scan_plot_and_open_wrapper_logic, generate_average_plot_logic # Import generate_average_plot_logic
from utils.frequency_bands import SCAN_BAND_RANGES, MHZ_TO_HZ, VBW_RBW_RATIO
from utils.instrument_control import set_debug_mode, debug_print, set_log_visa_commands_mode # Import debug_print and set_log_visa_commands_mode
from src.tabs import ReportConverterTab, MarkersDisplayTab # Import MarkersDisplayTab


class App(tk.Tk):
    """
    Main application class for the RF Spectrum Analyzer Controller.
    This class inherits from Tkinter's Tk class and manages the entire GUI,
    instrument communication, scan processes, and data visualization.
    """
    CONFIG_FILE = 'config.ini' # Define config file name here

    def _initialize_tkinter_vars(self):
        """
        Initializes Tkinter control variables and maps them to configuration keys.
        This method sets up the dynamic relationship between GUI elements and
        application settings, facilitating easy loading and saving.
        """
        # General Settings
        self.resource_var = tk.StringVar(self)
        self.resource_var.set("No resources found")
        self.gpib_device_var = tk.StringVar(self)
        self.scan_directory_var = tk.StringVar(self)
        self.scan_name_var = tk.StringVar(self)
        self.open_html_after_complete_var = tk.BooleanVar(self)
        
        # Separated Debugging Variables
        self.general_debug_enabled_var = tk.BooleanVar(self) # For general debug messages
        self.log_visa_commands_enabled_var = tk.BooleanVar(self) # For specific VISA command logging

        self.selected_bands_str_var = tk.StringVar(self) # Added for selected bands string

        # Scan Configuration Settings
        self.desired_rbw_var = tk.StringVar(self) # Resolution Bandwidth
        self.desired_vbw_display_var = tk.StringVar(self) # Video Bandwidth (display only, calculated)
        self.desired_max_hold_time_var = tk.StringVar(self) # Max Hold Time
        self.desired_cycle_wait_time_var = tk.StringVar(self) # Cycle Wait Time
        self.desired_reference_level_var = tk.StringVar(self) # Reference Level
        self.desired_freq_shift_var = tk.StringVar(self) # Frequency Shift
        self.desired_maxhold_enabled_var = tk.BooleanVar(self) # Max Hold Enabled
        self.desired_include_gov_markers_var = tk.BooleanVar(self) # Include Gov Markers
        self.desired_include_tv_markers_var = tk.BooleanVar(self) # Include TV Markers
        self.desired_high_sensitivity_var = tk.BooleanVar(self) # High Sensitivity
        self.desired_preamp_on_var = tk.BooleanVar(self) # Preamp On
        self.desired_scan_rbw_segmentation_var = tk.StringVar(self) # New: Scan RBW for segmentation
        self.desired_default_focus_width_var = tk.StringVar(self) # New: Default Focus Width

        # Map setting names to (last_used_key, default_key, tk_var) tuples
        # last_used_key: Key in [LAST_USED_SETTINGS]
        # default_key: Key in [DEFAULT_SETTINGS]
        # tk_var: The Tkinter variable instance
        self.setting_var_map = {
            'gpib_device_var': ('last_gpib_device', 'default_gpib_device', self.gpib_device_var),
            'scan_directory_var': ('last_scan_directory', 'default_scan_directory', self.scan_directory_var),
            'scan_name_var': ('last_scan_name', 'default_scan_name', self.scan_name_var),
            'open_html_after_complete_var': ('last_open_html_after_complete', 'default_open_html_after_complete', self.open_html_after_complete_var),
            'general_debug_enabled_var': ('last_general_debug_enabled', 'default_general_debug_enabled', self.general_debug_enabled_var), # Updated
            'log_visa_commands_enabled_var': ('last_log_visa_commands_enabled', 'default_log_visa_commands_enabled', self.log_visa_commands_enabled_var), # New
            'selected_bands_str_var': ('last_selected_bands', 'default_selected_bands', self.selected_bands_str_var), # Added to map
            'desired_rbw_var': ('last_scan_rbw_hz', 'default_scan_rbw_hz', self.desired_rbw_var), # This is RBW for instrument
            'desired_max_hold_time_var': ('last_maxhold_time_seconds', 'default_maxhold_time_seconds', self.desired_max_hold_time_var),
            'desired_cycle_wait_time_var': ('last_cycle_wait_time_seconds', 'default_cycle_wait_time_seconds', self.desired_cycle_wait_time_var),
            'desired_reference_level_var': ('last_reference_level_dbm', 'default_reference_level_dbm', self.desired_reference_level_var),
            'desired_freq_shift_var': ('last_freq_shift_hz', 'default_freq_shift_hz', self.desired_freq_shift_var),
            'desired_maxhold_enabled_var': ('last_maxhold_enabled', 'default_maxhold_enabled', self.desired_maxhold_enabled_var),
            'desired_include_gov_markers_var': ('last_include_gov_markers', 'default_include_gov_markers', self.desired_include_gov_markers_var),
            'desired_include_tv_markers_var': ('last_include_tv_markers', 'default_include_tv_markers', self.desired_include_tv_markers_var),
            'desired_high_sensitivity_var': ('last_high_sensitivity', 'default_high_sensitivity', self.desired_high_sensitivity_var),
            'desired_preamp_on_var': ('last_preamp_on', 'default_preamp_on', self.desired_preamp_on_var),
            'desired_scan_rbw_segmentation_var': ('last_scan_rbw_segmentation', 'default_scan_rbw_segmentation', self.desired_scan_rbw_segmentation_var), # New mapping
            'desired_default_focus_width_var': ('last_default_focus_width', 'default_default_focus_width', self.desired_default_focus_width_var) # New mapping
        }

        # Initialize band selection variables
        self.band_vars = []
        for band in self.SCAN_BAND_RANGES:
            var = tk.BooleanVar(self)
            var.set(True) # Default to all bands selected
            self.band_vars.append({"band": band, "var": var})

        # Trace Tkinter variables to update GUI elements and save config
        # RBW, Max Hold Time, Cycle Wait Time are now updated by sliders, so their direct entry traces are removed here
        self.desired_reference_level_var.trace_add("write", self._update_setting_color_callback)
        self.desired_freq_shift_var.trace_add("write", self._update_setting_color_callback)
        self.desired_maxhold_enabled_var.trace_add("write", self._update_setting_color_callback)
        self.desired_high_sensitivity_var.trace_add("write", self._update_setting_color_callback)
        self.desired_preamp_on_var.trace_add("write", self._update_setting_color_callback)
        self.desired_scan_rbw_segmentation_var.trace_add("write", self._update_setting_color_callback)
        self.desired_default_focus_width_var.trace_add("write", self._update_setting_color_callback)
        self.scan_directory_var.trace_add("write", self._update_setting_color_callback)
        self.scan_name_var.trace_add("write", self._update_setting_color_callback)
        self.open_html_after_complete_var.trace_add("write", self._update_setting_color_callback)
        
        # Debug mode traces for the new separate checkboxes
        self.general_debug_enabled_var.trace_add("write", self._update_general_debug_callback)
        self.log_visa_commands_enabled_var.trace_add("write", self._update_log_visa_commands_callback)


    def __init__(self):
        """
        Initializes the main application window and its components.
        Sets up the GUI layout, loads configuration, and initializes
        instrument-related variables.
        """
        super().__init__()
        self.title("RF Spectrum Analyzer Controller - V1.0.0")
        self.geometry("1400x780+100+100") # Default size and position
        
        # Configuration and Instrument Management
        self.config = configparser.ConfigParser()
        self.rm = None # PyVISA Resource Manager, initialized to None
        self.inst = None # Connected instrument
        self.instrument_list = []
        self.instrument_model = "Unknown" # To store detected instrument model (e.g., N9340B, N9342CN)

        # Attempt to initialize PyVISA Resource Manager
        try:
            self.rm = pyvisa.ResourceManager()
            print("✅ PyVISA Resource Manager initialized successfully.")
        except Exception as e:
            print(f"❌ Error initializing PyVISA Resource Manager: {e}")
            print("Please ensure NI-VISA or Keysight VISA is installed and configured correctly.")
            self.rm = None # Ensure rm is None if initialization fails

        # Scan and Plotting Data
        self.scanning = False
        self.paused = False
        self.scan_thread = None
        self.scan_cycle_count = 0
        self.current_freq_offset = 0.0 # Initialize frequency offset
        self.collected_scans_dataframes = [] # To store pandas DataFrames of collected scan data
        self.last_scanned_band_index = 0 # For resuming scans
        self.csv_filename_current_cycle = None # To store the filename of the current cycle's raw data CSV

        # Constants
        self.SCAN_BAND_RANGES = SCAN_BAND_RANGES # Frequency bands for scanning
        self.MHZ_TO_HZ = MHZ_TO_HZ
        self.VBW_RBW_RATIO = VBW_RBW_RATIO

        # Slider values and their corresponding Tkinter variables
        self.rbw_values = [5000, 10000, 25000, 50000, 100000, 250000, 500000, 1000000]
        self.rbw_val_to_idx = {val: i for i, val in enumerate(self.rbw_values)}
        self.max_hold_time_values = [0, 1, 3, 5, 10, 30, 60, 300, 600] # Example values
        self.max_hold_time_val_to_idx = {val: i for i, val in enumerate(self.max_hold_time_values)}
        self.cycle_wait_time_values = [0, 0.1, 0.5, 1, 2, 5, 10] # Example values
        self.cycle_wait_time_val_to_idx = {val: i for i, val in enumerate(self.cycle_wait_time_values)}


        # Tkinter Variables for Settings
        self.setting_var_map = {} # Map setting names to their Tkinter variables and config keys

        # Initialize Tkinter variables and load config
        self._initialize_tkinter_vars()
        load_config(self) # Load settings from config.ini

        # Initialize slider index variables after config load
        # Ensure default value is within range, or handle if not found
        # Convert string values from config to appropriate types for lookup
        current_rbw = int(float(self.desired_rbw_var.get()))
        current_max_hold = int(float(self.desired_max_hold_time_var.get()))
        current_cycle_wait = float(self.desired_cycle_wait_time_var.get())

        self.rbw_slider_index_var = tk.IntVar(self, value=self.rbw_val_to_idx.get(current_rbw, 0))
        self.max_hold_time_slider_index_var = tk.IntVar(self, value=self.max_hold_time_val_to_idx.get(current_max_hold, 0))
        # For float values in slider, direct lookup might fail due to precision. Find closest.
        closest_cycle_wait_idx = min(range(len(self.cycle_wait_time_values)), 
                                     key=lambda i: abs(self.cycle_wait_time_values[i] - current_cycle_wait))
        self.cycle_wait_time_slider_index_var = tk.IntVar(self, value=closest_cycle_wait_idx)

        # Trace slider index variables to update their corresponding setting variables
        self.rbw_slider_index_var.trace_add("write", lambda *args: update_scan_rbw_from_slider_index_logic(self, *args))
        self.max_hold_time_slider_index_var.trace_add("write", lambda *args: update_max_hold_time_from_slider_index_logic(self, *args))
        self.cycle_wait_time_slider_index_var.trace_add("write", lambda *args: update_cycle_wait_time_from_slider_index_logic(self, *args))


        # --- Debugging Config Load ---
        print("\n--- Debugging Config Load ---")
        debug_print(f"Scan Directory (after load): {self.scan_directory_var.get()}", file=__file__, function=inspect.currentframe().f_code.co_name)
        debug_print(f"Desired RBW (after load): {self.desired_rbw_var.get()}", file=__file__, function=inspect.currentframe().f_code.co_name)
        debug_print(f"General Debug Enabled (after load): {self.general_debug_enabled_var.get()}", file=__file__, function=inspect.currentframe().f_code.co_name) # Updated
        debug_print(f"Log VISA Commands Enabled (after load): {self.log_visa_commands_enabled_var.get()}", file=__file__, function=inspect.currentframe().f_code.co_name) # New
        debug_print(f"Max Hold Time (after load): {self.desired_max_hold_time_var.get()}", file=__file__, function=inspect.currentframe().f_code.co_name)
        print("-----------------------------\n")
        # --- End Debugging Config Load ---

        # Set up GUI
        self._create_widgets()
        self._redirect_stdout_to_console()
        print_art() # Corrected call from print_logo to print_art

        # Initial population of VISA resources on load
        populate_resources_logic(self)

        # Initialize MarkersDisplayTab to None; it will be created dynamically
        self.markers_display_tab = None

        # List to keep track of dynamically created tabs
        self.dynamic_tabs = []

        # Check for existing MARKERS.CSV and load it at startup
        # Moved this call to ensure all styles are fully set up before MarkersDisplayTab is initialized
        self._check_and_load_markers_csv()

        # Handle window close event - moved after on_closing definition
        self.protocol("WM_DELETE_WINDOW", self.on_closing)


    def on_closing(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Handles the application's closing event.
        Ensures the configuration is saved and the instrument is disconnected.
        """
        debug_print("Closing application...", file=file, function=function)
        print("Closing application...")
        save_config(self) # Save current settings before closing
        if self.inst:
            disconnect_instrument_logic(self)
        self.destroy() # Close the Tkinter window
        sys.exit(0) # Ensure the application fully exits

    def _reset_gui_on_disconnect_or_error(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Resets GUI elements to a disconnected state.
        Called when the instrument disconnects or an error occurs.
        """
        debug_print("Resetting GUI on disconnect or error...", file=file, function=function)
        self.inst = None
        self.instrument_model = "Unknown"
        self.connect_button.config(state=tk.NORMAL)
        self.disconnect_button.config(state=tk.DISABLED)
        self.start_scan_button.config(state=tk.DISABLED, style='Green.TButton') # Reset style
        self.stop_scan_button.config(state=tk.NORMAL) # Should be normal to allow stopping a hung scan
        self.pause_resume_button.config(state=tk.DISABLED, style='Orange.TButton') # Reset style
        self.apply_button.config(state=tk.DISABLED)
        self.query_presets_button.config(state=tk.DISABLED)
        # self.load_preset_button.config(state=tk.DISABLED) # Removed as buttons replace it
        self.plot_button.config(state=tk.DISABLED) # Disable plot button on refresh
        self._start_connect_button_blink() # Start blinking connect button
        self._stop_pause_button_blink() # Ensure pause button blinking stops
        self.reset_setting_colors_logic() # Reset colors as settings are no longer 'applied'

    # Blinking animation for connect/pause buttons
    def _start_connect_button_blink(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        debug_print("Starting connect button blink...", file=file, function=function)
        if not hasattr(self, '_blink_connect_id'):
            self._blink_connect_state = False
            self._blink_connect_id = self.after(500, self._blink_connect_button)
    
    def _stop_connect_button_blink(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        debug_print("Stopping connect button blink...", file=file, function=function)
        if hasattr(self, '_blink_connect_id'):
            self.after_cancel(self._blink_connect_id)
            del self._blink_connect_id
        self.connect_button.config(style='GreyText.TButton') # Reset to default grey style

    def _blink_connect_button(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        
        if self.inst: # Stop blinking if connected
            self._stop_connect_button_blink()
            return
        
        if self._blink_connect_state:
            self.connect_button.config(style='GreyText.TButton') # Grey with black text
        else:
            self.connect_button.config(style='OrangeRed.TButton') # Orange with red text
        self._blink_connect_state = not self._blink_connect_state
        self._blink_connect_id = self.after(500, self._blink_connect_button)

    def _start_pause_button_blink(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        debug_print("Starting pause button blink...", file=file, function=function)
        if not hasattr(self, '_blink_pause_id'):
            self._blink_pause_state = False
            self._blink_pause_id = self.after(500, self._blink_pause_button)

    def _stop_pause_button_blink(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        debug_print("Stopping pause button blink...", file=file, function=function)
        if hasattr(self, '_blink_pause_id'):
            self.after_cancel(self._blink_pause_id)
            del self._blink_pause_id
        self.pause_resume_button.config(style='Orange.TButton') # Reset to default style

    def _blink_pause_button(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        debug_print("Blinking pause button...", file=file, function=function)
        if not self.paused: # Stop blinking if not paused
            self._stop_pause_button_blink()
            return
        
        if self._blink_pause_state:
            self.pause_resume_button.config(style='Orange.TButton') # Default style
        else:
            self.pause_resume_button.config(style='Yellow.TButton') # Custom yellow style
        self._blink_pause_state = not self._blink_pause_state
        self._blink_pause_id = self.after(500, self._blink_pause_button)

    def _update_setting_color_callback(self, *args, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Callback function for Tkinter variable traces to change the background color
        of an Entry widget if its value differs from the instrument's current setting.
        This visually indicates unapplied settings.
        """
        debug_print("Updating setting color callback...", file=file, function=function)
        # This function needs to be able to identify which Entry widget corresponds
        # to the Tkinter variable that triggered the callback.
        # A more robust solution involves passing the Entry widget reference
        # or having a map from Tkinter var to Entry widget.
        # For now, this is a placeholder and would require more context to implement fully.
        # It's primarily here to prevent the AttributeError.
        pass

    def _update_general_debug_callback(self, *args, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Callback for the general debug mode checkbox.
        Updates the global debug mode setting.
        """
        debug_print("Updating general debug callback...", file=file, function=function)
        set_debug_mode(self.general_debug_enabled_var.get())
        debug_print(f"General Debug Mode set to: {self.general_debug_enabled_var.get()}", file=file, function=function)

    def _update_log_visa_commands_callback(self, *args, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Callback for the log VISA commands checkbox.
        Updates the global VISA command logging setting.
        """
        debug_print("Updating log VISA commands callback...", file=file, function=function)
        set_log_visa_commands_mode(self.log_visa_commands_enabled_var.get())
        debug_print(f"Log VISA Commands Mode set to: {self.log_visa_commands_enabled_var.get()}", file=file, function=function)

    def reset_setting_colors_logic(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Resets the background color of all setting entry widgets to their default,
        indicating that all settings are either applied or no longer need visual flagging.
        """
        debug_print("Resetting setting colors logic initiated.", file=file, function=function)
        # This function would need access to the Entry widgets to change their background.
        # A proper implementation would involve storing references to these widgets
        # in the self.setting_var_map or a similar structure.
        # For now, it's a placeholder to prevent errors.
        pass

    def _create_widgets(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Creates and arranges all GUI widgets within the application window.
        This includes notebook tabs, frames for settings, buttons, and console output.
        """
        debug_print("Creating GUI widgets...", file=file, function=function)
        # Configure ttk.Style for consistent widget appearance and dark theme
        style = ttk.Style(self)
        # Removed style.theme_use('clam') to prevent overriding with a light theme.

        # General dark theme configurations
        style.configure(".", background="#000000", foreground="white") # Default for all ttk widgets
        style.configure("TFrame", background="#000000")
        style.configure("TLabelframe", background="#000000", foreground="white", bordercolor="#4a4a4a")
        style.configure("TLabelframe.Label", background="#000000", foreground="white")
        style.configure("TLabel", background="#000000", foreground="white")
        
        # Textbox styling: Dark grey background, BLACK text, white insert cursor
        style.configure("TEntry", fieldbackground="#4a4a4a", foreground="black", insertbackground="white")
        
        style.configure("TCheckbutton", background="#000000", foreground="white", indicatoron=True)
        style.map("TCheckbutton", background=[('active', '#000000')], foreground=[('disabled', '#888888')])

        # Notebook (Tabs) styling
        style.configure("TNotebook", background="#000000", borderwidth=0)
        # Unselected tabs: Dark grey background, BLACK text
        style.configure("TNotebook.Tab", background="#3a3a3a", foreground="black", # Changed foreground to black
                        lightcolor="#3a3a3a", darkcolor="#3a3a3a", borderwidth=0, padding=[5, 2])
        # Selected tabs: Orange background, BLACK text
        style.map("TNotebook.Tab", background=[("selected", "orange")],
                                   foreground=[("selected", "black")]) # Changed foreground to black for selected tab

        # Button styling
        style.configure("TButton", background="#3a3a3a", foreground="white",
                        font=('TkDefaultFont', 9, 'bold'), borderwidth=1, relief="raised",
                        focuscolor="#6a6a6a") # Focus color for buttons
        style.map("TButton",
                  background=[('active', '#6a6a6a'), ('disabled', '#2a2a2a')],
                  foreground=[('disabled', '#888888')])

        # Custom styles for specific buttons
        style.configure("Green.TButton", background="green", foreground="black")
        style.map("Green.TButton", background=[('active', '#006400')]) # Darker green on active
        style.configure("Orange.TButton", background="orange", foreground="black")
        style.map("Orange.TButton", background=[('active', '#CC8400')]) # Darker orange on active
        style.configure("Red.TButton", background="red", foreground="black")
        style.map("Red.TButton", background=[('active', '#CC0000')]) # Darker red on active

        # New style for grey buttons with black text
        style.configure("GreyText.TButton", background="#4a4a4a", foreground="black")
        style.map("GreyText.TButton",
                  background=[('active', '#6a6a6a'), ('disabled', '#2a2a2a')],
                  foreground=[('disabled', '#888888')])

        # New style for blinking connect button (Orange background, Red text)
        style.configure("OrangeRed.TButton", background="orange", foreground="red")
        style.map("OrangeRed.TButton", background=[('active', '#CC8400')])


        # Treeview styling
        style.configure("Treeview",
                        background="#4a4a4a",
                        foreground="white",
                        fieldbackground="#4a4a4a",
                        bordercolor="#2e2e2e",
                        lightcolor="#4a4a4a",
                        darkcolor="#4a4a4a")
        style.map("Treeview",
                  background=[("selected", "#0078D7")], # Windows default blue selection
                  foreground=[("selected", "white")])
        style.configure("Treeview.Heading",
                        background="#3a3a3a",
                        foreground="white",
                        font=('TkDefaultFont', 10, 'bold'))

        # Scrollbar styling
        style.configure("Vertical.TScrollbar",
                        background="#4a4a4a",
                        troughcolor="#2e2e2e",
                        bordercolor="#2e2e2e",
                        arrowcolor="white")
        style.map("Vertical.TScrollbar",
                  background=[('active', '#6a6a6a')])

        # Scale (Slider) styling
        style.configure("TScale", background="#000000", troughcolor="#3a3a3a", sliderrelief="flat")
        style.map("TScale", background=[('active', '#6a6a6a')])

        # --- MarkersDisplayTab specific styles (moved from tabs.py) ---
        style.configure("Markers.TFrame", background="#000000")
        style.configure("Markers.TLabel", background="#000000", foreground="white")
        style.configure("Markers.Treeview.Heading", background="#3a3a3a", foreground="white")
        style.configure("Markers.Treeview", background="#4a4a4a", foreground="white", fieldbackground="#4a4a4a")
        style.map("Markers.Treeview", background=[("selected", "#0078D7")], foreground=[("selected", "white")])
        style.configure("Markers.TButton", background="#3a3a3a", foreground="white")
        style.map("Markers.TButton", background=[('active', '#6a6a6a')])
        
        # New styles for the inner treeview and buttons frame - Corrected style name for TLabelframe
        style.configure("Markers.Inner.TLabelframe", # Changed from Markers.Inner.LabelFrame
                        background="#333333", # Darker grey
                        foreground="white")
        style.configure("Markers.Inner.Treeview",
                        background="#333333", # Darker grey
                        foreground="white",
                        fieldbackground="#333333", # Darker grey
                        bordercolor="black",
                        lightcolor="#333333", # Darker grey
                        darkcolor="#333333") # Darker grey
        style.map("Markers.Inner.Treeview",
                  background=[("selected", "#F4902C")], # New highlight color
                  foreground=[("selected", "white")])
        style.configure("Markers.Inner.Frame", background="#333333")
        # --- End MarkersDisplayTab specific styles ---


        # Root window background
        self.configure(bg="#000000")

        # Main frames for 50/50 split (using grid)
        self.grid_columnconfigure(0, weight=1) # Left half for main controls
        self.grid_columnconfigure(1, weight=1) # Right half for console
        self.grid_rowconfigure(0, weight=1) # Only one row

        self.main_frame = ttk.Frame(self)
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        self.console_frame = ttk.Frame(self) # Parent is self (root window)
        self.console_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        
        # Configure console_frame's internal grid rows/columns
        self.console_frame.grid_rowconfigure(1, weight=1) # Make console_output expand vertically
        self.console_frame.grid_columnconfigure(0, weight=1) # Make content expand horizontally

        # Scan Control Frame (moved here, above console output)
        self.scan_control_frame = ttk.LabelFrame(self.console_frame, text="Scan Control", padding="10")
        self.scan_control_frame.grid(row=0, column=0, sticky="ew", pady=5)
        self.scan_control_frame.grid_columnconfigure(0, weight=1)
        self.scan_control_frame.grid_columnconfigure(1, weight=1)
        self.scan_control_frame.grid_columnconfigure(2, weight=1)

        # Buttons in scan control frame (using ttk.Button with custom styles)
        self.start_scan_button = ttk.Button(self.scan_control_frame, text="Start Scan", command=lambda: start_scan_thread_logic(self), state=tk.DISABLED, style='Green.TButton')
        self.start_scan_button.grid(row=0, column=0, padx=5, pady=2, sticky=tk.EW)
        self.pause_resume_button = ttk.Button(self.scan_control_frame, text="Pause Scan", command=lambda: toggle_pause_scan_logic(self), state=tk.DISABLED, style='Orange.TButton')
        self.pause_resume_button.grid(row=0, column=1, padx=5, pady=2, sticky=tk.EW)
        self.stop_scan_button = ttk.Button(self.scan_control_frame, text="Stop Scan", command=lambda: self.stop_scan(), state=tk.DISABLED, style='Red.TButton')
        self.stop_scan_button.grid(row=0, column=2, padx=5, pady=2, sticky=tk.EW)
        
        # Plot button also moved here
        self.plot_button = ttk.Button(self.scan_control_frame, text="Generate Average Plot", command=lambda: generate_average_plot_logic(self), state=tk.DISABLED, style='GreyText.TButton') # Applied new style
        self.plot_button.grid(row=1, column=0, columnspan=3, padx=5, pady=2, sticky=tk.EW) # Adjusted row and columnspan

        # Console output frame (now directly in console_frame, below scan_control_frame)
        self.console_output = scrolledtext.ScrolledText(self.console_frame, wrap=tk.WORD, bg="black", fg="white", font=("Courier New", 10))
        self.console_output.grid(row=1, column=0, sticky="nsew") # Placed in row 1 of console_frame (was row 2)
        self.console_output.config(state=tk.DISABLED) # Make it read-only

        # Configure tags for console output
        self.console_output.tag_config("green", foreground="green")
        self.console_output.tag_config("red", foreground="red")
        self.console_output.tag_config("yellow", foreground="yellow")
        self.console_output.tag_config("white", foreground="white") # Default output to white
        self.console_output.tag_config("debug_white", foreground="white") # Tag for debug_print output

        # Create a notebook (tabbed interface) - now gridded within main_frame
        self.notebook = ttk.Notebook(self.main_frame) # Parent is main_frame
        self.notebook.grid(row=0, column=0, sticky="nsew") # Using grid for main_frame's children. Only one column.

        # Create frames for each tab (now ttk.Frame)
        self.scan_settings_tab = ttk.Frame(self.notebook)
        self.preset_files_tab = ttk.Frame(self.notebook)
        self.report_converter_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.scan_settings_tab, text="Scan Configuration")
        self.notebook.add(self.preset_files_tab, text="Device Preset Files")
        self.notebook.add(self.report_converter_tab, text="Report Converter")

        # Store dynamic tabs for later management
        self.dynamic_tabs = [self.scan_settings_tab, self.preset_files_tab, self.report_converter_tab]


        # --- Scan Configuration Tab (self.scan_settings_tab) ---
        self._create_scan_settings_widgets(self.scan_settings_tab)

        # --- Device Preset Files Tab (self.preset_files_tab) ---
        self._create_preset_files_widgets(self.preset_files_tab)

        # --- Report Converter Tab (self.report_converter_tab) ---
        # The ReportConverterTab class itself needs to be updated to use ttk widgets and dark theme
        self.report_converter_frame = ReportConverterTab(self.report_converter_tab, app_instance=self)
        self.report_converter_frame.pack(expand=True, fill="both", padx=10, pady=10)

        # Configure grid weights for main_frame's internal layout
        self.main_frame.grid_columnconfigure(0, weight=1) # Only one column
        self.main_frame.grid_rowconfigure(0, weight=1) # Notebook expands vertically


    def on_closing(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Handles the application's closing event.
        Ensures the configuration is saved and the instrument is disconnected.
        """
        debug_print("Closing application...", file=file, function=function)
        print("Closing application...")
        save_config(self) # Save current settings before closing
        if self.inst:
            disconnect_instrument_logic(self)
        self.destroy() # Close the Tkinter window
        sys.exit(0) # Ensure the application fully exits

    def _reset_gui_on_disconnect_or_error(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Resets GUI elements to a disconnected state.
        Called when the instrument disconnects or an error occurs.
        """
        debug_print("Resetting GUI on disconnect or error...", file=file, function=function)
        self.inst = None
        self.instrument_model = "Unknown"
        self.connect_button.config(state=tk.NORMAL)
        self.disconnect_button.config(state=tk.DISABLED)
        self.start_scan_button.config(state=tk.DISABLED, style='Green.TButton') # Reset style
        self.stop_scan_button.config(state=tk.NORMAL) # Should be normal to allow stopping a hung scan
        self.pause_resume_button.config(state=tk.DISABLED, style='Orange.TButton') # Reset style
        self.apply_button.config(state=tk.DISABLED)
        self.query_presets_button.config(state=tk.DISABLED)
        # self.load_preset_button.config(state=tk.DISABLED) # Removed as buttons replace it
        self.plot_button.config(state=tk.DISABLED) # Disable plot button on refresh
        self._start_connect_button_blink() # Start blinking connect button
        self._stop_pause_button_blink() # Ensure pause button blinking stops
        self.reset_setting_colors_logic() # Reset colors as settings are no longer 'applied'

    # Blinking animation for connect/pause buttons
    def _start_connect_button_blink(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        debug_print("Starting connect button blink...", file=file, function=function)
        if not hasattr(self, '_blink_connect_id'):
            self._blink_connect_state = False
            self._blink_connect_id = self.after(500, self._blink_connect_button)
    
    def _stop_connect_button_blink(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        debug_print("Stopping connect button blink...", file=file, function=function)
        if hasattr(self, '_blink_connect_id'):
            self.after_cancel(self._blink_connect_id)
            del self._blink_connect_id
        self.connect_button.config(style='GreyText.TButton') # Reset to default grey style

    def _blink_connect_button(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        
        if self.inst: # Stop blinking if connected
            self._stop_connect_button_blink()
            return
        
        if self._blink_connect_state:
            self.connect_button.config(style='GreyText.TButton') # Grey with black text
        else:
            self.connect_button.config(style='OrangeRed.TButton') # Orange with red text
        self._blink_connect_state = not self._blink_connect_state
        self._blink_connect_id = self.after(500, self._blink_connect_button)

    def _start_pause_button_blink(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        debug_print("Starting pause button blink...", file=file, function=function)
        if not hasattr(self, '_blink_pause_id'):
            self._blink_pause_state = False
            self._blink_pause_id = self.after(500, self._blink_pause_button)

    def _stop_pause_button_blink(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        debug_print("Stopping pause button blink...", file=file, function=function)
        if hasattr(self, '_blink_pause_id'):
            self.after_cancel(self._blink_pause_id)
            del self._blink_pause_id
        self.pause_resume_button.config(style='Orange.TButton') # Reset to default style

    def _blink_pause_button(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        debug_print("Blinking pause button...", file=file, function=function)
        if not self.paused: # Stop blinking if not paused
            self._stop_pause_button_blink()
            return
        
        if self._blink_pause_state:
            self.pause_resume_button.config(style='Orange.TButton') # Default style
        else:
            self.pause_resume_button.config(style='Yellow.TButton') # Custom yellow style
        self._blink_pause_state = not self._blink_pause_state
        self._blink_pause_id = self.after(500, self._blink_pause_button)

    def _update_setting_color_callback(self, *args, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Callback function for Tkinter variable traces to change the background color
        of an Entry widget if its value differs from the instrument's current setting.
        This visually indicates unapplied settings.
        """
        debug_print("Updating setting color callback...", file=file, function=function)
        # This function needs to be able to identify which Entry widget corresponds
        # to the Tkinter variable that triggered the callback.
        # A more robust solution involves passing the Entry widget reference
        # or having a map from Tkinter var to Entry widget.
        # For now, this is a placeholder and would require more context to implement fully.
        # It's primarily here to prevent the AttributeError.
        pass

    def _update_general_debug_callback(self, *args, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Callback for the general debug mode checkbox.
        Updates the global debug mode setting.
        """
        debug_print("Updating general debug callback...", file=file, function=function)
        set_debug_mode(self.general_debug_enabled_var.get())
        debug_print(f"General Debug Mode set to: {self.general_debug_enabled_var.get()}", file=file, function=function)

    def _update_log_visa_commands_callback(self, *args, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Callback for the log VISA commands checkbox.
        Updates the global VISA command logging setting.
        """
        debug_print("Updating log VISA commands callback...", file=file, function=function)
        set_log_visa_commands_mode(self.log_visa_commands_enabled_var.get())
        debug_print(f"Log VISA Commands Mode set to: {self.log_visa_commands_enabled_var.get()}", file=file, function=function)

    def reset_setting_colors_logic(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Resets the background color of all setting entry widgets to their default,
        indicating that all settings are either applied or no longer need visual flagging.
        """
        debug_print("Resetting setting colors logic initiated.", file=file, function=function)
        # This function would need access to the Entry widgets to change their background.
        # A proper implementation would involve storing references to these widgets
        # in the self.setting_var_map or a similar structure.
        # For now, it's a placeholder to prevent errors.
        pass

    def _create_scan_settings_widgets(self, parent_frame, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Creates and arranges widgets for the 'Scan Configuration' tab.
        """
        debug_print("Creating scan settings widgets...", file=file, function=function)
        parent_frame.grid_columnconfigure(0, weight=1) # Make column 0 expandable
        parent_frame.grid_columnconfigure(1, weight=1) # Make column 1 expandable
        parent_frame.grid_rowconfigure(0, weight=0) # Row for instrument connection
        parent_frame.grid_rowconfigure(1, weight=1) # Row for main settings area

        # --- Instrument Connection Frame ---
        instrument_frame = ttk.LabelFrame(parent_frame, text="Instrument Connection", padding="10")
        instrument_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        instrument_frame.grid_columnconfigure(1, weight=1) # Make dropdown expandable

        ttk.Label(instrument_frame, text="VISA Resource:", style="TLabel").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.resource_dropdown = ttk.OptionMenu(instrument_frame, self.resource_var, "No resources found", *self.instrument_list)
        self.resource_dropdown.grid(row=0, column=1, padx=5, pady=2, sticky=tk.EW)
        
        self.connect_button = ttk.Button(instrument_frame, text="Connect", command=lambda: connect_instrument_logic(self), style='Green.TButton')
        self.connect_button.grid(row=0, column=2, padx=5, pady=2)
        self.disconnect_button = ttk.Button(instrument_frame, text="Disconnect", command=lambda: disconnect_instrument_logic(self), state=tk.DISABLED, style='Red.TButton')
        self.disconnect_button.grid(row=0, column=3, padx=5, pady=2)
        
        ttk.Button(instrument_frame, text="Refresh Resources", command=lambda: populate_resources_logic(self), style='GreyText.TButton').grid(row=1, column=0, columnspan=2, padx=5, pady=2, sticky=tk.EW)
        
        self.instrument_model_label = ttk.Label(instrument_frame, text="Model: Unknown", style="TLabel")
        self.instrument_model_label.grid(row=1, column=2, columnspan=2, padx=5, pady=2, sticky=tk.W)


        # --- Settings and Bands Frame (Main Content Area) ---
        settings_bands_frame = ttk.Frame(parent_frame)
        settings_bands_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)
        settings_bands_frame.grid_columnconfigure(0, weight=1) # Settings column
        settings_bands_frame.grid_columnconfigure(1, weight=1) # Bands column
        settings_bands_frame.grid_rowconfigure(0, weight=1) # Only one row for settings/bands

        # Left side: Scan Settings
        settings_frame = ttk.LabelFrame(settings_bands_frame, text="Scan Settings", padding="10")
        settings_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        settings_frame.grid_columnconfigure(1, weight=1) # Make entry fields expandable

        # Scan Directory
        ttk.Label(settings_frame, text="Scan Directory:", style="TLabel").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.scan_directory_entry = ttk.Entry(settings_frame, textvariable=self.scan_directory_var, width=30, style="TEntry")
        self.scan_directory_entry.grid(row=0, column=1, padx=5, pady=2, sticky=tk.EW)
        ttk.Button(settings_frame, text="Browse", command=lambda: self.scan_directory_var.set(filedialog.askdirectory() or self.scan_directory_var.get()), style='GreyText.TButton').grid(row=0, column=2, padx=5, pady=2)
        ttk.Button(settings_frame, text="Open Folder", command=lambda: open_output_folder_logic(self), style='GreyText.TButton').grid(row=0, column=3, padx=5, pady=2)

        # Scan Name
        ttk.Label(settings_frame, text="Scan Name:", style="TLabel").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.scan_name_entry = ttk.Entry(settings_frame, textvariable=self.scan_name_var, width=30, style="TEntry")
        self.scan_name_entry.grid(row=1, column=1, padx=5, pady=2, sticky=tk.EW)

        # Desired RBW (Read-only Entry + Slider)
        ttk.Label(settings_frame, text="Desired RBW (Hz):", style="TLabel").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.desired_rbw_entry = ttk.Entry(settings_frame, textvariable=self.desired_rbw_var, state='readonly', width=15, style="TEntry")
        self.desired_rbw_entry.grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)
        self.rbw_slider = ttk.Scale(settings_frame, from_=0, to=len(self.rbw_values) - 1, orient=tk.HORIZONTAL,
                                    variable=self.rbw_slider_index_var, command=lambda val: self.rbw_slider_index_var.set(round(float(val))))
        self.rbw_slider.grid(row=2, column=2, columnspan=2, padx=5, pady=2, sticky=tk.EW)


        # Desired VBW (Display Only)
        ttk.Label(settings_frame, text="Desired VBW (Hz):", style="TLabel").grid(row=3, column=0, padx=5, pady=2, sticky=tk.W)
        self.desired_vbw_display_label = ttk.Label(settings_frame, textvariable=self.desired_vbw_display_var, style="TLabel")
        self.desired_vbw_display_label.grid(row=3, column=1, padx=5, pady=2, sticky=tk.W)

        # Max Hold Time (Read-only Entry + Slider)
        ttk.Label(settings_frame, text="Max Hold Time (s):", style="TLabel").grid(row=4, column=0, padx=5, pady=2, sticky=tk.W)
        self.desired_max_hold_time_entry = ttk.Entry(settings_frame, textvariable=self.desired_max_hold_time_var, state='readonly', width=15, style="TEntry")
        self.desired_max_hold_time_entry.grid(row=4, column=1, padx=5, pady=2, sticky=tk.W)
        self.max_hold_time_slider = ttk.Scale(settings_frame, from_=0, to=len(self.max_hold_time_values) - 1, orient=tk.HORIZONTAL,
                                            variable=self.max_hold_time_slider_index_var, command=lambda val: self.max_hold_time_slider_index_var.set(round(float(val))))
        self.max_hold_time_slider.grid(row=4, column=2, columnspan=2, padx=5, pady=2, sticky=tk.EW)

        # Cycle Wait Time (Read-only Entry + Slider)
        ttk.Label(settings_frame, text="Cycle Wait Time (s):", style="TLabel").grid(row=5, column=0, padx=5, pady=2, sticky=tk.W)
        self.desired_cycle_wait_time_entry = ttk.Entry(settings_frame, textvariable=self.desired_cycle_wait_time_var, state='readonly', width=15, style="TEntry")
        self.desired_cycle_wait_time_entry.grid(row=5, column=1, padx=5, pady=2, sticky=tk.W)
        self.cycle_wait_time_slider = ttk.Scale(settings_frame, from_=0, to=len(self.cycle_wait_time_values) - 1, orient=tk.HORIZONTAL,
                                                variable=self.cycle_wait_time_slider_index_var, command=lambda val: self.cycle_wait_time_slider_index_var.set(round(float(val))))
        self.cycle_wait_time_slider.grid(row=5, column=2, columnspan=2, padx=5, pady=2, sticky=tk.EW)


        # Reference Level
        ttk.Label(settings_frame, text="Reference Level (dBm):", style="TLabel").grid(row=6, column=0, padx=5, pady=2, sticky=tk.W)
        self.desired_reference_level_entry = ttk.Entry(settings_frame, textvariable=self.desired_reference_level_var, width=15, style="TEntry")
        self.desired_reference_level_entry.grid(row=6, column=1, padx=5, pady=2, sticky=tk.W)

        # Frequency Shift
        ttk.Label(settings_frame, text="Frequency Shift (Hz):", style="TLabel").grid(row=7, column=0, padx=5, pady=2, sticky=tk.W)
        self.desired_freq_shift_entry = ttk.Entry(settings_frame, textvariable=self.desired_freq_shift_var, width=15, style="TEntry")
        self.desired_freq_shift_entry.grid(row=7, column=1, padx=5, pady=2, sticky=tk.EW)

        # Max Hold Enabled Checkbox
        ttk.Checkbutton(settings_frame, text="Max Hold Enabled", variable=self.desired_maxhold_enabled_var, style="TCheckbutton").grid(row=8, column=0, columnspan=2, padx=5, pady=2, sticky=tk.W)

        # High Sensitivity Checkbox
        ttk.Checkbutton(settings_frame, text="High Sensitivity", variable=self.desired_high_sensitivity_var, style="TCheckbutton").grid(row=9, column=0, columnspan=2, padx=5, pady=2, sticky=tk.W)

        # Preamp On Checkbox
        ttk.Checkbutton(settings_frame, text="Preamp On", variable=self.desired_preamp_on_var, style="TCheckbutton").grid(row=10, column=0, columnspan=2, padx=5, pady=2, sticky=tk.W)

        # Scan RBW Segmentation
        ttk.Label(settings_frame, text="Scan RBW Segmentation:", style="TLabel").grid(row=11, column=0, padx=5, pady=2, sticky=tk.W)
        self.desired_scan_rbw_segmentation_entry = ttk.Entry(settings_frame, textvariable=self.desired_scan_rbw_segmentation_var, width=15, style="TEntry")
        self.desired_scan_rbw_segmentation_entry.grid(row=11, column=1, padx=5, pady=2, sticky=tk.W)

        # Default Focus Width
        ttk.Label(settings_frame, text="Default Focus Width (Hz):", style="TLabel").grid(row=12, column=0, padx=5, pady=2, sticky=tk.W)
        self.desired_default_focus_width_entry = ttk.Entry(settings_frame, textvariable=self.desired_default_focus_width_var, width=15, style="TEntry")
        self.desired_default_focus_width_entry.grid(row=12, column=1, padx=5, pady=2, sticky=tk.W)

        # Debug Checkboxes
        ttk.Checkbutton(settings_frame, text="Enable General Debug", variable=self.general_debug_enabled_var, style="TCheckbutton").grid(row=13, column=0, columnspan=2, padx=5, pady=2, sticky=tk.W)
        ttk.Checkbutton(settings_frame, text="Log VISA Commands", variable=self.log_visa_commands_enabled_var, style="TCheckbutton").grid(row=14, column=0, columnspan=2, padx=5, pady=2, sticky=tk.W)
        
        # Open HTML After Complete Checkbox
        ttk.Checkbutton(settings_frame, text="Open HTML Plot After Scan", variable=self.open_html_after_complete_var, style="TCheckbutton").grid(row=15, column=0, columnspan=2, padx=5, pady=2, sticky=tk.W)

        # Apply Settings Button
        self.apply_button = ttk.Button(settings_frame, text="Apply Settings to Device", command=lambda: apply_settings_to_device_logic(self), state=tk.DISABLED, style='Orange.TButton')
        self.apply_button.grid(row=16, column=0, columnspan=4, padx=5, pady=10, sticky=tk.EW)

        # Restore Defaults Button
        ttk.Button(settings_frame, text="Restore Default Settings", command=lambda: restore_default_settings_logic(self), style='GreyText.TButton').grid(row=17, column=0, columnspan=4, padx=5, pady=2, sticky=tk.EW)


        # Right side: Frequency Bands Selection
        bands_frame = ttk.LabelFrame(settings_bands_frame, text="Frequency Bands", padding="10")
        bands_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        bands_frame.grid_columnconfigure(0, weight=1) # Make column for checkboxes expandable

        # Use a canvas with a scrollbar for bands if there are many
        bands_canvas = tk.Canvas(bands_frame, bg="#000000", highlightbackground="#000000") # tk.Canvas as ttk.Canvas doesn't exist
        bands_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        bands_scrollbar = ttk.Scrollbar(bands_frame, orient="vertical", command=bands_canvas.yview)
        bands_scrollbar.pack(side=tk.RIGHT, fill="y")

        bands_canvas.configure(yscrollcommand=bands_scrollbar.set)
        bands_canvas.bind('<Configure>', lambda e: bands_canvas.configure(scrollregion = bands_canvas.bbox("all")))

        self.inner_bands_frame = ttk.Frame(bands_canvas, style="TFrame")
        bands_canvas.create_window((0, 0), window=self.inner_bands_frame, anchor="nw")

        # Populate frequency band checkboxes
        for i, band_info in enumerate(self.band_vars):
            band_name = band_info["band"]["Band Name"]
            var = band_info["var"]
            cb = ttk.Checkbutton(self.inner_bands_frame, text=band_name, variable=var, style="TCheckbutton")
            cb.grid(row=i, column=0, sticky=tk.W, padx=2, pady=1)
            var.trace_add("write", self._update_setting_color_callback) # Trace changes to update apply button color

    def _create_preset_files_widgets(self, parent_frame, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Creates and arranges widgets for the 'Device Preset Files' tab.
        """
        debug_print("Creating preset files widgets...", file=file, function=inspect.currentframe().f_code.co_name)
        parent_frame.grid_columnconfigure(0, weight=1) # Make column expandable
        parent_frame.grid_rowconfigure(1, weight=1) # Make treeview row expandable

        # Preset Actions Frame
        preset_actions_frame = ttk.LabelFrame(parent_frame, text="Preset Actions", padding="10")
        preset_actions_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        preset_actions_frame.grid_columnconfigure(0, weight=1)
        preset_actions_frame.grid_columnconfigure(1, weight=1)

        self.query_presets_button = ttk.Button(preset_actions_frame, text="Query Device Presets", command=lambda: update_preset_buttons(self.inner_preset_buttons_frame), state=tk.DISABLED, style='GreyText.TButton') # Pass the frame
        self.query_presets_button.grid(row=0, column=0, padx=5, pady=5, sticky=tk.EW)
        
        ttk.Button(preset_actions_frame, text="Open Preset Folder", command=lambda: open_preset_folder_logic(self), style='GreyText.TButton').grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)

        # Preset Display Frame
        preset_display_frame = ttk.LabelFrame(parent_frame, text="Available Presets", padding="10")
        preset_display_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        preset_display_frame.grid_columnconfigure(0, weight=1)
        preset_display_frame.grid_rowconfigure(0, weight=1)

        # Use a canvas with a scrollbar for preset buttons
        preset_canvas = tk.Canvas(preset_display_frame, bg="#000000", highlightbackground="#000000")
        preset_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        preset_scrollbar = ttk.Scrollbar(preset_display_frame, orient="vertical", command=preset_canvas.yview)
        preset_scrollbar.pack(side=tk.RIGHT, fill="y")

        preset_canvas.configure(yscrollcommand=preset_scrollbar.set)
        preset_canvas.bind('<Configure>', lambda e: preset_canvas.configure(scrollregion = preset_canvas.bbox("all")))

        self.inner_preset_buttons_frame = ttk.Frame(preset_canvas, style="TFrame")
        preset_canvas.create_window((0, 0), window=self.inner_preset_buttons_frame, anchor="nw")

        self.inner_preset_buttons_frame.grid_columnconfigure(0, weight=1) # Make column expandable for buttons

        # Initial state: No presets loaded
        ttk.Label(self.inner_preset_buttons_frame, text="Click 'Query Device Presets' to load.", style="TLabel").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

    def _redirect_stdout_to_console(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Redirects stdout to the Tkinter scrolled text widget for console output.
        """
        debug_print("Redirecting stdout to console...", file=file, function=inspect.currentframe().f_code.co_name)
        sys.stdout = TextRedirector(self.console_output, "white") # Default tag for general output
        sys.stderr = TextRedirector(self.console_output, "red") # Error output in red

    def add_markers_tab(self, headers, rows, file=__file__, function=inspect.currentframe().f_code.co_name):
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
        debug_print("Adding/updating markers tab...", file=file, function=inspect.currentframe().f_code.co_name)
        # Check if "Markers Display" tab already exists and remove it
        for i, tab_id in enumerate(self.notebook.tabs()):
            tab_text = self.notebook.tab(tab_id, "text")
            if tab_text == "Markers Display":
                self.notebook.forget(tab_id)
                self.dynamic_tabs = [tab for tab in self.dynamic_tabs if tab != self.markers_display_tab] # Remove from dynamic tabs list
                del self.markers_display_tab # Delete the reference
                debug_print("Existing 'Markers Display' tab removed.", file=file, function=inspect.currentframe().f_code.co_name)
                break
            
        # Create and add the new MarkersDisplayTab
        self.markers_display_tab = MarkersDisplayTab(self.notebook, headers=headers, rows=rows, app_instance=self)
        self.notebook.add(self.markers_display_tab, text="Markers Display")
        self.dynamic_tabs.append(self.markers_display_tab) # Add to dynamic tabs list
        self.notebook.select(self.markers_display_tab)
        # No need for _update_tab_visibility() here, as notebook.add and notebook.select handle visibility.


    def _check_and_load_markers_csv(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Checks for the existence of 'MARKERS.CSV' in the default output directory
        and, if found, automatically loads its content and creates the "Markers Display" tab.
        """
        debug_print("Checking for existing MARKERS.CSV...", file=file, function=inspect.currentframe().f_code.co_name)
        # Use scan_directory_var for the output folder, which is initialized from config
        markers_csv_path = os.path.join(self.scan_directory_var.get(), 'MARKERS.CSV')
        
        if os.path.exists(markers_csv_path):
            print(f"Found existing MARKERS.CSV at: {markers_csv_path}. Attempting to load...")
            debug_print(f"Found existing MARKERS.CSV at: {markers_csv_path}. Attempting to load...", file=file, function=inspect.currentframe().f_code.co_name)
            headers = []
            rows = []
            try:
                with open(markers_csv_path, 'r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    headers = reader.fieldnames
                    for row in reader:
                        rows.append(row)
                
                if headers and rows:
                    # Debug print to check if style exists before creating the tab
                    # This is for debugging the "Layout Markers.Inner.LabelFrame not found" error
                    # if hasattr(self, 'style'): # Check if style object exists
                    #     print(f"DEBUG: Checking if style 'Markers.Inner.LabelFrame' exists: {self.style.lookup('Markers.Inner.LabelFrame', 'background') is not None}")
                    
                    self.add_markers_tab(headers, rows)
                    print("✅ MARKERS.CSV loaded successfully and 'Markers Display' tab created.")
                    debug_print("MARKERS.CSV loaded successfully and 'Markers Display' tab created.", file=file, function=inspect.currentframe().f_code.co_name)
                else:
                    print("ℹ️ MARKERS.CSV found but appears empty or malformed. Skipping tab creation.")
                    debug_print("MARKERS.CSV found but appears empty or malformed. Skipping tab creation.", file=file, function=inspect.currentframe().f_code.co_name)
            except Exception as e:
                print(f"❌ Error loading MARKERS.CSV: {e}")
                messagebox.showerror("Error Loading Markers", f"Failed to load MARKERS.CSV: {e}")
                debug_print(f"Error loading MARKERS.CSV: {e}", file=file, function=inspect.currentframe().f_code.co_name)
        else:
            print("ℹ️ No MARKERS.CSV found at startup. 'Markers Display' tab will not be automatically created.")
            debug_print("No MARKERS.CSV found at startup. 'Markers Display' tab will not be automatically created.", file=file, function=inspect.currentframe().f_code.co_name)


    def _run_scan(self, selected_bands, scan_rbw_segmentation, freq_shift_value, rbw_config_val, vbw_config_val, max_hold_time, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Wrapper for run_scan_logic to be called in a thread.
        """
        debug_print("Running scan logic in a thread...", file=file, function=inspect.currentframe().f_code.co_name)
        run_scan_logic(self, selected_bands, scan_rbw_segmentation, freq_shift_value, rbw_config_val, vbw_config_val, max_hold_time)

    def generate_single_scan_plot_and_open_wrapper(self, csv_file_path, output_html_path, auto_open_browser=True, single_marker_data=None, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Wrapper to call generate_single_scan_plot_and_open_wrapper_logic from the main thread.
        """
        debug_print("Generating single scan plot and opening wrapper...", file=file, function=inspect.currentframe().f_code.co_name)
        # Call the logic function with all parameters
        generate_single_scan_plot_and_open_wrapper_logic(
            self,
            csv_file_path,
            output_html_path,
            auto_open_browser,
            single_marker_data
        )

    def hide_all_dynamic_tabs(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Hides all dynamically created tabs from the notebook.
        """
        debug_print("Hiding all dynamic tabs...", file=file, function=inspect.currentframe().f_code.co_name)
        for tab_frame in self.dynamic_tabs:
            try:
                self.notebook.hide(tab_frame)
            except TclError as e:
                print(f"Warning: Could not hide tab {tab_frame} - {e}")
        print("Tabs updated: All dynamic tabs now hidden.")

    def show_all_dynamic_tabs(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Shows all dynamically created tabs in the notebook.
        This is useful for debugging or restoring a layout.
        """
        debug_print("Showing all dynamic tabs...", file=file, function=inspect.currentframe().f_code.co_name)
        # Re-add all dynamic tabs
        for tab_frame in self.dynamic_tabs:
            # Only add if not already present to avoid TclError
            if tab_frame not in self.notebook.tabs():
                # Re-add using its original text, assuming it was added with text initially
                # This is important if a tab was dynamically created (like Markers Display)
                # and then hidden. Its text property might still be available.
                
                # A more robust solution would be to store a mapping of tab_frame to tab_text
                # when the tabs are initially created. For now, we'll hardcode based on known tabs.
                if tab_frame == self.scan_settings_tab:
                    self.notebook.add(tab_frame, text="Scan Configuration")
                elif tab_frame == self.preset_files_tab:
                    self.notebook.add(tab_frame, text="Device Preset Files")
                elif tab_frame == self.report_converter_tab:
                    self.notebook.add(tab_frame, text="Report Converter")
                # Handle Markers Display Tab if it exists
                elif hasattr(self, 'markers_display_tab') and tab_frame == self.markers_display_tab:
                    self.notebook.add(tab_frame, text="Markers Display")
        print("Tabs updated: All tabs now visible.")

    def stop_scan(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Initiates the stopping of the scan.
        """
        debug_print("Stop scan requested.", file=file, function=inspect.currentframe().f_code.co_name)
        stop_scan_logic(self)


if __name__ == "__main__":
    app = App()
    app.mainloop()
