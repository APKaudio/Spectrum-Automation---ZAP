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
import time
import pyvisa
import configparser
import subprocess # For BeautifulSoup installation check
import inspect


# Import local modules
from src.config_manager import load_config, save_config
from src.gui_elements import TextRedirector, print_art
from src.instrument_logic import (
    populate_resources_logic, connect_instrument_logic, disconnect_instrument_logic,
    apply_settings_logic,
    query_current_instrument_settings_logic, query_device_presets_logic,
    load_selected_preset_logic
)
from src.scan_logic import update_connection_status_logic
from src.settings_logic import restore_default_settings_logic
from src.scan_controler_button_logic import ScanControlTab

from utils.instrument_control import set_debug_mode, set_log_visa_commands_mode, debug_print
from tabs.tab_instrument_preset import PresetFilesTab
from tabs.tab_marker_display import MarkersDisplayTab
from tabs.tab_plotting import PlottingTab
from tabs.tab_report_converter import ReportConverterTab
from tabs.tab_scan_configuration import ScanTab

from tabs.tab_instrument_connection import InstrumentTab
from tabs.tab_visa_interpreter import VisaInterpreterTab

# Import constants from frequency_bands.py
from ref.frequency_bands import SCAN_BAND_RANGES, MHZ_TO_HZ, VBW_RBW_RATIO


class App(tk.Tk):
    """
    Main application class for the RF Spectrum Analyzer Controller.
    This class inherits from Tkinter's Tk class to create the main window
    and manage the overall application flow, including GUI setup, instrument
    communication, and data processing.
    """
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    CONFIG_FILE = os.path.join(_script_dir, 'config.ini')

    DEFAULT_WINDOW_GEOMETRY = "1400x780+100+100"

    def __init__(self):
        """
        Initializes the main application window and sets up core components.
        """
        super().__init__()
        self.title("RF Spectrum Analyzer Controller")
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        self.config = configparser.ConfigParser()
        self.config.read(self.CONFIG_FILE)

        self.is_ready_to_save = False

        self._check_and_install_dependencies()

        self.rm = None
        self.inst = None
        self.instrument_model = None

        self.collected_scans_dataframes = []
        self.last_scan_markers = [] # Ensure this is initialized

        self.scanning = False
        self.scan_thread = None
        self.stop_scan_event = threading.Event()
        self.pause_scan_event = threading.Event()

        self.SCAN_BAND_RANGES = SCAN_BAND_RANGES
        self.MHZ_TO_HZ = MHZ_TO_HZ
        self.VBW_RBW_RATIO = VBW_RBW_RATIO

        self._setup_tkinter_vars()

        load_config(self)
        self._apply_saved_geometry()

        set_debug_mode(self.general_debug_enabled_var.get())
        set_log_visa_commands_mode(self.log_visa_commands_enabled_var.get())

        self._create_widgets()
        self._setup_styles()
        self._redirect_stdout_to_console()
        
        self.update_connection_status(self.inst is not None)
        
        print_art()

        if hasattr(self, 'scan_tab'):
            self.scan_tab._load_band_selections_from_config()
            debug_print("Called _load_band_selections_from_config on ScanTab during startup.", file=__file__, function=inspect.currentframe().f_code.co_name)

        self.is_ready_to_save = True
        debug_print("Application fully initialized and ready to save configuration.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _check_and_install_dependencies(self):
        """
        Checks for necessary Python packages (pyvisa, pandas, beautifulsoup4, pdfplumber, requests)
        and attempts to install them if missing.
        """
        dependencies = {
            "pyvisa": "pyvisa",
            "pandas": "pandas",
            "bs4": "beautifulsoup4", # For BeautifulSoup
            "pdfplumber": "pdfplumber",
            "requests": "requests" # Added requests to the dependency check
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

        # VISA resource variables
        self.resource_names = tk.StringVar(self) # Holds the list of available VISA resources
        self.selected_resource = tk.StringVar(self) # Holds the currently selected VISA resource


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
        
        # Instantiate InstrumentTab and add it to the notebook
        self.instrument_tab = InstrumentTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.left_notebook.add(self.instrument_tab, text="Instrument Connection")


        # --- ScanTab (New tab for Scan Configuration and Bands to Scan) ---
        self.scan_tab = ScanTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.left_notebook.add(self.scan_tab, text="Scan Configuration")


        # --- Existing Notebook (Tabs) for Instrument Presets, Markers, Plotting, Report Converter ---
        self.preset_files_tab = PresetFilesTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.left_notebook.add(self.preset_files_tab, text="Instrument Presets")

        self.markers_display_tab = MarkersDisplayTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.left_notebook.add(self.markers_display_tab, text="Markers Display")

        self.plotting_tab = PlottingTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.left_notebook.add(self.plotting_tab, text="Plotting")

        self.report_converter_tab = ReportConverterTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.left_notebook.add(self.report_converter_tab, text="Report Converter")

        # --- NEW: VISA Interpreter Tab ---
        self.visa_interpreter_tab = VisaInterpreterTab(self.left_notebook, app_instance=self, console_print_func=self._print_to_gui_console)
        self.left_notebook.add(self.visa_interpreter_tab, text="VISA Interpreter")


        # Bind tab selection event to update specific tab contents for the left notebook
        self.left_notebook.bind("<<NotebookTabChanged>>", self._on_tab_change)


        # --- Right Column Container Frame ---
        right_column_container = ttk.Frame(self, style='Dark.TFrame')
        right_column_container.grid(row=0, column=1, rowspan=1, padx=5, pady=5, sticky="nsew")
        right_column_container.grid_columnconfigure(0, weight=1)
        right_column_container.grid_rowconfigure(0, weight=0)
        right_column_container.grid_rowconfigure(1, weight=1)


        # --- Scan Control Buttons Frame (Moved into right_column_container) ---
        self.scan_control_tab = ScanControlTab(right_column_container, app_instance=self, console_print_func=self._print_to_gui_console)
        self.scan_control_tab.grid(row=0, column=0, padx=5, pady=5, sticky="new")


        # --- Application Console Frame (Moved into right_column_container) ---
        console_frame = ttk.LabelFrame(right_column_container, text="Application Console", style='Dark.TLabelframe')
        console_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        console_frame.grid_rowconfigure(0, weight=1)
        console_frame.grid_columnconfigure(0, weight=1)

        self.console_text = scrolledtext.ScrolledText(console_frame, wrap="word", bg="#2b2b2b", fg="#cccccc", insertbackground="white")
        self.console_text.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.console_text.config(state=tk.DISABLED)
        
        debug_print("Main application widgets created.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _setup_styles(self):
        """
        Configures and applies custom ttk styles for a modern dark theme.
        """
        debug_print("Setting up ttk styles...", file=__file__, function=inspect.currentframe().f_code.co_name)
        style = ttk.Style(self)

        # Overall theme
        style.theme_use('clam')

        # General background and foreground colors
        BG_DARK = "#1e1e1e"
        FG_LIGHT = "#cccccc"
        ACCENT_BLUE = "#007bff"
        ACCENT_GREEN = "#28a745"
        ACCENT_RED = "#dc3545"
        ACCENT_ORANGE = "#ffc107"
        ACCENT_PURPLE = "#6f42c1"

        style.configure('.', background=BG_DARK, foreground=FG_LIGHT, font=('Helvetica', 10))
        style.configure('TFrame', background=BG_DARK)
        style.configure('TLabel', background=BG_DARK, foreground=FG_LIGHT)
        style.configure('TEntry', fieldbackground="#3b3b3b", foreground="#ffffff", borderwidth=1, relief="flat")
        style.map('TEntry', fieldbackground=[('focus', '#4a4a4a')])
        style.configure('TCombobox', fieldbackground="#3b3b3b", foreground="#ffffff", selectbackground=ACCENT_BLUE, selectforeground="white")
        style.map('TCombobox', fieldbackground=[('readonly', '#3b3b3b')], arrowcolor=[('!disabled', FG_LIGHT)])

        # Buttons
        style.configure('TButton',
                        background="#4a4a4a",
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

        style.configure('Orange.TButton', background=ACCENT_ORANGE, foreground="#333333")
        style.map('Orange.TButton', background=[('active', '#e0a800'), ('disabled', '#d39e00')])

        style.configure('Purple.TButton', background=ACCENT_PURPLE, foreground="white")
        style.map('Purple.TButton', background=[('active', '#5a2d9e'), ('disabled', '#4d2482')])


        # Checkbuttons
        style.configure('TCheckbutton', background=BG_DARK, foreground=FG_LIGHT, indicatorcolor="#4a4a4a")
        style.map('TCheckbutton',
                background=[('active', BG_DARK)],
                foreground=[('disabled', '#808080')],
                indicatorcolor=[('selected', ACCENT_BLUE)])

        # LabelFrame
        style.configure('TLabelFrame', background=BG_DARK, foreground=FG_LIGHT, borderwidth=1, relief="solid")
        style.configure('TLabelFrame.Label', background=BG_DARK, foreground=FG_LIGHT, font=('Helvetica', 10, 'bold'))
        style.configure('Dark.TLabelframe', background="#1e1e1e", foreground="#cccccc")
        style.configure('Dark.TLabelframe.Label', background="#1e1e1e", foreground="#cccccc")
        style.configure('Dark.TFrame', background="#1e1e1e")

        # Notebook (Tabs)
        style.configure('TNotebook', background=BG_DARK, borderwidth=0)
        style.configure('TNotebook.Tab', background="#3b3b3b", foreground=FG_LIGHT, padding=[10, 5])
        style.map('TNotebook.Tab',
                background=[('selected', ACCENT_BLUE), ('active', '#4a4a4a')],
                foreground=[('selected', 'white')],
                expand=[('selected', [1, 1, 1, 0])])

        # Treeview (for MarkersDisplayTab and VisaInterpreterTab)
        style.configure('Treeview',
                        background="#3b3b3b",
                        foreground="#ffffff",
                        fieldbackground="#3b3b3b",
                        rowheight=25)
        style.map('Treeview',
                background=[('selected', ACCENT_BLUE)],
                foreground=[('selected', 'white')])
        
        style.configure('Treeview.Heading',
                        background="#4a4a4a",
                        foreground="white",
                        font=('Helvetica', 10, "bold"))
        style.map('Treeview.Heading',
                background=[('active', '#606060')])

        # Markers Tab Specific Styles
        style.configure("Markers.TFrame", background="#1e1e1e")
        style.configure("Markers.TLabel", background="#1e1e1e", foreground="#cccccc")
        style.configure("Markers.TButton",
                        background="#6a5acd",
                        foreground="white",
                        font=("Helvetica", 10, "bold"),
                        padding=[10, 5])
        style.map("Markers.TButton",
                background=[('active', '#8a7ad0')])

        style.configure("SelectedSpan.TButton",
                        background="#ff6347",
                        foreground="white",
                        font=("Helvetica", 10, "bold"),
                        padding=[10, 5])
        style.map("SelectedSpan.TButton",
                background=[('active', '#e05038')])

        style.configure("LargePreset.TButton",
                        background="#4a4a4a",
                        foreground="white",
                        font=("Helvetica", 14, "bold"),
                        padding=[30, 15, 30, 15])
        style.map("LargePreset.TButton",
                background=[('active', '#606060')])

        style.configure("SelectedPreset.TButton",
                        background="#007bff",
                        foreground="white",
                        font=("Helvetica", 14, "bold"),
                        padding=[30, 15, 30, 15])
        style.map("SelectedPreset.TButton",
                background=[('active', '#0056b3')])

        YAK_ORANGE = "#ff8c00"
        style.configure('LargeYAK.TButton',
                        font=('Helvetica', 100, 'bold'),
                        background=YAK_ORANGE,
                        foreground="white",
                        padding=[20, 10])
        style.map('LargeYAK.TButton',
                  background=[('active', '#e07b00'), ('disabled', '#cc7000')])


        debug_print("ttk styles set up.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _redirect_stdout_to_console(self):
        """
        Redirects standard output and error streams to the GUI's scrolled text widget.
        """
        debug_print("Redirecting stdout/stderr to GUI console...", file=__file__, function=inspect.currentframe().f_code.co_name)
        sys.stdout = TextRedirector(self.console_text, "stdout")
        sys.stderr = TextRedirector(self.console_text, "stderr")
        print("Application console initialized.")

    def _print_to_gui_console(self, message, overwrite=False):
        """
        A helper function to print messages to the GUI console from any thread.
        Uses after() to ensure thread safety.
        If overwrite is True, it attempts to overwrite the last line.
        """
        self.after(0, lambda: self._update_console_text(message, overwrite))

    def _update_console_text(self, message, overwrite):
        """
        Appends a message to the scrolled text widget.
        If overwrite is True, it deletes the last line before inserting.
        """
        self.console_text.config(state=tk.NORMAL)
        if overwrite:
            last_line_start = self.console_text.index("end-1c linestart")
            self.console_text.delete(last_line_start, tk.END)
        
        self.console_text.insert(tk.END, message + "\n")
        self.console_text.see(tk.END)
        self.console_text.config(state=tk.DISABLED)


    def _populate_resources(self):
        """
        Delegates the call to populate VISA resources to the InstrumentTab.
        """
        debug_print("Delegating populate VISA resources to InstrumentTab...", file=__file__, function=inspect.currentframe().f_code.co_name)
        if hasattr(self, 'instrument_tab'):
            self.instrument_tab._populate_resources()
        else:
            self._print_to_gui_console("⚠️ Warning: InstrumentTab not initialized. Cannot populate resources.")
            debug_print("InstrumentTab not initialized for _populate_resources.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _on_resource_selected(self, event):
        """
        Delegates the resource selection callback to the InstrumentTab.
        """
        debug_print("Delegating resource selection to InstrumentTab...", file=__file__, function=inspect.currentframe().f_code.co_name)
        if hasattr(self, 'instrument_tab'):
            self.instrument_tab._on_resource_selected(event)
        else:
            self._print_to_gui_console("⚠️ Warning: InstrumentTab not initialized. Cannot handle resource selection.")
            debug_print("InstrumentTab not initialized for _on_resource_selected.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _connect_instrument(self):
        """
        Delegates the connect instrument call to the InstrumentTab.
        """
        debug_print("Delegating connect instrument to InstrumentTab...", file=__file__, function=inspect.currentframe().f_code.co_name)
        if hasattr(self, 'instrument_tab'):
            self.instrument_tab._connect_instrument()
        else:
            self._print_to_gui_console("⚠️ Warning: InstrumentTab not initialized. Cannot connect instrument.")
            debug_print("InstrumentTab not initialized for _connect_instrument.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _disconnect_instrument(self):
        """
        Delegates the disconnect instrument call to the InstrumentTab.
        """
        debug_print("Delegating disconnect instrument to InstrumentTab...", file=__file__, function=inspect.currentframe().f_code.co_name)
        if hasattr(self, 'instrument_tab'):
            self.instrument_tab._disconnect_instrument()
        else:
            self._print_to_gui_console("⚠️ Warning: InstrumentTab not initialized. Cannot disconnect instrument.")
            debug_print("InstrumentTab not initialized for _disconnect_instrument.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _apply_instrument_settings(self):
        """
        Delegates the apply instrument settings call to the InstrumentTab.
        """
        debug_print("Delegating apply instrument settings to InstrumentTab...", file=__file__, function=inspect.currentframe().f_code.co_name)
        if hasattr(self, 'instrument_tab'):
            self.instrument_tab._apply_settings()
        else:
            self._print_to_gui_console("⚠️ Warning: InstrumentTab not initialized. Cannot apply settings.")
            debug_print("InstrumentTab not initialized for _apply_instrument_settings.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _load_selected_preset(self):
        """
        Loads the currently selected preset file onto the instrument.
        This function is called when the "Load Selected Preset" button is clicked.
        It delegates the actual loading logic to instrument_preset_tab.
        """
        debug_print("Loading selected preset...", file=__file__, function=inspect.currentframe().f_code.co_name)
        if hasattr(self, 'preset_files_tab'):
            self.preset_files_tab._load_selected_preset()
        else:
            self._print_to_gui_console("⚠️ Warning: Preset Files tab not initialized.")
            debug_print("PresetFilesTab not initialized.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _start_scan(self):
        """
        Initiates the scan process in a separate thread.
        """
        debug_print("Starting scan...", file=__file__, function=inspect.currentframe().f_code.co_name)
        if hasattr(self, 'scan_control_tab'):
            self.scan_control_tab._start_scan()
        else:
            self._print_to_gui_console("⚠️ Warning: Scan Control tab not initialized.")
            debug_print("ScanControlTab not initialized for _start_scan.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _pause_scan(self):
        """
        Pauses the active scan.
        """
        debug_print("Pausing scan...", file=__file__, function=inspect.currentframe().f_code.co_name)
        if hasattr(self, 'scan_control_tab'):
            self.scan_control_tab._pause_scan()
        else:
            self._print_to_gui_console("⚠️ Warning: Scan Control tab not initialized.")
            debug_print("ScanControlTab not initialized for _pause_scan.", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _stop_scan(self):
        """
        Stops the active scan.
        """
        debug_print("Stopping scan...", file=__file__, function=inspect.currentframe().f_code.co_name)
        if hasattr(self, 'scan_control_tab'):
            self.scan_control_tab._stop_scan()
        else:
            self._print_to_gui_console("⚠️ Warning: Scan Control tab not initialized.")
            debug_print("ScanControlTab not initialized for _stop_scan.", file=__file__, function=inspect.currentframe().f_code.co_name)


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
        for tk_var_name, (_, _, tk_var_instance) in self.setting_var_map.items():
            if isinstance(tk_var_instance, (tk.StringVar, tk.IntVar, tk.DoubleVar)):
                try:
                    for widget in self.winfo_children():
                        if isinstance(widget, (ttk.LabelFrame, tk.Frame, ttk.Notebook)):
                            for child in widget.winfo_children():
                                if isinstance(child, ttk.Frame):
                                    for grand_child in child.winfo_children():
                                        if isinstance(grand_child, ttk.Entry) and grand_child.cget('textvariable') == str(tk_var_instance):
                                            grand_child.config(style='TEntry')
                                            debug_print(f"Reset color for {tk_var_name}", file=__file__, function=inspect.currentframe().f_code.co_name)
                                            break
                                if isinstance(child, ttk.Entry) and child.cget('textvariable') == str(tk_var_instance):
                                    child.config(style='TEntry')
                                    debug_print(f"Reset color for {tk_var_name}", file=__file__, function=inspect.currentframe().f_code.co_name)
                                    break
                except Exception as e:
                    debug_print(f"Could not reset color for {tk_var_name}: {e}", file=__file__, function=inspect.currentframe().f_code.co_name)


    def _on_closing(self):
        """
        Handles the window closing event. Saves configuration before exiting,
        but only if the application is fully initialized.
        """
        debug_print("Application closing. Performing cleanup...", file=__file__, function=inspect.currentframe().f_code.co_name)
        
        if self.is_ready_to_save:
            debug_print("--- State of band_vars before saving ---", file=__file__, function=inspect.currentframe().f_code.co_name)
            for i, band_item in enumerate(self.band_vars):
                debug_print(f"  Band {band_item['band']['Band Name']}: {band_item['var'].get()}", file=__file__, function=inspect.currentframe().f_code.co_name)
            debug_print("--- End state of band_vars ---", file=__file__, function=inspect.currentframe().f_code.co_name)
            save_config(self)
        else:
            debug_print("Application closing prematurely or not fully initialized. Skipping config save.", file=__file__, function=inspect.currentframe().f_code.co_name)

        self.destroy()


    def _on_tab_change(self, event):
        """
        Handles tab change events in the Notebook.
        Calls _on_tab_selected for the newly selected tab if available.
        """
        selected_tab_id = self.left_notebook.select()
        selected_tab_widget = self.left_notebook.nametowidget(selected_tab_id)
        
        if hasattr(selected_tab_widget, '_on_tab_selected'):
            selected_tab_widget._on_tab_selected(event)
            debug_print(f"Tab changed to {selected_tab_widget.winfo_class()}. Calling _on_tab_selected.", file=__file__, function=inspect.currentframe().f_code.co_name)
        else:
            debug_print(f"Tab changed to {selected_tab_widget.winfo_class()}. No _on_tab_selected method found.", file=__file__, function=inspect.currentframe().f_code.co_name)

    def update_connection_status(self, is_connected):
        """
        A wrapper method in the App class to call the external logic function.
        This method will be called by other parts of the application.
        """
        update_connection_status_logic(self, is_connected, self._print_to_gui_console)


if __name__ == "__main__":
    app = App()
    app.mainloop()
