# user_interface.py
#
# This module is responsible for building and managing the graphical user interface (GUI)
# of the RF Spectrum Analyzer Controller application. It uses Tkinter to create the
# main window, input fields, buttons, and display areas, allowing users to interact
# with the spectrum analyzer, configure scan settings, view real-time output, and
# generate plots. It acts as the central hub for user interaction and orchestrates
# calls to other utility modules for instrument control, scanning, data processing,
# and plotting.
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
import xml.etree.ElementTree as ET # Added for SHW conversion

# --- BeautifulSoup Installation Check ---
# This block checks if BeautifulSoup4 is installed. If not, it attempts to install it.
# This is crucial for parsing HTML files.
try:
    from bs4 import BeautifulSoup
except ImportError:
    print("BeautifulSoup4 not found. Attempting to install...")
    try:
        # Use subprocess to run pip install for BeautifulSoup4
        subprocess.check_call([sys.executable, "-m", "pip", "install", "beautifulsoup4"])
        from bs4 import BeautifulSoup
        print("BeautifulSoup4 installed successfully.")
    except subprocess.CalledProcessError as e:
        # If installation fails, show an error message and exit
        messagebox.showerror("Installation Error", f"Error installing BeautifulSoup4: {e}\\nPlease install it manually by running: pip install beautifulsoup4")
        sys.exit(1)
    except Exception as e:
        # Catch any other unexpected errors during installation
        messagebox.showerror("Installation Error", f"An unexpected error occurred during BeautifulSoup4 installation: {e}")
        sys.exit(1)


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
from utils.averaging_utils import generate_historical_average_plot
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

# Import the new report converter utility functions
from utils.report_converter_utils import convert_html_report_to_csv, generate_csv_from_shw

# Define the config file name, ensuring it's in the same directory as the script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, 'config.ini')

class TextRedirector(object):
    """
    A class to redirect standard output (stdout) and standard error (stderr)
    to a Tkinter scrolled text widget. This allows all print statements and
    error messages from the application's backend to be displayed directly
    within the GUI's console area, providing real-time feedback to the user.
    """
    def __init__(self, widget, tag="stdout"):
        """
        Initializes the TextRedirector.

        Inputs:
            widget (tk.scrolledtext.ScrolledText): The Tkinter scrolled text widget
                                                  where output will be displayed.
            tag (str, optional): A tag for text formatting within the widget. Defaults to "stdout".
        Process:
            1. Stores the provided `widget` and `tag`.
            2. Initializes `last_char_was_cr` to False, used for handling carriage returns for line overwriting.
        Outputs: None
        """
        self.widget = widget
        self.tag = tag
        self.last_char_was_cr = False

    def write(self, str_val):
        """
        Writes the given string value to the Tkinter scrolled text widget.
        Handles carriage returns (`\r`) to overwrite the current line,
        useful for progress bars or dynamic console updates.

        Inputs:
            str_val (str): The string to write to the console.
        Process:
            1. Sets the widget state to `tk.NORMAL` to allow editing.
            2. Checks if `str_val` contains `\r`.
            3. If `\r` is present, splits the string and handles line deletion
               for overwriting if the previous character was also a carriage return.
            4. Inserts the string (or parts of it) into the widget at the end.
            5. Scrolls the widget to the end to show the latest output.
            6. Sets the widget state back to `tk.DISABLED` to prevent user editing.
            7. Updates Tkinter idle tasks to ensure immediate display.
        Outputs: None
        """
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
        """
        Required method for file-like objects. Does nothing in this implementation.
        """
        pass

class MarkersDisplayTab(tk.Frame):
    """
    A Tkinter Frame that displays extracted frequency markers in a hierarchical treeview
    and as clickable buttons.
    """
    def __init__(self, master=None, headers=None, rows=None, **kwargs):
        super().__init__(master, **kwargs)
        self.configure(bg="black")
        self.headers = headers if headers is not None else []
        self.rows = rows if rows is not None else []
        self.create_widgets()

    def create_widgets(self):
        # Main frame for the split layout
        main_split_frame = tk.Frame(self, bg="black")
        main_split_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main_split_frame.grid_columnconfigure(0, weight=1) # Left half
        main_split_frame.grid_columnconfigure(1, weight=1) # Right half
        main_split_frame.grid_rowconfigure(0, weight=1)

        # Left Half: Treeview for Zones and Groups
        tree_frame = tk.LabelFrame(main_split_frame, text="Zones & Groups", bg="black", fg="white", padx=5, pady=5)
        tree_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=5, pady=5)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.zone_group_tree = ttk.Treeview(tree_frame, show="tree") # Only show tree, not headings
        self.zone_group_tree.pack(fill=tk.BOTH, expand=True)

        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.zone_group_tree.yview)
        tree_scrollbar.pack(side=tk.RIGHT, fill="y")
        self.zone_group_tree.configure(yscrollcommand=tree_scrollbar.set)

        self._populate_zone_group_tree()

        # Right Half: Buttons for Devices
        buttons_frame = tk.LabelFrame(main_split_frame, text="Devices", bg="black", fg="white", padx=5, pady=5)
        buttons_frame.grid(row=0, column=1, sticky=tk.NSEW, padx=5, pady=5)
        
        # Use a canvas with a scrollbar for buttons if there are many
        buttons_canvas = tk.Canvas(buttons_frame, bg="black", highlightbackground="black")
        buttons_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        buttons_scrollbar = ttk.Scrollbar(buttons_frame, orient="vertical", command=buttons_canvas.yview)
        buttons_scrollbar.pack(side=tk.RIGHT, fill="y")

        buttons_canvas.configure(yscrollcommand=buttons_scrollbar.set)
        buttons_canvas.bind('<Configure>', lambda e: buttons_canvas.configure(scrollregion = buttons_canvas.bbox("all")))

        self.inner_buttons_frame = tk.Frame(buttons_canvas, bg="black")
        buttons_canvas.create_window((0, 0), window=self.inner_buttons_frame, anchor="nw")

        self._populate_device_buttons()

    def _populate_zone_group_tree(self):
        self.zone_group_tree.delete(*self.zone_group_tree.get_children()) # Clear existing items
        
        zones = {}
        for row in self.rows:
            zone_name = row.get("ZONE", "Unknown Zone")
            group_name = row.get("GROUP", "Unknown Group")

            if zone_name not in zones:
                zones[zone_name] = {}
            if group_name not in zones[zone_name]:
                zones[zone_name][group_name] = []
            zones[zone_name][group_name].append(row) # Store the full row for later use if needed

        for zone_name in sorted(zones.keys()):
            zone_id = self.zone_group_tree.insert("", "end", text=zone_name, open=True)
            for group_name in sorted(zones[zone_name].keys()):
                self.zone_group_tree.insert(zone_id, "end", text=group_name)

    def _populate_device_buttons(self):
        # Clear existing buttons
        for widget in self.inner_buttons_frame.winfo_children():
            widget.destroy()

        for row_data in self.rows:
            device = row_data.get("DEVICE", "N/A")
            name = row_data.get("NAME", "N/A")
            freq = row_data.get("FREQ", "N/A")

            button_text = f"{device}\n{name}\n{freq}"
            btn = tk.Button(self.inner_buttons_frame, text=button_text, 
                            font=('Arial', 9), bg='darkblue', fg='white',
                            activebackground='blue', activeforeground='white',
                            relief=tk.RAISED, bd=2, padx=5, pady=3, wraplength=150)
            btn.pack(fill=tk.X, pady=2) # Pack buttons vertically


class ReportConverterTab(tk.Frame):
    """
    A Tkinter Frame that encapsulates the functionality of the Report Converter.
    This includes converting HTML and SHW files to CSV format.
    """
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.configure(bg="black")
        self.create_widgets()

    def create_widgets(self):
        """
        Creates the widgets for the Report Converter tab.
        """
        # Create a label for instructions
        instruction_label = tk.Label(self, text="Click the button below to select a report file (.html or .shw)", 
                                     wraplength=350, bg="black", fg="white")
        instruction_label.pack(pady=10)

        # Create a button to open the file dialog
        select_button = tk.Button(self, text="Select Report File", command=self.select_file,
                                  font=('Arial', 12, 'bold'), bg='#4CAF50', fg='white',
                                  activebackground='#45a049', activeforeground='white',
                                  relief=tk.RAISED, bd=3, padx=10, pady=5)
        select_button.pack(pady=20)

    def select_file(self):
        """
        Opens a file dialog for the user to select an HTML or SHW file,
        then processes it accordingly.
        """
        file_path = filedialog.askopenfilename(
            title="Select an IAS HTML Report or a SHURE Wireless Workbench Show File",
            filetypes=[("Report Files", "*.html *.shw"), ("HTML files", "*.html"), ("SHW files", "*.shw")]
        )

        if not file_path:
            return # User cancelled file selection

        file_name = os.path.basename(file_path)
        base_name, extension = os.path.splitext(file_name)
        
        # Get the output directory from the main App instance's output_folder_var
        # self.master is the notebook, self.master.master is the main_frame, self.master.master.master is the App instance
        output_dir = self.master.master.master.output_folder_var.get()
        if not os.path.exists(output_dir):
            os.makedirs(output_dir) # Ensure the output directory exists
        
        output_csv_file = os.path.join(output_dir, "MARKERS.CSV")

        headers = []
        rows = []
        conversion_successful = False
        error_message = ""

        try:
            if extension.lower() == '.html':
                with open(file_path, 'r', encoding='utf-8') as f:
                    html_report_content = f.read()
                headers, rows = convert_html_report_to_csv(html_report_content)
                conversion_successful = True
            
            elif extension.lower() == '.shw':
                headers, rows = generate_csv_from_shw(file_path)
                conversion_successful = True
            
            else:
                messagebox.showwarning("Invalid File Type", "Please select a .html or .shw file.")
                return # Exit if file type is invalid

            if conversion_successful:
                if rows: # Only write if there's data to write
                    with open(output_csv_file, 'w', newline='', encoding='utf-8') as csvfile:
                        csv_writer = csv.DictWriter(csvfile, fieldnames=headers)
                        csv_writer.writeheader()
                        csv_writer.writerows(rows)
                    messagebox.showinfo("Success", f"Successfully converted '{file_name}' to '{os.path.basename(output_csv_file)}'")
                    
                    # Call the method on the main App instance to add the new tab
                    self.master.master.master.add_markers_tab(headers, rows)
                else:
                    messagebox.showwarning("No Data Extracted", f"No relevant data could be extracted from '{file_name}'. CSV file was not created.")

        except FileNotFoundError as e:
            error_message = f"File not found: {e}"
            messagebox.showerror("File Error", error_message)
        except ET.ParseError as e:
            error_message = f"Error parsing XML (SHW) file: {e}"
            messagebox.showerror("Parsing Error", error_message)
        except Exception as e:
            error_message = f"An unexpected error occurred during conversion: {e}"
            messagebox.showerror("Conversion Error", error_message)
        
        if error_message:
            print(f"❌ Conversion failed for {file_name}: {error_message}")


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
            6. Loads configuration settings from `config.ini` using `_load_config()`.
            7. Initializes sliders for RBW and Frequency Shift, binding them to update
               the relevant Tkinter variables.
            8. Creates the main GUI widgets by calling `create_widgets()`.
            9. Redirects `sys.stdout` and `sys.stderr` to the console output widget
               using `TextRedirector`.
            10. Prints ASCII art and a welcome message to the console.
            11. Populates available VISA resources using `populate_resources()`.
            12. Binds the window closing protocol to `on_closing` for saving settings.
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
        self.last_scan_data = None # This will still store the in-memory data for potential other uses
        self.collected_scans_dataframes = []
        self.instrument_model = None

        self.scan_cycle_count = 0
        self.current_freq_offset = 0

        self.desired_setting_entries = {}

        # Variable to control the blinking of the connect button
        self.blink_id = None
        self.blink_on = False

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

        print("--- RF Spectrum Analyzer Controller - GUI Initialized ---")
        self.create_widgets()
        self.after(0, self.populate_resources)

        

        # Bind the closing protocol to save config
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _load_config(self):
        """
        Loads configuration settings from `config.ini`.
        If the file or sections are missing, it ensures default settings are present.
        This function is called during application initialization to restore
        the last-used settings or apply defaults.

        Inputs: None
        Process:
            1. Reads `CONFIG_FILE` into `self.config` using `configparser`.
            2. Ensures `DEFAULT_SETTINGS` section exists, creating it if not.
            3. Dynamically generates default selected bands string from `SCAN_BAND_RANGES`.
            4. Populates `DEFAULT_SETTINGS` with predefined default values if any are missing.
            5. Ensures `LAST_USED_SETTINGS` section exists, creating it if not.
            6. Calls `_populate_vars_from_config()` to set Tkinter variables based on loaded config.
        Outputs: None
        """
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
        """
        Populates Tkinter variables with values from the loaded configuration.
        It prioritizes `LAST_USED_SETTINGS` and falls back to `DEFAULT_SETTINGS`
        if a value is not found or is invalid in the last used section.

        Inputs: None
        Process:
            1. Iterates through `self.setting_var_map`, which defines the mapping
               between Tkinter variables and config keys.
            2. For each setting, attempts to retrieve the value from `LAST_USED_SETTINGS`.
            3. If not found or parsing fails, attempts to retrieve from `DEFAULT_SETTINGS`.
            4. Converts the retrieved string value to the appropriate Python type (bool, float, int, str).
            5. Sets the corresponding Tkinter variable's value.
            6. Handles a special case for `resource_var` if no resource is found.
        Outputs: None
        """
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
        """
        Saves the current state of all configurable settings from the GUI's Tkinter
        variables to the `config.ini` file, specifically in the `LAST_USED_SETTINGS` section.
        This ensures that user preferences are preserved between application sessions.

        Inputs: None
        Process:
            1. Ensures the `LAST_USED_SETTINGS` section exists in `self.config`.
            2. Iterates through `self.setting_var_map` and saves the current value
               of each Tkinter variable to its corresponding `last_used_config_key`.
            3. Special handling to save the currently selected frequency bands.
            4. Saves the current window geometry.
            5. Writes the updated configuration to `CONFIG_FILE`.
        Outputs: None
        """
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
        """
        Handler for the window closing event. This function is called when
        the user attempts to close the application window. It ensures that
        the current configuration settings are saved before the application exits.

        Inputs: None
        Process:
            1. Calls `_save_config()` to persist current settings.
            2. Destroys the Tkinter root window, effectively closing the application.
        Outputs: None
        """
        self._save_config()
        self.destroy()

    def _update_debug_mode_global(self, *args):
        """
        Updates the global debug mode variable in the `instrument_control` module
        whenever the state of the debug mode checkbox in the GUI changes.
        This allows for dynamic enabling/disabling of verbose logging for VISA commands.

        Inputs:
            *args: Standard Tkinter trace arguments (not used).
        Process:
            1. Retrieves the current boolean value from `self.debug_mode_var`.
            2. Calls `set_debug_mode()` from `instrument_control` to update the global flag.
            3. Calls `_save_config()` to immediately persist the debug mode setting.
        Outputs: None
        """
        set_debug_mode(self.debug_mode_var.get())
        # Also update the config setting immediately
        # This specific debug mode update is handled by the generic _save_config now.
        # self.config['DEFAULT_SETTINGS']['DEFAULT_DEBUG_MODE'] = str(self.debug_mode_var.get())
        self._save_config()


    def _update_scan_rbw_from_slider_index(self, *args):
        """
        Updates the `desired_scan_rbw_segmentation_var` (which controls the instrument's RBW)
        based on the position of the RBW slider. This links the visual slider
        to the numerical RBW setting.

        Inputs:
            *args: Standard Tkinter trace arguments (not used).
        Process:
            1. Retrieves the current index from `self.rbw_slider_index_var`.
            2. Uses the index to look up the corresponding RBW value from `self.rbw_values`.
            3. Sets the `desired_scan_rbw_segmentation_var` to this float value.
            4. Includes error handling for invalid index access.
        Outputs: None
        """
        try:
            idx = self.rbw_slider_index_var.get()
            if 0 <= idx < len(self.rbw_values):
                self.desired_scan_rbw_segmentation_var.set(float(self.rbw_values[idx]))
        except Exception as e:
            print(f"Error updating scan RBW from slider index: {e}")

    def _update_freq_shift_from_slider_index(self, *args):
        """
        Updates the `shift_freq_var` (which controls the frequency offset)
        based on the position of the Frequency Shift slider. This links the visual slider
        to the numerical frequency offset setting.

        Inputs:
            *args: Standard Tkinter trace arguments (not used).
        Process:
            1. Retrieves the current index from `self.freq_shift_slider_index_var`.
            2. Uses the index to look up the corresponding frequency shift value from `self.freq_shift_values`.
            3. Sets the `shift_freq_var` to this float value.
            4. Includes error handling for invalid index access.
        Outputs: None
        """
        try:
            idx = self.freq_shift_slider_index_var.get()
            if 0 <= idx < len(self.freq_shift_values):
                self.shift_freq_var.set(float(self.freq_shift_values[idx]))
        except Exception as e:
            print(f"Error updating frequency shift from slider index: {e}")

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
               - Refresh, Connect, Disconnect buttons
            3. Creates `scan_settings_frame` for instrument configuration:
               - Restore Default Settings button
               - Entry fields and checkboxes for Reference Level, High Sensitivity, Preamplifier
               - Sliders and entry fields for Scan RBW and Frequency Shift
               - Sliders for Cycle Hold Time and Cycle Wait Time
               - Entry fields for Scan Name and Output Folder, with an "Open Folder" button
               - Checkboxes for including TV and Government band markers in plots
               - Checkbox for auto-opening HTML plots
               - Buttons for "Apply Settings to Device" and "Generate Plot (Average)"
            4. Creates `ttk.Notebook` for tabbed interface:
               - "Frequency Band Selection" tab with a scrollable canvas for frequency band checkboxes.
               - "Device Preset Files" tab with a "Load Selected Preset" button and a `ttk.Treeview`
                 to display device preset files.
               - "Report Converter" tab with the functionality from `import_markers.py`.
            5. Creates a `debug_frame` with a checkbox to enable/disable debug mode.
            6. Configures grid weights for responsive layout.
            7. Calls `update_vbw_display()` to initialize the VBW display.
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

        # Style for the Notebook (tabs)
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

        self.refresh_button = tk.Button(resource_frame, text="Refresh", command=lambda: self.populate_resources(), bg="darkgrey", fg="white")
        self.refresh_button.grid(row=0, column=2, padx=5, pady=2)

        self.connect_button = tk.Button(resource_frame, text="Connect", command=lambda: self.connect_instrument(), state=tk.DISABLED, bg="darkgrey", fg="white")
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

        # --- Tabbed Interface for Bands and Presets ---
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # Tab 1: Frequency Band Selection
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
        
        self._set_band_checkboxes_from_config()

        # Tab 2: Device Preset Files
        preset_files_tab = tk.Frame(self.notebook, bg="black")
        self.notebook.add(preset_files_tab, text="Device Preset Files")

        self.load_preset_button = tk.Button(preset_files_tab, text="Load Selected Preset", command=self.load_selected_preset, state=tk.DISABLED, bg="darkgrey", fg="white")
        self.load_preset_button.pack(pady=5)

        self.preset_tree = ttk.Treeview(preset_files_tab, columns=("Name",), show="headings", selectmode="browse")
        self.preset_tree.heading("Name", text="Preset File Name")
        self.preset_tree.column("Name", width=200, anchor="w")
        self.preset_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.preset_tree.tag_configure("Mon", foreground="blue")

        preset_scrollbar = ttk.Scrollbar(preset_files_tab, orient="vertical", command=self.preset_tree.yview)
        preset_scrollbar.pack(side=tk.RIGHT, fill="y")
        self.preset_tree.configure(yscrollcommand=preset_scrollbar.set)

        self.preset_tree.bind("<<TreeviewSelect>>", self._on_preset_select)
        
        # Tab 3: Report Converter
        report_converter_tab = ReportConverterTab(self.notebook, bg="black")
        self.notebook.add(report_converter_tab, text="Report Converter")

        # --- End Tabbed Interface ---

        debug_frame = tk.Frame(self.main_frame, bg="black")
        debug_frame.pack(pady=10, padx=10, fill=tk.X)
        tk.Checkbutton(debug_frame, text="Enable Debug Mode (Log VISA Commands)", variable=self.debug_mode_var, bg="black", fg="white", selectcolor="grey", activebackground="black", activeforeground="white").pack(anchor=tk.W)

        for i in range(5):
            resource_frame.grid_columnconfigure(i, weight=1)
        scan_settings_frame.grid_columnconfigure(0, weight=1)
        scan_settings_frame.grid_columnconfigure(1, weight=1)

        self.update_vbw_display()

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
            1. Creates a new `MarkersDisplayTab` instance, passing the extracted
               `headers` and `rows`.
            2. Adds this new tab to the `self.notebook` with the text "Markers Display".
            3. Selects the newly created tab to bring it into view.
        Outputs: None
        """
        # Check if a "Markers Display" tab already exists and remove it
        for i, tab_id in enumerate(self.notebook.tabs()):
            tab_text = self.notebook.tab(tab_id, "text")
            if tab_text == "Markers Display":
                self.notebook.forget(tab_id)
                print("Existing 'Markers Display' tab removed.")
                break

        markers_display_tab = MarkersDisplayTab(self.notebook, headers=headers, rows=rows, bg="black")
        self.notebook.add(markers_display_tab, text="Markers Display")
        self.notebook.select(markers_display_tab) # Switch to the new tab

    def _set_band_checkboxes_from_config(self):
        """
        Sets the state (checked/unchecked) of the frequency band checkboxes
        based on the `last_selected_bands` setting loaded from `config.ini`.
        If no last selected bands are found, all checkboxes are set to `True` (selected).

        Inputs: None
        Process:
            1. Retrieves the string of comma-separated band names from `self.last_selected_bands_str`.
            2. Splits the string into a list of selected band names.
            3. Iterates through `self.band_vars` (which holds each band's name and its Tkinter `BooleanVar`).
            4. If a band's name is found in the list of selected band names, its `BooleanVar` is set to `True`.
            5. Otherwise, it's set to `False`.
            6. If no `last_selected_bands` are found in config, all bands are set to `True`.
        Outputs: None
        """
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
        """
        Updates the displayed Video Bandwidth (VBW) value based on the
        current Resolution Bandwidth (RBW) setting. The VBW is typically
        set to approximately 1/3 of the RBW for spectrum analyzers.

        Inputs: None
        Process:
            1. Retrieves the current RBW value from `self.desired_scan_rbw_segmentation_var`.
            2. Calculates VBW as RBW / 3.
            3. Sets the `self.desired_vbw_display_var` to the calculated VBW (as an integer string).
            4. Handles `ValueError` if the RBW input is not a valid number.
        Outputs: None
        """
        try:
            scan_rbw_val = float(self.desired_scan_rbw_segmentation_var.get())
            self.desired_vbw_display_var.set(str(int(scan_rbw_val / 3)))
        except ValueError:
            self.desired_vbw_display_var.set("Invalid RBW")

    def open_output_folder(self):
        """
        Opens the specified output folder in the user's default file explorer.
        This provides a convenient way for users to access their saved scan data and plots.

        Inputs: None
        Process:
            1. Retrieves the output folder path from `self.output_folder_var`.
            2. Converts the path to an absolute path if it's relative.
            3. Checks if the folder exists, showing a warning if not.
            4. Uses `os.startfile` (Windows), `subprocess.run(['open', ...])` (macOS),
               or `subprocess.run(['xdg-open', ...])` (Linux) to open the folder.
            5. Prints success or error messages to the console and displays a messagebox on error.
        Outputs: None
        """
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
        """
        Establishes a connection to the selected VISA instrument (spectrum analyzer).
        Upon successful connection, it initializes the instrument with the desired
        settings from the GUI and updates the GUI state accordingly.

        Inputs: None
        Process:
            1. Retrieves the selected VISA resource string from `self.resource_var`.
            2. Performs initial checks for valid resource selection.
            3. If an existing connection exists, it attempts to close it first.
            4. Calls `connect_to_instrument()` from `instrument_control` to establish the connection.
            5. If connection is successful:
               - Updates the window title with the instrument model.
               - Retrieves current GUI settings for Reference Level, Sensitivity, Preamplifier, RBW, and VBW.
               - Calls `initialize_instrument()` from `instrument_control` to configure the device.
               - Queries and displays current instrument settings.
               - Enables/disables relevant GUI buttons (Start Scan, Disconnect, Apply, Load Preset).
               - Stops the connect button blinking.
               - Queries device presets (if not N9340B) and updates the preset treeview.
            6. If connection or initialization fails, displays appropriate error messages
               and resets the GUI to a disconnected state.
            7. Includes error handling for `ValueError` (invalid numeric inputs) and general exceptions.
        Outputs: None
        """
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
                    self._stop_connect_button_blink() # Stop blinking on successful connection

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
        """
        Closes the active connection to the spectrum analyzer.
        Resets the GUI state to reflect a disconnected instrument.

        Inputs: None
        Process:
            1. Calls `control_disconnect_instrument()` from `instrument_control`.
            2. If successful, sets `self.inst` and `self.instrument_model` to `None`.
            3. Resets the GUI button states using `_reset_gui_on_disconnect_or_error()`.
            4. Resets the window title.
            5. If resources are available, restarts the connect button blinking.
            6. Displays an error messagebox if disconnection fails.
        Outputs: None
        """
        if control_disconnect_instrument(self.inst):
            self.inst = None
            self.instrument_model = None
            print("Disconnected.")
            self._reset_gui_on_disconnect_or_error()
            self.title(f"RF Spectrum Analyzer Controller - {os.path.basename(sys.argv[0])}")
            # Start blinking again if resources are available
            if self.instrument_list and self.resource_var.get() != "No resources found":
                self._start_connect_button_blink()
        else:
            messagebox.showerror("Disconnect Error", "Failed to disconnect instrument.")
            
    def apply_settings_to_device(self):
        """
        Applies the currently configured settings from the GUI's input fields
        and checkboxes directly to the connected spectrum analyzer.

        Inputs: None
        Process:
            1. Checks if an instrument is connected, showing a warning if not.
            2. Retrieves the desired settings (Reference Level, High Sensitivity, Preamplifier, RBW, VBW).
            3. Calls `initialize_instrument()` from `instrument_control` to send these commands to the device.
            4. Displays success or error messages and resets the color of setting entries.
            5. Includes error handling for `ValueError` (invalid numeric inputs) and general exceptions.
        Outputs: None
        """
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
        """
        Resets the foreground color of all input entry widgets in the "Scan Configuration"
        section to white. This is typically called after settings are successfully applied
        to provide visual feedback that the settings are no longer "dirty" or unapplied.

        Inputs: None
        Process:
            1. Iterates through the `self.desired_setting_entries` dictionary.
            2. For each entry that is a `tk.Entry` widget, sets its `fg` (foreground)
               color to "white".
        Outputs: None
        """
        for key, entry_widget in self.desired_setting_entries.items():
            if isinstance(entry_widget, tk.Entry):
                entry_widget.config(fg="white")

    def _update_preset_tree(self, preset_files):
        """
        Updates the `ttk.Treeview` widget that displays the list of
        available preset files on the connected instrument.

        Inputs:
            preset_files (list): A list of strings, where each string is the name of a preset file.
        Process:
            1. Clears all existing items from the `preset_tree`.
            2. If `preset_files` is not empty, it inserts each preset name into the treeview,
               sorted alphabetically. It also applies a "Mon" tag for blue foreground if
               "MON" in the preset name (likely for 'Monitor' presets).
            3. If `preset_files` is empty, it inserts a "No .STA preset files found." message.
            4. Disables the "Load Selected Preset" button.
        Outputs: None
        """
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
        """
        Event handler for when a preset file is selected in the `preset_tree` Treeview.
        It enables the "Load Selected Preset" button if a valid preset is selected
        and the instrument is connected (and not an N9340B, which doesn't support presets).

        Inputs:
            event (tk.Event): The Tkinter event object (not directly used).
        Process:
            1. Checks if any items are selected in the `preset_tree`.
            2. If an item is selected and an instrument is connected (and it's not the N9340B model),
               the `load_preset_button` is enabled.
            3. Otherwise, the `load_preset_button` is disabled.
        Outputs: None
        """
        selected_items = self.preset_tree.selection()
        if selected_items and self.inst and self.instrument_model != "N9340B":
            self.load_preset_button.config(state=tk.NORMAL)
        else:
            self.load_preset_button.config(state=tk.DISABLED)

    def load_selected_preset(self):
        """
        Loads the currently selected preset file from the `preset_tree` onto the
        connected spectrum analyzer. This function acts as a wrapper to the
        `control_load_selected_preset` function in `instrument_control.py`.

        Inputs: None
        Process:
            1. Checks for instrument connection and displays a warning if not connected.
            2. Checks if the instrument model is N9340B, which does not support presets,
               and shows a warning if so.
            3. Retrieves the name of the selected preset from the `preset_tree`.
            4. Calls `control_load_selected_preset()` to send the command to the instrument.
            5. Displays success or error messages to the console and via messagebox.
        Outputs: None
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
        """
        Restores all configurable settings in the GUI to their default values
        as defined in the `DEFAULT_SETTINGS` section of `config.ini`.
        This provides a quick way for users to revert to a known good configuration.

        Inputs: None
        Process:
            1. Prints a message indicating restoration.
            2. Iterates through `self.setting_var_map`.
            3. For each setting, retrieves its default value from `config.ini` and converts it
               to the appropriate Python type.
            4. Sets the corresponding Tkinter variable to this default value.
            5. Handles special updates for slider widgets to reflect the new values.
            6. Resets all frequency band checkboxes to `True` (selected).
            7. Calls `update_vbw_display()` to refresh the VBW based on the restored RBW.
            8. Calls `reset_setting_colors()` to clear any visual indications of unapplied settings.
            9. Calls `_save_config()` to save these restored defaults as the new last-used settings.
            10. Displays an informational messagebox upon completion.
        Outputs: None
        """
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
        """
        Initiates the spectrum scanning process by launching a dedicated thread.
        This prevents the GUI from freezing during long scan operations.
        It performs initial checks, updates GUI button states, and saves the current configuration.

        Inputs: None
        Process:
            1. Checks if an instrument is connected and if a scan is already running.
            2. Saves the current GUI configuration to `config.ini`.
            3. Updates the state of various GUI buttons (Start, Stop, Pause/Resume, Connect, Disconnect, Apply, Load Preset).
            4. Sets `self.scanning` to `True` and `self.paused` to `False`.
            5. Retrieves scan parameters (Max Hold, RBW, Frequency Shift, selected bands).
            6. Creates the output directory if it doesn't exist.
            7. Initializes scan cycle count and frequency offset.
            8. Creates and starts a new `threading.Thread` that will execute the `_run_scan` method.
               The thread is set as a daemon so it terminates with the main application.
        Outputs: None
        """
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
        self._stop_connect_button_blink() # Stop blinking when scan starts

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
        """
        Toggles the paused state of the ongoing scan.
        When paused, the scan thread will temporarily halt its operations.

        Inputs: None
        Process:
            1. Checks if a scan is currently active.
            2. Toggles the `self.paused` boolean flag.
            3. Updates the text and background color of the `pause_resume_button`
               to reflect the current state ("Pause Scan" or "Resume Scan").
            4. Prints status messages to the console.
        Outputs: None
        """
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
        """
        The core logic for performing a continuous spectrum scan. This method runs
        in a separate thread to keep the GUI responsive. It iterates through selected
        frequency bands, performs sweeps, collects data, saves it to CSV, generates plots,
        and manages scan cycles and frequency shifting.

        Inputs:
            selected_bands (list): A list of dictionaries, each representing a frequency band to scan.
            scan_rbw_segmentation (float): The Resolution Bandwidth (RBW) to use for scan segments.
            freq_shift_value (float): The frequency offset (in Hz) to apply per scan cycle.
            rbw_config_val (float): RBW value to configure on the instrument.
            vbw_config_val (float): VBW value to configure on the instrument.
            max_hold_time (float): Duration in seconds for which MAX Hold should be active.

        Process:
            1. Enters a `while self.scanning` loop for continuous operation.
            2. Inside the loop, it checks `self.paused` and `self.scanning` flags
               to allow pausing and stopping the scan.
            3. Increments `self.scan_cycle_count` and applies `self.current_freq_offset`.
            4. Constructs a unique filename for the current scan cycle's raw CSV and HTML plot.
            5. Calls `scan_bands()` from `scan_instrument` to perform the actual instrument sweep
               for the selected bands. This function returns the scanned data, the index of the
               last successfully scanned band, and the path to the CSV file where raw data was saved.
            6. If the scan was not interrupted and data was collected:
               - Converts the scanned data into a Pandas DataFrame.
               - Appends the DataFrame to `self.collected_scans_dataframes` for later averaging.
               - Calls `generate_single_scan_plot_and_open_wrapper()` to create and potentially
                 open an HTML plot for the current cycle's data.
            7. Increments `self.scan_cycle_count` and updates `self.current_freq_offset`.
            8. Resets `self.current_freq_offset` and `self.scan_cycle_count` after 10 cycles.
            9. If `desired_cycle_wait_time_var` is greater than 0, pauses for the specified duration,
               allowing for pause/stop during the wait.
            10. Includes extensive `try-except` blocks to catch and report errors during the scan cycle.
            11. In the `finally` block, ensures `self.scanning` and `self.paused` are reset,
                and calls `reset_scan_buttons()` to update the GUI.
        Outputs: None (modifies `self.collected_scans_dataframes`, generates files, updates GUI)
        """
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
                
                rbw_str = f"RBW{int(scan_rbw_segmentation/1000):04d}"
                max_hold_time_val = float(self.desired_max_hold_time_var.get()) if self.desired_max_hold_var.get() else 0
                hold_str = f"HOLD{int(max_hold_time_val):02d}"
                offset_str = f"Offset{int(self.current_freq_offset)}"

                # IMPORTANT CHANGE: Include seconds in the timestamp for unique CSV/HTML files per cycle
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
            # Start blinking again if resources are available and not connected
            if not self.inst and self.instrument_list and self.resource_var.get() != "No resources found":
                self._start_connect_button_blink()

    def populate_resources(self):
        """
        Discovers and populates the list of available VISA resources (instruments)
        in the `resource_dropdown` menu. It also manages the state of connection-related
        buttons and the blinking effect for the "Connect" button.

        Inputs: None
        Process:
            1. Calls `list_visa_resources()` from `instrument_control` to get available resources.
            2. If resources are found:
               - Sets `resource_var` to the last used device from config, or the first available.
               - Clears and repopulates the `resource_dropdown` menu.
               - Enables the "Connect" button and starts its blinking effect.
            3. If no resources are found, sets `resource_var` to "No resources found"
               and disables the "Connect" button, stopping any blinking.
            4. Disables other scan/instrument control buttons until a connection is made.
            5. Includes error handling for `pyvisa` and general exceptions.
        Outputs: None
        """
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
                
                self.connect_button.config(state=tk.NORMAL)
                self._start_connect_button_blink() # Start blinking if resources are found and not connected
            else:
                self.resource_var.set("No resources found")
                menu = self.resource_dropdown["menu"]
                menu.delete(0, "end")
                menu.add_command(label="No resources found", command=tk._setit(self.resource_var, "No resources found"))
                self.connect_button.config(state=tk.DISABLED)
                self._stop_connect_button_blink() # Stop blinking if no resources
            
            self.start_scan_button.config(state=tk.DISABLED)
            self.disconnect_button.config(state=tk.DISABLED)
            self.apply_button.config(state=tk.DISABLED)
            self.load_preset_button.config(state=tk.DISABLED)
            self.pause_resume_button.config(state=tk.DISABLED)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to list VISA resources: {e}")
            self.resource_var.set("Error listing resources")
            self._reset_gui_on_disconnect_or_error()

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
        if self.blink_id is None: # Only start if not already blinking
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
        self.connect_button.config(bg="darkgrey", fg="white") # Reset to original color
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
            self.blink_id = self.after(500, self._blink_connect_button) # Schedule next blink after 500ms

    def stop_scan(self):
        """
        Initiates the process of stopping the ongoing spectrum scan.
        It sets internal flags that signal the scan thread to terminate gracefully.

        Inputs: None
        Process:
            1. Sets `self.scanning` to `False` to signal the scan thread to stop.
            2. Sets `self.paused` to `False` to ensure the thread isn't stuck in a pause state.
            3. Prints a message to the console indicating the stop attempt.
            4. Disables the "Stop Scan" button and resets the "Pause Scan" button.
        Outputs: None
        """
        self.scanning = False
        self.paused = False
        print("\nAttempting to stop scan... Please wait for current sweep to finish.")
        self.stop_scan_button.config(state=tk.DISABLED)
        self.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED)

    def reset_scan_buttons(self):
        """
        Resets the state of scan-related buttons in the GUI. This function is called
        after a scan completes normally or is stopped/interrupted, making the "Start Scan"
        button available again and disabling others.

        Inputs: None
        Process:
            1. Enables the "Start Scan" button.
            2. If an instrument is connected, enables the "Disconnect" and "Apply Settings" buttons.
            3. If an instrument is connected and a preset is selected (and not N9340B),
               enables the "Load Preset" button.
            4. Disables the "Stop Scan" and "Pause/Resume Scan" buttons.
            5. Resets the text of the "Pause/Resume Scan" button to "Pause Scan".
        Outputs: None
        """
        self.start_scan_button.config(state=tk.NORMAL)
        if self.inst:
            self.disconnect_button.config(state=tk.NORMAL)
            self.apply_button.config(state=tk.NORMAL)
            if self.preset_tree.selection() and self.instrument_model != "N9340B":
                self.load_preset_button.config(state=tk.NORMAL)
        self.stop_scan_button.config(state=tk.DISABLED)
        self.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED)

    def _reset_gui_on_disconnect_or_error(self):
        """
        Resets the state of various GUI elements to reflect a disconnected
        instrument or an error state. This ensures a consistent user experience
        when the connection is lost or problematic.

        Inputs: None
        Process:
            1. Disables all scan-related and instrument control buttons (Start, Stop, Pause/Resume, Disconnect, Apply, Load Preset).
            2. Clears the preset treeview and inserts a "No instrument connected." message.
            3. Enables the "Connect" button.
            4. Manages the blinking effect of the "Connect" button:
               - Starts blinking if resources are found and no instrument is connected.
               - Stops blinking if no resources are found or if an instrument is already connected (though this function implies disconnection).
        Outputs: None
        """
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
        # Start blinking if resources are found after a reset/error and not connected
        if self.instrument_list and self.resource_var.get() != "No resources found":
            self._start_connect_button_blink()
        else:
            self._stop_connect_button_blink() # Ensure blinking is off if no resources or already connected

    def generate_single_scan_plot_and_open_wrapper(self, csv_file_path, plot_title_suffix, output_html_path, auto_open_browser=True):
        """
        A wrapper function to facilitate plotting of a single scan's data.
        It reads the scan data from a specified CSV file, prepares it for plotting,
        and then calls `plot_single_scan_data` from `plotting_utils` to generate
        and potentially open the HTML plot.

        Inputs:
            csv_file_path (str): The full path to the CSV file containing the scan data.
            plot_title_suffix (str): A string to append to the plot's main title.
            output_html_path (str): The full path where the generated HTML plot should be saved.
            auto_open_browser (bool, optional): If True, the generated HTML plot will be
                                                automatically opened in the default web browser. Defaults to True.
        Process:
            1. Reads the scan data from the `csv_file_path` into a Pandas DataFrame,
               assuming no header and specific column names.
            2. Converts frequencies from MHz back to Hz, as `plot_single_scan_data` expects Hz.
            3. Calls `plot_single_scan_data` with the prepared data and plotting options.
            4. Prints success or error messages to the console and displays a messagebox on error.
        Outputs: None (generates an HTML file, potentially opens a browser)
        """
        # The 'output_html_path' parameter is already defined in the function signature.
        # The previous debug print and check are no longer needed as the saving logic
        # is now handled directly by plot_single_scan_data.

        print(f"Generating single scan plot from CSV: {csv_file_path}...")
        try:
            # Read data from the CSV file
            # Assuming no header and columns are Frequency_MHz, Power_dBm
            df = pd.read_csv(csv_file_path, header=None, names=["Frequency_MHz", "Power_dBm"])
            
            # Convert Frequency_MHz back to Frequency_Hz for plot_single_scan_data
            # which expects Frequency_Hz as its first column.
            scanned_data_from_csv = list(zip(df['Frequency_MHz'] * MHZ_TO_HZ, df['Power_dBm']))

            # Call plot_single_scan_data, which now handles saving and opening
            fig, saved_html_path = plot_single_scan_data(
                scanned_data_from_csv, 
                plot_title_suffix,
                include_tv_markers=self.include_tv_markers_var.get(),
                include_gov_markers=self.include_gov_markers_var.get(),
                output_html_path=output_html_path, # Pass the output path
                auto_open_browser=auto_open_browser # Pass the auto_open flag
            )
            
            if fig and saved_html_path:
                print(f"✅ Single scan plot generation complete: {saved_html_path}")
            else:
                print("🚫 Plotly figure was not generated or saved for single scan data.")
        except Exception as e:
            messagebox.showerror("Single Plot Error", f"Failed to generate single scan plot from CSV '{csv_file_path}': {e}")
            print(f"❌ Error generating single scan plot from CSV: {e}")

    def generate_average_plot(self):
        """
        Triggers the generation of a historical average, median, range, standard deviation,
        variance, and PSD plot from all relevant CSV files found in the current output folder.
        This plot also includes individual historical scans as overlay layers.
        This function acts as a wrapper to `generate_historical_average_plot` from `averaging_utils`.

        Inputs: None
        Process:
            1. Checks if a scan is currently in progress and warns the user if so.
            2. Calls `generate_historical_average_plot()` with the necessary Tkinter variables
               to retrieve scan name, output folder, and plotting options.
        Outputs: None (generates HTML plots and CSV files)
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
                # Magic: This line deletes the content of the last line.
                # "end-1c linestart" refers to the beginning of the last line.
                # "end-1c" refers to the character just before the very end of the text.
                # Together, they select the entire last line.
                self.console_output.delete("end-1c linestart", "end-1c")
            except TclError:
                # This can happen if the console is empty or the last line is not fully formed yet.
                pass
        self.console_output.insert(tk.END, text_to_display)
        self.console_output.see(tk.END)
        self.console_output.config(state=tk.DISABLED)
        self.console_output.update_idletasks()


def print_art():
    """
    Prints an ASCII art logo to the console output. This function is called
    during application startup to provide a visual brand element.

    Inputs: None
    Process:
        1. Uses a series of `print()` statements to output the multi-line ASCII art.
    Outputs: None (prints to console)
    """
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
