# src/scantab.py
import tkinter as tk
from tkinter import ttk
import inspect

# Import debug_print from utils
from utils.instrument_control import debug_print
# Import restore_default_settings_logic from src.settings_logic
from src.settings_logic import restore_default_settings_logic

class ScanTab(ttk.Frame):
    """
    A Tkinter Frame that contains the Scan Configuration settings
    and the band selection checkboxes.
    """
    def __init__(self, master=None, app_instance=None, console_print_func=None, **kwargs):
        """
        Initializes the ScanTab.

        Inputs:
            master (tk.Widget): The parent widget (the ttk.Notebook).
            app_instance (App): The main application instance, used for accessing
                                shared state like Tkinter variables and console print function.
            console_print_func (function): Function to print messages to the GUI console.
            **kwargs: Arbitrary keyword arguments for Tkinter Frame.
        """
        # Do NOT pass console_print_func to super().__init__ as it's not a ttk.Frame option
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self.console_print_func = console_print_func if console_print_func else print # Get console print function

        self._create_widgets()

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Scan Configuration tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Creating ScanTab widgets...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self.grid_columnconfigure(0, weight=1) # Allow elements to expand

        # --- Instrument settings for scan Frame ---
        instrument_settings_frame = ttk.LabelFrame(self, text="Instrument settings for scan", style='Dark.TLabelframe')
        instrument_settings_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        instrument_settings_frame.grid_columnconfigure(0, weight=1) # Allow labels/entries to expand
        instrument_settings_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(instrument_settings_frame, text="RBW Step Size (Hz):", style='TLabel').grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.rbw_step_size_hz_var, style='TEntry').grid(row=0, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(instrument_settings_frame, text="Cycle Wait Time (s):", style='TLabel').grid(row=1, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.cycle_wait_time_seconds_var, style='TEntry').grid(row=1, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(instrument_settings_frame, text="Max Hold Time (s):", style='TLabel').grid(row=2, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.maxhold_time_seconds_var, style='TEntry').grid(row=2, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(instrument_settings_frame, text="Scan RBW (Hz):", style='TLabel').grid(row=3, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.scan_rbw_hz_var, style='TEntry').grid(row=3, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(instrument_settings_frame, text="Reference Level (dBm):", style='TLabel').grid(row=4, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.reference_level_dbm_var, style='TEntry').grid(row=4, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(instrument_settings_frame, text="Frequency Shift (Hz):", style='TLabel').grid(row=5, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.freq_shift_hz_var, style='TEntry').grid(row=5, column=1, padx=2, pady=2, sticky="ew")

        ttk.Checkbutton(instrument_settings_frame, text="Max Hold Enabled", variable=self.app_instance.maxhold_enabled_var, style='TCheckbutton').grid(row=6, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(instrument_settings_frame, text="High Sensitivity", variable=self.app_instance.high_sensitivity_var, style='TCheckbutton').grid(row=7, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(instrument_settings_frame, text="Preamp On", variable=self.app_instance.preamp_on_var, style='TCheckbutton').grid(row=8, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        
        ttk.Label(instrument_settings_frame, text="RBW Segmentation (Hz):", style='TLabel').grid(row=9, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.scan_rbw_segmentation_var, style='TEntry').grid(row=9, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(instrument_settings_frame, text="Default Focus Width (MHz):", style='TLabel').grid(row=10, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.desired_default_focus_width_var, style='TEntry').grid(row=10, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(instrument_settings_frame, text="Number of Scan Cycles:", style='TLabel').grid(row=11, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.num_scan_cycles_var, style='TEntry').grid(row=11, column=1, padx=2, pady=2, sticky="ew")

        # Restore Defaults Button (now within ScanTab)
        ttk.Button(self, text="Restore Default Settings", command=self._restore_default_settings, style='Orange.TButton').grid(row=2, column=0, padx=5, pady=5, sticky="ew") # Placed below both frames

        # --- Select Bands to Scan Frame ---
        bands_frame = ttk.LabelFrame(self, text="Select Bands to Scan:", style='Dark.TLabelframe')
        bands_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew") # Placed below instrument_settings_frame
        bands_frame.grid_columnconfigure(0, weight=1) # Allow band checkboxes to expand
        bands_frame.grid_rowconfigure(0, weight=1) # Allow canvas to expand

        # Use a canvas and scrollbar for the bands if there are many
        bands_canvas = tk.Canvas(bands_frame, background="#1e1e1e", highlightthickness=0)
        bands_canvas.grid(row=0, column=0, sticky="nsew")
        
        bands_scrollbar = ttk.Scrollbar(bands_frame, orient="vertical", command=bands_canvas.yview)
        bands_scrollbar.grid(row=0, column=1, sticky="ns")
        
        bands_canvas.configure(yscrollcommand=bands_scrollbar.set)
        # Bind the canvas to the frame's configure event to update scrollregion
        bands_canvas.bind('<Configure>', lambda e: bands_canvas.configure(scrollregion = bands_canvas.bbox("all")))

        band_checkbox_frame = ttk.Frame(bands_canvas, style='Dark.TFrame')
        bands_canvas.create_window((0, 0), window=band_checkbox_frame, anchor="nw")
        
        band_checkbox_frame.grid_columnconfigure(0, weight=1) # Allow checkboxes to expand
        
        for i, band_item in enumerate(self.app_instance.band_vars):
            cb = ttk.Checkbutton(band_checkbox_frame, text=band_item["band"]["Band Name"], variable=band_item["var"], style='TCheckbutton')
            cb.grid(row=i, column=0, sticky="w", padx=2, pady=1)

        debug_print("ScanTab widgets created.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _restore_default_settings(self):
        """
        Restores all settings to their default values as defined in config.ini.
        This function is now part of the ScanTab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Restoring default settings from ScanTab...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        restore_default_settings_logic(self.app_instance)

    def _on_tab_selected(self, event):
        """
        Called when this tab is selected in the notebook.
        Can be used to refresh data or update UI elements specific to this tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("ScanTab selected.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        # Example: if you had data specific to this tab that needed refreshing
        # self.load_scan_settings()
