# src/scantab.py
import tkinter as tk
from tkinter import ttk, filedialog # Import filedialog
import inspect
import os # Import os for path manipulation
import subprocess # Import subprocess for opening directories

# Import debug_print from utils
from utils.instrument_control import debug_print
# Import restore_default_settings_logic from src.settings_logic
from src.settings_logic import restore_default_settings_logic
# Import save_config from config_manager
from src.config_manager import save_config

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

        # Configure columns to expand
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Scan Session Details Frame
        session_details_frame = ttk.LabelFrame(self, text="Scan Session Details", padding="10 10 10 10", style='Dark.TLabelframe')
        session_details_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        session_details_frame.grid_columnconfigure(0, weight=1)
        session_details_frame.grid_columnconfigure(1, weight=2)
        session_details_frame.grid_columnconfigure(2, weight=0) # For browse button
        # Add a row for the "Open Directory" button
        session_details_frame.grid_rowconfigure(2, weight=0) 


        ttk.Label(session_details_frame, text="Session Name:", style='TLabel').grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(session_details_frame, textvariable=self.app_instance.scan_name_var, style='TEntry').grid(row=0, column=1, columnspan=2, padx=2, pady=2, sticky="ew")

        ttk.Label(session_details_frame, text="Output Directory:", style='TLabel').grid(row=1, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(session_details_frame, textvariable=self.app_instance.output_folder_var, style='TEntry').grid(row=1, column=1, padx=2, pady=2, sticky="ew")
        ttk.Button(session_details_frame, text="Browse", command=self._browse_output_directory, style='TButton').grid(row=1, column=2, padx=2, pady=2, sticky="ew")

        # New: Open Directory Button
        ttk.Button(session_details_frame, text="Open Directory", command=self._open_output_directory, style='TButton').grid(row=2, column=0, columnspan=3, padx=2, pady=5, sticky="ew")


        # Instrument Settings Frame
        instrument_settings_frame = ttk.LabelFrame(self, text="Scan Configuration Settings", padding="10 10 10 10", style='Dark.TLabelframe')
        instrument_settings_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        instrument_settings_frame.grid_columnconfigure(0, weight=1)
        instrument_settings_frame.grid_columnconfigure(1, weight=1)

        # RBW Step Size
        ttk.Label(instrument_settings_frame, text="RBW Step Size (Hz):", style='TLabel').grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.rbw_step_size_hz_var, style='TEntry').grid(row=0, column=1, padx=2, pady=2, sticky="ew")

        # Cycle Wait Time
        ttk.Label(instrument_settings_frame, text="Cycle Wait Time (s):", style='TLabel').grid(row=1, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.cycle_wait_time_seconds_var, style='TEntry').grid(row=1, column=1, padx=2, pady=2, sticky="ew")

        # Max Hold Time
        ttk.Label(instrument_settings_frame, text="Max Hold Time (s):", style='TLabel').grid(row=2, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.maxhold_time_seconds_var, style='TEntry').grid(row=2, column=1, padx=2, pady=2, sticky="ew")

        # Scan RBW
        ttk.Label(instrument_settings_frame, text="Scan RBW (Hz):", style='TLabel').grid(row=3, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.scan_rbw_hz_var, style='TEntry').grid(row=3, column=1, padx=2, pady=2, sticky="ew")

        # Reference Level
        ttk.Label(instrument_settings_frame, text="Reference Level (dBm):", style='TLabel').grid(row=4, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.reference_level_dbm_var, style='TEntry').grid(row=4, column=1, padx=2, pady=2, sticky="ew")

        # Frequency Shift
        ttk.Label(instrument_settings_frame, text="Frequency Shift (Hz):", style='TLabel').grid(row=5, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.freq_shift_hz_var, style='TEntry').grid(row=5, column=1, padx=2, pady=2, sticky="ew")

        # Max Hold Enabled Checkbox
        ttk.Checkbutton(instrument_settings_frame, text="Max Hold Enabled", variable=self.app_instance.max_hold_enabled_var, style='TCheckbutton').grid(row=6, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        # High Sensitivity/Preamp Checkbox
        ttk.Checkbutton(instrument_settings_frame, text="High Sensitivity (Preamp On)", variable=self.app_instance.high_sensitivity_var, style='TCheckbutton').grid(row=7, column=0, columnspan=2, padx=5, pady=2, sticky="w")

        # RBW Segmentation
        ttk.Label(instrument_settings_frame, text="RBW Segmentation:", style='TLabel').grid(row=8, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.rbw_segmentation_var, style='TEntry').grid(row=8, column=1, padx=2, pady=2, sticky="ew")

        # Default Focus Width
        ttk.Label(instrument_settings_frame, text="Default Focus Width:", style='TLabel').grid(row=9, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.desired_default_focus_width_var, style='TEntry').grid(row=9, column=1, padx=2, pady=2, sticky="ew")

        # Number of Scan Cycles
        ttk.Label(instrument_settings_frame, text="Number of Scan Cycles:", style='TLabel').grid(row=10, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(instrument_settings_frame, textvariable=self.app_instance.num_scan_cycles_var, style='TEntry').grid(row=10, column=1, padx=2, pady=2, sticky="ew")

        # Frequency Band Selection Frame
        # Adjusted row to 2 as plotting options frame is removed
        band_selection_frame = ttk.LabelFrame(self, text="Frequency Band Selection", padding="10 10 10 10", style='Dark.TLabelframe')
        band_selection_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        band_selection_frame.grid_columnconfigure(0, weight=1)
        band_selection_frame.grid_rowconfigure(0, weight=1) # Allow checkbox frame to expand

        # Inner frame for band checkboxes to allow scrolling if many bands
        band_checkbox_frame = ttk.Frame(band_selection_frame, style='Dark.TFrame')
        band_checkbox_frame.grid(row=0, column=0, sticky="nsew")
        band_checkbox_frame.grid_columnconfigure(0, weight=1)

        # Create checkboxes for each band
        for i, band_item in enumerate(self.app_instance.band_vars):
            cb = ttk.Checkbutton(band_checkbox_frame, text=band_item["band"]["Band Name"], variable=band_item["var"], style='TCheckbutton')
            cb.grid(row=i, column=0, sticky="w", padx=2, pady=1)

        # Restore Defaults Button
        # Adjusted row to 3
        restore_defaults_button = ttk.Button(self, text="Restore Default Settings", command=self._restore_default_settings, style='TButton')
        restore_defaults_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        # New: Save Last Used to Config Button
        # Adjusted row to 4
        save_last_used_button = ttk.Button(self, text="Save Last Used to Config", command=self._save_last_used_config, style='TButton')
        save_last_used_button.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="ew")


        debug_print("ScanTab widgets created.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _browse_output_directory(self):
        """Opens a dialog to select the scan output directory."""
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        initial_dir = self.app_instance.output_folder_var.get()
        if not os.path.isdir(initial_dir):
            initial_dir = os.getcwd() # Fallback to current working directory

        folder_selected = filedialog.askdirectory(initialdir=initial_dir, title="Select Scan Output Directory")
        if folder_selected:
            self.app_instance.output_folder_var.set(folder_selected)
            self.console_print_func(f"✅ Output directory set to: {folder_selected}")
            debug_print(f"Output directory set to: {folder_selected}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.console_print_func("ℹ️ Output directory selection cancelled.")
            debug_print("Output directory selection cancelled.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _open_output_directory(self):
        """Opens the currently set output directory in the file explorer."""
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        output_path = self.app_instance.output_folder_var.get()

        if not output_path:
            self.console_print_func("⚠️ Warning: Output directory is not set.")
            debug_print("Output directory is not set.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        if not os.path.isdir(output_path):
            self.console_print_func(f"❌ Error: Directory does not exist: {output_path}")
            debug_print(f"Directory does not exist: {output_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        try:
            if sys.platform == "win32":
                os.startfile(output_path)
            elif sys.platform == "darwin": # macOS
                subprocess.Popen(["open", output_path])
            else: # Linux
                subprocess.Popen(["xdg-open", output_path])
            self.console_print_func(f"✅ Opened directory: {output_path}")
            debug_print(f"Opened directory: {output_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        except Exception as e:
            self.console_print_func(f"❌ Error opening directory: {e}")
            debug_print(f"Error opening directory '{output_path}': {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _save_last_used_config(self):
        """
        Saves the current application settings to config.ini.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Saving last used settings from ScanTab...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        save_config(self.app_instance)
        self.console_print_func("✅ Current settings saved to config.ini.")


    def _restore_default_settings(self):
        """
        Restores all settings to their default values as defined in config.ini.
        This function is now part of the ScanTab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Restoring default settings from ScanTab... (delegating to logic)", file=current_file, function=current_function, console_print_func=self.console_print_func)
        restore_default_settings_logic(self.app_instance, self.console_print_func)


    def _on_tab_selected(self, event):
        """
        Called when this tab is selected in the notebook.
        Can be used to refresh data or update UI elements specific to this tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("ScanTab selected.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        # No specific refresh logic needed for ScanTab upon selection for now.
        pass
