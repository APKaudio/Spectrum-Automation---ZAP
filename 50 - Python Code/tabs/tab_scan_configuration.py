# src/scantab.py
import tkinter as tk
from tkinter import ttk, filedialog
import inspect
import os
import subprocess
import sys # Import the sys module to fix NameError

from utils.instrument_control import debug_print
from src.settings_logic import restore_default_settings_logic, restore_last_used_settings_logic # Import the new logic function
from src.config_manager import save_config
from utils.frequency_bands import MHZ_TO_HZ # Import for display formatting

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
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self.console_print_func = console_print_func if console_print_func else print

        self._create_widgets()

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Scan Configuration tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Creating widgets for ScanTab...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # Configure grid for this tab
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0) # Settings frame
        self.grid_rowconfigure(1, weight=1) # Bands frame
        self.grid_rowconfigure(2, weight=0) # New frame for buttons

        # --- Scan Settings Frame ---
        settings_frame = ttk.LabelFrame(self, text="Scan Settings", style='Dark.TLabelframe')
        settings_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        settings_frame.grid_columnconfigure(1, weight=1) # Make entry fields expand

        # Helper to create labeled entry/checkbox rows
        def create_setting_row(parent, label_text, tk_var, row_num, is_checkbox=False, is_button=False, command=None):
            if is_button:
                button = ttk.Button(parent, text=label_text, command=command, style='Blue.TButton')
                button.grid(row=row_num, column=0, columnspan=2, pady=2, padx=5, sticky="ew")
                return button
            elif is_checkbox:
                checkbox = ttk.Checkbutton(parent, text=label_text, variable=tk_var, style='TCheckbutton')
                checkbox.grid(row=row_num, column=0, columnspan=2, pady=2, padx=5, sticky="w")
                return checkbox
            else:
                ttk.Label(parent, text=label_text, style='TLabel').grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
                entry = ttk.Entry(parent, textvariable=tk_var, style='TEntry')
                entry.grid(row=row_num, column=1, sticky="ew", padx=5, pady=2)
                return entry

        row = 0
        create_setting_row(settings_frame, "RBW Step Size (Hz):", self.app_instance.rbw_step_size_hz_var, row); row += 1
        create_setting_row(settings_frame, "Cycle Wait Time (s):", self.app_instance.cycle_wait_time_seconds_var, row); row += 1
        create_setting_row(settings_frame, "Max Hold Time (s):", self.app_instance.maxhold_time_seconds_var, row); row += 1
        create_setting_row(settings_frame, "Scan RBW (Hz):", self.app_instance.scan_rbw_hz_var, row); row += 1
        create_setting_row(settings_frame, "Reference Level (dBm):", self.app_instance.reference_level_dbm_var, row); row += 1
        create_setting_row(settings_frame, "Frequency Shift (Hz):", self.app_instance.freq_shift_hz_var, row); row += 1
        create_setting_row(settings_frame, "Max Hold Enabled:", self.app_instance.maxhold_enabled_var, row, is_checkbox=True); row += 1
        create_setting_row(settings_frame, "High Sensitivity:", self.app_instance.high_sensitivity_var, row, is_checkbox=True); row += 1
        create_setting_row(settings_frame, "Preamp On:", self.app_instance.preamp_on_var, row, is_checkbox=True); row += 1
        create_setting_row(settings_frame, "Scan RBW Segmentation (Hz):", self.app_instance.scan_rbw_segmentation_var, row); row += 1
        create_setting_row(settings_frame, "Default Focus Width (Hz):", self.app_instance.desired_default_focus_width_var, row); row += 1
        create_setting_row(settings_frame, "Number of Scan Cycles:", self.app_instance.num_scan_cycles_var, row); row += 1
        create_setting_row(settings_frame, "Scan Name:", self.app_instance.scan_name_var, row); row += 1
        
        # "Scan Export Folder" label and "Choose output directory" button
        ttk.Label(settings_frame, text="Scan Export Folder:", style='TLabel').grid(row=row, column=0, sticky="w", padx=5, pady=2)
        output_folder_entry = ttk.Entry(settings_frame, textvariable=self.app_instance.output_folder_var, style='TEntry')
        output_folder_entry.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
        choose_dir_button = ttk.Button(settings_frame, text="Choose output directory", command=self.app_instance._browse_output_folder, style='Blue.TButton')
        choose_dir_button.grid(row=row, column=2, padx=5, pady=2) # Place browse button in a new column
        settings_frame.grid_columnconfigure(2, weight=0) # Don't let browse button column expand
        row += 1

        # --- START OF MODIFICATION ---
        # Add "Browse Export Folder" button, now opening the folder directly
        ttk.Button(settings_frame, text="Browse Export Folder", command=self._open_export_folder, style='Blue.TButton').grid(row=row, column=0, columnspan=3, pady=2, padx=5, sticky="ew")
        row += 1
        # --- END OF MODIFICATION ---

        # --- Bands to Scan Frame ---
        bands_frame = ttk.LabelFrame(self, text="Bands to Scan", style='Dark.TLabelframe')
        bands_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        bands_frame.grid_columnconfigure(0, weight=1) # Allow band list to expand

        # Create a canvas and scrollbar for the band selection checkboxes
        bands_canvas = tk.Canvas(bands_frame, bg="#2b2b2b", highlightthickness=0)
        bands_canvas.grid(row=0, column=0, sticky="nsew")

        bands_scrollbar = ttk.Scrollbar(bands_frame, orient="vertical", command=bands_canvas.yview)
        bands_scrollbar.grid(row=0, column=1, sticky="ns")
        bands_canvas.configure(yscrollcommand=bands_scrollbar.set)

        self.bands_inner_frame = ttk.Frame(bands_canvas, style='Dark.TFrame')
        bands_canvas.create_window((0, 0), window=self.bands_inner_frame, anchor="nw", width=bands_canvas.winfo_width())
        self.bands_inner_frame.bind("<Configure>", lambda e: bands_canvas.configure(scrollregion=bands_canvas.bbox("all")))
        bands_canvas.bind('<Configure>', lambda e: bands_canvas.itemconfig(bands_canvas.find_withtag("all")[0], width=e.width))

        self.bands_inner_frame.grid_columnconfigure(0, weight=1) # Allow checkboxes to expand horizontally

        # Populate bands
        for i, band_item in enumerate(self.app_instance.band_vars):
            band_info = band_item["band"]
            band_var = band_item["var"]
            
            # Format the display text for each band
            display_text = (
                f"{band_info['Band Name']} - "
                f"Start: {band_info['Start MHz']:.3f} MHz "
                f"Stop: {band_info['Stop MHz']:.3f} MHz"
            )

            cb = ttk.Checkbutton(self.bands_inner_frame, text=display_text, variable=band_var, style='TCheckbutton')
            cb.grid(row=i, column=0, sticky="w", padx=5, pady=2)

        # New frame for buttons below the Bands to Scan box
        button_frame = ttk.LabelFrame(self, text="Configuration Actions", style='Dark.TLabelframe')
        button_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        button_frame.grid_columnconfigure(0, weight=1) # Allow buttons to expand

        # Add buttons to the new frame
        ttk.Button(button_frame, text="Save Current Settings", command=self._save_current_settings, style='Green.TButton').grid(row=0, column=0, pady=2, padx=5, sticky="ew")
        ttk.Button(button_frame, text="Restore Last Used Settings", command=self._restore_last_used_settings, style='Orange.TButton').grid(row=1, column=0, pady=2, padx=5, sticky="ew")
        ttk.Button(button_frame, text="Restore Default Settings", command=self._restore_default_settings, style='Red.TButton').grid(row=2, column=0, pady=2, padx=5, sticky="ew")
        ttk.Button(button_frame, text="Open Config.ini", command=self._open_config_file, style='Blue.TButton').grid(row=3, column=0, pady=2, padx=5, sticky="ew")

        debug_print("ScanTab widgets created.", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _save_current_settings(self):
        """
        Saves the current settings from the GUI elements to config.ini.
        This function is now part of the ScanTab.
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

    def _restore_last_used_settings(self):
        """
        Restores all settings to their last-used values as defined in config.ini.
        This function is now part of the ScanTab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Restoring last used settings from ScanTab... (delegating to logic)", file=current_file, function=current_function, console_print_func=self.console_print_func)
        restore_last_used_settings_logic(self.app_instance, self.console_print_func)

    def _open_config_file(self):
        """
        Opens the config.ini file using the default application associated with .ini files.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        config_file_path = self.app_instance.CONFIG_FILE
        debug_print(f"Attempting to open config file: {config_file_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if not os.path.exists(config_file_path):
            self.console_print_func(f"❌ Error: config.ini not found at {config_file_path}")
            debug_print(f"config.ini not found: {config_file_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        try:
            if sys.platform == "win32":
                os.startfile(config_file_path)
            elif sys.platform == "darwin": # macOS
                subprocess.Popen(["open", config_file_path])
            else: # Linux
                subprocess.Popen(["xdg-open", config_file_path])
            self.console_print_func(f"✅ Opened config.ini: {config_file_path}")
            debug_print(f"Opened config.ini: {config_file_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        except Exception as e:
            self.console_print_func(f"❌ Failed to open config.ini: {e}")
            debug_print(f"Error opening config.ini: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)

    # --- START OF MODIFICATION ---
    def _open_export_folder(self):
        """
        Opens the currently configured Scan Export Folder using the OS's default file explorer.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        export_folder_path = self.app_instance.output_folder_var.get()
        debug_print(f"Attempting to open export folder: {export_folder_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if not os.path.exists(export_folder_path):
            self.console_print_func(f"❌ Error: Export folder not found at {export_folder_path}. Please choose a valid directory.")
            debug_print(f"Export folder not found: {export_folder_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        try:
            if sys.platform == "win32":
                os.startfile(export_folder_path)
            elif sys.platform == "darwin": # macOS
                subprocess.Popen(["open", export_folder_path])
            else: # Linux
                subprocess.Popen(["xdg-open", export_folder_path])
            self.console_print_func(f"✅ Opened export folder: {export_folder_path}")
            debug_print(f"Opened export folder: {export_folder_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        except Exception as e:
            self.console_print_func(f"❌ Failed to open export folder: {e}")
            debug_print(f"Error opening export folder: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
    # --- END OF MODIFICATION ---

    def _on_tab_selected(self, event):
        """
        Called when this tab is selected in the notebook.
        Can be used to refresh data or update UI elements specific to this tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Scan Tab selected.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        # Ensure bands_inner_frame width is updated when tab is selected
        # This helps with proper wrapping and display of band names
        self.bands_inner_frame.update_idletasks()
        if self.bands_inner_frame.winfo_width() > 1: # Avoid division by zero or initial small width
            # This will trigger the canvas's create_window to update its width
            self.bands_inner_frame.event_generate("<Configure>", width=self.bands_inner_frame.winfo_width())

        # When the tab is selected, ensure the band checkboxes reflect the current state
        # This is primarily handled by load_config on app startup, but this ensures consistency
        # if settings are changed elsewhere or if the tab is revisited.
        self._load_band_selections_from_config()


    def _load_band_selections_from_config(self):
        """
        Loads the selected bands from config.ini into the checkboxes.
        This is called on tab selection to ensure the GUI reflects the config.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Loading band selections from config.ini...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        
        # Get the last used selected bands from config
        last_selected_bands_str = self.app_instance.config.get('LAST_USED_SETTINGS', 'last_selected_bands', fallback='')
        last_selected_band_names = [name.strip() for name in last_selected_bands_str.split(',') if name.strip()]

        # Set the state of each checkbox based on the loaded bands
        for band_item in self.app_instance.band_vars:
            band_name = band_item["band"]["Band Name"]
            band_item["var"].set(band_name in last_selected_band_names)
        debug_print(f"Loaded selected bands from config: {last_selected_band_names}", file=current_file, function=current_function, console_print_func=self.console_print_func)

