# src/instrument_preset_tab.py
import tkinter as tk
# from tkinter import messagebox, scrolledtext, filedialog, ttk # Removed messagebox
from tkinter import scrolledtext, filedialog, ttk # Keep other imports
import os
import sys
import inspect
import subprocess # Add this import for opening folders

# Import instrument_logic for setting focus frequency and loading presets
from src.instrument_logic import (
    load_selected_preset_logic,
    query_device_presets_logic
)
from utils.instrument_control import debug_print
from utils.frequency_bands import MHZ_TO_HZ # Import MHZ_TO_HZ for conversion

class PresetFilesTab(ttk.Frame):
    """
    A Tkinter Frame that displays available instrument preset files and allows loading them.
    """
    def __init__(self, master=None, app_instance=None, console_print_func=None, **kwargs):
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self.console_print_func = console_print_func if console_print_func else print # Use provided func or default print
        self.preset_files = [] # To store list of .sta files (either local or device)
        self.selected_preset = None # To store the currently selected preset name
        self.source_of_displayed_presets = "local" # "local" or "device"
        self.preset_buttons = {} # Dictionary to store references to preset buttons for dynamic updates
        self.current_selected_button = None # To keep track of the currently selected button widget

        self._create_widgets()
        # Initial population will happen in _on_tab_selected,
        # which is called when the tab is first shown.

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Preset Files tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Creating PresetFilesTab widgets.", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # Row for preset buttons frame

        # Control Frame for buttons (Load from Device, Open Folder)
        control_frame = ttk.Frame(self, style='TFrame')
        control_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        control_frame.grid_columnconfigure(0, weight=1)
        control_frame.grid_columnconfigure(1, weight=1)

        self.query_presets_button = ttk.Button(control_frame, text="Query Presets from Device", command=self._query_device_presets_threaded, style='Blue.TButton')
        self.query_presets_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.query_presets_button.config(state=tk.DISABLED) # Initially disabled until connected

        self.open_folder_button = ttk.Button(control_frame, text="Open Preset Folder", command=self._open_preset_folder, style='Blue.TButton')
        self.open_folder_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Frame to hold preset buttons (will be populated dynamically)
        self.preset_buttons_frame = ttk.LabelFrame(self, text="Available Presets", style='Dark.TLabelframe')
        self.preset_buttons_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        self.preset_buttons_frame.grid_columnconfigure(0, weight=1) # Allow buttons to expand horizontally

        # Canvas and Scrollbar for preset buttons
        self.preset_canvas = tk.Canvas(self.preset_buttons_frame, borderwidth=0, highlightthickness=0, bg="#1e1e1e")
        self.preset_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.preset_scrollbar = ttk.Scrollbar(self.preset_buttons_frame, orient="vertical", command=self.preset_canvas.yview)
        self.preset_scrollbar.pack(side=tk.RIGHT, fill="y")

        self.preset_canvas.configure(yscrollcommand=self.preset_scrollbar.set)
        self.preset_canvas.bind('<Configure>', lambda e: self.preset_canvas.configure(scrollregion = self.preset_canvas.bbox("all")))

        self.inner_preset_frame = ttk.Frame(self.preset_canvas, style='Dark.TFrame')
        self.preset_canvas.create_window((0, 0), window=self.inner_preset_frame, anchor="nw")

        self._populate_local_presets() # Initial population with local files

    def _on_tab_selected(self, event):
        """
        Callback for when this tab is selected.
        Refreshes the list of local presets.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Preset Files Tab selected. Refreshing local presets.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        self._populate_local_presets()
        # Update the state of the query_presets_button based on instrument connection
        if self.app_instance and self.app_instance.inst:
            self.query_presets_button.config(state=tk.NORMAL)
        else:
            self.query_presets_button.config(state=tk.DISABLED)


    def _populate_local_presets(self):
        """
        Populates the display with .sta files found in the local preset directory.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        self.source_of_displayed_presets = "local"
        
        preset_folder = os.path.join(os.getcwd(), "presets") # Assuming a 'presets' folder in the app directory
        if not os.path.exists(preset_folder):
            os.makedirs(preset_folder, exist_ok=True)
            self.console_print_func(f"Created local preset folder: {preset_folder}")
            debug_print(f"Created local preset folder: {preset_folder}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self.preset_files = []
        try:
            for f_name in os.listdir(preset_folder):
                if f_name.lower().endswith(".sta"):
                    self.preset_files.append(f_name)
            self.preset_files.sort() # Sort alphabetically
            debug_print(f"Found {len(self.preset_files)} local presets.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        except Exception as e:
            self.console_print_func(f"❌ Error listing local presets: {e}")
            debug_print(f"Error listing local presets: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        
        self._create_clickable_buttons()


    def _query_device_presets_threaded(self):
        """
        Starts a thread to query presets from the connected device.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Starting thread to query device presets...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        
        if not self.app_instance.inst:
            self.console_print_func("⚠️ Warning: No instrument connected to query presets from.")
            debug_print("No instrument connected for device preset query.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        self.console_print_func("Querying presets from device... This may take a moment.")
        # Disable button to prevent multiple queries
        self.query_presets_button.config(state=tk.DISABLED)
        
        query_thread = threading.Thread(target=self._run_device_preset_query)
        query_thread.daemon = True
        query_thread.start()

    def _run_device_preset_query(self):
        """
        Executes the device preset query logic in a separate thread.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Running device preset query in thread.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        
        try:
            device_presets = query_device_presets_logic(self.app_instance, self.console_print_func)
            self.app_instance.after(0, lambda: self._update_preset_display_from_device(device_presets))
        except Exception as e:
            self.app_instance.after(0, lambda: self.console_print_func(f"❌ Error querying device presets: {e}"))
            debug_print(f"Error in _run_device_preset_query: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        finally:
            self.app_instance.after(0, lambda: self.query_presets_button.config(state=tk.NORMAL)) # Re-enable button


    def _update_preset_display_from_device(self, presets):
        """
        Updates the GUI to display presets queried from the device.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Updating preset display with {len(presets)} device presets.", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self.source_of_displayed_presets = "device"
        self.preset_files = presets
        self.preset_files.sort() # Ensure sorted display
        self._create_clickable_buttons()
        self.console_print_func(f"Displaying {len(self.preset_files)} presets from the device.")


    def _create_clickable_buttons(self):
        """
        Clears existing preset buttons and creates new ones based on self.preset_files.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Creating clickable preset buttons...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # Clear existing buttons
        for widget in self.inner_preset_frame.winfo_children():
            widget.destroy()
        self.preset_buttons = {}
        self.current_selected_button = None

        if not self.preset_files:
            ttk.Label(self.inner_preset_frame, text="No presets found.", style='Markers.TLabel').grid(row=0, column=0, padx=5, pady=5)
            debug_print("No preset files to create buttons for.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        for i, preset_name in enumerate(self.preset_files):
            button = ttk.Button(self.inner_preset_frame, text=preset_name,
                                command=lambda p=preset_name: self._on_preset_button_click(p),
                                style='Markers.TButton') # Use a generic button style
            button.grid(row=i, column=0, sticky="ew", padx=2, pady=2)
            self.preset_buttons[preset_name] = button
        
        # Update scroll region after adding buttons
        self.inner_preset_frame.update_idletasks()
        self.preset_canvas.config(scrollregion=self.preset_canvas.bbox("all"))
        debug_print(f"Created {len(self.preset_files)} preset buttons.", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _on_preset_button_click(self, preset_name):
        """
        Handles a click on a preset button. Loads the selected preset to the instrument.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        self.console_print_func(f"\nAttempting to load preset: {preset_name}...")
        debug_print(f"Preset button clicked: {preset_name}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if not self.app_instance.inst:
            self.console_print_func("⚠️ Warning: No instrument connected. Cannot load preset.")
            debug_print("No instrument connected for loading preset.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        # Visually indicate selection
        if self.current_selected_button:
            self.current_selected_button.config(style='Markers.TButton') # Revert old button style
        
        clicked_button = self.preset_buttons.get(preset_name)
        if clicked_button:
            clicked_button.config(style='SelectedSpan.TButton') # Apply selected style
            self.current_selected_button = clicked_button
            self.selected_preset = preset_name
        
        # Start loading in a new thread to keep GUI responsive
        load_thread = threading.Thread(target=self._load_preset_threaded, args=(preset_name,))
        load_thread.daemon = True
        load_thread.start()


    def _load_preset_threaded(self, preset_name):
        """
        Loads the selected preset to the instrument in a separate thread.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Loading preset '{preset_name}' in thread.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        
        success, center_freq, span, rbw = load_selected_preset_logic(self.app_instance, preset_name, self.console_print_func) # Pass console_print_func

        if success:
            self.app_instance.after(0, lambda: self.console_print_func(f"✅ Preset '{preset_name}' loaded successfully."))
            # Update main app's instrument settings display
            self.app_instance.after(0, lambda: self.app_instance.current_center_freq_var.set(f"{center_freq:.3f} MHz" if center_freq else "N/A"))
            self.app_instance.after(0, lambda: self.app_instance.current_span_var.set(f"{span:.3f} MHz" if span else "N/A"))
            self.app_instance.after(0, lambda: self.app_instance.current_rbw_var.set(f"{rbw / 1000:.1f} kHz" if rbw else "N/A"))
        else:
            self.app_instance.after(0, lambda: self.console_print_func(f"❌ Error: Failed to load preset '{preset_name}'."))
        debug_print(f"Preset loading thread for '{preset_name}' finished. Success: {success}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _open_preset_folder(self):
        """
        Opens the local preset folder in the system's file explorer.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Attempting to open preset folder...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        preset_folder = os.path.join(os.getcwd(), "presets") # Assuming 'presets' folder in app directory
        
        # Ensure the directory exists before trying to open it
        if not os.path.exists(preset_folder):
            try:
                os.makedirs(preset_folder, exist_ok=True)
                self.console_print_func(f"Created preset folder: {preset_folder}")
                debug_print(f"Created preset folder: {preset_folder}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            except Exception as e:
                self.console_print_func(f"❌ Failed to create preset folder: {e}")
                debug_print(f"Failed to create preset folder: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
                return

        try:
            # Open the folder using the appropriate command for the OS
            if sys.platform == "win32":
                subprocess.Popen(['explorer', preset_folder]) # Use explorer for Windows
            elif sys.platform == "darwin": # macOS
                subprocess.Popen(["open", preset_folder])
            else: # Linux
                subprocess.Popen(["xdg-open", preset_folder])
            self.console_print_func(f"Opened preset folder: {preset_folder}")
            debug_print(f"Opened preset folder: {preset_folder}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        except Exception as e:
            self.console_print_func(f"❌ Could not open preset folder: {e}")
            debug_print(f"Error opening preset folder '{preset_folder}': {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)

