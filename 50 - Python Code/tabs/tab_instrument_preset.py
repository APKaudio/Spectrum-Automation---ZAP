# src/instrument_preset_tab.py
import tkinter as tk
from tkinter import scrolledtext, filedialog, ttk # Keep other imports
import os
import sys
import inspect
import subprocess # Add this import for opening folders

# Import instrument_logic for setting focus frequency and loading presets
from src.instrument_logic import (
    load_selected_preset_logic,
    # query_device_presets_logic, # Removed as it's now in preset_utils.py
    query_current_instrument_settings_logic # Import the logic function to get settings
)
# Import query_device_presets_logic from preset_utils.py
from utils.preset_utils import query_device_presets_logic, load_selected_preset as load_selected_preset_util # Import the logic function from its new home

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
        self.current_selected_button = None # Track the currently selected button for visual feedback

        self._create_widgets()
        # Bind the tab selection event
        if master: # Only bind if master (notebook) exists
            master.bind("<<NotebookTabChanged>>", self._on_tab_selected)

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Preset Files tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Creating PresetFilesTab widgets...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # Configure grid for two main columns (MON and Other)
        self.grid_columnconfigure(0, weight=1) # MON column
        self.grid_columnconfigure(1, weight=0) # Scrollbar for MON
        self.grid_columnconfigure(2, weight=1) # Other column
        self.grid_columnconfigure(3, weight=0) # Scrollbar for Other
        self.grid_rowconfigure(0, weight=1) # Main row for canvases
        self.grid_rowconfigure(1, weight=0) # For the Query button (and removed Open Folder button)

        # Frame for MON presets group label
        self.mon_label_frame = ttk.Frame(self, style='Dark.TFrame')
        self.mon_label_frame.grid(row=0, column=0, sticky="new", padx=10, pady=5)
        ttk.Label(self.mon_label_frame, text="MON Presets",
                  background="#333333", foreground="#F4902C", font=("Helvetica", 16, "bold")).pack(padx=5, pady=5)

        # Canvas for scrollable buttons for "MON" presets
        self.mon_preset_canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, bg='#333333')
        self.mon_preset_canvas.grid(row=0, column=0, sticky="nsew", padx=10, pady=(50, 5)) # Adjusted pady to make space for label

        self.mon_preset_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.mon_preset_canvas.yview)
        self.mon_preset_scrollbar.grid(row=0, column=1, sticky="ns")
        self.mon_preset_canvas.config(yscrollcommand=self.mon_preset_scrollbar.set)

        self.inner_mon_buttons_frame = ttk.Frame(self.mon_preset_canvas, style='Dark.TFrame')
        self.mon_preset_canvas.create_window((0, 0), window=self.inner_mon_buttons_frame, anchor="nw")

        # Configure inner frame for two columns
        self.inner_mon_buttons_frame.grid_columnconfigure(0, weight=1)
        self.inner_mon_buttons_frame.grid_columnconfigure(1, weight=1)

        self.inner_mon_buttons_frame.bind("<Configure>", lambda e: self.mon_preset_canvas.config(scrollregion=self.mon_preset_canvas.bbox("all")))
        self.mon_preset_canvas.bind('<Enter>', self._bind_mouse_wheel)
        self.mon_preset_canvas.bind('<Leave>', self._unbind_mouse_wheel)


        # Frame for Other presets group label
        self.other_label_frame = ttk.Frame(self, style='Dark.TFrame')
        self.other_label_frame.grid(row=0, column=2, sticky="new", padx=10, pady=5)
        ttk.Label(self.other_label_frame, text="Other Presets",
                  background="#333333", foreground="#F4902C", font=("Helvetica", 16, "bold")).pack(padx=5, pady=5)

        # Canvas for scrollable buttons for "Other" presets
        self.other_preset_canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, bg='#333333')
        self.other_preset_canvas.grid(row=0, column=2, sticky="nsew", padx=10, pady=(50, 5)) # Adjusted pady to make space for label

        self.other_preset_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.other_preset_canvas.yview)
        self.other_preset_scrollbar.grid(row=0, column=3, sticky="ns")
        self.other_preset_canvas.config(yscrollcommand=self.other_preset_scrollbar.set)

        self.inner_other_buttons_frame = ttk.Frame(self.other_preset_canvas, style='Dark.TFrame')
        self.other_preset_canvas.create_window((0, 0), window=self.inner_other_buttons_frame, anchor="nw")

        # Configure inner frame for two columns
        self.inner_other_buttons_frame.grid_columnconfigure(0, weight=1)
        self.inner_other_buttons_frame.grid_columnconfigure(1, weight=1)

        self.inner_other_buttons_frame.bind("<Configure>", lambda e: self.other_preset_canvas.config(scrollregion=self.inner_other_buttons_frame.bbox("all")))
        self.other_preset_canvas.bind('<Enter>', self._bind_mouse_wheel)
        self.other_preset_canvas.bind('<Leave>', self._unbind_mouse_wheel)

        # Query Presets Button
        self.query_presets_button = ttk.Button(self, text="Query Presets from Device", command=self._query_presets_from_device, state=tk.DISABLED, style='Blue.TButton')
        self.query_presets_button.grid(row=1, column=0, columnspan=4, padx=10, pady=5, sticky="ew") # Spanning across all columns

        # Removed the "Open Local Preset Folder" button
        # self.open_preset_folder_button = ttk.Button(
        #     self,
        #     text="Open Local Preset Folder",
        #     command=self._open_local_preset_folder,
        #     style='Purple.TButton'
        # )
        # self.open_preset_folder_button.grid(row=2, column=0, columnspan=4, padx=10, pady=5, sticky="ew")


        debug_print("PresetFilesTab widgets created.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _bind_mouse_wheel(self, event):
        """Binds mouse wheel events for the canvas."""
        event.widget.bind_all("<MouseWheel>", self._on_mouse_wheel)
        event.widget.bind_all("<Button-4>", self._on_mouse_wheel) # For Linux
        event.widget.bind_all("<Button-5>", self._on_mouse_wheel) # For Linux

    def _unbind_mouse_wheel(self, event):
        """Unbinds mouse wheel events for the canvas."""
        event.widget.unbind_all("<MouseWheel>")
        event.widget.unbind_all("<Button-4>")
        event.widget.unbind_all("<Button-5>")

    def _on_mouse_wheel(self, event):
        """Handles mouse wheel scrolling for the canvas."""
        if sys.platform == "darwin":
            event.widget.yview_scroll(-1 * int(event.delta), "units")
        elif event.num == 4: # Linux scroll up
            event.widget.yview_scroll(-1, "units")
        elif event.num == 5: # Linux scroll down
            event.widget.yview_scroll(1, "units")
        else: # Windows
            event.widget.yview_scroll(-1 * int(event.delta/120), "units")

    def _query_presets_from_device(self):
        """
        Queries the connected instrument for available preset files and updates the display.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        self.console_print_func("Querying device presets...")
        debug_print("Calling query_device_presets_logic...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        
        # Add debug print for app_instance.inst before calling the logic function
        self.console_print_func(f"Instrument instance in _query_presets_from_device: {self.app_instance.inst}")
        debug_print(f"Instrument instance in _query_presets_from_device: {self.app_instance.inst}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # Call the logic function from preset_utils
        presets = query_device_presets_logic(self.app_instance, self.console_print_func)
        
        if presets:
            self.preset_files = presets
            self.source_of_displayed_presets = "device"
            self.populate_preset_buttons(presets, source="device")
            # The main app's load preset button is managed by _on_preset_button_click
        else:
            self.preset_files = []
            self.source_of_displayed_presets = "device"
            self.populate_preset_buttons([], source="device") # Clear display
            # self.app_instance.load_preset_button.config(state=tk.DISABLED) # Removed, no longer exists in main_app

    def _load_selected_preset(self):
        """
        Loads the currently selected preset onto the instrument.
        This method is called by the main app's "Load Selected Preset" button.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        if self.selected_preset:
            self.console_print_func(f"Loading preset: {self.selected_preset}")
            debug_print(f"Calling load_selected_preset_logic for: {self.selected_preset}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            
            # Call the load utility function from preset_utils
            success, center_freq, span, rbw = load_selected_preset_util(self.app_instance.inst, self.selected_preset, self.console_print_func)
            
            if success:
                self.console_print_func(f"✅ Preset '{self.selected_preset}' loaded. Instrument settings updated.")
                # Update the button text with the new settings
                # Note: center_freq and span are already in MHz from load_selected_preset_logic
                self.update_preset_button_info(self.selected_preset, center_freq, span, rbw)
            else:
                self.console_print_func(f"❌ Failed to load preset '{self.selected_preset}'.")
                # If loading fails, revert the button style and clear selection
                if self.current_selected_button:
                    self.current_selected_button.config(style='LargePreset.TButton')
                    self.current_selected_button = None
                self.selected_preset = None

        else:
            self.console_print_func("⚠️ Warning: No preset selected to load.")
            debug_print("No preset selected to load.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            # If not connected, clear selected button and revert its style if any
            if self.current_selected_button:
                self.current_selected_button.config(style='LargePreset.TButton')
                self.current_selected_button = None
            self.selected_preset = None


    def _populate_local_preset_list(self):
        """
        Populates the list of .sta preset files from the C:\\PRESETS directory
        (on the local machine) and creates buttons for each.
        This is typically for initial display or if the user wants to see local files.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Populating local preset list from C:\\PRESETS\\...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        preset_dir = "C:\\PRESETS" # Hardcoded local path
        local_presets = []

        self.clear_preset_buttons() # Clear existing buttons and reset state

        if not os.path.exists(preset_dir):
            debug_print(f"Local preset directory not found: {preset_dir}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            self.console_print_func(f"ℹ️ Info: Local preset directory not found: {preset_dir}")
            self.source_of_displayed_presets = "local" # Still set source even if folder not found
            self.preset_files = []
            self.populate_preset_buttons([], source="local") # Call to show "No presets found"
            return

        try:
            for filename in os.listdir(preset_dir):
                if filename.lower().endswith(".sta"):
                    local_presets.append(filename) # Store full name with .STA extension
            local_presets.sort() # Sort alphabetically
            debug_print(f"Found {len(local_presets)} local preset files.", file=current_file, function=current_function, console_print_func=self.console_print_func)

        except Exception as e:
            debug_print(f"Error listing local preset files: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            self.console_print_func(f"❌ Error listing local preset files: {e}")
            self.source_of_displayed_presets = "local"
            self.preset_files = []
            self.populate_preset_buttons([], source="local") # Call to show "No presets found"
            return

        if not local_presets:
            self.source_of_displayed_presets = "local"
            self.preset_files = []
            self.populate_preset_buttons([], source="local") # Call to show "No presets found"
            return
        
        # Call the generic populate_preset_buttons with the local files
        self.populate_preset_buttons(local_presets, source="local")
        debug_print("Local preset buttons populated.", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def populate_preset_buttons(self, presets_list, source="unknown"):
        """
        Populates the display with clickable buttons for each preset in the given list,
        categorizing them into "MON" and "Other" groups, arranged in two columns.

        Inputs:
            presets_list (list): A list of preset names (strings) to display as buttons.
            source (str): The source of these presets ("local" or "device").
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Populating preset buttons with {len(presets_list)} items from {source}...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        self.clear_preset_buttons() # Clear existing buttons and reset state

        self.preset_files = presets_list # Update the stored list
        self.source_of_displayed_presets = source # Update the source

        mon_presets = sorted([p for p in presets_list if "MON" in p.upper()])
        other_presets = sorted([p for p in presets_list if "MON" not in p.upper()])

        # Populate MON presets
        if mon_presets:
            for i, preset_name in enumerate(mon_presets):
                row = i // 2 # Changed to 2 columns
                col = i % 2 # Changed to 2 columns
                button = ttk.Button(self.inner_mon_buttons_frame, text=preset_name,
                                    command=lambda name=preset_name: self._on_preset_button_click(name),
                                    style='LargePreset.TButton')
                button.grid(row=row, column=col, padx=5, pady=5, sticky="ew") # Use grid for 2 columns
                self.preset_buttons[preset_name] = button # Store button reference
        else:
            ttk.Label(self.inner_mon_buttons_frame, text="No MON presets found.",
                      background="#333333", foreground="white").grid(row=0, column=0, columnspan=2, padx=10, pady=10) # Changed columnspan to 2


        # Populate Other presets
        if other_presets:
            for i, preset_name in enumerate(other_presets):
                row = i // 2 # Changed to 2 columns
                col = i % 2 # Changed to 2 columns
                button = ttk.Button(self.inner_other_buttons_frame, text=preset_name,
                                    command=lambda name=preset_name: self._on_preset_button_click(name),
                                    style='LargePreset.TButton')
                button.grid(row=row, column=col, padx=5, pady=5, sticky="ew") # Use grid for 2 columns
                self.preset_buttons[preset_name] = button # Store button reference
        else:
            ttk.Label(self.inner_other_buttons_frame, text="No other presets found.",
                      background="#333333", foreground="white").grid(row=0, column=0, columnspan=2, padx=10, pady=10) # Changed columnspan to 2


        self.inner_mon_buttons_frame.update_idletasks()
        self.mon_preset_canvas.config(scrollregion=self.inner_mon_buttons_frame.bbox("all"))
        self.inner_other_buttons_frame.update_idletasks()
        self.other_preset_canvas.config(scrollregion=self.inner_other_buttons_frame.bbox("all"))
        debug_print("Preset buttons populated and scroll regions updated.", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _on_preset_button_click(self, preset_name):
        """
        Callback for individual preset buttons. Sets the selected preset and
        attempts to load it onto the instrument.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Preset button clicked: {preset_name}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        
        # If there was a previously selected button, revert its style
        if self.current_selected_button:
            self.current_selected_button.config(style='LargePreset.TButton') # Revert to default style

        self.selected_preset = preset_name
        
        if self.app_instance and self.app_instance.inst:
            # Set the style of the newly selected button to 'SelectedPreset.TButton'
            if preset_name in self.preset_buttons:
                self.current_selected_button = self.preset_buttons[preset_name]
                self.current_selected_button.config(style='SelectedPreset.TButton') # Set to selected style
            
            # Call the load utility function from preset_utils
            success, center_freq, span, rbw = load_selected_preset_util(self.app_instance.inst, self.selected_preset, self.console_print_func)
            
            if success:
                self.console_print_func(f"✅ Preset '{self.selected_preset}' loaded. Instrument settings updated.")
                # Update the button text with the new settings
                # Note: center_freq and span are already in MHz from load_selected_preset_logic
                self.update_preset_button_info(self.selected_preset, center_freq, span, rbw)
            else:
                self.console_print_func(f"❌ Failed to load preset '{self.selected_preset}'.")
                # If loading fails, revert the button style and clear selection
                if self.current_selected_button:
                    self.current_selected_button.config(style='LargePreset.TButton')
                    self.current_selected_button = None
                self.selected_preset = None

        else:
            self.console_print_func("⚠️ Warning: Please connect to an instrument first to load a preset.")
            debug_print("Not connected to instrument, cannot load preset.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            # If not connected, clear selected button and revert its style if any
            if self.current_selected_button:
                self.current_selected_button.config(style='LargePreset.TButton')
                self.current_selected_button = None
            self.selected_preset = None


    def update_preset_button_info(self, preset_name, center_freq_mhz, span_mhz, rbw_hz):
        """
        Updates the label of a specific preset button with frequency, span, and RBW info.
        Inputs:
            preset_name (str): The name of the preset.
            center_freq_mhz (float): Center frequency in MHz.
            span_mhz (float): Span in MHz.
            rbw_hz (float): RBW in Hz.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Updating button '{preset_name}' with C:{center_freq_mhz:.3f}MHz SP:{span_mhz:.3f}MHz RBW:{rbw_hz}Hz", file=current_file, function=current_function, console_print_func=self.console_print_func)
        if preset_name in self.preset_buttons:
            button = self.preset_buttons[preset_name]
            current_text = preset_name # Start with original preset name
            
            # Format frequency, span, and RBW. No division by MHZ_TO_HZ needed here.
            formatted_freq = f"C: {center_freq_mhz:.3f} MHz" if center_freq_mhz is not None else "N/A"
            formatted_span = f"SP: {span_mhz:.3f} MHz" if span_mhz is not None else "N/A"
            formatted_rbw = f"RBW: {rbw_hz / 1000:.1f} kHz" if rbw_hz is not None else "N/A" # This one is correct (Hz to kHz)

            # Combine original text with new info, using newlines
            new_text = f"{current_text}\n{formatted_freq}\n{formatted_span}\n{formatted_rbw}" # Add RBW to new line
            
            button.config(text=new_text)
            debug_print(f"Button '{preset_name}' text updated to:\n{new_text}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            debug_print(f"Button for preset '{preset_name}' not found for update.", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def get_selected_preset(self):
        """
        Returns the name of the currently selected preset.
        """
        return self.selected_preset

    def clear_preset_buttons(self):
        """
        Clears all preset buttons from the display.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Clearing preset buttons...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        for widget in self.inner_mon_buttons_frame.winfo_children():
            widget.destroy()
        for widget in self.inner_other_buttons_frame.winfo_children():
            widget.destroy()

        self.preset_buttons.clear() # Clear the stored button references
        self.current_selected_button = None # Clear the currently selected button

        self.inner_mon_buttons_frame.update_idletasks()
        self.mon_preset_canvas.config(scrollregion=self.inner_mon_buttons_frame.bbox("all"))
        self.inner_other_buttons_frame.update_idletasks()
        self.other_preset_canvas.config(scrollregion=self.inner_other_buttons_frame.bbox("all"))

        self.selected_preset = None # Clear selected preset
        self.preset_files = [] # Clear the stored list of presets
        self.source_of_displayed_presets = "unknown" # Reset source
        # self.app_instance.load_preset_button.config(state=tk.DISABLED) # Removed, no longer exists in main_app
        
        # Add placeholder labels if no presets are displayed
        ttk.Label(self.inner_mon_buttons_frame, text="No MON presets found.",
                      background="#333333", foreground="white").grid(row=0, column=0, columnspan=2, padx=10, pady=10)
        ttk.Label(self.inner_other_buttons_frame, text="No other presets found.",
                      background="#333333", foreground="white").grid(row=0, column=0, columnspan=2, padx=10, pady=10)


    def _on_tab_selected(self, event):
        """
        Callback when this tab is selected. Updates the state of the query presets button.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("PresetFilesTab selected.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        if self.app_instance and self.app_instance.inst:
            self.query_presets_button.config(state=tk.NORMAL)
            # Auto-query presets from device if N9342CN and not already displayed from device
            if self.app_instance.instrument_model == "N9342CN" and self.source_of_displayed_presets != "device":
                debug_print("Auto-querying device presets for N9342CN.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                self._query_presets_from_device()
            elif self.source_of_displayed_presets == "device" and self.preset_files:
                # If device presets were last displayed and we have them, re-display them
                debug_print("Re-populating with previously queried device presets.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                self.populate_preset_buttons(self.preset_files, source="device")
            else:
                # Otherwise, default to local presets or if no device presets were stored
                debug_print("Populating with local presets or no presets stored.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                self._populate_local_preset_list() 
        else:
            self.query_presets_button.config(state=tk.DISABLED)
            # Clear buttons if no instrument is connected
            self.clear_preset_buttons()

    def _open_local_preset_folder(self):
        """
        Opens the local preset folder (e.g., 'presets' directory in the app's root).
        Creates the folder if it doesn't exist.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        
        # Define the path to the local presets folder
        # Assuming it's a 'presets' directory relative to the script's location
        script_dir = os.path.dirname(os.path.abspath(__file__))
        # Go up one level from 'src' to the root, then into 'presets'
        preset_folder = os.path.join(os.path.dirname(script_dir), "presets")

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

