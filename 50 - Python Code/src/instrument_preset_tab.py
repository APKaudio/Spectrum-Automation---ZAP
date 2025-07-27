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
    def __init__(self, master=None, app_instance=None, **kwargs):
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self.preset_files = [] # To store list of .sta files (either local or device)
        self.selected_preset = None # To store the currently selected preset name
        self.source_of_displayed_presets = "local" # "local" or "device"
        self.preset_buttons = {} # Dictionary to store references to preset buttons for dynamic updates
        self.current_selected_button = None # To keep track of the currently selected button widget

        self._create_widgets()
        # Initial population will happen in _on_tab_selected,
        # which is called when the tab is first displayed.

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Preset Files tab.
        """
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0) # For buttons
        self.grid_rowconfigure(1, weight=1) # For preset display

        # Control buttons for presets
        control_frame = ttk.Frame(self)
        control_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        control_frame.grid_columnconfigure(0, weight=1)
        control_frame.grid_columnconfigure(1, weight=1)

        ttk.Button(control_frame, text="Load Local Presets", command=self._populate_local_preset_list, style='GreyText.TButton').grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        
        # Store a reference to query_presets_button for external control
        self.query_presets_button = ttk.Button(control_frame, text="Query Presets from Device", command=self._query_presets_from_device, state=tk.DISABLED, style='GreyText.TButton')
        self.query_presets_button.grid(row=0, column=1, padx=2, pady=2, sticky="ew")

        # Frame to hold the preset buttons (scrollable)
        preset_display_frame = ttk.Frame(self, style='Dark.TFrame')
        preset_display_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        preset_display_frame.grid_columnconfigure(0, weight=1)
        preset_display_frame.grid_rowconfigure(0, weight=1)

        # Use a canvas with a scrollbar for the preset buttons
        self.preset_canvas = tk.Canvas(preset_display_frame, borderwidth=0, highlightthickness=0, bg='#1e1e1e')
        self.preset_canvas.grid(row=0, column=0, sticky="nsew")

        preset_scrollbar = ttk.Scrollbar(preset_display_frame, orient="vertical", command=self.preset_canvas.yview)
        preset_scrollbar.grid(row=0, column=1, sticky="ns")
        self.preset_canvas.config(yscrollcommand=preset_scrollbar.set)

        self.inner_preset_frame = ttk.Frame(self.preset_canvas, style='Dark.TFrame')
        self.preset_canvas.create_window((0, 0), window=self.inner_preset_frame, anchor="nw")

        self.inner_preset_frame.bind("<Configure>", lambda e: self.preset_canvas.config(scrollregion=self.preset_canvas.bbox("all")))
        self.preset_canvas.bind('<Enter>', self._bind_preset_mouse_wheel)
        self.preset_canvas.bind('<Leave>', self._unbind_preset_mouse_wheel)

    def _bind_preset_mouse_wheel(self, event):
        """Binds mouse wheel events for the preset canvas."""
        event.widget.bind_all("<MouseWheel>", self._on_preset_mouse_wheel)
        event.widget.bind_all("<Button-4>", self._on_preset_mouse_wheel) # For Linux
        event.widget.bind_all("<Button-5>", self._on_preset_mouse_wheel) # For Linux

    def _unbind_preset_mouse_wheel(self, event):
        """Unbinds mouse wheel events for the preset canvas."""
        event.widget.unbind_all("<MouseWheel>")
        event.widget.unbind_all("<Button-4>")
        event.widget.unbind_all("<Button-5>")

    def _on_preset_mouse_wheel(self, event):
        """Handles mouse wheel scrolling for the preset canvas."""
        if sys.platform == "darwin":
            self.preset_canvas.yview_scroll(-1 * int(event.delta), "units")
        elif event.num == 4: # Linux scroll up
            self.preset_canvas.yview_scroll(-1, "units")
        elif event.num == 5: # Linux scroll down
            self.preset_canvas.yview_scroll(1, "units")
        else: # Windows
            self.preset_canvas.yview_scroll(-1 * int(event.delta/120), "units")

    def _populate_local_preset_list(self):
        """
        Scans the local C:\PRESETS directory for .sta files and populates the display.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__

        preset_folder = "C:\\PRESETS" # Hardcoded local preset folder
        if not os.path.isdir(preset_folder):
            print(f"🚫 Local preset folder not found: {preset_folder}")
            debug_print(f"Local preset folder not found: {preset_folder}", file=current_file, function=current_function)
            self.populate_preset_buttons([], source="local") # Clear display
            return

        local_presets = []
        try:
            for filename in os.listdir(preset_folder):
                if filename.lower().endswith(".sta"):
                    local_presets.append(filename)
            print(f"✅ Found {len(local_presets)} local presets in {preset_folder}")
            debug_print(f"Found {len(local_presets)} local presets.", file=current_file, function=current_function)
            self.populate_preset_buttons(local_presets, source="local")
        except Exception as e:
            print(f"❌ Error listing local presets: {e}")
            debug_print(f"Error listing local presets: {e}", file=current_file, function=current_function)
            self.populate_preset_buttons([], source="local") # Clear display on error

    def populate_preset_buttons(self, presets_list, source):
        """
        Populates the GUI with buttons for each preset.
        Clears existing buttons first.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__

        self.clear_preset_buttons()
        self.source_of_displayed_presets = source
        self.preset_files = presets_list
        self.preset_buttons = {} # Reset the dictionary

        if not presets_list:
            ttk.Label(self.inner_preset_frame, text=f"No {source} presets found.", foreground="#cccccc", background="#1e1e1e").grid(row=0, column=0, padx=5, pady=5, sticky="w")
            debug_print(f"No {source} presets to display.", file=current_file, function=current_function)
            return

        # Sort presets alphabetically
        sorted_presets = sorted(presets_list, key=lambda x: x.lower())

        row_idx = 0
        col_idx = 0
        # Display presets in a 3-column grid
        for i, preset_name in enumerate(sorted_presets):
            # Special handling for "MON" presets
            if preset_name.upper().startswith("MON"):
                button_text = f"MON\n{preset_name.replace('.sta', '')}"
                button_style = "SelectedPreset.TButton" if preset_name == self.selected_preset else "LargePreset.TButton"
            else:
                button_text = preset_name.replace('.sta', '')
                button_style = "SelectedPreset.TButton" if preset_name == self.selected_preset else "LargePreset.TButton"
            
            # Create a frame for each button to hold button and info label
            button_frame = ttk.Frame(self.inner_preset_frame, style='Dark.TFrame')
            button_frame.grid(row=row_idx, column=col_idx, padx=5, pady=5, sticky="nsew")
            button_frame.grid_columnconfigure(0, weight=1)

            btn = ttk.Button(button_frame, text=button_text,
                             command=lambda p=preset_name: self._on_preset_button_click(p),
                             style=button_style)
            btn.grid(row=0, column=0, sticky="ew")
            self.preset_buttons[preset_name] = {"button": btn, "info_label": None} # Store button reference

            # Add an info label below the button, initially empty
            info_label = ttk.Label(button_frame, text="", background="#1e1e1e", foreground="#cccccc", font=("Helvetica", 10))
            info_label.grid(row=1, column=0, sticky="ew")
            self.preset_buttons[preset_name]["info_label"] = info_label

            col_idx += 1
            if col_idx >= 3: # 3 columns per row
                col_idx = 0
                row_idx += 1
        debug_print(f"Populated {len(presets_list)} {source} preset buttons.", file=current_file, function=current_function)

    def _on_preset_button_click(self, preset_name):
        """
        Handles a click on a preset button, loads the preset, and updates button styles.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__

        if self.current_selected_button:
            # Reset the style of the previously selected button
            prev_preset_name = None
            for name, data in self.preset_buttons.items():
                if data["button"] == self.current_selected_button:
                    prev_preset_name = name
                    break
            if prev_preset_name:
                # Determine original style based on MON prefix
                if prev_preset_name.upper().startswith("MON"):
                    self.current_selected_button.config(style="LargePreset.TButton")
                else:
                    self.current_selected_button.config(style="LargePreset.TButton")
            self.current_selected_button = None

        # Set the new selected preset and update its style
        self.selected_preset = preset_name
        if preset_name in self.preset_buttons:
            self.current_selected_button = self.preset_buttons[preset_name]["button"]
            if preset_name.upper().startswith("MON"):
                self.current_selected_button.config(style="SelectedPreset.TButton")
            else:
                self.current_selected_button.config(style="SelectedPreset.TButton")
        
        print(f"Selected preset: {preset_name}")
        debug_print(f"Selected preset: {preset_name}", file=current_file, function=current_function)

        # Enable load preset button in main_app if an instrument is connected
        if self.app_instance.inst:
            self.app_instance.load_preset_button.config(state=tk.NORMAL)
        else:
            print("🚫 No instrument connected to load preset.")
            debug_print("No instrument connected to load preset.", file=current_file, function=current_function)
            self.app_instance.load_preset_button.config(state=tk.DISABLED)

    def update_preset_button_info(self, preset_name, center_freq_hz, span_hz, rbw_hz):
        """
        Updates the info label of a specific preset button with its current settings.
        """
        if preset_name in self.preset_buttons and self.preset_buttons[preset_name]["info_label"]:
            info_label = self.preset_buttons[preset_name]["info_label"]
            if center_freq_hz is not None and span_hz is not None and rbw_hz is not None:
                info_text = (f"Center: {center_freq_hz / MHZ_TO_HZ:.3f} MHz\n"
                             f"Span: {span_hz / MHZ_TO_HZ:.3f} MHz\n"
                             f"RBW: {rbw_hz:.0f} Hz")
                info_label.config(text=info_text)
            else:
                info_label.config(text="Info: N/A")

    def get_selected_preset(self):
        """Returns the name of the currently selected preset."""
        return self.selected_preset

    def _query_presets_from_device(self):
        """
        Calls the logic to query presets directly from the connected instrument.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__

        if not self.app_instance.inst:
            print("🚫 Not connected to instrument. Cannot query presets.")
            debug_print("Not connected to instrument, cannot query presets.", file=current_file, function=current_function)
            # tk.messagebox.showwarning("Not Connected", "Please connect to an instrument first to query presets.") # Removed messagebox
            return
        
        # Call the logic from instrument_logic.py
        query_device_presets_logic(self.app_instance)

    def clear_preset_buttons(self):
        """
        Removes all dynamically created preset buttons from the display.
        """
        for widget in self.inner_preset_frame.winfo_children():
            widget.destroy()
        self.preset_buttons = {}
        self.selected_preset = None
        self.current_selected_button = None
        debug_print("Cleared all preset buttons.", file=__file__, function=inspect.currentframe().f_code.co_name)

    def _on_tab_selected(self, event):
        """
        Callback for when this tab is selected in the main notebook.
        Refreshes the local preset list and updates the query button state.
        """
        debug_print("Instrument Presets Tab selected.", file=__file__, function=inspect.currentframe().f_code.co_name)
        self._populate_local_preset_list() # Always refresh local list when tab is selected

        # Update the state of the query presets button
        if self.app_instance.inst:
            self.query_presets_button.config(state=tk.NORMAL)
        else:
            self.query_presets_button.config(state=tk.DISABLED)

    def _open_preset_folder(self):
        """
        Opens the local preset folder in the system's file explorer.
        """
        current_file = inspect.currentframe().f_code.co_name
        current_function = inspect.currentframe().f_code.co_name

        preset_folder = "C:\\PRESETS" # Hardcoded local preset folder
        
        # Ensure the directory exists before trying to open it
        if not os.path.exists(preset_folder):
            try:
                os.makedirs(preset_folder, exist_ok=True)
                print(f"Created preset folder: {preset_folder}")
                debug_print(f"Created preset folder: {preset_folder}", file=current_file, function=current_function)
            except Exception as e:
                print(f"❌ Failed to create preset folder: {e}") # Changed messagebox to print
                debug_print(f"Failed to create preset folder: {e}", file=current_file, function=current_function)
                return

        try:
            # Open the folder using the appropriate command for the OS
            if sys.platform == "win32":
                subprocess.Popen(['explorer', preset_folder]) # Use explorer for Windows
            elif sys.platform == "darwin": # macOS
                subprocess.Popen(["open", preset_folder])
            else: # Linux
                subprocess.Popen(["xdg-open", preset_folder])
            print(f"Opened preset folder: {preset_folder}")
            debug_print(f"Opened preset folder: {preset_folder}", file=current_file, function=current_function)
        except Exception as e:
            print(f"❌ Could not open preset folder: {e}") # Changed messagebox to print
            debug_print(f"Error opening preset folder '{preset_folder}': {e}", file=current_file, function=current_function)

