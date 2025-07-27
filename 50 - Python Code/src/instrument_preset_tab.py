# src/instrument_preset_tab.py
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, ttk
import os
import sys
import inspect

# Import instrument_logic for setting focus frequency and loading presets
from src.instrument_logic import (
    load_selected_preset_logic,
    query_device_presets_logic
)
from utils.instrument_control import debug_print

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

        self._create_widgets()
        self._populate_local_preset_list() # Populate from local directory on initialization

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Preset Files tab.
        """
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0) # For the query button

        # Canvas for scrollable buttons
        self.preset_canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, bg='#333333')
        self.preset_canvas.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        self.preset_scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.preset_canvas.yview)
        self.preset_scrollbar.grid(row=0, column=1, sticky="ns")
        self.preset_canvas.config(yscrollcommand=self.preset_scrollbar.set)

        self.inner_buttons_frame = ttk.Frame(self.preset_canvas, style='Dark.TFrame')
        self.preset_canvas.create_window((0, 0), window=self.inner_buttons_frame, anchor="nw")

        self.inner_buttons_frame.bind("<Configure>", lambda e: self.preset_canvas.config(scrollregion=self.preset_canvas.bbox("all")))
        self.preset_canvas.bind('<Enter>', self._bind_mouse_wheel)
        self.preset_canvas.bind('<Leave>', self._unbind_mouse_wheel)

        # Query Presets Button
        self.query_presets_button = ttk.Button(self, text="Query Presets from Device", command=self._query_presets_from_device, state=tk.DISABLED, style='Accent.TButton')
        self.query_presets_button.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="ew")


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
            self.preset_canvas.yview_scroll(-1 * int(event.delta), "units")
        elif event.num == 4: # Linux scroll up
            self.preset_canvas.yview_scroll(-1, "units")
        elif event.num == 5: # Linux scroll down
            self.preset_canvas.yview_scroll(1, "units")
        else: # Windows
            self.preset_canvas.yview_scroll(-1 * int(event.delta/120), "units")

    def _populate_local_preset_list(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Populates the list of .sta preset files from the C:\\PRESETS directory
        (on the local machine) and creates buttons for each.
        This is typically for initial display or if the user wants to see local files.
        """
        debug_print("Populating local preset list from C:\\PRESETS\\...", file=file, function=function)
        preset_dir = "C:\\PRESETS" # Hardcoded local path
        local_presets = []

        # Clear existing buttons
        for widget in self.inner_buttons_frame.winfo_children():
            widget.destroy()

        if not os.path.exists(preset_dir):
            debug_print(f"Local preset directory not found: {preset_dir}", file=file, function=function)
            ttk.Label(self.inner_buttons_frame, text=f"Local preset directory not found: {preset_dir}",
                      background="#333333", foreground="white").pack(padx=10, pady=10)
            self.source_of_displayed_presets = "local" # Still set source even if folder not found
            self.preset_files = []
            return

        try:
            for filename in os.listdir(preset_dir):
                if filename.lower().endswith(".sta"):
                    local_presets.append(filename) # Store full name with .STA extension
            local_presets.sort() # Sort alphabetically
            debug_print(f"Found {len(local_presets)} local preset files.", file=file, function=function)

        except Exception as e:
            debug_print(f"Error listing local preset files: {e}", file=file, function=function)
            ttk.Label(self.inner_buttons_frame, text=f"Error loading local presets: {e}",
                      background="#333333", foreground="white").pack(padx=10, pady=10)
            self.source_of_displayed_presets = "local"
            self.preset_files = []
            return

        if not local_presets:
            ttk.Label(self.inner_buttons_frame, text="No .sta preset files found in C:\\PRESETS (local).",
                      background="#333333", foreground="white").pack(padx=10, pady=10)
            self.source_of_displayed_presets = "local"
            self.preset_files = []
            return
        
        # Call the generic populate_preset_buttons with the local files
        self.populate_preset_buttons(local_presets, source="local")
        debug_print("Local preset buttons populated.", file=file, function=function)


    def populate_preset_buttons(self, presets_list, source="unknown", file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Populates the display with clickable buttons for each preset in the given list.
        This method is used for both local and device-queried presets.

        Inputs:
            presets_list (list): A list of preset names (strings) to display as buttons.
            source (str): The source of these presets ("local" or "device").
        """
        debug_print(f"Populating preset buttons with {len(presets_list)} items from {source}...", file=file, function=function)
        # Clear existing buttons
        for widget in self.inner_buttons_frame.winfo_children():
            widget.destroy()

        self.preset_files = presets_list # Update the stored list
        self.source_of_displayed_presets = source # Update the source

        if not presets_list:
            ttk.Label(self.inner_buttons_frame, text=f"No presets found from {source}.",
                      background="#333333", foreground="white").pack(padx=10, pady=10)
            self.inner_buttons_frame.update_idletasks()
            self.preset_canvas.config(scrollregion=self.preset_canvas.bbox("all"))
            return

        for i, preset_name in enumerate(sorted(presets_list)): # Always sort for consistent display
            button = ttk.Button(self.inner_buttons_frame, text=preset_name, 
                                command=lambda name=preset_name: self._on_preset_button_click(name),
                                style='LargePreset.TButton')
            button.pack(fill=tk.X, padx=5, pady=2)

        self.inner_buttons_frame.update_idletasks()
        self.preset_canvas.config(scrollregion=self.preset_canvas.bbox("all"))
        debug_print("Preset buttons populated and scroll region updated.", file=file, function=function)


    def _on_preset_button_click(self, preset_name, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Callback for individual preset buttons. Sets the selected preset and
        attempts to load it onto the instrument.
        """
        debug_print(f"Preset button clicked: {preset_name}", file=file, function=function)
        self.selected_preset = preset_name
        
        if self.app_instance and self.app_instance.inst:
            self.app_instance.load_preset_button.config(state=tk.NORMAL) # Enable the main app's load button
            # Directly call the load logic from instrument_logic
            load_selected_preset_logic(self.app_instance, preset_name) # preset_name already includes .STA
        else:
            self.app_instance.load_preset_button.config(state=tk.DISABLED)
            messagebox.showwarning("Not Connected", "Please connect to an instrument first to load a preset.")

    def get_selected_preset(self):
        """
        Returns the name of the currently selected preset.
        """
        return self.selected_preset

    def _query_presets_from_device(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Queries the connected instrument for available presets and updates the display.
        This now calls the logic in instrument_logic.py which will then call
        this tab's populate_preset_buttons with the instrument's presets.
        """
        debug_print("Initiating query for presets from device...", file=file, function=function)
        # The query_device_presets_logic function in instrument_logic.py
        # now handles calling this tab's populate_preset_buttons directly.
        query_device_presets_logic(self.app_instance)
        print("Finished querying presets from device.")

    def clear_preset_buttons(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Clears all preset buttons from the display.
        """
        debug_print("Clearing preset buttons...", file=file, function=function)
        for widget in self.inner_buttons_frame.winfo_children():
            widget.destroy()
        self.inner_buttons_frame.update_idletasks()
        self.preset_canvas.config(scrollregion=self.preset_canvas.bbox("all"))
        self.selected_preset = None # Clear selected preset
        self.preset_files = [] # Clear the stored list of presets
        self.source_of_displayed_presets = "unknown" # Reset source
        if self.app_instance and hasattr(self.app_instance, 'load_preset_button'):
            self.app_instance.load_preset_button.config(state=tk.DISABLED)
        ttk.Label(self.inner_buttons_frame, text="No presets displayed.",
                      background="#333333", foreground="white").pack(padx=10, pady=10)


    def _on_tab_selected(self, event):
        """
        Callback when this tab is selected. Updates the state of the query presets button.
        """
        debug_print("PresetFilesTab selected.", file=__file__, function=inspect.currentframe().f_code.co_name)
        if self.app_instance and self.app_instance.inst:
            self.query_presets_button.config(state=tk.NORMAL)
            if self.source_of_displayed_presets == "device" and self.preset_files:
                # If device presets were last displayed and we have them, re-display them
                debug_print("Re-populating with previously queried device presets.", file=__file__, function=inspect.currentframe().f_code.co_name)
                self.populate_preset_buttons(self.preset_files, source="device")
            else:
                # Otherwise, default to local presets or if no device presets were stored
                debug_print("Populating with local presets or no presets stored.", file=__file__, function=inspect.currentframe().f_code.co_name)
                self._populate_local_preset_list() 
        else:
            self.query_presets_button.config(state=tk.DISABLED)
            # Clear buttons if no instrument is connected
            self.clear_preset_buttons()
