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

        self._create_widgets()
        # Initial population will happen in _on_tab_selected,
        # which is called when the tab is first displayed.

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Preset Files tab.
        """
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(2, weight=1) # Make both columns expandable
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0) # For the query button

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

        self.inner_other_buttons_frame.bind("<Configure>", lambda e: self.other_preset_canvas.config(scrollregion=self.inner_other_buttons_frame.bbox("all")))
        self.other_preset_canvas.bind('<Enter>', self._bind_mouse_wheel)
        self.other_preset_canvas.bind('<Leave>', self._unbind_mouse_wheel)

        # Query Presets Button
        self.query_presets_button = ttk.Button(self, text="Query Presets from Device", command=self._query_presets_from_device, state=tk.DISABLED, style='Accent.TButton')
        self.query_presets_button.grid(row=1, column=0, columnspan=4, padx=10, pady=5, sticky="ew") # Spanning across all columns


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

    def _populate_local_preset_list(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Populates the list of .sta preset files from the C:\\PRESETS directory
        (on the local machine) and creates buttons for each.
        This is typically for initial display or if the user wants to see local files.
        """
        debug_print("Populating local preset list from C:\\PRESETS\\...", file=file, function=function)
        preset_dir = "C:\\PRESETS" # Hardcoded local path
        local_presets = []

        self.clear_preset_buttons() # Clear existing buttons and reset state

        if not os.path.exists(preset_dir):
            debug_print(f"Local preset directory not found: {preset_dir}", file=file, function=function)
            ttk.Label(self.inner_mon_buttons_frame, text=f"Local preset directory not found: {preset_dir}",
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
            ttk.Label(self.inner_mon_buttons_frame, text=f"Error loading local presets: {e}",
                      background="#333333", foreground="white").pack(padx=10, pady=10)
            self.source_of_displayed_presets = "local"
            self.preset_files = []
            return

        if not local_presets:
            ttk.Label(self.inner_mon_buttons_frame, text="No .sta preset files found in C:\\PRESETS (local).",
                      background="#333333", foreground="white").pack(padx=10, pady=10)
            self.source_of_displayed_presets = "local"
            self.preset_files = []
            return
        
        # Call the generic populate_preset_buttons with the local files
        self.populate_preset_buttons(local_presets, source="local")
        debug_print("Local preset buttons populated.", file=file, function=function)


    def populate_preset_buttons(self, presets_list, source="unknown", file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Populates the display with clickable buttons for each preset in the given list,
        categorizing them into "MON" and "Other" groups.

        Inputs:
            presets_list (list): A list of preset names (strings) to display as buttons.
            source (str): The source of these presets ("local" or "device").
        """
        debug_print(f"Populating preset buttons with {len(presets_list)} items from {source}...", file=file, function=function)
        self.clear_preset_buttons() # Clear existing buttons and reset state

        self.preset_files = presets_list # Update the stored list
        self.source_of_displayed_presets = source # Update the source

        mon_presets = sorted([p for p in presets_list if "MON" in p.upper()])
        other_presets = sorted([p for p in presets_list if "MON" not in p.upper()])

        # Populate MON presets
        if mon_presets:
            # The "MON Presets" label is now part of _create_widgets and always present
            for preset_name in mon_presets:
                button = ttk.Button(self.inner_mon_buttons_frame, text=preset_name,
                                    command=lambda name=preset_name: self._on_preset_button_click(name),
                                    style='LargePreset.TButton') # This style should be defined in main_app.py with font size 40
                button.pack(fill=tk.X, padx=5, pady=5) # Increased pady for larger buttons
                self.preset_buttons[preset_name] = button # Store button reference
        else:
            ttk.Label(self.inner_mon_buttons_frame, text="No MON presets found.",
                      background="#333333", foreground="white").pack(padx=10, pady=10)

        # Populate Other presets
        if other_presets:
            # The "Other Presets" label is now part of _create_widgets and always present
            for preset_name in other_presets:
                button = ttk.Button(self.inner_other_buttons_frame, text=preset_name,
                                    command=lambda name=preset_name: self._on_preset_button_click(name),
                                    style='LargePreset.TButton') # This style should be defined in main_app.py with font size 40
                button.pack(fill=tk.X, padx=5, pady=5) # Increased pady for larger buttons
                self.preset_buttons[preset_name] = button # Store button reference
        else:
            ttk.Label(self.inner_other_buttons_frame, text="No other presets found.",
                      background="#333333", foreground="white").pack(padx=10, pady=10)


        self.inner_mon_buttons_frame.update_idletasks()
        self.mon_preset_canvas.config(scrollregion=self.inner_mon_buttons_frame.bbox("all"))
        self.inner_other_buttons_frame.update_idletasks()
        self.other_preset_canvas.config(scrollregion=self.inner_other_buttons_frame.bbox("all"))
        debug_print("Preset buttons populated and scroll regions updated.", file=file, function=function)


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
            # The load_selected_preset_logic will now return the queried settings
            load_selected_preset_logic(self.app_instance, preset_name) # preset_name already includes .STA
        else:
            self.app_instance.load_preset_button.config(state=tk.DISABLED)
            messagebox.showwarning("Not Connected", "Please connect to an instrument first to load a preset.")

    def update_preset_button_info(self, preset_name, center_freq_hz, span_hz, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Updates the label of a specific preset button with frequency and span info.
        """
        debug_print(f"Updating button '{preset_name}' with C:{center_freq_hz/MHZ_TO_HZ:.3f}MHz SP:{span_hz/MHZ_TO_HZ:.3f}MHz", file=file, function=function)
        if preset_name in self.preset_buttons:
            button = self.preset_buttons[preset_name]
            current_text = preset_name # Start with original preset name
            
            # Format frequency and span
            formatted_freq = f"C: {center_freq_hz/MHZ_TO_HZ:.3f} MHz"
            formatted_span = f"SP: {span_hz/MHZ_TO_HZ:.3f} MHz"

            # Combine original text with new info, using newlines
            new_text = f"{current_text}\n{formatted_freq}\n{formatted_span}"
            
            button.config(text=new_text)
            debug_print(f"Button '{preset_name}' text updated to:\n{new_text}", file=file, function=function)
        else:
            debug_print(f"Button for preset '{preset_name}' not found for update.", file=file, function=function)


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
        for widget in self.inner_mon_buttons_frame.winfo_children():
            widget.destroy()
        for widget in self.inner_other_buttons_frame.winfo_children():
            widget.destroy()

        self.preset_buttons.clear() # Clear the stored button references

        self.inner_mon_buttons_frame.update_idletasks()
        self.mon_preset_canvas.config(scrollregion=self.inner_mon_buttons_frame.bbox("all"))
        self.inner_other_buttons_frame.update_idletasks()
        self.other_preset_canvas.config(scrollregion=self.inner_other_buttons_frame.bbox("all"))

        self.selected_preset = None # Clear selected preset
        self.preset_files = [] # Clear the stored list of presets
        self.source_of_displayed_presets = "unknown" # Reset source
        if self.app_instance and hasattr(self.app_instance, 'load_preset_button'):
            self.app_instance.load_preset_button.config(state=tk.DISABLED)
        
        # Add placeholder labels if no presets are displayed
        ttk.Label(self.inner_mon_buttons_frame, text="No presets displayed.",
                      background="#333333", foreground="white").pack(padx=10, pady=10)
        ttk.Label(self.inner_other_buttons_frame, text="No presets displayed.",
                      background="#333333", foreground="white").pack(padx=10, pady=10)


    def _on_tab_selected(self, event):
        """
        Callback when this tab is selected. Updates the state of the query presets button.
        """
        debug_print("PresetFilesTab selected.", file=__file__, function=inspect.currentframe().f_code.co_name)
        if self.app_instance and self.app_instance.inst:
            self.query_presets_button.config(state=tk.NORMAL)
            # Auto-query presets from device if N9342CN and not already displayed from device
            if self.app_instance.instrument_model == "N9342CN" and self.source_of_displayed_presets != "device":
                debug_print("Auto-querying device presets for N9342CN.", file=__file__, function=inspect.currentframe().f_code.co_name)
                self._query_presets_from_device()
            elif self.source_of_displayed_presets == "device" and self.preset_files:
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

