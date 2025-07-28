# src/marker_tab.py
import tkinter as tk
from tkinter import scrolledtext, filedialog, ttk # Keep other imports
import os
import csv
import inspect
import json # Import json for serializing/deserializing row data

# Import set_marker_and_trace_modes_logic from marker_utils
from utils.marker_utils import set_marker_and_trace_modes_logic
from utils.instrument_control import debug_print, query_current_instrument_settings # Import query_current_instrument_settings
from ref.frequency_bands import MHZ_TO_HZ # Import MHZ_TO_HZ for conversion

# Import new marker utility functions and constants
from utils.marker_utils import SPAN_OPTIONS, set_span_logic


# Removed the hardcoded MARKERS_FILE_PATH, it will now be determined dynamically


class MarkersDisplayTab(ttk.Frame):
    """
    A Tkinter Frame that displays extracted frequency markers in a hierarchical treeview
    and as clickable buttons.
    """
    def __init__(self, master=None, headers=None, rows=None, app_instance=None, console_print_func=None, **kwargs):
        """
        Initializes the MarkersDisplayTab.

        Inputs:
            master (tk.Widget): The parent widget.
            headers (list): A list of column headers for the marker data.
            rows (list): A list of dictionaries, where each dictionary represents
                         a row of marker data with keys matching the headers.
            app_instance (App): The main application instance, used for accessing
                                shared state like instrument connection and focus width.
            console_print_func (function, optional): Function to use for console output.
            **kwargs: Arbitrary keyword arguments for Tkinter Frame.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Initializing MarkersDisplayTab...", file=current_file, function=current_function, console_print_func=console_print_func)

        super().__init__(master, **kwargs)
        self.headers = headers if headers is not None else []
        self.rows = rows if rows is not None else [] # Store full rows data
        self.app_instance = app_instance # Store reference to the main app instance
        self.console_print_func = console_print_func if console_print_func else print # Use provided func or default print

        # Apply style to the main frame (this style is now defined globally in main_app.py)
        self.config(style="Markers.TFrame") 
        self.last_selected_span_button = None # To keep track of the last selected span button
        self.current_span_hz = None # To store the currently active span value in Hz

        # Trace mode state variables
        self.live_mode_var = tk.BooleanVar(self, value=True) # Default to Live
        self.max_hold_mode_var = tk.BooleanVar(self, value=False)
        self.min_hold_mode_var = tk.BooleanVar(self, value=False)
        self.trace_mode_buttons = {} # To store references to trace mode buttons

        # For managing selected device button state across selections
        self.current_selected_device_button = None # Reference to the currently active button widget
        self.selected_device_unique_id = None # Unique ID of the currently selected device
        self.current_selected_device_data = None # NEW: Store the full data of the selected device

        self._create_widgets()


    def _create_widgets(self):
        """
        Creates the widgets for the Markers Display tab, including the treeview
        for zones/groups and the frame for device buttons.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Creating MarkersDisplayTab widgets...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        
        # Main frame for the split layout
        main_split_frame = ttk.Frame(self, style="Markers.TFrame") # Use ttk.Frame
        # Changed to grid layout to accommodate span control frame at the bottom
        main_split_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.grid_rowconfigure(0, weight=1) # Allow main_split_frame to expand
        self.grid_columnconfigure(0, weight=1)

        main_split_frame.grid_columnconfigure(0, weight=1) # Left half
        main_split_frame.grid_columnconfigure(1, weight=1) # Right half
        main_split_frame.grid_rowconfigure(0, weight=1) # Top row for treeview and device buttons
        main_split_frame.grid_rowconfigure(1, weight=0) # Bottom row for span controls (fixed height)
        main_split_frame.grid_rowconfigure(2, weight=0) # Bottom row for trace mode controls (fixed height)
        main_split_frame.grid_rowconfigure(3, weight=0) # NEW: Row for current span/trace display

        # Left Half: Treeview for Zones and Groups
        tree_frame = ttk.LabelFrame(main_split_frame, text="Zones & Groups", padding=(5,5,5,5), style='Dark.TLabelframe') 
        tree_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=5, pady=5)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.zone_group_tree = ttk.Treeview(tree_frame, show="tree", style="Markers.Inner.Treeview") # Apply the new style
        self.zone_group_tree.pack(fill=tk.BOTH, expand=True)

        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.zone_group_tree.yview)
        tree_scrollbar.pack(side=tk.RIGHT, fill="y")
        self.zone_group_tree.configure(yscrollcommand=tree_scrollbar.set)
        
        # Bind selection event to update device buttons
        self.zone_group_tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # Right Half: Buttons for Devices
        buttons_frame = ttk.LabelFrame(main_split_frame, text="Devices", padding=(5,5,5,5), style='Dark.TLabelframe') 
        buttons_frame.grid(row=0, column=1, sticky=tk.NSEW, padx=5, pady=5)
        
        # Use a canvas with a scrollbar for buttons if there are many
        self.buttons_canvas = tk.Canvas(buttons_frame, bg="#1e1e1e", highlightbackground="#1e1e1e") # tk.Canvas as ttk.Canvas doesn't exist
        self.buttons_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        buttons_scrollbar = ttk.Scrollbar(buttons_frame, orient="vertical", command=self.buttons_canvas.yview)
        buttons_scrollbar.pack(side=tk.RIGHT, fill="y")

        self.buttons_canvas.configure(yscrollcommand=buttons_scrollbar.set)
        self.buttons_canvas.bind('<Configure>', lambda e: self.buttons_canvas.configure(scrollregion = self.buttons_canvas.bbox("all")))

        self.inner_buttons_frame = ttk.Frame(self.buttons_canvas, style='Dark.TFrame') # Use ttk.Frame for consistency
        self.buttons_canvas.create_window((0, 0), window=self.inner_buttons_frame, anchor="nw")

        # Configure columns for the grid layout within inner_buttons_frame
        # Changed to 2 columns for buttons to make them wider
        self.inner_buttons_frame.grid_columnconfigure(0, weight=1)
        self.inner_buttons_frame.grid_columnconfigure(1, weight=1) 

        # Now call _populate_zone_group_tree after inner_buttons_frame is initialized
        self._populate_zone_group_tree() 

        # Initially populate with an empty list to clear any previous buttons
        self._populate_device_buttons([]) 

        # --- Span Control Buttons Frame (Bottom of main_split_frame) ---
        span_control_frame = ttk.LabelFrame(main_split_frame, text="Span for Device", padding=(5,5,5,5), style="Dark.TLabelframe") # LabelFrame with title
        span_control_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

        # Configure columns for the buttons within span_control_frame
        for i in range(len(SPAN_OPTIONS)): # Use length of SPAN_OPTIONS for column config
            span_control_frame.grid_columnconfigure(i, weight=1)

        # Create buttons
        self.span_buttons = {}
        col = 0
        # Iterate through SPAN_OPTIONS to create buttons
        for text_key, span_hz_value in SPAN_OPTIONS.items():
            # Format the second line of the button text
            display_value = f"{span_hz_value / MHZ_TO_HZ:.3f} MHz" if span_hz_value >= MHZ_TO_HZ else f"{span_hz_value / 1000:.0f} KHz"
            button_text = f"{text_key}\n{display_value}"

            btn = ttk.Button(span_control_frame, text=button_text, style="Markers.TButton", # Default style for unselected
                             command=lambda s=span_hz_value, t=text_key: self._on_span_button_click(s, self.span_buttons[t], t)) # Pass span_hz, button_widget, and text_key
            
            # Store button reference before gridding to ensure it's in the dict for the lambda
            self.span_buttons[text_key] = btn
            btn.grid(row=0, column=col, padx=2, pady=2, sticky="nsew")
            col += 1
        
        # Set "Normal" as the initially selected button and span
        # Find the "Normal" button and its span value
        normal_span_hz = SPAN_OPTIONS["Normal"]
        normal_button_widget = self.span_buttons["Normal"]
        self._on_span_button_click(normal_span_hz, normal_button_widget, "Normal")


        # --- New: Trace Mode Control Buttons Frame (Below Span Control) ---
        trace_mode_control_frame = ttk.LabelFrame(main_split_frame, text="Trace Mode", padding=(5,5,5,5), style="Dark.TLabelframe")
        trace_mode_control_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

        # Configure columns for the buttons within trace_mode_control_frame
        trace_mode_control_frame.grid_columnconfigure(0, weight=1)
        trace_mode_control_frame.grid_columnconfigure(1, weight=1)
        trace_mode_control_frame.grid_columnconfigure(2, weight=1)

        # Create Trace Mode buttons
        # The command will toggle the associated BooleanVar and then call the update logic
        btn_live = ttk.Button(trace_mode_control_frame, text="Live", style="Markers.TButton",
                              command=lambda: self._on_trace_mode_button_click("Live"))
        btn_max_hold = ttk.Button(trace_mode_control_frame, text="Max Hold", style="Markers.TButton",
                                  command=lambda: self._on_trace_mode_button_click("Max Hold"))
        btn_min_hold = ttk.Button(trace_mode_control_frame, text="Min Hold", style="Markers.TButton",
                                  command=lambda: self._on_trace_mode_button_click("Min Hold"))

        self.trace_mode_buttons["Live"] = {"button": btn_live, "var": self.live_mode_var}
        self.trace_mode_buttons["Max Hold"] = {"button": btn_max_hold, "var": self.max_hold_mode_var}
        self.trace_mode_buttons["Min Hold"] = {"button": btn_min_hold, "var": self.min_hold_mode_var}

        btn_live.grid(row=0, column=0, padx=2, pady=2, sticky="nsew")
        btn_max_hold.grid(row=0, column=1, padx=2, pady=2, sticky="nsew")
        btn_min_hold.grid(row=0, column=2, padx=2, pady=2, sticky="nsew")

        # Initialize button colors based on default states (Live is True, others False)
        self._update_trace_mode_button_styles() # Call this to set initial colors

        # --- NEW: Current Span and Trace Mode Display Box ---
        self.current_settings_frame = ttk.LabelFrame(main_split_frame, text="Current Instrument Settings", padding=(5,5,5,5), style="Dark.TLabelframe")
        self.current_settings_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        self.current_settings_frame.grid_columnconfigure(0, weight=1) # Allow label to expand

        self.current_span_label = ttk.Label(self.current_settings_frame, text="Span: N/A", style="Markers.TLabel")
        self.current_span_label.grid(row=0, column=0, sticky="w", padx=5, pady=2)

        self.current_trace_modes_label = ttk.Label(self.current_settings_frame, text="Trace Modes: N/A", style="Markers.TLabel")
        self.current_trace_modes_label.grid(row=1, column=0, sticky="w", padx=5, pady=2)
        # --- END NEW ---


    def _populate_zone_group_tree(self):
        """
        Populates the Treeview with zones and groups only.
        The tree structure will be: ZONE -> GROUP.
        If GROUP is empty, it will not create a group node, but the markers will still be associated with the zone.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Populating zone/group tree (2 levels)...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        self.zone_group_tree.delete(*self.zone_group_tree.get_children()) # Clear existing data

        # Nested dictionary to store data: {ZONE: {GROUP: [rows]}}
        nested_grouped_data = {}

        for row in self.rows:
            zone = row.get('ZONE', 'Uncategorized Zone').strip()
            group = row.get('GROUP', '').strip() # Get group, strip whitespace, default to empty string

            if zone not in nested_grouped_data:
                nested_grouped_data[zone] = {}
            
            # If group is empty, use a placeholder key to store markers directly under the zone.
            # Otherwise, use the group name.
            group_key = group if group else "__NO_GROUP__"

            if group_key not in nested_grouped_data[zone]:
                nested_grouped_data[zone][group_key] = []
            
            nested_grouped_data[zone][group_key].append(row)

        for zone_name in sorted(nested_grouped_data.keys()):
            zone_id = self.zone_group_tree.insert("", "end", text=zone_name, open=True, tags=('zone',))
            
            for group_key in sorted(nested_grouped_data[zone_name].keys()):
                if group_key != "__NO_GROUP__": # Only create a group node if a group name exists
                    self.zone_group_tree.insert(zone_id, "end", text=group_key, open=True, tags=('group',))
                # No leaf nodes for individual markers here, as per the new requirement.
                # The markers will be retrieved directly from the selected zone/group in _on_tree_select.

        # Clear device buttons when tree is repopulated
        # Removed: self._populate_device_buttons([]) # This was causing devices to disappear

    def _on_tree_select(self, event):
        """
        Handles selection events in the zone/group treeview.
        Populates the device buttons based on the selected zone or group.
        Also resets the selected device button state.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Tree item selected...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        selected_items = self.zone_group_tree.selection()
        
        # Reset selected device button state when tree selection changes
        if self.current_selected_device_button:
            self.current_selected_device_button.config(style="DeviceButton.TButton") # Revert old button to neutral
            self.current_selected_device_button = None
            self.selected_device_unique_id = None
            self.current_selected_device_data = None # NEW: Clear stored device data
        
        if not selected_items:
            self._populate_device_buttons([])
            return

        selected_rows_data = []

        for item_id in selected_items:
            item_tags = self.zone_group_tree.item(item_id, 'tags')
            
            if 'zone' in item_tags:
                zone_name = self.zone_group_tree.item(item_id, 'text')
                # Collect all markers belonging to this zone
                for row in self.rows:
                    if row.get('ZONE', '').strip() == zone_name:
                        selected_rows_data.append(row)
            elif 'group' in item_tags:
                group_name = self.zone_group_tree.item(item_id, 'text')
                parent_id = self.zone_group_tree.parent(item_id)
                zone_name = self.zone_group_tree.item(parent_id, 'text')
                # Collect all markers belonging to this specific group within this zone
                for row in self.rows:
                    if row.get('ZONE', '').strip() == zone_name and row.get('GROUP', '').strip() == group_name:
                        selected_rows_data.append(row)
            # No 'marker' tag check needed here as individual markers are not tree nodes anymore

        self._populate_device_buttons(selected_rows_data)

    def _populate_device_buttons(self, devices_to_display):
        """
        Populates the right-hand frame with clickable buttons for each device.
        Buttons will be approximately 50% of the device box width, and display
        NAME, DEVICE, and FREQ (in MHz) on three lines.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Populating device buttons with {len(devices_to_display)} devices...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        # Clear existing buttons
        for widget in self.inner_buttons_frame.winfo_children():
            widget.destroy()

        if not devices_to_display:
            ttk.Label(self.inner_buttons_frame, text="Select a zone or group from the left to display devices.",
                      background="#1e1e1e", foreground="#cccccc", style='Markers.TLabel').grid(row=0, column=0, columnspan=2, padx=5, pady=5)
            self.inner_buttons_frame.update_idletasks()
            self.buttons_canvas.config(scrollregion=self.buttons_canvas.bbox("all"))
            return

        row_idx = 0
        col_idx = 0
        for i, device_data in enumerate(devices_to_display):
            name = device_data.get('NAME', '').strip()
            device = device_data.get('DEVICE', '').strip()
            freq_mhz = device_data.get('FREQ') 

            if freq_mhz is not None:
                try:
                    # Create a unique ID for this device
                    unique_device_id = f"{device_data.get('ZONE', '')}-{device_data.get('GROUP', '')}-{device_data.get('DEVICE', '')}-{device_data.get('NAME', '')}-{freq_mhz}"
                    
                    # Format button text for three lines
                    display_name = name if name else "N/A Name"
                    display_device = device if device and device.lower() != "none - none - n/a" else "N/A Device"
                    
                    # Format frequency for display without decimal if it's a whole number
                    display_freq_mhz = int(float(freq_mhz)) if float(freq_mhz) == int(float(freq_mhz)) else f"{float(freq_mhz):.3f}"
                    button_text = f"{display_name}\n{display_device}\n{display_freq_mhz} MHz"
                    
                    # Use "DeviceButton.TButton" for the unselected state
                    btn = ttk.Button(self.inner_buttons_frame, text=button_text, style="DeviceButton.TButton", 
                                     command=lambda d=device_data, btn_w=None: self._on_device_button_click(d, btn_w))
                    
                    # Pass the button widget reference to the lambda after it's created
                    btn.configure(command=lambda d=device_data, u_id=unique_device_id, b_w=btn: self._on_device_button_click(d, b_w))

                    # Set initial style based on whether this device was previously selected
                    if unique_device_id == self.selected_device_unique_id:
                        btn.config(style="SelectedDevice.TButton")
                        self.current_selected_device_button = btn # Re-establish reference
                    
                    # Use sticky="nsew" to make buttons expand within their grid cells
                    btn.grid(row=row_idx, column=col_idx, padx=5, pady=5, sticky="nsew")
                    
                    col_idx += 1
                    if col_idx >= 2: # Two columns per row
                        col_idx = 0
                        row_idx += 1
                except ValueError:
                    debug_print(f"Could not convert frequency '{freq_mhz}' to float for button. Skipping.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            else:
                debug_print(f"Frequency not found for device '{name}'. Skipping button.", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # Ensure columns in inner_buttons_frame expand to fill available space
        self.inner_buttons_frame.grid_columnconfigure(0, weight=1)
        self.inner_buttons_frame.grid_columnconfigure(1, weight=1)
        # Ensure rows also expand if needed, though buttons will dictate row height
        for r in range(row_idx + 1):
            self.inner_buttons_frame.grid_rowconfigure(r, weight=1)


        self.inner_buttons_frame.update_idletasks() # Ensure layout is updated before calculating scrollregion
        self.buttons_canvas.config(scrollregion=self.buttons_canvas.bbox("all"))

    def _on_device_button_click(self, device_data, clicked_button_widget):
        """
        Callback for device buttons. Sets the instrument's focus frequency and a marker.
        Uses the currently selected span from the span buttons, and current trace modes.
        Manages button selection state.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__

        freq_hz = float(device_data.get('FREQ')) * MHZ_TO_HZ
        name = device_data.get('NAME', '').strip() # Ensure name is stripped for display
        unique_device_id = f"{device_data.get('ZONE', '')}-{device_data.get('GROUP', '')}-{device_data.get('DEVICE', '')}-{device_data.get('NAME', '')}-{device_data.get('FREQ', '')}"

        # Format frequency for display without decimal if it's a whole number
        display_freq_mhz = int(freq_hz / MHZ_TO_HZ) if (freq_hz / MHZ_TO_HZ) == int(freq_hz / MHZ_TO_HZ) else f"{freq_hz / MHZ_TO_HZ:.3f}"
        self.console_print_func(f"\nSetting instrument to '{name}' at {display_freq_mhz} MHz...")
        debug_print(f"Device button clicked: {name} at {freq_hz} Hz (ID: {unique_device_id})", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # Handle button visual toggle
        if self.current_selected_device_button and self.current_selected_device_button != clicked_button_widget:
            self.current_selected_device_button.config(style="DeviceButton.TButton") # Revert old button to neutral
        
        clicked_button_widget.config(style="SelectedDevice.TButton") # Select new button (orange)
        self.current_selected_device_button = clicked_button_widget
        self.selected_device_unique_id = unique_device_id
        self.current_selected_device_data = device_data # NEW: Store the full device data

        if self.app_instance and self.app_instance.inst:
            # Use the currently selected span, or fall back to default focus width
            span_to_use_hz = self.current_span_hz if self.current_span_hz is not None else \
                             float(self.app_instance.desired_default_focus_width_var.get()) * MHZ_TO_HZ
            
            # Call set_span_logic to set frequency, span, and trace modes
            # Pass the current state of the trace mode variables
            set_span_logic(self.app_instance.inst, span_to_use_hz, freq_hz, 
                           self.live_mode_var.get(), self.max_hold_mode_var.get(), self.min_hold_mode_var.get(),
                           self.console_print_func)
            
            # Keep set_marker_and_trace_modes_logic for marker setup (if it does more than just trace modes)
            set_marker_and_trace_modes_logic(self.app_instance, freq_hz, name, self.console_print_func) 
            self._update_current_settings_display() # Update display after device click
        else:
            self.console_print_func("⚠️ Warning: Cannot set focus frequency: Instrument not connected.")
            debug_print("Cannot set focus frequency: Instrument not connected.", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _on_span_button_click(self, span_hz, button_widget, button_text_key):
        """
        Callback for span control buttons. Changes the instrument's span and toggles button color/font.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        self.console_print_func(f"Setting span to {span_hz / MHZ_TO_HZ:.3f} MHz...")
        debug_print(f"Span button clicked: Setting span to {span_hz} Hz (Button: {button_text_key})", file=current_file, function=current_function, console_print_func=self.console_print_func)
        
        # Update the stored current span
        self.current_span_hz = span_hz

        # Toggle button styles for visual feedback (orange/blue accordingly)
        for text_key, btn in self.span_buttons.items():
            if btn == button_widget:
                # Apply the style that includes orange background
                btn.config(style="SelectedSpan.TButton") 
            else:
                # Revert to the default style for unselected buttons
                btn.config(style="Markers.TButton") 
        
        # If the instrument is connected, send the commands
        if self.app_instance and self.app_instance.inst:
            center_freq_to_use = None
            if self.current_selected_device_data:
                try:
                    center_freq_to_use = float(self.current_selected_device_data.get('FREQ')) * MHZ_TO_HZ
                    self.console_print_func(f"Re-centering on selected device at {center_freq_to_use / MHZ_TO_HZ:.3f} MHz with new span.")
                except ValueError:
                    debug_print("Could not convert selected device frequency to float for re-centering.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                    center_freq_to_use = None # Fallback if conversion fails

            # Call the centralized set_span_logic from marker_utils
            set_span_logic(self.app_instance.inst, span_to_use_hz=span_hz, center_freq_hz=center_freq_to_use, 
                           live_mode=self.live_mode_var.get(), max_hold_mode=self.max_hold_mode_var.get(), min_hold_mode=self.min_hold_mode_var.get(),
                           console_print_func=self.console_print_func)
            self._update_current_settings_display() # NEW: Update display after span change
        else:
            self.console_print_func("⚠️ Warning: Cannot set span: Instrument not connected.")
            debug_print("Cannot set span: Instrument not connected.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            

    def _on_trace_mode_button_click(self, mode_name):
        """
        Callback for trace mode buttons. Toggles the state of the clicked button's
        associated BooleanVar and updates the instrument's trace modes.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Trace mode button clicked: {mode_name}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # Toggle the state of the clicked button's variable
        # Since these are independent toggles, just invert the current state
        if mode_name == "Live":
            self.live_mode_var.set(not self.live_mode_var.get())
        elif mode_name == "Max Hold":
            self.max_hold_mode_var.set(not self.max_hold_mode_var.get())
        elif mode_name == "Min Hold":
            self.min_hold_mode_var.set(not self.min_hold_mode_var.get())

        # Update button styles
        self._update_trace_mode_button_styles()

        # If instrument is connected, send the updated trace mode commands
        if self.app_instance and self.app_instance.inst:
            # Get the current span (or default) to pass to set_span_logic
            span_to_use_hz = self.current_span_hz if self.current_span_hz is not None else \
                             float(self.app_instance.desired_default_focus_width_var.get()) * MHZ_TO_HZ
            
            # Determine center frequency if a device is selected
            center_freq_to_use = None
            if self.current_selected_device_data:
                try:
                    center_freq_to_use = float(self.current_selected_device_data.get('FREQ')) * MHZ_TO_HZ
                except ValueError:
                    debug_print("Could not convert selected device frequency to float for re-centering (trace mode click).", file=current_file, function=current_function, console_print_func=self.console_print_func)
                    center_freq_to_use = None # Fallback if conversion fails

            # Call set_span_logic with current span and updated trace modes
            set_span_logic(self.app_instance.inst, span_to_use_hz, center_freq_to_use,
                           self.live_mode_var.get(), self.max_hold_mode_var.get(), self.min_hold_mode_var.get(),
                           self.console_print_func)
            self._update_current_settings_display() # NEW: Update display after trace mode change
        else:
            self.console_print_func("⚠️ Warning: Cannot set trace mode: Instrument not connected.")
            debug_print("Cannot set trace mode: Instrument not connected.", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _update_trace_mode_button_styles(self):
        """
        Updates the visual style of the trace mode buttons based on their BooleanVar states.
        """
        for mode_name, data in self.trace_mode_buttons.items():
            button = data["button"]
            var = data["var"]
            if var.get():
                button.config(style="SelectedSpan.TButton") # Use orange for selected
            else:
                button.config(style="Markers.TButton") # Use default blue for unselected

    def _update_current_settings_display(self):
        """
        Updates the labels in the "Current Instrument Settings" display box.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Updating current settings display...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        display_span = "N/A"
        if self.current_span_hz is not None:
            if self.current_span_hz == 0.0: # Full Span case
                display_span = "Full Span"
            elif self.current_span_hz >= MHZ_TO_HZ:
                display_span = f"{self.current_span_hz / MHZ_TO_HZ:.3f} MHz"
            else:
                display_span = f"{self.current_span_hz / 1000:.0f} KHz"
        self.current_span_label.config(text=f"Span: {display_span}")
        debug_print(f"Displaying Span: {display_span}", file=current_file, function=current_function, console_print_func=self.console_print_func)


        active_modes = []
        if self.live_mode_var.get():
            active_modes.append("Live")
        if self.max_hold_mode_var.get():
            active_modes.append("Max Hold")
        if self.min_hold_mode_var.get():
            active_modes.append("Min Hold")
        
        if active_modes:
            self.current_trace_modes_label.config(text=f"Trace Modes: {', '.join(active_modes)}")
            debug_print(f"Displaying Trace Modes: {', '.join(active_modes)}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.current_trace_modes_label.config(text="Trace Modes: None Active")
            debug_print("Displaying Trace Modes: None Active", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def update_markers_data(self, headers, rows):
        """
        Updates the data displayed in the markers tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Updating markers data...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        self.headers = headers
        self.rows = rows
        self._populate_zone_group_tree() # Repopulate the treeview with new data
        # Removed: self._populate_device_buttons([]) # This was causing devices to disappear
        self._update_current_settings_display() # NEW: Update display after data load


    def _on_tab_selected(self, event):
        """
        Callback when this tab is selected. Checks for and loads MARKERS.CSV.
        Also queries the instrument for current span and trace modes to update GUI.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        self.console_print_func("MarkersDisplayTab selected. Checking for MARKERS.CSV...")
        debug_print("MarkersDisplayTab selected. Checking for MARKERS.CSV...", file=current_file, function=current_function, console_print_func=self.console_print_func)
        
        # Dynamically determine MARKERS.CSV path from the main app's output folder
        markers_file_path = None
        if self.app_instance and hasattr(self.app_instance, 'output_folder_var'):
            output_folder = self.app_instance.output_folder_var.get()
            if output_folder:
                markers_file_path = os.path.join(output_folder, 'MARKERS.CSV')
                debug_print(f"Attempting to load MARKERS.CSV from configured output folder: {markers_file_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            else:
                self.console_print_func("⚠️ Warning: Output folder not configured in main app. Cannot check for MARKERS.CSV.")
                debug_print("Output folder not configured in main app. Cannot check for MARKERS.CSV.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.console_print_func("⚠️ Warning: App instance or output_folder_var not available. Cannot check for MARKERS.CSV.")
            debug_print("App instance or output_folder_var not available. Cannot check for MARKERS.CSV.", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if markers_file_path and os.path.exists(markers_file_path):
            debug_print(f"MARKERS.CSV found at: {markers_file_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            try:
                headers = []
                rows = []
                with open(markers_file_path, mode='r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    headers = reader.fieldnames
                    for row_data in reader:
                        rows.append(row_data)
                
                if headers and rows:
                    debug_print(f"Loaded {len(rows)} markers from MARKERS.CSV.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                    self.update_markers_data(headers, rows)
                    # NEW: After updating data and populating tree, try to select the first item
                    if self.zone_group_tree.get_children():
                        first_item = self.zone_group_tree.get_children()[0]
                        self.zone_group_tree.selection_set(first_item)
                        self.zone_group_tree.focus(first_item)
                        self._on_tree_select(None) # Manually trigger the selection logic
                    self.console_print_func(f"✅ Loaded {len(rows)} markers from MARKERS.CSV.")
                else:
                    debug_print("MARKERS.CSV is empty or has no data rows.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                    self.console_print_func("ℹ️ Info: The MARKERS.CSV file was found but contains no data.")
                    self.update_markers_data([], []) # Clear any existing display
            except Exception as e:
                debug_print(f"Error loading MARKERS.CSV: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
                self.console_print_func(f"❌ Error Loading Markers: An error occurred while loading MARKERS.CSV: {e}")
                self.update_markers_data([], []) # Clear any existing display on error
        else:
            debug_print(f"MARKERS.CSV not found or path not determined. Path: {markers_file_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            self.console_print_func("ℹ️ Info: MARKERS.CSV not found. Please generate a report first.")
            self.update_markers_data([], []) # Ensure display is clear if file doesn't exist

        # Query instrument for current span and trace modes on tab selection
        if self.app_instance and self.app_instance.inst:
            center_freq, current_span, rbw = query_current_instrument_settings(self.app_instance.inst, MHZ_TO_HZ, self.console_print_func)
            if current_span is not None:
                current_span_hz = current_span * MHZ_TO_HZ
                self.current_span_hz = current_span_hz # Update stored span

                # Find and highlight the corresponding span button
                found_match = False
                for text_key, span_hz_value in SPAN_OPTIONS.items():
                    if abs(span_hz_value - current_span_hz) < 1: # Allow for small floating point differences
                        self.span_buttons[text_key].config(style="SelectedSpan.TButton")
                        self.last_selected_span_button = self.span_buttons[text_key]
                        self.console_print_func(f"✅ GUI span button updated to match instrument: {text_key}")
                        found_match = True
                    else:
                        self.span_buttons[text_key].config(style="Markers.TButton") # Revert others
                
                if not found_match: # If no matching button, default to "Normal" and highlight it
                    normal_span_hz = SPAN_OPTIONS["Normal"]
                    normal_button_widget = self.span_buttons["Normal"]
                    normal_button_widget.config(style="SelectedSpan.TButton")
                    self.last_selected_span_button = normal_button_widget
                    self.current_span_hz = normal_span_hz # Ensure internal state is "Normal"
                    self.console_print_func("ℹ️ Instrument span did not match a button. Defaulting GUI span to 'Normal'.")

            else: # If instrument not connected or span not queried, default GUI span to "Normal"
                normal_span_hz = SPAN_OPTIONS["Normal"]
                normal_button_widget = self.span_buttons["Normal"]
                normal_button_widget.config(style="SelectedSpan.TButton")
                self.last_selected_span_button = normal_button_widget
                self.current_span_hz = normal_span_hz # Ensure internal state is "Normal"
                self.console_print_func("ℹ️ Instrument not connected or span not queried. Defaulting GUI span to 'Normal'.")


            # Query and update trace modes (assuming instrument_control.py can query them)
            # This would require new query commands in instrument_control.py like :TRAC1:MODE?
            # For now, we'll assume the instrument is in a known state or rely on initial setup.
            # If you have SCPI commands to query current trace modes, you'd add them here
            # and update self.live_mode_var, self.max_hold_mode_var, self.min_hold_mode_var
            # then call self._update_trace_mode_button_styles()
        else: # If no instrument, ensure "Normal" is selected by default in GUI
            normal_span_hz = SPAN_OPTIONS["Normal"]
            normal_button_widget = self.span_buttons["Normal"]
            normal_button_widget.config(style="SelectedSpan.TButton")
            self.last_selected_span_button = normal_button_widget
            self.current_span_hz = normal_span_hz # Ensure internal state is "Normal"
            self.console_print_func("ℹ️ Instrument not connected. Defaulting GUI span to 'Normal'.")

        self._update_current_settings_display() # NEW: Always update display on tab selection
