# src/marker_logic.py
import tkinter as tk
# from tkinter import messagebox, scrolledtext, filedialog, ttk # Removed messagebox
from tkinter import scrolledtext, filedialog, ttk # Keep other imports
import os
import csv
import inspect
import json # Import json for serializing/deserializing row data

# Import instrument_logic for setting focus frequency
from src.instrument_logic import set_focus_frequency_logic, set_marker_and_trace_modes_logic
from utils.instrument_control import debug_print # Import debug_print

# Removed the hardcoded MARKERS_FILE_PATH, it will now be determined dynamically


class MarkersDisplayTab(ttk.Frame):
    """
    A Tkinter Frame that displays extracted frequency markers in a hierarchical treeview
    and as clickable buttons.
    """
    def __init__(self, master=None, headers=None, rows=None, app_instance=None, **kwargs):
        """
        Initializes the MarkersDisplayTab.

        Inputs:
            master (tk.Widget): The parent widget.
            headers (list): A list of column headers for the marker data.
            rows (list): A list of dictionaries, where each dictionary represents
                         a row of marker data with keys matching the headers.
            app_instance (App): The main application instance, used for accessing
                                shared state like instrument connection and focus width.
            **kwargs: Arbitrary keyword arguments for Tkinter Frame.
        Process:
            1. Calls the parent `ttk.Frame` constructor.
            2. Stores `headers`, `rows`, and `app_instance`.
            3. Initializes `self.tree` (Treeview widget) and `self.device_buttons_frame`.
            4. Calls `_create_widgets` to set up the GUI.
            5. Calls `update_markers_data` to populate the initial display if data is provided.
            6. Initializes `current_span_hz` for span control.
            7. Initializes `self.current_selected_span_button` to track the active span button.
        Outputs: None
        """
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self.headers = headers if headers is not None else []
        self.rows = rows if rows is not None else []
        self.current_span_hz = 10000000 # Default span 10 MHz
        self.current_selected_span_button = None # To track the currently selected span button

        self._create_widgets()

        if self.headers and self.rows:
            self.update_markers_data(self.headers, self.rows)

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Markers Display tab.
        """
        self.grid_columnconfigure(0, weight=1) # Treeview column
        self.grid_columnconfigure(1, weight=1) # Device buttons column
        self.grid_rowconfigure(0, weight=1) # Main content row
        self.grid_rowconfigure(1, weight=0) # Span control buttons row

        # Left Pane: Treeview for Zones & Groups
        tree_frame = ttk.LabelFrame(self, text="Zones & Groups")
        tree_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_frame, style="Markers.Treeview", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew")

        tree_scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        tree_scrollbar_y.grid(row=0, column=1, sticky="ns")
        self.tree.config(yscrollcommand=tree_scrollbar_y.set)

        tree_scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        tree_scrollbar_x.grid(row=1, column=0, sticky="ew")
        self.tree.config(xscrollcommand=tree_scrollbar_x.set)

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # Right Pane: Frame for Device Buttons (scrollable)
        device_buttons_outer_frame = ttk.LabelFrame(self, text="Devices")
        device_buttons_outer_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        device_buttons_outer_frame.grid_rowconfigure(0, weight=1)
        device_buttons_outer_frame.grid_columnconfigure(0, weight=1)

        self.device_canvas = tk.Canvas(device_buttons_outer_frame, borderwidth=0, highlightthickness=0, bg='#1e1e1e')
        self.device_canvas.grid(row=0, column=0, sticky="nsew")

        device_scrollbar = ttk.Scrollbar(device_buttons_outer_frame, orient="vertical", command=self.device_canvas.yview)
        device_scrollbar.grid(row=0, column=1, sticky="ns")
        self.device_canvas.config(yscrollcommand=device_scrollbar.set)

        self.device_buttons_frame = ttk.Frame(self.device_canvas, style='Dark.TFrame')
        self.device_canvas.create_window((0, 0), window=self.device_buttons_frame, anchor="nw")

        self.device_buttons_frame.bind("<Configure>", lambda e: self.device_canvas.config(scrollregion=self.device_canvas.bbox("all")))
        self.device_canvas.bind('<Enter>', self._bind_device_mouse_wheel)
        self.device_canvas.bind('<Leave>', self._unbind_device_mouse_wheel)

        # Bottom Pane: Span Control Buttons
        span_control_frame = ttk.LabelFrame(self, text="Span Control")
        span_control_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        span_control_frame.grid_columnconfigure(0, weight=1)
        span_control_frame.grid_columnconfigure(1, weight=1)
        span_control_frame.grid_columnconfigure(2, weight=1)
        span_control_frame.grid_columnconfigure(3, weight=1)
        span_control_frame.grid_columnconfigure(4, weight=1)

        # Define span options (in Hz)
        span_options = [
            (10000, "10 KHz"),
            (100000, "100 KHz"),
            (1000000, "1 MHz"),
            (10000000, "10 MHz"), # Default
            (100000000, "100 MHz")
        ]

        for i, (span_hz, text) in enumerate(span_options):
            btn = ttk.Button(span_control_frame, text=text,
                             command=lambda s=span_hz: self._on_span_button_click(s, btn, text),
                             style='Markers.TButton')
            btn.grid(row=0, column=i, padx=2, pady=2, sticky="ew")
            if span_hz == self.current_span_hz:
                self.current_selected_span_button = btn
                btn.config(style='SelectedSpan.TButton') # Apply selected style

    def _bind_device_mouse_wheel(self, event):
        """Binds mouse wheel events for the device buttons canvas."""
        event.widget.bind_all("<MouseWheel>", self._on_device_mouse_wheel)
        event.widget.bind_all("<Button-4>", self._on_device_mouse_wheel) # For Linux
        event.widget.bind_all("<Button-5>", self._on_device_mouse_wheel) # For Linux

    def _unbind_device_mouse_wheel(self, event):
        """Unbinds mouse wheel events for the device buttons canvas."""
        event.widget.unbind_all("<MouseWheel>")
        event.widget.unbind_all("<Button-4>")
        event.widget.unbind_all("<Button-5>")

    def _on_device_mouse_wheel(self, event):
        """Handles mouse wheel scrolling for the device buttons canvas."""
        if sys.platform == "darwin":
            self.device_canvas.yview_scroll(-1 * int(event.delta), "units")
        elif event.num == 4: # Linux scroll up
            self.device_canvas.yview_scroll(-1, "units")
        elif event.num == 5: # Linux scroll down
            self.device_canvas.yview_scroll(1, "units")
        else: # Windows
            self.device_canvas.yview_scroll(-1 * int(event.delta/120), "units")

    def _populate_zone_group_tree(self):
        """
        Populates the treeview with zones and groups from the marker data.
        """
        self.tree.delete(*self.tree.get_children()) # Clear existing tree
        
        # Define treeview columns
        self.tree["columns"] = ("Frequency", "Type")
        self.tree.column("#0", width=200, minwidth=150, stretch=tk.NO, anchor="w") # Zone/Group Name
        self.tree.column("Frequency", width=100, minwidth=80, stretch=tk.NO, anchor="center")
        self.tree.column("Type", width=100, minwidth=80, stretch=tk.NO, anchor="center")

        self.tree.heading("#0", text="Zone / Group")
        self.tree.heading("Frequency", text="Freq (MHz)")
        self.tree.heading("Type", text="Type")

        # Group data by ZONE, then by GROUP
        zones = {}
        for row in self.rows:
            zone = row.get("ZONE", "Uncategorized Zone")
            group = row.get("GROUP", "Uncategorized Group")
            if zone not in zones:
                zones[zone] = {}
            if group not in zones[zone]:
                zones[zone][group] = []
            zones[zone][group].append(row)

        for zone, groups in sorted(zones.items()):
            zone_id = self.tree.insert("", "end", text=zone, open=False, tags=("zone",))
            for group, devices in sorted(groups.items()):
                # Store the full device data with the group for easy retrieval
                group_id = self.tree.insert(zone_id, "end", text=group, open=False, tags=("group",),
                                            values=("", "", json.dumps(devices))) # Store devices as JSON string
                
        debug_print(f"Treeview populated with {len(zones)} zones.", file=__file__, function=inspect.currentframe().f_code.co_name)

    def _on_tree_select(self, event):
        """
        Handles selection in the treeview. When a group is selected,
        it populates the device buttons frame with devices from that group.
        """
        selected_item = self.tree.focus()
        if not selected_item:
            self._populate_device_buttons([]) # Clear buttons if nothing selected
            return

        item_tags = self.tree.item(selected_item, "tags")
        if "group" in item_tags:
            # Retrieve the devices data stored as a JSON string in the values
            item_values = self.tree.item(selected_item, "values")
            if len(item_values) > 2 and item_values[2]:
                devices_json = item_values[2]
                devices = json.loads(devices_json)
                self._populate_device_buttons(devices)
            else:
                selfprint("🚫 No device data found for this group.")
                debug_print("No device data found for selected group.", file=__file__, function=inspect.currentframe().f_code.co_name)
                self._populate_device_buttons([]) # Clear buttons
        else:
            self._populate_device_buttons([]) # Clear buttons if a zone or empty space is selected

    def _populate_device_buttons(self, devices_to_display):
        """
        Creates and arranges buttons for each device in the right pane.
        """
        # Clear existing buttons
        for widget in self.device_buttons_frame.winfo_children():
            widget.destroy()

        if not devices_to_display:
            ttk.Label(self.device_buttons_frame, text="Select a group from the left to see devices.",
                      foreground="#cccccc", background="#1e1e1e").grid(row=0, column=0, padx=5, pady=5, sticky="w")
            return

        row_idx = 0
        col_idx = 0
        for i, device in enumerate(devices_to_display):
            device_name = device.get("NAME", "N/A")
            device_freq = device.get("FREQ", "N/A")
            device_type = device.get("DEVICE", "N/A") # Using DEVICE column for type if available, otherwise "N/A"

            button_text = f"{device_name}\n{device_freq} MHz\n({device_type})"
            
            # Use a frame for each button to allow for more complex layouts if needed
            btn_frame = ttk.Frame(self.device_buttons_frame, style='Dark.TFrame')
            btn_frame.grid(row=row_idx, column=col_idx, padx=5, pady=5, sticky="nsew")
            btn_frame.grid_columnconfigure(0, weight=1)

            btn = ttk.Button(btn_frame, text=button_text,
                             command=lambda f=device_freq, n=device_name: self._on_device_button_click(f, n),
                             style='Markers.TButton')
            btn.grid(row=0, column=0, sticky="ew")

            col_idx += 1
            if col_idx >= 2: # 2 columns per row for device buttons
                col_idx = 0
                row_idx += 1
        debug_print(f"Populated {len(devices_to_display)} device buttons.", file=__file__, function=inspect.currentframe().f_code.co_name)

    def _on_device_button_click(self, frequency_mhz, device_name):
        """
        Handles a click on a device button. Sets the instrument's focus frequency
        and a marker at that frequency.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__

        if not self.app_instance.inst:
            print("🚫 Not connected to instrument. Cannot set frequency or marker.")
            debug_print("Not connected to instrument, cannot set frequency or marker.", file=current_file, function=current_function)
            # tk.messagebox.showwarning("Not Connected", "Please connect to an instrument first.") # Removed messagebox
            return

        try:
            freq_hz = float(frequency_mhz) * MHZ_TO_HZ
            span_hz = self.current_span_hz # Use the currently selected span

            # Set focus frequency and span
            set_focus_frequency_logic(self.app_instance, freq_hz, span_hz)
            
            # Set marker
            set_marker_and_trace_modes_logic(self.app_instance, freq_hz, device_name)

        except ValueError:
            print(f"❌ Invalid frequency value: {frequency_mhz} MHz. Cannot set instrument focus.")
            debug_print(f"Invalid frequency value: {frequency_mhz} MHz.", file=current_file, function=current_function)
            # tk.messagebox.showerror("Invalid Frequency", f"Could not set instrument frequency. Invalid value: {frequency_mhz} MHz") # Removed messagebox
        except Exception as e:
            print(f"❌ An unexpected error occurred while setting frequency/marker: {e}")
            debug_print(f"Error setting frequency/marker: {e}", file=current_file, function=current_function)
            # tk.messagebox.showerror("Error", f"An unexpected error occurred: {e}") # Removed messagebox

    def _on_span_button_click(self, span_hz, button_widget, button_text):
        """
        Handles a click on a span control button. Updates the current span and
        highlights the selected button.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__

        # Reset style of previously selected button
        if self.current_selected_span_button:
            self.current_selected_span_button.config(style='Markers.TButton')
        
        # Set new selected button and update its style
        self.current_selected_span_button = button_widget
        self.current_selected_span_button.config(style='SelectedSpan.TButton')
        self.current_span_hz = span_hz
        print(f"Selected span: {button_text} ({span_hz / MHZ_TO_HZ:.3f} MHz)")
        debug_print(f"Selected span: {button_text} ({span_hz / MHZ_TO_HZ:.3f} MHz)", file=current_file, function=current_function)

        # Optionally, apply the new span to the instrument immediately if connected
        if self.app_instance.inst:
            # We don't have a specific center frequency to set here,
            # so we might just update the span without changing center freq,
            # or use the last known center freq if available from app_instance.
            # For now, let's just log the change. The next device button click will use it.
            print("Note: Span change will apply on next device button click or explicit frequency set.")
            debug_print("Span change will apply on next device button click.", file=current_file, function=current_function)
        else:
            print("🚫 Not connected to instrument. Span change will take effect when connected.")
            debug_print("Not connected to instrument. Span change will take effect when connected.", file=current_file, function=current_function)


    def update_markers_data(self, headers, rows):
        """
        Updates the internal marker data and refreshes the display.
        """
        self.headers = headers
        self.rows = rows
        self._populate_zone_group_tree()
        self._populate_device_buttons([]) # Clear device buttons when data is updated

    def _on_tab_selected(self, event):
        """
        Callback for when this tab is selected.
        Dynamically determines MARKERS.CSV path and loads data.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__

        debug_print("Markers Display Tab selected. Attempting to load MARKERS.CSV...", file=current_file, function=current_function)

        # Determine MARKERS.CSV path based on the current output folder
        output_folder = self.app_instance.output_folder_var.get()
        markers_file_path = None
        if output_folder and os.path.isdir(output_folder):
            markers_file_path = os.path.join(output_folder, "MARKERS.CSV")
            debug_print(f"Looking for MARKERS.CSV at: {markers_file_path}", file=current_file, function=current_function)
        else:
            print("🚫 Output directory not set or does not exist. Cannot load MARKERS.CSV.")
            debug_print("Output directory not set or does not exist.", file=current_file, function=current_function)
            self.update_markers_data([], []) # Clear any existing display
            return

        if markers_file_path and os.path.exists(markers_file_path):
            try:
                headers = []
                rows = []
                with open(markers_file_path, 'r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    headers = reader.fieldnames
                    for row_data in reader:
                        # Convert FREQ to float if possible, keep original if not
                        if "FREQ" in row_data and row_data["FREQ"]:
                            try:
                                row_data["FREQ"] = float(row_data["FREQ"])
                            except ValueError:
                                # Keep as string if conversion fails, or set to None/error indicator
                                row_data["FREQ"] = "Invalid" # Or keep as original string
                        rows.append(row_data)
                
                if headers and rows:
                    debug_print(f"Loaded {len(rows)} markers from MARKERS.CSV.", file=current_file, function=current_function)
                    self.update_markers_data(headers, rows)
                else:
                    debug_print("MARKERS.CSV is empty or has no data rows.", file=current_file, function=current_function)
                    # Changed messagebox to debug_print
                    debug_print("No Markers: The MARKERS.CSV file was found but contains no data.", file=current_file, function=current_function)
                    self.update_markers_data([], []) # Clear any existing display
            except Exception as e:
                debug_print(f"Error loading MARKERS.CSV: {e}", file=current_file, function=current_function)
                # Changed messagebox to debug_print
                debug_print(f"Error Loading Markers: An error occurred while loading MARKERS.CSV: {e}", file=current_file, function=current_function)
                self.update_markers_data([], []) # Clear any existing display on error
        else:
            debug_print(f"MARKERS.CSV not found or path not determined. Path: {markers_file_path}", file=current_file, function=current_function)
            # Changed messagebox to debug_print
            debug_print("No Markers File: MARKERS.CSV not found. Please generate a report first.", file=current_file, function=current_function)
            self.update_markers_data([], []) # Ensure display is clear if file doesn't exist

