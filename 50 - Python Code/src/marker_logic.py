# src/marker_logic.py
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, ttk
import os
import csv
import inspect
import json # Import json for serializing/deserializing row data

# Import instrument_logic for setting focus frequency
from src.instrument_logic import set_focus_frequency_logic, set_marker_and_trace_modes_logic
from utils.instrument_control import debug_print # Import debug_print

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
            3. Configures the frame's style and layout.
            4. Calls `create_widgets()` to build the GUI elements.
        Outputs: None (modifies GUI state)
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Initializing MarkersDisplayTab...", file=current_file, function=current_function)

        super().__init__(master, **kwargs)
        self.headers = headers if headers is not None else []
        self.rows = rows if rows is not None else [] # Store full rows data
        self.app_instance = app_instance # Store reference to the main app instance

        # Configure style for this frame's widgets
        style = ttk.Style()
        style.configure("Markers.TFrame", background="#000000") # Dark background for the main frame
        style.configure("Markers.TLabel", background="#000000", foreground="white")
        style.configure("Markers.Treeview.Heading", background="#3a3a3a", foreground="white")
        style.configure("Markers.Treeview", background="#4a4a4a", foreground="white", fieldbackground="#4a4a4a")
        style.map("Markers.Treeview", background=[("selected", "#0078D7")], foreground=[("selected", "white")])
        
        # Updated button style for orange background and black text, and larger size
        style.configure("Markers.TButton", 
                        background="#F4902C", # Orange background
                        foreground="black",   # Black text
                        font=("Helvetica", 12, "bold"), # Larger font
                        padding=[10, 10, 10, 10]) # More padding for "pushable" feel
        style.map("Markers.TButton", 
                  background=[('active', '#FFB050')]) # Lighter orange on active

        # New styles for the inner treeview and buttons frame
        style.configure("Markers.Inner.Treeview",
                        background="#333333", # Darker grey
                        foreground="white",
                        fieldbackground="#333333", # Darker grey
                        bordercolor="black",
                        lightcolor="#333333", # Darker grey
                        darkcolor="#333333") # Darker grey
        style.map("Markers.Inner.Treeview",
                  background=[("selected", "#0078D7")], # Reverted to blue highlight for treeview
                  foreground=[("selected", "white")])
        
        # Configure the base TLabelFrame style and its label part
        style.configure("TLabelFrame", background="#333333", foreground="white")
        style.configure("TLabelFrame.Label", background="#333333", foreground="white")
        
        # Define a style for the inner_buttons_frame (ttk.Frame)
        style.configure("Markers.Inner.Frame", background="#333333") 
        style.layout("Markers.Inner.Frame",
                     [('TFrame.border', {'sticky': 'nswe', 'border': '1', 'children':
                       [('TFrame.padding', {'sticky': 'nswe', 'children':
                         [('TFrame.contents', {'sticky': 'nswe'})]})]})]) 

        self.config(style="Markers.TFrame") # Apply style to the main frame

        self.create_widgets()


    def create_widgets(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Creates the widgets for the Markers Display tab, including the treeview
        for zones/groups and the frame for device buttons.
        """
        debug_print("Creating MarkersDisplayTab widgets...", file=file, function=function)
        # Main frame for the split layout
        main_split_frame = ttk.Frame(self, style="Markers.TFrame") # Use ttk.Frame
        main_split_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0) # Removed outer padding as it's already on self
        main_split_frame.grid_columnconfigure(0, weight=1) # Left half
        main_split_frame.grid_columnconfigure(1, weight=1) # Right half
        main_split_frame.grid_rowconfigure(0, weight=1)

        # Left Half: Treeview for Zones and Groups
        tree_frame = ttk.LabelFrame(main_split_frame, text="Zones & Groups", padding=(5,5,5,5)) 
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
        buttons_frame = ttk.LabelFrame(main_split_frame, text="Devices", padding=(5,5,5,5)) 
        buttons_frame.grid(row=0, column=1, sticky=tk.NSEW, padx=5, pady=5)
        
        # Use a canvas with a scrollbar for buttons if there are many
        self.buttons_canvas = tk.Canvas(buttons_frame, bg="#333333", highlightbackground="#333333") # tk.Canvas as ttk.Canvas doesn't exist
        self.buttons_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        buttons_scrollbar = ttk.Scrollbar(buttons_frame, orient="vertical", command=self.buttons_canvas.yview)
        buttons_scrollbar.pack(side=tk.RIGHT, fill="y")

        self.buttons_canvas.configure(yscrollcommand=buttons_scrollbar.set)
        self.buttons_canvas.bind('<Configure>', lambda e: self.buttons_canvas.configure(scrollregion = self.buttons_canvas.bbox("all")))

        self.inner_buttons_frame = ttk.Frame(self.buttons_canvas, style="Markers.Inner.Frame") # Use ttk.Frame
        self.buttons_canvas.create_window((0, 0), window=self.inner_buttons_frame, anchor="nw")

        # Configure columns for the grid layout within inner_buttons_frame
        # Changed to 2 columns for buttons to make them wider
        self.inner_buttons_frame.grid_columnconfigure(0, weight=1)
        self.inner_buttons_frame.grid_columnconfigure(1, weight=1) 

        # Now call _populate_zone_group_tree after inner_buttons_frame is initialized
        self._populate_zone_group_tree() 

        # Initially populate with an empty list to clear any previous buttons
        self._populate_device_buttons([]) 

    def _populate_zone_group_tree(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Populates the Treeview with zones/groups and their associated devices.
        Assumes 'ZONE', 'GROUP', 'DEVICE', 'NAME', and 'FREQ' are available in self.rows.
        The tree structure will be: ZONE -> GROUP -> DEVICE -> NAME (with FREQ).
        """
        debug_print("Populating zone/group tree...", file=file, function=function)
        self.zone_group_tree.delete(*self.zone_group_tree.get_children()) # Clear existing data

        # Nested dictionary to store data: {ZONE: {GROUP: [rows]}}
        nested_grouped_data = {}

        for row in self.rows:
            zone = row.get('ZONE', 'Uncategorized Zone')
            group = row.get('GROUP', 'Uncategorized Group')
            
            if zone not in nested_grouped_data:
                nested_grouped_data[zone] = {}
            if group not in nested_grouped_data[zone]:
                nested_grouped_data[zone][group] = []
            
            nested_grouped_data[zone][group].append(row)

        for zone_name in sorted(nested_grouped_data.keys()):
            zone_id = self.zone_group_tree.insert("", "end", text=zone_name, open=True, tags=('zone',))
            
            for group_name in sorted(nested_grouped_data[zone_name].keys()):
                group_id = self.zone_group_tree.insert(zone_id, "end", text=group_name, open=True, tags=('group',))
                
                # Leaf nodes for individual devices/markers under the group
                for i, row in enumerate(nested_grouped_data[zone_name][group_name]):
                    device_name = row.get('DEVICE', 'N/A')
                    name = row.get('NAME', f"Marker {i+1}")
                    freq = row.get('FREQ', 'N/A')
                    
                    # Store the full row dictionary as a JSON string in values for easy retrieval on click
                    row_json = json.dumps(row)
                    
                    # Display frequency in MHz for consistency with the CSV output
                    display_freq = f"{float(freq):.3f} MHz" if isinstance(freq, (int, float)) else freq
                    
                    # Display DEVICE and NAME (FREQ) as leaf nodes
                    self.zone_group_tree.insert(group_id, "end", text=f"{device_name} - {name} ({display_freq})",
                                                values=(zone_name, group_name, device_name, name, freq, row_json), tags=('marker',))

        # Clear device buttons when tree is repopulated
        self._populate_device_buttons([])

    def _on_tree_select(self, event, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Handles selection events in the zone/group treeview.
        Populates the device buttons based on the selected group or individual device.
        """
        debug_print("Tree item selected...", file=file, function=function)
        selected_items = self.zone_group_tree.selection()
        if not selected_items:
            self._populate_device_buttons([])
            return

        selected_rows_data = []

        for item_id in selected_items:
            item_tags = self.zone_group_tree.item(item_id, 'tags')
            
            if 'marker' in item_tags: # It's a leaf node (actual marker)
                full_row_data_json = self.zone_group_tree.item(item_id, 'values')[-1]
                try:
                    full_row_data = json.loads(full_row_data_json)
                    selected_rows_data.append(full_row_data)
                except json.JSONDecodeError:
                    debug_print(f"Error decoding JSON for marker data: {full_row_data_json}", file=file, function=function)
            else: # It's a parent node (zone, group, or device), get all its descendant markers
                # Recursively get all children that are 'marker' tags
                def get_all_marker_children(parent_item_id):
                    children_markers = []
                    for child_id in self.zone_group_tree.get_children(parent_item_id):
                        child_tags = self.zone_group_tree.item(child_id, 'tags')
                        if 'marker' in child_tags:
                            child_row_data_json = self.zone_group_tree.item(child_id, 'values')[-1]
                            try:
                                children_markers.append(json.loads(child_row_data_json))
                            except json.JSONDecodeError:
                                debug_print(f"Error decoding JSON for child marker data: {child_row_data_json}", file=file, function=function)
                        else:
                            children_markers.extend(get_all_marker_children(child_id)) # Recurse for deeper levels
                    return children_markers
                
                selected_rows_data.extend(get_all_marker_children(item_id))
        
        self._populate_device_buttons(selected_rows_data)

    def _populate_device_buttons(self, devices_to_display, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Populates the right-hand frame with clickable buttons for each device.
        Buttons will be approximately 1/3 of the window width, orange with black text,
        and display NAME, DEVICE, and FREQ (in MHz) on three lines.
        """
        debug_print(f"Populating device buttons with {len(devices_to_display)} devices...", file=file, function=function)
        # Clear existing buttons
        for widget in self.inner_buttons_frame.winfo_children():
            widget.destroy()

        if not devices_to_display:
            ttk.Label(self.inner_buttons_frame, text="Select a zone, group, or device from the left.",
                      background="#333333", foreground="white").grid(row=0, column=0, columnspan=2, padx=5, pady=5)
            self.inner_buttons_frame.update_idletasks()
            self.buttons_canvas.config(scrollregion=self.buttons_canvas.bbox("all"))
            return

        row_idx = 0
        col_idx = 0
        for i, device_data in enumerate(devices_to_display):
            name = device_data.get('NAME', f'Unknown Marker {i+1}')
            device = device_data.get('DEVICE', 'N/A Device')
            freq_mhz = device_data.get('FREQ') 

            if freq_mhz is not None:
                try:
                    frequency_hz = float(freq_mhz) * 1_000_000 # Convert MHz to Hz for instrument
                    
                    # Format button text for three lines
                    button_text = f"{name}\n{device}\n{float(freq_mhz):.3f} MHz"
                    
                    btn = ttk.Button(self.inner_buttons_frame, text=button_text, style="Markers.TButton",
                                     command=lambda f=frequency_hz, n=name: self._on_device_button_click(f, n))
                    
                    # Use sticky="nsew" to make buttons expand within their grid cells
                    btn.grid(row=row_idx, column=col_idx, padx=5, pady=5, sticky="nsew")
                    
                    col_idx += 1
                    if col_idx >= 2: # Two columns per row
                        col_idx = 0
                        row_idx += 1
                except ValueError:
                    debug_print(f"Could not convert frequency '{freq_mhz}' to float for button. Skipping.", file=file, function=function)
            else:
                debug_print(f"Frequency not found for device '{name}'. Skipping button.", file=file, function=function)

        # Ensure columns in inner_buttons_frame expand to fill available space
        self.inner_buttons_frame.grid_columnconfigure(0, weight=1)
        self.inner_buttons_frame.grid_columnconfigure(1, weight=1)
        # Ensure rows also expand if needed, though buttons will dictate row height
        for r in range(row_idx + 1):
            self.inner_buttons_frame.grid_rowconfigure(r, weight=1)


        self.inner_buttons_frame.update_idletasks() # Ensure layout is updated before calculating scrollregion
        self.buttons_canvas.config(scrollregion=self.buttons_canvas.bbox("all"))

    def _on_device_button_click(self, freq, name, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Callback for device buttons. Sets the instrument's focus frequency and a marker.
        """
        debug_print(f"Device button clicked: {name} at {freq} Hz", file=file, function=function)
        if self.app_instance and self.app_instance.inst:
            set_focus_frequency_logic(self.app_instance, freq)
            set_marker_and_trace_modes_logic(self.app_instance, freq, name)
        else:
            debug_print("Cannot set focus frequency: Instrument not connected.", file=file, function=function)
            messagebox.showwarning("Not Connected", "Please connect to an instrument first.")

    def update_markers_data(self, headers, rows, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Updates the data displayed in the markers tab.
        """
        debug_print("Updating markers data...", file=file, function=function)
        self.headers = headers
        self.rows = rows
        self._populate_zone_group_tree() # Repopulate the treeview with new data
        self._populate_device_buttons([]) # Clear device buttons when new data loaded

