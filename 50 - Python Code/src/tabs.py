# src/tabs.py
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, ttk
import os
import csv
import xml.etree.ElementTree as ET
import sys
import inspect # Import inspect module

# Import the new report converter utility functions
from utils.report_converter_utils import convert_html_report_to_csv, generate_csv_from_shw
# Import instrument_logic for setting focus frequency
from src.instrument_logic import set_focus_frequency_logic, set_marker_and_trace_modes_logic # Ensure both are imported
from utils.instrument_control import debug_print, write_safe # Import debug_print and write_safe

class MarkersDisplayTab(ttk.Frame): # Changed from tk.Frame to ttk.Frame
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

        # REMOVED: self.pack(fill="both", expand=True, padx=10, pady=10)
        # This line was causing the tab content to "explode" over the notebook tabs.
        # The notebook itself handles the packing of its tabs.

        # Configure style for this frame's widgets
        style = ttk.Style()
        style.configure("Markers.TFrame", background="#71BE32")
        style.configure("Markers.TLabel", background="#73AF43", foreground="white")
        style.configure("Markers.Treeview.Heading", background="#3a3a3a", foreground="white")
        style.configure("Markers.Treeview", background="#4a4a4a", foreground="white", fieldbackground="#4a4a4a")
        style.map("Markers.Treeview", background=[("selected", "#0078D7")], foreground=[("selected", "white")])
        style.configure("Markers.TButton", background="#3a3a3a", foreground="white")
        style.map("Markers.TButton", background=[('active', '#6a6a6a')])
        
        # New styles for the inner treeview and buttons frame
        style.configure("Markers.Inner.Treeview",
                        background="#333333", # Darker grey
                        foreground="white",
                        fieldbackground="#333333", # Darker grey
                        bordercolor="black",
                        lightcolor="#333333", # Darker grey
                        darkcolor="#333333") # Darker grey
        style.map("Markers.Inner.Treeview",
                  background=[("selected", "#F4902C")], # New highlight color
                  foreground=[("selected", "white")])
        style.configure("Markers.Inner.LabelFrame", background="#333333", foreground="white")
        style.configure("Markers.Inner.Frame", background="#333333")

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
        tree_frame = ttk.LabelFrame(main_split_frame, text="Zones & Groups", style="Markers.Inner.LabelFrame", padding=(5,5,5,5)) # Use ttk.LabelFrame, apply padding
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

        self._populate_zone_group_tree()

        # Right Half: Buttons for Devices
        buttons_frame = ttk.LabelFrame(main_split_frame, text="Devices", style="Markers.Inner.LabelFrame", padding=(5,5,5,5)) # Use ttk.LabelFrame, apply padding
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
        self.inner_buttons_frame.grid_columnconfigure(0, weight=1)
        self.inner_buttons_frame.grid_columnconfigure(1, weight=1) # Allow for two columns

        # Initially populate with an empty list to clear any previous buttons
        self._populate_device_buttons([])

    def _populate_zone_group_tree(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Populates the Treeview with zones/groups and their associated devices.
        Assumes 'Device' and 'Frequency_MHz' are available in self.rows.
        """
        debug_print("Populating zone/group tree...", file=file, function=function)
        self.zone_group_tree.delete(*self.zone_group_tree.get_children()) # Clear existing data

        grouped_data = {}
        # Group by 'Device' or 'Zone' - prioritizing 'Device' for now
        # You might need to adjust this key based on your actual SHW/HTML structure
        group_key = 'Device' # Or 'Zone', 'Band', etc.

        for row in self.rows:
            group_name = row.get(group_key, 'Uncategorized')
            if group_name not in grouped_data:
                grouped_data[group_name] = []
            grouped_data[group_name].append(row)

        for group_name, group_rows in sorted(grouped_data.items()):
            parent_id = self.zone_group_tree.insert("", "end", text=group_name, open=True) # No values for parent
            for i, row in enumerate(group_rows):
                # Assuming 'Name' or a similar identifier for individual devices/markers
                device_name = row.get('Name', row.get('Marker', f"Item {i+1}"))
                freq_mhz = row.get('Frequency_MHz', 'N/A')
                power_dbm = row.get('Peak_Power_dBm', 'N/A') # Assuming this key exists

                # Store the full row dictionary as the last value for easy retrieval on click
                self.zone_group_tree.insert(parent_id, "end", text=f"{device_name} ({freq_mhz} MHz)",
                                            values=(device_name, freq_mhz, power_dbm, row))

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
            # Check if it's a parent (group) or a child (individual device)
            if self.zone_group_tree.parent(item_id): # It's a child
                # Retrieve the full row data stored in values (last element of values tuple)
                full_row_data = self.zone_group_tree.item(item_id, 'values')[-1]
                if isinstance(full_row_data, dict):
                    selected_rows_data.append(full_row_data)
            else: # It's a parent (group), get all its children's data
                for child_id in self.zone_group_tree.get_children(item_id):
                    full_row_data = self.zone_group_tree.item(child_id, 'values')[-1]
                    if isinstance(full_row_data, dict):
                        selected_rows_data.append(full_row_data)
        
        self._populate_device_buttons(selected_rows_data)

    def _populate_device_buttons(self, devices_to_display, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Populates the right-hand frame with clickable buttons for each device.
        """
        debug_print(f"Populating device buttons with {len(devices_to_display)} devices...", file=file, function=function)
        # Clear existing buttons
        for widget in self.inner_buttons_frame.winfo_children():
            widget.destroy()

        if not devices_to_display:
            ttk.Label(self.inner_buttons_frame, text="Select a group or device from the left.",
                      background="#333333", foreground="white").grid(row=0, column=0, columnspan=2, padx=5, pady=5)
            self.inner_buttons_frame.update_idletasks()
            self.buttons_canvas.config(scrollregion=self.buttons_canvas.bbox("all"))
            return

        row_idx = 0
        col_idx = 0
        for i, device_data in enumerate(devices_to_display):
            freq_mhz = device_data.get('Frequency_MHz')
            name = device_data.get('Name', device_data.get('Marker', 'Unknown Device')) # Use 'Marker' if 'Name' not found
            
            if freq_mhz is not None:
                try:
                    frequency_hz = float(freq_mhz) * self.app_instance.MHZ_TO_HZ
                    button_text = f"{name}: {float(freq_mhz):.3f} MHz"
                    btn = ttk.Button(self.inner_buttons_frame, text=button_text, style="Markers.TButton",
                                     command=lambda f=frequency_hz, n=name: self._on_device_button_click(f, n))
                    btn.grid(row=row_idx, column=col_idx, padx=5, pady=5, sticky="ew")
                    
                    col_idx += 1
                    if col_idx >= 2: # Two columns per row
                        col_idx = 0
                        row_idx += 1
                except ValueError:
                    debug_print(f"Could not convert frequency '{freq_mhz}' to float for button. Skipping.", file=file, function=function)
            else:
                debug_print(f"Frequency not found for device '{name}'. Skipping button.", file=file, function=function)

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


class ReportConverterTab(ttk.Frame): # Changed from tk.Frame to ttk.Frame
    """
    A Tkinter Frame that provides functionality to convert spectrum analyzer
    report files (HTML or SHW) into CSV format.
    """
    def __init__(self, master=None, app_instance=None, **kwargs):
        """
        Initializes the ReportConverterTab.

        Inputs:
            master (tk.Widget): The parent widget.
            app_instance (App): The main application instance, used for accessing
                                shared state like output directory.
            **kwargs: Arbitrary keyword arguments for Tkinter Frame.
        Process:
            1. Calls the parent `ttk.Frame` constructor.
            2. Stores `app_instance`.
            3. Configures the frame's style and layout.
            4. Creates widgets for file selection, conversion, and output.
        Outputs: None (modifies GUI state)
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Initializing ReportConverterTab...", file=current_file, function=current_function)

        super().__init__(master, **kwargs)
        self.app_instance = app_instance # Store reference to the main app instance
        
        # REMOVED: self.pack(fill="both", expand=True, padx=10, pady=10)
        # This line was causing the tab content to "explode" over the notebook tabs.
        # The notebook itself handles the packing of its tabs.

        # Configure style for this frame's widgets
        style = ttk.Style()
        style.configure("Converter.TFrame", background="#169721")
        style.configure("Converter.TLabel", background="#000000", foreground="white")
        style.configure("Converter.TEntry", fieldbackground="#4a4a4a", foreground="black", insertbackground="white")
        style.configure("Converter.TButton", background="#3a3a3a", foreground="white")
        style.map("Converter.TButton", background=[('active', '#6a6a6a')])

        self.config(style="Converter.TFrame")

        # File selection frame
        file_frame = ttk.LabelFrame(self, text="Select Report File", style="Converter.TFrame", padding="10")
        file_frame.pack(fill="x", pady=5)

        ttk.Label(file_frame, text="File Path:", style="Converter.TLabel").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.file_path_var = tk.StringVar()
        self.file_path_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50, style="Converter.TEntry")
        self.file_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(file_frame, text="Browse", command=self.browse_file, style="Converter.TButton").grid(row=0, column=2, padx=5, pady=5)
        
        file_frame.grid_columnconfigure(1, weight=1) # Make entry expand

        # Conversion button
        self.convert_button = ttk.Button(self, text="Convert to CSV", command=self.select_file, style="Converter.TButton")
        self.convert_button.pack(pady=10)

        # Output console for conversion messages
        output_frame = ttk.LabelFrame(self, text="Conversion Output", style="Converter.TFrame", padding="10")
        output_frame.pack(fill="both", expand=True, pady=5)

        self.output_text = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, bg="pink", fg="white", font=("Courier New", 10))
        self.output_text.pack(fill="both", expand=True)
        self.output_text.config(state=tk.DISABLED) # Make it read-only

    def browse_file(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """Opens a file dialog to select an HTML or SHW report file."""
        debug_print("Browsing for file...", file=file, function=function)
        file_path = filedialog.askopenfilename(
            title="Select Report File",
            filetypes=[("HTML files", "*.html"), ("SHW files", "*.shw"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            debug_print(f"Selected file: {file_path}", file=file, function=function)

    def select_file(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Converts the selected report file (HTML or SHW) to CSV format.
        """
        debug_print("Converting file to CSV...", file=file, function=function)
        input_file = self.file_path_var.get()
        if not input_file:
            messagebox.showwarning("No File Selected", "Please select an HTML or SHW file to convert.")
            debug_print("Conversion aborted: No file selected.", file=file, function=function)
            return

        self.output_text.config(state=tk.NORMAL)
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, f"Attempting to convert: {os.path.basename(input_file)}\n", "cyan")
        self.output_text.config(state=tk.DISABLED)

        error_message = None
        try:
            file_name, file_extension = os.path.splitext(os.path.basename(input_file))
            
            # Use app_instance.scan_directory_var for the output folder
            output_dir = self.app_instance.scan_directory_var.get()
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                debug_print(f"Created output directory: {output_dir}", file=file, function=function)

            if file_extension.lower() == '.html':
                output_csv_file = os.path.join(output_dir, f"{file_name}.csv")
                headers, rows = convert_html_report_to_csv(input_file, output_csv_file)
            elif file_extension.lower() == '.shw':
                output_csv_file = os.path.join(output_dir, f"{file_name}.csv")
                # Corrected call: generate_csv_from_shw should only take input_file
                headers, rows = generate_csv_from_shw(input_file) 
            else:
                messagebox.showerror("Unsupported File Type", "Only HTML (.html) and SHW (.shw) files are supported for conversion.")
                debug_print(f"Unsupported file type: {file_extension}", file=file, function=function)
                return

            if rows: # Only write if there's data to write
                with open(output_csv_file, 'w', newline='', encoding='utf-8') as csvfile:
                    csv_writer = csv.DictWriter(csvfile, fieldnames=headers)
                    csv_writer.writeheader()
                    csv_writer.writerows(rows)
                messagebox.showinfo("Success", f"Successfully converted '{file_name}' to '{os.path.basename(output_csv_file)}'")
                
                # Call the method on the main App instance to add the new tab
                # This call is correct based on the App class structure.
                self.app_instance.add_markers_tab(headers, rows)
            else:
                messagebox.showwarning("No Data Extracted", f"No relevant data could be extracted from '{file_name}'. CSV file was not created.")

        except FileNotFoundError as e:
            error_message = f"File not found: {e}"
            messagebox.showerror("File Error", error_message)
        except ET.ParseError as e:
            error_message = f"Error parsing XML (SHW) file: {e}"
            messagebox.showerror("Parsing Error", error_message)
        except Exception as e:
            error_message = f"An unexpected error occurred during conversion: {e}"
            messagebox.showerror("Conversion Error", error_message)
        
        if error_message:
            print(f"❌ Conversion failed for {file_name}: {error_message}")
            debug_print(f"Conversion failed for {file_name}: {error_message}", file=file, function=function)
