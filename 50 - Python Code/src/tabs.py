# src/tabs.py
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, ttk
import os
import csv
import xml.etree.ElementTree as ET
import subprocess # For BeautifulSoup installation check if kept here
import sys

# BeautifulSoup Installation Check (can be moved to a separate install_deps.py if preferred)
try:
    from bs4 import BeautifulSoup
except ImportError:
    print("BeautifulSoup4 not found. Attempting to install...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "beautifulsoup4"])
        from bs4 import BeautifulSoup
        print("BeautifulSoup4 installed successfully.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Installation Error", f"Error installing BeautifulSoup4: {e}\\nPlease install it manually by running: pip install beautifulsoup4")
        sys.exit(1)
    except Exception as e:
        messagebox.showerror("Installation Error", f"An unexpected error occurred during BeautifulSoup4 installation: {e}")
        sys.exit(1)

# Import the new report converter utility functions
from utils.report_converter_utils import convert_html_report_to_csv, generate_csv_from_shw
# Import instrument_logic for setting focus frequency
from src.instrument_logic import set_focus_frequency_logic

class MarkersDisplayTab(tk.Frame):
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
            1. Calls the parent `tk.Frame` constructor.
            2. Configures the background color.
            3. Stores the `headers`, `rows` data, and `app_instance`.
            4. Calls `create_widgets()` to build the GUI elements.
        Outputs: None
        """
        super().__init__(master, **kwargs)
        self.configure(bg="black")
        self.headers = headers if headers is not None else []
        self.rows = rows if rows is not None else [] # Store full rows data
        self.app_instance = app_instance # Store reference to the main app instance
        self.create_widgets()

    def create_widgets(self):
        """
        Creates the widgets for the Markers Display tab, including the treeview
        for zones/groups and the frame for device buttons.

        Inputs: None
        Process:
            1. Creates a main split frame for layout.
            2. Configures grid weights for the main split frame.
            3. Creates a `tk.LabelFrame` for "Zones & Groups" to hold the treeview.
            4. Initializes `self.zone_group_tree` (ttk.Treeview) and its scrollbar.
            5. Binds the `<<TreeviewSelect>>` event to `self._on_tree_select` to handle selections.
            6. Calls `_populate_zone_group_tree()` to initially populate the treeview.
            7. Creates a `tk.LabelFrame` for "Devices" to hold the device buttons.
            8. Initializes a `tk.Canvas` and `ttk.Scrollbar` for the device buttons
               to allow scrolling if many buttons are present.
            9. Creates `self.inner_buttons_frame` inside the canvas to hold the actual buttons.
            10. Calls `_populate_device_buttons([])` to ensure the device button area is initially empty.
        Outputs: None
        """
        # Main frame for the split layout
        main_split_frame = tk.Frame(self, bg="black")
        main_split_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main_split_frame.grid_columnconfigure(0, weight=1) # Left half
        main_split_frame.grid_columnconfigure(1, weight=1) # Right half
        main_split_frame.grid_rowconfigure(0, weight=1)

        # Left Half: Treeview for Zones and Groups
        tree_frame = tk.LabelFrame(main_split_frame, text="Zones & Groups", bg="black", fg="white", padx=5, pady=5)
        tree_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=5, pady=5)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.zone_group_tree = ttk.Treeview(tree_frame, show="tree") # Only show tree, not headings
        self.zone_group_tree.pack(fill=tk.BOTH, expand=True)

        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.zone_group_tree.yview)
        tree_scrollbar.pack(side=tk.RIGHT, fill="y")
        self.zone_group_tree.configure(yscrollcommand=tree_scrollbar.set)
        
        # Bind selection event to update device buttons
        self.zone_group_tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        self._populate_zone_group_tree()

        # Right Half: Buttons for Devices
        buttons_frame = tk.LabelFrame(main_split_frame, text="Devices", bg="black", fg="white", padx=5, pady=5)
        buttons_frame.grid(row=0, column=1, sticky=tk.NSEW, padx=5, pady=5)
        
        # Use a canvas with a scrollbar for buttons if there are many
        self.buttons_canvas = tk.Canvas(buttons_frame, bg="black", highlightbackground="black") # Store canvas as instance variable
        self.buttons_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        buttons_scrollbar = ttk.Scrollbar(buttons_frame, orient="vertical", command=self.buttons_canvas.yview)
        buttons_scrollbar.pack(side=tk.RIGHT, fill="y")

        self.buttons_canvas.configure(yscrollcommand=buttons_scrollbar.set)
        self.buttons_canvas.bind('<Configure>', lambda e: self.buttons_canvas.configure(scrollregion = self.buttons_canvas.bbox("all")))

        self.inner_buttons_frame = tk.Frame(self.buttons_canvas, bg="black")
        self.buttons_canvas.create_window((0, 0), window=self.inner_buttons_frame, anchor="nw")

        # Configure columns for the grid layout within inner_buttons_frame
        self.inner_buttons_frame.grid_columnconfigure(0, weight=1)
        self.inner_buttons_frame.grid_columnconfigure(1, weight=1) # Allow for two columns

        # Initially populate with an empty list to clear any previous buttons
        # This call is moved here to ensure self.inner_buttons_frame is already created.
        self._populate_device_buttons([])

    def _populate_zone_group_tree(self):
        """
        Populates the `self.zone_group_tree` (Treeview) with zones and groups
        extracted from the `self.rows` data. Handles cases where groups are blank.

        Inputs: None
        Process:
            1. Clears all existing items and selections from the treeview.
            2. Organizes the `self.rows` data into a nested dictionary `zones_data`
               structured as `{zone_name: {group_name: [list_of_rows]}}`.
               An empty string is used for `group_name` if the CSV 'GROUP' field is blank.
            3. Iterates through sorted zone names:
               - Inserts a parent node for each `zone_name`. The `values` for a zone node
                 are `("zone", zone_name, None)`.
               - Determines if the current zone has any *named* groups.
               - If named groups exist, iterates through sorted group names:
                 - Inserts a child node for each *non-empty* `group_name` under its zone.
                   The `values` for a group node are `("group", zone_name, group_name)`.
                   Blank groups are not given their own sub-node.
            4. Calls `_populate_device_buttons([])` to ensure the device display is empty initially.
        Outputs: None
        """
        self.zone_group_tree.delete(*self.zone_group_tree.get_children()) # Clear existing items
        self.zone_group_tree.selection_set([]) # Clear selection on repopulate
        
        zones_data = {} # {zone_name: {group_name: [list_of_rows]}}

        for row in self.rows:
            zone_name = row.get("ZONE", "Unknown Zone").strip()
            group_name = row.get("GROUP", "").strip() # Use empty string for no group

            if zone_name not in zones_data:
                zones_data[zone_name] = {}
            if group_name not in zones_data[zone_name]:
                zones_data[zone_name][group_name] = []
            zones_data[zone_name][group_name].append(row)

        for zone_name in sorted(zones_data.keys()):
            # Insert zone node. Value tuple: (node_type, zone_name, group_name_if_group)
            # For a zone node, group_name_if_group is None
            zone_id = self.zone_group_tree.insert("", "end", text=zone_name, open=True, values=("zone", zone_name, None))
            
            # Determine if this zone has any named groups
            has_named_groups = any(g for g in zones_data[zone_name].keys() if g)

            if has_named_groups:
                # If there are named groups, insert them as children
                for group_name in sorted(zones_data[zone_name].keys()):
                    if group_name: # Only insert if group name is not empty
                        group_id = self.zone_group_tree.insert(zone_id, "end", text=group_name, values=("group", zone_name, group_name))
            # If no named groups, the devices with blank groups are implicitly under the zone node itself.
            # No need to create a "blank group" child node as per user request.

        # Initially clear device buttons
        # This call is now handled by create_widgets after inner_buttons_frame is created.
        # self._populate_device_buttons([])

    def _on_tree_select(self, event):
        """
        Event handler for when an item in the `self.zone_group_tree` (Treeview) is selected.
        This function filters the device data based on the selected zone or group
        and updates the displayed device buttons accordingly.

        Inputs:
            event (tk.Event): The Tkinter event object (not directly used, but part of binding).
        Process:
            1. Retrieves the currently selected item(s) from the treeview.
            2. If no item is selected, calls `_populate_device_buttons([])` to clear buttons.
            3. If an item is selected:
               - Extracts the `node_type` ("zone" or "group"), `selected_zone`, and `selected_group`
                 from the selected item's `values` tuple.
               - Initializes an empty list `filtered_devices`.
               - If `node_type` is "zone":
                 - Iterates through all `self.rows` and adds any row where `ZONE` matches `selected_zone`
                   to `filtered_devices`. This displays all devices under the selected zone.
               - If `node_type` is "group":
                 - Iterates through all `self.rows` and adds any row where `ZONE` matches `selected_zone`
                   AND `GROUP` matches `selected_group` to `filtered_devices`. This displays devices
                   only for the specific selected group.
               - Calls `self._populate_device_buttons(filtered_devices)` to update the display.
        Outputs: None
        """
        selected_items = self.zone_group_tree.selection()
        if not selected_items:
            self._populate_device_buttons([]) # Clear buttons if nothing selected
            return

        selected_item_id = selected_items[0]
        item_values = self.zone_group_tree.item(selected_item_id, 'values')
        
        if not item_values or len(item_values) < 2: # Ensure values tuple has at least type and zone
            self._populate_device_buttons([])
            return

        node_type = item_values[0]
        selected_zone = item_values[1]
        selected_group = item_values[2] if len(item_values) > 2 else None # Will be None for zone nodes

        filtered_devices = []

        if node_type == "zone":
            # If a zone node is selected, display all devices belonging to that zone,
            # regardless of whether they have a named group or a blank group.
            for row in self.rows:
                if row.get("ZONE", "").strip() == selected_zone:
                    filtered_devices.append(row)
        elif node_type == "group":
            # If a group node is selected, display only devices belonging to that specific zone AND group.
            for row in self.rows:
                if row.get("ZONE", "").strip() == selected_zone and row.get("GROUP", "").strip() == selected_group:
                    filtered_devices.append(row)
        
        self._populate_device_buttons(filtered_devices)

    def _populate_device_buttons(self, devices_to_display):
        """
        Populates the `self.inner_buttons_frame` with buttons for each device
        in the `devices_to_display` list. Clears any existing buttons first.

        Inputs:
            devices_to_display (list): A list of dictionaries, where each dictionary
                                       represents a device row to be displayed as a button.
        Process:
            1. Destroys all existing widgets (buttons) within `self.inner_buttons_frame`.
            2. Configures `self.inner_buttons_frame` columns to expand.
            3. Initializes `row_num` and `col_num` for grid placement.
            4. If `devices_to_display` is empty, displays a message.
            5. Iterates through each `row_data` in `devices_to_display`.
            6. Extracts "DEVICE", "NAME", and "FREQ" from the `row_data`.
            7. Constructs the button text.
            8. Creates a `tk.Button` for each device, using `grid` to place it.
            9. Increments `col_num` and `row_num` to move to the next grid cell.
            10. Updates the scroll region of the `buttons_canvas` to ensure all buttons are scrollable.
        Outputs: None
        """
        # Clear existing buttons
        for widget in self.inner_buttons_frame.winfo_children():
            widget.destroy()

        # Configure columns for the grid layout within inner_buttons_frame
        # Already done in create_widgets, but re-doing here for safety if called independently
        self.inner_buttons_frame.grid_columnconfigure(0, weight=1)
        self.inner_buttons_frame.grid_columnconfigure(1, weight=1) # Allow for two columns
        
        num_columns = 2 # Define the number of columns for the button grid
        row_num = 0
        col_num = 0

        if not devices_to_display:
            # Display a message if no devices are to be shown
            no_devices_label = tk.Label(self.inner_buttons_frame, text="Select a Zone or Group to view devices.",
                                        bg="black", fg="grey", font=('Arial', 10, 'italic'), wraplength=250)
            no_devices_label.grid(row=0, column=0, columnspan=num_columns, pady=20, sticky="nsew")
            # Access the main App instance to call _update_console_line
            # self.master is the notebook, self.master.master is the main_frame, self.master.master.master is the App instance
            self.app_instance._update_console_line("No devices to display. Select a Zone or Group in the tree.", overwrite=True)
        else:
            for row_data in devices_to_display:
                device = row_data.get("DEVICE", "N/A")
                name = row_data.get("NAME", "N/A")
                freq = row_data.get("FREQ", "N/A")

                button_text = f"{device}\n{name}\n{freq}"
                
                # Bind the button click to the new _on_device_button_click method
                btn = tk.Button(self.inner_buttons_frame, text=button_text, 
                                font=('Arial', 9), bg='darkblue', fg='white',
                                activebackground='blue', activeforeground='white',
                                relief=tk.RAISED, bd=2, padx=5, pady=3, wraplength=150,
                                command=lambda f=freq, n=name: self._on_device_button_click(f, n))
                
                # Use grid to place buttons
                btn.grid(row=row_num, column=col_num, padx=2, pady=2, sticky="nsew")
                
                col_num += 1
                if col_num >= num_columns:
                    col_num = 0
                    row_num += 1
        
        # Update the scroll region after adding/removing buttons
        self.inner_buttons_frame.update_idletasks()
        # Corrected: Use self.buttons_canvas to update its scrollregion
        self.buttons_canvas.config(scrollregion=self.buttons_canvas.bbox("all"))

    def _on_device_button_click(self, freq, name):
        """
        Handles the event when a device button is clicked. It pauses any ongoing scan
        and then attempts to set the instrument's center frequency and span based on
        the clicked device's frequency and the default focus width.

        Inputs:
            freq (str/float): The frequency of the clicked device (from CSV).
            name (str): The name of the clicked device (from CSV).
        Process:
            1. Checks if an `app_instance` is available. If not, prints an error.
            2. If a scan is currently running (`self.app_instance.scanning` is True),
               it calls `self.app_instance.toggle_pause_scan()` to pause it.
            3. Retrieves the `default_focus_width` from `self.app_instance.default_focus_width_var`.
            4. Converts the `freq` to a float (assuming MHz from CSV, convert to Hz).
            5. Calls `set_focus_frequency_logic` from `src.instrument_logic` to send
               commands to the connected instrument.
            6. Prints success or failure messages to the console.
        Outputs: None (modifies instrument state, updates GUI console)
        """
        if not self.app_instance:
            print("Error: App instance not available in MarkersDisplayTab.")
            return

        if self.app_instance.inst:
            # Pause any ongoing scan
            if self.app_instance.scanning:
                self.app_instance.toggle_pause_scan()
                print("Scan paused to set focus frequency.")
            
            try:
                center_frequency_hz = float(freq) * 1_000_000 # Assuming freq from CSV is in MHz
                span_hz = self.app_instance.default_focus_width_var.get()

                print(f"Attempting to set instrument focus for '{name}' at {center_frequency_hz / 1_000_000:.3f} MHz with span {span_hz} Hz...")
                
                success = set_focus_frequency_logic(
                    self.app_instance,
                    center_frequency_hz,
                    name,
                    span_hz
                )
                if success:
                    print(f"✅ Instrument focused on '{name}' at {center_frequency_hz / 1_000_000:.3f} MHz with span {span_hz} Hz.")
                else:
                    print(f"❌ Failed to set instrument focus for '{name}'. See console for details.")
            except ValueError:
                messagebox.showerror("Input Error", f"Invalid frequency value for device '{name}': {freq}")
                print(f"❌ Invalid frequency value for device '{name}': {freq}")
            except Exception as e:
                messagebox.showerror("Instrument Error", f"An error occurred while setting instrument focus: {e}")
                print(f"❌ Error setting instrument focus: {e}")
        else:
            messagebox.showwarning("Not Connected", "Please connect to an instrument first to set focus frequency.")
            print("🚫 Cannot set focus frequency: Instrument not connected.")


class ReportConverterTab(tk.Frame):
    """
    A Tkinter Frame that encapsulates the functionality of the Report Converter.
    This includes converting HTML and SHW files to CSV format.
    """
    def __init__(self, master=None, app_instance=None, **kwargs):
        """
        Initializes the ReportConverterTab.

        Inputs:
            master (tk.Widget): The parent widget.
            app_instance (App): The main application instance, used for accessing
                                shared state like the default focus width variable.
            **kwargs: Arbitrary keyword arguments for Tkinter Frame.
        Process:
            1. Calls the parent `tk.Frame` constructor.
            2. Configures the background color.
            3. Stores the `app_instance`.
            4. Calls `create_widgets()` to build the GUI elements.
        Outputs: None
        """
        super().__init__(master, **kwargs)
        self.configure(bg="black")
        self.app_instance = app_instance # Store reference to the main app instance
        self.create_widgets()

    def create_widgets(self):
        """
        Creates the widgets for the Report Converter tab, including the file selection
        button and the new "Default width of the focus" slider.

        Inputs: None
        Process:
            1. Creates a label for instructions.
            2. Creates a button to open the file dialog, linked to `self.select_file`.
            3. Creates a new label and slider for "Default width of the focus":
               - The slider's values are 5000, 10000, 25000, 50000, 100000 Hz.
               - The slider's variable is linked to `self.app_instance.default_focus_width_var`.
        Outputs: None
        """
        # Create a label for instructions
        instruction_label = tk.Label(self, text="Click the button below to select a report file (.html or .shw)", 
                                     wraplength=350, bg="black", fg="white")
        instruction_label.pack(pady=10)

        # Create a button to open the file dialog
        select_button = tk.Button(self, text="Select Report File", command=self.select_file,
                                  font=('Arial', 12, 'bold'), bg='#4CAF50', fg='white',
                                  activebackground='#45a049', activeforeground='white',
                                  relief=tk.RAISED, bd=3, padx=10, pady=5)
        select_button.pack(pady=20)

        # New: Default width of the focus slider
        if self.app_instance and hasattr(self.app_instance, 'default_focus_width_var'):
            focus_width_frame = tk.LabelFrame(self, text="Focus Scan Width", padx=10, pady=5, bg="black", fg="white")
            focus_width_frame.pack(pady=10, padx=10, fill=tk.X)

            tk.Label(focus_width_frame, text="Default Width (Hz):", bg="black", fg="white").pack(side=tk.LEFT, padx=5)
            
            # Define the values for the slider
            focus_width_values = [5000, 10000, 25000, 50000, 100000]
            # Create a mapping from value to index for the slider
            focus_width_val_to_idx = {val: i for i, val in enumerate(focus_width_values)}
            
            # Create an IntVar for the slider's index
            self.focus_width_slider_index_var = tk.IntVar(self, value=focus_width_val_to_idx.get(int(self.app_instance.default_focus_width_var.get()), 0))

            def update_focus_width_from_slider(*args):
                try:
                    idx = self.focus_width_slider_index_var.get()
                    if 0 <= idx < len(focus_width_values):
                        self.app_instance.default_focus_width_var.set(float(focus_width_values[idx]))
                except Exception as e:
                    print(f"Error updating focus width from slider: {e}")

            def update_focus_width_slider_from_var(*args):
                try:
                    val = float(self.app_instance.default_focus_width_var.get())
                    if val in focus_width_val_to_idx:
                        self.focus_width_slider_index_var.set(focus_width_val_to_idx[val])
                    else:
                        closest_val = min(focus_width_values, key=lambda x: abs(x - val))
                        self.focus_width_slider_index_var.set(focus_width_val_to_idx[closest_val])
                except ValueError:
                    pass

            self.focus_width_slider_index_var.trace_add("write", update_focus_width_from_slider)
            self.app_instance.default_focus_width_var.trace_add("write", update_focus_width_slider_from_var)

            self.focus_width_slider = tk.Scale(focus_width_frame,
                                               variable=self.focus_width_slider_index_var,
                                               from_=0, to=len(focus_width_values) - 1,
                                               orient=tk.HORIZONTAL, showvalue=0, # Don't show numeric value on slider
                                               resolution=1,
                                               bg="black", fg="white", troughcolor="grey", highlightbackground="black",
                                               length=200)
            self.focus_width_slider.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

            # Display the current value next to the slider
            self.focus_width_display_label = tk.Label(focus_width_frame, textvariable=self.app_instance.default_focus_width_var, bg="black", fg="white")
            self.focus_width_display_label.pack(side=tk.LEFT, padx=5)

            # Manually update the slider's initial position based on the variable's initial value
            update_focus_width_slider_from_var()


    def select_file(self):
        """
        Opens a file dialog for the user to select an HTML or SHW file,
        then processes it accordingly.

        Inputs: None
        Process:
            1. Opens a file dialog allowing selection of .html or .shw files.
            2. If no file is selected, returns.
            3. Extracts file name and extension.
            4. Determines the output directory using the main App instance's `output_folder_var`.
            5. Ensures the output directory exists.
            6. Sets the `output_csv_file` path to "MARKERS.CSV" within the output directory.
            7. Initializes `headers`, `rows`, `conversion_successful`, and `error_message` variables.
            8. Attempts to convert the selected file:
               - If .html, reads content and calls `convert_html_report_to_csv`.
               - If .shw, calls `generate_csv_from_shw`.
               - If invalid type, shows a warning and returns.
            9. If conversion is successful and rows are extracted:
               - Writes the extracted `headers` and `rows` to the `output_csv_file`.
               - Shows a success messagebox.
               - Calls `self.master.master.master.add_markers_tab(headers, rows)` to add/update the Markers Display tab.
            10. If no data is extracted, shows a warning.
            11. Includes comprehensive error handling for `FileNotFoundError`, `ET.ParseError`, and general `Exception`.
            12. If an error occurs, prints an error message to the console and shows an error messagebox.
        Outputs: None (creates CSV file, updates GUI tabs, shows messageboxes)
        """
        file_path = filedialog.askopenfilename(
            title="Select an IAS HTML Report or a SHURE Wireless Workbench Show File",
            filetypes=[("Report Files", "*.html *.shw"), ("HTML files", "*.html"), ("SHW files", "*.shw")]
        )

        if not file_path:
            return # User cancelled file selection

        file_name = os.path.basename(file_path)
        base_name, extension = os.path.splitext(file_name)
        
        # Get the output directory from the main App instance's output_folder_var
        # self.master is the notebook, self.master.master is the main_frame, self.master.master.master is the App instance
        output_dir = self.app_instance.output_folder_var.get()
        if not os.path.exists(output_dir):
            os.makedirs(output_dir) # Ensure the output directory exists
        
        output_csv_file = os.path.join(output_dir, "MARKERS.CSV")

        headers = []
        rows = []
        conversion_successful = False
        error_message = ""

        try:
            if extension.lower() == '.html':
                with open(file_path, 'r', encoding='utf-8') as f:
                    html_report_content = f.read()
                headers, rows = convert_html_report_to_csv(html_report_content)
                conversion_successful = True
            
            elif extension.lower() == '.shw':
                headers, rows = generate_csv_from_shw(file_path)
                conversion_successful = True
            
            else:
                messagebox.showwarning("Invalid File Type", "Please select a .html or .shw file.")
                return # Exit if file type is invalid

            if conversion_successful:
                if rows: # Only write if there's data to write
                    with open(output_csv_file, 'w', newline='', encoding='utf-8') as csvfile:
                        csv_writer = csv.DictWriter(csvfile, fieldnames=headers)
                        csv_writer.writeheader()
                        csv_writer.writerows(rows)
                    messagebox.showinfo("Success", f"Successfully converted '{file_name}' to '{os.path.basename(output_csv_file)}'")
                    
                    # Call the method on the main App instance to add the new tab
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
