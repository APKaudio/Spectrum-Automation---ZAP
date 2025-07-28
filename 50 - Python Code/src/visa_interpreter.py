# src/visa_interpreter.py
import tkinter as tk
from tkinter import ttk, filedialog
import csv
import os
import inspect # Import inspect module for debug_print

from utils.instrument_control import debug_print # Import debug_print

class VisaInterpreterTab(ttk.Frame):
    """
    A Tkinter Frame that provides a user-editable cell editor for VISA commands.
    It displays command types and the commands themselves, allowing users to
    modify, add, or remove entries.
    """
    def __init__(self, master=None, app_instance=None, console_print_func=None, **kwargs):
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self.console_print_func = console_print_func if console_print_func else print

        self.data_file = "visa_commands.csv" # File to save/load user-edited commands

        self._create_widgets()
        self._load_data() # Load existing data when the tab is initialized

        # Bind double-click event for cell editing
        self.tree.bind("<Double-1>", self._on_double_click_edit)
        # Bind <Return> key to save and close editor
        self.tree.bind("<Return>", self._on_enter_edit)
        # Bind <Escape> key to close editor without saving
        self.tree.bind("<Escape>", self._on_escape_edit)

        debug_print("VisaInterpreterTab initialized.", file=__file__, function=inspect.currentframe().f_code.co_name, console_print_func=self.console_print_func)

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the VISA Interpreter tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Creating VisaInterpreterTab widgets...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1) # Treeview takes most space
        self.grid_rowconfigure(1, weight=0) # Buttons row

        # Treeview for displaying and editing commands
        # Changed columns to include "Command Type" and "Variable"
        columns = ("Command Type", "VISA Command", "Variable")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", style='Treeview')
        self.tree.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # Configure column headings
        self.tree.heading("Command Type", text="Command Type", anchor=tk.W)
        self.tree.heading("VISA Command", text="VISA Command", anchor=tk.W)
        self.tree.heading("Variable", text="Variable", anchor=tk.W) # New heading for Variable

        # Configure column widths (adjust as needed)
        self.tree.column("Command Type", width=120, minwidth=100, stretch=False) # Fixed width for type
        self.tree.column("VISA Command", width=350, minwidth=250, stretch=True)
        self.tree.column("Variable", width=100, minwidth=80, stretch=True) # New column width for Variable

        # Scrollbars
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=vsb.set)

        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=hsb.set)

        # Button Frame
        button_frame = ttk.Frame(self, style='Dark.TFrame')
        button_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        button_frame.grid_columnconfigure(2, weight=1)

        ttk.Button(button_frame, text="Add Row", command=self._add_row, style='Blue.TButton').grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ttk.Button(button_frame, text="Delete Selected Row", command=self._delete_row, style='Red.TButton').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(button_frame, text="Save Commands", command=self._save_data, style='Green.TButton').grid(row=0, column=2, padx=5, pady=5, sticky="ew")

        self.editor = None # To hold the Entry widget for editing

        debug_print("VisaInterpreterTab widgets created.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _load_data(self):
        """
        Loads VISA commands from a CSV file or uses default commands if the file doesn't exist.
        Handles both old 2-column and new 3-column CSV formats.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Loading data from {self.data_file}...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self.tree.delete(*self.tree.get_children()) # Clear existing data

        commands = []
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    header = next(reader) # Skip header row
                    for row in reader:
                        if len(row) == 3: # New format: Command Type, VISA Command, Variable
                            commands.append(row)
                        elif len(row) == 2: # Old format: Command Type Description, VISA Command
                            # Infer Command Type and add empty Variable column
                            command_text = row[1]
                            command_type_prefix = "GET" if command_text.strip().endswith("?") else ("DO" if command_text.strip() in ["*RST", ":SYSTem:DISPlay:UPDate"] else "SET")
                            description = row[0].replace(" - GET", "").replace(" - SET", "").replace(" - DO", "").strip()
                            commands.append([f"{description} - {command_type_prefix}", command_text, ""])
                        else:
                            self.console_print_func(f"⚠️ Skipping malformed row in {self.data_file}: {row}")
                            debug_print(f"Skipping malformed row: {row}", file=current_file, function=current_function, console_print_func=self.console_print_func)

                self.console_print_func(f"✅ Loaded {len(commands)} commands from {self.data_file}.")
                debug_print(f"Loaded {len(commands)} commands from {self.data_file}.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            except Exception as e:
                self.console_print_func(f"❌ Error loading commands from {self.data_file}: {e}. Loading defaults.")
                debug_print(f"Error loading {self.data_file}: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
                commands = self._get_default_commands()
        else:
            self.console_print_func(f"ℹ️ {self.data_file} not found. Loading default commands.")
            debug_print(f"{self.data_file} not found. Loading default commands.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            commands = self._get_default_commands()

        for cmd_type, command, variable in commands:
            self.tree.insert("", "end", values=(cmd_type, command, variable))
        debug_print(f"Displayed {len(commands)} commands in Treeview.", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _save_data(self):
        """
        Saves the current commands from the Treeview to the CSV file.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Saving data to {self.data_file}...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        data_to_save = []
        for item_id in self.tree.get_children():
            values = self.tree.item(item_id, 'values')
            data_to_save.append(list(values)) # Convert tuple to list for consistency

        try:
            with open(self.data_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                # Updated header to include 'Variable'
                writer.writerow(["Command Type", "VISA Command", "Variable"])
                writer.writerows(data_to_save)
            self.console_print_func(f"✅ Saved {len(data_to_save)} commands to {self.data_file}.")
            debug_print(f"Saved {len(data_to_save)} commands to {self.data_file}.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        except Exception as e:
            self.console_print_func(f"❌ Error saving commands to {self.data_file}: {e}")
            debug_print(f"Error saving {self.data_file}: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _add_row(self):
        """
        Adds a new empty row to the Treeview with default values.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        # New rows will have empty strings for all three columns
        self.tree.insert("", "end", values=("Set", "", "")) # Default to "Set" for new rows
        self.console_print_func("✅ Added a new empty row.")
        debug_print("Added new row.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _delete_row(self):
        """
        Deletes the selected row(s) from the Treeview.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        selected_items = self.tree.selection()
        if not selected_items:
            self.console_print_func("⚠️ No row selected to delete.")
            debug_print("No row selected for deletion.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        for item in selected_items:
            self.tree.delete(item)
        self.console_print_func(f"✅ Deleted {len(selected_items)} selected row(s).")
        debug_print(f"Deleted {len(selected_items)} row(s).", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _on_double_click_edit(self, event):
        """
        Handles double-click event to enable in-cell editing.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Double-click detected for editing.", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if self.editor: # If an editor is already open, destroy it first
            self._save_and_destroy_editor()

        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)

        if not item or not column:
            debug_print("No item or column identified for editing.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        # Get column index (e.g., #1 for first column, #2 for second, #3 for third)
        col_idx = int(column[1:]) - 1 
        
        # Get current value
        current_values = self.tree.item(item, 'values')
        current_value = current_values[col_idx]

        # Get bounding box of the cell
        x, y, width, height = self.tree.bbox(item, column)

        # Create an Entry widget for editing
        self.editor = ttk.Entry(self.tree, style='TEntry')
        self.editor.place(x=x, y=y, width=width, height=height)
        self.editor.insert(0, current_value)
        self.editor.focus_set()

        # Store item and column info for saving
        self.editor.item = item
        self.editor.column = col_idx
        debug_print(f"Editor created for item {item}, column {col_idx} with value '{current_value}'.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _save_and_destroy_editor(self):
        """
        Saves the edited content from the Entry widget back to the Treeview
        and destroys the Entry widget. Automatically updates "Command Type"
        if the "VISA Command" column is edited.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        if self.editor:
            new_value = self.editor.get()
            item = self.editor.item
            col_idx = self.editor.column

            # Get current values as a list, update the specific column, then set back
            current_values = list(self.tree.item(item, 'values'))
            current_values[col_idx] = new_value

            # If the edited column is the VISA Command column (index 1), update the Command Type (index 0)
            if col_idx == 1: # This is the "VISA Command" column
                visa_command = new_value
                command_type_prefix = "GET" if visa_command.strip().endswith("?") else ("DO" if visa_command.strip() in ["*RST", ":SYSTem:DISPlay:UPDate"] else "SET")
                # Preserve the original description part if it exists
                original_description = current_values[0].split(" - ")[0] if " - " in current_values[0] else ""
                current_values[0] = f"{original_description} - {command_type_prefix}".strip()

            self.tree.item(item, values=current_values)
            self.editor.destroy()
            self.editor = None
            self.console_print_func(f"✅ Cell updated: {current_values}")
            debug_print(f"Editor destroyed. Cell updated to: '{new_value}'", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _on_enter_edit(self, event):
        """Handles Enter key press to save and destroy editor."""
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Enter key pressed in editor.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        self._save_and_destroy_editor()

    def _on_escape_edit(self, event):
        """Handles Escape key press to destroy editor without saving."""
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Escape key pressed in editor.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        if self.editor:
            self.editor.destroy()
            self.editor = None
            self.console_print_func("ℹ️ Cell edit cancelled.")
            debug_print("Editor destroyed without saving.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _get_default_commands(self):
        """
        Returns a list of default VISA commands with their automatically determined
        type ("GET", "SET", or "DO") and an empty variable column.
        """
        default_categorized_commands = [
            # System/Identification
            ("System/Identification", "*IDN?"),
            ("System/Identification", "*RST"),
            ("System/Identification", ":SYSTem:ERRor?"),
            ("System/Identification", ":SYSTem:DISPlay:UPDate"),

            # Frequency/Span/Sweep
            ("Frequency/Span/Sweep", ":SENSe:FREQuency:CENTer"),
            ("Frequency/Span/Sweep", ":SENSe:FREQuency:CENTer?"),
            ("Frequency/Span/Sweep", ":SENSe:FREQuency:SPAN"),
            ("Frequency/Span/Sweep", ":SENSe:FREQuency:SPAN?"),
            ("Frequency/Span/Sweep", ":FREQuency:STARt?"),
            ("Frequency/Span/Sweep", ":FREQuency:STOP?"),
            ("Frequency/Span/Sweep", ":SENSe:SWEep:POINts?"),
            ("Frequency/Span/Sweep", ":SENSe:SWEep:TIME:AUTO ON"),
            ("Frequency/Span/Sweep", ":SENSe:X:SPACing LINear"),
            ("Frequency/Span/Sweep", ":FREQuency:OFFSet?"),
            ("Frequency/Span/Sweep", ":INPut:RFSense:FREQuency:SHIFt?"),
            ("Frequency/Span/Sweep", ":INPut:RFSense:FREQuency:SHIFt"),


            # Bandwidth (RBW/VBW)
            ("Bandwidth (RBW/VBW)", ":SENSe:BANDwidth:RESolution"),
            ("Bandwidth (RBW/VBW)", ":SENSe:BANDwidth:RESolution?"),
            ("Bandwidth (RBW/VBW)", ":SENSe:BANDwidth:VIDeo"),
            ("Bandwidth (RBW/VBW)", ":SENSe:BANDwidth:VIDeo?"),
            ("Bandwidth (RBW/VBW)", ":SENSe:BANDwidth:RESolution:AUTO ON"),
            ("Bandwidth (RBW/VBW)", ":SENSe:BANDwidth:VIDeo:AUTO ON"),

            # Amplitude/Reference Level/Attenuation/Gain
            ("Amplitude/Level/Gain", ":DISPlay:WINDow:TRACe:Y:RLEVel"),
            ("Amplitude/Level/Gain", ":DISPlay:WINDow:TRACe:Y:RLEVel?"),
            ("Amplitude/Level/Gain", ":INPut:ATTenuation:AUTO"),
            ("Amplitude/Level/Gain", ":INPut:ATTenuation:AUTO?"),
            ("Amplitude/Level/Gain", ":INPut:GAIN:STATe"),
            ("Amplitude/Level/Gain", ":INPut:GAIN:STATe?"),
            ("Amplitude/Level/Gain", ":POWer:ATTenuation:AUTO ON"),
            ("Amplitude/Level/Gain", ":POWer:ATTenuation 0"),
            ("Amplitude/Level/Gain", ":POWer:ATTenuation 10"),
            ("Amplitude/Level/Gain", ":POWer:GAIN ON"),
            ("Amplitude/Level/Gain", ":POWer:GAIN OFF"),
            ("Amplitude/Level/Gain", ":POWer:GAIN 1"),
            ("Amplitude/Level/Gain", ":POWer:HSENsitive ON"),
            ("Amplitude/Level/Gain", ":POWer:HSENsitive OFF"),

            # Trace/Display
            ("Trace/Display", ":TRACe:DATA? TRACE1"),
            ("Trace/Display", ":TRACe1:MODE WRITe"), # Specific mode added for clarity
            ("Trace/Display", ":TRAC2:MODE MAXHold"),
            ("Trace/Display", ":TRAC2:MODE AVERage"),
            ("Trace/Display", ":TRAC3:MODE MINHold"),
            ("Trace/Display", ":DISPlay:WINDow:TRACe:TYPE?"),
            ("Trace/Display", ":DISPlay:WINDow:TRACe:Y:SCALe:SPACing LOGarithmic"),
            ("Trace/Display", ":TRACe:FORMat:DATA ASCii"), # For N9340B
            ("Trace/Display", ":FORMat:DATA ASCii"), # General

            # Marker
            ("Marker", ":CALCulate:MARKer1:MAX"),
            ("Marker", ":CALCulate:MARKer1:STATe"),
            ("Marker", ":CALCulate:MARKer1:X?"),
            ("Marker/Display", ":CALCulate:MARKer1:Y?"),

            # Memory/Preset
            ("Memory/Preset", ":MMEMory:CATalog:STATe?"),
            ("Memory/Preset", ":MMEMory:LOAD:STATe"),
            ("Memory/Preset", ":MMEMory:STORe:STATe"),
        ]

        processed_commands = []
        for category, cmd in default_categorized_commands:
            command_type_prefix = ""
            if cmd.strip().endswith("?"):
                command_type_prefix = "GET"
            elif cmd.strip() in ["*RST", ":SYSTem:DISPlay:UPDate"]:
                command_type_prefix = "DO"
            else:
                command_type_prefix = "SET"
            
            full_command_type = f"{category} - {command_type_prefix}"
            processed_commands.append((full_command_type, cmd, "")) # Add empty string for Variable
        return processed_commands

    def _on_tab_selected(self, event):
        """
        Called when this tab is selected in the notebook.
        Can be used to refresh data or update UI elements specific to this tab.
        For the interpreter, we ensure data is loaded/reloaded.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("VISA Interpreter Tab selected. Reloading data.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        self._load_data() # Reload data to ensure it's up-to-date
