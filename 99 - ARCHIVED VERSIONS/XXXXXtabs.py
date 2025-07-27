# src/tabs.py
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, ttk
import os
import csv
import xml.etree.ElementTree as ET
import sys
import inspect # Import inspect module
import threading # Added: Import the threading module

# Import the new report converter utility functions
from utils.report_converter_utils import convert_html_report_to_csv, generate_csv_from_shw, convert_pdf_report_to_csv
# Import instrument_logic for setting focus frequency and loading presets
from src.instrument_logic import (
    set_focus_frequency_logic,
    set_marker_and_trace_modes_logic,
    load_selected_preset_logic, # Ensure load_selected_preset_logic is imported
    query_device_presets_logic # Ensure query_device_presets_logic is imported
)
from utils.instrument_control import debug_print, write_safe # Import debug_print and write_safe
from src.gui_elements import TextRedirector # Import TextRedirector - FIX: Ensure this is present

# REMOVED MarkersDisplayTab from here. It is now solely in src/marker_logic.py

class ReportConverterTab(ttk.Frame): # Changed from tk.Frame to ttk.Frame
    """
    A Tkinter Frame that provides functionality to convert spectrum analyzer
    report files (HTML, SHW, or Soundbase PDF) into CSV format.
    """
    def __init__(self, master=None, app_instance=None, **kwargs):
        """
        Initializes the ReportConverterTab.

        Inputs:
            master (tk.Widget): The parent widget.
            app_instance (App): The main application instance, used for accessing
                                shared state like output directory.
            **kwargs: Arbitrary keyword arguments for Tkinter Frame.
        """
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self.output_csv_path = None # To store the path of the last generated CSV

        self._create_widgets()

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Report Converter tab.
        """
        # Configure grid for responsiveness
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # Row for message area should expand

        # Frame for buttons
        button_frame = ttk.Frame(self)
        button_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        button_frame.grid_columnconfigure(2, weight=1)


        # Convert HTML to CSV Button
        ttk.Button(button_frame, text="Convert HTML to CSV", command=self._convert_html_to_csv, style='Blue.TButton').grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        # Convert SHW to CSV Button
        ttk.Button(button_frame, text="Convert SHW to CSV", command=self._convert_shw_to_csv, style='Blue.TButton').grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Convert Soundbase PDF to CSV Button
        ttk.Button(button_frame, text="Convert Soundbase PDF to CSV", command=self._convert_pdf_to_csv, style='Blue.TButton').grid(row=0, column=2, padx=5, pady=5, sticky="ew")


        # Message/Output Area
        self.message_text = scrolledtext.ScrolledText(self, wrap="word", height=10, bg="#2b2b2b", fg="#cccccc", insertbackground="white")
        self.message_text.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        # Removed _redirect_output_to_message_area() from here. Redirection is now temporary.

    # Removed _redirect_output_to_message_area and _restore_output methods
    # as redirection is now handled within _convert_and_display for temporary scope.

    def _convert_html_to_csv(self):
        """
        Prompts user for an HTML file, converts it to CSV, and saves it.
        """
        file_path = filedialog.askopenfilename(
            title="Select HTML Report File",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")]
        )
        if not file_path:
            print("HTML conversion cancelled.")
            return

        # Run conversion in a separate thread to keep GUI responsive
        conversion_thread = threading.Thread(target=self._convert_and_display, args=(file_path, "HTML"))
        conversion_thread.start()

    def _convert_shw_to_csv(self):
        """
        Prompts user for an SHW (XML) file, converts it to CSV, and saves it.
        """
        file_path = filedialog.askopenfilename(
            title="Select SHW Report File",
            filetypes=[("SHW files", "*.shw"), ("XML files", "*.xml"), ("All files", "*.*")]
        )
        if not file_path:
            print("SHW conversion cancelled.")
            return

        # Run conversion in a separate thread to keep GUI responsive
        conversion_thread = threading.Thread(target=self._convert_and_display, args=(file_path, "SHW"))
        conversion_thread.start()

    def _convert_pdf_to_csv(self):
        """
        Prompts user for a Soundbase PDF file, converts it to CSV, and saves it.
        """
        file_path = filedialog.askopenfilename(
            title="Select Soundbase PDF Report File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not file_path:
            print("PDF conversion cancelled.")
            return

        # Run conversion in a separate thread to keep GUI responsive
        conversion_thread = threading.Thread(target=self._convert_and_display, args=(file_path, "PDF"))
        conversion_thread.start()

    def _convert_and_display(self, file_path, file_type, file=__file__, function=inspect.currentframe().f_code.co_name):
        debug_print(f"Starting conversion for {file_path} (type: {file_type})...", file=file, function=function)
        error_message = None
        output_csv_file = ""

        # Temporarily redirect stdout and stderr to the conversion console
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = TextRedirector(self.message_text, "stdout")
        sys.stderr = TextRedirector(self.message_text, "stderr", file=file, function=function) # Also redirect stderr for warnings

        try:
            # Clear previous messages in the conversion console
            self.message_text.config(state=tk.NORMAL)
            self.message_text.delete(1.0, tk.END)
            self.message_text.config(state=tk.DISABLED)

            # This message will now go to the tab's console
            print(f"Processing '{os.path.basename(file_path)}'...") 

            file_name = os.path.basename(file_path)
            headers = []
            rows = []

            if file_type == "HTML":
                with open(file_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                headers, rows = convert_html_report_to_csv(html_content)
            elif file_type == "SHW":
                headers, rows = generate_csv_from_shw(file_path)
            elif file_type == "PDF":
                headers, rows = convert_pdf_report_to_csv(file_path)
            
            if headers and rows:
                output_folder = self.app_instance.output_folder_var.get()
                if not output_folder or not os.path.isdir(output_folder):
                    error_message = "Invalid Output Folder. Please set a valid output directory in the Scan Configuration tab."
                    print(f"❌ Conversion failed: {error_message}") # This will go to the tab's console
                    messagebox.showwarning("Invalid Output Folder", error_message) # Still show messagebox
                    return # Exit early if output folder is invalid

                output_csv_filename = "MARKERS.CSV"
                self.output_csv_path = os.path.join(output_folder, output_csv_filename)
                
                with open(self.output_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=headers)
                    writer.writeheader()
                    writer.writerows(rows)
                
                print(f"\n✅ Successfully converted '{file_name}' to '{os.path.basename(self.output_csv_path)}'")
                
                # Display MARKERS.CSV file contents in the tab's console
                print(f"\n--- Contents of {os.path.basename(self.output_csv_path)} ---")
                try:
                    with open(self.output_csv_path, 'r', encoding='utf-8') as f:
                        csv_content = f.read()
                        print(csv_content)
                    print(f"--- End of {os.path.basename(self.output_csv_path)} ---")
                except Exception as e:
                    print(f"❌ Error reading {os.path.basename(self.output_csv_path)}: {e}")

                # Call the method on the main App instance to update the Markers Display tab
                if hasattr(self.app_instance, 'markers_display_tab'):
                    self.app_instance.markers_display_tab.update_markers_data(headers, rows)
                else:
                    debug_print("MarkersDisplayTab instance not found on app_instance.", file=__file__, function=inspect.currentframe().f_code.co_name)
            else:
                error_message = f"No relevant data could be extracted from '{file_name}'. CSV file was not created."
                messagebox.showwarning("No Data Extracted", error_message)
                print(f"🚫 {error_message}") # This will go to the tab's console

        except FileNotFoundError as e:
            error_message = f"File not found: {e}"
            messagebox.showerror("File Error", error_message)
            print(f"❌ {error_message}") # This will go to the tab's console
        except ET.ParseError as e:
            error_message = f"Error parsing XML (SHW) file: {e}"
            messagebox.showerror("Parsing Error", error_message)
            print(f"❌ {error_message}") # This will go to the tab's console
        except Exception as e:
            error_message = f"An unexpected error occurred during conversion: {e}"
            messagebox.showerror("Conversion Error", error_message)
            print(f"❌ {error_message}") # This will go to the tab's console
        
        finally:
            # Restore stdout and stderr
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            if error_message:
                # This debug_print will now go to the main console
                debug_print(f"Conversion failed for {file_name}: {error_message}", file=file, function=function)


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
                                style='LargePreset.TButton') # Changed style to 'LargePreset.TButton'
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

