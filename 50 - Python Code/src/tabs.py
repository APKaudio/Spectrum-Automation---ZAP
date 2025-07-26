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
    load_selected_preset_logic,
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
        sys.stderr = TextRedirector(self.message_text, "stderr") # Also redirect stderr for warnings

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
        self.preset_files = [] # To store list of .sta files
        self.selected_preset = None # To store the currently selected preset name

        self._create_widgets()
        self._populate_preset_list() # Populate on initialization

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

    def _populate_preset_list(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Populates the list of .sta preset files from the C:\\PRESETS directory
        and creates buttons for each.
        """
        preset_dir = "C:\\PRESETS" # Hardcoded directory as per original code
        self.preset_files = []

        # Clear existing buttons
        for widget in self.inner_buttons_frame.winfo_children():
            widget.destroy()

        if not os.path.exists(preset_dir):
            debug_print(f"Preset directory not found: {preset_dir}", file=file, function=function)
            ttk.Label(self.inner_buttons_frame, text=f"Preset directory not found: {preset_dir}").pack(padx=10, pady=10)
            return

        try:
            for filename in os.listdir(preset_dir):
                if filename.lower().endswith(".sta"):
                    self.preset_files.append(os.path.splitext(filename)[0]) # Store name without extension
            self.preset_files.sort() # Sort alphabetically
            debug_print(f"Found {len(self.preset_files)} preset files.", file=file, function=function)

        except Exception as e:
            debug_print(f"Error listing preset files: {e}", file=file, function=function)
            ttk.Label(self.inner_buttons_frame, text=f"Error loading presets: {e}").pack(padx=10, pady=10)
            return

        if not self.preset_files:
            ttk.Label(self.inner_buttons_frame, text="No .sta preset files found in C:\\PRESETS.").pack(padx=10, pady=10)
            return

        for i, preset_name in enumerate(self.preset_files):
            # Use ttk.Button for consistent styling
            button = ttk.Button(self.inner_buttons_frame, text=preset_name, 
                                command=lambda name=preset_name: self._on_preset_button_click(name),
                                style='GreyText.TButton') # Apply a consistent style
            button.pack(fill=tk.X, padx=5, pady=2) # Pack to fill horizontally

        # Update scroll region after adding buttons
        self.inner_buttons_frame.update_idletasks()
        self.preset_canvas.config(scrollregion=self.preset_canvas.bbox("all"))
        debug_print("Preset buttons populated and scroll region updated.", file=file, function=function)


    def _on_preset_button_click(self, preset_name, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Callback for individual preset buttons. Sets the selected preset and updates
        the load preset button state in the main app.
        """
        debug_print(f"Preset button clicked: {preset_name}", file=file, function=function)
        self.selected_preset = preset_name
        
        # Update the load_preset_button state in the main app
        if self.app_instance and self.app_instance.inst:
            self.app_instance.load_preset_button.config(state=tk.NORMAL)
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
        """
        if self.app_instance and self.app_instance.inst:
            debug_print("Querying presets from device...", file=file, function=function)
            # This calls the logic function from instrument_logic.py
            query_device_presets_logic(self.app_instance)
            # The populate_resources_logic (or a similar refresh) should update the dropdown
            # For presets, it might involve a different mechanism if the device itself lists them.
            # For now, assuming query_device_presets_logic updates self.app_instance.instrument_list
            # or a similar structure that _populate_preset_list can use.
            # If query_device_presets_logic updates a different list, this tab needs to be informed
            # and call _populate_preset_list with that data.
            # As per current design, _populate_preset_list reads from C:\PRESETS, not directly from device query.
            # So, this button should ideally trigger a refresh of the local file list if the device
            # somehow syncs with the local directory, or if the device itself returns a list of presets.
            # Given the current `load_selected_preset` uses `C:\PRESETS`, this button should likely refresh that.
            self._populate_preset_list() # Re-populate from local directory after query attempt
            print("Finished querying presets from device (local list refreshed).")
        else:
            debug_print("Cannot query presets: Instrument not connected.", file=file, function=function)
            messagebox.showwarning("Not Connected", "Please connect to an instrument first to query presets.")

    def _on_tab_selected(self, event):
        """
        Callback when this tab is selected. Updates the state of the query presets button.
        """
        debug_print("PresetFilesTab selected.", file=__file__, function=inspect.currentframe().f_code.co_name)
        if self.app_instance and self.app_instance.inst:
            self.query_presets_button.config(state=tk.NORMAL)
            # Optionally, refresh the preset list every time the tab is selected
            self._populate_preset_list()
        else:
            self.query_presets_button.config(state=tk.DISABLED)
