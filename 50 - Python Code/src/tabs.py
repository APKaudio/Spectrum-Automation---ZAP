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

# REMOVED MarkersDisplayTab from here. It is now solely in src/marker_logic.py

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
