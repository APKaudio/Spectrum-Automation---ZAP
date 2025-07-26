# src/tabs.py
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, ttk
import os
import csv
import xml.etree.ElementTree as ET
import sys
import inspect # Import inspect module
import threading # Added: Import the threading module
import io # Import io for StringIO or similar

# Import the new report converter utility functions
from utils.report_converter_utils import convert_html_report_to_csv, generate_csv_from_shw, convert_pdf_report_to_csv # Added PDF converter
# Import instrument_logic for setting focus frequency
from src.instrument_logic import set_focus_frequency_logic, set_marker_and_trace_modes_logic # Ensure both are imported
from utils.instrument_control import debug_print # Import debug_print
from src.gui_elements import TextRedirector # Import TextRedirector

class ReportConverterTab(ttk.Frame): # Changed from tk.Frame to ttk.Frame
    """
    A Tkinter Frame that provides functionality to convert spectrum analyzer
    report files (HTML, SHW, or PDF) into CSV format.
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
            4. Calls `create_widgets()` to build the GUI elements.
        Outputs: None (modifies GUI state)
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Initializing ReportConverterTab...", file=current_file, function=current_function)

        super().__init__(master, **kwargs)
        self.app_instance = app_instance

        # Configure style for this frame's widgets
        style = ttk.Style()
        style.configure("ReportConverter.TFrame", background="#000000") # Dark background
        style.configure("ReportConverter.TLabel", background="#000000", foreground="white")
        # Changed button foreground to black as requested
        style.configure("ReportConverter.TButton", background="#3a3a3a", foreground="black") 
        style.map("ReportConverter.TButton", background=[('active', '#6a6a6a')])
        style.configure("ReportConverter.TEntry", fieldbackground="#4a4a4a", foreground="black", insertbackground="white")
        style.configure("ReportConverter.TLabelframe", background="#333333", foreground="white")
        style.configure("ReportConverter.TLabelframe.Label", background="#333333", foreground="white")

        self.config(style="ReportConverter.TFrame") # Apply style to the main frame

        self.create_widgets()


    def create_widgets(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Creates the widgets for the Report Converter tab, including file selection,
        conversion buttons, and output display.
        """
        debug_print("Creating ReportConverterTab widgets...", file=file, function=function)
        # File selection frame
        file_selection_frame = ttk.LabelFrame(self, text="Select Report File", style="ReportConverter.TLabelframe", padding="10")
        file_selection_frame.pack(pady=10, padx=10, fill=tk.X)

        self.file_path_var = tk.StringVar()
        self.file_path_entry = ttk.Entry(file_selection_frame, textvariable=self.file_path_var, width=50, style="ReportConverter.TEntry")
        self.file_path_entry.grid(row=0, column=0, padx=5, pady=5, sticky=tk.EW)

        ttk.Button(file_selection_frame, text="Browse HTML", command=lambda: self._browse_file("HTML"), style="ReportConverter.TButton").grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_selection_frame, text="Browse SHW", command=lambda: self._browse_file("SHW"), style="ReportConverter.TButton").grid(row=0, column=2, padx=5, pady=5)
        # New button for PDF import
        ttk.Button(file_selection_frame, text="Sound Base PDF Import", command=lambda: self._browse_file("PDF"), style="ReportConverter.TButton").grid(row=0, column=3, padx=5, pady=5)


        file_selection_frame.grid_columnconfigure(0, weight=1)

        # Conversion button frame
        conversion_button_frame = ttk.Frame(self, style="ReportConverter.TFrame")
        conversion_button_frame.pack(pady=5, padx=10, fill=tk.X)

        ttk.Button(conversion_button_frame, text="Convert to CSV and Display Markers", command=self._on_convert_button_click, style="ReportConverter.TButton").pack(pady=5)

        # Console output for conversion
        self.conversion_console = scrolledtext.ScrolledText(self, wrap=tk.WORD, height=10, bg="black", fg="white", font=("Courier New", 10))
        self.conversion_console.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.conversion_console.config(state=tk.DISABLED)

    def _browse_file(self, file_type, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Opens a file dialog to select an HTML, SHW, or PDF file.
        """
        debug_print(f"Browsing for {file_type} file...", file=file, function=function)
        file_path = ""
        if file_type == "HTML":
            file_path = filedialog.askopenfilename(filetypes=[("HTML files", "*.html"), ("All files", "*.*")])
        elif file_type == "SHW":
            file_path = filedialog.askopenfilename(filetypes=[("SHW files", "*.shw"), ("All files", "*.*")])
        elif file_type == "PDF": # New PDF file type
            file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        
        if file_path:
            self.file_path_var.set(file_path)
            self.conversion_console.config(state=tk.NORMAL)
            self.conversion_console.delete(1.0, tk.END)
            # Initial message will now be handled by debug_print when conversion starts
            self.conversion_console.config(state=tk.DISABLED)
            debug_print(f"File selected: {file_path}", file=file, function=function)

    def _on_convert_button_click(self, file=__file__, function=inspect.currentframe().f_code.co_name):
        """
        Handles the conversion button click event.
        Determines file type and calls the appropriate conversion logic.
        """
        debug_print("Convert button clicked...", file=file, function=function) # This is the first debug message
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showwarning("No File Selected", "Please select an HTML, SHW, or PDF file to convert.")
            return

        file_extension = os.path.splitext(file_path)[1].lower()
        if file_extension == ".html":
            file_type = "HTML"
        elif file_extension == ".shw":
            file_type = "SHW"
        elif file_extension == ".pdf":
            file_type = "PDF"
        else:
            messagebox.showerror("Unsupported File Type", "Selected file is not a supported HTML, SHW, or PDF format.")
            return

        self.conversion_console.config(state=tk.NORMAL)
        self.conversion_console.delete(1.0, tk.END)
        self.conversion_console.config(state=tk.DISABLED)

        # Run conversion in a separate thread to keep GUI responsive
        conversion_thread = threading.Thread(target=self._convert_and_display, args=(file_path, file_type))
        conversion_thread.start()

    def _convert_and_display(self, file_path, file_type, file=__file__, function=inspect.currentframe().f_code.co_name):
        debug_print(f"Starting conversion for {file_path} (type: {file_type})...", file=file, function=function)
        error_message = None
        output_csv_file = "" # Initialize here for finally block

        # Temporarily redirect stdout to the conversion console
        old_stdout = sys.stdout
        sys.stdout = TextRedirector(self.conversion_console, "stdout")

        try:
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
                output_csv_file = os.path.join(self.app_instance.scan_directory_var.get(), 'MARKERS.CSV')
                
                with open(output_csv_file, 'w', newline='', encoding='utf-8') as csvfile:
                    csv_writer = csv.DictWriter(csvfile, fieldnames=headers)
                    csv_writer.writeheader()
                    csv_writer.writerows(rows)
                
                # Removed messagebox.showinfo
                print(f"\n✅ Successfully converted '{file_name}' to '{os.path.basename(output_csv_file)}'")
                
                # Display MARKERS.CSV file contents
                print(f"\n--- Contents of {os.path.basename(output_csv_file)} ---")
                try:
                    with open(output_csv_file, 'r', encoding='utf-8') as f:
                        csv_content = f.read()
                        print(csv_content)
                    print(f"--- End of {os.path.basename(output_csv_file)} ---")
                except Exception as e:
                    print(f"❌ Error reading {os.path.basename(output_csv_file)}: {e}")

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
        
        finally:
            # Restore stdout
            sys.stdout = old_stdout
            if error_message:
                print(f"❌ Conversion failed for {file_name}: {error_message}")
                debug_print(f"Conversion failed for {file_name}: {error_message}", file=file, function=function)

