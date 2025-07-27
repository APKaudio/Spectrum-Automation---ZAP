# src/report_converter_tab.py
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, ttk
import os
import csv
import xml.etree.ElementTree as ET
import sys
import inspect
import threading

# Import the new report converter utility functions
from utils.report_converter_utils import convert_html_report_to_csv, generate_csv_from_shw, convert_pdf_report_to_csv
from src.gui_elements import TextRedirector
from utils.instrument_control import debug_print # Import debug_print

class ReportConverterTab(ttk.Frame):
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
        # Removed 'file=file, function=function' as TextRedirector does not accept them
        sys.stdout = TextRedirector(self.message_text, "stdout")
        sys.stderr = TextRedirector(self.message_text, "stderr") # Corrected: removed file= and function=

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

