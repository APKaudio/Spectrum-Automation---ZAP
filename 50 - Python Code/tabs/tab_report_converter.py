# src/report_converter_tab.py
import tkinter as tk
# from tkinter import messagebox, scrolledtext, filedialog, ttk # Removed messagebox
from tkinter import scrolledtext, filedialog, ttk # Keep other imports
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
    def __init__(self, master=None, app_instance=None, console_print_func=None, **kwargs):
        """
        Initializes the ReportConverterTab.

        Inputs:
            master (tk.Widget): The parent widget.
            app_instance (App): The main application instance, used for accessing
                                shared state like output directory.
            console_print_func (function, optional): Function to use for console output.
            **kwargs: Arbitrary keyword arguments for Tkinter Frame.
        """
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self.console_print_func = console_print_func if console_print_func else print # Use provided func or default print
        self.output_csv_path = None # To store the path of the last generated CSV

        self._create_widgets()

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Report Converter tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Creating ReportConverterTab widgets.", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # Configure grid for responsiveness
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # Row for message area should expand

        # Frame for buttons
        button_frame = ttk.Frame(self, style='TFrame') # Apply TFrame style
        button_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        button_frame.grid_columnconfigure(2, weight=1)

        # Convert HTML to CSV Button
        self.html_button = ttk.Button(button_frame, text="Convert HTML to CSV", command=self._convert_html_to_csv, style='Blue.TButton')
        self.html_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        # Convert SHW to CSV Button
        self.shw_button = ttk.Button(button_frame, text="Convert SHW to CSV", command=self._convert_shw_to_csv, style='Blue.TButton')
        self.shw_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Convert Soundbase PDF to CSV Button
        self.pdf_button = ttk.Button(button_frame, text="Convert Soundbase PDF to CSV", command=self._convert_pdf_to_csv, style='Blue.TButton')
        self.pdf_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

        # Conversion Log Console
        conversion_log_frame = ttk.LabelFrame(self, text="Conversion Log", style='Dark.TLabelframe')
        conversion_log_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        conversion_log_frame.grid_columnconfigure(0, weight=1)
        conversion_log_frame.grid_rowconfigure(0, weight=1)

        self.conversion_console = scrolledtext.ScrolledText(conversion_log_frame, wrap="word", height=10, bg="#2b2b2b", fg="#cccccc", insertbackground="white")
        self.conversion_console.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.conversion_console.config(state=tk.DISABLED) # Make it read-only

    def _browse_report_file(self):
        """
        Opens a file dialog to select a report file for conversion.
        This method is now primarily for the Entry field, not directly used by the conversion buttons.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Opening file browser for report conversion.", file=current_file, function=current_function, console_print_func=self.console_print_func)

        file_path = filedialog.askopenfilename(
            title="Select Report File",
            filetypes=[("Report Files", "*.html *.shw *.pdf"), ("All files", "*.*")]
        )
        if file_path:
            self.report_file_path_var.set(file_path)
            self.console_print_func(f"Selected file: {os.path.basename(file_path)}")
            debug_print(f"Selected report file: {file_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _convert_html_to_csv(self):
        """
        Prompts user for an HTML file, converts it to CSV, and saves it.
        """
        file_path = filedialog.askopenfilename(
            title="Select HTML Report File",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")]
        )
        if not file_path:
            self.console_print_func("ℹ️ Info: HTML conversion cancelled.")
            return

        self._disable_buttons()
        conversion_thread = threading.Thread(target=self._perform_conversion, args=(file_path, "HTML"))
        conversion_thread.daemon = True
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
            self.console_print_func("ℹ️ Info: SHW conversion cancelled.")
            return

        self._disable_buttons()
        conversion_thread = threading.Thread(target=self._perform_conversion, args=(file_path, "SHW"))
        conversion_thread.daemon = True
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
            self.console_print_func("ℹ️ Info: PDF conversion cancelled.")
            return

        self._disable_buttons()
        conversion_thread = threading.Thread(target=self._perform_conversion, args=(file_path, "PDF"))
        conversion_thread.daemon = True
        conversion_thread.start()

    def _disable_buttons(self):
        """Disables all conversion buttons during a conversion process."""
        self.html_button.config(state=tk.DISABLED)
        self.shw_button.config(state=tk.DISABLED)
        self.pdf_button.config(state=tk.DISABLED)

    def _enable_buttons(self):
        """Enables all conversion buttons after a conversion process."""
        self.html_button.config(state=tk.NORMAL)
        self.shw_button.config(state=tk.NORMAL)
        self.pdf_button.config(state=tk.NORMAL)

    def _perform_conversion(self, file_path, file_type):
        """
        Performs the actual file conversion based on file type.
        Redirects print statements to the conversion_console during the process.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Performing conversion for {file_path} (type: {file_type}) in thread.", file=current_file, function=current_function, console_print_func=self.console_print_func)

        headers = []
        rows = []
        error_message = None
        output_csv_path = None

        # Temporarily redirect stdout and stderr to the conversion_console
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = TextRedirector(self.conversion_console, "stdout")
        sys.stderr = TextRedirector(self.conversion_console, "stderr")

        try:
            # Clear previous conversion log
            self.conversion_console.config(state=tk.NORMAL)
            self.conversion_console.delete(1.0, tk.END)
            self.conversion_console.config(state=tk.DISABLED)
            self.console_print_func(f"Processing '{os.path.basename(file_path)}'...")

            file_name = os.path.basename(file_path)
            output_dir = self.app_instance.output_folder_var.get()
            os.makedirs(output_dir, exist_ok=True) # Ensure output directory exists
            
            base_name, ext = os.path.splitext(file_name)
            # The output CSV filename will always be MARKERS.CSV for the MarkersDisplayTab to find it
            output_csv_path = os.path.join(output_dir, "MARKERS.CSV")

            if file_type == 'HTML':
                self.console_print_func("Detected HTML file. Converting...")
                with open(file_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                headers, rows = convert_html_report_to_csv(html_content, console_print_func=self.console_print_func)
            elif file_type == 'SHW':
                self.console_print_func("Detected SHW file. Converting...")
                headers, rows = generate_csv_from_shw(file_path, console_print_func=self.console_print_func)
            elif file_type == 'PDF':
                self.console_print_func("Detected PDF file. Converting...")
                headers, rows = convert_pdf_report_to_csv(file_path, console_print_func=self.console_print_func)
            else:
                error_message = f"Unsupported file type: {file_type}. This should not happen."
                self.console_print_func(f"❌ {error_message}")
                debug_print(f"Unsupported file type: {file_type}", file=current_file, function=current_function, console_print_func=self.console_print_func)

            if not error_message and headers and rows:
                with open(output_csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=headers)
                    writer.writeheader()
                    writer.writerows(rows)
                self.output_csv_path = output_csv_path
                
                self.console_print_func(f"\n✅ Successfully converted '{file_name}' to '{os.path.basename(self.output_csv_path)}'")
                
                # Display MARKERS.CSV file contents in the tab's console
                self.console_print_func(f"\n--- Contents of {os.path.basename(self.output_csv_path)} ---")
                try:
                    with open(self.output_csv_path, 'r', encoding='utf-8') as f:
                        csv_content = f.read()
                        self.console_print_func(csv_content)
                    self.console_print_func(f"--- End of {os.path.basename(self.output_csv_path)} ---")
                except Exception as e:
                    self.console_print_func(f"❌ Error reading {os.path.basename(self.output_csv_path)}: {e}")

                # Call the method on the main App instance to update the Markers Display tab
                if hasattr(self.app_instance, 'markers_display_tab'):
                    self.app_instance.markers_display_tab.update_markers_data(headers, rows)
                else:
                    debug_print("MarkersDisplayTab instance not found on app_instance.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            else:
                error_message = f"No relevant data could be extracted from '{file_name}'. CSV file was not created."
                self.console_print_func(f"🚫 {error_message}") # This will go to the tab's console

        except FileNotFoundError as e:
            error_message = f"File not found: {e}"
            self.console_print_func(f"❌ {error_message}") # This will go to the tab's console
        except ET.ParseError as e:
            error_message = f"Error parsing XML (SHW) file: {e}"
            self.console_print_func(f"❌ {error_message}") # This will go to the tab's console
        except Exception as e:
            error_message = f"An unexpected error occurred during conversion: {e}"
            self.console_print_func(f"❌ {error_message}") # This will go to the tab's console
        
        finally:
            # Restore stdout and stderr
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            if error_message:
                # This debug_print will now go to the main console
                debug_print(f"Conversion failed for {file_name}: {error_message}", file=current_file, function=current_function, console_print_func=self.console_print_func)
                # Schedule a print on the main thread for the main console
                self.app_instance.after(0, lambda: self.app_instance._print_to_gui_console(f"❌ Conversion failed for {file_name}. See Report Converter Log for details."))
            else:
                self.app_instance.after(0, lambda: self.app_instance._print_to_gui_console(f"✅ Successfully converted {file_name} to CSV: {output_csv_path}"))

            # Re-enable conversion buttons on the main thread
            self.app_instance.after(0, self._enable_buttons)

