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
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0) # File selection
        self.grid_rowconfigure(1, weight=1) # Console output

        # File Selection Frame
        file_selection_frame = ttk.LabelFrame(self, text="Select Report File to Convert")
        file_selection_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        file_selection_frame.grid_columnconfigure(0, weight=1)

        self.file_path_var = tk.StringVar()
        self.file_path_entry = ttk.Entry(file_selection_frame, textvariable=self.file_path_var, state='readonly')
        self.file_path_entry.grid(row=0, column=0, padx=5, pady=2, sticky="ew")

        ttk.Button(file_selection_frame, text="Browse", command=self._browse_file, style='GreyText.TButton').grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(file_selection_frame, text="Convert to CSV", command=self._start_conversion_thread, style='Accent.TButton').grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        # Console output for conversion process
        self.console_text = scrolledtext.ScrolledText(self, wrap="word", height=10, bg="#1a1a1a", fg="#cccccc", insertbackground="white")
        self.console_text.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")

    def _browse_file(self):
        """
        Opens a file dialog to select a report file (HTML, SHW, or PDF).
        """
        file_path = filedialog.askopenfilename(
            filetypes=[("Report Files", "*.html *.shw *.pdf"),
                       ("HTML Files", "*.html"),
                       ("SHW (XML) Files", "*.shw"),
                       ("PDF Files", "*.pdf"),
                       ("All Files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            debug_print(f"Selected file for conversion: {file_path}", file=__file__, function=inspect.currentframe().f_code.co_name)

    def _start_conversion_thread(self):
        """
        Starts the file conversion process in a separate thread to keep the GUI responsive.
        """
        file_path = self.file_path_var.get()
        if not file_path:
            print("🚫 Please select a file to convert first.")
            # tk.messagebox.showwarning("No File Selected", "Please select a file to convert first.") # Removed messagebox
            return

        print(f"\nStarting conversion for: {os.path.basename(file_path)}...")
        debug_print(f"Starting conversion for: {os.path.basename(file_path)}", file=__file__, function=inspect.currentframe().f_code.co_name)

        # Disable conversion button during process
        self.master.master.master.children['!button'].config(state=tk.DISABLED) # Access the convert button via its parent hierarchy
        # This is a bit brittle, a direct reference to the convert button would be better if it were an instance variable.
        # For now, assuming the button is always the second child of file_selection_frame.
        
        # Redirect stdout/stderr for this tab's console
        self.old_stdout = sys.stdout
        self.old_stderr = sys.stderr
        sys.stdout = TextRedirector(self.console_text, "stdout")
        sys.stderr = TextRedirector(self.console_text, "stderr")

        conversion_thread = threading.Thread(target=self._convert_file_logic, args=(file_path,))
        conversion_thread.daemon = True
        conversion_thread.start()

    def _convert_file_logic(self, file_path):
        """
        Contains the actual file conversion logic, running in a separate thread.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        
        file_name = os.path.basename(file_path)
        output_folder = self.app_instance.output_folder_var.get()
        if not output_folder:
            output_folder = os.path.dirname(file_path) # Default to source directory if not set
            print(f"⚠️ Output directory not set. Using source directory: {output_folder}")
            debug_print(f"Output directory not set. Using source directory: {output_folder}", file=current_file, function=current_function)

        os.makedirs(output_folder, exist_ok=True) # Ensure output directory exists

        base_name = os.path.splitext(file_name)[0]
        output_csv_file = os.path.join(output_folder, f"{base_name}_extracted.csv")
        markers_csv_file = os.path.join(output_folder, "MARKERS.CSV") # Consistent markers file

        headers = []
        rows = []
        error_message = None

        try:
            if file_path.lower().endswith(".html"):
                print("Detected HTML file. Extracting data...")
                with open(file_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                headers, rows = convert_html_report_to_csv(html_content, console_print_func=print)
            elif file_path.lower().endswith(".shw"):
                print("Detected SHW (XML) file. Extracting data...")
                headers, rows = generate_csv_from_shw(file_path, console_print_func=print)
            elif file_path.lower().endswith(".pdf"):
                print("Detected PDF file. Extracting data...")
                headers, rows = convert_pdf_report_to_csv(file_path, console_print_func=print)
            else:
                error_message = "Unsupported file type. Please select an HTML, SHW, or PDF file."
                print(f"❌ {error_message}")
                # tk.messagebox.showerror("Unsupported File Type", error_message) # Removed messagebox
                return

            if headers and rows:
                # Save extracted data to the main output CSV
                with open(output_csv_file, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=headers)
                    writer.writeheader()
                    writer.writerows(rows)
                print(f"✅ Conversion complete. Data saved to: {output_csv_file}")
                debug_print(f"Conversion complete. Data saved to: {output_csv_file}", file=current_file, function=current_function)
                self.output_csv_path = output_csv_file # Store for potential future use

                # Also save to MARKERS.CSV
                with open(markers_csv_file, 'w', newline='', encoding='utf-8') as markers_file:
                    writer = csv.DictWriter(markers_file, fieldnames=headers)
                    writer.writeheader()
                    writer.writerows(rows)
                print(f"✅ Markers also saved to: {markers_csv_file}")
                debug_print(f"Markers also saved to: {markers_csv_file}", file=current_file, function=current_function)

                # Update the MarkersDisplayTab if it exists
                if hasattr(self.app_instance, 'markers_display_tab') and self.app_instance.markers_display_tab:
                    self.app_instance.markers_display_tab.update_markers_data(headers, rows)
                else:
                    debug_print("MarkersDisplayTab instance not found on app_instance.", file=__file__, function=inspect.currentframe().f_code.co_name)
            else:
                error_message = f"No relevant data could be extracted from '{file_name}'. CSV file was not created."
                # Removed messagebox.showwarning
                print(f"🚫 {error_message}") # This will go to the tab's console

        except FileNotFoundError as e:
            error_message = f"File not found: {e}"
            # Removed messagebox.showerror
            print(f"❌ {error_message}") # This will go to the tab's console
        except ET.ParseError as e:
            error_message = f"Error parsing XML (SHW) file: {e}"
            # Removed messagebox.showerror
            print(f"❌ {error_message}") # This will go to the tab's console
        except Exception as e:
            error_message = f"An unexpected error occurred during conversion: {e}"
            # Removed messagebox.showerror
            print(f"❌ {error_message}") # This will go to the tab's console
        
        finally:
            # Restore stdout and stderr
            sys.stdout = self.old_stdout
            sys.stderr = self.old_stderr
            if error_message:
                # This debug_print will now go to the main console
                debug_print(f"Conversion failed for {file_name}: {error_message}", file=current_file, function=current_function)
                # Schedule a messagebox on the main thread if an error occurred
                self.app_instance.after(0, lambda: tk.messagebox.showerror("Conversion Error", f"Conversion failed for {file_name}.\nSee console for details."))
            else:
                self.app_instance.after(0, lambda: tk.messagebox.showinfo("Conversion Complete", f"Successfully converted {file_name} to CSV."))


            # Re-enable conversion button
            self.master.master.master.children['!button'].config(state=tk.NORMAL) # Access the convert button via its parent hierarchy


