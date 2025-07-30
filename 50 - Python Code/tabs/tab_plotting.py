# src/tab_plotting.py
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import os
import pandas as pd
import inspect
import webbrowser # For opening HTML plot in browser
import re # Added for regex in file grouping
import platform # For opening folder cross-platform

from utils.instrument_control import debug_print
from src.plot_logic import plot_single_scan_data # Import only plot_single_scan_data
# Import generate_current_cycle_average_csv_and_plot from its correct location

from process_math.averaging_utils import average_scan # NEW import

class PlottingTab(ttk.Frame):
    """
    A Tkinter Frame that provides functionality for plotting scan data.
    """
    def __init__(self, master=None, app_instance=None, console_print_func=None, **kwargs):
        """
        Initializes the PlottingTab.

        Inputs:
            master (tk.Widget): The parent widget.
            app_instance (App): The main application instance, used for accessing
                                shared state like collected_scans_dataframes and output directory.
            console_print_func (function): Function to print messages to the GUI console.
            **kwargs: Arbitrary keyword arguments for Tkinter Frame.
        """
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self.console_print_func = console_print_func if console_print_func else print # Use provided func or print
        self.current_plot_file = None # To store the path of the last generated plot HTML
        self.last_opened_folder = None # To remember the last opened folder for averaging
        self.last_applied_math_folder = None # To store the path of the last folder created by applied math

        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self._create_widgets()
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _create_widgets(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # SCAN Plotting Options Frame (for single scan and current cycle average)
        plotting_options_frame = ttk.LabelFrame(self, text="SCAN Plotting Options", padding="10")
        plotting_options_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # Inner frame for 50/50 split
        scan_options_inner_frame = ttk.Frame(plotting_options_frame)
        scan_options_inner_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew", columnspan=2)
        scan_options_inner_frame.grid_columnconfigure(0, weight=1)
        scan_options_inner_frame.grid_columnconfigure(1, weight=1)

        # Left 50% - HTML Output Options
        html_output_options_frame = ttk.LabelFrame(scan_options_inner_frame, text="HTML Output Options", padding="10")
        html_output_options_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        html_output_options_frame.grid_columnconfigure(0, weight=1)

        self.open_html_after_complete_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(html_output_options_frame, text="Plot the HTML after every scan", variable=self.open_html_after_complete_var).grid(row=0, column=0, padx=5, pady=2, sticky="w")

        self.create_html_var = tk.BooleanVar(value=True) # New checkbox
        ttk.Checkbutton(html_output_options_frame, text="Create HTML", variable=self.create_html_var,
                        command=self._on_create_html_checkbox_changed).grid(row=1, column=0, padx=5, pady=2, sticky="w")


        # Right 50% - Scan Markers to Plot
        scan_markers_to_plot_frame = ttk.LabelFrame(scan_options_inner_frame, text="Scan Markers to Plot", padding="10")
        scan_markers_to_plot_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        scan_markers_to_plot_frame.grid_columnconfigure(0, weight=1)

        self.include_scan_tv_markers_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(scan_markers_to_plot_frame, text="Include TV Band Markers", variable=self.include_scan_tv_markers_var,
                        command=self._on_scan_marker_checkbox_changed).grid(row=0, column=0, padx=5, pady=2, sticky="w")

        self.include_scan_gov_markers_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(scan_markers_to_plot_frame, text="Include Government Band Markers", variable=self.include_scan_gov_markers_var,
                        command=self._on_scan_marker_checkbox_changed).grid(row=1, column=0, padx=5, pady=2, sticky="w")

        self.include_scan_markers_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(scan_markers_to_plot_frame, text="Include Markers", variable=self.include_scan_markers_var,
                        command=self._on_scan_marker_checkbox_changed).grid(row=2, column=0, padx=5, pady=2, sticky="w")

        self.include_scan_intermod_markers_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(scan_markers_to_plot_frame, text="Include Intermodulations", variable=self.include_scan_intermod_markers_var,
                        command=self._on_scan_marker_checkbox_changed).grid(row=3, column=0, padx=5, pady=2, sticky="w")


        # Plotting Buttons below the split frames
        self.plot_button = ttk.Button(plotting_options_frame, text="Plot Single Scan", command=self._plot_single_scan)
        self.plot_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew") # Adjusted row
        # self.plot_button.config(state=tk.DISABLED) # Will be enabled if a file is selected

        self.plot_average_button = ttk.Button(plotting_options_frame, text="Plot Current Cycle Average (All Traces)", command=self._plot_current_cycle_average)
        self.plot_average_button.grid(row=2, column=0, padx=5, pady=5, sticky="ew") # Adjusted row
        self.plot_average_button.config(state=tk.DISABLED) # Disable until data is available

        self.open_last_plot_button = ttk.Button(plotting_options_frame, text="Open Last Plot", command=self._open_last_plot)
        self.open_last_plot_button.grid(row=3, column=0, padx=5, pady=5, sticky="ew") # Adjusted row

        # --- Plotting Averages from Folder Section ---
        self.averaging_folder_frame = ttk.LabelFrame(self, text="Plotting Averages from Folder", padding="10")
        self.averaging_folder_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        # Open Folder to Average Button (Top of this section)
        self.open_folder_button = ttk.Button(self.averaging_folder_frame, text="Open Folder to Average", command=self._open_folder_for_averaging)
        self.open_folder_button.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        # Discovered Series of Scans Frame
        self.discovered_series_frame = ttk.LabelFrame(self.averaging_folder_frame, text="Discovered Series of Scans", padding="10")
        self.discovered_series_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        self.dynamic_avg_buttons_frame = ttk.Frame(self.discovered_series_frame) # This frame will hold dynamically created buttons.
        self.dynamic_avg_buttons_frame.pack(fill="both", expand=True) # Use pack for dynamic buttons within this frame

        # Math and Markers Container Frame (for 50/50 split)
        math_and_markers_frame = ttk.Frame(self.averaging_folder_frame)
        math_and_markers_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        math_and_markers_frame.grid_columnconfigure(0, weight=1) # Make columns expandable
        math_and_markers_frame.grid_columnconfigure(1, weight=1)

        # Apply Math Frame
        apply_math_frame = ttk.LabelFrame(math_and_markers_frame, text="Apply Math", padding="10")
        apply_math_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        apply_math_frame.grid_columnconfigure(0, weight=1) # Make column expandable

        self.avg_type_vars = {
            "Average": tk.BooleanVar(value=True), # Set to True by default
            "Median": tk.BooleanVar(value=True), # Set to True by default
            "Range": tk.BooleanVar(value=True), # Set to True by default
            "Std Dev": tk.BooleanVar(value=True), # Set to True by default
            "Variance": tk.BooleanVar(value=True), # Set to True by default
            "PSD (dBm/Hz)": tk.BooleanVar(value=True) # Set to True by default
        }
        for i, (text, var) in enumerate(self.avg_type_vars.items()):
            ttk.Checkbutton(apply_math_frame, text=text, variable=var,
                            command=self._on_avg_type_checkbox_changed).grid(row=i, column=0, padx=5, pady=2, sticky="w")

        # Markers to Plot Frame (Includes TV, Government, General, and Intermodulations)
        markers_to_plot_frame = ttk.LabelFrame(math_and_markers_frame, text="Markers to Plot", padding="10")
        markers_to_plot_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        markers_to_plot_frame.grid_columnconfigure(0, weight=1) # Make column expandable

        self.include_tv_markers_var = tk.BooleanVar(value=True) # Default to True
        ttk.Checkbutton(markers_to_plot_frame, text="Include TV Band Markers", variable=self.include_tv_markers_var,
                        command=self._on_multi_file_marker_checkbox_changed).grid(row=0, column=0, padx=5, pady=2, sticky="w")

        self.include_gov_markers_var = tk.BooleanVar(value=False) # Default to False
        ttk.Checkbutton(markers_to_plot_frame, text="Include Government Band Markers", variable=self.include_gov_markers_var,
                        command=self._on_multi_file_marker_checkbox_changed).grid(row=1, column=0, padx=5, pady=2, sticky="w")

        self.include_markers_var = tk.BooleanVar(value=True) # General Markers for multi-file plots
        ttk.Checkbutton(markers_to_plot_frame, text="Include Markers", variable=self.include_markers_var,
                        command=self._on_multi_file_marker_checkbox_changed).grid(row=2, column=0, padx=5, pady=2, sticky="w")

        self.include_intermod_markers_var = tk.BooleanVar(value=False) # Intermodulations
        ttk.Checkbutton(markers_to_plot_frame, text="Include Intermodulations", variable=self.include_intermod_markers_var,
                        command=self._on_multi_file_marker_checkbox_changed).grid(row=3, column=0, padx=5, pady=2, sticky="w")


        # Make Averages Frame
        make_averages_frame = ttk.LabelFrame(self.averaging_folder_frame, text="Make Averages", padding="10")
        make_averages_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        make_averages_frame.grid_columnconfigure(0, weight=1) # Make column expandable

        self.generate_csv_button = ttk.Button(make_averages_frame, text="Generate CSV of Selected Series of Scan", command=self._generate_csv_selected_series)
        self.generate_csv_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.generate_csv_button.config(state=tk.DISABLED) # Disable until group is selected

        self.open_applied_math_folder_button = ttk.Button(make_averages_frame, text="Open Folder of Applied Math", command=self._open_applied_math_folder)
        self.open_applied_math_folder_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.open_applied_math_folder_button.config(state=tk.DISABLED) # Disable until a folder is created

        self.generate_plot_averages_button = ttk.Button(make_averages_frame, text="Generate Plot of Averages", command=self._generate_selected_average_plot)
        self.generate_plot_averages_button.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
        self.generate_plot_averages_button.config(state=tk.DISABLED) # Disable until group is selected

        self.generate_plot_averages_with_scan_button = ttk.Button(make_averages_frame, text="Generate Plot of Averages with Scan", command=self._generate_plot_averages_with_scan)
        self.generate_plot_averages_with_scan_button.grid(row=3, column=0, padx=5, pady=5, sticky="ew")
        self.generate_plot_averages_with_scan_button.config(state=tk.DISABLED) # Disable until group is selected

        self.generate_plot_scans_over_time_button = ttk.Button(make_averages_frame, text="Generate Plot of Scans Over Time", command=self._generate_plot_scans_over_time)
        self.generate_plot_scans_over_time_button.grid(row=4, column=0, padx=5, pady=5, sticky="ew")
        self.generate_plot_scans_over_time_button.config(state=tk.DISABLED) # Disable until group is selected


        # Configure column weights for resizing
        self.grid_columnconfigure(0, weight=1)
        plotting_options_frame.grid_columnconfigure(0, weight=1)
        self.averaging_folder_frame.grid_columnconfigure(0, weight=1)
        self.averaging_folder_frame.grid_columnconfigure(1, weight=1) # For the two columns inside
        self.discovered_series_frame.grid_columnconfigure(0, weight=1)
        self.dynamic_avg_buttons_frame.grid_columnconfigure(0, weight=1) # Ensure dynamic buttons expand
        make_averages_frame.grid_columnconfigure(0, weight=1)


        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _on_avg_type_checkbox_changed(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        selected_types = [avg_type for avg_type, var in self.avg_type_vars.items() if var.get()]
        debug_print(f"Checkbox changed. Currently selected average types: {selected_types}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        self.console_print_func(f"Selected average types: {', '.join(selected_types) if selected_types else 'None'}")

    def _on_create_html_checkbox_changed(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        state = "Enabled" if self.create_html_var.get() else "Disabled"
        debug_print(f"Create HTML checkbox changed. State: {state}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        self.console_print_func(f"Create HTML: {state}")


    def _on_scan_marker_checkbox_changed(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        selected_markers = []
        if self.include_scan_tv_markers_var.get():
            selected_markers.append("TV Band Markers")
        if self.include_scan_gov_markers_var.get():
            selected_markers.append("Government Band Markers")
        if self.include_scan_markers_var.get():
            selected_markers.append("General Markers")
        if self.include_scan_intermod_markers_var.get():
            selected_markers.append("Intermodulations")

        debug_print(f"SCAN Plotting Options - Selected markers: {selected_markers}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        self.console_print_func(f"SCAN Plotting Options - Markers: {', '.join(selected_markers) if selected_markers else 'None'}")

    def _on_multi_file_marker_checkbox_changed(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        selected_markers = []
        if self.include_tv_markers_var.get():
            selected_markers.append("TV Band Markers")
        if self.include_gov_markers_var.get():
            selected_markers.append("Government Band Markers")
        if self.include_markers_var.get():
            selected_markers.append("General Markers")
        if self.include_intermod_markers_var.get():
            selected_markers.append("Intermodulations")

        debug_print(f"Multi-File Plotting - Selected markers: {selected_markers}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        self.console_print_func(f"Multi-File Plotting - Markers: {', '.join(selected_markers) if selected_markers else 'None'}")


    def _plot_single_scan(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        file_path = filedialog.askopenfilename(
            title="Select a CSV file to plot",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            self.console_print_func("File selection cancelled. No single scan plot generated.")
            debug_print("File selection cancelled for single plot.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        try:
            df = pd.read_csv(file_path, header=None, names=['Frequency', 'Amplitude'])
            scan_name = os.path.splitext(os.path.basename(file_path))[0]
            self.console_print_func(f"Plotting single scan from: {scan_name}")
            debug_print(f"Plotting single scan for: {scan_name}", file=current_file, function=current_function, console_print_func=self.console_print_func)

            output_dir = self.app_instance.config.get('Paths', 'output_directory')
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                debug_print(f"Created output directory: {output_dir}", file=current_file, function=current_function, console_print_func=self.console_print_func)

            html_filename = os.path.join(output_dir, f"{scan_name.replace(' ', '_')}_single_scan_plot.html")
            debug_print(f"Output HTML filename: {html_filename}", file=current_file, function=current_function, console_print_func=self.console_print_func)

            fig, plot_html_path_return = plot_single_scan_data(
                df,
                f"Single Scan: {scan_name}",
                include_tv_markers=self.include_scan_tv_markers_var.get(),
                include_gov_markers=self.include_scan_gov_markers_var.get(),
                include_markers=self.include_scan_markers_var.get(),
                output_html_path=html_filename if self.create_html_var.get() else None, # Only create HTML if checkbox is checked
                console_print_func=self.console_print_func
            )

            if fig:
                self.current_plot_file = plot_html_path_return
                if self.create_html_var.get():
                    self.console_print_func(f"✅ Single scan plot saved to: {self.current_plot_file}")
                    debug_print(f"Plot saved successfully to: {self.current_plot_file}", file=current_file, function=current_function, console_print_func=self.console_print_func)
                else:
                    self.console_print_func("✅ Single scan plot data processed (HTML not saved as per setting).")
                    debug_print("Plot data processed, HTML not saved.", file=current_file, function=current_function, console_print_func=self.console_print_func)

                if self.open_html_after_complete_var.get() and self.create_html_var.get() and plot_html_path_return:
                    self.console_print_func(f"Opening plot in browser: {self.current_plot_file}")
                    webbrowser.open_new_tab(self.current_plot_file)
                    debug_print(f"Plot opened in browser: {self.current_plot_file}", file=current_file, function=current_function, console_print_func=self.console_print_func)
                elif self.open_html_after_complete_var.get() and not self.create_html_var.get():
                    self.console_print_func("HTML plot not opened because 'Create HTML' is unchecked.")
                    debug_print("HTML not opened as 'Create HTML' is unchecked.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            else:
                self.console_print_func("🚫 Plotly figure was not generated for single scan.")
                debug_print("Plotly figure not generated for single scan.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        except Exception as e:
            self.console_print_func(f"❌ Error plotting single scan: {e}")
            debug_print(f"Error plotting single scan: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _plot_current_cycle_average(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if not self.app_instance.collected_scans_dataframes:
            self.console_print_func("No collected scan dataframes to average.")
            debug_print("No collected scan dataframes for current cycle average.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            debug_print(f"Exiting {current_function} (no data)", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        output_dir = self.app_instance.config.get('Paths', 'output_directory')
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            debug_print(f"Created output directory: {output_dir}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # Get selected average types for current cycle plot
        selected_avg_types = [
            avg_type for avg_type, var in self.avg_type_vars.items() if var.get()
        ]
        if not selected_avg_types:
            self.console_print_func("Warning: No average type selected for current cycle plot. Please select at least one type.")
            debug_print(f"Exiting {current_function} (no selected_avg_types for current cycle)", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        debug_print(f"Calling generate_current_cycle_average_csv_and_plot with {len(self.app_instance.collected_scans_dataframes)} dataframes and selected types: {selected_avg_types}.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        # Pass all relevant variables to the averaging utility
        fig, plot_html_path_return = generate_current_cycle_average_csv_and_plot(
            self.app_instance.collected_scans_dataframes,
            output_dir,
            include_tv_markers_var=self.include_scan_tv_markers_var,
            include_gov_markers_var=self.include_scan_gov_markers_var,
            include_markers_var=self.include_scan_markers_var,
            open_html_after_complete_var=self.open_html_after_complete_var, # Pass the Tkinter BooleanVar directly
            create_html_var=self.create_html_var, # Pass the new create HTML checkbox state
            console_print_func=self.console_print_func,
            selected_avg_types=selected_avg_types, # Pass the selected average types
            scan_rbw_hz=self.app_instance.config.getfloat('Instrument', 'rbw_hz') if self.app_instance.config.has_option('Instrument', 'rbw_hz') else None
        )

        if fig:
            self.current_plot_file = plot_html_path_return
            if self.create_html_var.get():
                self.console_print_func(f"✅ Current cycle averaged plot saved to: {self.current_plot_file}")
                debug_print(f"Current cycle averaged plot saved to: {self.current_plot_file}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            else:
                self.console_print_func("✅ Current cycle averaged plot data processed (HTML not saved as per setting).")
                debug_print("Plot data processed, HTML not saved.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            # The function itself opens the browser if open_html_after_complete_var is True and create_html_var is True
        else:
            self.console_print_func("🚫 Plotly figure was not generated for current cycle averaged data.")
            debug_print("Plotly figure not generated for current cycle averaged data.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _open_last_plot(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if self.current_plot_file and os.path.exists(self.current_plot_file):
            self.console_print_func(f"Opening last plot: {self.current_plot_file}")
            webbrowser.open_new_tab(self.current_plot_file)
            debug_print(f"Opened last plot: {self.current_plot_file}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.console_print_func("Error: No plot available or file not found. Please generate a plot first.")
            debug_print("No plot available or file not found.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _open_folder_for_averaging(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        folder_path = filedialog.askdirectory(initialdir=self.last_opened_folder)
        debug_print(f"Selected folder_path: {folder_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        if folder_path:
            self.last_opened_folder = folder_path
            self.console_print_func(f"Selected folder for averaging: {folder_path}")
            self._find_and_group_csv_files(folder_path)
            # Enable relevant buttons after a folder is selected and groups are found
            self.generate_csv_button.config(state=tk.NORMAL)
            self.generate_plot_averages_button.config(state=tk.NORMAL)
            self.generate_plot_averages_with_scan_button.config(state=tk.NORMAL)
            self.generate_plot_scans_over_time_button.config(state=tk.NORMAL)
            debug_print("Averaging buttons enabled (folder selected).", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.console_print_func("Folder selection cancelled.")
            # Disable buttons if folder selection is cancelled
            self.generate_csv_button.config(state=tk.DISABLED)
            self.generate_plot_averages_button.config(state=tk.DISABLED)
            self.generate_plot_averages_with_scan_button.config(state=tk.DISABLED)
            self.generate_plot_scans_over_time_button.config(state=tk.DISABLED)
            debug_print("Folder selection cancelled. Averaging buttons disabled.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _find_and_group_csv_files(self, folder_path):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function} with folder_path: {folder_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
        debug_print(f"Found CSV files: {csv_files}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        if not csv_files:
            self.console_print_func("No CSV files found in the selected folder.")
            self._clear_dynamic_buttons()
            debug_print(f"Exiting {current_function} (no CSVs)", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        file_groups = {}
        for filename in csv_files:
            base_name = os.path.splitext(filename)[0]
            # Updated regex for grouping: be more flexible with the ending part before date/time
            # It tries to capture the main prefix before any RBW/HOLD/Offset/timestamp.
            match = re.match(r"([^\d_ -]+(?:[_ -][^\d_ -]+)*?)_RBW\d+K?_HOLD\d+.*", base_name)
            prefix = base_name # Default to full base name if no clear pattern
            if match:
                prefix = match.group(1).strip() # Get the part before RBW/HOLD
                # Refine prefix: remove trailing underscores/hyphens if they were part of the non-digit group
                prefix = re.sub(r"[_ -]+$", "", prefix)
            else:
                # Fallback if the more specific pattern doesn't match (e.g., for INTERMOD.csv or simpler names)
                # Try to split by common delimiters if no complex pattern is found
                if '_' in base_name:
                    prefix = base_name.split('_')[0]
                elif '-' in base_name:
                    prefix = base_name.split('-')[0]
                else:
                    prefix = base_name # Use full name if no common delimiters

            if not prefix: # Ensure prefix is not empty
                prefix = base_name # Fallback to full base name

            if prefix not in file_groups:
                file_groups[prefix] = []
            file_groups[prefix].append(os.path.join(folder_path, filename))
        debug_print(f"Grouped CSV files: {file_groups}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self._clear_dynamic_buttons() # Clear any previous buttons

        if not file_groups:
            self.console_print_func("No identifiable groups of CSV files found.")
            debug_print(f"Exiting {current_function} (no file groups)", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        self.console_print_func(f"Found {len(file_groups)} groups of similar CSV files.")

        self.grouped_csv_files = file_groups
        self.selected_group_prefix = None

        row_start = 0
        for i, (prefix, files) in enumerate(file_groups.items()):
            group_text = f"Group '{prefix}' ({len(files)} files)"
            btn = ttk.Button(self.dynamic_avg_buttons_frame, text=group_text,
                             command=lambda p=prefix: self._select_group_for_plotting(p))
            btn.grid(row=row_start + i, column=0, padx=5, pady=2, sticky="ew")
            btn.config(style='Orange.TButton')

        try:
            style = ttk.Style()
            style.configure('Orange.TButton', background='orange', foreground='black')
        except Exception as e:
            debug_print(f"Could not apply orange style: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _select_group_for_plotting(self, prefix):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function} with prefix: {prefix}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self.selected_group_prefix = prefix
        self.console_print_func(f"Selected group for plotting: '{prefix}'")
        debug_print(f"Selected group for plotting: '{prefix}'", file=current_file, function=current_function, console_print_func=self.console_print_func)

        for widget in self.dynamic_avg_buttons_frame.winfo_children():
            if isinstance(widget, ttk.Button):
                if widget.cget("text").startswith(f"Group '{prefix}'"):
                    widget.config(relief="sunken", style='SelectedOrange.TButton')
                    debug_print(f"Highlighted button for group: {prefix}", file=current_file, function=current_function, console_print_func=self.console_print_func)
                else:
                    widget.config(relief="raised", style='Orange.TButton')
        try:
            style = ttk.Style()
            style.configure('SelectedOrange.TButton', background='darkorange', foreground='white')
        except Exception as e:
            debug_print(f"Could not apply selected orange style: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _clear_dynamic_buttons(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        for widget in self.dynamic_avg_buttons_frame.winfo_children():
            widget.destroy()
        debug_print(f"Cleared dynamic buttons.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _generate_csv_selected_series(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if not hasattr(self, 'grouped_csv_files') or not self.grouped_csv_files:
            self.console_print_func("Warning: No data. Please select a folder and identify CSV file groups first.")
            debug_print(f"Exiting {current_function} (no grouped_csv_files)", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        if not self.selected_group_prefix:
            self.console_print_func("Warning: No group selected. Please click on one of the group buttons to select files for averaging.")
            debug_print(f"Exiting {current_function} (no selected_group_prefix)", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        files_to_average = self.grouped_csv_files[self.selected_group_prefix]
        debug_print(f"Files to average for selected group '{self.selected_group_prefix}': {files_to_average}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        if not files_to_average:
            self.console_print_func("Error: No files found for the selected group.")
            debug_print(f"Exiting {current_function} (files_to_average is empty)", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        selected_avg_types = [
            avg_type for avg_type, var in self.avg_type_vars.items() if var.get()
        ]
        if not selected_avg_types:
            self.console_print_func("Warning: No average type selected. Please select at least one type of average to generate CSVs (e.g., Average, Median).")
            debug_print(f"Exiting {current_function} (no selected_avg_types)", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        output_dir_base = self.last_opened_folder # Use the folder where files were opened as base for new subfolder

        # Call the new average_scan function
        aggregated_df, output_folder_path = average_scan(
            file_paths=files_to_average,
            selected_avg_types=selected_avg_types,
            plot_title_prefix=self.selected_group_prefix,
            output_html_path_base=output_dir_base,
            console_print_func=self.console_print_func
        )

        if output_folder_path:
            self.last_applied_math_folder = output_folder_path
            self.console_print_func(f"CSV files generated in: {output_folder_path}")
            self.open_applied_math_folder_button.config(state=tk.NORMAL) # Enable the button
            self.console_print_func("🎉 CSV generation complete! 🎉")
        else:
            self.console_print_func("🚫 CSV generation failed.")
            self.open_applied_math_folder_button.config(state=tk.DISABLED)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _open_applied_math_folder(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if self.last_applied_math_folder and os.path.exists(self.last_applied_math_folder):
            self.console_print_func(f"Opening folder: {self.last_applied_math_folder}")
            try:
                if platform.system() == "Windows":
                    os.startfile(self.last_applied_math_folder)
                elif platform.system() == "Darwin": # macOS
                    subprocess.Popen(["open", self.last_applied_math_folder])
                else: # Linux and other Unix-like systems
                    subprocess.Popen(["xdg-open", self.last_applied_math_folder])
            except Exception as e:
                self.console_print_func(f"❌ Error opening folder: {e}")
                debug_print(f"Error opening folder {self.last_applied_math_folder}: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.console_print_func("No folder of applied math available. Please generate CSVs first.")
            debug_print("No applied math folder to open.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _generate_selected_average_plot(self): # Renamed this to match the button command
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if not hasattr(self, 'grouped_csv_files') or not self.grouped_csv_files:
            self.console_print_func("Warning: No data. Please select a folder and identify CSV file groups first.")
            debug_print(f"Exiting {current_function} (no grouped_csv_files)", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        if not self.selected_group_prefix:
            self.console_print_func("Warning: No group selected. Please click on one of the group buttons to select files for averaging.")
            debug_print(f"Exiting {current_function} (no selected_group_prefix)", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        files_to_average = self.grouped_csv_files[self.selected_group_prefix]
        debug_print(f"Files to average for selected group '{self.selected_group_prefix}': {files_to_average}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        if not files_to_average:
            self.console_print_func("Error: No files found for the selected group.")
            debug_print(f"Exiting {current_function} (files_to_average is empty)", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        selected_avg_types = [
            avg_type for avg_type, var in self.avg_type_vars.items() if var.get()
        ]
        debug_print(f"Selected average types BEFORE check: {selected_avg_types}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if not selected_avg_types:
            self.console_print_func("Warning: No average type selected. Please select at least one type of average to plot (e.g., Average, Median).")
            debug_print(f"Exiting {current_function} (no selected_avg_types)", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        # Use the last_opened_folder as the base output directory for multi-file averages
        output_dir = self.last_opened_folder
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            debug_print(f"Ensured base output directory exists: {output_dir}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        plot_base_name = f"{self.selected_group_prefix}_averaged_plot"
        html_filename = os.path.join(output_dir, f"{plot_base_name}.html") # This might be overwritten by the averaging_utils function
        debug_print(f"Output HTML filename (initial): {html_filename}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self.console_print_func(f"Generating plot for group '{self.selected_group_prefix}' with types: {', '.join(selected_avg_types)}")
        debug_print(f"Calling generate_multi_file_average_and_plot with files: {files_to_average} and types: {selected_avg_types}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        fig, plot_html_path_return = generate_multi_file_average_and_plot(
            file_paths=files_to_average,
            selected_avg_types=selected_avg_types,
            plot_title_prefix=self.selected_group_prefix,
            include_tv_markers=self.include_tv_markers_var.get(),
            include_gov_markers=self.include_gov_markers_var.get(),
            include_markers=self.include_markers_var.get(), # Pass the general markers variable
            output_html_path_base=output_dir, # Pass the base output directory
            open_html_after_complete=self.open_html_after_complete_var.get(),
            console_print_func=self.console_print_func
        )

        if fig:
            self.current_plot_file = plot_html_path_return
            self.console_print_func(f"✅ Multi-file averaged plot saved to: {self.current_plot_file}")
            debug_print(f"Multi-file averaged plot saved to: {self.current_plot_file}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.console_print_func("🚫 Plotly figure was not generated for multi-file averaged data.")
            debug_print("Plotly figure not generated for multi-file averaged data.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _generate_plot_averages_with_scan(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        self.console_print_func("Function to generate plot of averages with scan is not yet implemented.")
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _generate_plot_scans_over_time(self):
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        self.console_print_func("Function to generate plot of scans over time is not yet implemented.")
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _on_tab_selected(self, event):
        """
        Callback for when this tab is selected.
        This can be used to refresh data or update UI elements specific to this tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)

        debug_print("Plotting Tab selected.", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if self.app_instance.collected_scans_dataframes:
            self.plot_button.config(state=tk.NORMAL)
            self.plot_average_button.config(state=tk.NORMAL)
            debug_print("Plotting buttons enabled (data available).", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.plot_button.config(state=tk.DISABLED)
            self.plot_average_button.config(state=tk.DISABLED)
            debug_print("Plotting buttons disabled (no data available).", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # Also update the state of the multi-file averaging buttons based on whether a folder was selected
        if self.last_opened_folder and os.path.exists(self.last_opened_folder) and hasattr(self, 'grouped_csv_files') and self.grouped_csv_files:
            self.generate_csv_button.config(state=tk.NORMAL)
            self.generate_plot_averages_button.config(state=tk.NORMAL)
            self.generate_plot_averages_with_scan_button.config(state=tk.NORMAL)
            self.generate_plot_scans_over_time_button.config(state=tk.NORMAL)
            debug_print("Multi-file averaging buttons enabled (folder and groups exist).", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.generate_csv_button.config(state=tk.DISABLED)
            self.generate_plot_averages_button.config(state=tk.DISABLED)
            self.generate_plot_averages_with_scan_button.config(state=tk.DISABLED)
            self.generate_plot_scans_over_time_button.config(state=tk.DISABLED)
            debug_print("Multi-file averaging buttons disabled (no folder or groups).", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # Enable/Disable "Open Folder of Applied Math" button based on whether a folder was previously created
        if self.last_applied_math_folder and os.path.exists(self.last_applied_math_folder):
            self.open_applied_math_folder_button.config(state=tk.NORMAL)
            debug_print("Open Applied Math Folder button enabled.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.open_applied_math_folder_button.config(state=tk.DISABLED)
            debug_print("Open Applied Math Folder button disabled.", file=current_file, function=current_function, console_print_func=self.console_print_func)


        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=self.console_print_func)
