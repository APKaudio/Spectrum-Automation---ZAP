# src/plotting_tab.py
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import os
import pandas as pd
import inspect
import webbrowser # For opening HTML plot in browser

from utils.instrument_control import debug_print
from src.plot_logic import plot_single_scan_data # Import only plot_single_scan_data
# Import generate_current_cycle_average_csv_and_plot from its correct location
from process_math.averaging_utils import generate_current_cycle_average_csv_and_plot

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
        self.console_print_func = console_print_func if console_print_func else print # Use provided func or default print
        self.last_plot_path = None # To store the path of the last generated plot

        self._create_widgets()

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Plotting tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Creating PlottingTab widgets...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Plotting Controls
        plot_control_frame = ttk.LabelFrame(self, text="Plotting Controls", style='Dark.TLabelframe')
        plot_control_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        plot_control_frame.grid_columnconfigure(0, weight=1)
        plot_control_frame.grid_columnconfigure(1, weight=1)

        self.plot_button = ttk.Button(plot_control_frame, text="Plot Last Scan", command=self._plot_last_scan, style='Blue.TButton')
        self.plot_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.plot_average_button = ttk.Button(plot_control_frame, text="Plot Average Scan", command=self._plot_average_scan, style='Blue.TButton')
        self.plot_average_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.open_plot_button = ttk.Button(plot_control_frame, text="Open Last Plot in Browser", command=self._open_last_plot_in_browser, state=tk.DISABLED, style='Blue.TButton')
        self.open_plot_button.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        # Plotting Options (Moved from main_app)
        plotting_options_frame = ttk.LabelFrame(self, text="Plotting Options", style='Dark.TLabelframe')
        plotting_options_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        plotting_options_frame.grid_columnconfigure(0, weight=1)

        ttk.Checkbutton(plotting_options_frame, text="Include Government Markers", variable=self.app_instance.include_gov_markers_var, style='TCheckbutton').grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(plotting_options_frame, text="Include TV Markers", variable=self.app_instance.include_tv_markers_var, style='TCheckbutton').grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(plotting_options_frame, text="Include Custom Markers (from CSV)", variable=self.app_instance.include_markers_var, style='TCheckbutton').grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(plotting_options_frame, text="Open HTML Plot After Scan Complete", variable=self.app_instance.open_html_after_complete_var, style='TCheckbutton').grid(row=3, column=0, sticky="w", padx=5, pady=2)

        debug_print("PlottingTab widgets created.", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _plot_last_scan(self):
        """
        Plots the data from the last completed scan.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Attempting to plot last scan...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if not self.app_instance.collected_scans_dataframes:
            self.console_print_func("⚠️ Warning: No scan data available to plot. Please run a scan first.")
            debug_print("No scan data for plotting.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        last_df = self.app_instance.collected_scans_dataframes[-1]
        output_dir = self.app_instance.output_folder_var.get()
        scan_name = self.app_instance.scan_name_var.get()

        try:
            plot_path = plot_single_scan_data(
                last_df,
                output_dir,
                scan_name,
                self.app_instance.include_gov_markers_var.get(),
                self.app_instance.include_tv_markers_var.get(),
                self.app_instance.include_markers_var.get(),
                self.app_instance.markers_display_tab.headers, # Pass headers from MarkersDisplayTab
                self.app_instance.markers_display_tab.rows,    # Pass rows from MarkersDisplayTab
                self.console_print_func
            )
            if plot_path:
                self.last_plot_path = plot_path
                self.open_plot_button.config(state=tk.NORMAL)
                self.console_print_func(f"✅ Last scan plotted to: {plot_path}")
                debug_print(f"Last scan plotted to: {plot_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            else:
                self.console_print_func("❌ Error: Failed to generate plot for the last scan.")
                debug_print("Failed to generate plot for last scan.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        except Exception as e:
            self.console_print_func(f"❌ Error plotting last scan: {e}")
            debug_print(f"Error plotting last scan: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _plot_average_scan(self):
        """
        Generates and plots the average of all collected scan data.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Attempting to plot average scan...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if not self.app_instance.collected_scans_dataframes:
            self.console_print_func("⚠️ Warning: No scan data available to average plot. Please run a scan first.")
            debug_print("No scan data for average plotting.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        output_dir = self.app_instance.output_folder_var.get()
        scan_name = self.app_instance.scan_name_var.get()

        try:
            plot_path = generate_current_cycle_average_csv_and_plot(
                self.app_instance.collected_scans_dataframes,
                output_dir,
                scan_name,
                self.app_instance.include_gov_markers_var.get(),
                self.app_instance.include_tv_markers_var.get(),
                self.app_instance.include_markers_var.get(),
                self.app_instance.markers_display_tab.headers, # Pass headers from MarkersDisplayTab
                self.app_instance.markers_display_tab.rows,    # Pass rows from MarkersDisplayTab
                self.console_print_func
            )
            if plot_path:
                self.last_plot_path = plot_path
                self.open_plot_button.config(state=tk.NORMAL)
                self.console_print_func(f"✅ Average scan plotted to: {plot_path}")
                debug_print(f"Average scan plotted to: {plot_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            else:
                self.console_print_func("❌ Error: Failed to generate average plot.")
                debug_print("Failed to generate average plot.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        except Exception as e:
            self.console_print_func(f"❌ Error plotting average scan: {e}")
            debug_print(f"Error plotting average scan: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)


    def _open_last_plot_in_browser(self):
        """
        Opens the last generated HTML plot in the default web browser.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Attempting to open last plot in browser...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if self.last_plot_path and os.path.exists(self.last_plot_path):
            try:
                webbrowser.open_new_tab(f"file:///{os.path.abspath(self.last_plot_path)}")
                self.console_print_func(f"✅ Opened plot: {self.last_plot_path}")
                debug_print(f"Opened plot: {self.last_plot_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            except Exception as e:
                self.console_print_func(f"❌ Error opening plot in browser: {e}")
                debug_print(f"Error opening plot in browser: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.console_print_func("⚠️ Warning: No plot available or file not found. Please generate a plot first.")
            debug_print("No plot available or file not found.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _on_tab_selected(self, event):
        """
        Callback for when this tab is selected.
        This can be used to refresh data or update UI elements specific to this tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Plotting Tab selected.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        # Enable plot button if there's data to plot
        if self.app_instance.collected_scans_dataframes:
            self.plot_button.config(state=tk.NORMAL)
            self.plot_average_button.config(state=tk.NORMAL)
        else:
            self.plot_button.config(state=tk.DISABLED)
            self.plot_average_button.config(state=tk.DISABLED)
        # Keep open plot button state as is, it's enabled when a plot is generated

