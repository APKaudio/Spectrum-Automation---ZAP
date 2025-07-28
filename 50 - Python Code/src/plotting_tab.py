# src/plotting_tab.py
import tkinter as tk
from tkinter import ttk
import inspect # Import inspect for debug_print

# Import plotting logic from utils
from utils.averaging_utils import generate_current_cycle_average_csv_and_plot as generate_average_plot_logic
from utils.instrument_control import debug_print # Import debug_print

class PlottingTab(ttk.Frame):
    """
    A Tkinter Frame that serves as a tab for plotting configuration and actions.
    It includes options for including various markers and generating plots.
    """
    def __init__(self, master=None, app_instance=None, console_print_func=None, **kwargs):
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self.console_print_func = console_print_func if console_print_func else print # Use provided func or default print
        self._create_widgets()

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Plotting tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Creating PlottingTab widgets.", file=current_file, function=current_function, console_print_func=self.console_print_func)

        # Configure grid for this tab
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1) # Row for settings
        self.grid_rowconfigure(1, weight=0) # Row for plot generation button

        # Plotting Settings Frame
        plotting_settings_frame = ttk.LabelFrame(self, text="Plotting Settings", style='Dark.TLabelframe')
        plotting_settings_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        plotting_settings_frame.grid_columnconfigure(0, weight=1) # Checkboxes expand

        # Checkboxes for plot options
        ttk.Checkbutton(plotting_settings_frame, text="Include Government Markers", variable=self.app_instance.include_gov_markers_var, style='TCheckbutton').grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(plotting_settings_frame, text="Include TV Channel Markers", variable=self.app_instance.include_tv_markers_var, style='TCheckbutton').grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(plotting_settings_frame, text="Include Extracted Markers", variable=self.app_instance.include_markers_var, style='TCheckbutton').grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Checkbutton(plotting_settings_frame, text="Open HTML Plot After Generation", variable=self.app_instance.open_html_after_complete_var, style='TCheckbutton').grid(row=3, column=0, sticky="w", padx=5, pady=2)

        # Plot Generation Button
        self.plot_button = ttk.Button(self, text="Generate Averaged Plot", 
                                      command=lambda: generate_average_plot_logic(
                                          self.app_instance.collected_scans_dataframes,
                                          self.app_instance.scan_name_var,
                                          self.app_instance.output_folder_var,
                                          self.app_instance.open_html_after_complete_var,
                                          self.app_instance.include_tv_markers_var,
                                          self.app_instance.include_gov_markers_var,
                                          self.app_instance.include_markers_var,
                                          self.console_print_func # Pass console_print_func
                                      ),
                                      state=tk.DISABLED, style='Blue.TButton')
        self.plot_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        debug_print("PlottingTab widgets created.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _on_tab_selected(self, event):
        """
        Callback for when this tab is selected.
        This can be used to refresh data or update UI elements specific to this tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Plotting Tab selected.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        # Example: You might want to update the state of the plot_button here
        # based on whether there's data available to plot in app_instance.collected_scans_dataframes
        if self.app_instance.collected_scans_dataframes:
            self.plot_button.config(state=tk.NORMAL)
        else:
            self.plot_button.config(state=tk.DISABLED)

