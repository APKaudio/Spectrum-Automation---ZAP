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
    def __init__(self, master=None, app_instance=None, **kwargs):
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self._create_widgets()

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Plotting tab.
        """
        # Configure grid for this tab
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1) # Row for settings
        self.grid_rowconfigure(1, weight=1) # Row for plot generation button

        # Plotting Settings Frame
        plotting_settings_frame = ttk.LabelFrame(self, text="Plotting Settings")
        plotting_settings_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        plotting_settings_frame.grid_columnconfigure(0, weight=1)

        # Checkboxes for plotting markers
        ttk.Checkbutton(plotting_settings_frame, text="Include Gov Markers", variable=self.app_instance.include_gov_markers_var,
                        command=lambda: self.app_instance._on_setting_change(self.app_instance.include_gov_markers_var, 'last_include_gov_markers')).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(plotting_settings_frame, text="Include TV Markers", variable=self.app_instance.include_tv_markers_var,
                        command=lambda: self.app_instance._on_setting_change(self.app_instance.include_tv_markers_var, 'last_include_tv_markers')).grid(row=1, column=0, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(plotting_settings_frame, text="Include Custom Markers (MARKERS.CSV)", variable=self.app_instance.include_markers_var,
                        command=lambda: self.app_instance._on_setting_change(self.app_instance.include_markers_var, 'last_include_markers')).grid(row=2, column=0, padx=5, pady=2, sticky="w")

        # Checkbox for opening HTML after complete
        ttk.Checkbutton(plotting_settings_frame, text="Open HTML After Complete", variable=self.app_instance.open_html_after_complete_var,
                        command=lambda: self.app_instance._on_setting_change(self.app_instance.open_html_after_complete_var, 'last_open_html_after_complete')).grid(row=3, column=0, padx=5, pady=2, sticky="w")

        # Plot Generation Button
        # Store a reference to this button for external control (e.g., enabling/disabling)
        self.plot_button = ttk.Button(self, text="Generate Average Plot",
                                      command=lambda: generate_average_plot_logic(
                                          self.app_instance.collected_scans_dataframes,
                                          self.app_instance.scan_name_var,
                                          self.app_instance.output_folder_var,
                                          self.app_instance.include_tv_markers_var,
                                          self.app_instance.include_gov_markers_var,
                                          self.app_instance.include_markers_var,
                                          self.app_instance.open_html_after_complete_var
                                      ),
                                      state=tk.DISABLED, style='Blue.TButton')
        self.plot_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        debug_print("PlottingTab widgets created.", file=__file__, function=inspect.currentframe().f_code.co_name)

    def _on_tab_selected(self, event):
        """
        Callback for when this tab is selected.
        This can be used to refresh data or update UI elements specific to this tab.
        """
        debug_print("Plotting Tab selected.", file=__file__, function=inspect.currentframe().f_code.co_name)
        # Example: You might want to update the state of the plot_button here
        # based on whether there's data available to plot in app_instance.collected_scans_dataframes
        if self.app_instance.collected_scans_dataframes:
            self.plot_button.config(state=tk.NORMAL)
        else:
            self.plot_button.config(state=tk.DISABLED)

