# src/instrument_tab.py
import tkinter as tk
from tkinter import ttk
import inspect

# Import instrument control logic functions
from src.instrument_logic import (
    populate_resources_logic, connect_instrument_logic, disconnect_instrument_logic,
    apply_settings_logic, query_current_instrument_settings_logic
)
# Removed: from src.scan_logic import update_connection_status_logic # This import is no longer needed here
from utils.instrument_control import debug_print, set_debug_mode, log_visa_command, query_safe # Import debug control functions and query_safe
from utils.frequency_bands import MHZ_TO_HZ # Import for display conversion

class InstrumentTab(ttk.Frame):
    """
    A Tkinter Frame that provides functionality for connecting to, disconnecting from,
    and querying the spectrum analyzer. It also displays available VISA resources
    and current instrument settings.
    """
    def __init__(self, master=None, app_instance=None, console_print_func=None, **kwargs):
        """
        Initializes the InstrumentTab.

        Inputs:
            master (tk.Widget): The parent widget (the ttk.Notebook).
            app_instance (App): The main application instance, used for accessing
                                shared state like Tkinter variables and console print function.
            console_print_func (function): Function to print messages to the GUI console.
            **kwargs: Arbitrary keyword arguments for Tkinter Frame.
        """
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self.console_print_func = console_print_func if console_print_func else print

        # Use the shared variables from app_instance
        self.selected_resource = self.app_instance.selected_resource
        self.resource_names = self.app_instance.resource_names

        # Tkinter variables for displaying current instrument settings
        self.current_center_freq_var = tk.StringVar(self)
        self.current_span_var = tk.StringVar(self)
        self.current_rbw_var = tk.StringVar(self)
        self.current_ref_level_var = tk.StringVar(self)
        self.current_freq_shift_var = tk.StringVar(self)
        self.current_max_hold_var = tk.StringVar(self)
        self.current_high_sensitivity_var = tk.StringVar(self)


        self._create_widgets()

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Instrument Connection tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Creating InstrumentTab widgets...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1) # For buttons

        # Frame for resource selection and connection buttons
        connection_frame = ttk.LabelFrame(self, text="Instrument Connection", padding="10 10 10 10", style='Dark.TLabelframe')
        connection_frame.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        connection_frame.grid_columnconfigure(0, weight=1)
        connection_frame.grid_columnconfigure(1, weight=2) # Resource dropdown
        connection_frame.grid_columnconfigure(2, weight=1) # Populate button

        ttk.Label(connection_frame, text="VISA Resource:", style='TLabel').grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.resource_dropdown = ttk.OptionMenu(connection_frame, self.selected_resource, "", *self.resource_names.get().split(),
                                                command=lambda x: debug_print(f"Selected resource: {x}", file=current_file, function=current_function, console_print_func=self.console_print_func))
        self.resource_dropdown.config(width=40)
        self.resource_dropdown.grid(row=0, column=1, padx=5, pady=2, sticky="ew")

        self.populate_button = ttk.Button(connection_frame, text="Populate Resources", command=self._populate_resources)
        self.populate_button = ttk.Button(connection_frame, text="Populate Resources", command=self._populate_resources)
        self.populate_button.grid(row=0, column=2, padx=5, pady=2, sticky="ew")

        self.connect_button = ttk.Button(connection_frame, text="Connect", command=self._connect_instrument)
        self.connect_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        self.disconnect_button = ttk.Button(connection_frame, text="Disconnect", command=self._disconnect_instrument, state=tk.DISABLED)
        self.disconnect_button.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.apply_settings_button = ttk.Button(connection_frame, text="Apply Settings", command=self._apply_settings, state=tk.DISABLED)
        self.apply_settings_button.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

        self.query_settings_button = ttk.Button(connection_frame, text="Query Instrument", command=self._query_settings, state=tk.DISABLED)
        self.query_settings_button.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        # Frame for Current Instrument Values Display
        current_values_frame = ttk.LabelFrame(self, text="Current Instrument Values", padding="10 10 10 10", style='Dark.TLabelframe')
        current_values_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        current_values_frame.grid_columnconfigure(0, weight=1)
        current_values_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(current_values_frame, text="Center Freq (MHz):", style='TLabel').grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(current_values_frame, textvariable=self.current_center_freq_var, state='readonly', style='TEntry').grid(row=0, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(current_values_frame, text="Span (MHz):", style='TLabel').grid(row=1, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(current_values_frame, textvariable=self.current_span_var, state='readonly', style='TEntry').grid(row=1, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(current_values_frame, text="RBW (Hz):", style='TLabel').grid(row=2, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(current_values_frame, textvariable=self.current_rbw_var, state='readonly', style='TEntry').grid(row=2, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(current_values_frame, text="Reference Level (dBm):", style='TLabel').grid(row=3, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(current_values_frame, textvariable=self.current_ref_level_var, state='readonly', style='TEntry').grid(row=3, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(current_values_frame, text="Frequency Shift (Hz):", style='TLabel').grid(row=4, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(current_values_frame, textvariable=self.current_freq_shift_var, state='readonly', style='TEntry').grid(row=4, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(current_values_frame, text="Max Hold:", style='TLabel').grid(row=5, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(current_values_frame, textvariable=self.current_max_hold_var, state='readonly', style='TEntry').grid(row=5, column=1, padx=2, pady=2, sticky="ew")

        ttk.Label(current_values_frame, text="High Sensitivity:", style='TLabel').grid(row=6, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(current_values_frame, textvariable=self.current_high_sensitivity_var, state='readonly', style='TEntry').grid(row=6, column=1, padx=2, pady=2, sticky="ew")

        # Debug Options Frame
        debug_frame = ttk.LabelFrame(self, text="Debug Options", padding="10 10 10 10", style='Dark.TLabelframe')
        debug_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        debug_frame.grid_columnconfigure(0, weight=1)
        debug_frame.grid_columnconfigure(1, weight=1)

        self.general_debug_checkbox = ttk.Checkbutton(debug_frame, text="Enable General Debug",
                                                      variable=self.app_instance.general_debug_enabled_var,
                                                      command=self._toggle_general_debug, style='TCheckbutton')
        self.general_debug_checkbox.grid(row=0, column=0, padx=5, pady=2, sticky="w")

        self.log_visa_commands_checkbox = ttk.Checkbutton(debug_frame, text="Log VISA Commands",
                                                          variable=self.app_instance.log_visa_commands_enabled_var,
                                                          command=self._toggle_visa_logging, style='TCheckbutton')
        self.log_visa_commands_checkbox.grid(row=1, column=0, padx=5, pady=2, sticky="w")


        debug_print("InstrumentTab widgets created.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _populate_resources(self):
        """Calls the logic function to populate VISA resources."""
        populate_resources_logic(self.app_instance, self.console_print_func)
        # Update dropdown menu after populating resources
        menu = self.resource_dropdown["menu"]
        menu.delete(0, "end")
        resources = self.resource_names.get().split()
        for resource in resources:
            menu.add_command(label=resource, command=tk._setit(self.selected_resource, resource))
        if self.selected_resource.get() == "" and resources:
            self.selected_resource.set(resources[0]) # Set first as default if none selected

    def _connect_instrument(self):
        """Calls the logic function to connect to the instrument."""
        if connect_instrument_logic(self.app_instance, self.console_print_func):
            self._query_settings_display() # Update display after connection
        # Trigger full GUI update via main app
        self.app_instance.update_connection_status(self.app_instance.inst is not None)

    def _disconnect_instrument(self):
        """Calls the logic function to disconnect from the instrument."""
        disconnect_instrument_logic(self.app_instance, self.console_print_func)
        self._clear_settings_display() # Clear display after disconnect
        # Trigger full GUI update via main app
        self.app_instance.update_connection_status(self.app_instance.inst is not None)

    def _apply_settings(self):
        """Calls the logic function to apply settings to the instrument."""
        if apply_settings_logic(self.app_instance, self.console_print_func):
            self._query_settings_display() # Update display after applying settings

    def _query_settings(self):
        """Calls the logic function to query current instrument settings and updates display."""
        self._query_settings_display()

    def _query_settings_display(self):
        """
        Queries the current settings from the instrument and updates the display variables.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Querying current instrument settings for display...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if not self.app_instance.inst:
            self.console_print_func("⚠️ Warning: No instrument connected. Cannot query settings for display.")
            debug_print("No instrument connected. Cannot query settings for display.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            self._clear_settings_display()
            return False

        try:
            # Query Center Frequency
            center_freq_str = query_safe(self.app_instance.inst, ":SENSe:FREQuency:CENTer?", self.console_print_func)
            self.current_center_freq_var.set(f"{float(center_freq_str) / MHZ_TO_HZ:.3f}" if center_freq_str else "N/A")

            # Query Span
            span_str = query_safe(self.app_instance.inst, ":SENSe:FREQuency:SPAN?", self.console_print_func)
            self.current_span_var.set(f"{float(span_str) / MHZ_TO_HZ:.3f}" if span_str else "N/A")

            # Query RBW
            rbw_str = query_safe(self.app_instance.inst, ":SENSe:BANDwidth:RESolution?", self.console_print_func)
            self.current_rbw_var.set(f"{float(rbw_str):.0f}" if rbw_str else "N/A")

            # Query Reference Level
            ref_level_str = query_safe(self.app_instance.inst, ":DISPlay:WINDow:TRACe:Y:RLEVel?", self.console_print_func)
            self.current_ref_level_var.set(f"{float(ref_level_str):.1f}" if ref_level_str else "N/A")


            self.console_print_func("✅ Current instrument settings displayed.")
            debug_print("Current instrument settings displayed.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return True
        except Exception as e:
            self.console_print_func(f"❌ Error querying instrument settings for display: {e}")
            debug_print(f"Error querying instrument settings for display: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            self._clear_settings_display()
            return False

    def _clear_settings_display(self):
        """Clears the displayed instrument settings."""
        self.current_center_freq_var.set("")
        self.current_span_var.set("")
        self.current_rbw_var.set("")
        self.current_ref_level_var.set("")
        self.current_freq_shift_var.set("")
        self.current_max_hold_var.set("")
        self.current_high_sensitivity_var.set("")

    def _on_tab_selected(self, event):
        """
        Callback for when this tab is selected.
        This can be used to refresh data or update UI elements specific to this tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Instrument Tab selected.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        # Ensure buttons are in the correct state when tab is selected
        # Pass the buttons from this tab to the update_connection_status_logic
        self.app_instance.update_connection_status(self.app_instance.inst is not None)
        
        # Also, query and display current settings if connected
        if self.app_instance.inst:
            self._query_settings_display()
        else:
            self._clear_settings_display()

        # Set initial state of debug checkboxes based on app_instance variables
        self.general_debug_checkbox.config(variable=self.app_instance.general_debug_enabled_var)
        self.log_visa_commands_checkbox.config(variable=self.app_instance.log_visa_commands_enabled_var)

    def _toggle_general_debug(self):
        """Toggles the global debug mode based on checkbox state."""
        set_debug_mode(self.app_instance.general_debug_enabled_var.get())
        self.console_print_func(f"Debug Mode: {'Enabled' if self.app_instance.general_debug_enabled_var.get() else 'Disabled'}")

    def _toggle_visa_logging(self):
        """Toggles the global VISA command logging based on checkbox state."""
        log_visa_command(self.app_instance.log_visa_commands_enabled_var.get())
        self.console_print_func(f"VISA Command Logging: {'Enabled' if self.app_instance.log_visa_commands_enabled_var.get() else 'Disabled'}")

