# src/scan_control.py
import tkinter as tk
from tkinter import ttk, filedialog
import threading
import time
import os
from datetime import datetime
import pandas as pd # Keep pandas for DataFrame operations on stitched data
import inspect

# Import scan-related logic
from utils.scan_instrument import scan_bands # Simplified scan_bands
from process_math.scan_stitch import process_and_stitch_scan_data # New import for data stitching
from ref.frequency_bands import MHZ_TO_HZ
from utils.instrument_control import debug_print
from src.plot_logic import plot_single_scan_data, _open_plot_in_browser
from process_math.averaging_utils import generate_current_cycle_average_csv_and_plot as generate_average_plot_logic


class ScanControlTab(ttk.Frame):
    """
    A Tkinter Frame that provides controls for starting, pausing, and stopping spectrum scans.
    """
    def __init__(self, master=None, app_instance=None, console_print_func=None, **kwargs):
        super().__init__(master, **kwargs)
        self.app_instance = app_instance
        self.console_print_func = console_print_func if console_print_func else print

        self.scan_thread = None
        self.is_scanning = False
        self.is_paused = False # New state for pausing

        self._create_widgets()
        # Initial state update will be triggered by main_app after all tabs are initialized
        # self.update_scan_button_states() # Removed: avoid calling before app_instance is fully set up

    def _create_widgets(self):
        """
        Creates and arranges the widgets for the Scan Control tab.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Creating ScanControlTab widgets...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)

        # Scan Control Buttons
        self.start_button = ttk.Button(self, text="Start Scan", command=self._start_scan, style='Green.TButton')
        self.start_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.pause_button = ttk.Button(self, text="Pause Scan", command=self._pause_scan, style='Orange.TButton', state=tk.DISABLED)
        self.pause_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.stop_button = ttk.Button(self, text="Stop Scan", command=self._stop_scan, style='Red.TButton', state=tk.DISABLED)
        self.stop_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

        debug_print("ScanControlTab widgets created.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _start_scan(self):
        """
        Starts the spectrum analyzer scan in a new thread to keep the GUI responsive.
        Manages scan state and button enablement.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Attempting to start scan...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if not self.app_instance.inst:
            self.console_print_func("⚠️ Warning: Please connect to an instrument first.")
            debug_print("Scan start failed: No instrument connected.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return

        if self.is_scanning and not self.is_paused:
            self.console_print_func("ℹ️ Info: Scan is already in progress.")
            debug_print("Scan already in progress.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            return
        
        if self.is_paused:
            self.is_paused = False
            self.console_print_func("▶️ Scan resumed.")
            debug_print("Scan resumed.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            # Trigger full GUI update after state change
            self.app_instance.update_connection_status(self.app_instance.inst is not None)
            return

        self.is_scanning = True
        self.is_paused = False
        self.scan_thread = threading.Thread(target=self._run_scan)
        self.scan_thread.daemon = True # Allow the thread to exit with the main application
        self.scan_thread.start()
        self.console_print_func("▶️ Scan started...")
        debug_print("Scan thread started.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        # Trigger full GUI update after state change
        self.app_instance.update_connection_status(self.app_instance.inst is not None)


    def _pause_scan(self):
        """
        Pauses the currently running scan.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Attempting to pause scan...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if self.is_scanning and not self.is_paused:
            self.is_paused = True
            self.console_print_func("⏸️ Scan paused.")
            debug_print("Scan paused.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            # Trigger full GUI update after state change
            self.app_instance.update_connection_status(self.app_instance.inst is not None)
        elif self.is_paused:
            self.console_print_func("ℹ️ Info: Scan is already paused.")
            debug_print("Scan already paused.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.console_print_func("ℹ️ Info: No scan is currently running to pause.")
            debug_print("No scan running to pause.", file=current_file, function=current_function, console_print_func=self.console_print_func)

    def _stop_scan(self):
        """
        Stops the currently running scan.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Attempting to stop scan...", file=current_file, function=current_function, console_print_func=self.console_print_func)

        if self.is_scanning:
            self.is_scanning = False # Signal the thread to stop
            self.is_paused = False # Ensure paused state is reset
            # Wait for the thread to finish (optional, but good for cleanup)
            if self.scan_thread and self.scan_thread.is_alive():
                # Setting the stop event will allow the scan_bands function to exit gracefully
                self.app_instance.stop_scan_event.set()
                self.scan_thread.join(timeout=2.0) # Give it a bit more time to clean up
                if self.scan_thread.is_alive():
                    self.console_print_func("⚠️ Warning: Scan thread did not terminate gracefully.")
                    debug_print("Scan thread did not terminate gracefully.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            self.console_print_func("⏹️ Scan stopped by user.")
            debug_print("Scan stopped by user.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        else:
            self.console_print_func("ℹ️ Info: No scan is currently running.")
            debug_print("No scan is currently running.", file=current_file, function=current_function, console_print_func=self.console_print_func)
        
        # Trigger full GUI update after state change
        self.app_instance.update_connection_status(self.app_instance.inst is not None)

    def _run_scan(self):
        """
        The main scan loop, run in a separate thread. This orchestrates
        the segment sweeps and then stitches the data.
        """
        current_function = inspect.currentframe().f_code.co_name
        current_file = __file__
        debug_print("Entering _run_scan function.", file=current_file, function=current_function, console_print_func=self.console_print_func)

        try:
            selected_bands = [item["band"] for item in self.app_instance.band_vars if item["var"].get()]
            if not selected_bands:
                self.app_instance.after(0, lambda: self.console_print_func("⚠️ Warning: No frequency bands selected for scan."))
                debug_print("No frequency bands selected.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                self.is_scanning = False
                self.app_instance.after(0, self.app_instance.update_connection_status, self.app_instance.inst is not None)
                return

            num_scan_cycles = int(self.app_instance.num_scan_cycles_var.get())
            cycle_wait_time = float(self.app_instance.cycle_wait_time_seconds_var.get())
            scan_name = self.app_instance.scan_name_var.get()
            output_dir = self.app_instance.output_folder_var.get()
            open_html_after_complete = self.app_instance.open_html_after_complete_var.get()
            include_tv_markers = self.app_instance.include_tv_markers_var.get()
            include_gov_markers = self.app_instance.include_gov_markers_var.get()
            include_markers = self.app_instance.include_markers_var.get()

            # Convert Tkinter StringVar values to appropriate types
            rbw_hz_val = float(self.app_instance.scan_rbw_hz_var.get())
            ref_level_dbm_val = float(self.app_instance.reference_level_dbm_var.get())
            freq_shift_hz_val = float(self.app_instance.freq_shift_hz_var.get())
            maxhold_enabled_val = self.app_instance.maxhold_enabled_var.get()
            high_sensitivity_val = self.app_instance.high_sensitivity_var.get()
            preamp_on_val = self.app_instance.preamp_on_var.get()
            rbw_step_size_hz_val = float(self.app_instance.rbw_step_size_hz_var.get())
            max_hold_time_seconds_val = float(self.app_instance.maxhold_time_seconds_var.get()) # Get max hold time

            # Ensure output directory exists
            os.makedirs(output_dir, exist_ok=True)
            debug_print(f"Ensured output directory exists: {output_dir}", file=current_file, function=current_function, console_print_func=self.console_print_func)

            self.app_instance.collected_scans_dataframes = [] # Clear previous scan data
            self.app_instance.last_scan_markers = [] # Clear previous markers

            overall_start_freq_hz = min(band["Start MHz"] for band in selected_bands) * MHZ_TO_HZ
            overall_stop_freq_hz = max(band["Stop MHz"] for band in selected_bands) * MHZ_TO_HZ

            for cycle in range(num_scan_cycles):
                # Reset stop event for each new cycle
                self.app_instance.stop_scan_event.clear()

                while self.is_paused:
                    self.app_instance.after(0, self.app_instance.update_connection_status, self.app_instance.inst is not None)
                    time.sleep(0.1)

                if not self.is_scanning:
                    self.app_instance.after(0, lambda c=cycle: self.console_print_func(f"Scan cycle {c + 1}/{num_scan_cycles} interrupted."))
                    debug_print(f"Scan cycle {cycle + 1}/{num_scan_cycles} interrupted.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                    break
                
                self.app_instance.after(0, lambda c=cycle: self.console_print_func(f"Scanning Cycle {c + 1}/{num_scan_cycles}..."))
                debug_print(f"Starting scan cycle {cycle + 1}/{num_scan_cycles}.", file=current_file, function=current_function, console_print_func=self.console_print_func)

                # Call the simplified scan_bands to collect raw data
                last_successful_band_index, raw_scan_data, markers_data = scan_bands(
                    app_instance_ref=self.app_instance,
                    inst=self.app_instance.inst,
                    selected_bands=selected_bands,
                    rbw_hz=rbw_hz_val,
                    ref_level_dbm=ref_level_dbm_val,
                    freq_shift_hz=freq_shift_hz_val,
                    maxhold_enabled=maxhold_enabled_val,
                    high_sensitivity=high_sensitivity_val,
                    preamp_on=preamp_on_val,
                    rbw_step_size_hz=rbw_step_size_hz_val,
                    max_hold_time_seconds=max_hold_time_seconds_val, # Pass max_hold_time_seconds
                    scan_name=scan_name,
                    output_folder=output_dir,
                    stop_event=self.app_instance.stop_scan_event,
                    pause_event=self.app_instance.pause_scan_event,
                    log_visa_commands_enabled=self.app_instance.log_visa_commands_enabled_var.get(),
                    general_debug_enabled=self.app_instance.general_debug_enabled_var.get(),
                    app_console_update_func=self.console_print_func
                )

                if raw_scan_data is not None and len(raw_scan_data) > 0:
                    self.app_instance.after(0, lambda: self.console_print_func(f"Processing raw data for cycle {cycle + 1}..."))
                    debug_print(f"Processing {len(raw_scan_data)} raw points for cycle {cycle + 1}.", file=current_file, function=current_function, console_print_func=self.console_print_func)

                    # Use the new process_and_stitch_scan_data function
                    scan_df = process_and_stitch_scan_data(
                        raw_scan_data,
                        overall_start_freq_hz,
                        overall_stop_freq_hz,
                        self.console_print_func
                    )

                    if not scan_df.empty:
                        self.app_instance.collected_scans_dataframes.append(scan_df)
                        self.app_instance.last_scan_markers = markers_data # Store markers from this scan (still placeholder)
                        self.app_instance.after(0, lambda: self.console_print_func(f"✅ Data collected and stitched for cycle {cycle + 1}. ({scan_df.shape[0]} points)"))
                        debug_print(f"Data collected and stitched for cycle {cycle + 1}. DataFrame shape: {scan_df.shape}", file=current_file, function=current_function, console_print_func=self.console_print_func)

                        # Generate and save single scan plot
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        plot_filename = os.path.join(output_dir, f"{scan_name}_Scan_{timestamp}.html")
                        
                        self.app_instance.after(0, lambda: self.console_print_func(f"Generating plot for cycle {cycle + 1}..."))
                        debug_print(f"Generating plot for cycle {cycle + 1} to {plot_filename}.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                        
                        fig, html_path = plot_single_scan_data(
                            scan_df,
                            f"{scan_name} - Cycle {cycle + 1} - {timestamp}",
                            include_tv_markers,
                            include_gov_markers,
                            include_markers,
                            output_html_path=plot_filename,
                            console_print_func=self.console_print_func
                        )
                        if fig:
                            self.app_instance.after(0, lambda: self.console_print_func(f"✅ Plot saved to: {html_path}"))
                            debug_print(f"Plot saved to: {html_path}", file=current_file, function=current_function, console_print_func=self.console_print_func)
                            if hasattr(self.app_instance, 'plotting_tab') and hasattr(self.app_instance.plotting_tab, 'last_plot_path'):
                                self.app_instance.plotting_tab.last_plot_path = html_path
                            if open_html_after_complete:
                                self.app_instance.after(0, lambda p=html_path: _open_plot_in_browser(p, self.console_print_func))
                        else:
                            self.app_instance.after(0, lambda: self.console_print_func(f"🚫 Failed to generate plot for cycle {cycle + 1}."))
                            debug_print(f"Failed to generate plot for cycle {cycle + 1}.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                    else:
                        self.app_instance.after(0, lambda: self.console_print_func(f"🚫 No data after stitching for cycle {cycle + 1}."))
                        debug_print(f"No data after stitching for cycle {cycle + 1}.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                else:
                    self.app_instance.after(0, lambda: self.console_print_func(f"🚫 No raw data collected for cycle {cycle + 1}."))
                    debug_print(f"No raw data collected for cycle {cycle + 1}.", file=current_file, function=current_function, console_print_func=self.console_print_func)

                if not self.is_scanning: # Check if stop was requested during the scan_bands call
                    self.app_instance.after(0, lambda c=cycle: self.console_print_func(f"Scan cycle {c + 1}/{num_scan_cycles} interrupted after data collection."))
                    debug_print(f"Scan cycle {cycle + 1}/{num_scan_cycles} interrupted after data collection.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                    break

                if cycle < num_scan_cycles - 1:
                    self.app_instance.after(0, lambda: self.console_print_func(f"Waiting {cycle_wait_time} seconds before next cycle..."))
                    debug_print(f"Waiting {cycle_wait_time} seconds before next cycle.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                    time.sleep(cycle_wait_time)

            self.is_scanning = False
            self.is_paused = False
            self.app_instance.after(0, lambda: self.console_print_func("✅ Scan complete."))
            debug_print("Scan complete. Re-enabling buttons.", file=current_file, function=current_function, console_print_func=self.console_print_func)

            # After scan, if there's collected data, update the markers tab
            if self.app_instance.last_scan_markers:
                if hasattr(self.app_instance, 'markers_display_tab') and self.app_instance.last_scan_markers:
                    headers = list(self.app_instance.last_scan_markers[0].keys())
                    self.app_instance.after(0, lambda h=headers, r=self.app_instance.last_scan_markers: self.app_instance.markers_display_tab.update_markers_data(h, r))
                    self.app_instance.after(0, lambda: self.console_print_func(f"📊 Markers data updated in Markers tab with {len(self.app_instance.last_scan_markers)} entries."))
                    debug_print(f"Markers data updated in Markers tab with {len(self.app_instance.last_scan_markers)} entries.", file=current_file, function=current_function, console_print_func=self.console_print_func)
                else:
                    self.app_instance.after(0, lambda: self.console_print_func("ℹ️ No markers extracted during scan to update Markers tab."))
                    debug_print("No markers extracted during scan or markers tab not found.", file=current_file, function=current_function, console_print_func=self.console_print_func)
            
            self.app_instance.after(0, self.app_instance.update_connection_status, self.app_instance.inst is not None)

        except Exception as e:
            self.is_scanning = False
            self.is_paused = False
            self.app_instance.after(0, lambda: self.console_print_func(f"❌ An error occurred during scan: {e}"))
            debug_print(f"Error during scan: {e}", file=current_file, function=current_function, console_print_func=self.console_print_func)
            self.app_instance.after(0, self.app_instance.update_connection_status, self.app_instance.inst is not None)

