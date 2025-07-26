# src/plot_logic.py
import pandas as pd
import plotly.graph_objects as go
import plotly.offline as pyo
import os
import sys
from tkinter import messagebox
from datetime import datetime # Ensure datetime is imported correctly

# Import plotting functions and constants from utils.plotting_utils
from utils.plotting_utils import plot_single_scan_data, plot_multi_trace_data, _open_plot_in_browser

# Import constants from frequency_bands.py
from utils.frequency_bands import (
    MHZ_TO_HZ,
    TV_PLOT_BAND_MARKERS,
    GOV_PLOT_BAND_MARKERS
)

def generate_single_scan_plot_and_open_wrapper_logic(app_instance, csv_file_path, output_html_path, auto_open_browser=True, single_marker_data=None):
    """
    Generates a Plotly HTML plot for a single scan and optionally opens it in a browser.
    This function is a wrapper to be called from the main thread using `app_instance.after()`.

    Inputs:
        app_instance (App): The main application instance, providing access to settings.
        csv_file_path (str): The full path to the CSV file containing the scan data.
        output_html_path (str): The full path where the generated HTML plot should be saved.
        auto_open_browser (bool, optional): If True, the generated HTML plot will be
                                            automatically opened in the default web browser. Defaults to True.
        single_marker_data (tuple, optional): A tuple (frequency_hz, name) for a single marker to highlight.
                                              Defaults to None.
    Process:
        1. Calls `plot_single_scan_data` from `utils.plotting_utils` with all necessary parameters.
           It now passes the `csv_file_path` directly, and `plot_single_scan_data` will handle reading it.
    Outputs: None (generates HTML file, may open browser)
    """
    # plot_single_scan_data will now load the data directly from the CSV
    plot_single_scan_data(
        csv_file_path, # Pass the CSV file path directly
        app_instance.include_tv_markers_var.get(), # Pass include_tv_markers
        app_instance.include_gov_markers_var.get(), # Pass include_gov_markers
        output_html_path,
        auto_open_browser,
        single_marker_data # Pass the single marker data
    )

def generate_average_plot_logic(app_instance):
    """
    Generates an averaged Plotly HTML plot from all collected scan dataframes
    and optionally opens it in a browser. This function now loads individual
    scan dataframes from their respective CSV files for historical overlays.

    Inputs:
        app_instance (App): The main application instance, providing access to
                            collected scan data, settings, and output folder.
    Process:
        1. Checks if any scan data has been collected. If not, shows a warning.
        2. Loads historical scan data from CSVs if available, converting frequencies to MHz.
        3. Concatenates all collected DataFrames (current and historical) and calculates
           aggregated metrics (average, median, range, std dev).
        4. Calls `plot_multi_trace_data` from `utils.plotting_utils` to generate the plot.
        5. Handles potential errors during plotting.
    Outputs: None (modifies app_instance state, generates files)
    """
    if not app_instance.collected_scans_dataframes:
        messagebox.showwarning("No Data", "No scan data collected yet to generate an average plot.")
        print("🚫 No scan data collected for average plot.")
        return

    try:
        # Concatenate all dataframes for the current scan session to calculate aggregated metrics
        # collected_scans_dataframes now contains DataFrames already loaded from CSVs
        all_scans_df = pd.concat(app_instance.collected_scans_dataframes)
        
        # Group by Frequency_MHz and calculate the mean, median, min, max, std, variance
        aggregated_df = all_scans_df.groupby('Frequency_MHz')['Power_dBm'].agg(
            Average_dBm='mean',
            Median_dBm='median',
            Min_dBm='min',
            Max_dBm='max',
            Std_Dev_dBm='std',
            Variance_dBm='var'
        ).reset_index()
        
        # Calculate Range (Max-Min)
        aggregated_df['Range_dBm'] = aggregated_df['Max_dBm'] - aggregated_df['Min_dBm']

        # Handle NaN values for Std_Dev_dBm and Variance_dBm (e.g., if only one data point)
        aggregated_df['Std_Dev_dBm'] = aggregated_df['Std_Dev_dBm'].fillna(0)
        aggregated_df['Variance_dBm'] = aggregated_df['Variance_dBm'].fillna(0)

        aggregated_df = aggregated_df.sort_values(by='Frequency_MHz')

        # Load historical CSVs for overlays
        historical_dfs_for_overlays = []
        output_folder = app_instance.output_folder_var.get()
        scan_name_prefix = app_instance.scan_name_var.get()

        # Regex to match scan files from previous runs
        # It looks for filenames like "MyScan_RBW..._HOLD..._Offset..._YYYYMMDD_HHMMSS.csv"
        # and excludes the current cycle's filename
        current_cycle_filename_pattern = f"{scan_name_prefix}_RBW.*_HOLD.*_Offset.*_{datetime.now().strftime('%Y%m%d_%H%M')}" # Match current minute
        
        for filename in os.listdir(output_folder):
            if filename.endswith(".csv") and filename.startswith(scan_name_prefix) and \
               not filename.startswith(app_instance.scan_name_var.get() + "_Cycle"): # Exclude single cycle plots
                
                # Check if this filename is from the *current* scan session (current run)
                # We want to exclude files generated in the current run from being historical overlays
                # This is tricky because filenames are generated with a timestamp.
                # A robust way is to compare against the list of collected_scans_dataframes' original paths
                
                # For simplicity, let's assume files already in collected_scans_dataframes are from current run.
                # And other files matching the pattern are historical.
                # This might need refinement if a user manually puts old files in the current output folder.
                
                full_path = os.path.join(output_folder, filename)
                
                # Check if this file is one of the files collected in the current session
                # (This is a simplified check; a more robust one would involve comparing file paths)
                is_current_session_file = False
                # This part needs adjustment if collected_scans_dataframes only stores DFs, not paths.
                # Since it now stores DFs loaded from CSVs, we can't easily check original paths.
                # For now, we'll rely on the assumption that files matching the current timestamp pattern are current.
                
                if app_instance.csv_filename_current_cycle and filename == os.path.basename(app_instance.csv_filename_current_cycle): # Assuming this is set in app_instance
                    continue # Skip the current cycle's raw data file
                if "Averaged_Scan_" in filename:
                    continue # Skip previously generated average plots

                try:
                    # Read CSV without header and assign column names
                    hist_df = pd.read_csv(full_path, header=None)
                    hist_df.columns = ['Frequency_MHz', 'Power_dBm']
                    
                    # Ensure Power_dBm column exists
                    if 'Power_dBm' not in hist_df.columns:
                        print(f"Warning: Historical CSV '{filename}' does not contain 'Power_dBm' column. Skipping.")
                        continue

                    # Extract timestamp for display name
                    match = re.search(r'(\d{8}_\d{6})\.csv$', filename)
                    display_name = match.group(1) if match else filename

                    historical_dfs_for_overlays.append({'name': display_name, 'df': hist_df[['Frequency_MHz', 'Power_dBm']]})
                except Exception as e:
                    print(f"Warning: Could not load historical CSV '{filename}': {e}")
        
        # Call the multi-trace plotting utility
        plot_multi_trace_data(
            aggregated_df,
            f"{app_instance.scan_name_var.get()} - Averaged, Median & Range (Historical)",
            app_instance.include_tv_markers_var.get(),
            app_instance.include_gov_markers_var.get(),
            historical_dfs_for_overlays,
            os.path.join(output_folder, f"{app_instance.scan_name_var.get()}_Averaged_Plot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
        )

    except Exception as e:
        messagebox.showerror("Plotting Error", f"An error occurred during average plot generation: {e} in {os.path.basename(__file__)}")
        print(f"❌ Error generating average plot: {e} in {os.path.basename(__file__)}")
