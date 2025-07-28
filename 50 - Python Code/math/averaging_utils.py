# averaging_utils.py
#
# This module provides utility functions for processing and analyzing collected
# spectrum scan data. It includes functionalities for calculating various
# statistical measures such as average, median, range, standard deviation,
# variance, and power spectral density (PSD) from multiple scan cycles.
# It is crucial for generating insightful plots and CSV reports from the raw scan data.
#
# Author: Anthony Peter Kuzub
# Blog: www.Like.audio (Contributor to this project)
#
# Professional services for customizing and tailoring this software to your specific
# application can be negotiated. There is no change to use, modify, or fork this software.
#
# Build Log: https://like.audio/category/software/spectrum-scanner/
# Source Code: https://github.com/APKaudio/
#
import pandas as pd
import numpy as np # Ensure numpy is imported for std, var, and log10
import os
import csv
from datetime import datetime
import re
# import tkinter as tk # For messagebox # Removed
# from tkinter import messagebox # Removed

# Import plotting functions and constants
from utils.plotting_utils import plot_multi_trace_data, _open_plot_in_browser
from utils.frequency_bands import (
    MHZ_TO_HZ,
    TV_PLOT_BAND_MARKERS,
    GOV_PLOT_BAND_MARKERS
)
from utils.instrument_control import debug_print # Import debug_print
import inspect # Import inspect module

def generate_current_cycle_average_csv_and_plot(collected_scans_dataframes, scan_name_var, output_folder_var, open_html_after_complete_var, include_tv_markers_var, include_gov_markers_var, include_markers_var, console_print_func):
    """
    Generates an averaged CSV and an interactive HTML plot from collected scan data.
    This function processes multiple scan dataframes, calculates various statistics
    (average, median, range, std dev, variance, PSD), saves them to CSV, and generates a plot.

    Inputs:
        collected_scans_dataframes (list): A list of pandas DataFrames, each representing a scan cycle.
        scan_name_var (tk.StringVar): Tkinter variable for the base name of the scan.
        output_folder_var (tk.StringVar): Tkinter variable for the output directory.
        open_html_after_complete_var (tk.BooleanVar): Tkinter variable to open HTML plot after generation.
        include_tv_markers_var (tk.BooleanVar): Tkinter variable to include TV markers on plot.
        include_gov_markers_var (tk.BooleanVar): Tkinter variable to include Government markers on plot.
        include_markers_var (tk.BooleanVar): Tkinter variable to include extracted markers on plot.
        console_print_func (function): Function to use for console output.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print("Generating averaged CSV and plot...", file=current_file, function=current_function, console_print_func=console_print_func)

    if not collected_scans_dataframes:
        console_print_func("⚠️ Warning: No scan data available to generate average plot.")
        debug_print("No scan data for average plot.", file=current_file, function=current_function, console_print_func=console_print_func)
        return

    scan_name = scan_name_var.get()
    output_folder = output_folder_var.get()

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        console_print_func(f"Created output directory: {output_folder}")
        debug_print(f"Created output directory: {output_folder}", file=current_file, function=current_function, console_print_func=console_print_func)

    # Ensure all dataframes have the same frequency points for aggregation
    # This assumes frequency is the first column and is consistent.
    # A more robust solution might involve resampling or merging.
    base_df = collected_scans_dataframes[0].copy()
    frequencies = base_df.iloc[:, 0]

    # Extract all power level columns
    all_power_levels = [df.iloc[:, 1] for df in collected_scans_dataframes] # Assuming power is second column

    # Convert list of Series to a DataFrame for easier aggregation
    power_levels_df = pd.DataFrame(all_power_levels).T
    power_levels_df.columns = [f'Cycle_{i+1}' for i in range(len(collected_scans_dataframes))]

    # Calculate statistics
    average_levels = power_levels_df.mean(axis=1)
    median_levels = power_levels_df.median(axis=1)
    range_levels = power_levels_df.max(axis=1) - power_levels_df.min(axis=1)
    std_dev_levels = power_levels_df.std(axis=1)
    variance_levels = power_levels_df.var(axis=1)

    # Calculate Power Spectral Density (PSD)
    # Assuming the frequency step is constant for PSD calculation
    if len(frequencies) > 1:
        freq_step_hz = (frequencies.iloc[1] - frequencies.iloc[0]) * MHZ_TO_HZ
    else:
        freq_step_hz = 1 # Default to 1 Hz if only one frequency point
        console_print_func("⚠️ Warning: Only one frequency point, PSD calculation may not be meaningful.")
        debug_print("Single frequency point, PSD may be inaccurate.", file=current_file, function=current_function, console_print_func=console_print_func)

    # Convert dBm to Watts for linear averaging, then to dBm/Hz for PSD
    power_watts = 10**((average_levels - 30) / 10) # Convert dBm to Watts
    psd_dbm_per_hz = 10 * np.log10(power_watts / freq_step_hz) + 30 # Convert to dBm/Hz

    # Create aggregated DataFrame
    aggregated_df = pd.DataFrame({
        'Frequency (MHz)': frequencies,
        'Average Power (dBm)': average_levels,
        'Median Power (dBm)': median_levels,
        'Range (dB)': range_levels,
        'Standard Deviation (dB)': std_dev_levels,
        'Variance (dB^2)': variance_levels,
        'PSD (dBm/Hz)': psd_dbm_per_hz
    })

    # Prepare historical data for plotting overlays
    historical_dfs_for_overlays = []
    for i, df in enumerate(collected_scans_dataframes):
        historical_dfs_for_overlays.append({
            'name': f'Cycle {i+1} (Raw)',
            'df': df,
            'x_col': df.columns[0], # Frequency column
            'y_col': df.columns[1]  # Power level column
        })

    # Save aggregated data to CSV
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_filename = os.path.join(output_folder, f"{scan_name}_averaged_data_{timestamp}.csv")
    try:
        aggregated_df.to_csv(csv_filename, index=False)
        console_print_func(f"✅ Averaged data saved to: {csv_filename}")
        debug_print(f"Averaged data saved: {csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)
    except Exception as e:
        console_print_func(f"❌ Failed to save averaged CSV: {e}")
        debug_print(f"Failed to save averaged CSV: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        # messagebox.showerror("CSV Save Error", f"Could not save averaged CSV: {e}") # Removed
        return

    # Save historical raw data to separate CSVs if there's more than one cycle
    if len(collected_scans_dataframes) > 1:
        historical_csv_folder = os.path.join(output_folder, f"{scan_name}_raw_cycles_{timestamp}")
        os.makedirs(historical_csv_folder, exist_ok=True)
        debug_print(f"Created historical CSV folder: {historical_csv_folder}", file=current_file, function=current_function, console_print_func=console_print_func)
        try:
            for i, df in enumerate(collected_scans_dataframes):
                raw_csv_path = os.path.join(historical_csv_folder, f"{scan_name}_raw_cycle_{i+1}.csv")
                df.to_csv(raw_csv_path, index=False)
                debug_print(f"Saved raw cycle {i+1} to: {raw_csv_path}", file=current_file, function=current_function, console_print_func=console_print_func)
            console_print_func(f"✅ Historical raw scan data saved to: {historical_csv_folder}")
        except Exception as e:
            console_print_func(f"❌ Failed to save historical CSVs: {e}")
            debug_print(f"Failed to save historical CSVs: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
            # messagebox.showerror("CSV Save Error", f"Could not save historical CSVs: {e}") # Removed
            return

    # Plotting the historical averaged, median, and range data, PLUS historical overlays
    try:
        plot_title_only_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S") # Use only date and time for title
        html_filename = os.path.join(output_folder, f"{scan_name}_averaged_plot_{timestamp}.html")
        fig, plot_html_path_return = plot_multi_trace_data(
            aggregated_df,
            f"{scan_name} - Averaged, Median, Range, Std Dev, Variance & PSD (Historical {plot_title_only_datetime})", # Updated plot title
            include_tv_markers_var.get(),
            include_gov_markers_var.get(),
            include_markers_var.get(), # Pass include_markers_var
            historical_dfs_with_names=historical_dfs_for_overlays, # Pass historical data for overlays
            output_html_path=html_filename, # Pass the desired full path for the HTML file
            console_print_func=console_print_func # Pass console_print_func
        )

        if fig:
            fig.write_html(plot_html_path_return, auto_open=False)
            console_print_func(f"✅ Historical averaged plot saved to: {plot_html_path_return}")
            if open_html_after_complete_var.get():
                _open_plot_in_browser(plot_html_path_return, console_print_func) # Pass console_print_func
        else:
            console_print_func("🚫 Plotly figure was not generated for historical averaged data.")

    except Exception as e:
        console_print_func(f"❌ Error generating historical averaged plot: {e}")
        debug_print(f"Error generating historical averaged plot: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        # messagebox.showerror("Plotting Error", f"An error occurred while generating the historical averaged plot: {e}") # Removed

