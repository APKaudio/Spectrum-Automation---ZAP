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
from utils.plotting_utils import plot_multi_trace_data, _open_plot_in_browser # Changed to utils.plotting_utils
from ref.frequency_bands import (
    MHZ_TO_HZ,
    TV_PLOT_BAND_MARKERS,
    GOV_PLOT_BAND_MARKERS
)
from utils.instrument_control import debug_print # Import debug_print
import inspect # Import inspect module

def generate_current_cycle_average_csv_and_plot(
    collected_scans_dataframes,
    output_dir,
    include_tv_markers_var,
    include_gov_markers_var,
    open_html_after_complete_var,
    console_print_func,
    include_markers_var=None # Ensure this is compatible with plotting_utils.py if used
):
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Input collected_scans_dataframes (keys): {collected_scans_dataframes.keys() if collected_scans_dataframes else 'None'}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Input output_dir: {output_dir}", file=current_file, function=current_function, console_print_func=console_print_func)

    if not collected_scans_dataframes:
        console_print_func("No scan dataframes provided for averaging.")
        debug_print("No scan dataframes for current cycle averaging.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (no dataframes)", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    last_scan_name = list(collected_scans_dataframes.keys())[-1]
    plot_title_only_datetime = datetime.now().strftime("%Y-%m-%d %H%M%S")
    scan_name = f"Cycle_{plot_title_only_datetime}"
    debug_print(f"Scan name for current cycle: {scan_name}", file=current_file, function=current_function, console_print_func=console_print_func)

    aligned_dfs = []
    first_df_freq = collected_scans_dataframes[list(collected_scans_dataframes.keys())[0]]['Frequency (Hz)']
    debug_print(f"Reference frequency axis from first dataframe (first few points): {first_df_freq.head().tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)

    for df_name, df in collected_scans_dataframes.items():
        if 'Frequency (Hz)' in df.columns and 'Power (dBm)' in df.columns:
            aligned_df = df.set_index('Frequency (Hz)').reindex(first_df_freq).reset_index()
            aligned_dfs.append(aligned_df.set_index('Frequency (Hz)')['Power (dBm)'])
            debug_print(f"Aligned {df_name} for power levels.", file=current_file, function=current_function, console_print_func=console_print_func)
        else:
            console_print_func(f"Skipping {df_name}: Missing 'Frequency (Hz)' or 'Power (dBm)' column.")
            debug_print(f"Skipping {df_name} due to missing columns.", file=current_file, function=current_function, console_print_func=console_print_func)
            continue

    if not aligned_dfs:
        console_print_func("No valid scan data to average after alignment.")
        debug_print("No valid scan data to average after alignment.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (no aligned data)", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    power_levels_df = pd.concat(aligned_dfs, axis=1)
    power_levels_df.columns = [f"Scan_{i+1}" for i in range(len(aligned_dfs))]
    debug_print(f"Combined power_levels_df shape: {power_levels_df.shape}", file=current_file, function=current_function, console_print_func=console_print_func)

    average_levels = power_levels_df.mean(axis=1)
    median_levels = power_levels_df.median(axis=1)
    range_levels = power_levels_df.max(axis=1) - power_levels_df.min(axis=1)
    std_dev_levels = power_levels_df.std(axis=1)
    variance_levels = power_levels_df.var(axis=1)
    debug_print("Calculated average, median, range, std dev, variance.", file=current_file, function=current_function, console_print_func=console_print_func)

    aggregated_df = pd.DataFrame({
        'Frequency (Hz)': power_levels_df.index,
        'Average (dBm)': average_levels,
        'Median (dBm)': median_levels,
        'Range (dBm)': range_levels,
        'Std Dev (dBm)': std_dev_levels,
        'Variance (dBm^2)': variance_levels,
    }).reset_index(drop=True)
    debug_print(f"Aggregated DataFrame columns: {aggregated_df.columns.tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)

    csv_filename = os.path.join(output_dir, f"{scan_name}_aggregated_data.csv")
    aggregated_df.to_csv(csv_filename, index=False)
    console_print_func(f"✅ Aggregated data saved to: {csv_filename}")
    debug_print(f"Aggregated data saved to: {csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

    html_filename = os.path.join(output_dir, f"{scan_name}_aggregated_plot.html")
    debug_print(f"Plot HTML filename: {html_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

    historical_dfs_for_overlays = [
        {'name': df_name, 'df': df}
        for df_name, df in collected_scans_dataframes.items()
    ]
    debug_print(f"Prepared {len(historical_dfs_for_overlays)} historical dataframes for overlays.", file=current_file, function=current_function, console_print_func=console_print_func)

    fig, plot_html_path_return = plot_multi_trace_data(
        aggregated_df,
        f"{scan_name} - Averaged, Median, Range, Std Dev, Variance & PSD (Historical {plot_title_only_datetime})",
        include_tv_markers_var.get(),
        include_gov_markers_var.get(),
        historical_dfs_with_names=historical_dfs_for_overlays,
        output_html_path=html_filename,
        console_print_func=console_print_func
    )

    if fig:
        fig.write_html(plot_html_path_return, auto_open=False)
        console_print_func(f"✅ Historical averaged plot saved to: {plot_html_path_return}")
        if open_html_after_complete_var.get():
            _open_plot_in_browser(plot_html_path_return, console_print_func)
            debug_print(f"Opened plot in browser: {plot_html_path_return}", file=current_file, function=current_function, console_print_func=console_print_func)
    else:
        console_print_func("🚫 Plotly figure was not generated for historical averaged data.")
        debug_print("Plotly figure not generated for historical averaged data.", file=current_file, function=current_function, console_print_func=console_print_func)

    debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    return fig, plot_html_path_return


def generate_multi_file_average_and_plot(
    file_paths,
    selected_avg_types,
    plot_title_prefix,
    include_tv_markers,
    include_gov_markers,
    output_html_path,
    open_html_after_complete,
    console_print_func
):
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Input file_paths ({len(file_paths)} files): {file_paths}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Input selected_avg_types: {selected_avg_types}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Input output_html_path: {output_html_path}", file=current_file, function=current_function, console_print_func=console_print_func)

    if not file_paths:
        console_print_func("No file paths provided for multi-file averaging.")
        debug_print("No file paths provided for multi-file averaging.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (no file paths)", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    all_scans_dfs = []
    for f_path in file_paths:
        try:
            df = pd.read_csv(f_path)
            if 'Frequency (Hz)' in df.columns and 'Power (dBm)' in df.columns:
                all_scans_dfs.append(df)
                debug_print(f"Successfully read file: {os.path.basename(f_path)}", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                console_print_func(f"Skipping {os.path.basename(f_path)}: Missing 'Frequency (Hz)' or 'Power (dBm)' columns.")
                debug_print(f"Skipping {os.path.basename(f_path)} due to missing columns.", file=current_file, function=current_function, console_print_func=console_print_func)
        except Exception as e:
            console_print_func(f"Error reading {os.path.basename(f_path)}: {e}")
            debug_print(f"Error reading {os.path.basename(f_path)}: {e}", file=current_file, function=current_function, console_print_func=console_print_func)

    if not all_scans_dfs:
        console_print_func("No valid scan data could be loaded from the selected files for averaging.")
        debug_print("No valid scan data could be loaded from the selected files for averaging.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (no valid data)", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    reference_freq = all_scans_dfs[0]['Frequency (Hz)']
    debug_print(f"Reference frequency axis from first loaded file (first few points): {reference_freq.head().tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)
    power_levels_aligned = []

    for df in all_scans_dfs:
        aligned_power_series = df.set_index('Frequency (Hz)')['Power (dBm)'].reindex(reference_freq)
        power_levels_aligned.append(aligned_power_series)
        debug_print(f"Aligned one dataframe's power levels.", file=current_file, function=current_function, console_print_func=console_print_func)

    power_levels_df = pd.concat(power_levels_aligned, axis=1)
    power_levels_df.columns = [f"File_{i+1}" for i in range(len(power_levels_aligned))]
    debug_print(f"Combined and aligned power_levels_df shape: {power_levels_df.shape}", file=current_file, function=current_function, console_print_func=console_print_func)

    aggregated_df_columns = {'Frequency (Hz)': reference_freq}
    debug_print(f"Calculating selected average types: {selected_avg_types}", file=current_file, function=current_function, console_print_func=console_print_func)

    if "Average" in selected_avg_types:
        aggregated_df_columns['Average (dBm)'] = power_levels_df.mean(axis=1)
    if "Median" in selected_avg_types:
        aggregated_df_columns['Median (dBm)'] = power_levels_df.median(axis=1)
    if "Range" in selected_avg_types:
        aggregated_df_columns['Range (dBm)'] = power_levels_df.max(axis=1) - power_levels_df.min(axis=1)
    if "Std Dev" in selected_avg_types:
        aggregated_df_columns['Std Dev (dBm)'] = power_levels_df.std(axis=1)
    if "Variance" in selected_avg_types:
        aggregated_df_columns['Variance (dBm^2)'] = power_levels_df.var(axis=1)
    if "PSD (dBm/Hz)" in selected_avg_types:
        console_print_func("Warning: PSD (dBm/Hz) calculation requires Resolution Bandwidth (RBW) which is not available from CSVs. Plotting Average as a proxy.")
        debug_print("PSD selected, using Average as proxy.", file=current_file, function=current_function, console_print_func=console_print_func)
        aggregated_df_columns['PSD (dBm/Hz)'] = power_levels_df.mean(axis=1)

    aggregated_df = pd.DataFrame(aggregated_df_columns).reset_index(drop=True)
    debug_print(f"Final aggregated_df columns: {aggregated_df.columns.tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)

    plot_title_suffix = ", ".join(selected_avg_types)
    plot_title = f"{plot_title_prefix} - {plot_title_suffix} (Multi-File Average)"
    debug_print(f"Generated plot title: {plot_title}", file=current_file, function=current_function, console_print_func=console_print_func)

    fig, plot_html_path_return = plot_multi_trace_data(
        aggregated_df,
        plot_title,
        include_tv_markers,
        include_gov_markers,
        historical_dfs_with_names=None,
        output_html_path=output_html_path,
        console_print_func=console_print_func
    )

    if fig and open_html_after_complete:
        _open_plot_in_browser(plot_html_path_return, console_print_func)
        debug_print(f"Opened plot in browser: {plot_html_path_return}", file=current_file, function=current_function, console_print_func=console_print_func)
    elif not fig:
        console_print_func("🚫 Plotly figure was not generated for multi-file averaged data.")
        debug_print("Plotly figure not generated for multi-file averaged data.", file=current_file, function=current_function, console_print_func=console_print_func)

    debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    return fig, plot_html_path_return
