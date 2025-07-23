# averaging_utils.py

import pandas as pd
import numpy as np
import os
import csv
from datetime import datetime
import re
import tkinter as tk # For messagebox
from tkinter import messagebox

# Import plotting functions and constants
from utils.plotting_utils import plot_multi_trace_data, _open_plot_in_browser
from utils.frequency_bands import (
    MHZ_TO_HZ,
    TV_PLOT_BAND_MARKERS,
    GOV_PLOT_BAND_MARKERS
)


def generate_current_cycle_average_csv_and_plot(collected_scans_dataframes, scan_name_var, output_folder_var, open_html_after_complete_var, include_tv_markers_var, include_gov_markers_var):
    """
    Calculates average, median, and range from collected scan data (from current scan cycle),
    saves them to separate CSVs in a new subfolder, and plots them with overlays.
    This function is called on the main Tkinter thread via self.after().
    """
    if not collected_scans_dataframes:
        print("No scan data collected for current cycle averaging.")
        return

    print("\n📊 Generating averaged, median, and range data for current cycle...")

    # Combine all scan data into a single DataFrame for easier processing
    # Concatenate the collected_scans_dataframes vertically
    combined_current_scans_df = pd.concat(collected_scans_dataframes)

    # Now group by Frequency_MHz and calculate average, median, min, and max
    aggregated_df = combined_current_scans_df.groupby('Frequency_MHz')['Power_dBm'].agg(
        Average_Power_dBm='mean',
        Median_Power_dBm='median',
        Min_Power_dBm='min',
        Max_Power_dBm='max'
    ).reset_index()

    # Calculate Range_dBm (Max - Min)
    aggregated_df['Range_dBm'] = aggregated_df['Max_Power_dBm'] - aggregated_df['Min_Power_dBm']

    # Create a unique subfolder for this averaged cycle's outputs
    scan_name = scan_name_var.get() if scan_name_var.get() else "UnnamedScan"
    # Use HHMMSS for more unique folder names
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    subfolder_name = f"{scan_name}_CYCLE_AVG_{timestamp_str}"
    output_dir = os.path.join(output_folder_var.get(), subfolder_name)
    os.makedirs(output_dir, exist_ok=True)

    # Define file paths for the aggregated data
    average_csv_filename = os.path.join(output_dir, f"{scan_name}_average_{timestamp_str}.csv")
    median_csv_filename = os.path.join(output_dir, f"{scan_name}_median_{timestamp_str}.csv")
    range_csv_filename = os.path.join(output_dir, f"{scan_name}_range_{timestamp_str}.csv")
    html_filename = os.path.join(output_dir, f"{scan_name}_average_plot_{timestamp_str}.html")

    # Save aggregated data to CSVs
    try:
        # Save Average
        aggregated_df[['Frequency_MHz', 'Average_Power_dBm']].to_csv(average_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Current cycle average data saved to: {average_csv_filename}")

        # Save Median
        aggregated_df[['Frequency_MHz', 'Median_Power_dBm']].to_csv(median_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Current cycle median data saved to: {median_csv_filename}")

        # Save Range (Max-Min)
        aggregated_df[['Frequency_MHz', 'Range_dBm']].to_csv(range_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Current cycle range data saved to: {range_csv_filename}")

    except Exception as e:
        print(f"❌ Failed to save current cycle CSVs: {e}")
        messagebox.showerror("CSV Save Error", f"Could not save current cycle CSVs: {e}")
        return

    # Plotting the current cycle averaged, median, and range data
    try:
        plot_title_only_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        fig, plot_html_path_return = plot_multi_trace_data(
            aggregated_df,
            plot_title_only_datetime,
            include_tv_markers_var.get(),
            include_gov_markers_var.get(),
            historical_dfs_with_names=None, # No historical overlays for current cycle plot
            output_html_path=html_filename
        )

        if fig:
            # fig.write_html(plot_html_path_return, auto_open=False) # Handled by plot_multi_trace_data
            print(f"✅ Current cycle averaged plot saved to: {plot_html_path_return}")
            if open_html_after_complete_var.get():
                _open_plot_in_browser(plot_html_path_return)
        else:
            print("🚫 Plotly figure was not generated for current cycle averaged data.")

    except Exception as e:
        messagebox.showerror("Plot Generation Error", f"Could not generate current cycle averaged plot: {e}")
        print(f"❌ Failed to generate current cycle averaged plot: {e}")


def generate_historical_average_plot(scan_name_var, output_folder_var, open_html_after_complete_var, include_tv_markers_var, include_gov_markers_var):
    """
    Collects all individual scan CSVs from the specified output folder,
    normalizes their frequencies based on offset, calculates historical average, median, and range,
    and plots them with individual historical scans as overlays.
    """
    scan_name = scan_name_var.get() if scan_name_var.get() else "UnnamedScan"
    base_output_dir = output_folder_var.get()

    print("\n📊 Collecting and normalizing historical scan data for averaging...")

    all_normalized_data_dfs = [] # Store DataFrames with normalized frequencies
    historical_dfs_for_overlays = [] # Store for plotting overlays (original frequencies with their names)

    # Regex to find individual scan CSVs and capture the offset
    # Pattern: scan_name_RBW<digits>K_HOLD<digits>_Offset<optional_minus><digits>_YYYYMMDD_HHMMSS.csv
    scan_file_pattern = re.compile(rf"^{re.escape(scan_name)}_RBW\d+K_HOLD\d+_Offset(-?\d+)_\d{{8}}_\d{{6}}\.csv$")

    # Walk through the base output directory to find all relevant CSVs
    for root, _, files in os.walk(base_output_dir):
        for file in files:
            match = scan_file_pattern.match(file)
            if match:
                file_path = os.path.join(root, file)
                try:
                    df = pd.read_csv(file_path, header=None, names=["Frequency_MHz", "Power_dBm"])
                    if not df.empty:
                        # Store original df for overlays
                        historical_dfs_for_overlays.append({"name": os.path.basename(file).replace(".csv", ""), "df": df.copy()})

                        # Extract offset and normalize frequencies
                        offset_hz = int(match.group(1))
                        df['Frequency_MHz_Normalized'] = df['Frequency_MHz'] - (offset_hz / MHZ_TO_HZ)
                        
                        # Add a unique scan ID for pivoting later
                        df['Scan_ID'] = os.path.basename(file).replace(".csv", "")
                        
                        all_normalized_data_dfs.append(df[['Frequency_MHz_Normalized', 'Power_dBm', 'Scan_ID']])
                        print(f"  Loaded and normalized historical scan: {file} (Offset: {offset_hz} Hz)")
                    else:
                        print(f"  Skipping empty historical scan file: {file}")
                except Exception as e:
                    print(f"❌ Error loading or normalizing historical scan file {file}: {e}")

    if not all_normalized_data_dfs:
        messagebox.showwarning("No Historical Data", "No historical scan data found to generate an average plot. Please run some scans first.")
        print("🚫 No historical scan data found for averaging.")
        return

    # Concatenate all normalized dataframes
    combined_normalized_df = pd.concat(all_normalized_data_dfs)

    # --- Pivot and Aggregate ---
    print("\n📈 Pivoting and aggregating data to calculate historical average, median, min, and max...")

    # Pivot the table to have normalized frequencies as index and Scan_ID as columns
    # This will create NaNs where a specific frequency point isn't present in a scan
    pivot_df = combined_normalized_df.pivot_table(
        index='Frequency_MHz_Normalized',
        columns='Scan_ID',
        values='Power_dBm'
    )

    # Calculate aggregated statistics across the scan columns for each frequency
    aggregated_df = pd.DataFrame({
        'Average_Power_dBm': pivot_df.mean(axis=1),
        'Median_Power_dBm': pivot_df.median(axis=1),
        'Min_Power_dBm': pivot_df.min(axis=1),
        'Max_Power_dBm': pivot_df.max(axis=1)
    }).reset_index() # Reset index to make Frequency_MHz_Normalized a column

    # Rename the frequency column back to 'Frequency_MHz' for consistency with plotting functions
    aggregated_df.rename(columns={'Frequency_MHz_Normalized': 'Frequency_MHz'}, inplace=True)

    # Calculate Range_dBm (Max - Min) - this should now be correct and non-zero if there's variation
    aggregated_df['Range_dBm'] = aggregated_df['Max_Power_dBm'] - aggregated_df['Min_Power_dBm']
    print("✅ Historical average, median, min, max, and range calculated.")


    # Create a unique subfolder for this historical averaged plot's outputs
    # Use HHMMSS for more unique folder names
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    subfolder_name = f"{scan_name}_HISTORICAL_{timestamp_str}"
    output_dir = os.path.join(base_output_dir, subfolder_name)
    os.makedirs(output_dir, exist_ok=True)

    # Define file paths for the aggregated historical data
    average_csv_filename = os.path.join(output_dir, f"{scan_name}_HISTORICAL_{timestamp_str}_average.csv")
    median_csv_filename = os.path.join(output_dir, f"{scan_name}_HISTORICAL_{timestamp_str}_median.csv")
    range_csv_filename = os.path.join(output_dir, f"{scan_name}_HISTORICAL_{timestamp_str}_range.csv")
    html_filename = os.path.join(output_dir, f"{scan_name}_HISTORICAL_average_plot_{timestamp_str}.html")

    # Save aggregated historical data to CSVs
    try:
        # Save Average
        aggregated_df[['Frequency_MHz', 'Average_Power_dBm']].to_csv(average_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Historical average data saved to: {average_csv_filename}")

        # Save Median
        aggregated_df[['Frequency_MHz', 'Median_Power_dBm']].to_csv(median_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Historical median data saved to: {median_csv_filename}")

        # Save Range (Max-Min)
        aggregated_df[['Frequency_MHz', 'Range_dBm']].to_csv(range_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Historical range data saved to: {range_csv_filename}")

    except Exception as e:
        print(f"❌ Failed to save historical CSVs: {e}")
        messagebox.showerror("CSV Save Error", f"Could not save historical CSVs: {e}")
        return

    # Plotting the historical averaged, median, and range data, PLUS historical overlays
    try:
        plot_title_only_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S") # Use only date and time for title
        fig, plot_html_path_return = plot_multi_trace_data(
            aggregated_df,
            plot_title_only_datetime,
            include_tv_markers_var.get(),
            include_gov_markers_var.get(),
            historical_dfs_with_names=historical_dfs_for_overlays, # Pass historical data for overlays
            output_html_path=html_filename # Pass the desired full path for the HTML file
        )

        if fig:
            # fig.write_html(plot_html_path_return, auto_open=False) # Handled by plot_multi_trace_data
            print(f"✅ Historical averaged plot saved to: {plot_html_path_return}")
            if open_html_after_complete_var.get():
                _open_plot_in_browser(plot_html_path_return)
        else:
            print("🚫 Plotly figure was not generated for historical averaged data.")

    except Exception as e:
        messagebox.showerror("Plot Generation Error", f"Could not generate historical averaged plot: {e}")
        print(f"❌ Failed to generate or save historical averaged plot: {e}")
