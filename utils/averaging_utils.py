# averaging_utils.py

import pandas as pd
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

    # Now group by Frequency_MHz and apply the aggregations on 'Power_dBm'
    # The collected_scans_dataframes already have 'Frequency_MHz'
    aggregated_df = combined_current_scans_df.groupby('Frequency_MHz')['Power_dBm'].agg(
        Average_dBm='mean',
        Median_dBm='median',
        Max_dBm='max', # Intermediate for Range
        Min_dBm='min'  # Intermediate for Range
    ).reset_index() # Reset index to make Frequency_MHz a column again

    # Calculate Range (Max - Min)
    aggregated_df['Range_dBm'] = aggregated_df['Max_dBm'] - aggregated_df['Min_dBm']

    # Drop intermediate Max_dBm and Min_dBm columns
    aggregated_df = aggregated_df.drop(columns=['Max_dBm', 'Min_dBm'])

    # Generate base filename based on current time (HHMM - no seconds) and scan name
    timestamp_hm = datetime.now().strftime("%H%M") # HHMM format (no seconds)
    base_filename = f"{scan_name_var.get()}_{timestamp_hm}_CurrentCycle" # Added _CurrentCycle for clarity

    # Define the new subfolder path
    subfolder_path = os.path.join(output_folder_var.get(), base_filename)
    os.makedirs(subfolder_path, exist_ok=True) # Ensure the new subfolder exists

    # Define paths for individual CSVs and HTML plot within the new subfolder
    average_csv_filename = os.path.join(subfolder_path, f"{base_filename}_average.csv")
    median_csv_filename = os.path.join(subfolder_path, f"{base_filename}_median.csv")
    range_csv_filename = os.path.join(subfolder_path, f"{base_filename}_range.csv")
    html_filename = os.path.join(subfolder_path, f"{base_filename}_plot.html") # HTML filename in new folder

    # Save to separate CSVs (no headers)
    try:
        aggregated_df[['Frequency_MHz', 'Average_dBm']].to_csv(average_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Current cycle average data saved to: {average_csv_filename}")

        aggregated_df[['Frequency_MHz', 'Median_dBm']].to_csv(median_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Current cycle median data saved to: {median_csv_filename}")

        aggregated_df[['Frequency_MHz', 'Range_dBm']].to_csv(range_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Current cycle range data saved to: {range_csv_filename}")

    except Exception as e:
        print(f"❌ Failed to save current cycle CSVs: {e}")
        messagebox.showerror("CSV Save Error", f"Could not save current cycle CSVs: {e}")
        return

    # Plotting the averaged, median, and range data
    try:
        fig, plot_html_path_return = plot_multi_trace_data(
            aggregated_df,
            f"{scan_name_var.get()} - Averaged, Median & Range (Current Cycle {timestamp_hm})", # Include HHMM in plot title
            include_tv_markers_var.get(),
            include_gov_markers_var.get(),
            output_html_path=html_filename # Pass the desired full path for the HTML file
        )

        if fig:
            fig.write_html(plot_html_path_return, auto_open=False)
            print(f"✅ Averaged plot for current cycle saved to: {plot_html_path_return}")
            if open_html_after_complete_var.get():
                _open_plot_in_browser(plot_html_path_return)
        else:
            print("🚫 Plotly figure was not generated for current cycle averaged data.")

    except Exception as e:
        messagebox.showerror("Plot Error", f"Could not generate or save current cycle averaged plot: {e}")
        print(f"❌ Failed to generate or save current cycle averaged plot: {e}")


def generate_historical_average_plot(scan_name_var, output_folder_var, open_html_after_complete_var, include_tv_markers_var, include_gov_markers_var):
    """
    Generates an average, median, and range plot from ALL relevant CSV files
    found in the current output folder base. This is triggered by the button.
    This plot also includes all individual historical scans as overlay layers.
    The output CSVs and HTML plot are saved into a new, dedicated subfolder.
    """
    base_output_dir = output_folder_var.get()
    if not os.path.exists(base_output_dir):
        messagebox.showwarning("Folder Not Found", f"The output folder '{base_output_dir}' does not exist. Please ensure it exists and contains scan data.")
        return

    all_historical_dfs_normalized = [] # To store DataFrames with normalized frequencies for aggregation
    historical_dfs_for_overlays = [] # To store (DataFrame, name) for plotting as overlays

    # Regex to extract components from filename.
    # Made RBW and HOLD digits flexible (\d+).
    # Made Offset optional and also account for _SESSION in older files.
    # Made timestamp flexible for seconds (\d{2})?
    filename_pattern = re.compile(
        r'^(?P<prefix>.*?)_RBW(?P<rbw>\d+K)_HOLD(?P<hold>\d+)'
        r'(?:_Offset(?P<offset>-?\d+)|_SESSION)?' # Non-capturing group for optional Offset or SESSION
        r'_(?P<datetime>\d{8}_\d{4}(?:\d{2})?)\.csv$' # Timestamp with optional seconds
    )

    print("📊 Collecting and normalizing historical scan data for averaging...")

    # Iterate directly through files in the base output directory
    for file_name in os.listdir(base_output_dir):
        # Filter for CSVs that match the scan name prefix and are not the 'averaged_cycle' or 'HISTORICAL' files
        if file_name.endswith(".csv") and file_name.startswith(scan_name_var.get() + "_") and "_averaged_cycle.csv" not in file_name and "_HISTORICAL_" not in file_name:
            csv_path = os.path.join(base_output_dir, file_name)
            try:
                # Extract components for layer name and offset using named groups
                match = filename_pattern.match(file_name)
                if not match:
                    print(f"Skipping '{file_name}': Filename does not match expected pattern. Omitting from average.")
                    continue

                # Get captured groups
                groups = match.groupdict()
                scan_name_prefix = groups['prefix']
                rbw_val = groups['rbw']
                hold_val = groups['hold']
                offset_str = groups['offset'] # Will be None if _SESSION was matched
                datetime_val = groups['datetime']

                current_offset_hz = 0.0
                if offset_str is not None:
                    current_offset_hz = float(offset_str) # Get the offset in Hz

                # Read CSV, assuming no header as per the new csv_utils.py
                df = pd.read_csv(csv_path, header=None, names=["Frequency_MHz", "Power_dBm"]).copy()

                # Calculate original (un-shifted) frequencies for averaging
                df['Original_Frequency_Hz'] = df['Frequency_MHz'] * MHZ_TO_HZ - current_offset_hz
                df['Original_Frequency_MHz'] = df['Original_Frequency_Hz'] / MHZ_TO_HZ

                # For historical overlays, we want the *actual* frequencies
                # Parse datetime string, handling both HHMM and HHMMSS formats
                if len(datetime_val) == 12: # YYYYMMDD_HHMM
                    display_datetime = datetime.strptime(datetime_val, '%Y%m%d_%H%M').strftime('%Y-%m-%d %H:%M')
                elif len(datetime_val) == 14: # YYYYMMDD_HHMMSS
                    display_datetime = datetime.strptime(datetime_val, '%Y%m%d_%H%M%S').strftime('%Y-%m-%d %H:%M:%S')
                else:
                    display_datetime = datetime_val # Fallback if format is unexpected

                display_name = f"{scan_name_prefix}_{rbw_val}_{hold_val}_Offset{int(current_offset_hz)} ({display_datetime})"
                historical_dfs_for_overlays.append((df[['Frequency_MHz', 'Power_dBm']].copy(), display_name))

                # For averaging, use the original (un-shifted) frequencies
                all_historical_dfs_normalized.append(df[['Original_Frequency_MHz', 'Power_dBm']].copy())

            except Exception as e:
                print(f"Error reading historical CSV {csv_path}: {e}")

    if not all_historical_dfs_normalized:
        messagebox.showwarning("No Data Found", f"No valid historical scan data CSV files with prefix '{scan_name_var.get()}_' and correct format found in '{base_output_dir}' to generate an average plot.")
        return

    print("📊 Generating historical averaged, median, and range data from normalized frequencies...")

    # Concatenate all historical DataFrames (with normalized frequencies) vertically for aggregation
    combined_all_scans_df_normalized = pd.concat(all_historical_dfs_normalized)

    # Now group by Original_Frequency_MHz and apply the aggregations on 'Power_dBm'
    aggregated_df = combined_all_scans_df_normalized.groupby('Original_Frequency_MHz')['Power_dBm'].agg(
        Average_dBm='mean',
        Median_dBm='median',
        Max_dBm='max', # Intermediate for Range
        Min_dBm='min'  # Intermediate for Range
    ).reset_index() # Reset index to make Original_Frequency_MHz a column again

    # Calculate Range (Max - Min)
    aggregated_df['Range_dBm'] = aggregated_df['Max_dBm'] - aggregated_df['Min_dBm']

    # Drop intermediate Max_dBm and Min_dBm columns
    aggregated_df = aggregated_df.drop(columns=['Max_dBm', 'Min_dBm'])

    # Rename the frequency column back to Frequency_MHz for plotting consistency
    aggregated_df.rename(columns={'Original_Frequency_MHz': 'Frequency_MHz'}, inplace=True)

    # Generate filename based on current time (HHMM - no seconds) and scan name
    timestamp_hm = datetime.now().strftime("%H%M") # HHMM format (no seconds)
    # Plot title should only be date and time
    plot_title_only_datetime = datetime.now().strftime("%Y-%m-%d %H:%M")
    base_filename = f"{scan_name_var.get()}_HISTORICAL_{timestamp_hm}" # Distinct filename for historical average

    # Define the new subfolder path
    subfolder_path = os.path.join(output_folder_var.get(), base_filename)
    os.makedirs(subfolder_path, exist_ok=True) # Ensure the new subfolder exists

    # Define paths for individual CSVs and HTML plot within the new subfolder
    average_csv_filename = os.path.join(subfolder_path, f"{base_filename}_average.csv")
    median_csv_filename = os.path.join(subfolder_path, f"{base_filename}_median.csv")
    range_csv_filename = os.path.join(subfolder_path, f"{base_filename}_range.csv")
    html_filename = os.path.join(subfolder_path, f"{base_filename}_plot.html") # HTML filename in new folder

    # Save to separate CSVs (no headers)
    try:
        aggregated_df[['Frequency_MHz', 'Average_dBm']].to_csv(average_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Historical average data saved to: {average_csv_filename}")

        aggregated_df[['Frequency_MHz', 'Median_dBm']].to_csv(median_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Historical median data saved to: {median_csv_filename}")

        aggregated_df[['Frequency_MHz', 'Range_dBm']].to_csv(range_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Historical range data saved to: {range_csv_filename}")

    except Exception as e:
        print(f"❌ Failed to save historical CSVs: {e}")
        messagebox.showerror("CSV Save Error", f"Could not save historical CSVs: {e}")
        return

    # Plotting the historical averaged, median, and range data, PLUS historical overlays
    try:
        fig, plot_html_path_return = plot_multi_trace_data(
            aggregated_df,
            plot_title_only_datetime, # Use only date and time for title
            include_tv_markers_var.get(),
            include_gov_markers_var.get(),
            historical_dfs_with_names=historical_dfs_for_overlays, # Pass historical data for overlays
            output_html_path=html_filename # Pass the desired full path for the HTML file
        )

        if fig:
            fig.write_html(plot_html_path_return, auto_open=False)
            print(f"✅ Historical averaged plot saved to: {plot_html_path_return}")
            if open_html_after_complete_var.get():
                _open_plot_in_browser(plot_html_path_return)
        else:
            print("🚫 Plotly figure was not generated for historical averaged data.")

    except Exception as e:
        messagebox.showerror("Plot Error", f"Could not generate or save historical averaged plot: {e}")
        print(f"❌ Failed to generate or save historical averaged plot: {e}")
