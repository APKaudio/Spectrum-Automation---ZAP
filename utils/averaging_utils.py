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
    saves them to a CSV, and plots them with overlays.
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

    # Frequency_MHz is already present from the groupby operation, no need to re-add.

    # Generate filename based on current time (HMS) and scan name
    timestamp_hms = datetime.now().strftime("%H%M") # HMS format
    base_filename = f"{scan_name_var.get()}_{timestamp_hms}"

    csv_filename = os.path.join(output_folder_var.get(), f"{base_filename}_averaged_cycle.csv") # Added _cycle to distinguish
    html_filename = os.path.join(output_folder_var.get(), f"{base_filename}_averaged_cycle.html") # Added _cycle to distinguish

    # Ensure output directory exists
    os.makedirs(output_folder_var.get(), exist_ok=True)

    # Save to CSV
    try:
        # Select columns for CSV: Frequency_MHz, Average_dBm, Median_dBm, Range_dBm
        aggregated_df.to_csv(csv_filename, index=False, float_format='%.3f',
                             columns=['Frequency_MHz', 'Average_dBm', 'Median_dBm', 'Range_dBm'])
        print(f"✅ Averaged data for current cycle saved to: {csv_filename}")
    except Exception as e:
        print(f"❌ Failed to save averaged CSV for current cycle: {e}")
        messagebox.showerror("CSV Save Error", f"Could not save averaged CSV for current cycle: {e}")
        return

    # Plotting the averaged, median, and range data
    try:
        fig, plot_html_path_return = plot_multi_trace_data(
            aggregated_df,
            f"{scan_name_var.get()} - Averaged, Median & Range (Current Cycle {timestamp_hms})", # Include HMS in plot title
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
    """
    base_output_dir = output_folder_var.get()
    if not os.path.exists(base_output_dir):
        messagebox.showwarning("Folder Not Found", f"The output folder '{base_output_dir}' does not exist. Please ensure it exists and contains scan data.")
        return

    all_historical_dfs_normalized = [] # To store DataFrames with normalized frequencies for aggregation
    historical_dfs_for_overlays = [] # To store (DataFrame, name) for plotting as overlays

    # Regex to extract datetime string and offset from filename: e.g., MyScan_RBW####K_HOLD##_Offset####_YYYYMMDD_HHMMSS.csv
    # Changed Offset(\d+) to Offset(-?\d+) to allow for negative offsets if they were ever introduced.
    filename_pattern = re.compile(r'^(.*)_RBW(\d{4}K)_HOLD(\d{2})_Offset(-?\d+)_(\d{8}_\d{6})\.csv$')

    print("📊 Collecting and normalizing historical scan data for averaging...")

    # Iterate directly through files in the base output directory
    for file_name in os.listdir(base_output_dir):
        # Filter for CSVs that match the scan name prefix and are not the 'averaged_cycle' or 'HISTORICAL' files
        if file_name.endswith(".csv") and file_name.startswith(scan_name_var.get() + "_") and "_averaged_cycle.csv" not in file_name and "_HISTORICAL_" not in file_name:
            csv_path = os.path.join(base_output_dir, file_name)
            try:
                # Extract components for layer name and offset
                match = filename_pattern.match(file_name)
                if not match:
                    print(f"Skipping '{file_name}': Filename does not match expected pattern with Offset. Omitting from average.")
                    continue

                scan_name_prefix, rbw_val, hold_val, offset_str, datetime_val = match.groups()
                current_offset_hz = float(offset_str) # Get the offset in Hz

                # Read CSV, assuming the header is "Frequency (MHz)", "Level (dBm)"
                df = pd.read_csv(csv_path).copy()

                # Strip whitespace from column names
                df.columns = df.columns.str.strip()

                # Ensure columns are as expected and rename for internal consistency
                if "Frequency (MHz)" in df.columns and "Level (dBm)" in df.columns:
                    df.rename(columns={"Frequency (MHz)": "Frequency_MHz", "Level (dBm)": "Power_dBm"}, inplace=True)
                    
                    # Calculate original (un-shifted) frequencies for averaging
                    df['Original_Frequency_Hz'] = df['Frequency_MHz'] * MHZ_TO_HZ - current_offset_hz
                    df['Original_Frequency_MHz'] = df['Original_Frequency_Hz'] / MHZ_TO_HZ

                    # For historical overlays, we want the *actual* frequencies
                    display_datetime = datetime.strptime(datetime_val, '%Y%m%d_%H%M%S').strftime('%Y-%m-%d %H:%M:%S')
                    display_name = f"{scan_name_prefix}_{rbw_val}_{hold_val}_Offset{offset_str} ({display_datetime})"
                    historical_dfs_for_overlays.append((df[['Frequency_MHz', 'Power_dBm']].copy(), display_name))

                    # For averaging, use the original (un-shifted) frequencies
                    all_historical_dfs_normalized.append(df[['Original_Frequency_MHz', 'Power_dBm']].copy())

                else:
                    print(f"Skipping {file_name}: Missing expected columns or incorrect format. Expected 'Frequency (MHz)' and 'Level (dBm)'.")
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

    # Generate filename based on current time (HMS) and scan name
    timestamp_hms = datetime.now().strftime("%H%M") # HMS format
    # Plot title should only be date and time
    plot_title_only_datetime = datetime.now().strftime("%Y-%m-%d %H:%M") 
    base_filename = f"{scan_name_var.get()}_HISTORICAL_{timestamp_hms}" # Distinct filename for historical average

    csv_filename = os.path.join(output_folder_var.get(), f"{base_filename}_averaged.csv")
    html_filename = os.path.join(output_folder_var.get(), f"{base_filename}_averaged.html")

    # Ensure output directory exists
    os.makedirs(output_folder_var.get(), exist_ok=True)

    # Save to CSV
    try:
        # Select columns for CSV: Frequency_MHz, Average_dBm, Median_dBm, Range_dBm
        aggregated_df.to_csv(csv_filename, index=False, float_format='%.3f',
                             columns=['Frequency_MHz', 'Average_dBm', 'Median_dBm', 'Range_dBm'])
        print(f"✅ Historical averaged data saved to: {csv_filename}")
    except Exception as e:
        print(f"❌ Failed to save historical averaged CSV: {e}")
        messagebox.showerror("CSV Save Error", f"Could not save historical averaged CSV: {e}")
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
