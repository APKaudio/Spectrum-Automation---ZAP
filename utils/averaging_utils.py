# averaging_utils.py

import pandas as pd
import numpy as np # Ensure numpy is imported for std, var, and log10
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


def generate_current_cycle_average_csv_and_plot(collected_scans_dataframes, scan_name_var, output_folder_var, open_html_after_complete_var, include_tv_markers_var, include_gov_markers_var, scan_rbw_hz):
    """
    Calculates average, median, range, standard deviation, variance, and power spectral density (PSD)
    from collected scan data (from current scan cycle), saves them to separate CSVs in a new subfolder,
    and plots them with overlays.
    This function is called on the main Tkinter thread via self.after().

    Args:
        collected_scans_dataframes (list): List of pandas DataFrames from current scan cycle.
        scan_name_var (tk.StringVar): Tkinter variable holding the scan name.
        output_folder_var (tk.StringVar): Tkinter variable holding the output folder path.
        open_html_after_complete_var (tk.BooleanVar): Tkinter variable for auto-opening HTML.
        include_tv_markers_var (tk.BooleanVar): Tkinter variable for including TV markers.
        include_gov_markers_var (tk.BooleanVar): Tkinter variable for including Government markers.
        scan_rbw_hz (float): The Resolution Bandwidth (RBW) in Hz used for the current scans.
    """
    if not collected_scans_dataframes:
        print("No scan data collected for current cycle averaging.")
        return

    print("\n📊 Generating averaged, median, range, std dev, variance, and PSD data for current cycle...")

    # Combine all scan data into a single DataFrame for easier processing
    combined_current_scans_df = pd.concat(collected_scans_dataframes)

    # Now group by Frequency_MHz and apply the aggregations on 'Power_dBm'
    aggregated_df = combined_current_scans_df.groupby('Frequency_MHz')['Power_dBm'].agg(
        Average_dBm='mean',
        Median_dBm='median',
        Max_dBm='max', # Intermediate for Range
        Min_dBm='min',  # Intermediate for Range
        Std_Dev_dBm='std', # Standard Deviation
        Variance_dBm='var' # Variance
    ).reset_index() # Reset index to make Frequency_MHz a column again

    # Calculate Range (Max - Min)
    aggregated_df['Range_dBm'] = aggregated_df['Max_dBm'] - aggregated_df['Min_dBm']

    # Calculate Power Spectral Density (PSD)
    # PSD (dBm/Hz) = Power (dBm) - 10 * log10(RBW_Hz)
    # Ensure RBW_Hz is not zero or negative to avoid log errors
    if scan_rbw_hz > 0:
        aggregated_df['PSD_dBm_Hz'] = aggregated_df['Average_dBm'] - 10 * np.log10(scan_rbw_hz)
    else:
        aggregated_df['PSD_dBm_Hz'] = np.nan # Set to NaN if RBW is invalid
        print(f"Warning: Invalid RBW ({scan_rbw_hz} Hz) for PSD calculation in current cycle. PSD will be NaN.")


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
    std_dev_csv_filename = os.path.join(subfolder_path, f"{base_filename}_std_dev.csv") # New CSV for Std Dev
    variance_csv_filename = os.path.join(subfolder_path, f"{base_filename}_variance.csv") # New CSV for Variance
    psd_csv_filename = os.path.join(subfolder_path, f"{base_filename}_psd.csv") # New CSV for PSD
    html_filename = os.path.join(subfolder_path, f"{base_filename}_plot.html") # HTML filename in new folder

    # Save to separate CSVs (no headers)
    try:
        aggregated_df[['Frequency_MHz', 'Average_dBm']].to_csv(average_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Current cycle average data saved to: {average_csv_filename}")

        aggregated_df[['Frequency_MHz', 'Median_dBm']].to_csv(median_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Current cycle median data saved to: {median_csv_filename}")

        aggregated_df[['Frequency_MHz', 'Range_dBm']].to_csv(range_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Current cycle range data saved to: {range_csv_filename}")

        aggregated_df[['Frequency_MHz', 'Std_Dev_dBm']].to_csv(std_dev_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Current cycle standard deviation data saved to: {std_dev_csv_filename}")

        aggregated_df[['Frequency_MHz', 'Variance_dBm']].to_csv(variance_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Current cycle variance data saved to: {variance_csv_filename}")
        
        if 'PSD_dBm_Hz' in aggregated_df.columns:
            aggregated_df[['Frequency_MHz', 'PSD_dBm_Hz']].to_csv(psd_csv_filename, index=False, float_format='%.3f', header=False)
            print(f"✅ Current cycle PSD data saved to: {psd_csv_filename}")


    except Exception as e:
        print(f"❌ Failed to save current cycle CSVs: {e}")
        messagebox.showerror("CSV Save Error", f"Could not save current cycle CSVs: {e}")
        return

    # Plotting the averaged, median, and range data
    try:
        fig, plot_html_path_return = plot_multi_trace_data(
            aggregated_df,
            f"{scan_name_var.get()} - Averaged, Median, Range, Std Dev, Variance & PSD (Current Cycle {timestamp_hm})", # Updated plot title
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
    Generates an average, median, range, standard deviation, variance, and PSD plot from ALL relevant CSV files
    found in the current output folder base. This is triggered by the button.
    This plot also includes all individual historical scans as overlay layers.
    The output CSVs and HTML plot are saved into a new, dedicated subfolder.
    """
    scan_name = scan_name_var.get() if scan_name_var.get() else "UnnamedScan"
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
        if file_name.endswith(".csv") and file_name.startswith(scan_name + "_") and "_averaged_cycle.csv" not in file_name and "_HISTORICAL_" not in file_name:
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
                rbw_val_str = groups['rbw'] # e.g., "100K"
                hold_val = groups['hold']
                offset_str = groups['offset'] # Will be None if _SESSION was matched
                datetime_val = groups['datetime']

                current_offset_hz = 0.0
                if offset_str is not None:
                    current_offset_hz = float(offset_str) # Get the offset in Hz

                # Extract RBW in Hz from filename (e.g., "100K" -> 100000)
                rbw_hz = float(rbw_val_str.replace('K', '000')) if 'K' in rbw_val_str else float(rbw_val_str)


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

                display_name = f"{scan_name_prefix}_{rbw_val_str}_{hold_val}_Offset{int(current_offset_hz)} ({display_datetime})"
                historical_dfs_for_overlays.append({"name": display_name, "df": df[['Frequency_MHz', 'Power_dBm']].copy()})

                # For averaging, use the original (un-shifted) frequencies and include RBW for PSD
                df_normalized_for_agg = df[['Original_Frequency_MHz', 'Power_dBm']].copy()
                df_normalized_for_agg['RBW_Hz'] = rbw_hz # Add RBW to each row for PSD calculation
                all_historical_dfs_normalized.append(df_normalized_for_agg)

            except Exception as e:
                print(f"Error reading historical CSV {csv_path}: {e}")

    if not all_historical_dfs_normalized:
        messagebox.showwarning("No Data Found", f"No valid historical scan data CSV files with prefix '{scan_name}_' and correct format found in '{base_output_dir}' to generate an average plot.")
        return

    print("📊 Generating historical averaged, median, range, std dev, variance, and PSD data from normalized frequencies...")

    # Concatenate all historical DataFrames (with normalized frequencies) vertically for aggregation
    combined_all_scans_df_normalized = pd.concat(all_historical_dfs_normalized)

    # First, calculate PSD for each individual power measurement before aggregation
    combined_all_scans_df_normalized['PSD_dBm_Hz'] = combined_all_scans_df_normalized.apply(
        lambda row: row['Power_dBm'] - 10 * np.log10(row['RBW_Hz']) if row['RBW_Hz'] > 0 else np.nan, axis=1
    )

    aggregated_df = combined_all_scans_df_normalized.groupby('Original_Frequency_MHz').agg(
        Average_dBm=('Power_dBm', 'mean'),
        Median_dBm=('Power_dBm', 'median'),
        Max_dBm=('Power_dBm', 'max'),
        Min_dBm=('Power_dBm', 'min'),
        Std_Dev_dBm=('Power_dBm', 'std'), # Standard Deviation
        Variance_dBm=('Power_dBm', 'var'), # Variance
        Average_PSD_dBm_Hz=('PSD_dBm_Hz', 'mean') # Average PSD
    ).reset_index()

    # Calculate Range (Max - Min)
    aggregated_df['Range_dBm'] = aggregated_df['Max_dBm'] - aggregated_df['Min_dBm']

    # Drop intermediate Max_dBm and Min_dBm columns
    aggregated_df = aggregated_df.drop(columns=['Max_dBm', 'Min_dBm'])

    # Rename the frequency column back to Frequency_MHz for plotting consistency
    aggregated_df.rename(columns={'Original_Frequency_MHz': 'Frequency_MHz'}, inplace=True)

    print("✅ Historical average, median, min, max, range, std dev, variance, and PSD calculated.")


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
    std_dev_csv_filename = os.path.join(output_dir, f"{scan_name}_HISTORICAL_{timestamp_str}_std_dev.csv") # New CSV for Std Dev
    variance_csv_filename = os.path.join(output_dir, f"{scan_name}_HISTORICAL_{timestamp_str}_variance.csv") # New CSV for Variance
    psd_csv_filename = os.path.join(output_dir, f"{scan_name}_HISTORICAL_{timestamp_str}_psd.csv") # New CSV for PSD
    html_filename = os.path.join(output_dir, f"{scan_name}_HISTORICAL_plot_{timestamp_str}.html") # HTML filename in new folder

    # Save aggregated historical data to CSVs
    try:
        # Save Average
        aggregated_df[['Frequency_MHz', 'Average_dBm']].to_csv(average_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Historical average data saved to: {average_csv_filename}")

        # Save Median
        aggregated_df[['Frequency_MHz', 'Median_dBm']].to_csv(median_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Historical median data saved to: {median_csv_filename}")

        # Save Range (Max-Min)
        aggregated_df[['Frequency_MHz', 'Range_dBm']].to_csv(range_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Historical range data saved to: {range_csv_filename}")

        # Save Standard Deviation
        aggregated_df[['Frequency_MHz', 'Std_Dev_dBm']].to_csv(std_dev_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Historical standard deviation data saved to: {std_dev_csv_filename}")

        # Save Variance
        aggregated_df[['Frequency_MHz', 'Variance_dBm']].to_csv(variance_csv_filename, index=False, float_format='%.3f', header=False)
        print(f"✅ Historical variance data saved to: {variance_csv_filename}")

        # Save PSD
        if 'Average_PSD_dBm_Hz' in aggregated_df.columns:
            aggregated_df[['Frequency_MHz', 'Average_PSD_dBm_Hz']].to_csv(psd_csv_filename, index=False, float_format='%.3f', header=False)
            print(f"✅ Historical PSD data saved to: {psd_csv_filename}")


    except Exception as e:
        print(f"❌ Failed to save historical CSVs: {e}")
        messagebox.showerror("CSV Save Error", f"Could not save historical CSVs: {e}")
        return

    # Plotting the historical averaged, median, and range data, PLUS historical overlays
    try:
        plot_title_only_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S") # Use only date and time for title
        fig, plot_html_path_return = plot_multi_trace_data(
            aggregated_df,
            f"{scan_name} - Averaged, Median, Range, Std Dev, Variance & PSD (Historical {plot_title_only_datetime})", # Updated plot title
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
        messagebox.showerror("Plot Generation Error", f"Could not generate historical averaged plot: {e}")
        print(f"❌ Failed to generate or save historical averaged plot: {e}")
