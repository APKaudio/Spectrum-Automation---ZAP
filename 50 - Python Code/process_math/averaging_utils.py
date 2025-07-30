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

# Import plotting functions and constants
from utils.plotting_utils import plot_multi_trace_data, _open_plot_in_browser
from ref.frequency_bands import (
    MHZ_TO_HZ,
    TV_PLOT_BAND_MARKERS,
    GOV_PLOT_BAND_MARKERS
)
from utils.instrument_control import debug_print # Import debug_print
import inspect # Import inspect module

# --- Helper Functions for Calculations and Folder Management ---

def _create_output_subfolder(base_output_dir, prefix, timestamp_str, console_print_func):
    """
    Creates a new subfolder for scan outputs based on a prefix and timestamp.

    Inputs:
        base_output_dir (str): The base directory where the subfolder will be created.
        prefix (str): A descriptive prefix for the subfolder name (e.g., scan name, group name).
        timestamp_str (str): A timestamp string (e.g., YYYYMMDD_HHMMSS) for uniqueness.
        console_print_func (function): Function to print messages to the GUI console.

    Returns:
        str: The full path to the newly created subfolder.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)

    subfolder_name = f"{prefix}_{timestamp_str}"
    output_dir_full = os.path.join(base_output_dir, subfolder_name)
    os.makedirs(output_dir_full, exist_ok=True)
    console_print_func(f"Created output subfolder: {output_dir_full}")
    debug_print(f"Created output subfolder: {output_dir_full}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    return output_dir_full

def _calculate_average(power_levels_df, console_print_func):
    """Calculates the average power levels across all traces."""
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    average_levels = power_levels_df.mean(axis=1)
    debug_print(f"Calculated Average. First 5 values: {average_levels.head().tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    return average_levels

def _calculate_median(power_levels_df, console_print_func):
    """Calculates the median power levels across all traces."""
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    median_levels = power_levels_df.median(axis=1)
    debug_print(f"Calculated Median. First 5 values: {median_levels.head().tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    return median_levels

def _calculate_range(power_levels_df, console_print_func):
    """Calculates the range (max - min) of power levels across all traces."""
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    range_levels = power_levels_df.max(axis=1) - power_levels_df.min(axis=1)
    debug_print(f"Calculated Range. First 5 values: {range_levels.head().tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    return range_levels

def _calculate_std_dev(power_levels_df, console_print_func):
    """Calculates the standard deviation of power levels across all traces."""
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    std_dev_levels = power_levels_df.std(axis=1)
    debug_print(f"Calculated Std Dev. First 5 values: {std_dev_levels.head().tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    return std_dev_levels

def _calculate_variance(power_levels_df, console_print_func):
    """Calculates the variance of power levels across all traces."""
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    variance_levels = power_levels_df.var(axis=1)
    debug_print(f"Calculated Variance. First 5 values: {variance_levels.head().tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    return variance_levels

def _calculate_psd(power_levels_df, rbw_values, console_print_func):
    """
    Calculates the Power Spectral Density (PSD) from power levels and RBW values.
    If multiple traces, calculates PSD for each then averages.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    
    psd_levels = pd.Series([np.nan] * len(power_levels_df.index), index=power_levels_df.index) # Initialize with NaN

    if not rbw_values or all(rbw is None or rbw <= 0 for rbw in rbw_values):
        console_print_func("Warning: Resolution Bandwidth (RBW) not provided or invalid for PSD calculation. PSD will be NaN.")
        debug_print("RBW missing or invalid for PSD calculation.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (invalid RBW)", file=current_file, function=current_function, console_print_func=console_print_func)
        return psd_levels

    linear_power_mW_traces = 10**(power_levels_df / 10)
    psd_traces = []
    
    # Ensure rbw_values matches the number of columns in power_levels_df
    if len(rbw_values) != power_levels_df.shape[1]:
        console_print_func(f"Warning: Number of RBW values ({len(rbw_values)}) does not match number of power traces ({power_levels_df.shape[1]}). Using first RBW for all traces for PSD calculation.")
        debug_print(f"RBW count mismatch. Using first RBW for all traces.", file=current_file, function=current_function, console_print_func=console_print_func)
        # Fallback to using the first valid RBW for all traces if mismatch
        valid_rbw = next((rbw for rbw in rbw_values if rbw is not None and rbw > 0), None)
        if valid_rbw:
            rbw_values = [valid_rbw] * power_levels_df.shape[1]
        else:
            console_print_func("Error: No valid RBW found for PSD calculation. PSD will be NaN.")
            debug_print("No valid RBW found for PSD.", file=current_file, function=current_function, console_print_func=console_print_func)
            return psd_levels

    for i, col in enumerate(power_levels_df.columns):
        rbw = rbw_values[i]
        if rbw is not None and rbw > 0:
            psd_trace = 10 * np.log10(linear_power_mW_traces[col] / rbw)
            psd_traces.append(psd_trace)
            debug_print(f"Calculated PSD for trace {i+1} with RBW {rbw}.", file=current_file, function=current_function, console_print_func=console_print_func)
        else:
            psd_traces.append(pd.Series([np.nan] * len(power_levels_df.index), index=power_levels_df.index))
            console_print_func(f"Warning: RBW not available or invalid for trace {i+1} for PSD calculation. PSD for this trace will be NaN.")
            debug_print(f"Skipping PSD for trace {i+1} due to invalid RBW.", file=current_file, function=current_function, console_print_func=console_print_func)

    if psd_traces:
        combined_psd_df = pd.concat(psd_traces, axis=1)
        psd_levels = combined_psd_df.mean(axis=1)
        debug_print(f"Calculated multi-trace averaged PSD. First 5 values: {psd_levels.head().tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)
    else:
        console_print_func("No valid PSD data could be calculated for any trace. PSD column will be NaN.")
        debug_print("No valid PSD data for multi-trace average.", file=current_file, function=current_function, console_print_func=console_print_func)

    debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    return psd_levels


# --- Main Averaging Functions ---

def generate_current_cycle_average_csv_and_plot(
    collected_scans_dataframes,
    output_dir_base, # Renamed to output_dir_base to indicate it's the base folder
    include_tv_markers_var,
    include_gov_markers_var,
    open_html_after_complete_var,
    console_print_func,
    selected_avg_types, # NEW: Added selected_avg_types parameter
    scan_rbw_hz=None # Added RBW for PSD calculation, needed for current cycle
):
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Input collected_scans_dataframes (keys): {collected_scans_dataframes.keys() if collected_scans_dataframes else 'None'}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Input output_dir_base: {output_dir_base}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Input scan_rbw_hz: {scan_rbw_hz}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Input selected_avg_types: {selected_avg_types}", file=current_file, function=current_function, console_print_func=console_print_func)


    if not collected_scans_dataframes:
        console_print_func("No scan dataframes provided for current cycle averaging.")
        debug_print("No scan dataframes for current cycle averaging.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (no dataframes)", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    # Get the last collected scan name as a base for the output file
    # Assuming scan_name is part of the collected_scans_dataframes keys, or derive from app_instance
    # For simplicity, let's use a generic name or assume app_instance provides a scan_name_var
    # If not provided, a generic timestamped name will be used.
    scan_name_prefix = "CurrentScan" # Default prefix
    if collected_scans_dataframes:
        # Try to infer a scan name from the first collected scan if available
        first_scan_key = list(collected_scans_dataframes.keys())[0]
        # Attempt to extract a meaningful prefix, e.g., "MyScan_RBW100K_HOLD0_Offset0"
        match = re.match(r"([^\d_ -]+(?:[_ -]?[^\d_ -]+)*?)_?\d{8}_\d{6}", first_scan_key)
        if match:
            scan_name_prefix = re.sub(r"[_ -]+$", "", match.group(1).strip())
        else:
            scan_name_prefix = first_scan_key.split('_')[0] # Fallback to first part of name

    plot_title_only_datetime = datetime.now().strftime("%Y-%m-%d %H%M%S")
    
    # Use the new helper function to create the output subfolder
    output_dir_full = _create_output_subfolder(
        output_dir_base,
        f"{scan_name_prefix}_CurrentCycle", # Prefix for current cycle
        datetime.now().strftime("%Y%m%d_%H%M%S"), # Timestamp for subfolder
        console_print_func
    )
    debug_print(f"Output subfolder for current cycle: {output_dir_full}", file=current_file, function=current_function, console_print_func=console_print_func)


    aligned_dfs = []
    # Use the frequency axis of the first collected DataFrame as the reference
    # Ensure the first DataFrame has the expected columns, or assign them if not
    first_df_key = list(collected_scans_dataframes.keys())[0]
    first_df_raw = collected_scans_dataframes[first_df_key].copy() # Work on a copy

    # Assume no header and assign column names if they are not already present
    if 'Frequency (Hz)' not in first_df_raw.columns or 'Power (dBm)' not in first_df_raw.columns:
        if first_df_raw.shape[1] >= 2:
            first_df_raw.columns = ['Frequency (Hz)', 'Power (dBm)']
            debug_print(f"Assigned implied columns to first dataframe: {first_df_key}", file=current_file, function=current_function, console_print_func=console_print_func)
        else:
            console_print_func(f"Error: First collected dataframe '{first_df_key}' does not have enough columns for Frequency/Power. Skipping current cycle average.")
            debug_print(f"First dataframe '{first_df_key}' has insufficient columns.", file=current_file, function=current_function, console_print_func=console_print_func)
            return None, None

    # Ensure frequency column is numeric for the reference and drop duplicates
    first_df_raw['Frequency (Hz)'] = pd.to_numeric(first_df_raw['Frequency (Hz)'], errors='coerce')
    first_df_raw.dropna(subset=['Frequency (Hz)'], inplace=True)
    first_df_raw.drop_duplicates(subset=['Frequency (Hz)'], keep='first', inplace=True) # Ensure unique frequencies
    if first_df_raw.empty:
        console_print_func(f"Error: First collected dataframe '{first_df_key}' became empty after cleaning non-numeric or duplicate frequencies. Skipping current cycle average.")
        debug_print(f"First dataframe '{first_df_key}' empty after frequency cleanup.", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    first_df_freq = first_df_raw['Frequency (Hz)']
    debug_print(f"Reference frequency axis from first dataframe (first few points): {first_df_freq.head().tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)


    for df_name, df_original in collected_scans_dataframes.items():
        df = df_original.copy() # Work on a copy
        
        # Assume no header and assign column names if they are not already present
        if 'Frequency (Hz)' not in df.columns or 'Power (dBm)' not in df.columns:
            if df.shape[1] >= 2:
                df.columns = ['Frequency (Hz)', 'Power (dBm)']
                debug_print(f"Assigned implied columns to dataframe: {df_name}", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                console_print_func(f"Skipping {df_name}: Does not contain enough columns for Frequency/Power. Found {df.shape[1]} columns.")
                debug_print(f"Skipping {df_name} due to insufficient columns.", file=current_file, function=current_function, console_print_func=console_print_func)
                continue # Skip to the next dataframe

        # Ensure numeric types and drop NaNs for processing, then drop duplicates
        df['Frequency (Hz)'] = pd.to_numeric(df['Frequency (Hz)'], errors='coerce')
        df['Power (dBm)'] = pd.to_numeric(df['Power (dBm)'], errors='coerce')
        df.dropna(subset=['Frequency (Hz)', 'Power (dBm)'], inplace=True)
        df.drop_duplicates(subset=['Frequency (Hz)'], keep='first', inplace=True) # Ensure unique frequencies

        if df.empty:
            console_print_func(f"Warning: Dataframe {df_name} became empty after cleaning non-numeric or duplicate data. Skipping.")
            debug_print(f"Dataframe {df_name} empty after data cleanup.", file=current_file, function=current_function, console_print_func=console_print_func)
            continue

        # Reindex to the first_df_freq to align
        aligned_df = df.set_index('Frequency (Hz)').reindex(first_df_freq).reset_index()
        aligned_dfs.append(aligned_df.set_index('Frequency (Hz)')['Power (dBm)']) # Only power levels
        debug_print(f"Aligned {df_name} for power levels.", file=current_file, function=current_function, console_print_func=console_print_func)

    if not aligned_dfs:
        console_print_func("No valid scan data to average after alignment.")
        debug_print("No valid scan data to average after alignment.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (no aligned data)", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    power_levels_df = pd.concat(aligned_dfs, axis=1)
    power_levels_df.columns = [f"Scan_{i+1}" for i in range(len(aligned_dfs))] # Rename columns
    debug_print(f"Combined power_levels_df shape: {power_levels_df.shape}", file=current_file, function=current_function, console_print_func=console_print_func)

    # Prepare aggregated DataFrame for plotting and CSV export
    aggregated_df_columns = {'Frequency (Hz)': power_levels_df.index}
    debug_print(f"Calculating selected average types: {selected_avg_types}", file=current_file, function=current_function, console_print_func=console_print_func)

    if "Average" in selected_avg_types:
        aggregated_df_columns['Average (dBm)'] = _calculate_average(power_levels_df, console_print_func)
    if "Median" in selected_avg_types:
        aggregated_df_columns['Median (dBm)'] = _calculate_median(power_levels_df, console_print_func)
    if "Range" in selected_avg_types:
        aggregated_df_columns['Range (dBm)'] = _calculate_range(power_levels_df, console_print_func)
    if "Std Dev" in selected_avg_types:
        aggregated_df_columns['Std Dev (dBm)'] = _calculate_std_dev(power_levels_df, console_print_func)
    if "Variance" in selected_avg_types:
        aggregated_df_columns['Variance (dBm^2)'] = _calculate_variance(power_levels_df, console_print_func)
    if "PSD (dBm/Hz)" in selected_avg_types:
        # For current cycle, all traces share the same RBW from instrument settings
        # So we pass a list of the same RBW for _calculate_psd
        rbw_list_for_psd = [scan_rbw_hz] * power_levels_df.shape[1] if scan_rbw_hz is not None else [None] * power_levels_df.shape[1]
        aggregated_df_columns['PSD (dBm/Hz)'] = _calculate_psd(power_levels_df, rbw_list_for_psd, console_print_func)


    # Only create DataFrame with columns that were selected
    final_columns = ['Frequency (Hz)'] + [col for col in selected_avg_types if col in aggregated_df_columns]
    aggregated_df = pd.DataFrame({col: aggregated_df_columns[col] for col in final_columns}).reset_index(drop=True)

    debug_print(f"Aggregated DataFrame columns: {aggregated_df.columns.tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Aggregated DataFrame head:\n{aggregated_df.head()}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Aggregated DataFrame info:\n{aggregated_df.info()}", file=current_file, function=current_function, console_print_func=console_print_func)


    # --- Save to separate CSVs in the new subfolder ---
    try:
        if "Average" in selected_avg_types:
            average_csv_filename = os.path.join(output_dir_full, f"{scan_name_prefix}_CurrentCycle_{datetime.now().strftime('%Y%m%d_%H%M%S')}_average.csv")
            aggregated_df[['Frequency (Hz)', 'Average (dBm)']].to_csv(average_csv_filename, index=False, float_format='%.3f', header=False, sep=',')
            console_print_func(f"✅ Current cycle average data saved to: {average_csv_filename}")
            debug_print(f"Saved average CSV to: {average_csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

        if "Median" in selected_avg_types:
            median_csv_filename = os.path.join(output_dir_full, f"{scan_name_prefix}_CurrentCycle_{datetime.now().strftime('%Y%m%d_%H%M%S')}_median.csv")
            aggregated_df[['Frequency (Hz)', 'Median (dBm)']].to_csv(median_csv_filename, index=False, float_format='%.3f', header=False, sep=',')
            console_print_func(f"✅ Current cycle median data saved to: {median_csv_filename}")
            debug_print(f"Saved median CSV to: {median_csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

        if "Range" in selected_avg_types:
            range_csv_filename = os.path.join(output_dir_full, f"{scan_name_prefix}_CurrentCycle_{datetime.now().strftime('%Y%m%d_%H%M%S')}_range.csv")
            aggregated_df[['Frequency (Hz)', 'Range (dBm)']].to_csv(range_csv_filename, index=False, float_format='%.3f', header=False, sep=',')
            console_print_func(f"✅ Current cycle range data saved to: {range_csv_filename}")
            debug_print(f"Saved range CSV to: {range_csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

        if "Std Dev" in selected_avg_types:
            std_dev_csv_filename = os.path.join(output_dir_full, f"{scan_name_prefix}_CurrentCycle_{datetime.now().strftime('%Y%m%d_%H%M%S')}_std_dev.csv")
            aggregated_df[['Frequency (Hz)', 'Std Dev (dBm)']].to_csv(std_dev_csv_filename, index=False, float_format='%.3f', header=False, sep=',')
            console_print_func(f"✅ Current cycle standard deviation data saved to: {std_dev_csv_filename}")
            debug_print(f"Saved std dev CSV to: {std_dev_csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

        if "Variance" in selected_avg_types:
            variance_csv_filename = os.path.join(output_dir_full, f"{scan_name_prefix}_CurrentCycle_{datetime.now().strftime('%Y%m%d_%H%M%S')}_variance.csv")
            aggregated_df[['Frequency (Hz)', 'Variance (dBm^2)']].to_csv(variance_csv_filename, index=False, float_format='%.3f', header=False, sep=',')
            console_print_func(f"✅ Current cycle variance data saved to: {variance_csv_filename}")
            debug_print(f"Saved variance CSV to: {variance_csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

        if "PSD (dBm/Hz)" in selected_avg_types and 'PSD (dBm/Hz)' in aggregated_df.columns and not aggregated_df['PSD (dBm/Hz)'].isnull().all():
            psd_csv_filename = os.path.join(output_dir_full, f"{scan_name_prefix}_CurrentCycle_{datetime.now().strftime('%Y%m%d_%H%M%S')}_psd.csv")
            aggregated_df[['Frequency (Hz)', 'PSD (dBm/Hz)']].to_csv(psd_csv_filename, index=False, float_format='%.3f', header=False, sep=',')
            console_print_func(f"✅ Current cycle PSD data saved to: {psd_csv_filename}")
            debug_print(f"Saved PSD CSV to: {psd_csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)
        elif "PSD (dBm/Hz)" in selected_avg_types:
            console_print_func("PSD data is all NaN or not present, skipping PSD CSV export for current cycle.")
            debug_print("Skipping PSD CSV export for current cycle.", file=current_file, function=current_function, console_print_func=console_print_func)

    except Exception as e:
        console_print_func(f"❌ Failed to save current cycle CSVs: {e}")
        debug_print(f"Error saving current cycle CSVs: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (CSV save error)", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None


    # Generate plot
    html_filename = os.path.join(output_dir_full, f"{scan_name_prefix}_CurrentCycle_Plot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
    debug_print(f"Plot HTML filename: {html_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

    historical_dfs_for_overlays = [
        {'name': df_name, 'df': df}
        for df_name, df_original in collected_scans_dataframes.items()
    ]
    # Re-process historical_dfs_for_overlays to ensure they have correct headers if needed for plotting_utils
    processed_historical_dfs = []
    for item in historical_dfs_for_overlays:
        df = item['df'].copy()
        # Assume no header and assign column names if they are not already present
        if 'Frequency (Hz)' not in df.columns or 'Power (dBm)' not in df.columns:
            if df.shape[1] >= 2:
                df.columns = ['Frequency (Hz)', 'Power (dBm)']
                df['Frequency (Hz)'] = pd.to_numeric(df['Frequency (Hz)'], errors='coerce')
                df['Power (dBm)'] = pd.to_numeric(df['Power (dBm)'], errors='coerce')
                df.dropna(subset=['Frequency (Hz)', 'Power (dBm)'], inplace=True)
                df.drop_duplicates(subset=['Frequency (Hz)'], keep='first', inplace=True) # Ensure unique frequencies
                if not df.empty:
                    processed_historical_dfs.append({'name': item['name'], 'df': df})
                else:
                    console_print_func(f"Warning: Historical dataframe '{item['name']}' became empty after cleaning for plotting overlay. Skipping.")
            else:
                console_print_func(f"Warning: Historical dataframe '{item['name']}' has insufficient columns for plotting overlay. Skipping.")
        else:
            processed_historical_dfs.append(item) # Use original if already has correct columns

    debug_print(f"Prepared {len(processed_historical_dfs)} historical dataframes for overlays after processing.", file=current_file, function=current_function, console_print_func=console_print_func)


    fig, plot_html_path_return = plot_multi_trace_data(
        aggregated_df,
        f"{scan_name_prefix} - Averaged, Median, Range, Std Dev, Variance & PSD (Current Cycle {plot_title_only_datetime})",
        include_tv_markers_var.get(),
        include_gov_markers_var.get(),
        historical_dfs_with_names=processed_historical_dfs, # Pass the processed historical data
        output_html_path=html_filename,
        console_print_func=console_print_func
    )

    if fig:
        # fig.write_html is already called inside plot_multi_trace_data if output_html_path is provided
        console_print_func(f"✅ Current cycle averaged plot saved to: {plot_html_path_return}")
        if open_html_after_complete_var.get():
            _open_plot_in_browser(plot_html_path_return, console_print_func)
            debug_print(f"Opened plot in browser: {plot_html_path_return}", file=current_file, function=current_function, console_print_func=console_print_func)
    else:
        console_print_func("🚫 Plotly figure was not generated for current cycle averaged data.")
        debug_print("Plotly figure not generated for current cycle averaged data.", file=current_file, function=current_function, console_print_func=console_print_func)

    debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    return fig, plot_html_path_return


def generate_multi_file_average_and_plot(
    file_paths,
    selected_avg_types,
    plot_title_prefix,
    include_tv_markers,
    include_gov_markers,
    output_html_path_base, # Renamed to output_html_path_base for clarity
    open_html_after_complete,
    console_print_func
):
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Input file_paths ({len(file_paths)} files): {file_paths}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Input selected_avg_types: {selected_avg_types}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Input output_html_path_base: {output_html_path_base}", file=current_file, function=current_function, console_print_func=console_print_func)

    if not file_paths:
        console_print_func("No file paths provided for multi-file averaging.")
        debug_print("No file paths provided for multi-file averaging.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (no file paths)", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    all_scans_dfs = []
    # Regex to extract RBW and Offset from filename for PSD calculation and frequency normalization
    # This pattern expects filenames like "PREFIX_RBW10K_HOLD0_Offset0_YYYYMMDD_HHMMSS.csv"
    filename_pattern = re.compile(
        r'.*_RBW(?P<rbw_val>\d+K?)_HOLD\d+_(?:Offset(?P<offset_val>-?\d+))?_(?P<date_time>\d{8}_\d{6})\.csv$'
    )

    for df_idx, f_path in enumerate(file_paths): # Use enumerate to get index for debugging
        debug_print(f"Processing file {df_idx+1}/{len(file_paths)}: {os.path.basename(f_path)}", file=current_file, function=current_function, console_print_func=console_print_func)
        try:
            # Read CSV without header
            df = pd.read_csv(f_path, header=None)
            
            # Check if the DataFrame has at least two columns before assigning names
            if df.shape[1] < 2:
                console_print_func(f"Skipping {os.path.basename(f_path)}: CSV does not contain at least two columns (Frequency, Power). Found {df.shape[1]} columns.")
                debug_print(f"Skipping {os.path.basename(f_path)}: Insufficient columns in CSV.", file=current_file, function=current_function, console_print_func=console_print_func)
                continue # Skip to the next file

            # Assign implied column names
            df.columns = ['Frequency (Hz)', 'Power (dBm)']
            debug_print(f"Successfully read CSV: {os.path.basename(f_path)} with implied columns: {df.columns.tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)
            
            # Ensure 'Frequency (Hz)' column is numeric and handle potential errors during set_index
            df['Frequency (Hz)'] = pd.to_numeric(df['Frequency (Hz)'], errors='coerce')
            
            # Drop rows where 'Frequency (Hz)' became NaN due to coercion
            df.dropna(subset=['Frequency (Hz)'], inplace=True)
            # Drop duplicate frequencies, keeping the first
            df.drop_duplicates(subset=['Frequency (Hz)'], keep='first', inplace=True)

            if df.empty:
                console_print_func(f"Warning: File {os.path.basename(f_path)} became empty after cleaning non-numeric or duplicate frequencies. Skipping.")
                debug_print(f"File {os.path.basename(f_path)} empty after frequency cleanup.", file=current_file, function=current_function, console_print_func=console_print_func)
                continue


            file_name = os.path.basename(f_path)
            match = filename_pattern.match(file_name)
            debug_print(f"Regex match result for '{file_name}': {match}", file=current_file, function=current_function, console_print_func=console_print_func)
            
            rbw_hz = None
            current_offset_hz = 0.0

            if match:
                rbw_str = match.group('rbw_val')
                if 'K' in rbw_str:
                    rbw_hz = float(rbw_str.replace('K', '')) * 1000
                else:
                    rbw_hz = float(rbw_str)

                offset_str = match.group('offset_val')
                if offset_str:
                    current_offset_hz = float(offset_str)

                # Normalize frequency if an offset was applied
                df['Frequency (Hz)'] = df['Frequency (Hz)'] - current_offset_hz
                debug_print(f"File {file_name}: Extracted RBW={rbw_hz}, Offset={current_offset_hz}. Frequency normalized.", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                debug_print(f"File {file_name}: Filename pattern mismatch. RBW and Offset not extracted. Assuming no offset and default RBW for PSD.", file=current_file, function=current_function, console_print_func=console_print_func)
                console_print_func(f"Warning: Filename '{file_name}' did not match expected pattern for RBW/Offset. PSD calculation might be inaccurate.")
                # If pattern doesn't match, PSD will rely on default/NaN RBW later.

            df['RBW_Hz'] = rbw_hz # Add RBW to the dataframe for later PSD calculation
            all_scans_dfs.append(df)
            debug_print(f"Added DF from {os.path.basename(f_path)} to all_scans_dfs. Current count: {len(all_scans_dfs)}", file=current_file, function=current_function, console_print_func=console_print_func)
        except Exception as e:
            console_print_func(f"Error reading {os.path.basename(f_path)}: {e}")
            debug_print(f"Error reading {os.path.basename(f_path)}: {e}", file=current_file, function=current_function, console_print_func=console_print_func)

    if not all_scans_dfs:
        console_print_func("No valid scan data could be loaded from the selected files for averaging.")
        debug_print("No valid scan data could be loaded from the selected files for averaging.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (no valid data)", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    # Use the frequency axis of the first loaded DataFrame as the reference for alignment
    # Ensure reference_freq is from a valid DataFrame that was actually loaded
    try:
        reference_freq = all_scans_dfs[0]['Frequency (Hz)']
    except IndexError:
        console_print_func("Error: No valid dataframes available to establish a reference frequency. Cannot proceed.")
        debug_print("No valid dataframes to establish reference frequency.", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    debug_print(f"Reference frequency axis from first loaded file (first few points): {reference_freq.head().tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)
    power_levels_aligned = []
    rbw_values = [] # Collect RBW values for PSD calculation

    for df_idx, df in enumerate(all_scans_dfs): # Use enumerate for better debugging
        try:
            # Set 'Frequency (Hz)' as index, then reindex to the common reference_freq
            df_indexed = df.set_index('Frequency (Hz)')
            aligned_power_series = df_indexed['Power (dBm)'].reindex(reference_freq)

            if aligned_power_series.empty or aligned_power_series.isnull().all():
                console_print_func(f"Warning: Aligned power series for file {os.path.basename(file_paths[df_idx])} is empty or all NaNs after reindexing. Skipping this file for averaging.")
                debug_print(f"Aligned power series for file {os.path.basename(file_paths[df_idx])} is empty/NaNs after reindex.", file=current_file, function=current_function, console_print_func=console_print_func)
                continue # Skip this file if alignment results in empty/NaNs

            power_levels_aligned.append(aligned_power_series)
            rbw_values.append(df['RBW_Hz'].iloc[0]) # Assuming RBW is constant per file
            debug_print(f"Aligned dataframe {df_idx+1}'s power levels and collected RBW: {df['RBW_Hz'].iloc[0]}", file=current_file, function=current_function, console_print_func=console_print_func)

        except Exception as e:
            console_print_func(f"Error during alignment for file {os.path.basename(file_paths[df_idx])}: {e}. Skipping this file.")
            debug_print(f"Error during alignment for file {os.path.basename(file_paths[df_idx])}: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
            continue # Skip to the next file if an error occurs during alignment

    debug_print(f"After alignment loop: len(power_levels_aligned) = {len(power_levels_aligned)}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"After alignment loop: len(rbw_values) = {len(rbw_values)}", file=current_file, function=current_function, console_print_func=console_print_func)

    if not power_levels_aligned: # Check if anything was successfully aligned
        console_print_func("No dataframes successfully aligned for concatenation. Cannot proceed with averaging.")
        debug_print("No dataframes successfully aligned for concatenation.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (no aligned data for concat)", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None


    power_levels_df = pd.concat(power_levels_aligned, axis=1)
    power_levels_df.columns = [f"File_{i+1}" for i in range(len(power_levels_aligned))] # Name columns
    debug_print(f"Combined and aligned power_levels_df shape: {power_levels_df.shape}", file=current_file, function=current_function, console_print_func=console_print_func)

    # Prepare aggregated DataFrame for plotting with selected average types
    aggregated_df_columns = {'Frequency (Hz)': reference_freq}
    debug_print(f"Calculating selected average types: {selected_avg_types}", file=current_file, function=current_function, console_print_func=console_print_func)

    if "Average" in selected_avg_types:
        aggregated_df_columns['Average (dBm)'] = _calculate_average(power_levels_df, console_print_func)
    if "Median" in selected_avg_types:
        aggregated_df_columns['Median (dBm)'] = _calculate_median(power_levels_df, console_print_func)
    if "Range" in selected_avg_types:
        aggregated_df_columns['Range (dBm)'] = _calculate_range(power_levels_df, console_print_func)
    if "Std Dev" in selected_avg_types:
        aggregated_df_columns['Std Dev (dBm)'] = _calculate_std_dev(power_levels_df, console_print_func)
    if "Variance" in selected_avg_types:
        aggregated_df_columns['Variance (dBm^2)'] = _calculate_variance(power_levels_df, console_print_func)
    if "PSD (dBm/Hz)" in selected_avg_types:
        aggregated_df_columns['PSD (dBm/Hz)'] = _calculate_psd(power_levels_df, rbw_values, console_print_func)


    aggregated_df = pd.DataFrame(aggregated_df_columns).reset_index(drop=True)
    debug_print(f"Final aggregated_df columns: {aggregated_df.columns.tolist()}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Aggregated DataFrame head:\n{aggregated_df.head()}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Aggregated DataFrame info:\n{aggregated_df.info()}", file=current_file, function=current_function, console_print_func=console_print_func)


    # --- Create a new subfolder for this multi-file averaged plot's outputs ---
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    subfolder_name = f"{plot_title_prefix}_MultiFileAverage_{timestamp_str}"
    output_dir_full = _create_output_subfolder(
        output_html_path_base,
        f"{plot_title_prefix}_MultiFileAverage", # Prefix for multi-file average
        timestamp_str, # Timestamp for subfolder
        console_print_func
    )
    debug_print(f"Output subfolder for multi-file average: {output_dir_full}", file=current_file, function=current_function, console_print_func=console_print_func)

    # Define plot title
    plot_title_suffix = ", ".join(selected_avg_types)
    plot_title = f"{plot_title_prefix} - {plot_title_suffix} (Multi-File Average)"
    debug_print(f"Generated plot title: {plot_title}", file=current_file, function=current_function, console_print_func=console_print_func)

    # Define the actual HTML output path within the new subfolder
    html_output_path_final = os.path.join(output_dir_full, f"{plot_title_prefix}_MultiFileAverage_Plot_{timestamp_str}.html")
    debug_print(f"Final HTML output path: {html_output_path_final}", file=current_file, function=current_function, console_print_func=console_print_func)

    # --- Save aggregated data to separate CSVs in the new subfolder ---
    try:
        if "Average" in selected_avg_types:
            average_csv_filename = os.path.join(output_dir_full, f"AVERAGE_{plot_title_prefix}_MultiFileAverage_{timestamp_str}.csv")
            aggregated_df[['Frequency (Hz)', 'Average (dBm)']].to_csv(average_csv_filename, index=False, float_format='%.3f', header=False, sep=',')
            console_print_func(f"✅ Multi-file average data saved to: {average_csv_filename}")
            debug_print(f"Saved average CSV to: {average_csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

        if "Median" in selected_avg_types:
            median_csv_filename = os.path.join(output_dir_full, f"MEDIAN_{plot_title_prefix}_MultiFileAverage_{timestamp_str}.csv")
            aggregated_df[['Frequency (Hz)', 'Median (dBm)']].to_csv(median_csv_filename, index=False, float_format='%.3f', header=False, sep=',')
            console_print_func(f"✅ Multi-file median data saved to: {median_csv_filename}")
            debug_print(f"Saved median CSV to: {median_csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

        if "Range" in selected_avg_types:
            range_csv_filename = os.path.join(output_dir_full, f"RANGE_{plot_title_prefix}_MultiFileAverage_{timestamp_str}.csv")
            aggregated_df[['Frequency (Hz)', 'Range (dBm)']].to_csv(range_csv_filename, index=False, float_format='%.3f', header=False, sep=',')
            console_print_func(f"✅ Multi-file range data saved to: {range_csv_filename}")
            debug_print(f"Saved range CSV to: {range_csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

        if "Std Dev" in selected_avg_types:
            std_dev_csv_filename = os.path.join(output_dir_full, f"STDDEV_{plot_title_prefix}_MultiFileAverage_{timestamp_str}.csv")
            aggregated_df[['Frequency (Hz)', 'Std Dev (dBm)']].to_csv(std_dev_csv_filename, index=False, float_format='%.3f', header=False, sep=',')
            console_print_func(f"✅ Multi-file standard deviation data saved to: {std_dev_csv_filename}")
            debug_print(f"Saved std dev CSV to: {std_dev_csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

        if "Variance" in selected_avg_types:
            variance_csv_filename = os.path.join(output_dir_full, f"VARIANCE_{plot_title_prefix}_MultiFileAverage_{timestamp_str}.csv")
            aggregated_df[['Frequency (Hz)', 'Variance (dBm^2)']].to_csv(variance_csv_filename, index=False, float_format='%.3f', header=False, sep=',')
            console_print_func(f"✅ Multi-file variance data saved to: {variance_csv_filename}")
            debug_print(f"Saved variance CSV to: {variance_csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)

        if "PSD (dBm/Hz)" in selected_avg_types and 'PSD (dBm/Hz)' in aggregated_df.columns and not aggregated_df['PSD (dBm/Hz)'].isnull().all():
            psd_csv_filename = os.path.join(output_dir_full, f"PSD_{plot_title_prefix}_MultiFileAverage_{timestamp_str}.csv")
            aggregated_df[['Frequency (Hz)', 'PSD (dBm/Hz)']].to_csv(psd_csv_filename, index=False, float_format='%.3f', header=False, sep=',')
            console_print_func(f"✅ Multi-file PSD data saved to: {psd_csv_filename}")
            debug_print(f"Saved PSD CSV to: {psd_csv_filename}", file=current_file, function=current_function, console_print_func=console_print_func)
        elif "PSD (dBm/Hz)" in selected_avg_types:
            console_print_func("PSD data is all NaN or not present, skipping PSD CSV export for multi-file average.")
            debug_print("Skipping PSD CSV export for multi-file average.", file=current_file, function=current_function, console_print_func=console_print_func)

    except Exception as e:
        console_print_func(f"❌ Failed to save multi-file aggregated CSVs: {e}")
        debug_print(f"Error saving multi-file aggregated CSVs: {e}", file=current_file, function=current_function, console_print_func=console_print_func)


    # Generate plot using plot_multi_trace_data
    fig, plot_html_path_return = plot_multi_trace_data(
        aggregated_df,
        plot_title,
        include_tv_markers,
        include_gov_markers,
        historical_dfs_with_names=None, # No historical overlays for this multi-file average from external folder
        output_html_path=html_output_path_final, # Use the final path in the new subfolder
        console_print_func=console_print_func
    )

    if fig:
        # fig.write_html is already called inside plot_multi_trace_data if output_html_path is provided
        if open_html_after_complete:
            _open_plot_in_browser(plot_html_path_return, console_print_func)
            debug_print(f"Opened plot in browser: {plot_html_path_return}", file=current_file, function=current_function, console_print_func=console_print_func)
    elif not fig:
        console_print_func("🚫 Plotly figure was not generated for multi-file averaged data.")
        debug_print("Plotly figure not generated for multi-file averaged data.", file=current_file, function=current_function, console_print_func=console_print_func)

    debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
    return fig, plot_html_path_return
