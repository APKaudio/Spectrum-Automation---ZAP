# plotting_utils.py
#
# This module provides functions for generating interactive plots of spectrum analyzer data
# using Plotly. It supports plotting single scan traces, as well as aggregated data
# (average, median, range, standard deviation, variance, PSD) with historical overlays.
# It also includes functionalities for adding frequency band markers (TV, Government)
# and saving plots to HTML files.
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
import plotly.express as px
import plotly.graph_objects as go
import webbrowser
import os
import tkinter as tk # For messagebox
from tkinter import messagebox
import re # Added import for regular expressions
import csv # New: Import csv for MARKERS.CSV
import inspect # Import inspect for debug_print

# Import constants from frequency_bands.py
try:
    from utils.frequency_bands import ( # Changed to relative import
        MHZ_TO_HZ,
        TV_PLOT_BAND_MARKERS,
        GOV_PLOT_BAND_MARKERS
    )
except ImportError:
    print("Error: frequency_bands.py not found. Please ensure it's in the same directory.")
    # Define dummy values to prevent errors if file is missing
    MHZ_TO_HZ = 1_000_000
    TV_PLOT_BAND_MARKERS = []
    GOV_PLOT_BAND_MARKERS = []

# Import debug_print from instrument_control
from utils.instrument_control import debug_print


def _open_plot_in_browser(file_path):
    """
    Opens the generated HTML plot in the default web browser.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    try:
        webbrowser.open(f"file:///{os.path.abspath(file_path)}")
        debug_print(f"Opened plot in browser: {os.path.abspath(file_path)}", file=current_file, function=current_function)
    except Exception as e:
        messagebox.showerror("Browser Error", f"Failed to open plot in browser: {e}")
        debug_print(f"❌ Failed to open plot in browser: {e}", file=current_file, function=current_function)

def plot_single_scan_data(
    df,
    plot_title,
    output_html_path,
    include_tv_markers=False,
    include_gov_markers=False,
    include_markers_from_csv=False, # New parameter for custom markers
    markers_csv_path=None,          # Path to MARKERS.CSV
    y_range_min=-100,
    y_range_max=0
):
    """
    Generates an interactive Plotly HTML plot for a single scan's frequency vs. amplitude data.
    Includes options to overlay TV channel markers and Government band markers.

    Inputs:
        df (pd.DataFrame): DataFrame with 'Frequency (MHz)' and 'Amplitude (dBm)' columns.
        plot_title (str): Title of the plot.
        output_html_path (str): Full path to save the HTML plot.
        include_tv_markers (bool): Whether to include TV channel markers.
        include_gov_markers (bool): Whether to include Government band markers.
        include_markers_from_csv (bool): Whether to include custom markers from MARKERS.CSV.
        markers_csv_path (str): Path to the MARKERS.CSV file.
        y_range_min (int): Minimum Y-axis (Amplitude) value.
        y_range_max (int): Maximum Y-axis (Amplitude) value.

    Returns:
        tuple: A tuple containing the Plotly figure object and the output HTML path,
               or (None, None) if an error occurs.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if df.empty:
        debug_print("🚫 Cannot plot: DataFrame is empty.", file=current_file, function=current_function)
        return None, None

    # DataFrame should already have correct column names from scan_logic.py's pd.read_csv(..., names=...)
    debug_print(f"DataFrame columns received in plot_single_scan_data: {df.columns.tolist()}", file=current_file, function=current_function)
    debug_print(f"DataFrame head received in plot_single_scan_data:\n{df.head()}", file=current_file, function=current_function)

    # Check if the required columns exist
    if 'Frequency (MHz)' not in df.columns or 'Amplitude (dBm)' not in df.columns:
        debug_print("❌ Required columns 'Frequency (MHz)' or 'Amplitude (dBm)' are missing. Cannot plot.", file=current_file, function=current_function)
        return None, None

    fig = go.Figure()

    # Add the main scan trace
    fig.add_trace(go.Scatter(
        x=df['Frequency (MHz)'],
        y=df['Amplitude (dBm)'],
        mode='lines',
        name='Scan Trace',
        line=dict(color='cyan', width=2)
    ))

    # Add TV channel markers if enabled
    if include_tv_markers:
        for marker in TV_PLOT_BAND_MARKERS:
            start_freq = marker["Start MHz"]
            stop_freq = marker["Stop MHz"]
            band_name = marker["Band Name"]

            # Add shaded regions for TV bands
            fig.add_shape(
                type="rect",
                x0=start_freq, y0=y_range_min, x1=stop_freq, y1=y_range_max,
                fillcolor="rgba(255, 165, 0, 0.1)",  # Light orange, semi-transparent
                line_width=0,
                layer="below"
            )
            # Add text annotation for the band name
            fig.add_annotation(
                x=(start_freq + stop_freq) / 2,
                y=y_range_max - 5,  # Position slightly below max Y
                text=band_name,
                showarrow=False,
                font=dict(color="orange", size=8),
                bgcolor="rgba(0,0,0,0.5)",
                bordercolor="orange",
                borderwidth=0.5
            )

    # Add Government band markers if enabled
    if include_gov_markers:
        for marker in GOV_PLOT_BAND_MARKERS:
            start_freq = marker["Start MHz"]
            stop_freq = marker["Stop MHz"]
            band_name = marker["Band Name"]

            # Add shaded regions for Government bands
            fig.add_shape(
                type="rect",
                x0=start_freq, y0=y_range_min, x1=stop_freq, y1=y_range_max,
                fillcolor="rgba(144, 238, 144, 0.1)",  # Light green, semi-transparent
                line_width=0,
                layer="below"
            )
            # Add text annotation for the band name
            fig.add_annotation(
                x=(start_freq + stop_freq) / 2,
                y=y_range_max - 10,  # Position slightly below TV markers
                text=band_name,
                showarrow=False,
                font=dict(color="lightgreen", size=8),
                bgcolor="rgba(0,0,0,0.5)",
                bordercolor="lightgreen",
                borderwidth=0.5
            )

    # Add custom markers from MARKERS.CSV if enabled
    if include_markers_from_csv and markers_csv_path and os.path.exists(markers_csv_path):
        debug_print(f"Attempting to load custom markers from: {os.path.abspath(markers_csv_path)}", file=current_file, function=current_function)
        try:
            with open(markers_csv_path, mode='r', newline='', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                debug_print(f"MARKERS.CSV fieldnames: {reader.fieldnames}", file=current_file, function=current_function)

                # Ensure 'FREQ' column exists in the header
                if 'FREQ' not in reader.fieldnames:
                    debug_print(f"Error: MARKERS.CSV at {markers_csv_path} does not have a 'FREQ' column. Headers found: {reader.fieldnames}. Skipping custom markers.", file=current_file, function=current_function)
                else:
                    for row in reader:
                        try:
                            freq_mhz = float(row['FREQ'])
                            zone = row.get('ZONE', 'N/A')
                            group = row.get('GROUP', 'N/A')
                            device = row.get('DEVICE', 'N/A')
                            name = row.get('NAME', 'N/A')

                            # Add vertical dashed line
                            fig.add_vline(
                                x=freq_mhz,
                                line_dash="dash",
                                line_color="white",
                                line_width=3,
                                layer="above" # Ensure markers are visible above the trace
                            )
                            # Add text annotation
                            annotation_text = (
                                f"Zone: {zone}<br>"
                                f"Group: {group}<br>"
                                f"Device: {device}<br>"
                                f"Freq: {freq_mhz:.3f} MHz"
                            )
                            fig.add_annotation(
                                x=freq_mhz,
                                y=y_range_max, # Position at the top of the plot
                                text=annotation_text,
                                showarrow=False,
                                font=dict(color="white", size=9),
                                bgcolor="rgba(0,0,0,0.7)", # Semi-transparent black background
                                bordercolor="white",
                                borderwidth=0.5,
                                # Adjust xanchor/yanchor for positioning relative to the line
                                xanchor="left", # Anchor text to the left of the line
                                yanchor="top",  # Anchor text to the top of the plot
                                yshift=-5,      # Shift down slightly from the top edge
                                xshift=5        # Shift right slightly from the line
                            )
                        except ValueError:
                            debug_print(f"Warning: Could not parse frequency for marker row: {row}", file=current_file, function=current_function)
                        except KeyError as e:
                            debug_print(f"Warning: Missing expected column '{e}' in marker row: {row}", file=current_file, function=current_function)
            debug_print(f"Finished loading custom markers from {os.path.abspath(markers_csv_path)}.", file=current_file, function=current_function)
        except Exception as e:
            debug_print(f"Error loading MARKERS.CSV for plotting: {e}", file=current_file, function=current_function)
    elif include_markers_from_csv and not os.path.exists(markers_csv_path):
        debug_print(f"Info: MARKERS.CSV not found at {markers_csv_path}. Skipping custom markers.", file=current_file, function=current_function)


    fig.update_layout(
        title={
            'text': plot_title,
            'y':0.9,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top',
            'font': dict(size=16, color='white')
        },
        xaxis=dict(
            title='Frequency (MHz)',
            showgrid=True,
            gridcolor='rgba(255,255,255,0.1)',
            zeroline=False,
            range=[df['Frequency (MHz)'].min() if not df.empty else None,
                   df['Frequency (MHz)'].max() if not df.empty else None]
        ),
        yaxis=dict(
            title='Amplitude (dBm)',
            showgrid=True,
            gridcolor='rgba(255,255,255,0.1)',
            zeroline=False,
            autorange=False, # Explicitly set autorange to False
            range=[y_range_min, y_range_max]
        ),
        plot_bgcolor='black',
        paper_bgcolor='black',
        font=dict(color='white'),
        legend=dict(
            orientation="v", # Vertical orientation
            yanchor="top",   # Anchor to the top
            y=0.98,          # Position near the top right, adjusted slightly down
            xanchor="right", # Anchor to the right
            x=0.98,          # Position near the right edge
            bgcolor="rgba(0,0,0,0.5)", # Semi-transparent background for readability
            bordercolor="white",
            borderwidth=1,
            font=dict(size=9) # Slightly smaller font for compactness
        )
    )

    # If an output path is provided, save the figure
    if output_html_path:
        # Ensure the directory exists
        output_dir = os.path.dirname(output_html_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            debug_print(f"Created directory for plot: {output_dir}", file=current_file, function=current_function)
        fig.write_html(output_html_path, auto_open=False)
        debug_print(f"✅ Plot saved to: {output_html_path}", file=current_file, function=current_function)
        return fig, output_html_path
    else:
        debug_print("🚫 No output path provided, plot not saved.", file=current_file, function=current_function)
        return fig, None


def plot_multi_trace_data(
    aggregated_df,
    plot_title,
    include_tv_markers=False,
    include_gov_markers=False,
    historical_dfs_with_names=None,
    output_html_path=None,
    y_range_min=-100,
    y_range_max=0
):
    """
    Generates an interactive Plotly HTML plot for aggregated scan data (average, median, etc.)
    and can include historical overlays.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if aggregated_df.empty:
        debug_print("🚫 Cannot plot: Aggregated DataFrame is empty.", file=current_file, function=current_function)
        return None, None

    # DataFrame should already have correct column names from scan_logic.py's pd.read_csv(..., names=...)
    debug_print(f"Aggregated DataFrame columns received in plot_multi_trace_data: {aggregated_df.columns.tolist()}", file=current_file, function=current_function)
    debug_print(f"Aggregated DataFrame head received in plot_multi_trace_data:\n{aggregated_df.head()}", file=current_file, function=current_function)

    # Check if the required columns exist
    if 'Frequency (MHz)' not in aggregated_df.columns or 'Amplitude (dBm)' not in aggregated_df.columns:
        debug_print("❌ Required columns 'Frequency (MHz)' or 'Amplitude (dBm)' are missing in aggregated_df. Cannot plot.", file=current_file, function=current_function)
        return None, None

    fig = go.Figure()

    # Add traces for aggregated data
    if 'Average Amplitude (dBm)' in aggregated_df.columns:
        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency (MHz)'],
            y=aggregated_df['Average Amplitude (dBm)'],
            mode='lines',
            name='Average',
            line=dict(color='lime', width=2)
        ))
    if 'Median Amplitude (dBm)' in aggregated_df.columns:
        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency (MHz)'],
            y=aggregated_df['Median Amplitude (dBm)'],
            mode='lines',
            name='Median',
            line=dict(color='yellow', width=1, dash='dot')
        ))
    if 'Range (dB)' in aggregated_df.columns:
        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency (MHz)'],
            y=aggregated_df['Range (dB)'],
            mode='lines',
            name='Range',
            line=dict(color='orange', width=1, dash='dash')
        ))
    if 'Standard Deviation (dB)' in aggregated_df.columns:
        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency (MHz)'],
            y=aggregated_df['Standard Deviation (dB)'],
            mode='lines',
            name='Std Dev',
            line=dict(color='magenta', width=1, dash='longdash')
        ))
    if 'Variance (dB^2)' in aggregated_df.columns:
        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency (MHz)'],
            y=aggregated_df['Variance (dB^2)'],
            mode='lines',
            name='Variance',
            line=dict(color='purple', width=1, dash='shortdot')
        ))
    if 'PSD (dBm/Hz)' in aggregated_df.columns:
        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency (MHz)'],
            y=aggregated_df['PSD (dBm/Hz)'],
            mode='lines',
            name='PSD',
            line=dict(color='red', width=1, dash='solid')
        ))

    # Add historical overlays if provided
    if historical_dfs_with_names:
        for hist_df, hist_name in historical_dfs_with_names:
            debug_print(f"Historical DataFrame columns received for {hist_name}: {hist_df.columns.tolist()}", file=current_file, function=current_function)
            debug_print(f"Historical DataFrame head received for {hist_name}:\n{hist_df.head()}", file=current_file, function=current_function)

            # Check if the required columns exist
            if 'Frequency (MHz)' not in hist_df.columns or 'Amplitude (dBm)' not in hist_df.columns:
                debug_print(f"❌ Required columns 'Frequency (MHz)' or 'Amplitude (dBm)' are missing in historical_df {hist_name}. Skipping this historical plot.", file=current_file, function=current_function)
                continue # Skip this historical plot if columns are missing

            if 'Frequency (MHz)' in hist_df.columns and 'Average Amplitude (dBm)' in hist_df.columns:
                fig.add_trace(go.Scatter(
                    x=hist_df['Frequency (MHz)'],
                    y=hist_df['Average Amplitude (dBm)'],
                    mode='lines',
                    name=f'Historical Avg: {hist_name}',
                    line=dict(color='grey', width=1, dash='solid', opacity=0.7)
                ))

    # Add TV channel markers if enabled
    if include_tv_markers:
        for marker in TV_PLOT_BAND_MARKERS:
            start_freq = marker["Start MHz"]
            stop_freq = marker["Stop MHz"]
            band_name = marker["Band Name"]

            fig.add_shape(
                type="rect",
                x0=start_freq, y0=y_range_min, x1=stop_freq, y1=y_range_max,
                fillcolor="rgba(255, 165, 0, 0.1)",
                line_width=0,
                layer="below"
            )
            fig.add_annotation(
                x=(start_freq + stop_freq) / 2,
                y=y_range_max - 5,
                text=band_name,
                showarrow=False,
                font=dict(color="orange", size=8),
                bgcolor="rgba(0,0,0,0.5)",
                bordercolor="orange",
                borderwidth=0.5
            )

    # Add Government band markers if enabled
    if include_gov_markers:
        for marker in GOV_PLOT_BAND_MARKERS:
            start_freq = marker["Start MHz"]
            stop_freq = marker["Stop MHz"]
            band_name = marker["Band Name"]

            fig.add_shape(
                type="rect",
                x0=start_freq, y0=y_range_min, x1=stop_freq, y1=y_range_max,
                fillcolor="rgba(144, 238, 144, 0.1)",
                line_width=0,
                layer="below"
            )
            fig.add_annotation(
                x=(start_freq + stop_freq) / 2,
                y=y_range_max - 10,
                text=band_name,
                showarrow=False,
                font=dict(color="lightgreen", size=8),
                bgcolor="rgba(0,0,0,0.5)",
                bordercolor="lightgreen",
                borderwidth=0.5
            )

    fig.update_layout(
        title={
            'text': plot_title,
            'y':0.9,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top',
            'font': dict(size=16, color='white')
        },
        xaxis=dict(
            title='Frequency (MHz)',
            showgrid=True,
            gridcolor='rgba(255,255,255,0.1)',
            zeroline=False,
            range=[aggregated_df['Frequency (MHz)'].min() if not aggregated_df.empty else None,
                   aggregated_df['Frequency (MHz)'].max() if not aggregated_df.empty else None]
        ),
        yaxis=dict(
            title='Amplitude (dBm)',
            showgrid=True,
            gridcolor='rgba(255,255,255,0.1)',
            zeroline=False,
            autorange=False, # Explicitly set autorange to False
            range=[y_range_min, y_range_max]
        ),
        plot_bgcolor='black',
        paper_bgcolor='black',
        font=dict(color='white'),
        legend=dict(
            orientation="v", # Vertical orientation
            yanchor="top",   # Anchor to the top
            y=0.98,          # Position near the top right, adjusted slightly down
            xanchor="right", # Anchor to the right
            x=0.98,          # Position near the right edge
            bgcolor="rgba(0,0,0,0.5)", # Semi-transparent background for readability
            bordercolor="white",
            borderwidth=1,
            font=dict(size=9) # Slightly smaller font for compactness
        )
    )

    # If an output path is provided, save the figure
    if output_html_path:
        # Ensure the directory exists
        output_dir = os.path.dirname(output_html_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            debug_print(f"Created directory for plot: {output_dir}", file=current_file, function=current_function)
        fig.write_html(output_html_path, auto_open=False)
        debug_print(f"✅ Plot saved to: {output_html_path}", file=current_file, function=current_function)
        return fig, output_html_path
    else:
        debug_print("🚫 No output path provided, plot not saved.", file=current_file, function=current_function)
        return fig, None
