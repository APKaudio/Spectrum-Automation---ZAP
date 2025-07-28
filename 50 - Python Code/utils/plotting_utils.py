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
# import tkinter as tk # For messagebox - KEEP for _open_plot_in_browser if it uses it # Removed
# from tkinter import messagebox # Removed messagebox import
import re # Added import for regular expressions
import csv # New: Import csv for MARKERS.CSV
import inspect # Import inspect for debug_print

# Import constants from frequency_bands.py
try:
    from ref.frequency_bands import ( # Changed to relative import
        MHZ_TO_HZ,
        TV_PLOT_BAND_MARKERS,
        GOV_PLOT_BAND_MARKERS
    )
except ImportError:
    # Fallback for direct script execution outside the package structure
    print("Could not import frequency_bands.py directly. Ensure it's in PYTHONPATH or run from project root.")
    MHZ_TO_HZ = 1_000_000
    TV_PLOT_BAND_MARKERS = []
    GOV_PLOT_BAND_MARKERS = []

# Import debug_print from instrument_control
from utils.instrument_control import debug_print


def _open_plot_in_browser(html_file_path, console_print_func):
    """
    Opens the generated HTML plot in the default web browser.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Attempting to open plot in browser: {html_file_path}", file=current_file, function=current_function, console_print_func=console_print_func)
    try:
        webbrowser.open(f'file://{os.path.realpath(html_file_path)}')
        console_print_func(f"✅ Plot opened in browser: {os.path.basename(html_file_path)}")
    except Exception as e:
        console_print_func(f"❌ Error opening plot in browser: {e}")
        debug_print(f"Error opening plot in browser: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
        # tk.messagebox.showerror("Browser Error", f"Could not open plot in browser: {e}") # Removed


def _load_markers_from_csv_for_plotting(output_folder, console_print_func):
    """
    Loads marker data from the MARKERS.CSV file for plotting purposes.
    The MARKERS.CSV is expected to be in the output_folder.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    
    markers_file_path = os.path.join(output_folder, "MARKERS.CSV")
    debug_print(f"Attempting to load markers from: {markers_file_path}", file=current_file, function=current_function, console_print_func=console_print_func)

    markers = []
    if os.path.exists(markers_file_path):
        try:
            with open(markers_file_path, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    try:
                        # Ensure FREQ is converted to float (MHz)
                        row['FREQ'] = float(row['FREQ'])
                        markers.append(row)
                    except ValueError:
                        console_print_func(f"⚠️ Warning: Invalid frequency value in MARKERS.CSV row: {row}. Skipping this marker.")
                        debug_print(f"Invalid FREQ in MARKERS.CSV: {row}", file=current_file, function=current_function, console_print_func=console_print_func)
            debug_print(f"Loaded {len(markers)} markers from MARKERS.CSV.", file=current_file, function=current_function, console_print_func=console_print_func)
        except Exception as e:
            console_print_func(f"❌ Error loading MARKERS.CSV for plotting: {e}")
            debug_print(f"Error loading MARKERS.CSV: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
            # tk.messagebox.showerror("File Error", f"Could not load MARKERS.CSV: {e}") # Removed
    else:
        debug_print(f"MARKERS.CSV not found at {markers_file_path}. No custom markers will be plotted.", file=current_file, function=current_function, console_print_func=console_print_func)
    return markers


def plot_single_scan_data(df, title, output_html_path=None, console_print_func=None):
    """
    Generates an interactive Plotly HTML plot for a single scan.

    Inputs:
        df (pd.DataFrame): DataFrame containing 'Frequency (MHz)' and 'Power Level (dBm)'.
        title (str): Title of the plot.
        output_html_path (str, optional): Path to save the HTML plot. If None, plot is not saved.
        console_print_func (function, optional): Function to use for console output.
    Returns:
        plotly.graph_objects.Figure: The Plotly figure object.
        str: The path where the HTML plot was saved (or None if not saved).
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Plotting single scan data with title: {title}", file=current_file, function=current_function, console_print_func=console_print_func)

    if df.empty:
        console_print_func("⚠️ Warning: DataFrame is empty. Cannot generate plot.")
        debug_print("Empty DataFrame for single scan plot.", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    # Ensure column names are correct
    if 'Frequency (MHz)' not in df.columns or 'Power Level (dBm)' not in df.columns:
        console_print_func("❌ Error: DataFrame must contain 'Frequency (MHz)' and 'Power Level (dBm)' columns.")
        debug_print("Missing required columns in DataFrame for single scan plot.", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    fig = px.line(df, x='Frequency (MHz)', y='Power Level (dBm)', 
                  title=title,
                  labels={'Frequency (MHz)': 'Frequency (MHz)', 'Power Level (dBm)': 'Power Level (dBm)'})

    fig.update_layout(
        template="plotly_dark", # Dark theme
        hovermode="x unified",
        xaxis_title="Frequency (MHz)",
        yaxis_title="Power Level (dBm)",
        font=dict(family="Inter, sans-serif"),
        margin=dict(l=40, r=40, t=80, b=40),
        height=600 # Fixed height for consistency
    )
    debug_print("Plotly figure created for single scan.", file=current_file, function=current_function, console_print_func=console_print_func)

    if output_html_path:
        # Ensure the directory exists
        output_dir = os.path.dirname(output_html_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            debug_print(f"Created directory for single scan plot: {output_dir}", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Saving single scan plot to: {output_html_path}", file=current_file, function=current_function, console_print_func=console_print_func)
        fig.write_html(output_html_path, auto_open=False)
        debug_print(f"✅ Single scan plot saved to: {output_html_path}", file=current_file, function=current_function, console_print_func=console_print_func)
        return fig, output_html_path
    else:
        debug_print("No output_html_path provided for single scan plot. Not saving to file.", file=current_file, function=current_function, console_print_func=console_print_func)
        return fig, None


def plot_multi_trace_data(aggregated_df, title, include_tv_markers, include_gov_markers, include_markers, historical_dfs_with_names=None, output_html_path=None, console_print_func=None):
    """
    Generates an interactive Plotly HTML plot for aggregated scan data,
    with options for historical overlays and various markers.

    Inputs:
        aggregated_df (pd.DataFrame): DataFrame containing aggregated data (e.g., 'Frequency (MHz)',
                                     'Average Power (dBm)', 'Median Power (dBm)', 'Range (dB)',
                                     'Standard Deviation (dB)', 'Variance (dB^2)', 'PSD (dBm/Hz)').
        title (str): Title of the plot.
        include_tv_markers (bool): Whether to include TV channel markers.
        include_gov_markers (bool): Whether to include Government band markers.
        include_markers (bool): Whether to include markers loaded from MARKERS.CSV.
        historical_dfs_with_names (list of dict, optional): List of dictionaries, each with
                                                            'name', 'df', 'x_col', 'y_col' for historical overlays.
        output_html_path (str, optional): Path to save the HTML plot. If None, plot is not saved.
        console_print_func (function, optional): Function to use for console output.
    Returns:
        plotly.graph_objects.Figure: The Plotly figure object.
        str: The path where the HTML plot was saved (or None if not saved).
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Plotting multi-trace data with title: {title}", file=current_file, function=current_function, console_print_func=console_print_func)

    if aggregated_df.empty:
        console_print_func("⚠️ Warning: Aggregated DataFrame is empty. Cannot generate plot.")
        debug_print("Empty aggregated DataFrame for multi-trace plot.", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    fig = go.Figure()

    # Add aggregated traces
    fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['Average Power (dBm)'],
                             mode='lines', name='Average Power (dBm)',
                             line=dict(color='cyan', width=2)))
    fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['Median Power (dBm)'],
                             mode='lines', name='Median Power (dBm)',
                             line=dict(color='lightgreen', width=1, dash='dot')))
    fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['Range (dB)'],
                             mode='lines', name='Range (dB)',
                             line=dict(color='orange', width=1, dash='dash')))
    fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['Standard Deviation (dB)'],
                             mode='lines', name='Standard Deviation (dB)',
                             line=dict(color='magenta', width=1, dash='dashdot')))
    fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['Variance (dB^2)'],
                             mode='lines', name='Variance (dB^2)',
                             line=dict(color='red', width=1, dash='longdash')))
    fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['PSD (dBm/Hz)'],
                             mode='lines', name='PSD (dBm/Hz)',
                             line=dict(color='yellow', width=2, dash='solid')))

    # Add historical raw scan overlays
    if historical_dfs_with_names:
        for i, hist_data in enumerate(historical_dfs_with_names):
            hist_df = hist_data['df']
            hist_name = hist_data['name']
            x_col = hist_data['x_col']
            y_col = hist_data['y_col']
            if not hist_df.empty and x_col in hist_df.columns and y_col in hist_df.columns:
                fig.add_trace(go.Scatter(x=hist_df[x_col], y=hist_df[y_col],
                                         mode='lines', name=hist_name,
                                         line=dict(color=f'rgba(100, 100, 255, {0.2 + i * 0.1})', width=0.5), # Faded blueish
                                         showlegend=True))
                debug_print(f"Added historical overlay: {hist_name}", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                debug_print(f"Skipping empty or malformed historical DataFrame: {hist_name}", file=current_file, function=current_function, console_print_func=console_print_func)


    shapes = []
    annotations = []

    # Add TV Channel Markers
    if include_tv_markers:
        for marker in TV_PLOT_BAND_MARKERS:
            # Draw vertical line
            shapes.append(dict(
                type="line",
                xref="x", yref="paper",
                x0=marker["Start MHz"], y0=0, x1=marker["Start MHz"], y1=1,
                line=dict(color="gray", width=1, dash="dot"),
                name=marker["Band Name"]
            ))
            shapes.append(dict(
                type="line",
                xref="x", yref="paper",
                x0=marker["Stop MHz"], y0=0, x1=marker["Stop MHz"], y1=1,
                line=dict(color="gray", width=1, dash="dot"),
                name=marker["Band Name"]
            ))
            # Add annotation for band name
            annotations.append(dict(
                x=(marker["Start MHz"] + marker["Stop MHz"]) / 2,
                y=1.02, # Position above the plot
                xref="x", yref="paper",
                text=marker["Band Name"],
                showarrow=False,
                font=dict(size=8, color="gray"),
                xanchor="center", yanchor="bottom"
            ))
        debug_print("Added TV channel markers to plot.", file=current_file, function=current_function, console_print_func=console_print_func)

    # Add Government/Commercial Band Markers
    if include_gov_markers:
        for marker in GOV_PLOT_BAND_MARKERS:
            # Draw vertical line
            shapes.append(dict(
                type="line",
                xref="x", yref="paper",
                x0=marker["Start MHz"], y0=0, x1=marker["Start MHz"], y1=1,
                line=dict(color="purple", width=1, dash="dot"),
                name=marker["Band Name"]
            ))
            shapes.append(dict(
                type="line",
                xref="x", yref="paper",
                x0=marker["Stop MHz"], y0=0, x1=marker["Stop MHz"], y1=1,
                line=dict(color="purple", width=1, dash="dot"),
                name=marker["Band Name"]
            ))
            # Add annotation for band name
            annotations.append(dict(
                x=(marker["Start MHz"] + marker["Stop MHz"]) / 2,
                y=1.05, # Position slightly higher than TV markers
                xref="x", yref="paper",
                text=marker["Band Name"],
                showarrow=False,
                font=dict(size=8, color="purple"),
                xanchor="center", yanchor="bottom"
            ))
        debug_print("Added Government/Commercial band markers to plot.", file=current_file, function=current_function, console_print_func=console_print_func)

    # Add Markers from MARKERS.CSV
    if include_markers:
        output_folder = os.path.dirname(output_html_path) if output_html_path else os.getcwd() # Assume current dir if no output path
        custom_markers = _load_markers_from_csv_for_plotting(output_folder, console_print_func) # Pass console_print_func
        for marker in custom_markers:
            freq_mhz = marker.get("FREQ")
            name = marker.get("NAME", "Custom Marker")
            if freq_mhz is not None:
                shapes.append(dict(
                    type="line",
                    xref="x", yref="paper",
                    x0=freq_mhz, y0=0, x1=freq_mhz, y1=1,
                    line=dict(color="red", width=2, dash="solid"),
                    name=name
                ))
                annotations.append(dict(
                    x=freq_mhz,
                    y=0.98, # Position near top of plot
                    xref="x", yref="paper",
                    text=f"{name} ({freq_mhz:.3f} MHz)",
                    showarrow=True,
                    arrowhead=2,
                    ax=0, ay=-30, # Arrow points down from text
                    font=dict(size=9, color="red"),
                    xanchor="center", yanchor="bottom"
                ))
        debug_print(f"Added {len(custom_markers)} custom markers from MARKERS.CSV to plot.", file=current_file, function=current_function, console_print_func=console_print_func)


    fig.update_layout(
        title={
            'text': title,
            'y':0.9,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top'
        },
        template="plotly_dark", # Dark theme
        hovermode="x unified",
        xaxis_title="Frequency (MHz)",
        yaxis_title="Power Level (dBm)",
        font=dict(family="Inter, sans-serif"),
        margin=dict(l=40, r=40, t=80, b=40),
        height=600, # Fixed height for consistency
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right", # Anchor to the right
            x=0.98,          # Position near the right edge
            bgcolor="rgba(0,0,0,0.5)", # Semi-transparent background for readability
            bordercolor="white",
            borderwidth=1,
            font=dict(size=9) # Slightly smaller font for compactness
        ),
        # Add collected shapes and annotations to the layout
        shapes=shapes,
        annotations=annotations
    )
    debug_print("Plotly multi-trace layout updated with traces, shapes, and annotations.", file=current_file, function=current_function, console_print_func=console_print_func)

    # If an output path is provided, save the figure
    if output_html_path:
        # Ensure the directory exists
        output_dir = os.path.dirname(output_html_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            debug_print(f"Created directory for multi-trace plot: {output_dir}", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Saving multi-trace plot to: {output_html_path}", file=current_file, function=current_function, console_print_func=console_print_func)
        fig.write_html(output_html_path, auto_open=False)
        debug_print(f"✅ Multi-trace plot saved to: {output_html_path}", file=current_file, function=current_function, console_print_func=console_print_func)
        return fig, output_html_path
    else:
        debug_print("No output_html_path provided for multi-trace plot. Not saving to file.", file=current_file, function=current_function, console_print_func=console_print_func)
        return fig, None

