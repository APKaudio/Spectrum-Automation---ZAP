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
import tkinter as tk # For messagebox - KEEP for _open_plot_in_browser if it uses it
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
    print("Error: frequency_bands.py not found. Please ensure it's in the same directory.")
    # Define dummy values to prevent errors if file is missing
    MHZ_TO_HZ = 1_000_000
    TV_PLOT_BAND_MARKERS = []
    GOV_PLOT_BAND_MARKERS = []

# Import debug_print from instrument_control
from utils.instrument_control import debug_print # Import debug_print


def _open_plot_in_browser(file_path):
    """
    Opens the generated HTML plot in the default web browser.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    try:
        if os.path.exists(file_path):
            webbrowser.open_new_tab(f"file:///{os.path.abspath(file_path)}")
            debug_print(f"Opened plot in browser: {file_path}", file=current_file, function=current_function)
        else:
            print(f"🚫 Error: Plot file not found at {file_path}")
            debug_print(f"Plot file not found: {file_path}", file=current_file, function=current_function)
            # messagebox.showerror("File Not Found", f"Plot file not found: {file_path}") # Removed messagebox
    except Exception as e:
        print(f"❌ Error opening plot in browser: {e}")
        debug_print(f"Error opening plot in browser: {e}", file=current_file, function=current_function)
        # messagebox.showerror("Browser Error", f"Could not open plot in browser: {e}") # Removed messagebox


def plot_single_scan_data(df, title, include_tv_markers, include_gov_markers, include_custom_markers, output_folder_for_markers, output_html_path=None):
    """
    Generates an interactive Plotly HTML plot for a single scan DataFrame.
    Can include TV channel, Government band, and custom markers.

    Inputs:
        df (pd.DataFrame): DataFrame with 'Frequency (MHz)' and 'Level (dBm)' columns.
        title (str): Title of the plot.
        include_tv_markers (bool): Whether to include TV channel markers.
        include_gov_markers (bool): Whether to include Government band markers.
        include_custom_markers (bool): Whether to include custom markers from MARKERS.CSV.
        output_folder_for_markers (str): The base output folder to look for MARKERS.CSV.
        output_html_path (str, optional): Full path to save the HTML plot. If None, returns figure.

    Returns:
        tuple: (plotly.graph_objects.Figure, str) if output_html_path is provided,
               (plotly.graph_objects.Figure, None) if not.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if df.empty:
        print("🚫 Cannot plot: DataFrame is empty.")
        debug_print("Cannot plot: DataFrame is empty.", file=current_file, function=current_function)
        return None, None

    fig = go.Figure()

    # Add the main scan trace
    fig.add_trace(go.Scatter(x=df['Frequency (MHz)'], y=df['Level (dBm)'],
                             mode='lines', name='Scan Trace',
                             line=dict(color='cyan', width=1)))

    shapes = []
    annotations = []

    # Function to add markers
    def add_markers(markers, color, name_prefix, line_dash='dot'):
        for i, marker in enumerate(markers):
            start_freq = marker["Start MHz"]
            stop_freq = marker["Stop MHz"]
            band_name = marker["Band Name"]

            # Add shaded rectangle for the band
            shapes.append(
                dict(
                    type="rect",
                    xref="x", yref="paper",
                    x0=start_freq, y0=0,
                    x1=stop_freq, y1=1,
                    fillcolor=color,
                    opacity=0.1,
                    layer="below",
                    line_width=0,
                )
            )
            # Add annotation for the band name
            annotations.append(
                dict(
                    x=(start_freq + stop_freq) / 2,
                    y=1.02, # Position above the plot area
                    xref="x", yref="paper",
                    text=band_name,
                    showarrow=False,
                    font=dict(size=8, color=color),
                    xanchor="center", yanchor="bottom",
                    bgcolor="rgba(0,0,0,0.5)", # Semi-transparent background
                    bordercolor=color,
                    borderwidth=0.5,
                    borderpad=1
                )
            )

    # Add TV channel markers
    if include_tv_markers:
        debug_print("Including TV markers in plot.", file=current_file, function=current_function)
        add_markers(TV_PLOT_BAND_MARKERS, 'lightgreen', 'TV Channel')

    # Add Government band markers
    if include_gov_markers:
        debug_print("Including Government markers in plot.", file=current_file, function=current_function)
        add_markers(GOV_PLOT_BAND_MARKERS, 'pink', 'Gov Band')

    # Add custom markers from MARKERS.CSV
    if include_custom_markers:
        debug_print("Including custom markers from MARKERS.CSV in plot.", file=current_file, function=current_function)
        markers_file_path = os.path.join(output_folder_for_markers, "MARKERS.CSV")
        if os.path.exists(markers_file_path):
            try:
                custom_markers = []
                with open(markers_file_path, 'r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        try:
                            freq_mhz = float(row.get("FREQ"))
                            name = row.get("NAME", "Custom Marker")
                            custom_markers.append({"Band Name": name, "Start MHz": freq_mhz, "Stop MHz": freq_mhz})
                        except (ValueError, TypeError):
                            print(f"⚠️ Warning: Skipping invalid custom marker row: {row}")
                            debug_print(f"Skipping invalid custom marker row: {row}", file=current_file, function=current_function)
                            continue
                
                if custom_markers:
                    add_markers(custom_markers, 'yellow', 'Custom Marker')
                    debug_print(f"Added {len(custom_markers)} custom markers from MARKERS.CSV.", file=current_file, function=current_function)
                else:
                    print("ℹ️ MARKERS.CSV found but contains no valid custom marker data.")
                    debug_print("MARKERS.CSV found but contains no valid custom marker data.", file=current_file, function=current_function)

            except Exception as e:
                print(f"❌ Error loading custom markers from MARKERS.CSV: {e}")
                debug_print(f"Error loading custom markers from MARKERS.CSV: {e}", file=current_file, function=current_function)
        else:
            print(f"ℹ️ MARKERS.CSV not found at {markers_file_path}. Skipping custom markers.")
            debug_print(f"MARKERS.CSV not found at {markers_file_path}. Skipping custom markers.", file=current_file, function=current_function)


    # Determine x-axis range
    x_range_min = df['Frequency (MHz)'].min()
    x_range_max = df['Frequency (MHz)'].max()

    # Determine y-axis range (add some padding)
    y_range_min = df['Level (dBm)'].min() - 5
    y_range_max = df['Level (dBm)'].max() + 5

    fig.update_layout(
        title={
            'text': title,
            'y':0.95,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top',
            'font': dict(color='white') # Dark mode title color
        },
        xaxis=dict(
            title='Frequency (MHz)',
            showgrid=True,
            gridcolor='rgba(255,255,255,0.1)', # Light grid for dark mode
            zeroline=False,
            range=[x_range_min, x_range_max] if x_range_min is not None and x_range_max is not None else None
        ),
        yaxis=dict(
            title='Level (dBm)',
            showgrid=True,
            gridcolor='rgba(255,255,255,0.1)',
            zeroline=False,
            autorange=True,
            range=[y_range_min, y_range_max]
        ),
        plot_bgcolor='black', # Dark mode plot background
        paper_bgcolor='black', # Dark mode paper background
        font=dict(color='white'), # Default font color for labels, etc.
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
        ),
        # Add collected shapes and annotations to the layout
        shapes=shapes,
        annotations=annotations
    )
    debug_print("Plotly layout updated with traces, shapes, and annotations.", file=current_file, function=current_function)

    # If an output path is provided, save the figure
    if output_html_path:
        # Ensure the directory exists
        output_dir = os.path.dirname(output_html_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            debug_print(f"Created directory for plot: {output_dir}", file=current_file, function=current_function)
        debug_print(f"Saving plot to: {output_html_path}", file=current_file, function=current_function)
        fig.write_html(output_html_path, auto_open=False)
        debug_print(f"✅ Plot saved to: {output_html_path}", file=current_file, function=current_function)
        return fig, output_html_path
    else:
        debug_print("No output_html_path provided, returning Plotly figure directly.", file=current_file, function=current_function)
        return fig, None


def plot_multi_trace_data(aggregated_df, title, include_tv_markers, include_gov_markers, include_custom_markers, output_folder_for_markers, historical_dfs_with_names=None, output_html_path=None):
    """
    Generates an interactive Plotly HTML plot for multiple traces (e.g., average, median, historical).

    Inputs:
        aggregated_df (pd.DataFrame): DataFrame with aggregated data (e.g., 'Frequency (MHz)', 'Average Level (dBm)').
        title (str): Title of the plot.
        include_tv_markers (bool): Whether to include TV channel markers.
        include_gov_markers (bool): Whether to include Government band markers.
        include_custom_markers (bool): Whether to include custom markers from MARKERS.CSV.
        output_folder_for_markers (str): The base output folder to look for MARKERS.CSV.
        historical_dfs_with_names (list of tuples): List of (DataFrame, name) for historical overlays.
        output_html_path (str, optional): Full path to save the HTML plot. If None, returns figure.

    Returns:
        tuple: (plotly.graph_objects.Figure, str) if output_html_path is provided,
               (plotly.graph_objects.Figure, None) if not.
    """
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__

    if aggregated_df.empty:
        print("🚫 Cannot plot: Aggregated DataFrame is empty.")
        debug_print("Cannot plot: Aggregated DataFrame is empty.", file=current_file, function=current_function)
        return None, None

    fig = go.Figure()

    # Add aggregated traces
    if 'Average Level (dBm)' in aggregated_df.columns:
        fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['Average Level (dBm)'],
                                 mode='lines', name='Average Level', line=dict(color='cyan', width=2)))
    if 'Median Level (dBm)' in aggregated_df.columns:
        fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['Median Level (dBm)'],
                                 mode='lines', name='Median Level', line=dict(color='lightgreen', width=1, dash='dot')))
    if 'Max Level (dBm)' in aggregated_df.columns and 'Min Level (dBm)' in aggregated_df.columns:
        # Add a shaded area for the range
        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency (MHz)'].tolist() + aggregated_df['Frequency (MHz)'].tolist()[::-1],
            y=aggregated_df['Max Level (dBm)'].tolist() + aggregated_df['Min Level (dBm)'].tolist()[::-1],
            fill='toself',
            fillcolor='rgba(255,165,0,0.1)', # Light orange fill
            line=dict(color='rgba(255,255,255,0)'),
            name='Min/Max Range',
            hoverinfo='skip',
            showlegend=True
        ))
        fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['Max Level (dBm)'],
                                 mode='lines', name='Max Level', line=dict(color='orange', width=1, dash='dash')))
        fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['Min Level (dBm)'],
                                 mode='lines', name='Min Level', line=dict(color='yellow', width=1, dash='dash')))
    
    if 'Standard Deviation (dB)' in aggregated_df.columns:
        fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['Standard Deviation (dB)'],
                                 mode='lines', name='Std Dev (dB)', line=dict(color='magenta', width=1, dash='dash')))
    
    if 'Variance (dB²)' in aggregated_df.columns:
        fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['Variance (dB²)'],
                                 mode='lines', name='Variance (dB²)', line=dict(color='purple', width=1, dash='dash')))
    
    if 'PSD (dBm/Hz)' in aggregated_df.columns:
        fig.add_trace(go.Scatter(x=aggregated_df['Frequency (MHz)'], y=aggregated_df['PSD (dBm/Hz)'],
                                 mode='lines', name='PSD (dBm/Hz)', line=dict(color='white', width=1, dash='solid')))

    # Add historical overlays
    if historical_dfs_with_names:
        debug_print(f"Including {len(historical_dfs_with_names)} historical overlays.", file=current_file, function=current_function)
        colors = ['red', 'blue', 'green', 'yellow', 'purple', 'orange', 'pink'] # Cycle through colors
        for i, (hist_df, hist_name) in enumerate(historical_dfs_with_names):
            if 'Frequency (MHz)' in hist_df.columns and 'Average Level (dBm)' in hist_df.columns:
                fig.add_trace(go.Scatter(x=hist_df['Frequency (MHz)'], y=hist_df['Average Level (dBm)'],
                                         mode='lines', name=f'{hist_name} (Avg)',
                                         line=dict(color=colors[i % len(colors)], width=0.8, dash='dash')))
                debug_print(f"Added historical overlay: {hist_name}", file=current_file, function=current_function)


    shapes = []
    annotations = []

    # Function to add markers (re-using the logic from plot_single_scan_data)
    def add_markers(markers, color, name_prefix, line_dash='dot'):
        for i, marker in enumerate(markers):
            start_freq = marker["Start MHz"]
            stop_freq = marker["Stop MHz"]
            band_name = marker["Band Name"]

            # Add shaded rectangle for the band
            shapes.append(
                dict(
                    type="rect",
                    xref="x", yref="paper",
                    x0=start_freq, y0=0,
                    x1=stop_freq, y1=1,
                    fillcolor=color,
                    opacity=0.1,
                    layer="below",
                    line_width=0,
                )
            )
            # Add annotation for the band name
            annotations.append(
                dict(
                    x=(start_freq + stop_freq) / 2,
                    y=1.02, # Position above the plot area
                    xref="x", yref="paper",
                    text=band_name,
                    showarrow=False,
                    font=dict(size=8, color=color),
                    xanchor="center", yanchor="bottom",
                    bgcolor="rgba(0,0,0,0.5)", # Semi-transparent background
                    bordercolor=color,
                    borderwidth=0.5,
                    borderpad=1
                )
            )

    # Add TV channel markers
    if include_tv_markers:
        debug_print("Including TV markers in multi-trace plot.", file=current_file, function=current_function)
        add_markers(TV_PLOT_BAND_MARKERS, 'lightgreen', 'TV Channel')

    # Add Government band markers
    if include_gov_markers:
        debug_print("Including Government markers in multi-trace plot.", file=current_file, function=current_function)
        add_markers(GOV_PLOT_BAND_MARKERS, 'pink', 'Gov Band')

    # Add custom markers from MARKERS.CSV
    if include_custom_markers:
        debug_print("Including custom markers from MARKERS.CSV in multi-trace plot.", file=current_file, function=current_function)
        markers_file_path = os.path.join(output_folder_for_markers, "MARKERS.CSV")
        if os.path.exists(markers_file_path):
            try:
                custom_markers = []
                with open(markers_file_path, 'r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        try:
                            freq_mhz = float(row.get("FREQ"))
                            name = row.get("NAME", "Custom Marker")
                            custom_markers.append({"Band Name": name, "Start MHz": freq_mhz, "Stop MHz": freq_mhz})
                        except (ValueError, TypeError):
                            print(f"⚠️ Warning: Skipping invalid custom marker row: {row}")
                            debug_print(f"Skipping invalid custom marker row: {row}", file=current_file, function=current_function)
                            continue
                
                if custom_markers:
                    add_markers(custom_markers, 'yellow', 'Custom Marker')
                    debug_print(f"Added {len(custom_markers)} custom markers from MARKERS.CSV.", file=current_file, function=current_function)
                else:
                    print("ℹ️ MARKERS.CSV found but contains no valid custom marker data for multi-trace plot.")
                    debug_print("MARKERS.CSV found but contains no valid custom marker data for multi-trace plot.", file=current_file, function=current_function)

            except Exception as e:
                print(f"❌ Error loading custom markers from MARKERS.CSV for multi-trace plot: {e}")
                debug_print(f"Error loading custom markers from MARKERS.CSV for multi-trace plot: {e}", file=current_file, function=current_function)
        else:
            print(f"ℹ️ MARKERS.CSV not found at {markers_file_path}. Skipping custom markers for multi-trace plot.")
            debug_print(f"MARKERS.CSV not found at {markers_file_path}. Skipping custom markers for multi-trace plot.", file=current_file, function=current_function)


    # Determine x-axis range
    x_range_min = aggregated_df['Frequency (MHz)'].min()
    x_range_max = aggregated_df['Frequency (MHz)'].max()

    # Determine y-axis range (add some padding)
    all_levels = []
    if 'Average Level (dBm)' in aggregated_df.columns: all_levels.extend(aggregated_df['Average Level (dBm)'].tolist())
    if 'Median Level (dBm)' in aggregated_df.columns: all_levels.extend(aggregated_df['Median Level (dBm)'].tolist())
    if 'Max Level (dBm)' in aggregated_df.columns: all_levels.extend(aggregated_df['Max Level (dBm)'].tolist())
    if 'Min Level (dBm)' in aggregated_df.columns: all_levels.extend(aggregated_df['Min Level (dBm)'].tolist())
    if 'Standard Deviation (dB)' in aggregated_df.columns: all_levels.extend(aggregated_df['Standard Deviation (dB)'].tolist())
    if 'Variance (dB²)' in aggregated_df.columns: all_levels.extend(aggregated_df['Variance (dB²)'].tolist())
    if 'PSD (dBm/Hz)' in aggregated_df.columns: all_levels.extend(aggregated_df['PSD (dBm/Hz)'].tolist())

    for hist_df, _ in (historical_dfs_with_names if historical_dfs_with_names else []):
        if 'Average Level (dBm)' in hist_df.columns:
            all_levels.extend(hist_df['Average Level (dBm)'].tolist())

    if all_levels:
        y_range_min = min(all_levels) - 5
        y_range_max = max(all_levels) + 5
    else:
        y_range_min = -100
        y_range_max = 0 # Default if no data

    fig.update_layout(
        title={
            'text': title,
            'y':0.95,
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top',
            'font': dict(color='white') # Dark mode title color
        },
        xaxis=dict(
            title='Frequency (MHz)',
            showgrid=True,
            gridcolor='rgba(255,255,255,0.1)',
            zeroline=False,
            range=[x_range_min, x_range_max] if x_range_min is not None and x_range_max is not None else None
        ),
        yaxis=dict(
            title='Level (dBm)' if 'Level (dBm)' in aggregated_df.columns else 'Value', # Dynamic Y-axis title
            showgrid=True,
            gridcolor='rgba(255,255,255,0.1)',
            zeroline=False,
            autorange=False, # Set to False to use manual range
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
        ),
        # Add collected shapes and annotations to the layout
        shapes=shapes,
        annotations=annotations
    )
    debug_print("Plotly multi-trace layout updated with traces, shapes, and annotations.", file=current_file, function=current_function)

    # If an output path is provided, save the figure
    if output_html_path:
        # Ensure the directory exists
        output_dir = os.path.dirname(output_html_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            debug_print(f"Created directory for multi-trace plot: {output_dir}", file=current_file, function=current_function)
        debug_print(f"Saving multi-trace plot to: {output_html_path}", file=current_file, function=current_function)
        fig.write_html(output_html_path, auto_open=False)
        debug_print(f"✅ Multi-trace plot saved to: {output_html_path}", file=current_file, function=current_function)
        return fig, output_html_path
    else:
        debug_print("No output_html_path provided, returning Plotly multi-trace figure directly.", file=current_file, function=current_function)
        return fig, None

