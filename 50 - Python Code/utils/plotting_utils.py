# utils/plotting_utils.py (or src/plot_logic.py, assuming this is the full version used)
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
    # Fallback if frequency_bands.py is not in ref, assume it's directly accessible
    try:
        from frequency_bands import (
            MHZ_TO_HZ,
            TV_PLOT_BAND_MARKERS,
            GOV_PLOT_BAND_MARKERS
        )
    except ImportError:
        print("Error: frequency_bands.py not found in 'ref' or current directory.")
        # Define placeholders to prevent errors if file is completely missing
        MHZ_TO_HZ = 1_000_000
        TV_PLOT_BAND_MARKERS = []
        GOV_PLOT_BAND_MARKERS = []

# Assuming debug_print exists or define a dummy one for standalone testing
# This needs to be imported from utils.instrument_control if that's the standard
try:
    from utils.instrument_control import debug_print
except ImportError:
    print("Warning: debug_print not found in utils.instrument_control. Using dummy.")
    def debug_print(*args, **kwargs):
        pass # Dummy function


def _open_plot_in_browser(html_file_path, console_print_func=None):
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function} with html_file_path: {html_file_path}", file=current_file, function=current_function, console_print_func=console_print_func)
    try:
        if os.path.exists(html_file_path):
            webbrowser.open_new_tab(html_file_path)
            if console_print_func:
                console_print_func(f"Plot opened in browser: {html_file_path}")
            debug_print(f"Successfully opened plot: {html_file_path}", file=current_file, function=current_function, console_print_func=console_print_func)
        else:
            if console_print_func:
                console_print_func(f"Error: HTML file not found at {html_file_path}")
            debug_print(f"Error: HTML file not found at {html_file_path}", file=current_file, function=current_function, console_print_func=console_print_func)
    except Exception as e:
        if console_print_func:
            console_print_func(f"Failed to open plot in browser: {e}")
        debug_print(f"Failed to open plot in browser: {e}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)


def plot_single_scan_data(
    df,
    plot_title,
    include_tv_markers=False,
    include_gov_markers=False,
    output_html_path=None,
    y_range_min_override=None,
    y_range_max_override=0,
    console_print_func=None
):
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function} with plot_title: {plot_title}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"DataFrame shape: {df.shape}, include_tv_markers: {include_tv_markers}, include_gov_markers: {include_gov_markers}", file=current_file, function=current_function, console_print_func=console_print_func)

    if df.empty:
        if console_print_func:
            console_print_func("DataFrame is empty, cannot plot single scan.")
        debug_print("DataFrame is empty, cannot plot single scan.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (empty df)", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df['Frequency (Hz)'],
        y=df['Power (dBm)'],
        mode='lines',
        name='Current Scan',
        line=dict(color='blue')
    ))
    debug_print("Added main scan trace.", file=current_file, function=current_function, console_print_func=console_print_func)

    shapes = []
    annotations = []

    if include_tv_markers:
        debug_print("Adding TV Band Markers.", file=current_file, function=current_function, console_print_func=console_print_func)
        for band in TV_PLOT_BAND_MARKERS:
            if 'start_mhz' in band and 'end_mhz' in band:
                start_hz = band['start_mhz'] * MHZ_TO_HZ
                end_hz = band['end_mhz'] * MHZ_TO_HZ
                shapes.append(
                    dict(
                        type="rect", xref="x", yref="paper", x0=start_hz, y0=0, x1=end_hz, y1=1,
                        fillcolor=band.get('fill_color', "rgba(0,0,0,0)"), line=dict(color=band.get('line_color', "red"), width=1),
                        opacity=0.3, layer="below"
                    )
                )
                mid_hz = (start_hz + end_hz) / 2
                annotations.append(
                    dict(
                        x=mid_hz, y=1.02, xref="x", yref="paper", text=band['name'], showarrow=False,
                        font=dict(color=band.get('text_color', "red"), size=8), xanchor="center", yanchor="bottom"
                    )
                )
            else:
                debug_print(f"Warning: TV Band Marker {band.get('name', 'Unknown')} missing 'start_mhz' or 'end_mhz'.", file=current_file, function=current_function, console_print_func=console_print_func)


    if include_gov_markers:
        debug_print("Adding Government Band Markers.", file=current_file, function=current_function, console_print_func=console_print_func)
        for band in GOV_PLOT_BAND_MARKERS:
            if 'start_mhz' in band and 'end_mhz' in band:
                start_hz = band['start_mhz'] * MHZ_TO_HZ
                end_hz = band['end_mhz'] * MHZ_TO_HZ
                shapes.append(
                    dict(
                        type="rect", xref="x", yref="paper", x0=start_hz, y0=0, x1=end_hz, y1=1,
                        fillcolor=band.get('fill_color', "rgba(0,0,0,0)"), line=dict(color=band.get('line_color', "green"), width=1),
                        opacity=0.3, layer="below"
                    )
                )
                mid_hz = (start_hz + end_hz) / 2
                annotations.append(
                    dict(
                        x=mid_hz, y=1.05, xref="x", yref="paper", text=band['name'], showarrow=False,
                        font=dict(color=band.get('text_color', "green"), size=8), xanchor="center", yanchor="bottom"
                    )
                )
            else:
                debug_print(f"Warning: Government Band Marker {band.get('name', 'Unknown')} missing 'start_mhz' or 'end_mhz'.", file=current_file, function=current_function, console_print_func=console_print_func)


    fig.update_layout(
        title={'text': plot_title, 'y':0.9, 'x':0.5, 'xanchor': 'center', 'yanchor': 'top'},
        xaxis_title="Frequency (Hz)", yaxis_title="Power (dBm)", hovermode="x unified",
        template="plotly_dark", margin=dict(l=50, r=50, t=80, b=50), height=600,
        legend=dict(yanchor="top", y=0.99, xanchor="right", x=0.98, bgcolor="rgba(0,0,0,0.5)", bordercolor="white", borderwidth=1, font=dict(size=9)),
        shapes=shapes, annotations=annotations
    )
    debug_print("Plotly layout updated with shapes and annotations.", file=current_file, function=current_function, console_print_func=console_print_func)

    if y_range_min_override is not None or y_range_max_override is not None:
        y_axis_range = [y_range_min_override, y_range_max_override]
        fig.update_yaxes(range=y_axis_range)
        debug_print(f"Applied Y-axis range override: {y_axis_range}", file=current_file, function=current_function, console_print_func=console_print_func)


    if output_html_path:
        output_dir = os.path.dirname(output_html_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            debug_print(f"Created directory for plot: {output_dir}", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Saving plot to: {output_html_path}", file=current_file, function=current_function, console_print_func=console_print_func)
        fig.write_html(output_html_path, auto_open=False)
        debug_print(f"✅ Plot saved to: {output_html_path}", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
        return fig, output_html_path
    else:
        debug_print("No output_html_path provided, returning figure object only.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
        return fig, None


def plot_multi_trace_data(
    aggregated_df,
    plot_title,
    include_tv_markers=False,
    include_gov_markers=False,
    include_markers=True, # Added this to handle markers.csv
    historical_dfs_with_names=None,
    output_html_path=None,
    y_range_min_override=None,
    y_range_max_override=0,
    console_print_func=None
):
    current_function = inspect.currentframe().f_code.co_name
    current_file = __file__
    debug_print(f"Entering {current_function} with plot_title: {plot_title}", file=current_file, function=current_function, console_print_func=console_print_func)
    debug_print(f"Aggregated DataFrame shape: {aggregated_df.shape}, Historical DFS count: {len(historical_dfs_with_names) if historical_dfs_with_names else 0}", file=current_file, function=current_function, console_print_func=console_print_func)

    if aggregated_df.empty and not historical_dfs_with_names:
        if console_print_func:
            console_print_func("Aggregated DataFrame is empty and no historical data, cannot plot multi-trace data.")
        debug_print("Aggregated DataFrame is empty and no historical data.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function} (no data)", file=current_file, function=current_function, console_print_func=console_print_func)
        return None, None

    fig = go.Figure()

    if historical_dfs_with_names:
        debug_print("Adding historical overlays.", file=current_file, function=current_function, console_print_func=console_print_func)
        for i, item in enumerate(historical_dfs_with_names):
            df = item['df']
            name = item['name']
            if 'Frequency (Hz)' in df.columns and 'Power (dBm)' in df.columns:
                fig.add_trace(go.Scatter(
                    x=df['Frequency (Hz)'],
                    y=df['Power (dBm)'],
                    mode='lines',
                    name=f'{name} (Historical)',
                    line=dict(color=f'rgba(150, 150, 150, {0.5 - i * 0.1})', width=1),
                    showlegend=True
                ))
                debug_print(f"Added historical trace for: {name}", file=current_file, function=current_function, console_print_func=console_print_func)
            else:
                debug_print(f"Skipping historical DF '{name}': Missing 'Frequency (Hz)' or 'Power (dBm)'.", file=current_file, function=current_function, console_print_func=console_print_func)

    debug_print("Adding traces from aggregated_df.", file=current_file, function=current_function, console_print_func=console_print_func)
    for column in aggregated_df.columns:
        if column == 'Frequency (Hz)':
            continue
        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency (Hz)'],
            y=aggregated_df[column],
            mode='lines',
            name=column,
            line=dict(width=2)
        ))
        debug_print(f"Added aggregated trace for column: {column}", file=current_file, function=current_function, console_print_func=console_print_func)

    shapes = []
    annotations = []

    if include_tv_markers:
        debug_print("Adding TV Band Markers to multi-trace plot.", file=current_file, function=current_function, console_print_func=console_print_func)
        for band in TV_PLOT_BAND_MARKERS:
            if 'start_mhz' in band and 'end_mhz' in band:
                start_hz = band['start_mhz'] * MHZ_TO_HZ
                end_hz = band['end_mhz'] * MHZ_TO_HZ
                shapes.append(
                    dict(
                        type="rect", xref="x", yref="paper", x0=start_hz, y0=0, x1=end_hz, y1=1,
                        fillcolor=band.get('fill_color', "rgba(0,0,0,0)"), line=dict(color=band.get('line_color', "red"), width=1),
                        opacity=0.3, layer="below"
                    )
                )
                mid_hz = (start_hz + end_hz) / 2
                annotations.append(
                    dict(
                        x=mid_hz, y=1.02, xref="x", yref="paper", text=band['name'], showarrow=False,
                        font=dict(color=band.get('text_color', "red"), size=8), xanchor="center", yanchor="bottom"
                    )
                )
            else:
                debug_print(f"Warning: TV Band Marker {band.get('name', 'Unknown')} missing 'start_mhz' or 'end_mhz'.", file=current_file, function=current_function, console_print_func=console_print_func)

    if include_gov_markers:
        debug_print("Adding Government Band Markers to multi-trace plot.", file=current_file, function=current_function, console_print_func=console_print_func)
        for band in GOV_PLOT_BAND_MARKERS:
            if 'start_mhz' in band and 'end_mhz' in band:
                start_hz = band['start_mhz'] * MHZ_TO_HZ
                end_hz = band['end_mhz'] * MHZ_TO_HZ
                shapes.append(
                    dict(
                        type="rect", xref="x", yref="paper", x0=start_hz, y0=0, x1=end_hz, y1=1,
                        fillcolor=band.get('fill_color', "rgba(0,0,0,0)"), line=dict(color=band.get('line_color', "green"), width=1),
                        opacity=0.3, layer="below"
                    )
                )
                mid_hz = (start_hz + end_hz) / 2
                annotations.append(
                    dict(
                        x=mid_hz, y=1.05, xref="x", yref="paper", text=band['name'], showarrow=False,
                        font=dict(color=band.get('text_color', "green"), size=8), xanchor="center", yanchor="bottom"
                    )
                )
            else:
                debug_print(f"Warning: Government Band Marker {band.get('name', 'Unknown')} missing 'start_mhz' or 'end_mhz'.", file=current_file, function=current_function, console_print_func=console_print_func)


    fig.update_layout(
        title={'text': plot_title, 'y':0.9, 'x':0.5, 'xanchor': 'center', 'yanchor': 'top'},
        xaxis_title="Frequency (Hz)", yaxis_title="Power (dBm)", hovermode="x unified",
        template="plotly_dark", margin=dict(l=50, r=50, t=80, b=50), height=600,
        legend=dict(yanchor="top", y=0.99, xanchor="right", x=0.98, bgcolor="rgba(0,0,0,0.5)", bordercolor="white", borderwidth=1, font=dict(size=9)),
        shapes=shapes, annotations=annotations
    )
    debug_print("Plotly multi-trace layout updated with traces, shapes, and annotations.", file=current_file, function=current_function, console_print_func=console_print_func)

    if y_range_min_override is not None or y_range_max_override is not None:
        y_axis_range = [y_range_min_override, y_range_max_override]
        fig.update_yaxes(range=y_axis_range)
        debug_print(f"Applied Y-axis range override: {y_axis_range}", file=current_file, function=current_function, console_print_func=console_print_func)


    if output_html_path:
        output_dir = os.path.dirname(output_html_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            debug_print(f"Created directory for multi-trace plot: {output_dir}", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Saving multi-trace plot to: {output_html_path}", file=current_file, function=current_function, console_print_func=console_print_func)
        fig.write_html(output_html_path, auto_open=False)
        debug_print(f"✅ Multi-trace plot saved to: {output_html_path}", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
        return fig, output_html_path
    else:
        debug_print("No output_html_path provided, returning figure object only.", file=current_file, function=current_function, console_print_func=console_print_func)
        debug_print(f"Exiting {current_function}", file=current_file, function=current_function, console_print_func=console_print_func)
        return fig, None
