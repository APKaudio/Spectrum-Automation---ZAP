# plotting_utils.py

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import webbrowser
import os
import tkinter as tk # For messagebox
from tkinter import messagebox

# Import constants from frequency_bands.py
try:
    from frequency_bands import (
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


def _open_plot_in_browser(plot_path):
    """Helper to open an HTML plot in the default web browser."""
    try:
        webbrowser.open(plot_path)
        print(f"✅ Opened plot in browser: {plot_path}")
    except Exception as e:
        print(f"❌ Failed to open plot in browser: {e}")
        # Use a simple Tkinter messagebox for errors, as this function might be called from a thread
        # and direct console print might not be noticed by the user.
        # Ensure Tkinter root is initialized if not already (e.g., if called standalone)
        if not tk._default_root:
            root = tk.Tk()
            root.withdraw() # Hide the main window
            messagebox.showwarning("Open Plot Error", f"Could not open plot in web browser automatically: {e}")
            root.destroy()
        else:
            messagebox.showwarning("Open Plot Error", f"Could not open plot in web browser automatically: {e}")


def plot_single_scan_data(scanned_data, plot_title_suffix, include_tv_markers=True, include_gov_markers=True):
    """
    Generates an interactive Plotly HTML plot from scanned data,
    including overlays for TV and Government frequency bands based on flags.
    This function now only returns the Plotly figure object. The saving and
    opening of the HTML file are handled by the calling function.
    """
    if not scanned_data:
        print("No data to plot.")
        return None # Return None for fig

    # Ensure scanned_data is in the expected (Frequency_Hz, Power_dBm) tuple format
    df = pd.DataFrame(scanned_data, columns=['Frequency_Hz', 'Power_dBm'])
    df['Frequency_MHz'] = df['Frequency_Hz'] / MHZ_TO_HZ

    fig = px.line(df, x='Frequency_MHz', y='Power_dBm', 
                  title=f'RF Spectrum Scan ({plot_title_suffix})',
                  labels={'Frequency_MHz': 'Frequency (MHz)', 'Power_dBm': 'Power (dBm)'},
                  line_shape='linear') # 'linear' connects points with straight lines

    # Determine Y-axis range for marker positioning
    y_range_min = df['Power_dBm'].min()
    # Set max Y-axis to 0 dBm as requested
    y_range_max = 0 # Fixed maximum to 0 dBm
    # Add some padding to the y-range for better text visibility (only to min, max is fixed)
    y_padding = (y_range_max - y_range_min) * 0.1
    y_range_min -= y_padding
    
    fig.update_layout(yaxis_range=[y_range_min, y_range_max])


    # Determine X-axis range for marker visibility check
    x_min_data = df['Frequency_MHz'].min()
    x_max_data = df['Frequency_MHz'].max()

    # --- Add TV Band Markers ---
    if include_tv_markers:
        # Define colors for the TV band markers and text
        tv_marker_line_color = "rgba(255, 255, 0, 0.7)"  # Bright yellow, semi-transparent
        tv_marker_text_color = "yellow"
        tv_band_fill_color = "rgba(255, 255, 0, 0.05)"    # Very light yellow, highly transparent fill

        for band in TV_PLOT_BAND_MARKERS:
            # Only add markers if they are within the actual scanned frequency range
            if band["Start MHz"] < x_max_data and band["Stop MHz"] > x_min_data:
                # Add a shaded rectangle to represent the frequency band allocation
                fig.add_shape(
                    type="rect",
                    x0=band["Start MHz"],
                    y0=y_range_min,  # Span full Y-axis range
                    x1=band["Stop MHz"],
                    y1=y_range_max,  # Span full Y-axis range
                    line=dict(
                        color=tv_marker_line_color,
                        width=0.3,
                        dash="dot",
                    ),
                    fillcolor=tv_band_fill_color,
                    layer="below",
                )

                # Add text markers using go.Scatter with mode='text'
                x_center = (band["Start MHz"] + band["Stop MHz"]) / 2
                y_text_position = y_range_max - (y_range_max - y_range_min) * 0.05

                fig.add_trace(go.Scatter(
                    x=[x_center],
                    y=[y_text_position],
                    mode='text',
                    text=[f"{band['Band Name']}<br>{band['Start MHz']:.1f}-{band['Stop MHz']:.1f} MHz"],
                    textfont=dict(
                        size=8,
                        color=tv_marker_text_color
                    ),
                    showlegend=False,
                    hoverinfo='text',
                    name=f"Band Label: {band['Band Name']}"
                ))

    # --- Add Government Band Markers ---
    if include_gov_markers:
        # Define colors for the Government band markers and text
        gov_marker_line_color = "rgba(255, 0, 0, 0.9)"  # Red, semi-transparent
        gov_marker_text_color = "red"
        gov_band_fill_color = "rgba(255, 0, 0, 0.1)"    # Very light red, highly transparent fill

        # Define the four y-offsets for staggering
        y_offset_level_1 = 0.20
        y_offset_level_2 = 0.25
        y_offset_level_3 = 0.30
        y_offset_level_4 = 0.35
        y_offset_levels = [y_offset_level_1, y_offset_level_2, y_offset_level_3, y_offset_level_4]

        for i, band in enumerate(GOV_PLOT_BAND_MARKERS):
            # Only add markers if they are within the actual scanned frequency range
            if band["Start MHz"] < x_max_data and band["Stop MHz"] > x_min_data:
                # Add a shaded rectangle to represent the frequency band allocation
                fig.add_shape(
                    type="rect",
                    x0=band["Start MHz"],
                    y0=y_range_min,
                    x1=band["Stop MHz"],
                    y1=y_range_max,
                    line=dict(
                        color=gov_marker_line_color,
                        width=0.3,
                        dash="dot",
                    ),
                    fillcolor=gov_band_fill_color,
                    layer="below",
                )

                # Add text markers using go.Scatter with mode='text'
                x_center = (band["Start MHz"] + band["Stop MHz"]) / 2

                # Determine the Y position based on staggering using modulo for 4 levels
                current_y_offset = y_offset_levels[i % len(y_offset_levels)]
                y_text_position = y_range_max - (y_range_max - y_range_min) * current_y_offset

                fig.add_trace(go.Scatter(
                    x=[x_center],
                    y=[y_text_position],
                    mode='text',
                    text=[f"{band['Band Name']}<br>{band['Start MHz']:.1f}-{band['Stop MHz']:.1f} MHz"],
                    textfont=dict(
                        size=8,
                        color=gov_marker_text_color
                    ),
                    showlegend=False,
                    hoverinfo='text',
                    name=f"Band Label: {band['Band Name']}"
                ))

    # Apply Dark Mode Theme
    fig.update_layout(template="plotly_dark")

    fig.update_layout(hovermode="x unified") # Show all traces on hover

    # This function now only returns the figure, the filename is handled by the caller.
    return fig

def plot_multi_trace_data(aggregated_df, plot_title_suffix, include_tv_markers=True, include_gov_markers=True, historical_dfs_with_names=None, output_html_path=None):
    """
    Generates an interactive Plotly HTML plot from aggregated scan data (Average, Median, Range),
    including overlays for TV and Government frequency bands based on flags.
    Optionally includes historical scan data as additional layers.

    Args:
        aggregated_df (pd.DataFrame): DataFrame with 'Frequency_Hz', 'Average_dBm', 'Median_dBm', 'Range_dBm'.
        plot_title_suffix (str): Suffix for the plot title.
        include_tv_markers (bool): Whether to include TV band markers.
        include_gov_markers (bool): Whether to include Government band markers.
        historical_dfs_with_names (list of tuple): Optional. List of (DataFrame, name) tuples for historical scans.
        output_html_path (str): The full path including filename for the HTML output.
    Returns:
        tuple: (plotly.graph_objects.Figure, str) The Plotly figure object and the output HTML path.
    """
    if aggregated_df.empty and not historical_dfs_with_names:
        print("No data to plot (neither aggregated nor historical).")
        return None, None

    fig = go.Figure()

    # Add Average trace (RED)
    if not aggregated_df.empty:
        # The 'Frequency_MHz' column should already be present in aggregated_df
        # from the generate_average_plot function.
        # This check is a safeguard but should ideally not be needed if data preparation is correct.
        if 'Frequency_MHz' not in aggregated_df.columns:
            # If it's truly missing, and Frequency_Hz is present, derive it.
            # This indicates an an issue in data preparation before calling this function.
            if 'Frequency_Hz' in aggregated_df.columns:
                aggregated_df['Frequency_MHz'] = aggregated_df['Frequency_Hz'] / MHZ_TO_HZ
            else:
                print("Error: Neither 'Frequency_MHz' nor 'Frequency_Hz' found in aggregated_df.")
                return None, None # Cannot plot without frequency data

        fig.add_trace(go.Scatter(x=aggregated_df['Frequency_MHz'], y=aggregated_df['Average_dBm'],
                                 mode='lines', name='Average Power (dBm)',
                                 line=dict(color='red', width=3))) # Changed to RED

        # Add Median trace (YELLOW)
        fig.add_trace(go.Scatter(x=aggregated_df['Frequency_MHz'], y=aggregated_df['Median_dBm'],
                                 mode='lines', name='Median Power (dBm)',
                                 line=dict(color='yellow', width=2))) # Changed to YELLOW

        # Add Range trace (GREEN)
        fig.add_trace(go.Scatter(x=aggregated_df['Frequency_MHz'], y=aggregated_df['Range_dBm'],
                                 mode='lines', name='Range (Max - Min) (dB)',
                                 line=dict(color='green', width=2))) # Changed to GREEN

    # Add historical data as additional layers
    if historical_dfs_with_names:
        for hist_df, hist_name in historical_dfs_with_names:
            # FIX: Ensure we are operating on a copy to avoid SettingWithCopyWarning
            hist_df_copy = hist_df.copy() 
            # The 'Frequency_MHz' column should already be present in hist_df_copy
            if 'Frequency_MHz' not in hist_df_copy.columns:
                # If it's truly missing, and Frequency_Hz is present, derive it.
                if 'Frequency_Hz' in hist_df_copy.columns:
                    hist_df_copy['Frequency_MHz'] = hist_df_copy['Frequency_Hz'] / MHZ_TO_HZ
                else:
                    print(f"Error: Neither 'Frequency_MHz' nor 'Frequency_Hz' found in historical_df for {hist_name}. Skipping.")
                    continue # Skip this historical DataFrame if frequency data is missing

            fig.add_trace(go.Scatter(x=hist_df_copy['Frequency_MHz'], y=hist_df_copy['Power_dBm'],
                                     mode='lines', name=f"Scan: {hist_name}",
                                     line=dict(color='rgba(100, 100, 255, 0.5)', width=1, dash='dot'), # Lighter, thinner, dashed
                                     showlegend=True)) # Ensure legend entry for each historical scan


    fig.update_layout(title=f'RF Spectrum Scan - {plot_title_suffix}',
                      xaxis_title='Frequency (MHz)',
                      yaxis_title='Power / Range (dBm)', # Y-axis label accommodates range and power
                      template="plotly_dark",
                      hovermode="x unified")

    # Determine Y-axis range for marker positioning, considering all plotted traces
    all_y_values = []
    if not aggregated_df.empty:
        all_y_values.extend(aggregated_df['Average_dBm'].tolist())
        all_y_values.extend(aggregated_df['Median_dBm'].tolist())
        all_y_values.extend(aggregated_df['Range_dBm'].tolist()) # Include Range_dBm in max calculation
    if historical_dfs_with_names:
        for hist_df, _ in historical_dfs_with_names:
            # Ensure 'Power_dBm' exists before extending
            if 'Power_dBm' in hist_df.columns:
                all_y_values.extend(hist_df['Power_dBm'].tolist())

    if all_y_values:
        # Calculate y_range_max as the maximum of all relevant y-values, ensuring it's at least 0
        y_range_max = max(0, max(all_y_values))
    else:
        y_range_max = 0 # Default if no data

    y_range_min = min(all_y_values) if all_y_values else -100 # Default if no data
    y_padding = (y_range_max - y_range_min) * 0.1
    y_range_min -= y_padding
    
    fig.update_layout(yaxis_range=[y_range_min, y_range_max]) # Apply the updated Y-axis range


    # Determine X-axis range for marker visibility check
    x_min_data = float('inf')
    x_max_data = float('-inf')

    if not aggregated_df.empty:
        if 'Frequency_MHz' in aggregated_df.columns:
            x_min_data = min(x_min_data, aggregated_df['Frequency_MHz'].min())
            x_max_data = max(x_max_data, aggregated_df['Frequency_MHz'].max())
    if historical_dfs_with_names:
        for hist_df, _ in historical_dfs_with_names:
            if 'Frequency_MHz' in hist_df.columns:
                x_min_data = min(x_min_data, hist_df['Frequency_MHz'].min())
                x_max_data = max(x_max_data, hist_df['Frequency_MHz'].max())

    # --- Add TV Band Markers ---
    if include_tv_markers:
        # Define colors for the TV band markers and text
        tv_marker_line_color = "rgba(255, 255, 0, 0.7)"  # Bright yellow, semi-transparent
        tv_marker_text_color = "yellow"
        tv_band_fill_color = "rgba(255, 255, 0, 0.05)"    # Very light yellow, highly transparent fill

        for band in TV_PLOT_BAND_MARKERS:
            # Only add markers if they are within the actual scanned frequency range
            if band["Start MHz"] < x_max_data and band["Stop MHz"] > x_min_data:
                # Add a shaded rectangle to represent the frequency band allocation
                fig.add_shape(
                    type="rect",
                    x0=band["Start MHz"],
                    y0=y_range_min,  # Span full Y-axis range
                    x1=band["Stop MHz"],
                    y1=y_range_max,  # Span full Y-axis range
                    line=dict(
                        color=tv_marker_line_color,
                        width=0.3,
                        dash="dot",
                    ),
                    fillcolor=tv_band_fill_color,
                    layer="below",
                )

                # Add text markers using go.Scatter with mode='text'
                x_center = (band["Start MHz"] + band["Stop MHz"]) / 2
                y_text_position = y_range_max - (y_range_max - y_range_min) * 0.05

                fig.add_trace(go.Scatter(
                    x=[x_center],
                    y=[y_text_position],
                    mode='text',
                    text=[f"{band['Band Name']}<br>{band['Start MHz']:.1f}-{band['Stop MHz']:.1f} MHz"],
                    textfont=dict(
                        size=8,
                        color=tv_marker_text_color
                    ),
                    showlegend=False,
                    hoverinfo='text',
                    name=f"Band Label: {band['Band Name']}"
                ))

    # --- Add Government Band Markers ---
    if include_gov_markers:
        # Define colors for the Government band markers and text
        gov_marker_line_color = "rgba(255, 0, 0, 0.9)"  # Red, semi-transparent
        gov_marker_text_color = "red"
        gov_band_fill_color = "rgba(255, 0, 0, 0.1)"    # Very light red, highly transparent fill

        # Define the four y-offsets for staggering
        y_offset_level_1 = 0.20
        y_offset_level_2 = 0.25
        y_offset_level_3 = 0.30
        y_offset_level_4 = 0.35
        y_offset_levels = [y_offset_level_1, y_offset_level_2, y_offset_level_3, y_offset_level_4]

        for i, band in enumerate(GOV_PLOT_BAND_MARKERS):
            # Only add markers if they are within the actual scanned frequency range
            if band["Start MHz"] < x_max_data and band["Stop MHz"] > x_min_data:
                # Add a shaded rectangle to represent the frequency band allocation
                fig.add_shape(
                    type="rect",
                    x0=band["Start MHz"],
                    y0=y_range_min,
                    x1=band["Stop MHz"],
                    y1=y_range_max,
                    line=dict(
                        color=gov_marker_line_color,
                        width=0.3,
                        dash="dot",
                    ),
                    fillcolor=gov_band_fill_color,
                    layer="below",
                )

                # Add text markers using go.Scatter with mode='text'
                x_center = (band["Start MHz"] + band["Stop MHz"]) / 2

                # Determine the Y position based on staggering using modulo for 4 levels
                current_y_offset = y_offset_levels[i % len(y_offset_levels)]
                y_text_position = y_range_max - (y_range_max - y_range_min) * current_y_offset

                fig.add_trace(go.Scatter(
                    x=[x_center],
                    y=[y_text_position],
                    mode='text',
                    text=[f"{band['Band Name']}<br>{band['Start MHz']:.1f}-{band['Stop MHz']:.1f} MHz"],
                    textfont=dict(
                        size=8,
                        color=gov_marker_text_color
                    ),
                    showlegend=False,
                    hoverinfo='text',
                    name=f"Band Label: {band['Band Name']}"
                ))

    # Apply Dark Mode Theme
    fig.update_layout(template="plotly_dark")

    fig.update_layout(hovermode="x unified") # Show all traces on hover

    # This function now returns the figure and the output HTML path (which was passed in)
    return fig, output_html_path
