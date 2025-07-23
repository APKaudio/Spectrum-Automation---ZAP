# plotting_utils.py

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import webbrowser
import os
import tkinter as tk # For messagebox
from tkinter import messagebox
import re # Added import for regular expressions

# Import constants from frequency_bands.py
try:
    from .frequency_bands import ( # Changed to relative import
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
            messagebox.showerror("Plot Open Error", f"Could not open plot in browser: {e}")
            root.destroy()
        else:
            messagebox.showerror("Plot Open Error", f"Could not open plot in browser: {e}")


def plot_single_scan_data(scan_data, plot_title_suffix, include_tv_markers=True, include_gov_markers=True):
    """
    Plots a single scan's frequency vs. amplitude data using Plotly.

    Args:
        scan_data (list of tuples): List of (frequency_hz, amplitude_dbm) data points.
        plot_title_suffix (str): Suffix for the plot title (e.g., timestamp).
        include_tv_markers (bool): Whether to include TV band markers.
        include_gov_markers (bool): Whether to include Government band markers.

    Returns:
        plotly.graph_objects.Figure: The Plotly figure object, or None if no data.
    """
    if not scan_data:
        print("No scan data provided for plotting.")
        return None

    # Convert to DataFrame for easier plotting
    df = pd.DataFrame(scan_data, columns=['Frequency_Hz', 'Power_dBm'])
    df['Frequency_MHz'] = df['Frequency_Hz'] / MHZ_TO_HZ

    # Determine Y-axis range based on data
    y_range_min = df['Power_dBm'].min() if not df.empty else -120 # Lowest value from data, default to -120
    y_range_max = 0 # Always 0 dBm as requested

    # Create the plot
    fig = px.line(df, x='Frequency_MHz', y='Power_dBm',
                  title=f'Spectrum Scan - {plot_title_suffix}',
                  labels={'Frequency_MHz': 'Frequency (MHz)', 'Power_dBm': 'Power (dBm)'})

    # Update layout for better readability
    fig.update_layout(
        xaxis_title='Frequency (MHz)',
        yaxis_title='Power (dBm)',
        hovermode='x unified',
        title_x=0.5, # Center the title
        margin=dict(l=40, r=40, t=60, b=40) # Adjust margins
    )

    # Set Y-axis range
    fig.update_yaxes(range=[y_range_min, y_range_max])

    # Add band markers as shaded regions and text
    y_offset_levels = [0.05, 0.1, 0.15, 0.2] # Levels for staggering text markers

    if include_tv_markers and TV_PLOT_BAND_MARKERS:
        tv_band_fill_color = 'rgba(0, 255, 0, 0.1)' # Light green transparent
        tv_marker_line_color = 'rgba(0, 255, 0, 0.5)'
        tv_marker_text_color = tv_marker_line_color # Text color matches line color as requested

        for i, band in enumerate(TV_PLOT_BAND_MARKERS):
            fig.add_shape(
                type="rect",
                xref="x", yref="paper",
                x0=band["Start MHz"], y0=0,
                x1=band["Stop MHz"], y1=1,
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
            current_y_offset = y_offset_levels[i % len(y_offset_levels)]
            y_text_position = y_range_max - (y_range_max - y_range_min) * current_y_offset

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

    if include_gov_markers and GOV_PLOT_BAND_MARKERS:
        gov_band_fill_color = 'rgba(255, 0, 0, 0.1)' # Light red transparent
        gov_marker_line_color = 'rgba(255, 0, 0, 0.5)'
        gov_marker_text_color = gov_marker_line_color # Text color matches line color as requested

        for i, band in enumerate(GOV_PLOT_BAND_MARKERS):
            fig.add_shape(
                type="rect",
                xref="x", yref="paper",
                x0=band["Start MHz"], y0=0,
                x1=band["Stop MHz"], y1=1,
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

    return fig


def plot_multi_trace_data(
    aggregated_df,
    plot_title_suffix,
    include_tv_markers=True,
    include_gov_markers=True,
    historical_dfs_with_names=None,
    output_html_path=None # Added output_html_path parameter
):
    """
    Plots aggregated (average, median, range) and optionally historical scan data using Plotly.

    Args:
        aggregated_df (pd.DataFrame): DataFrame with 'Frequency_MHz', 'Average_dBm', 'Median_dBm', 'Range_dBm'.
        plot_title_suffix (str): Suffix for the plot title (e.g., datetime for historical plot).
        include_tv_markers (bool): Whether to include TV band markers.
        include_gov_markers (bool): Whether to include Government band markers.
        historical_dfs_with_names (list of tuples): Optional list of (DataFrame, name_str) for overlaying historical scans.
        output_html_path (str): The full path including filename where the HTML plot should be saved.

    Returns:
        plotly.graph_objects.Figure: The Plotly figure object, or None if no data.
        str: The path where the HTML plot was saved (same as output_html_path).
    """
    if aggregated_df.empty and not historical_dfs_with_names:
        print("No data provided for multi-trace plotting.")
        return None, None

    fig = go.Figure()

    # Determine Y-axis range based on data
    all_power_data = []
    if not aggregated_df.empty:
        all_power_data.extend(aggregated_df['Average_dBm'].tolist())
        all_power_data.extend(aggregated_df['Median_dBm'].tolist())
        # Include Range_dBm for Y-max calculation if it's positive
        all_power_data.extend(aggregated_df['Range_dBm'].tolist())

    if historical_dfs_with_names:
        for hist_df, _ in historical_dfs_with_names:
            if not hist_df.empty:
                all_power_data.extend(hist_df['Power_dBm'].tolist())


    y_range_min = min(all_power_data) if all_power_data else -120 # Lowest value from all data, default to -120
    # y_range_max should be the maximum of all_power_data, but at least 0
    y_range_max = max(0, max(all_power_data) if all_power_data else 0)


    # Add aggregated traces (Average, Median, Range) - now at the top
    if not aggregated_df.empty:
        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency_MHz'],
            y=aggregated_df['Average_dBm'],
            mode='lines',
            name='Average Power (dBm)', # Changed name
            line=dict(color='red', width=3), # Changed to RED, width 3
            hoverinfo='name+x+y'
        ))

        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency_MHz'],
            y=aggregated_df['Median_dBm'],
            mode='lines',
            name='Median Power (dBm)', # Changed name
            line=dict(color='orange', width=2), # Changed to ORANGE
            hoverinfo='name+x+y'
        ))

        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency_MHz'],
            y=aggregated_df['Range_dBm'],
            mode='lines',
            name='Range (Max - Min) (dB)', # Changed name
            line=dict(color='yellow', width=2), # Changed to YELLOW
            hoverinfo='name+x+y'
        ))

    # Add historical overlays after aggregated traces
    if historical_dfs_with_names:
        for hist_df, name in historical_dfs_with_names:
            if not hist_df.empty:
                # Extract only the date and time for the name
                # Assuming name format is "MyScan_RBW100K_HOLD1_Offset0 (YYYY-MM-DD HH:MM)" or similar
                # We need to parse the original name to get just the datetime part
                # The name passed from averaging_utils.py is already formatted as "Prefix_RBW_HOLD_Offset (YYYY-MM-DD HH:MM:SS)"
                # So we just need to extract the part in parentheses
                match_datetime = re.search(r'\((.*?)\)', name)
                if match_datetime:
                    display_name = match_datetime.group(1) # Get the content inside parentheses
                else:
                    # Fallback if no datetime in parentheses, use the full name as is
                    display_name = name

                fig.add_trace(go.Scatter(
                    x=hist_df['Frequency_MHz'],
                    y=hist_df['Power_dBm'],
                    mode='lines',
                    name=display_name, # Only date and time as the name
                    line=dict(color='rgba(100, 100, 255, 0.3)', dash='dot'), # Light blue, semi-transparent, and dotted
                    hoverinfo='name+x+y'
                ))


    # Update layout for better readability
    fig.update_layout(
        title=f'Aggregated Spectrum Scan Data - {plot_title_suffix}',
        xaxis_title='Frequency (MHz)',
        yaxis_title='Power (dBm)',
        hovermode='x unified',
        title_x=0.5, # Center the title
        margin=dict(l=40, r=40, t=60, b=40) # Adjust margins
    )

    # Set Y-axis range
    fig.update_yaxes(range=[y_range_min, y_range_max])


    # Add band markers as shaded regions and text
    y_offset_levels = [0.05, 0.1, 0.15, 0.2] # Levels for staggering text markers

    if include_tv_markers and TV_PLOT_BAND_MARKERS:
        tv_band_fill_color = 'rgba(0, 255, 0, 0.1)' # Light green transparent
        tv_marker_line_color = 'rgba(0, 255, 0, 0.5)'
        tv_marker_text_color = tv_marker_line_color # Text color matches line color as requested

        for i, band in enumerate(TV_PLOT_BAND_MARKERS):
            fig.add_shape(
                type="rect",
                xref="x", yref="paper",
                x0=band["Start MHz"], y0=0,
                x1=band["Stop MHz"], y1=1,
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
            current_y_offset = y_offset_levels[i % len(y_offset_levels)]
            y_text_position = y_range_max - (y_range_max - y_range_min) * current_y_offset

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

    if include_gov_markers and GOV_PLOT_BAND_MARKERS:
        gov_band_fill_color = 'rgba(255, 0, 0, 0.1)' # Light red transparent
        gov_marker_line_color = 'rgba(255, 0, 0, 0.5)'
        gov_marker_text_color = gov_marker_line_color # Text color matches line color as requested

        for i, band in enumerate(GOV_PLOT_BAND_MARKERS):
            fig.add_shape(
                type="rect",
                xref="x", yref="paper",
                x0=band["Start MHz"], y0=0,
                x1=band["Stop MHz"], y1=1,
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

    # This function now returns the figure and the output path
    return fig, output_html_path

