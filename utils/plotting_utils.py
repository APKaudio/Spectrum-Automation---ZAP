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
        messagebox.showerror("Plot Open Error", f"Could not open plot in browser: {e}")


def plot_single_scan_data(scan_data, plot_title_suffix, include_tv_markers=True, include_gov_markers=True, output_html_path=None, auto_open_browser=True):
    """
    Plots a single scan's frequency vs. amplitude data using Plotly.
    Also handles saving the plot to an HTML file and optionally opening it.

    Args:
        scan_data (list): A list of (frequency_hz, amplitude_dbm) tuples.
        plot_title_suffix (str): A suffix to add to the plot title (e.g., timestamp, scan name).
        include_tv_markers (bool): Whether to include TV band markers on the plot.
        include_gov_markers (bool): Whether to include Government band markers on the plot.
        output_html_path (str, optional): Full path to save the HTML plot. If None, plot is not saved.
        auto_open_browser (bool): If True and output_html_path is provided, opens the plot in browser.

    Returns:
        plotly.graph_objects.Figure: The Plotly figure object, or None if no data.
        str: The path to the saved HTML file, or None if not saved.
    """
    if not scan_data:
        print("No data to plot for single scan.")
        return None, None

    # Convert data to DataFrame
    df = pd.DataFrame(scan_data, columns=['Frequency_Hz', 'Power_dBm'])
    df['Frequency_MHz'] = df['Frequency_Hz'] / MHZ_TO_HZ

    # Determine plot title
    plot_title = f"Spectrum Scan - {plot_title_suffix}"

    fig = go.Figure()

    # Add the main scan trace
    fig.add_trace(go.Scatter(
        x=df['Frequency_MHz'],
        y=df['Power_dBm'],
        mode='lines',
        name='Scan Data',
        line=dict(color='cyan', width=1.5),
        showlegend=False # Hide this trace from the legend
    ))

    # Set Y-axis range: Max always 0 dBm, Min based on data with padding
    y_range_max = 0 # Fixed maximum at 0 dBm
    y_range_min = df['Power_dBm'].min() - 5 # Min based on data with 5 dB padding

    # Ensure a reasonable default range if data is flat or single point
    # Also handle cases where min_power_dbm is very high (e.g., positive)
    if y_range_max <= y_range_min:
        y_range_min = -100 # Fallback to a common low value if min is unexpectedly high or data is flat
        y_range_max = 0 # Keep max at 0

    # Get X-axis range from data
    x_range_min = df['Frequency_MHz'].min()
    x_range_max = df['Frequency_MHz'].max()

    # Define colors for markers
    tv_band_fill_color = 'rgba(255, 165, 0, 0.1)' # Orange, semi-transparent
    tv_marker_line_color = 'orange'
    tv_marker_text_color = 'orange'

    gov_band_fill_color = 'rgba(0, 128, 0, 0.1)' # Green, semi-transparent
    gov_marker_line_color = 'lightgreen'
    gov_marker_text_color = 'lightgreen'

    # Staggered Y-offset levels for text markers
    y_offset_levels = [0.05, 0.10, 0.15, 0.20] # 5%, 10%, 15%, 20% from top of plot area


    # Add TV Band Markers
    if include_tv_markers and TV_PLOT_BAND_MARKERS:
        for i, band in enumerate(TV_PLOT_BAND_MARKERS):
            if not isinstance(band, dict): # Defensive check
                print(f"Warning: Expected dictionary for TV band marker, but got {type(band)}: {band}. Skipping.")
                continue
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

            # Determine the Y position based on staggering
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

    # Add Government Band Markers
    if include_gov_markers and GOV_PLOT_BAND_MARKERS:
        for i, band in enumerate(GOV_PLOT_BAND_MARKERS):
            if not isinstance(band, dict): # Defensive check
                print(f"Warning: Expected dictionary for Government band marker, but got {type(band)}: {band}. Skipping.")
                continue
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
    fig.update_layout(
        template="plotly_dark",
        title={
            'text': plot_title,
            'y':1.0, # Set Y position to the top of the plot area
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'bottom', # Align the bottom of the title with y=1.0
            'font': dict(size=16, color='white')
        },
        xaxis_title="Frequency (MHz)",
        yaxis_title="Amplitude (dBm)",
        hovermode="x unified",
        margin=dict(l=50, r=50, t=100, b=50), # Increased top margin to make space for the title
        xaxis=dict(
            showgrid=True,
            gridcolor='rgba(255,255,255,0.1)',
            zeroline=False,
            rangeslider=dict(
                visible=True,
                thickness=0.05,
                # Reverted: Removed attempts to hide rangeslider yaxis labels, as it's not directly supported
            ),
            type="linear",
            range=[x_range_min, x_range_max] if x_range_min is not None and x_range_max is not None else None
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='rgba(255,255,255,0.1)',
            zeroline=False,
            # Explicitly set autorange to True for Y-axis interactivity
            autorange=True
        ),
        plot_bgcolor='black',
        paper_bgcolor='black',
        font=dict(color='white'),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )

    # If an output path is provided, save the figure
    if output_html_path:
        # Ensure the directory exists
        output_dir = os.path.dirname(output_html_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created directory for plot: {output_dir}")
        fig.write_html(output_html_path, auto_open=False) # auto_open=False as opening is handled by _open_plot_in_browser
        print(f"✅ Plot saved to: {output_html_path}")
        if auto_open_browser: # Open in browser if requested
            _open_plot_in_browser(output_html_path)

    return fig, output_html_path


def plot_multi_trace_data(
    aggregated_df,
    plot_title_only_datetime,
    include_tv_markers_var,
    include_gov_markers_var,
    historical_dfs_with_names=None,
    output_html_path=None
):
    """
    Plots aggregated (averaged, median, range) and optional historical scan data
    on a single Plotly graph.

    Args:
        aggregated_df (pd.DataFrame): DataFrame with 'Frequency_MHz', 'Average_Power_dBm',
                                      'Median_Power_dBm', 'Min_Power_dBm', 'Max_Power_dBm' columns.
        plot_title_only_datetime (str): Suffix for the plot title.
        include_tv_markers (bool): Whether to include TV band markers.
        include_gov_markers (bool): Whether to include Government band markers.
        historical_dfs_with_names (list of dict): List of dictionaries, each with 'name' (str)
                                                  and 'df' (pd.DataFrame) for historical overlays.
                                                  Each historical df must have 'Frequency_MHz' and 'Power_dBm'.
        output_html_path (str): Full path to save the HTML plot.

    Returns:
        tuple: (plotly.graph_objects.Figure, str): The Plotly figure object and the output HTML path.
               Returns (None, None) if no data to plot.
    """
    if aggregated_df.empty and not historical_dfs_with_names:
        print("No data to plot for multi-trace or historical average.")
        return None, None

    plot_title = f"Aggregated Spectrum Scan - {plot_title_only_datetime}"
    fig = go.Figure()

    # Add historical overlays first, with lighter color and lower opacity
    if historical_dfs_with_names:
        for i, hist_item in enumerate(historical_dfs_with_names):
            hist_df = hist_item['df']
            full_name = hist_item['name']
            
            # Extract only the date and time part from the filename
            # Example: MyScan_RBW100K_HOLD0_Offset0_20250723_104243
            match = re.search(r'(\d{8}_\d{6})$', full_name)
            display_name = match.group(1) if match else full_name # Fallback to full name if regex fails
            
            if not hist_df.empty:
                # Use a very light grey or transparent color for historical overlays
                line_color = 'rgba(244, 144, 44, .8)' 
                
                fig.add_trace(go.Scatter(
                    x=hist_df['Frequency_MHz'],
                    y=hist_df['Power_dBm'],
                    mode='lines',
                    name=f'{display_name}', # Removed "Historical:" prefix
                    line=dict(color=line_color, width=0.5, dash='solid'), # Changed to dashed line, width 0.5
                    hoverinfo='x+y+name', # Show frequency, power, and name on hover
                    showlegend=True # Show in legend
                ))

    # Add aggregated traces (Average, Median, Range)
    if not aggregated_df.empty:
        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency_MHz'],
            y=aggregated_df['Average_Power_dBm'],
            mode='lines',
            name='Average Power',
            line=dict(color='cyan', width=2, dash='solid') # Changed to solid line
        ))

        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency_MHz'],
            y=aggregated_df['Median_Power_dBm'],
            mode='lines',
            name='Median Power',
            line=dict(color='yellow', width=1.5, dash='solid') # Changed to solid line
        ))

        # Add Range as a filled area (min to max)
        if 'Min_Power_dBm' in aggregated_df.columns and 'Max_Power_dBm' in aggregated_df.columns:
            fig.add_trace(go.Scatter(
                x=aggregated_df['Frequency_MHz'],
                y=aggregated_df['Max_Power_dBm'],
                mode='lines',
                name='Max Power',
                line=dict(color='red', width=0.8, dash='solid'), # Changed to solid line
                fill=None # No fill for the upper boundary initially
            ))
            fig.add_trace(go.Scatter(
                x=aggregated_df['Frequency_MHz'],
                y=aggregated_df['Min_Power_dBm'],
                mode='lines',
                name='Min Power',
                line=dict(color='green', width=0.8, dash='solid'), # Changed to solid line
                fill='tonexty', # Fills the area between this trace and the 'Max Power' trace
                fillcolor='rgba(255,0,0,0.02)' # Even lighter red fill for the range
            ))
            # NEW: Add a trace for the Range_dBm itself
            if 'Range_dBm' in aggregated_df.columns:
                fig.add_trace(go.Scatter(
                    x=aggregated_df['Frequency_MHz'],
                    y=aggregated_df['Range_dBm'], # Plotting the calculated range directly
                    mode='lines',
                    name='Range (Max-Min)',
                    line=dict(color='magenta', width=1, dash='solid'), # Changed to solid magenta line
                    hoverinfo='x+y+name',
                    showlegend=True
                ))
        elif 'Range_dBm' in aggregated_df.columns:
            # Fallback if only Range_dBm is available, assume it's centered around Average
            # This is a less accurate representation of true min/max.
            fig.add_trace(go.Scatter(
                x=aggregated_df['Frequency_MHz'],
                y=aggregated_df['Average_Power_dBm'] + (aggregated_df['Range_dBm'] / 2),
                mode='lines',
                name='Upper Range',
                line=dict(color='red', width=0.8, dash='solid'),
                fill=None
            ))
            fig.add_trace(go.Scatter(
                x=aggregated_df['Frequency_MHz'],
                y=aggregated_df['Average_Power_dBm'] - (aggregated_df['Range_dBm'] / 2),
                mode='lines',
                name='Lower Range',
                line=dict(color='green', width=0.8, dash='solid'),
                fill='tonexty',
                fillcolor='rgba(255,0,0,0.02)'
            ))
            # NEW: Add a trace for the Range_dBm itself (for this fallback case)
            fig.add_trace(go.Scatter(
                x=aggregated_df['Frequency_MHz'],
                y=aggregated_df['Range_dBm'], # Plotting the calculated range directly
                mode='lines',
                name='Range (Max-Min)',
                line=dict(color='magenta', width=1, dash='solid'), # Changed to solid magenta line
                hoverinfo='x+y+name',
                showlegend=True
            ))
        else:
            print("Warning: No 'Min_Power_dBm', 'Max_Power_dBm', or 'Range_dBm' found for plotting range.")


    # Determine initial Y-axis range based on all data (including historical)
    all_power_data_for_range_calc = []
    if not aggregated_df.empty:
        all_power_data_for_range_calc.extend(aggregated_df['Average_Power_dBm'].tolist())
        if 'Min_Power_dBm' in aggregated_df.columns and 'Max_Power_dBm' in aggregated_df.columns:
            all_power_data_for_range_calc.extend(aggregated_df['Min_Power_dBm'].tolist())
            all_power_data_for_range_calc.extend(aggregated_df['Max_Power_dBm'].tolist())
        
        # Always include the Range_dBm values if available, as they are now plotted
        if 'Range_dBm' in aggregated_df.columns:
            all_power_data_for_range_calc.extend(aggregated_df['Range_dBm'].tolist())
    
    if historical_dfs_with_names:
        for hist_item in historical_dfs_with_names:
            if not hist_item['df'].empty:
                all_power_data_for_range_calc.extend(hist_item['df']['Power_dBm'].tolist())

    y_range_min = min(all_power_data_for_range_calc) - 5 if all_power_data_for_range_calc else -100
    y_range_max = max(all_power_data_for_range_calc) + 5 if all_power_data_for_range_calc else 0

    if y_range_max <= y_range_min: # Handle cases with very flat or single-point data
        y_range_min = -100
        y_range_max = 0


    # Get X-axis range from all data (aggregated and historical)
    all_freq_data = []
    if not aggregated_df.empty:
        all_freq_data.extend(aggregated_df['Frequency_MHz'].tolist())
    if historical_dfs_with_names:
        for hist_item in historical_dfs_with_names:
            if not hist_item['df'].empty:
                all_freq_data.extend(hist_item['df']['Frequency_MHz'].tolist())

    x_range_min = min(all_freq_data) if all_freq_data else None
    x_range_max = max(all_freq_data) if all_freq_data else None


    # Define colors for markers
    tv_band_fill_color = 'rgba(255, 165, 0, 0.1)' # Orange, semi-transparent
    tv_marker_line_color = 'orange'
    tv_marker_text_color = 'orange'

    gov_band_fill_color = 'rgba(0, 128, 0, 0.1)' # Green, semi-transparent
    gov_marker_line_color = 'lightgreen'
    gov_marker_text_color = 'lightgreen'

    # Staggered Y-offset levels for text markers
    y_offset_levels = [0.05, 0.10, 0.15, 0.20] # 5%, 10%, 15%, 20% from top of plot area

    # Add TV Band Markers
    if include_tv_markers_var and TV_PLOT_BAND_MARKERS:
        for i, band in enumerate(TV_PLOT_BAND_MARKERS):
            if not isinstance(band, dict): # Defensive check
                print(f"Warning: Expected dictionary for TV band marker, but got {type(band)}: {band}. Skipping.")
                continue
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
                    color=tv_marker_text_color
                ),
                showlegend=False,
                hoverinfo='text',
                name=f"Band Label: {band['Band Name']}"
            ))

    # Add Government Band Markers
    if include_gov_markers_var and GOV_PLOT_BAND_MARKERS:
        for i, band in enumerate(GOV_PLOT_BAND_MARKERS):
            if not isinstance(band, dict): # Defensive check
                print(f"Warning: Expected dictionary for Government band marker, but got {type(band)}: {band}. Skipping.")
                continue
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
    fig.update_layout(
        template="plotly_dark",
        title={
            'text': plot_title,
            'y':1.0, # Set Y position to the top of the plot area
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'bottom', # Align the bottom of the title with y=1.0
            'font': dict(size=16, color='white')
        },
        xaxis_title="Frequency (MHz)",
        yaxis_title="Amplitude (dBm)",
        hovermode="x unified",
        margin=dict(l=50, r=50, t=100, b=50), # Increased top margin to make space for the title
        xaxis=dict(
            showgrid=True,
            gridcolor='rgba(255,255,255,0.1)',
            zeroline=False,
            rangeslider=dict(
                visible=False, # Removed range slider from average plot
                thickness=0.05,
            ),
            type="linear",
            range=[x_range_min, x_range_max] if x_range_min is not None and x_range_max is not None else None
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='rgba(255,255,255,0.1)',
            zeroline=False,
            # Explicitly set autorange to True for Y-axis interactivity
            autorange=True,
            # Set initial range based on calculated min/max from aggregated data (with padding)
            range=[y_range_min, y_range_max]
        ),
        plot_bgcolor='black',
        paper_bgcolor='black',
        font=dict(color='white'),
        legend=dict(
            orientation="v", # Changed to vertical orientation
            yanchor="top",   # Anchor to the top
            y=1,             # Position at the top
            xanchor="right",  # Anchor to the left
            x=0              # Position at the left
        )
    )

    # If an output path is provided, save the figure
    if output_html_path:
        # Ensure the directory exists
        output_dir = os.path.dirname(output_html_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created directory for plot: {output_dir}")
        fig.write_html(output_html_path, auto_open=False) # auto_open=False as opening is handled by _open_plot_in_browser
        print(f"✅ Plot saved to: {output_html_path}")

    return fig, output_html_path
