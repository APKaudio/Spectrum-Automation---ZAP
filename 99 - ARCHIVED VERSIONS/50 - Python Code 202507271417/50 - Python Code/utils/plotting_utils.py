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
    """
    Helper function to open an HTML plot file in the default web browser.

    Inputs:
        plot_path (str): The full file path to the HTML plot to open.
    Process:
        1. Attempts to open the `plot_path` using `webbrowser.open()`.
        2. Prints a success message to the console.
        3. Catches any `Exception` during the opening process and displays
           a Tkinter messagebox error, as this function might be called from a thread.
    Outputs:
        None. (Side effect: Opens a web browser window.)
    """
    try:
        webbrowser.open(plot_path)
        print(f"✅ Opened plot in browser: {plot_path}")
    except Exception as e:
        print(f"❌ Failed to open plot in browser: {e}")
        # Use a simple Tkinter messagebox for errors, as this function might be called from a thread
        # and direct console print might not be noticed by the user.
        messagebox.showerror("Plot Open Error", f"Could not open plot in browser: {e}")


def plot_single_scan_data(csv_file_path, include_tv_markers=True, include_gov_markers=True, output_html_path=None, auto_open_browser=True, single_marker_data=None):
    """
    Plots a single scan's frequency vs. amplitude data using Plotly.
    It now reads data directly from a CSV file without headers.
    Also handles saving the plot to an HTML file and optionally opening it.

    Inputs:
        csv_file_path (str): The full path to the CSV file containing the scan data.
        include_tv_markers (bool): If True, adds shaded regions and text labels
                                   for North American TV broadcast bands.
        include_gov_markers (bool): If True, adds shaded regions and text labels
                                    for common Government/Commercial frequency bands.
        output_html_path (str, optional): The full file path where the generated
                                          HTML plot should be saved. If None, the plot
                                          is not saved to a file.
        auto_open_browser (bool): If True and `output_html_path` is provided,
                                  the generated HTML plot will be automatically
                                  opened in the default web browser.
        single_marker_data (tuple, optional): A tuple (frequency_hz, name) for a single marker to highlight.
                                              Defaults to None.
    Process:
        1. **Data Loading**: Reads the input `csv_file_path` into a pandas DataFrame,
           specifying `header=None` and then manually assigning 'Frequency_MHz' and 'Power_dBm' columns.
           **Crucially, it now converts 'Frequency_MHz' and 'Power_dBm' to numeric type.**
        2. **Plotly Figure Initialization**: Creates an empty `go.Figure` object.
        3. **Main Trace Addition**: Adds a `go.Scatter` trace for the main scan data
           (Frequency_MHz vs. Power_dBm) with specific line styling.
        4. **Y-Axis Range Determination**: Sets the Y-axis (amplitude) range, fixing the maximum at 0 dBm
           and dynamically setting the minimum based on the data's lowest power value, with some padding.
        5. **X-Axis Range Determination**: Determines the X-axis (frequency) range from the scan data.
        6. **Band Marker Addition (TV & Government)**:
           - If `include_tv_markers` or `include_gov_markers` is True, iterates through the respective
             `TV_PLOT_BAND_MARKERS` or `GOV_PLOT_BAND_MARKERS` lists (from `frequency_bands.py`).
           - For each band, it adds a rectangular `go.layout.Shape` to create a shaded background,
             and a `go.Scatter` trace with `mode='text'` to add text labels for the band name and frequency range.
           - **Staggering Labels**: Uses `y_offset_levels` and modulo arithmetic to stagger the Y-position
             of the text labels, preventing overlap when multiple bands are close together.
        7. **Single Marker Addition (Optional)**: If `single_marker_data` is provided, adds a marker.
        8. **Layout Configuration**: Applies a "plotly_dark" theme and configures plot title,
           axis labels, grid lines, rangeslider visibility, and legend orientation for better aesthetics.
        9. **HTML Export (Optional)**: If `output_html_path` is provided, it ensures the output directory exists
           and then saves the Plotly figure to an HTML file using `fig.write_html()`.
        10. **Browser Opening (Optional)**: If `auto_open_browser` is True, calls `_open_plot_in_browser`
            to open the saved HTML file.

    Outputs:
        tuple: `(plotly.graph_objects.Figure, str)` - The generated Plotly figure object and the
               full path to the saved HTML file (or None if not saved). Returns `(None, None)`
               if data loading fails.
    """
    try:
        # Read CSV without header and assign column names
        df = pd.read_csv(csv_file_path, header=None)
        df.columns = ['Frequency_MHz', 'Power_dBm']
        
        # Explicitly convert 'Frequency_MHz' and 'Power_dBm' to numeric, coercing errors to NaN
        df['Frequency_MHz'] = pd.to_numeric(df['Frequency_MHz'], errors='coerce')
        df['Power_dBm'] = pd.to_numeric(df['Power_dBm'], errors='coerce')
        
        # Drop any rows where conversion failed
        df.dropna(subset=['Frequency_MHz', 'Power_dBm'], inplace=True)

    except FileNotFoundError:
        print(f"🚫 Plotting Error: CSV file not found at {csv_file_path}")
        messagebox.showerror("Plotting Error", f"Scan data CSV file not found: {csv_file_path}")
        return None, None
    except Exception as e:
        print(f"🚫 Plotting Error: Could not read CSV file {csv_file_path}: {e}")
        messagebox.showerror("Plotting Error", f"Could not read CSV file {csv_file_path}: {e}")
        return None, None

    if df.empty:
        print("No data to plot for single scan after loading from CSV.")
        return None, None

    # Determine plot title from the CSV filename
    plot_title = os.path.basename(csv_file_path).replace('.csv', '') # Remove .csv extension for title

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

    # Set Y-axis range: Min based on data with 5 dB padding, Max based on data with 5 dB padding
    if not df.empty:
        y_range_min = df['Power_dBm'].min() - 5
        y_range_max = df['Power_dBm'].max() + 5
    else:
        y_range_min = -100 # Fallback default
        y_range_max = 0    # Fallback default

    # Ensure a reasonable default range if data is flat or single point
    if y_range_max <= y_range_min:
        y_range_min = -100 # Fallback to a common low value if min is unexpectedly high or data is flat
        y_range_max = 0 # Keep max at 0 (or a sensible default if all data is negative)

    # Get X-axis range from data
    if not df.empty:
        x_range_min = df['Frequency_MHz'].min()
        x_range_max = df['Frequency_MHz'].max()
    else:
        x_range_min = None
        x_range_max = None

    # Define colors for markers
    tv_band_fill_color = 'rgba(255, 255, 0, 0.1)' # Yellow, semi-transparent
    tv_marker_line_color = 'yellow'
    tv_marker_text_color = 'yellow'

    gov_band_fill_color = 'rgba(255, 0, 0, 0.1)' # Red, semi-transparent
    gov_marker_line_color = 'red'
    gov_marker_text_color = 'red'

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
    
    # Add single marker if provided
    if single_marker_data:
        freq_hz, name = single_marker_data
        freq_mhz = freq_hz / MHZ_TO_HZ
        
        # Find the power level at or near the marker frequency
        # This is a simple nearest-neighbor lookup; for more accuracy, interpolation could be used.
        closest_point = df.iloc[(df['Frequency_MHz'] - freq_mhz).abs().argsort()[:1]]
        if not closest_point.empty:
            power_dbm = closest_point['Power_dBm'].iloc[0]
        else:
            power_dbm = -100 # Default if no data point is found (shouldn't happen with valid CSV)

        fig.add_trace(go.Scatter(
            x=[freq_mhz],
            y=[power_dbm],
            mode='markers+text',
            marker=dict(color='white', size=12, symbol='star', line=dict(color='black', width=2)),
            text=[f"{name}<br>{freq_mhz:.3f} MHz"],
            textposition="top center",
            textfont=dict(color='white', size=12),
            name=f'Selected: {name}',
            hovertemplate=f'<b>{{text}}</b><br>Frequency: %{{x:.3f}} MHz<br>Power: %{{y:.2f}} dBm<extra></extra>'
        ))
        print(f"Added single marker for '{name}' at {freq_mhz:.3f} MHz to plot.")


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
            autorange=False, # Explicitly set autorange to False
            range=[x_range_min, x_range_max] if x_range_min is not None and x_range_max is not None else None
        ),
        yaxis=dict(
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
    plot_title_full, # Changed to plot_title_full to use the complete title
    include_tv_markers_var,
    include_gov_markers_var,
    historical_dfs_with_names=None,
    output_html_path=None
):
    """
    Plots aggregated (averaged, median, range, std dev, variance, PSD) and optional historical scan data
    on a single Plotly graph. This function is designed to visualize trends and variations
    across multiple scan cycles.

    Inputs:
        aggregated_df (pd.DataFrame): A pandas DataFrame containing the aggregated statistical
                                      data. Expected columns: 'Frequency_MHz', 'Average_dBm',
                                      'Median_dBm', 'Range_dBm', 'Std_Dev_dBm', 'Variance_dBm',
                                      and 'Average_PSD_dBm_Hz'.
        plot_title_full (str): The complete title string for the plot.
        include_tv_markers_var (bool): A boolean indicating whether to
                                       include TV band markers on the plot.
        include_gov_markers_var (bool): A boolean indicating whether to
                                        include Government band markers on the plot.
        historical_dfs_with_names (list of dict, optional): A list of dictionaries, where each
                                                            dictionary represents a historical scan
                                                            to be overlaid. Each dict should have
                                                            'name' (str for legend) and 'df' (pd.DataFrame
                                                            with 'Frequency_MHz' and 'Power_dBm').
        output_html_path (str, optional): The full file path where the generated HTML plot
                                          should be saved. If None, the plot is not saved to a file.

    Process:
        1. **Figure Initialization**: Creates an empty `go.Figure` object.
        2. **Historical Overlays (Optional)**:
           - If `historical_dfs_with_names` is provided, iterates through each historical scan.
           - Extracts the display name (often a timestamp) from the historical scan's name.
           - Reads the historical CSV into a DataFrame without headers and assigns columns.
           - Adds each historical scan as a `go.Scatter` trace with a light, semi-transparent color
             and a dotted line style, making them visually distinct as background layers.
           - These overlays are shown in the legend.
        3. **Aggregated Traces**:
           - If `aggregated_df` is not empty, adds `go.Scatter` traces for:
             - 'Average Power (dBm)' (red, solid line)
             - 'Median Power (dBm)' (yellow, solid line)
             - 'Range (Max-Min) (dB)' (magenta, solid line)
             - 'Standard Deviation (dB)' (lime, solid line)
        4. **Axis Range Determination**: Calculates the overall min/max for both Y-axis (amplitude)
           and X-axis (frequency) by considering all data (aggregated and historical overlays)
           to ensure the plot encompasses all relevant information.
        5. **Band Marker Addition (TV & Government)**:
           - Similar to `plot_single_scan_data`, adds shaded rectangular shapes and text labels
             for TV and Government frequency bands based on `include_tv_markers_var` and `include_gov_markers_var`.
           - Uses staggering for text labels to prevent overlap.
        6. **Layout Configuration**: Applies a "plotly_dark" theme, sets the plot title,
           axis labels, grid lines, rangeslider, and positions the legend vertically
           at the top-right for better readability with multiple traces.
        7. **HTML Export (Optional)**: If `output_html_path` is provided, ensures the output directory exists
           and saves the Plotly figure to an HTML file.

    Outputs:
        tuple: `(plotly.graph_objects.Figure, str)` - The generated Plotly figure object and the
               full path to the saved HTML file (or None if not saved). Returns `(None, None)`
               if no data (neither aggregated nor historical) is provided.
    """
    if aggregated_df.empty and not historical_dfs_with_names:
        print("No data to plot for multi-trace or historical average.")
        return None, None

    fig = go.Figure()

    # Add historical overlays first so aggregated traces appear on top
    if historical_dfs_with_names:
        for i, hist_item in enumerate(historical_dfs_with_names):
            hist_df = hist_item['df']
            full_name = hist_item['name']
            
            # Extract only the date and time part from the filename for display
            match_paren = re.search(r'\((\d{8}_\d{6})\)', full_name)
            match_end = re.search(r'(\d{8}_\d{6})$', full_name)

            if match_paren:
                display_name = match_paren.group(1)
            elif match_end:
                display_name = match_end.group(1)
            else:
                display_name = full_name # Fallback to full name if no match

            if not hist_df.empty:
                # Explicitly convert 'Frequency_MHz' and 'Power_dBm' to numeric
                hist_df['Frequency_MHz'] = pd.to_numeric(hist_df['Frequency_MHz'], errors='coerce')
                hist_df['Power_dBm'] = pd.to_numeric(hist_df['Power_dBm'], errors='coerce')
                hist_df.dropna(subset=['Frequency_MHz', 'Power_dBm'], inplace=True)

                # Use a more opaque orange for historical overlays
                line_color = 'rgba(244, 144, 44, 0.4)' # Increased opacity from 0.2 to 0.7
                
                fig.add_trace(go.Scatter(
                    x=hist_df['Frequency_MHz'],
                    y=hist_df['Power_dBm'],
                    mode='lines',
                    name=display_name, # Now only shows date and time
                    line=dict(color=line_color, width=1, dash='dash'), # Changed to dotted line for overlays
                    hoverinfo='x+y+name', # Show frequency, power, and name on hover
                    showlegend=True # Show in legend
                ))

    # Add aggregated traces (Average, Median, Range, Std Dev) after historical traces
    if not aggregated_df.empty:
        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency_MHz'],
            y=aggregated_df['Average_dBm'],
            mode='lines',
            name='Average Power (dBm)',
            line=dict(color='red', width=2, dash='solid')
        ))

        fig.add_trace(go.Scatter(
            x=aggregated_df['Frequency_MHz'],
            y=aggregated_df['Median_dBm'],
            mode='lines',
            name='Median Power (dBm)',
            line=dict(color='yellow', width=1.5, dash='solid')
        ))

        # Add Range (Max-Min)
        if 'Range_dBm' in aggregated_df.columns:
            fig.add_trace(go.Scatter(
                x=aggregated_df['Frequency_MHz'],
                y=aggregated_df['Range_dBm'],
                mode='lines',
                name='Range (Max-Min) (dB)',
                line=dict(color='magenta', width=1.5, dash='solid')
            ))
        
        # Add Standard Deviation
        if 'Std_Dev_dBm' in aggregated_df.columns:
            fig.add_trace(go.Scatter(
                x=aggregated_df['Frequency_MHz'],
                y=aggregated_df['Std_Dev_dBm'],
                mode='lines',
                name='Standard Deviation (dB)',
                line=dict(color='lime', width=1.5, dash='solid') # Greenish color
            ))


    # Determine initial Y-axis range based on all data (including historical and new metrics)
    all_y_data_for_range_calc = []
    if not aggregated_df.empty:
        all_y_data_for_range_calc.extend(aggregated_df['Average_dBm'].tolist())
        all_y_data_for_range_calc.extend(aggregated_df['Median_dBm'].tolist())
        if 'Range_dBm' in aggregated_df.columns:
            all_y_data_for_range_calc.extend(aggregated_df['Range_dBm'].tolist())
        if 'Std_Dev_dBm' in aggregated_df.columns:
            all_y_data_for_range_calc.extend(aggregated_df['Std_Dev_dBm'].tolist())
    
    if historical_dfs_with_names:
        for hist_item in historical_dfs_with_names:
            if not hist_item['df'].empty:
                all_y_data_for_range_calc.extend(hist_item['df']['Power_dBm'].tolist())

    y_range_min = min(all_y_data_for_range_calc) - 5 if all_y_data_for_range_calc else -100
    y_range_max = max(all_y_data_for_range_calc) + 5 if all_y_data_for_range_calc else 0

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
    tv_band_fill_color = 'rgba(255, 255, 0, 0.1)' # Yellow, semi-transparent
    tv_marker_line_color = 'yellow'
    tv_marker_text_color = 'yellow'

    gov_band_fill_color = 'rgba(255, 0, 0, 0.1)' # Red, semi-transparent
    gov_marker_line_color = 'red'
    gov_marker_text_color = 'red'

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
            'text': plot_title_full, # Use the full title passed as argument
            'y':0.95, # Adjusted Y position slightly lower
            'x':0.5,
            'xanchor': 'center',
            'yanchor': 'top', # Anchor the top of the title to this y position
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
            autorange=False, # Explicitly set autorange to False
            range=[x_range_min, x_range_max] if x_range_min is not None and x_range_max is not None else None
        ),
        yaxis=dict(
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
            print(f"Created directory for plot: {output_dir}")
        fig.write_html(output_html_path, auto_open=False) # auto_open=False as opening is handled by _open_plot_in_browser
        print(f"✅ Plot saved to: {output_html_path}")

    return fig, output_html_path
