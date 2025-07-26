# src/plot_logic.py
import pandas as pd
import plotly.graph_objects as go
import plotly.offline as pyo
import os
import sys
from tkinter import messagebox
from datetime import datetime # Ensure datetime is imported correctly

# Import constants from frequency_bands.py
from utils.frequency_bands import (
    MHZ_TO_HZ,
    TV_PLOT_BAND_MARKERS,
    GOV_PLOT_BAND_MARKERS
)

def generate_single_scan_plot_and_open_wrapper_logic(app_instance, csv_file_path, plot_title_suffix, output_html_path, auto_open_browser=True, single_marker_data=None):
    """
    Generates a Plotly HTML plot for a single scan and optionally opens it in a browser.
    This function is a wrapper to be called from the main thread using `app_instance.after()`.

    Inputs:
        app_instance (App): The main application instance, providing access to settings.
        csv_file_path (str): The full path to the CSV file containing the scan data.
        plot_title_suffix (str): A string to append to the plot's main title.
        output_html_path (str): The full path where the generated HTML plot should be saved.
        auto_open_browser (bool, optional): If True, the generated HTML plot will be
                                            automatically opened in the default web browser. Defaults to True.
        single_marker_data (tuple, optional): A tuple (frequency_hz, name) for a single marker to highlight.
                                              Defaults to None.
    Process:
        1. Calls `_generate_single_scan_plot_and_open` with all necessary parameters.
    Outputs: None (generates HTML file, may open browser)
    """
    _generate_single_scan_plot_and_open(
        app_instance.output_folder_var.get(),
        csv_file_path,
        plot_title_suffix,
        output_html_path,
        app_instance.include_tv_markers_var.get(),
        app_instance.include_gov_markers_var.get(),
        auto_open_browser,
        single_marker_data # Pass the single marker data
    )

def _generate_single_scan_plot_and_open(output_dir, csv_file, plot_title_suffix, output_html_path, include_tv_markers, include_gov_markers, auto_open_browser, single_marker_data=None):
    """
    Internal function to generate a Plotly HTML plot for a single scan.

    Inputs:
        output_dir (str): The directory where the plot will be saved.
        csv_file (str): The path to the CSV file containing the scan data.
        plot_title_suffix (str): A string to append to the plot's main title.
        output_html_path (str): The full path where the generated HTML plot should be saved.
        include_tv_markers (bool): Whether to include TV band markers.
        include_gov_markers (bool): Whether to include government band markers.
        auto_open_browser (bool): If True, the generated HTML plot will be automatically opened.
        single_marker_data (tuple, optional): A tuple (frequency_hz, name) for a single marker to highlight.
                                              Defaults to None.
    Process:
        1. Reads scan data from the specified CSV file into a pandas DataFrame.
        2. **Calculates 'Frequency_MHz' column if not present.**
        3. Creates a Plotly `Figure` object.
        4. Adds the main scan trace (Frequency vs. Power).
        5. Adds TV band markers if `include_tv_markers` is True.
        6. Adds Government band markers if `include_gov_markers` is True.
        7. **Adds a single marker if `single_marker_data` is provided.**
        8. Configures plot layout (title, axes labels, theme).
        9. Saves the plot as an HTML file.
        10. Optionally opens the HTML file in the default web browser.
        11. Handles potential `FileNotFoundError` during CSV reading.
    Outputs: None (generates HTML file, may open browser)
    """
    try:
        df = pd.read_csv(csv_file)
    except FileNotFoundError:
        messagebox.showerror("Plotting Error", f"Scan data CSV file not found: {csv_file} in {os.path.basename(__file__)}")
        print(f"❌ Plotting Error: Scan data CSV file not found: {csv_file} in {os.path.basename(__file__)}")
        return

    # Check if 'Frequency_Hz' column exists. If not, it's an error.
    if 'Frequency_Hz' not in df.columns:
        messagebox.showerror("Plotting Error", f"Required column 'Frequency_Hz' not found in CSV: {csv_file} in {os.path.basename(__file__)}")
        print(f"❌ Plotting Error: Missing required column 'Frequency_Hz' in CSV: {csv_file} in {os.path.basename(__file__)}")
        return

    # Always calculate 'Frequency_MHz' from 'Frequency_Hz' to ensure consistency
    df['Frequency_MHz'] = df['Frequency_Hz'] / MHZ_TO_HZ


    fig = go.Figure()

    # Add main scan trace
    fig.add_trace(go.Scatter(x=df['Frequency_MHz'], y=df['Power_dBm'], mode='lines', name='Scan Data',
                             line=dict(color='cyan', width=2)))

    # Add TV Band Markers
    if include_tv_markers:
        for band in TV_PLOT_BAND_MARKERS:
            fig.add_vrect(x0=band['Start MHz'], x1=band['Stop MHz'],
                          fillcolor="LightSalmon", opacity=0.2, layer="below", line_width=0,
                          annotation_text=band['Band Name'], annotation_position="top left",
                          annotation_font_color="white")

    # Add Government Band Markers
    if include_gov_markers:
        for band in GOV_PLOT_BAND_MARKERS:
            fig.add_vrect(x0=band['Start MHz'], x1=band['Stop MHz'],
                          fillcolor="MediumPurple", opacity=0.2, layer="below", line_width=0,
                          annotation_text=band['Band Name'], annotation_position="top left",
                          annotation_font_color="white")

    # Add single marker if provided
    if single_marker_data:
        freq_hz, name = single_marker_data
        freq_mhz = freq_hz / MHZ_TO_HZ
        
        # Find the power level at or near the marker frequency
        # This is a simple nearest-neighbor lookup; for more accuracy, interpolation could be used.
        closest_point = df.iloc[(df['Frequency_Hz'] - freq_hz).abs().argsort()[:1]]
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


    # Update layout
    fig.update_layout(
        title={
            'text': f"Spectrum Scan - {plot_title_suffix}",
            'yref': 'paper',
            'y': 0.9,
            'x': 0.5,
            'xanchor': 'center',
            'yanchor': 'top',
            'font': dict(color='white')
        },
        xaxis_title="Frequency (MHz)",
        yaxis_title="Power (dBm)",
        plot_bgcolor='black',
        paper_bgcolor='black',
        font=dict(color='white'),
        hovermode="x unified",
        xaxis=dict(gridcolor='gray', zerolinecolor='gray', showline=True, linecolor='white'),
        yaxis=dict(gridcolor='gray', zerolinecolor='gray', showline=True, linecolor='white')
    )

    # Save and open
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    pyo.plot(fig, filename=output_html_path, auto_open=auto_open_browser, include_plotlyjs='cdn')
    print(f"Plot saved to: {output_html_path}")

def generate_average_plot_logic(app_instance):
    """
    Generates an averaged Plotly HTML plot from all collected scan dataframes
    and optionally opens it in a browser.

    Inputs:
        app_instance (App): The main application instance, providing access to
                            collected scan data, settings, and output folder.
    Process:
        1. Checks if any scan data has been collected. If not, shows a warning.
        2. Concatenates all collected DataFrames and calculates the mean power for each frequency.
        3. Creates a Plotly `Figure` object.
        4. Adds the averaged scan trace.
        5. Adds TV band markers if `include_tv_markers` is True.
        6. Adds Government band markers if `include_gov_markers` is True.
        7. Configures plot layout.
        8. Saves the plot as an HTML file.
        9. Optionally opens the HTML file in the default web browser.
        10. Handles potential errors during plotting.
    Outputs: None (generates HTML file, may open browser)
    """
    if not app_instance.collected_scans_dataframes:
        messagebox.showwarning("No Data", "No scan data collected yet to generate an average plot.")
        print("🚫 No scan data collected for average plot.")
        return

    try:
        # Concatenate all dataframes and calculate the mean for each frequency
        all_scans_df = pd.concat(app_instance.collected_scans_dataframes)
        # Group by Frequency_MHz and calculate the mean of Power_dBm
        averaged_df = all_scans_df.groupby('Frequency_MHz')['Power_dBm'].mean().reset_index()
        averaged_df = averaged_df.sort_values(by='Frequency_MHz')

        fig = go.Figure()

        # Add averaged scan trace
        fig.add_trace(go.Scatter(x=averaged_df['Frequency_MHz'], y=averaged_df['Power_dBm'], mode='lines', name='Averaged Scan Data',
                                 line=dict(color='lime', width=2)))

        # Add TV Band Markers
        if app_instance.include_tv_markers_var.get():
            for band in TV_PLOT_BAND_MARKERS:
                fig.add_vrect(x0=band['Start MHz'], x1=band['Stop MHz'],
                              fillcolor="LightSalmon", opacity=0.2, layer="below", line_width=0,
                              annotation_text=band['Band Name'], annotation_position="top left",
                              annotation_font_color="white")

        # Add Government Band Markers
        if app_instance.include_gov_markers_var.get():
            for band in GOV_PLOT_BAND_MARKERS:
                fig.add_vrect(x0=band['Start MHz'], x1=band['Stop MHz'],
                              fillcolor="MediumPurple", opacity=0.2, layer="below", line_width=0,
                              annotation_text=band['Band Name'], annotation_position="top left",
                              annotation_font_color="white")

        # Update layout
        plot_title_suffix = f"Averaged Scan ({len(app_instance.collected_scans_dataframes)} cycles)"
        fig.update_layout(
            title={
                'text': f"Spectrum Scan - {plot_title_suffix}",
                'yref': 'paper',
                'y': 0.9,
                'x': 0.5,
                'xanchor': 'center',
                'yanchor': 'top',
                'font': dict(color='white')
            },
            xaxis_title="Frequency (MHz)",
            yaxis_title="Power (dBm)",
            plot_bgcolor='black',
            paper_bgcolor='black',
            font=dict(color='white'),
            hovermode="x unified",
            xaxis=dict(gridcolor='gray', zerolinecolor='gray', showline=True, linecolor='white'),
            yaxis=dict(gridcolor='gray', zerolinecolor='gray', showline=True, linecolor='white')
        )

        # Save and open
        output_dir = app_instance.output_folder_var.get()
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        datetime_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_html_path = os.path.join(output_dir, f"Averaged_Scan_{datetime_str}.html")

        pyo.plot(fig, filename=output_html_path, auto_open=True, include_plotlyjs='cdn')
        print(f"✅ Averaged plot saved to: {output_html_path}")

    except Exception as e:
        messagebox.showerror("Plotting Error", f"An error occurred during average plot generation: {e} in {os.path.basename(__file__)}")
        print(f"❌ Error generating average plot: {e} in {os.path.basename(__file__)}")