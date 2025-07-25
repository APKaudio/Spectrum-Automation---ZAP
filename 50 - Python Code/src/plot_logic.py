# src/plot_logic.py
import tkinter as tk
from tkinter import messagebox
import pandas as pd

from utils.plotting_utils import plot_single_scan_data, _open_plot_in_browser
from utils.averaging_utils import generate_historical_average_plot
from utils.frequency_bands import MHZ_TO_HZ

def generate_single_scan_plot_and_open_wrapper_logic(app_instance, csv_file_path, plot_title_suffix, output_html_path, auto_open_browser=True):
    print(f"Generating single scan plot from CSV: {csv_file_path}...")
    try:
        df = pd.read_csv(csv_file_path, header=None, names=["Frequency_MHz", "Power_dBm"])
        scanned_data_from_csv = list(zip(df['Frequency_MHz'] * MHZ_TO_HZ, df['Power_dBm']))

        fig, saved_html_path = plot_single_scan_data(
            scanned_data_from_csv, 
            plot_title_suffix,
            include_tv_markers=app_instance.include_tv_markers_var.get(),
            include_gov_markers=app_instance.include_gov_markers_var.get(),
            output_html_path=output_html_path,
            auto_open_browser=auto_open_browser
        )
        
        if fig and saved_html_path:
            print(f"✅ Single scan plot generation complete: {saved_html_path}")
        else:
            print("🚫 Plotly figure was not generated or saved for single scan data.")
    except Exception as e:
        messagebox.showerror("Single Plot Error", f"Failed to generate single scan plot from CSV '{csv_file_path}': {e}")
        print(f"❌ Error generating single scan plot from CSV: {e}")

def generate_average_plot_logic(app_instance):
    if app_instance.scanning and not app_instance.paused:
        messagebox.showwarning("Plotting Error", "Cannot generate historical average plot while a scan is in progress. Please pause or stop the scan first.")
        return

    generate_historical_average_plot(
        app_instance.scan_name_var,
        app_instance.output_folder_var,
        app_instance.open_html_after_complete_var,
        app_instance.include_tv_markers_var,
        app_instance.include_gov_markers_var
    )
