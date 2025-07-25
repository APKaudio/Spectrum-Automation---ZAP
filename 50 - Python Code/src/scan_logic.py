# src/scan_logic.py
import tkinter as tk
from tkinter import messagebox
import threading
import time
import os
from datetime import datetime
import pandas as pd

from utils.scan_instrument import scan_bands
from utils.frequency_bands import MHZ_TO_HZ

def start_scan_thread_logic(app_instance):
    if not app_instance.inst:
        messagebox.showwarning("Not Connected", "Please connect to an instrument first.")
        return
    
    if app_instance.scanning:
        messagebox.showwarning("Scan in Progress", "A scan is already running.")
        return

    from src.config_manager import save_config
    save_config(app_instance)

    app_instance.start_scan_button.config(state=tk.DISABLED)
    app_instance.stop_scan_button.config(state=tk.NORMAL)
    app_instance.pause_resume_button.config(state=tk.NORMAL)
    app_instance.connect_button.config(state=tk.DISABLED)
    app_instance.disconnect_button.config(state=tk.DISABLED)
    app_instance.apply_button.config(state=tk.DISABLED)
    app_instance.load_preset_button.config(state=tk.DISABLED)
    app_instance._stop_connect_button_blink()

    app_instance.scanning = True
    app_instance.paused = False
    app_instance.pause_resume_button.config(text="Pause Scan")

    print("\nStarting continuous spectrum scan...")
    
    max_hold_enabled = app_instance.desired_max_hold_var.get()
    max_hold_time = float(app_instance.desired_max_hold_time_var.get()) if max_hold_enabled else 0
    
    scan_rbw_segmentation = float(app_instance.desired_scan_rbw_segmentation_var.get())
    freq_shift_value = float(app_instance.shift_freq_var.get()) 

    rbw_config_val = scan_rbw_segmentation
    vbw_config_val = int(rbw_config_val / 3)

    selected_bands = [item["band"] for item in app_instance.band_vars if item["var"].get()]
    if not selected_bands:
        messagebox.showwarning("No Bands Selected", "Please select at least one frequency band to scan.")
        print("🚫 No bands selected for scan.")
        stop_scan_logic(app_instance)
        return

    base_output_dir = app_instance.output_folder_var.get()
    if not os.path.exists(base_output_dir):
        os.makedirs(base_output_dir)
        print(f"Created base output directory: {base_output_dir}")

    app_instance.scan_cycle_count = 0
    app_instance.current_freq_offset = 0

    scan_thread = threading.Thread(target=run_scan_logic, 
                                   args=(app_instance, selected_bands, 
                                         scan_rbw_segmentation, freq_shift_value, 
                                         rbw_config_val, vbw_config_val, max_hold_time))
    scan_thread.daemon = True
    scan_thread.start()

def toggle_pause_scan_logic(app_instance):
    if app_instance.scanning:
        app_instance.paused = not app_instance.paused
        if app_instance.paused:
            app_instance.pause_resume_button.config(text="Resume Scan", bg="blue")
            print("Scan Paused. Click Resume to continue.")
            print("Scan paused.")
        else:
            app_instance.pause_resume_button.config(text="Pause Scan", bg="orange")
            print("Scan Resumed.")
            print("Scan resumed.")
    else:
        messagebox.showwarning("Scan Not Active", "No scan is currently running to pause or resume.")

def run_scan_logic(app_instance, selected_bands, scan_rbw_segmentation, freq_shift_value, rbw_config_val, vbw_config_val, max_hold_time):
    try:
        while app_instance.scanning:
            while app_instance.paused:
                print("Scan Paused. Click Resume to continue.")
                time.sleep(0.5)
                if not app_instance.scanning:
                    print("\nScan process finished (interrupted).")
                    print("Scan interrupted by user.")
                    break

            if not app_instance.scanning:
                print("\nScan process finished (interrupted).")
                print("Scan interrupted by user.")
                break

            print(f"\n--- Starting Scan Cycle {app_instance.scan_cycle_count + 1} ---")
            print(f"Current Frequency Offset: {app_instance.current_freq_offset} Hz (Applied to all band frequencies)")
            print(f"Scan RBW: {scan_rbw_segmentation} Hz (Constant)")

            scan_name = app_instance.scan_name_var.get()
            if not scan_name:
                scan_name = "UnnamedScan"
            
            rbw_str = f"RBW{int(scan_rbw_segmentation/1000):04d}"
            max_hold_time_val = float(app_instance.desired_max_hold_time_var.get()) if app_instance.desired_max_hold_var.get() else 0
            hold_str = f"HOLD{int(max_hold_time_val):02d}"
            offset_str = f"Offset{int(app_instance.current_freq_offset)}"

            datetime_str = datetime.now().strftime("%Y%m%d_%H%M%S") 
            
            html_plot_path_for_single_scan = os.path.join(app_instance.output_folder_var.get(), f"{scan_name}_{rbw_str}_{hold_str}_{offset_str}_{datetime_str}.html")

            try:
                scanned_data, last_successful_band_index, current_scan_csv_path = scan_bands(
                    app_instance, app_instance.inst, selected_bands, 
                    scan_rbw_segmentation, rbw_config_val, 
                    vbw_config_val, max_hold_time, app_instance.current_freq_offset
                ) 
                
                if not app_instance.scanning:
                    print("\nScan process finished (interrupted after band scan).")
                    print("Scan interrupted by user.")
                    if scanned_data:
                        plot_suffix = f"{scan_name}_{rbw_str}_{hold_str}_{offset_str}_{datetime_str}_INTERRUPTED"
                        html_plot_path_for_single_scan_interrupted = os.path.join(app_instance.output_folder_var.get(), f"{plot_suffix}.html")
                        app_instance.after(0, app_instance.generate_single_scan_plot_and_open_wrapper, current_scan_csv_path, plot_suffix, html_plot_path_for_single_scan_interrupted, False)
                    break

                if scanned_data:
                    df_scan = pd.DataFrame(scanned_data, columns=['Frequency_Hz', 'Power_dBm'])
                    df_scan['Frequency_MHz'] = df_scan['Frequency_Hz'] / MHZ_TO_HZ
                    app_instance.collected_scans_dataframes.append(df_scan[['Frequency_MHz', 'Power_dBm']].copy())
                    print(f"✅ Stored scan data for averaging. Total scans collected: {len(app_instance.collected_scans_dataframes)}")

                app_instance.last_scan_data = scanned_data
                
                print(f"Cycle scan finished. Plot will be generated from CSV: {current_scan_csv_path}")

                plot_suffix = f"{scan_name}_{rbw_str}_{hold_str}_{offset_str}_{datetime_str}"
                app_instance.after(0, app_instance.generate_single_scan_plot_and_open_wrapper, current_scan_csv_path, plot_suffix, html_plot_path_for_single_scan, app_instance.open_html_after_complete_var.get()) 
                
                app_instance.scan_cycle_count += 1
                app_instance.current_freq_offset += freq_shift_value

                if app_instance.scan_cycle_count >= 10:
                    print(f"🎉 {app_instance.scan_cycle_count} scan cycles completed. Resetting frequency offset to 0 Hz.")
                    app_instance.current_freq_offset = 0
                    app_instance.scan_cycle_count = 0

                if not app_instance.scanning:
                    print("\nScan process finished (interrupted).")
                    print("Scan interrupted by user.")
                    break
                
                wait_time = float(app_instance.desired_cycle_wait_time_var.get())
                if wait_time > 0:
                    print(f"Waiting {wait_time} seconds for next cycle...")
                    print(f"Waiting {wait_time} seconds before next scan cycle...")
                    for _ in range(int(wait_time * 10)):
                        while app_instance.paused:
                            print("Scan Paused. Click Resume to continue.")
                            time.sleep(0.1)
                            if not app_instance.scanning:
                                print("\nScan process finished (interrupted during pause in wait).")
                                print("Scan interrupted during pause in wait.")
                                break
                        
                        if not app_instance.scanning:
                            print("\nScan process finished (interrupted during wait).")
                            print("Scan interrupted during wait.")
                            break
                        time.sleep(0.1)

            except Exception as e:
                app_instance.after(0, messagebox.showerror, "Scan Cycle Error", f"An error occurred during scan cycle: {e}")
                print(f"❌ Scan cycle encountered an error: {e}")
                print(f"Scan cycle error: {e}")
                app_instance.scanning = False
                break

        print("\nContinuous scan process terminated.")
        print("Continuous scan terminated.")
        
    except Exception as e:
        app_instance.after(0, messagebox.showerror, "Scan Thread Error", f"An unexpected error occurred in main scan thread: {e}")
        print(f"❌ Main scan thread encountered an error: {e}")
        print(f"Main scan thread error: {e}")
    finally:
        app_instance.scanning = False
        app_instance.paused = False
        app_instance.after(100, reset_scan_buttons_logic, app_instance)
        if not app_instance.inst and app_instance.instrument_list and app_instance.resource_var.get() != "No resources found":
            app_instance._start_connect_button_blink()

def stop_scan_logic(app_instance):
    app_instance.scanning = False
    app_instance.paused = False
    print("\nAttempting to stop scan... Please wait for current sweep to finish.")
    app_instance.stop_scan_button.config(state=tk.DISABLED)
    app_instance.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED)

def reset_scan_buttons_logic(app_instance):
    app_instance.start_scan_button.config(state=tk.NORMAL)
    if app_instance.inst:
        app_instance.disconnect_button.config(state=tk.NORMAL)
        app_instance.apply_button.config(state=tk.NORMAL)
        if app_instance.preset_tree.selection() and app_instance.instrument_model != "N9340B":
            app_instance.load_preset_button.config(state=tk.NORMAL)
    app_instance.stop_scan_button.config(state=tk.DISABLED)
    app_instance.pause_resume_button.config(text="Pause Scan", state=tk.DISABLED)
