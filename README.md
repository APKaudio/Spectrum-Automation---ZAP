# **RF Spectrum Analyzer Controller: User Manual**

**This application provides a graphical interface to control your Keysight N9340B/N9342CN spectrum analyzer. It automates the process of scanning frequency bands, collecting data, and generating interactive plots. A key feature is its ability to perform a series of scans with a small frequency shift on each, allowing you to build a more detailed picture of the spectrum by averaging the results.**

![**Configuration Screen**](https://github.com/APKaudio/Spectrum-Automation---ZAP/blob/main/Screen%20shots%20and%20Demo%20Scan/V10.8.1%20screen.png?raw=true)

## **1. Getting Started**

**Before you run the application, ensure you have Python installed. The application will automatically check for and offer to install necessary libraries like **`<span class="selected">pyvisa</span>`**, **`<span class="selected">pandas</span>`**, **`<span class="selected">plotly</span>`**, and **`<span class="selected">numpy</span>`** if they are missing.**

**To run the application:**

1. **Save the provided Python code as **`<span class="selected">ScanV10.7.9.py</span>`**.**
2. **Make sure the **`<span class="selected">frequency_bands.py</span>`** file (which defines common frequency ranges) is in the same directory as **`<span class="selected">ScanV10.7.9.py</span>`**.**
3. **Execute the script using a Python interpreter (e.g., **`<span class="selected">python ScanV10.7.9.py</span>`** from your command prompt).**

## **2. Connecting to Your Instrument**

**The "Instrument Connection" section at the top of the left panel allows you to establish communication with your spectrum analyzer.**

* ****VISA Resource:****** This dropdown lists all detected instruments connected to your computer via VISA (Virtual Instrument Software Architecture).**
  * **Click the dropdown to see available instruments (e.g., "ASRL1::INSTR", "USB0::...::INSTR").**
* ****Refresh Button:****** If your instrument isn't listed, or if you've just connected it, click "Refresh" to update the list of available VISA resources.**
* ****Connect Button:****** Select your instrument from the dropdown and click "Connect." The application will attempt to establish communication. If successful, the console output will confirm the connection and display your instrument's model. The "Start Scan" and "Apply Settings to Device" buttons will become active.**
* ****Disconnect Button:****** Click this to safely close the connection to your instrument.**

## **3. Scan Configuration (Push to Device)**

**This section allows you to define how the spectrum analyzer will perform its scans. These settings are "pushed" to the device, meaning they configure the hardware directly.**

* ****Reference Level (dBm):****** Sets the maximum power level displayed on the analyzer's screen. Adjust this based on the expected signal strengths to avoid clipping or to optimize visibility of weak signals.**
* ****High Sensitivity (Preamplifier):****** Check this box to enable the instrument's high sensitivity mode, which typically turns on an internal preamplifier. This is useful for detecting very weak signals.**
* ****Preamplifier (ON/OFF):****** This checkbox also controls the instrument's preamplifier. It's generally recommended to use "High Sensitivity" for overall weak signal reception.**
* ****Scan RBW (for Segmentation, Hz):****** This is the Resolution Bandwidth that the application uses to segment the overall frequency bands into smaller sweeps. A smaller RBW provides finer frequency resolution but increases scan time. This value remains constant during a scan cycle.**
* ****Frequency Shift (Hz, per cycle):****** This is a unique feature.**
  * **It defines a frequency increment (e.g., 1000 Hz).**
  * **For each *****subsequent***** full scan cycle (after the first), the entire frequency range of *****all***** selected bands will be shifted by this amount.**
  * **For example, if a band is 200-300 MHz and the shift is 1000 Hz:**
    * **Scan 1: 200 MHz to 300 MHz**
    * **Scan 2: 200.001 MHz to 300.001 MHz**
    * **Scan 3: 200.002 MHz to 300.002 MHz**
    * **...and so on.**
  * ****Reset Logic:****** After 10 full scan cycles (meaning the total shift has reached 10 times the "Frequency Shift" value, e.g., 10,000 Hz if the shift is 1000 Hz), the frequency offset will automatically reset to 0 Hz, and the cycle count will restart. This ensures you cover a full "shifted" range before repeating the pattern.**
* ****Max Hold Time (s):****** When enabled (and Max Hold trace is active on the instrument), this is the duration in seconds that the analyzer will wait to capture the maximum power level at each frequency point across multiple sweeps within a segment. A longer time ensures a more accurate "peak" reading.**
* ****Cycle Wait Time (s):****** The duration the application will pause between completing one full scan cycle (all selected bands) and starting the next.**
* ****Scan Name:****** A descriptive name for your scan. This name will be used as part of the filenames for your saved data and plots.**
* ****Output Folder (Base):****** The directory where all your CSV data and HTML plots will be saved.**
  * ****Open Folder Button:****** Click this to quickly open the specified output folder in your system's file explorer.**
* ****Plot Options:****
  * ****Include TV Band Markers:****** If checked, the generated plots will include shaded regions and labels for common TV broadcast frequency bands.**
  * ****Include Government Band Markers:****** If checked, the generated plots will include shaded regions and labels for common government/military frequency bands.**
  * ****Open HTML Plot After Each Cycle:****** If checked, a new interactive HTML plot will automatically open in your web browser after each complete scan cycle.**
* ****Apply Settings to Device Button:****** Click this after changing any settings in this section to send the updated configuration to your connected spectrum analyzer.**

## **4. Frequency Band Selection**

**This section allows you to choose which predefined frequency ranges the application will scan.**

* **A list of common frequency bands (e.g., FM Radio, Cellular, Wi-Fi) is displayed with their start and stop frequencies.**
* **Check the box next to each band you wish to include in your scan. By default, all bands are selected.**

## **5. Device Preset Files (C:\PRESETS)**

**This section allows you to manage preset configurations stored directly on your spectrum analyzer.**

* **The list will show **`<span class="selected">.STA</span>`** (state) files found in the **`<span class="selected">C:\PRESETS\</span>`** directory on your connected instrument.**
* **Files containing "MON" in their name will be highlighted in blue for easy identification.**
* ****Load Selected Preset Button:****** Select a preset file from the list and click this button to load that configuration onto your spectrum analyzer. This is useful for quickly recalling specific instrument setups.**

## **6. Controlling the Scan**

**The buttons at the top of the console output area control the scanning process.**

* ****Start Scan Button:****** Initiates the continuous scanning process based on your configured settings and selected bands. This will run in a loop, applying frequency shifts as configured.**
* ****Pause Scan / Resume Scan Button:****** Toggles the scanning process between paused and running states.**
* ****Stop Scan Button:****** Halts the ongoing scanning process. The current sweep will complete, and then the scan will stop.**

![Creating Plot](https://github.com/APKaudio/Spectrum-Automation---ZAP/blob/main/Screen%20shots%20and%20Demo%20Scan/newplot%20(16).png?raw=true)

## **7. Generating Plots**

**The application automatically generates plots after each scan cycle, and you can also generate a comprehensive historical plot.**

* ****Generate Plot (Average) Button:****** Click this button to create a powerful historical analysis plot.**
  * **It will scan your **`<span class="selected">Output Folder</span>`** for all previously saved CSV files that match your current "Scan Name" and contain frequency offset information.**
  * ****Smart Averaging:****** It intelligently "un-shifts" the frequencies of these historical scans back to a common reference point. This allows it to accurately calculate the ******Average Power**** **, ******Median Power**** **, and ******Range (Max - Min)****** across all your shifted scans for each frequency point.**
  * **The generated plot will display these average, median, and range traces.**
  * **Crucially, it will also overlay *****each individual historical scan***** (at its *****original, shifted***** frequencies) as a transparent line, allowing you to see how the individual shifted scans contribute to the overall average and range.**
  * **This plot provides a comprehensive view of the spectrum, accounting for the frequency shifting you've applied.**
  * 

## **8. Outputs**

**The application produces the following files in your specified "Output Folder":**

* ****CSV Files (Comma Separated Values):****
  
  * ****Individual Scan Files:****** For each completed scan cycle, a CSV file will be created with a filename like **`<span class="selected">[ScanName]_RBW[RBWValue]K_HOLD[HoldTime]_[OffsetValue]_[YYYYMMDD_HHMMSS].csv</span>`**.**
    * **Example: **`<span class="selected">MyScan_RBW0010K_HOLD03_Offset1000_20250721_171942.csv</span>`
    * **These files contain two columns: "Frequency (MHz)" (the *****actual***** scanned frequency, including any applied offset for that cycle) and "Level (dBm)".**
  * ****Averaged Cycle Files (if applicable):****** If you use the **`<span class="selected">_generate_current_cycle_average_csv_and_plot</span>`** function (which happens internally after a set of scans or on demand), you might see files like **`<span class="selected">[ScanName]_[HHMM]_averaged_cycle.csv</span>`**. These contain the average, median, and range for a batch of recent scans.**
  * ****Historical Averaged Files:****** When you click "Generate Plot (Average)", a CSV file like **`<span class="selected">[ScanName]_HISTORICAL_[HHMM]_averaged.csv</span>`** is created. This file contains the normalized frequency, average power, median power, and range (Max-Min) calculated from *****all***** relevant historical individual scan CSVs.**
* ****HTML Plot Files:****
  
  * ****Individual Scan Plots:****** For each completed scan cycle, an interactive HTML plot (e.g., **`<span class="selected">MyScan_RBW0010K_HOLD03_Offset1000_20250721_171942.html</span>`**) is generated. You can open these in any web browser to view the spectrum data for that specific scan, including any applied frequency offset.**
  * ****Historical Averaged Plots:****** When you click "Generate Plot (Average)", a comprehensive HTML plot (e.g., **`<span class="selected">MyScan_HISTORICAL_[HHMM]_averaged.html</span>`**) is created. This plot displays the averaged, median, and range traces (on a normalized frequency axis) along with transparent overlays of all the individual shifted scans that contributed to the average. This allows for in-depth analysis of spectral trends over your shifted measurement cycles.**
* 

![Commands](https://github.com/APKaudio/Spectrum-Automation---ZAP/blob/main/Screen%20shots%20and%20Demo%20Scan/V10.8.1%20commands%20screen.png?raw=true)

## **9. Console Output**

**The large black area on the right side of the application window displays real-time messages from the application. This includes:**

* **Connection status and instrument identification.**
* **SCPI commands being sent to the instrument (if debug mode is enabled).**
* **Scan progress updates, including current band, segment, and frequency range.**
* **Information about data saving and plot generation.**
* **Any errors or warnings encountered during operation.**

## 🚀 Interactive Setup Guide

*For a smooth setup of `ScanV10.8.1.py`*

---

## 1. Get Python 🐍

The first step is to install Python, the programming language needed to run the script. This section will guide you to the official download page and highlight the most important part of the installation process.

**Go to the official Python website to download the correct installer for your operating system (Windows, macOS, or Linux):**

👉 [**Download Python**](https://www.python.org/downloads/)

> ⚠️ **Critical Tip for Windows Users!**
> On the first screen of the installer, you **must** check the box that says
> **"Add Python X.Y to PATH"**.
> This allows your computer to find Python easily from any folder.

---

## 2. Install Helper Tools 🛠️

The script needs a few helper tools (libraries) to function correctly. Use `pip`, Python's package installer, in your **Command Prompt (Windows)** or **Terminal (macOS/Linux)**.

```bash
pip install pyvisa
```

```bash
pip install numpy
```

```bash
pip install pandas
```

```bash
pip install plotly
```

---

## 3. Prepare Your Workspace 🏠

Before running the script:

### 1. Save the Script

Save the `ScanV9.3.py` file into a folder. For example, create a folder named `ScanApp` on your C: drive (Windows) or in your home directory (macOS/Linux).

### 2. Navigate to the Folder

Use the `cd` (change directory) command. Example for Windows:

```bash
cd C:\ScanApp
```

---

## 4. Run the Script! ▶️

You're all set! With Python installed and your workspace ready, you can now run the script:

```bash
python ScanV9.3.py
```

