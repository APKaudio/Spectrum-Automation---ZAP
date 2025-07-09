
# 🚀 Interactive Setup Guide

*For a smooth setup of `ScanV9.3.py`*

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


# 📊 Spectrum Analyzer Scan Configuration – Parameter Guide

![User Interface](https://github.com/APKaudio/Spectrum-Automation---ZAP/blob/main/Screen%20shots%20and%20Demo%20Scan/User%20interface.png?raw=true)

This user interface configures and initiates a multi-band scan on a connected spectrum analyzer instrument using PyVISA. Here's what each field and control does:

---

## 🧪 Scan Parameters (Left Panel)

| Field | Description |
|-------|-------------|
| **Scan Series Name** | Sets a name prefix for the scan output files (e.g., CSV and HTML). Useful for organizing multiple scans. |
| **RBW Step Size (Hz) for Scan** | Defines how finely the instrument steps across frequency during scanning. A smaller value yields higher resolution (e.g., 10000 Hz = 10 kHz). |
| **Max Hold Time (seconds)** | Time (in seconds) to hold each segment for peak detection (using MAXHold trace mode). Higher values allow better signal capture. |
| **Cycle Wait Time (seconds)** | Delay between consecutive full scan cycles. Useful in continuous monitoring mode. |

---

## ⚙️ Instrument Initial Configuration (Middle Panel)

| Field | Description |
|-------|-------------|
| **Clear and Reset Instrument (*CLS; *RST)** | Sends SCPI commands to clear and reset the instrument before each scan. |
| **Set RBW (e.g., 1KHZ, 10KHZ)** | Controls the instrument’s Resolution Bandwidth setting during scan (affects frequency resolution). |
| **Set VBW (e.g., 1KHZ, 10KHZ)** | Controls the Video Bandwidth (affects signal smoothing). |
| **Preamplifier ON for high sensitivity** | Enables the internal preamplifier for better detection of weak signals. |
| **Display set to logarithmic scale** | Sets the vertical axis to logarithmic (dBm) scale, typical for RF spectrum. |
| **Display Reference Level (dBm)** | Sets the top level of the display in dBm. This affects vertical scale of measurement. |
| **Trace 1 set to max hold** | Sets trace mode to MAXHold, which keeps the peak value observed. Important for detecting intermittent signals. |

---

## 🖼️ Plot Options (Right Panel)

| Option | Description |
|--------|-------------|
| **Include Government Band Markers** | Overlays colored regions in the HTML plot for known government spectrum allocations (defined in script). |
| **Include TV Channel Markers** | Adds markers for TV channels (e.g., CH2, CH14) based on known frequency allocations. |
| **Open HTML after complete** | Automatically opens the resulting HTML plot after scan completion. |

---

## 📡 Select Frequency Bands to Scan (Bottom Panel)

| Option | Description |
|--------|-------------|
| **Low VHF+FM (50-110 MHz)** through **2 GHz Cams (2000-2390 MHz)** | Allows users to choose which frequency bands to scan. These are defined in the `SCAN_BAND_RANGES` list in the script. Only selected bands will be scanned. |

---

## 🖱️ Button Bar (Top Row)

| Button | Function |
|--------|----------|
| **Start Scan** | Initiates the scan using all parameters configured in the GUI. |
| **System Restart** | Sends the `:SYSTem:POWer:RESet` SCPI command to reboot the instrument. |
| **Reset Instrument** | Sends a soft reset `*RST` command to restore defaults without reboot. |
| **Quit** | Exits the GUI. |

---

## 🔌 VISA Instrument Connection

At the top, the **Connected Instrument** label displays the VISA address of the attached instrument (e.g., `USB0::0x0957::...`). This is detected automatically on startup.

**That's it! Happy scanning! 😊**
