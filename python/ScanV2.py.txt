import pyvisa
import time
import csv
from datetime import datetime
import os
import argparse

# -----------------------------
# Command-line argument parsing
# -----------------------------
parser = argparse.ArgumentParser(description="Spectrum Analyzer Sweep and CSV Export")

parser.add_argument('--SCANname', type=str, default="25kz scan ",
                    help='Prefix for the output CSV filename')

parser.add_argument('--startFreq', type=float, default=400e6,
                    help='Start frequency in Hz')

parser.add_argument('--endFreq', type=float, default=650e6,
                    help='End frequency in Hz')

parser.add_argument('--stepSize', type=float, default=25000,
                    help='Step size in Hz')

args = parser.parse_args()

file_prefix = args.SCANname
start_freq = args.startFreq
end_freq = args.endFreq
step = args.stepSize

# ---------------------
# VISA connection setup
# ---------------------
visa_address = 'USB0::0x0957::0xFFEF::CN03480580::0::INSTR'
rm = pyvisa.ResourceManager()
inst = rm.open_resource(visa_address)

# Reset and configure
inst.write("*CLS")
inst.write("*RST")
inst.query("*OPC?")

inst.write(":DISP:WIND:TRAC:Y:SCAL LOG")       # Set Y axis to dBm
inst.write(":DISP:WIND:TRAC:Y:RLEV -30")       # Set reference level
inst.write(":CALC:MARK1 ON")                   # Enable marker 1
inst.write(":CALC:MARK1:MODE POS")
inst.write(":CALC:MARK1:ACT")

time.sleep(2)  # Allow instrument to settle

# -----------------
# File & directory
# -----------------
scan_dir = os.path.join(os.getcwd(), "N9340 Scans")
os.makedirs(scan_dir, exist_ok=True)

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
filename = os.path.join(scan_dir, f"{file_prefix}--{timestamp}.csv")

# -------------------------
# Sweep and write to CSV
# -------------------------
segment_width = 10_000_000  # 10 MHz
step_int = int(step)
scan_limit = int(end_freq)

with open(filename, mode='w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["Frequency (MHz)", "Level (dBm)"])

    current_block_start = int(start_freq)

    while current_block_start < scan_limit:
        current_block_stop = current_block_start + segment_width

        print(f"Sweeping range {current_block_start / 1e6:.3f} to {current_block_stop / 1e6:.3f} MHz")
        inst.write(f":FREQ:START {current_block_start}")
        inst.write(f":FREQ:STOP {current_block_stop}")
        inst.write(":INIT:IMM")
        time.sleep(2)

        current_freq = current_block_start
        while current_freq <= current_block_stop:
            inst.write(f":CALC:MARK1:X {current_freq}")
            level_raw = inst.query(":CALC:MARK1:Y?").strip()

            try:
                level = float(level_raw)
                level_formatted = f"{level:.1f}"
                freq_mhz = current_freq / 1_000_000
                print(f"{freq_mhz:.3f} MHz : {level_formatted} dBm")
                writer.writerow([freq_mhz, level_formatted])

                
            except ValueError:
                level_formatted = level_raw

            
            current_freq += step_int

        current_block_start = current_block_stop

# -----------------
# Cleanup
# -----------------
try:
    inst.write("SYST:LOC")  # may not work on N9340B, but safe to try
except pyvisa.VisaIOError:
    pass

inst.close()
print(f"\nScan complete. Results saved to '{filename}'")
