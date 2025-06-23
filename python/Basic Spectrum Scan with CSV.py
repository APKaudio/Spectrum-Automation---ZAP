import pyvisa
import time
import csv
from datetime import datetime
import os
import argparse

# Parse command-line arguments
parser = argparse.ArgumentParser(description="Spectrum Analyzer Sweep and CSV Export")

parser.add_argument('--SCANname', type=str, default="25kz scan ",
                    help='Prefix for the output CSV filename')

parser.add_argument('--startFreq', type=float, default=400e6,
                    help='Start frequency in Hz (default: 400000000)')

parser.add_argument('--endFreq', type=float, default=600e6,
                    help='End frequency in Hz (default: 600000000)')

parser.add_argument('--stepSize', type=float, default=25000,
                    help='Step size in Hz (default: 25000)')

args = parser.parse_args()

file_prefix = args.SCANname
start_freq = args.startFreq
end_freq = args.endFreq
step = args.stepSize

visa_address = 'USB0::0x0957::0xFFEF::CN03480580::0::INSTR'

# Create VISA resource manager and open instrument
rm = pyvisa.ResourceManager()
inst = rm.open_resource(visa_address)

# Clear and configure sweep
inst.write("*CLS")
inst.write(f":FREQ:START {start_freq}")
inst.write(f":FREQ:STOP {end_freq}")
inst.write(":INIT:CONT OFF")
inst.write(":INIT:IMM")
time.sleep(2)  # wait for sweep to complete

# Create 'SCANs' directory if it doesn't exist
scan_dir = os.path.join(os.getcwd(), "N9340 Scans")
os.makedirs(scan_dir, exist_ok=True)

# Prepare CSV filename with timestamp inside SCANs folder
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
filename = os.path.join(scan_dir, f"{file_prefix}--{timestamp}.csv")

# Open CSV file and write header
with open(filename, mode='w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["Frequency (Hz)", "Level (dBm)"])

    # Sweep and read levels
    current_freq = int(start_freq)
    end_freq_int = int(end_freq)
    step_int = int(step)

    while current_freq <= end_freq_int:
        inst.write(f":CALC:MARK1:X {current_freq}")
        level = inst.query(":CALC:MARK1:Y?").strip()
        print(f"{current_freq} Hz : {level} dBm")
        writer.writerow([current_freq, level])
        current_freq += step_int

# Close instrument connection
inst.close()

print(f"\nScan complete. Results saved to '{filename}'")
