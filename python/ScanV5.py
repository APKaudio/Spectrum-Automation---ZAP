import pyvisa
import time
import csv
from datetime import datetime
import os
import argparse
import sys

# Define constants for better readability and easier modification
MHZ_TO_HZ = 1_000_000  # Conversion factor from MHz to Hz
DEFAULT_SEGMENT_WIDTH_HZ = 10_000_000  # 10 MHz segment for sweeping (Corrected from 0_000_000)
DEFAULT_STEP_SIZE_HZ = 25000  # 25 kHz step size
DEFAULT_CYCLE_WAIT_TIME_SECONDS = 30  # 5 minutes wait between full scan cycles
DEFAULT_MAXHOLD_TIME_SECONDS = 10

# Define the frequency bands to scan
# Frequencies are stored in MHz for readability, will be converted to Hz for instrument commands
BAND_RANGES = [
    {"Band Name": "Low VHF+FM", "Start MHz": 54.000, "Stop MHz": 109.000},
    {"Band Name": "High VHF+216", "Start MHz": 174.000, "Stop MHz": 218.000},
    {"Band Name": "UHF -1", "Start MHz": 400.000, "Stop MHz": 700.000},
    {"Band Name": "UHF -2", "Start MHz": 700.000, "Stop MHz": 900.000},
    {"Band Name": "900 ISM-STL", "Start MHz": 900.000, "Stop MHz": 962.000},
    {"Band Name": "AFTRCC-1", "Start MHz": 1435.000, "Stop MHz": 1535.000},
    {"Band Name": "DECT-ALL", "Start MHz": 1880.000, "Stop MHz": 2000.000},
    {"Band Name": "2 GHz Cams", "Start MHz": 2000.000, "Stop MHz": 2395.000},
]

VISA_ADDRESSES = {
    'apk': 'USB0::0x0957::0xFFEF::CN03480580::0::INSTR',
    'zap': 'USB1::0x0957::0xFFEF::SG05300002::0::INSTR'
}

def setup_arguments():
    """
    Parses command-line arguments for the spectrum analyzer sweep.
    Allows customization of filename prefix, frequency range, step size, and user.
    """
    parser = argparse.ArgumentParser(description="Spectrum Analyzer Sweep and CSV Export")

    parser.add_argument('--SCANname', type=str, default="25kz scan ",
                        help='Prefix for the output CSV filename')
    parser.add_argument('--startFreq', type=float, default=None,
                        help='Start frequency in Hz (overrides default bands if provided with --endFreq)')
    parser.add_argument('--endFreq', type=float, default=None,
                        help='End frequency in Hz (overrides default bands if provided with --startFreq)')
    parser.add_argument('--stepSize', type=float, default=DEFAULT_STEP_SIZE_HZ,
                        help='Step size in Hz')
    parser.add_argument('--user', type=str, choices=['apk', 'zap'], default='zap',
                        help='Specify who is running the program: "apk" or "zap". Default is "zap".')

    return parser.parse_args()

def display_predefined_bands():
    """
    Prints a formatted table of the predefined frequency bands.
    """
    print("Predefined Scan Areas:")
    print("=" * 40)
    print(f"{'Band Name':<15} {'Start MHz':<12} {'Stop MHz':<12}")
    print("-" * 40)
    for band in BAND_RANGES:
        print(f"{band['Band Name']:<15} {band['Start MHz']:<12.3f} {band['Stop MHz']:<12.3f}")
    print("=" * 40 + "\n")

def initialize_instrument(visa_address):
    """
    Establishes a VISA connection to the instrument and performs initial configuration.

    Args:
        visa_address (str): The VISA address of the spectrum analyzer.

    Returns:
        pyvisa.resources.Resource: The instrument object if connection is successful, else None.
    """
    rm = pyvisa.ResourceManager()
    inst = None
    try:
        inst = rm.open_resource(visa_address)
        print(f"Connected to instrument at {visa_address}")

        # Clear and reset the instrument
        inst.write("*CLS")
        inst.write("*RST")
        inst.query("*OPC?") # Wait for operations to complete

        # Configure preamplifier for high sensitivity
        inst.write(":POWer:GAIN ON") # Equivalent to ':POWer:GAIN 1' for most Keysight instruments
        print("Preamplifier turned ON for high sensitivity.")

        # Configure display and marker settings
        inst.write(":DISP:WIND:TRAC:Y:SCAL LOG") # Logarithmic scale (dBm)
        inst.write(":DISP:WIND:TRAC:Y:RLEV -30") # Reference level
        inst.write(":CALC:MARK1 ON") # Enable Marker 1
        inst.write(":CALC:MARK1:MODE POS") # Marker 1 mode to position
        inst.write(":CALC:MARK1:ACT") # Activate Marker 1



        time.sleep(2) # Allow instrument to settle
        return inst
    except pyvisa.VisaIOError as e:
        print(f"VISA Error: Could not connect to or communicate with the instrument at {visa_address}: {e}")
        print("Please ensure the instrument is connected, powered on, and the VISA address is correct.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred during instrument initialization: {e}")
        return None

def perform_segment_sweep(inst, segment_start_freq, segment_stop_freq, step_size_hz, band_name, writer):
    """
    Performs a frequency sweep for a given segment and writes data to CSV.

    Args:
        inst (pyvisa.resources.Resource): The connected instrument object.
        segment_start_freq (int): Start frequency of the current segment in Hz.
        segment_stop_freq (int): Stop frequency of the current segment in Hz.
        step_size_hz (int): Step size in Hz.
        band_name (str): Name of the current frequency band.
        writer (csv._writer._writer): CSV writer object to write data.
    """
    print(f"  Sweeping segment: {segment_start_freq / MHZ_TO_HZ:.3f} MHz to {segment_stop_freq / MHZ_TO_HZ:.3f} MHz")

    # Set the start and stop frequencies for the instrument's sweep for this segment.
    inst.write(f":FREQ:START {segment_start_freq}")
    inst.write(f":FREQ:STOP {segment_stop_freq}")

    # Enable MAX hold on Trace 1
    inst.write(":TRACe1:MODE MAXHold")
    print("Trace 1 set to MAX Hold mode.")

    #Add settling time for max hold values to show up
    print(f"Waiting {DEFAULT_MAXHOLD_TIME_SECONDS} seconds for MAX hold to settle...")
    time.sleep(DEFAULT_MAXHOLD_TIME_SECONDS)

    


    #inst.write(":INIT:IMM") # Initiate a single immediate sweep.
    inst.query("*OPC?") # Wait for sweep to finish

    current_freq = segment_start_freq
    while current_freq <= segment_stop_freq:
        inst.write(f":CALC:MARK1:X {current_freq}") # Set Marker 1 to the current frequency.
        level_raw = inst.query(":CALC:MARK1:Y?").strip() # Query the Y-axis value (level in dBm).

        try:
            level = float(level_raw)
            level_formatted = f"{level:.1f}"
            freq_mhz = current_freq / MHZ_TO_HZ
            print(f"    Band: {band_name:<15} Freq: {freq_mhz:.3f} MHz : {level_formatted} dBm")
            writer.writerow([freq_mhz, level_formatted])
        except ValueError:
            print(f"    Warning: Could not parse level '{level_raw}' at {current_freq / MHZ_TO_HZ:.3f} MHz")
            writer.writerow([current_freq / MHZ_TO_HZ, level_raw]) # Write raw string if parsing fails
        except pyvisa.VisaIOError as e:
            print(f"    VISA Error during data collection at {current_freq / MHZ_TO_HZ:.3f} MHz: {e}")
            writer.writerow([current_freq / MHZ_TO_HZ, "ERROR"])
            break # Exit segment sweep if instrument communication fails
        current_freq += step_size_hz

def run_scan_cycle(args):
    """
    Manages one full cycle of scanning, connecting to the instrument,
    performing sweeps across defined bands or a custom range, and saving data.

    Args:
        args (argparse.Namespace): Parsed command-line arguments.
    """
    file_prefix = args.SCANname
    step_hz = int(args.stepSize)
    user_running = args.user

    visa_address = VISA_ADDRESSES.get(user_running, VISA_ADDRESSES['zap']) # Default to 'zap'

    inst = initialize_instrument(visa_address)
    if inst is None:
        return # Skip scan if instrument connection failed

    try:
        # Determine the list of ranges to scan for the current cycle
        ranges_to_scan = []
        if args.startFreq is not None and args.endFreq is not None:
            # If custom range provided, use it
            ranges_to_scan.append({
                "Band Name": "Custom Scan",
                "Start MHz": args.startFreq / MHZ_TO_HZ,
                "Stop MHz": args.endFreq / MHZ_TO_HZ
            })
        else:
            # Otherwise, iterate through all predefined bands
            ranges_to_scan = BAND_RANGES

        # File & directory setup
        scan_dir = os.path.join(os.getcwd(), "N9340 Scans")
        os.makedirs(scan_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H-%M-%S")
        filename = os.path.join(scan_dir, f"{file_prefix}{timestamp}.csv")

        # Open the CSV file and perform sweeps
        with open(filename, mode='w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            # Write header row to CSV
            writer.writerow(["Frequency (MHz)", "Level (dBm)"])

            for band_info in ranges_to_scan:
                band_name = band_info["Band Name"]
                current_band_start_freq = int(band_info["Start MHz"] * MHZ_TO_HZ)
                current_band_stop_freq = int(band_info["Stop MHz"] * MHZ_TO_HZ)

                print(f"\nScanning Band: {band_name}")
                print(f"Full Band Range: {current_band_start_freq / MHZ_TO_HZ:.3f} MHz to {current_band_stop_freq / MHZ_TO_HZ:.3f} MHz")

                current_block_start = current_band_start_freq
                scan_limit = current_band_stop_freq

                # Loop through frequency blocks until the end frequency of the current band is reached.
                while current_block_start < scan_limit:
                    current_block_stop = min(current_block_start + DEFAULT_SEGMENT_WIDTH_HZ, scan_limit)
                    perform_segment_sweep(inst, current_block_start, current_block_stop, step_hz, band_name, writer)
                    current_block_start = current_block_stop # Move to the start of the next block.

        print(f"\nScan cycle complete. Results saved to '{filename}'")

    except pyvisa.VisaIOError as e:
        print(f"VISA Error during scan cycle: {e}")
    except Exception as e:
        print(f"An unexpected error occurred during the scan cycle: {e}")
    finally:
        # Cleanup: Close instrument connection
        if inst and inst.session != 0:
            try:
                inst.write("SYST:LOC") # Attempt to send instrument to local control
                inst.close()
                print("Instrument connection closed.")
            except pyvisa.VisaIOError:
                pass # Ignore if connection is already broken or command not supported

def wait_with_interrupt(wait_time_seconds):
    """
    Provides a timed delay with user interruption capability.
    Allows skipping the wait, quitting the program, or resuming.

    Args:
        wait_time_seconds (int): The total time to wait in seconds.
    """
    print("\n" + "="*50)
    print(f"Next full scan cycle in {wait_time_seconds // 60} minutes.")
    print("Press Ctrl+C at any time during the countdown to interact.")
    print("="*50)

    seconds_remaining = wait_time_seconds
    skip_wait = False

    while seconds_remaining > 0:
        minutes = seconds_remaining // 60
        seconds = seconds_remaining % 60
        sys.stdout.write(f"\rTime until next scan: {minutes:02d}:{seconds:02d} ")
        sys.stdout.flush()

        try:
            time.sleep(1)
        except KeyboardInterrupt:
            sys.stdout.write("\n")
            sys.stdout.flush()
            choice = input("Countdown interrupted. (S)kip wait, (Q)uit program, or (R)esume countdown? ").strip().lower()
            if choice == 's':
                skip_wait = True
                print("Skipping remaining wait time. Starting next scan shortly...")
                break
            elif choice == 'q':
                print("Exiting program.")
                sys.exit(0)
            else:
                print("Resuming countdown...")
        seconds_remaining -= 1

    if not skip_wait:
        sys.stdout.write("\r" + " "*50 + "\r") # Clear the last countdown line
        sys.stdout.flush()
        print("Starting next scan now!")

    print("\n" + "="*50 + "\n")

def main():
    """
    Main function to run the spectrum analyzer sweep program.
    """
    args = setup_arguments()

    # Determine if custom range or any arguments were provided to skip table display
    skip_table_display = (args.startFreq is not None and args.endFreq is not None) or \
                         (len(sys.argv) > 1 and not (len(sys.argv) == 2 and sys.argv[1] == os.path.basename(__file__)))
    # The condition `len(sys.argv) > 1 and not (len(sys.argv) == 2 and sys.argv[1] == os.path.basename(__file__))
    # is a more robust way to check if arguments were *actually* passed,
    # as sys.argv[0] is always the script name itself.

    if not skip_table_display:
        display_predefined_bands()

    while True:
        run_scan_cycle(args)
        wait_with_interrupt(DEFAULT_CYCLE_WAIT_TIME_SECONDS)

if __name__ == "__main__":
    main()
