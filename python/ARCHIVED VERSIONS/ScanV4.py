import pyvisa
import time
import csv
from datetime import datetime
import os
import argparse
import sys

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

# ------------------------------------------------------------------------------
# Command-line argument parsing
# This section defines and parses command-line arguments, allowing users to
# customize the scan parameters (filename, frequency range, step size) when
# running the script.
# ------------------------------------------------------------------------------
parser = argparse.ArgumentParser(description="Spectrum Analyzer Sweep and CSV Export")

# Define an argument for the prefix of the output CSV filename
parser.add_argument('--SCANname', type=str, default="25kz scan ",
                    help='Prefix for the output CSV filename')

# Define optional arguments for custom start and end frequencies.
# If both are provided, they will override the predefined band scanning.
parser.add_argument('--startFreq', type=float, default=None,
                    help='Start frequency in Hz (overrides default bands if provided with --endFreq)')
parser.add_argument('--endFreq', type=float, default=None,
                    help='End frequency in Hz (overrides default bands if provided with --startFreq)')

# Define an argument for the step size
parser.add_argument('--stepSize', type=float, default=25000,
                    help='Step size in Hz')
                    
# Add an argument to choose who is running the program (apk or zap)
parser.add_argument('--user', type=str, choices=['apk', 'zap'], default='zap',
                    help='Specify who is running the program: "apk" or "zap". Default is "zap".')

# Parse the arguments provided by the user
args = parser.parse_args()

# Assign parsed arguments to variables for easy access
file_prefix = args.SCANname
step = args.stepSize
user_running = args.user

# Determine if a custom range is specified via command-line arguments
scan_custom_range = False
custom_scan_details = None

if args.startFreq is not None and args.endFreq is not None:
    # If both startFreq and endFreq are provided, perform a single custom scan.
    scan_custom_range = True
    custom_scan_details = {
        "Band Name": "Custom Scan",
        "Start MHz": args.startFreq / 1e6, # Store in MHz for consistency, convert to Hz later
        "Stop MHz": args.endFreq / 1e6
    }
    # If a custom range is explicitly given, skip the display of predefined bands.
    skip_table_display = True
elif len(sys.argv) > 1:
    # If any command-line argument is provided (even if not startFreq/endFreq),
    # skip the table display as per user's request.
    skip_table_display = True
else:
    # If no command-line arguments are provided, display the table.
    skip_table_display = False

# Display the table of scan areas if no arguments were added
if not skip_table_display:
    print("Predefined Scan Areas:")
    print("=" * 40)
    print(f"{'Band Name':<15} {'Start MHz':<12} {'Stop MHz':<12}")
    print("-" * 40)
    for band in BAND_RANGES:
        print(f"{band['Band Name']:<15} {band['Start MHz']:<12.3f} {band['Stop MHz']:<12.3f}")
    print("=" * 40 + "\n")

# Define the waiting time in seconds between full scan cycles
WAIT_TIME_SECONDS = 300 # 5 minutes

# ------------------------------------------------------------------------------
# Main program loop
# The entire scanning process will now run continuously with a delay.
# ------------------------------------------------------------------------------
while True:
    # Determine the list of ranges to scan for the current cycle
    if scan_custom_range:
        # If a custom range was provided, only scan that range
        ranges_to_scan = [custom_scan_details]
    else:
        # Otherwise, iterate through all predefined bands
        ranges_to_scan = BAND_RANGES

    # --------------------------------------------------------------------------
    # VISA connection setup
    # This section establishes communication with the spectrum analyzer using the
    # PyVISA library, opens the specified instrument resource, and performs initial
    # configuration commands.
    # --------------------------------------------------------------------------
    # Define the VISA addresses for different users.
    apk_visa_address = 'USB0::0x0957::0xFFEF::CN03480580::0::INSTR'
    zap_visa_address = 'USB1::0x0957::0xFFEF::SG05300002::0::INSTR'
    
    # Select the VISA address based on the --user argument
    if user_running == 'apk':
        visa_address = apk_visa_address
    else:  # default is 'zap'
        visa_address = zap_visa_address

    # Create a ResourceManager object, which is the entry point for PyVISA.
    rm = pyvisa.ResourceManager()
    inst = None # Initialize inst to None to ensure it's defined before finally block

    try:
        # Open the connection to the specified instrument resource.
        inst = rm.open_resource(visa_address)
        print(f"Connected to instrument at {visa_address}")

        # Clear the instrument's status byte and error queue.
        inst.write("*CLS")
        # Reset the instrument to its default settings.
        inst.write("*RST")
        # Query the Operation Complete (OPC) bit to ensure the previous commands have
        # finished executing before proceeding. This is important for synchronization.
        inst.query("*OPC?")

        # Configure the preamplifier for high sensitivity.
        inst.write(":POWer:GAIN ON")
        print("Preamplifier turned ON.")
        inst.write(":POWer:GAIN 1") # '1' is equivalent to 'ON'
        print("Preamplifier turned ON for high sensitivity.")

        # Configure the display: Set Y-axis scale to logarithmic (dBm).
        inst.write(":DISP:WIND:TRAC:Y:SCAL LOG")
        # Configure the display: Set the reference level for the Y-axis.
        inst.write(":DISP:WIND:TRAC:Y:RLEV -30")
        # Enable Marker 1. Markers are used to read values at specific frequencies.
        inst.write(":CALC:MARK1 ON")
        # Set Marker 1 mode to position, meaning it can be moved to a specific frequency.
        inst.write(":CALC:MARK1:MODE POS")
        # Activate Marker 1, making it ready for use.
        inst.write(":CALC:MARK1:ACT")

        # Set the instrument to single sweep mode.
        # This ensures that after each :INIT:IMM command, the instrument performs one
        # sweep and then holds the trace data until another sweep is initiated.
        inst.write(":INITiate:CONTinuous OFF")

        # Pause execution for 2 seconds to allow the instrument to settle after configuration.
        time.sleep(2)

        # --------------------------------------------------------------------------
        # File & directory setup
        # This section prepares the output directory and generates a unique filename
        # for the CSV export based on the current timestamp and user-defined prefix.
        # --------------------------------------------------------------------------
        # Define the directory where scan results will be saved.
        scan_dir = os.path.join(os.getcwd(), "N9340 Scans")
        # Create the directory if it doesn't already exist.
        os.makedirs(scan_dir, exist_ok=True)

        # Generate a timestamp for the filename to ensure uniqueness.
        timestamp = datetime.now().strftime("%Y%m%d_%H-%M-%S")
        # Construct the full path for the output CSV file.
        filename = os.path.join(scan_dir, f"{file_prefix}--{timestamp}.csv")

        # --------------------------------------------------------------------------
        # Sweep and write to CSV for each defined band/range
        # This is the core logic, performing the frequency sweep in segments,
        # reading data, and writing it to the CSV.
        # --------------------------------------------------------------------------
        # Open the CSV file in write mode (`'w'`). `newline=''` prevents extra blank rows.
        with open(filename, mode='w', newline='') as csvfile:
            # Create a CSV writer object.
            writer = csv.writer(csvfile)
            
            # Iterate through each defined band or the custom range
            for band_info in ranges_to_scan:
                band_name = band_info["Band Name"]
                # Convert start and stop frequencies from MHz to Hz for instrument commands
                current_band_start_freq = int(band_info["Start MHz"] * 1e6)
                current_band_stop_freq = int(band_info["Stop MHz"] * 1e6)

                # Display the name of the band being scanned
                print(f"\nScanning Band: {band_name}")
                print(f"Full Band Range: {current_band_start_freq / 1e6:.3f} MHz to {current_band_stop_freq / 1e6:.3f} MHz")

                # Define the width of each frequency segment for sweeping.
                segment_width = 10_000_000  # 10 MHz
                # Convert step size to integer, as some instrument commands might expect integers.
                step_int = int(step)
                
                # Initialize the start of the current frequency block.
                current_block_start = current_band_start_freq
                scan_limit = current_band_stop_freq

                # Loop through frequency blocks until the end frequency of the current band is reached.
                while current_block_start < scan_limit:
                    # Calculate the end frequency for the current block.
                    current_block_stop = current_block_start + segment_width
                    # Ensure the block stop doesn't exceed the overall scan limit for the band.
                    if current_block_stop > scan_limit:
                        current_block_stop = scan_limit

                    # Print the current sweep segment range to the console for user feedback.
                    print(f"  Sweeping segment: {current_block_start / 1e6:.3f} MHz to {current_block_stop / 1e6:.3f} MHz")

                    # Set the start and stop frequencies for the instrument's sweep for this segment.
                    inst.write(f":FREQ:START {current_block_start}")
                    inst.write(f":FREQ:STOP {current_block_stop}")
                    # Initiate a single immediate sweep.
                    inst.write(":INIT:IMM")
                    # Query Operation Complete to ensure the sweep has finished before reading markers.
                    inst.query("*OPC?")

                    # Initialize the current frequency for data point collection within the segment.
                    current_freq = current_block_start
                    # Loop through each frequency step within the current segment.
                    while current_freq <= current_block_stop:
                        # Set Marker 1 to the current frequency.
                        inst.write(f":CALC:MARK1:X {current_freq}")
                        # Query the Y-axis value (level in dBm) at Marker 1's position.
                        level_raw = inst.query(":CALC:MARK1:Y?").strip()

                        try:
                            # Attempt to convert the raw level string to a float.
                            level = float(level_raw)
                            # Format the level to one decimal place for consistent output.
                            level_formatted = f"{level:.1f}"
                            # Convert frequency from Hz to MHz for readability in console and CSV.
                            freq_mhz = current_freq / 1_000_000
                            #the frequency, band name, and level to the console.
                            print(f"    Band: {band_name:<15} Freq: {freq_mhz:.3f} MHz : {level_formatted} dBm")
                            # Write the frequency and formatted level to the CSV file.
                            # The CSV only includes frequency and level, not the band name.
                            writer.writerow([freq_mhz, level_formatted])

                        except ValueError:
                            # If the raw level cannot be converted (e.g., error message from instrument)
                            level_formatted = level_raw
                            print(f"    Warning: Could not parse level '{level_raw}' at {current_freq / 1e6:.3f} MHz")
                            writer.writerow([current_freq / 1_000_000, level_formatted])

                        # Increment the current frequency by the step size.
                        current_freq += step_int
                    
                    # Move to the start of the next block.
                    current_block_start = current_block_stop

    except pyvisa.VisaIOError as e:
        print(f"VISA Error: Could not connect to or communicate with the instrument: {e}")
        print("Please ensure the instrument is connected and the VISA address is correct.")
        # Decide if you want to exit or retry after a connection error.
        # For now, it will proceed to the wait and then try again.
    except Exception as e:
        print(f"An unexpected error occurred during the scan: {e}")
        # Continue to the wait or exit if the error is critical.
    finally:
        # ----------------------------------------------------------------------
        # Cleanup
        # This section ensures that the instrument is returned to a safe state and
        # the VISA connection is properly closed after the scan is complete.
        # ----------------------------------------------------------------------
        # Check if the instrument object exists and the session is still open.
        if inst and inst.session != 0:
            try:
                # Attempt to send the instrument to local control.
                inst.write("SYST:LOC")
            except pyvisa.VisaIOError:
                # Ignore if command is not supported or connection is already broken.
                pass
            finally:
                # Close the instrument connection.
                inst.close()
                print("Instrument connection closed.")
        
        # Print a confirmation message indicating the scan completion and output file.
        # Only print if filename was successfully created (i.e., not an early error before file setup).
        if 'filename' in locals():
            print(f"\nScan cycle complete. Results saved to '{filename}'")

    # --------------------------------------------------------------------------
    # Countdown and Interruptible Wait
    # This section provides a timed delay between scan cycles, allowing for
    # user interruption to skip the wait or quit the program.
    # --------------------------------------------------------------------------
    print("\n" + "="*50)
    print(f"Next full scan cycle in {WAIT_TIME_SECONDS // 60} minutes.")
    print("Press Ctrl+C at any time during the countdown to interact.")
    print("="*50)

    seconds_remaining = WAIT_TIME_SECONDS
    skip_wait = False

    while seconds_remaining > 0:
        minutes = seconds_remaining // 60
        seconds = seconds_remaining % 60
        # Print countdown, overwriting the same line for a dynamic display.
        sys.stdout.write(f"\rTime until next scan: {minutes:02d}:{seconds:02d} ")
        sys.stdout.flush() # Ensure the output is immediately written to the console.

        try:
            time.sleep(1) # Wait for 1 second
        except KeyboardInterrupt:
            # Handle Ctrl+C interruption
            sys.stdout.write("\n") # Move to a new line after Ctrl+C
            sys.stdout.flush()
            choice = input("Countdown interrupted. (S)kip wait, (Q)uit program, or (R)esume countdown? ").strip().lower()
            if choice == 's':
                skip_wait = True
                print("Skipping remaining wait time. Starting next scan shortly...")
                break # Exit the countdown loop
            elif choice == 'q':
                print("Exiting program.")
                sys.exit(0) # Exit the entire script
            else:
                print("Resuming countdown...")
                # Continue the loop from where it left off
        seconds_remaining -= 1

    if not skip_wait:
        # Clear the last countdown line before printing the next message
        sys.stdout.write("\r" + " "*50 + "\r")
        sys.stdout.flush()
        print("Starting next scan now!")
    
    print("\n" + "="*50 + "\n") # Add some spacing for clarity between cycles
