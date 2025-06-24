import pyvisa
import time
import csv
from datetime import datetime
import os
import argparse
import sys

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

# Define an argument for the start frequency
parser.add_argument('--startFreq', type=float, default=400e6,
                    help='Start frequency in Hz')

# Define an argument for the end frequency
parser.add_argument('--endFreq', type=float, default=650e6,
                    help='End frequency in Hz')

# Define an argument for the step size
parser.add_argument('--stepSize', type=float, default=25000,
                    help='Step size in Hz')

# Parse the arguments provided by the user
args = parser.parse_args()

# Assign parsed arguments to variables for easy access
file_prefix = args.SCANname
start_freq = args.startFreq
end_freq = args.endFreq
step = args.stepSize

# Define the waiting time in seconds
WAIT_TIME_SECONDS = 300 # 5 minutes

# ------------------------------------------------------------------------------
# Main program loop
# The entire scanning process will now run continuously with a delay.
# ------------------------------------------------------------------------------
while True:
    # --------------------------------------------------------------------------
    # VISA connection setup
    # This section establishes communication with the spectrum analyzer using the
    # PyVISA library, opens the specified instrument resource, and performs initial
    # configuration commands.
    # --------------------------------------------------------------------------
    # Define the VISA address of the spectrum analyzer. This typically identifies
    # the instrument on the bus (e.g., USB, LAN, GPIB).
    visa_address = 'USB0::0x0957::0xFFEF::CN03480580::0::INSTR'

    # Create a ResourceManager object, which is the entry point for PyVISA.
    rm = pyvisa.ResourceManager()

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
        # It creates a subdirectory named "N9340 Scans" in the current working directory.
        scan_dir = os.path.join(os.getcwd(), "N9340 Scans")
        # Create the directory if it doesn't already exist. `exist_ok=True` prevents
        # an error if the directory already exists.
        os.makedirs(scan_dir, exist_ok=True)

        # Generate a timestamp for the filename to ensure uniqueness.
        timestamp = datetime.now().strftime("%Y%m%d_%H-%M-%S")
        # Construct the full path for the output CSV file.
        filename = os.path.join(scan_dir, f"{file_prefix}--{timestamp}.csv")

        # --------------------------------------------------------------------------
        # Sweep and write to CSV
        # This is the core logic of the script, performing the frequency sweep in
        # segments, reading data from the spectrum analyzer, and writing it to the CSV.
        # --------------------------------------------------------------------------
        # Define the width of each frequency segment for sweeping.
        # Sweeping in segments helps manage memory and performance on some instruments.
        segment_width = 10_000_000  # 10 MHz

        # Convert step size to integer, as some instrument commands might expect integers.
        step_int = int(step)
        # Convert end frequency to integer, for consistent comparison in loops.
        scan_limit = int(end_freq)

        # Open the CSV file in write mode (`'w'`). `newline=''` prevents extra blank rows.
        with open(filename, mode='w', newline='') as csvfile:
            # Create a CSV writer object.
            writer = csv.writer(csvfile)
            # Initialize the start of the current frequency block.
            current_block_start = int(start_freq)

            # Loop through frequency blocks until the end frequency is reached.
            while current_block_start < scan_limit:
                # Calculate the end frequency for the current block.
                current_block_stop = current_block_start + segment_width
                # Ensure the block stop doesn't exceed the overall scan limit.
                if current_block_stop > scan_limit:
                    current_block_stop = scan_limit

                # Print the current sweep range to the console for user feedback.
                print(f"Sweeping range {current_block_start / 1e6:.3f} to {current_block_stop / 1e6:.3f} MHz")

                # Set the start frequency for the instrument's sweep.
                inst.write(f":FREQ:START {current_block_start}")
                # Set the stop frequency for the instrument's sweep.
                inst.write(f":FREQ:STOP {current_block_stop}")
                # Initiate a single immediate sweep.
                inst.write(":INIT:IMM")
                # Query Operation Complete to ensure the sweep has finished before reading markers.
                # This replaces the fixed time.sleep(2) for more robust synchronization.
                inst.query("*OPC?")

                # Initialize the current frequency for data point collection within the block.
                current_freq = current_block_start
                # Loop through each frequency step within the current block.
                while current_freq <= current_block_stop:
                    # Set Marker 1 to the current frequency.
                    inst.write(f":CALC:MARK1:X {current_freq}")
                    # Query the Y-axis value (level in dBm) at Marker 1's position.
                    # .strip() removes any leading/trailing whitespace or newline characters.
                    level_raw = inst.query(":CALC:MARK1:Y?").strip()

                    try:
                        # Attempt to convert the raw level string to a float.
                        level = float(level_raw)
                        # Format the level to one decimal place for consistent output.
                        level_formatted = f"{level:.1f}"
                        # Convert frequency from Hz to MHz for readability.
                        freq_mhz = current_freq / 1_000_000
                        # Print the frequency and level to the console.
                        print(f"{freq_mhz:.3f} MHz : {level_formatted} dBm")
                        # Write the frequency and formatted level to the CSV file.
                        writer.writerow([freq_mhz, level_formatted])

                    except ValueError:
                        # If the raw level cannot be converted to a float (e.g., if it's an error message),
                        # use the raw string directly.
                        level_formatted = level_raw
                        # Optionally, you might want to log this error or write a placeholder.
                        print(f"Warning: Could not parse level '{level_raw}' at {current_freq / 1e6:.3f} MHz")
                        writer.writerow([current_freq / 1_000_000, level_formatted])

                    # Increment the current frequency by the step size.
                    current_freq += step_int

                # Move to the start of the next block.
                current_block_start = current_block_stop

    except pyvisa.VisaIOError as e:
        print(f"VISA Error: Could not connect to or communicate with the instrument: {e}")
        print("Please ensure the instrument is connected and the VISA address is correct.")
        # Decide if you want to exit or retry after a connection error
        # For now, it will proceed to the wait and then try again.
    except Exception as e:
        print(f"An unexpected error occurred during the scan: {e}")
        # Continue to the wait or exit if the error is critical
    finally:
        # ----------------------------------------------------------------------
        # Cleanup
        # This section ensures that the instrument is returned to a safe state and
        # the VISA connection is properly closed after the scan is complete.
        # ----------------------------------------------------------------------
        if 'inst' in locals() and inst.session != 0: # Check if inst object exists and is not closed
            try:
                # Attempt to send the instrument to local control.
                inst.write("SYST:LOC")
            except pyvisa.VisaIOError:
                pass # Ignore if command is not supported or connection is already broken
            finally:
                inst.close()
                print("Instrument connection closed.")
        
        # Print a confirmation message indicating the scan completion and output file.
        if 'filename' in locals(): # Only print if filename was successfully created
            print(f"\nScan complete. Results saved to '{filename}'")

    # --------------------------------------------------------------------------
    # Countdown and Interruptible Wait
    # --------------------------------------------------------------------------
    print("\n" + "="*50)
    print(f"Next scan in {WAIT_TIME_SECONDS // 60} minutes.")
    print("Press Ctrl+C at any time during the countdown to interact.")
    print("="*50)

    seconds_remaining = WAIT_TIME_SECONDS
    skip_wait = False

    while seconds_remaining > 0:
        minutes = seconds_remaining // 60
        seconds = seconds_remaining % 60
        # Print countdown, overwriting the same line
        sys.stdout.write(f"\rTime until next scan: {minutes:02d}:{seconds:02d} ")
        sys.stdout.flush() # Ensure the output is immediately written to the console

        try:
            time.sleep(1)
        except KeyboardInterrupt:
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
        # Clear the last countdown line
        sys.stdout.write("\r" + " "*50 + "\r")
        sys.stdout.flush()
        print("Starting next scan now!")
    
    print("\n" + "="*50 + "\n") # Add some spacing for clarity between cycles

