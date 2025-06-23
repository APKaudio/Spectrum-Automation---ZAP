# Define frequency sweep range (in Hz)
$startFreq = 400e6
$endFreq   = 600e6
$stepFreq  = 1e6   # 1 MHz step — adjust for finer or coarser resolution

# VISA address for USB connection
$visaAddress = "USB0::0x0957::0xFFEF::CN03480580::0::INSTR"

# Load VISA COM object
$visa = New-Object -ComObject "IOActiveX.IOActiveXCtrl"
$visa.IO = $visa.IOManager.Open($visaAddress)

# Clear instrument and set up sweep
$visa.IO.Write("*CLS")
$visa.IO.Write(":FREQ:START $startFreq")
$visa.IO.Write(":FREQ:STOP $endFreq")
$visa.IO.Write(":FREQ:SPAN ($endFreq - $startFreq)")
$visa.IO.Write(":BAND 100kHz")  # Set resolution bandwidth if needed
$visa.IO.Write(":INIT:CONT OFF")  # Turn off continuous sweep
$visa.IO.Write(":INIT:IMM")       # Trigger single sweep
Start-Sleep -Seconds 2

# Loop over the frequency range using a marker
for ($freq = $startFreq; $freq -le $endFreq; $freq += $stepFreq) {
    # Move marker to target frequency
    $visa.IO.Write(":CALC:MARK1:X $freq")

    # Query amplitude at this frequency
    $visa.IO.Write(":CALC:MARK1:Y?")
    $level = $visa.IO.ReadString().Trim()

    # Output the frequency and level
    "{0,10} Hz : {1,8} dBm" -f [math]::Round($freq), $level
}

# Close the connection
$visa.IO.Close()
