"""
	Author	: Ayush Chinmay
	Created	: 22 November 2023
	Updated	: 27 November 2023
	
	Description: Sweep through frequencies, capture data, and generate the frequency response curve for any circuit

	Notes:
		- Currently only supports the Agilent InfiniiVision DSO-X 3012A Oscilloscope
		- Enter the start/stop frequencies and number of steps when prompted
		- If no input is given, the default values are used: [Start: 100Hz | Stop: 10kHz | Steps: 25]
		- The data is saved to a CSV file in the /Data folder
		- The Bode Plot is saved to an image file in the /Img folder
		- The data is plotted using matplotlib

	Software Setup:
		1) Download NI-VISA, NI-Device Drivers, and IVI Foundation Shared Components
		2) Install the required modules by running the `setupModule.bat`
		3) Run the Code > Enter start/stop frequencies (in Hz) > Enter the number of steps > wait...

	Hardware Setup:
		Components:
			- Agilent InfiniiVision DSO-X 3012A Oscilloscope
			- 3 X BNC-Banana Cables
		Setup:
			1) Attach a BNC-Banana to `Gen-Out` port on the oscilloscope
			2) Attach a BNC-Banana to `Channel-1` port; 
				Connect the Banana-Plug sides of Channel-1 with Gen-Out Banand Plugs
				These will be the INPUT for the circuit
			3) Attach a BNC-Banana to `Channel-2` port
				This will be the OUTPUT from the circuit
"""

# =====================[ Import Modules ]==================
# *********************************************************
import os
import sys
import math
import array
import string
from time import sleep
from time import localtime
import pyvisa
import matplotlib.pyplot as plt
from comtypes.client import GetModule
from comtypes.client import CreateObject
# *********************************************************

if not hasattr(sys, "frozen"):									# Run GetModule once to generate comtypes.gen.VisaComLib.
	GetModule("C:/Program Files (x86)/IVI Foundation/VISA/VisaCom/GlobMgr.dll")
import comtypes.gen.VisaComLib as VisaComLib

# ===========[ Global Variables ]=========== #
myScope = CreateObject("VISA.BasicFormattedIO", interface=VisaComLib.IFormattedIO488)	# Create a VISA object
dir_path = os.path.dirname(os.path.realpath(__file__))			# Get the current directory path


# ===========[ Send String Command ]=========== #
def do_command(command):
	"""
	Sends a command, checks for errors, and returns nothing
	@param command: Command to be executed
	"""
	myScope.WriteString("%s" % command, True)					# Send the command to the instrument.
	check_instrument_errors(command)							# Check for instrument errors.


# ===========[ Send Binary Block Command ]=========== #
def do_command_ieee_block(command, data):
	"""
	Sends a command, checks for errors, and returns a string
	@param command: Command to be executed
	@param data: Data to be sent
	"""
	myScope.WriteIEEEBlock(command, data, True)					# Send the command to the instrument.
	check_instrument_errors(command)							# Check for instrument errors.


# ===========[ Query String ]=========== #
def do_query_string(query):
	"""
	Sends a query, checks for errors, and returns a string
	@param query: Query to be executed
	"""
	myScope.WriteString("%s" % query, True)						# Send the command to the instrument.
	result = myScope.ReadString()								# Read the string from the instrument.
	check_instrument_errors(query)								# Check for instrument errors.
	return result


# ===========[ Query Binary Block ]=========== #
def do_query_ieee_block(query):
	"""
	Sends a query, checks for errors, and returns a string
	@param query: Query to be executed
	"""
	myScope.WriteString("%s" % query, True)						# Send the command to the instrument.
	result = myScope.ReadIEEEBlock(VisaComLib.BinaryType_UI1, False, True)	# Read the binary block from the instrument.
	check_instrument_errors(query)								# Check for instrument errors.
	return result

# ===========[ Query Number ]=========== #
def do_query_number(query):
	"""
	Sends a query, checks for errors, and returns a value
	@param query: Query to be executed
	"""
	myScope.WriteString("%s" % query, True)						# Send the command to the instrument.
	result = myScope.ReadNumber(VisaComLib.ASCIIType_R8, True)	# Read the number from the instrument.
	check_instrument_errors(query)								# Check for instrument errors.
	return result


# ===========[ Query List of Numbers ]=========== #
def do_query_numbers(query):
	"""
	Sends a query, checks for errors, and returns a list of values
	@param query: Query to be executed
	"""
	myScope.WriteString("%s" % query, True)						# Send the command to the instrument.
	result = myScope.ReadList(VisaComLib.ASCIIType_R8, ",;")	# Read the list of numbers from the instrument.
	check_instrument_errors(query)								# Check for instrument errors.
	return result


# ===========[ Check Instrument Error ]=========== #
def check_instrument_errors(command):
	"""
	Checks for instrument errors
	If there is an error, print the error and exits the program
	@param command: Command to be executed
	"""
	while True:	# Loop until no more errors are found.
		myScope.WriteString(":SYSTem:ERRor?", True)	# Check for instrument errors.
		error_string = myScope.ReadString()			# Read the error string from the instrument.
		if error_string:	# If there is an error string value.
			if error_string.find("+0,", 0, 3) == -1:# Not "No error".
				print("ERROR: %s, command: '%s'" % (error_string, command))
				print("Exited because of error.")
				sys.exit(1)		# Exit the program with an error.
			else:	# "No error"
				break
		else: # :SYSTem:ERRor? should always return string.
			print("ERROR: :SYSTem:ERRor? returned nothing, command: '%s'" % command)
			print("Exited because of error.")
			sys.exit(1)			# Exit the program with an error.


# ===========[ Save Scope Setup ]=========== #
def save_setup():
	"""
	Saves the current scope setup to a file that can be loaded later
	"""
	setup_bytes = do_query_ieee_block(":SYSTem:SETup?")
	nLength = len(setup_bytes)			# Get the length of the setup bytes
	f = open("/Data/setup.stp", "wb")	# Open the file in write-binary mode
	f.write(bytearray(setup_bytes))		# Write the setup bytes to the file
	f.close()
	print("Setup bytes saved: %d" % nLength)	# Print the number of bytes saved


# ===========[ Load Default Setup ]=========== #
def default_setup():
	"""
	Loads the default scope setup
	"""
	do_command("*CLS")	# Clear the instrument status
	do_command("*RST")	# Reset the instrument to default settings


# ===========[ Autoscale ]=========== #
def autoscale():
	"""
	Autoscale the scope display
	"""
	# print(">> [INFO] AUTOSCALE\n")
	do_command(":AUToscale")


# ===========[ Get Port ID ]=========== #
def get_deviceID():
	# Returns the string containing VISA ID for the Oscilloscope
	rm = pyvisa.ResourceManager()
	scopeDevice = rm.list_resources()[0]
	print(f"DEVICE ID: {scopeDevice}")
	return scopeDevice


# ===========[ Initialize oscilloscope ]=========== #
def initialize():
	global myScope
	rm = CreateObject("VISA.GlobalRM", interface=VisaComLib.IResourceManager)
	# myScope.IO = rm.Open("USB0::0x0957::0x17A9::MY51250110::0::INSTR")
	myScope.IO = rm.Open(get_deviceID())
	# Clear the interface.
	myScope.IO.Clear()
	print("Interface cleared.")
	# Set the Timeout to 15 seconds.
	myScope.IO.Timeout = 15000 # 15 seconds.
	print("Timeout set to 15000 milliseconds.")

	# Get and display the device's *IDN? string.
	idn_string = do_query_string("*IDN?")
	print("Identification string '%s'" % idn_string)
	# # Clear status and load the default setup.
	default_setup()
	sleep(1)


# ===========[ Capture Waveform ]=========== #
def capture_waveform():
	"""
	Digitizes the scope display to acquire waveform measurements, and returns them as a tuple
	@return: Tuple containing Frequency, Amplitude, Phase Difference for channel-1 and channel-2
	"""
	# Start capturing the waveform
	do_command(":SINGLE")								# Capture a single waveform using :SINGLE.
	do_command(":ACQuire:TYPE NORMal")					# Set the acquisition type to Normal
	qresult = do_query_string(":ACQuire:TYPE?")			# Query the acquisition type
	do_command(f":DIGitize CHANnel1,CHANnel2")			# Capture an acquisition using :DIGitize.
	# print(f"Acquire type: {qresult}")

	# Make measurements on the waveform
	do_command(f":MEASure:SOURce CHANnel1,CHANnel2")	# Set active sources for measurements
	qresult = do_query_string(":MEASure:SOURce?")		# Query the active sources for measurements
	print(f"MEASURE SOURCE: {qresult}{40*'='}")

	# Measure Frequency on Channel 1
	do_command(":MEASure:FREQuency CHANnel1")			# Set the measurement type to Frequency on Channel 1
	qresult = do_query_string(":MEASure:FREQuency? CHANnel1")	# Query the measurement value on Channel 1
	meas_freq1 = float(qresult)							# Convert the string to float
	print(f"FREQUENCY   [CHAN 1]:  {meas_freq1:<5.3f} {'Hz':<4s}")

	# Measure Peak-Peak Voltage on Channel 1
	do_command(":MEASure:VPP CHANnel1")					# Set the measurement type to Peak-Peak Voltage	 on Channel 1
	qresult = do_query_string(":MEASure:VPP? CHANnel1")	# Query the measurement value on Channel 1
	meas_ampl1 = float(qresult)							# Convert the string to float
	print(f"AMPLITUDE   [CHAN 1]:  {meas_ampl1:<5.3f} {'Vpp':<4s}")

	# Measure Frequency on Channel 2
	do_command(":MEASure:FREQuency CHANnel2")			# Set the measurement type to Frequency on Channel 2
	qresult = do_query_string(":MEASure:FREQuency? CHANnel2")	# Query the measurement value on Channel 2
	meas_freq2 = float(qresult)							# Convert the string to float
	print(f"FREQUENCY   [CHAN 2]:  {meas_freq2:<5.3f} {'Hz':<4s}")

	# Measure Peak-Peak Voltage on Channel 2
	do_command(":MEASure:VPP CHANnel2")					# Set the measurement type to Peak-Peak Voltage	 on Channel 2
	qresult = do_query_string(":MEASure:VPP? CHANnel2")	# Query the measurement value on Channel 2
	meas_ampl2 = float(qresult)							# Convert the string to float
	print(f"AMPLITUDE   [CHAN 2]:  {meas_ampl2:<5.3f} {'Vpp':<4s}")

	# Measure Phase shift between Channel 1 and 2
	do_command(":MEASure:PHASe CHANnel2,CHANnel1")		# Set the measurement type to Phase shift between Channel 2 and 1
	qresult = do_query_string(":MEASure:PHASe? CHANnel2,CHANnel1")	# Query the measurement value between Channel 2 and 1
	phase_diff = float(qresult)							# Convert the string to float
	print(f"PHASE SHIFT [CH 2->1]: {phase_diff:<5.3f} {'deg':<4s}\n")

	return (meas_freq1, meas_ampl1, meas_freq2, meas_ampl2, phase_diff)


# ===========[ Setup Wave Generator ]=========== #
def generate_waveform(freq, ampl=1.0, offset=0.0):
	"""
	Function to generate a sinusoidal waveform using the built-in function generator
	@param freq: Frequency in Hertz
	@param ampl: Amplitude in Volts
	@param offset: DC Offset in Volts
	"""
	do_command(":WGEN:OUTput 1");						# Enable the output 
	do_command(":WGEN:FUNCtion SIN");					# Set the waveform to Sine
	do_command(f":WGEN:VOLTage {ampl:.3f}");			# Set the amplitude (Vpp)
	do_command(f":WGEN:FREQuency {freq:.3f}");			# Set the frequency	(Hz)
	do_command(f":WGEN:VOLTage:OFFset {offset:.3f}")	# Set the DC Offset	(V)
	return (freq, ampl, offset)


# ===========[ Sweep Function ]=========== #
def sweep_frequency(start_freq, end_freq, points):
	"""
	Sweep through the frequencies and generate a sinusoidal waveform using the built-in function generator
	@param start_freq: Starting Frequency (Hz)
	@param end_freq: End Frequency (Hz)
	@param points: Number of points to sweep through
	"""
	# Calcualte the time-step
	counter = 1
	error_cnt = 0
	step = (end_freq - start_freq) / (points-1)
	curr_freq = start_freq

	# fn = "bode_measure.csv"
	curr_time = localtime()
	fname = f"Bode_{curr_time[1]}-{curr_time[2]}-{curr_time[0]}_{curr_time[3]}-{curr_time[4]}-{curr_time[5]}"
	fp = open(f"{dir_path}/Data/{fname}.csv", 'w')
	fp.write("CH1_FREQ [Hz], CH1_AMPL [Vpp], CH2_FREQ [Hz], CH2_AMPL [Vpp], PHASE_DIFF [Deg], GAIN [dB]\n")

	# Generate the waveform at the starting frequency
	generate_waveform(curr_freq)
	sleep(3)				# Wait for the waveform to settle
	autoscale()				# Autoscale the scope to get a good view of the waveform

	# Capture the initial Waveform
	freq1, ampl1, freq2, ampl2, phase_diff = capture_waveform()

	# Sweep through the frequencies and generate the waveform
	while (curr_freq <= end_freq):
		print(f"\nSTEP [{counter}/{points:d}] : {curr_freq:5.2f} Hz")
		generate_waveform(curr_freq)	# Generate the waveform at the current frequency
		autoscale()						# Autoscale the scope to get a good view of the waveform
		# Capture Waveform and get the measurements
		freq1, ampl1, freq2, ampl2, phase_diff = capture_waveform()

		# Check for errors
		if (freq2 > 1.5*freq1) or (freq1 > 10E9) or (ampl1 > 10E9):
			error_cnt += 1
			curr_freq += 2
			print("\n>> [ERROR] Erroneous Signal Detected!")
			print(f">> [INFO] Retrying at {curr_freq:.0f} Hz...")
			if (error_cnt > 4):
				print(f"\n>> [ERROR] Unable to read signal at {curr_freq:5.2f} Hz!")
				print(">> [ERROR] Please modify frequency range and try again...")
				print(f">> [INFO] Suggested Range: 100Hz - {curr_freq-step:.0f}Hz")
				break
			continue
		# Calculate the gain and write to file
		gain = 20*math.log10(float(ampl2)/float(ampl1))
		fp.write(f"{freq1}, {ampl1}, {freq2}, {ampl2}, {phase_diff}, {gain}\n")
		curr_freq += step
		counter += 1
		sleep(0.75)
	fp.close()
	# Save the data to a file and print the path
	print(f"Measurement Data saved to path: /Data/{fname}.csv")
	plot_bode(fname)


# ===========[ Sweep Function ]=========== #
def plot_bode(fn):
	"""
	Parse the data file and plot the Frequency/Phase Response
	Display the plot and save it to an image file
	@param fn: file name of the data file
	"""
	fp = open(f"{dir_path}/Data/{fn}.csv", 'r')
	freq_arr  = []		# Array to store the Frequency: CH1 [Hz]
	phase_arr = []		# Array to store the Phase Difference: Phase(CH2)-Phase(CH1) [Deg]
	gain_arr  = []		# Array to store the Gain: 20*log10(V2/V1) [dB]

	# Parse the data file
	for lin in fp:
		try:
			f1, a1, f2, a2, pd, gain = lin.strip("\n").split(", ")
			freq_arr.append(float(f1))
			phase_arr.append(float(pd))
			gain_arr.append(float(gain))
			# print(f"freq: {float(f1)} | gain: {20*math.log10(float(a2)/float(a1))}")
		except:
			continue

	# Plot the Bode Plot
	plt.figure(figsize=(9, 7), layout='constrained', num="Frequency Response Plots", facecolor='#f0f0f0')
	plt.style.use('seaborn-v0_8-bright')

	# Plot the Gain Response
	plt.subplot(211)
	plt.grid(True, which='both', linestyle='--', linewidth=0.5)
	plt.minorticks_on()
	plt.plot(freq_arr, gain_arr, color='#00798C')
	plt.title("Frequency Response", family='sans', fontweight='bold', color='#00798C')
	plt.xlabel("Frequency (Hz)", family='sans', style='italic')
	plt.ylabel("Gain (dB)", family='sans', style='italic')
	plt.xscale('log', base=10)

	# Plot the Phase Response
	plt.subplot(212)
	plt.grid(True, which='both', linestyle='--', linewidth=0.5)
	plt.minorticks_on()
	plt.plot(freq_arr, phase_arr, color='#D1495B')
	plt.title("Phase Response", family='sans', fontweight='bold', color='#D1495B')
	plt.xlabel("Frequency (Hz)", family='sans', style='italic')
	plt.ylabel("Phase (Deg)", family='sans', style='italic')
	plt.xscale('log', base=10)

	# Save the plot and display
	plt.savefig(f"{dir_path}/Img/{fn}.png", bbox_inches='tight')
	print(f"Bode Plot Image saved to path: /Img/{fn}.png")
	plt.show()


# ===========[ Main Code ]=========== #
def main():
	"""
	Initialize the oscilloscope
	Take user input for start/stop frequencies and number of steps
	Sweep through the frequencies and generate the waveform
	Capture the waveform and generate the frequency response curve
	"""
	initialize()
	try:	# Take user input for start/stop frequencies and number of steps
		fstart = float(input(">> Enter start frequency [Hz]: "))
		fstop = float(input(">> Enter stop frequency [Hz]: "))
		steps = int(input(">> Enter the desired number of steps: "))
	except:	# If no/invalid input is given, use the default values
		fstart = 100.0
		fstop = 10000.0
		steps = 25
		print(">> [ERROR] Invalid Inputs! Using default values...")
		print(f">> Start: {fstart:5.2f} Hz | Stop: {fstop:5.2f} Hz | Steps: {steps:d}\n")
	sweep_frequency(fstart, fstop, steps)	# Sweep through the frequencies and generate the waveform
	sys.exit()								# Exit the program

if __name__ == '__main__':
	main()