"""
	Author	: Ayush Chinmay
	Created	: 29 November 2023
	Updated	: 29 November 2023
	
	Description: Use the plotly package to plot an interactive Bode Plot for the given data file

	Notes:
		- Requires the plotly package to be installed, if not installed, it will be installed automatically
		- Data files in the Data folder will be automatically detected and listed
		- Select the file number to plot the Bode Plot
		- Interactive bode plot will be displayed in the browser

	Running the Program:
		1) Run the Python program
		2) Enter a number from the prompt to select the data file to plot
"""

import os
import pip

# Import the required packages, if not installed, install them
try:	
	import plotly.graph_objects as go
	import plotly.subplots as sp
except ImportError:
	pip.main(['install', 'plotly']) 
	import plotly.graph_objects as go
	import plotly.subplots as sp


# ===========[ Global Variables ]=========== #
dir_path = os.path.dirname(os.path.realpath(__file__))			# Get the current directory path

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
			f1, a1, f2, a2, pd, gain = map(float, lin.strip("\n").split(", "))
			freq_arr.append(f1)
			phase_arr.append(pd)
			gain_arr.append(gain)
			# print(f"freq: {float(f1)} | gain: {20*math.log10(float(a2)/float(a1))}")
		except:
			continue
	
	# Colors
	red = '#EF476F'
	blue = '#118AB2'
	purple = '#7b2cbf'
	yellow = '#FFD166'

	# Create the Bode Plot Figure
	fig = sp.make_subplots(rows=2, cols=1, vertical_spacing=0.1, 
		subplot_titles=("Gain Response", "Phase Response"))
	fig.update_layout(title_text="Bode Plot", height=1000, width=1400, title_x=0.5, title_font={"size":22, "color":"#0c0c0c"})

	# Plot the Gain Response
	gain_trace = go.Scatter(x=freq_arr, y=gain_arr, name="Gain (dB)",\
		mode='lines+markers', marker={"symbol":"circle", "size":5, "color":red}, \
		hovertemplate="Freq: %{x:.3f} Hz<br>Gain: %{y:.3f} dB")
	fig.add_trace(gain_trace, row=1, col=1)
	fig.update_yaxes(title_text="Gain (dB)", row=1, col=1, autorangeoptions_clipmin=-30, autorangeoptions_clipmax=30)

	# Plot the Phase Response
	phase_trace = go.Scatter(x=freq_arr, y=phase_arr, name="Phase (deg)",\
		mode='lines+markers', marker={"symbol":"circle", "size":5, "color":blue}, \
		hovertemplate="Freq: %{x:.3f} Hz<br>Phase: %{y:.3f} deg")
	fig.add_trace(phase_trace, row=2, col=1)
	fig.update_yaxes(title_text="Phase (deg)", row=2, col=1, autorangeoptions_clipmin=-180, autorangeoptions_clipmax=180)

	# Update the Layout
	fig.update_annotations(font_size=18)
	fig.update_xaxes(title_text="Frequency (Hz)", title_font={"size":16}, type="log", matches='x', autorangeoptions_clipmin=freq_arr[0]-1, autorangeoptions_clipmax=freq_arr[-1]+10)
	fig.update_yaxes(title_font={"size":16})
	fig.update_traces(line={"smoothing":0.8}, line_shape='spline')

	# Find Cutoff Frequency
	cutoff_index = gain_arr.index(min(gain_arr, key=lambda x:abs(x - (max(gain_arr)-3))))
	fig.add_hline(y=gain_arr[cutoff_index], annotation_text="-3dB Line", annotation_font_size=14, annotation_font_color=purple, line_width=1.0, line_dash="dash", line_color=purple, row=1, col=1)
	fig.add_hline(y=phase_arr[cutoff_index], annotation_text="-3dB Line", annotation_font_size=14, annotation_font_color=purple, line_width=1.0, line_dash="dash", line_color=purple, row=2, col=1)

	# # Save the plot and display
	# fig.write_image(f"{dir_path}/Img/{fn}.png")
	# print(f"Bode Plot Image saved to path: /Img/{fn}.png")
	fig.show()
	


# ===========[ Main Function ]=========== #
def main():
	"""
	Main Function
	"""
	# List all files
	data_files = os.listdir(f"{dir_path}/Data/")
	if len(data_files) == 0:
		print("[ERROR] No Data Files found!")
		exit(1)

	print("[INFO] Following Data-Files were recognized...")
	for i, f in enumerate(data_files):
		print(f"\t[{i+1}] {f}")
	
	# Get user input and plot the Bode Plot
	try:
		file = int(input("  >> Enter the File Number to plot: "))
		plot_bode(data_files[file-1].split(".")[0])
	except:
		print("[ERROR] Invalid Input!")
		exit(1)

# ===========[ Driver Function ]=========== #
if __name__ == '__main__':
	main()