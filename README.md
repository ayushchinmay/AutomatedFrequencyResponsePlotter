# Automated Frequency Response Plotter
Python script that allows the users to sweep over a specified range of frequencies, calculate the frequency/phase response, and generate a bode-plot for any given physical-circuit.

## File/Folder Descriptions
* **/Data:**  *Data folder that stores the '.csv' data-files saved by the `FrequencyResponseCapture.py` program*
* **/Img:**  *Image folder that stores the '.png' data-files saved by the `FrequencyResponseCapture.py` program*
* **FrequencyResponseCapture.py:**  *Python program to sweep frequencies and generate bode-plot*
* **installPythonModules.bat:**  *Batch script to install required python modules automatically*
* **interactivePlot.py:**  *Python program to open an interactive plot using the data-files saved in the */Data* folder*
* **requirements.txt:**  *Text file containing names of the required python modules*
* **README.md:**  *README file containing the project description*

# Program Description
The program flow consists of the following processes:
![Program Execution Flowchart](https://github.com/ayushchinmay/AutomatedFrequencyResponsePlotter/blob/main/readme_references/program-flow.png)
* Find the device ID to create a VISA instrument control resource object.
* Initialize the oscilloscope to default settings and check for errors.
* If no errors were encountered, take user input for start/stop frequencies and the number of iteration steps.
* Start the waveform generator on the oscilloscope at the starting frequency, and autoscale to get a good view of the waveform.
* Step through the frequencies with increments calculated from the desired number of iterations
* Get the waveform measurements such as input/output signal frequency, amplitude (peak-to-peak) and the phase difference
* If the measurements captured are valid, save them to a *`.csv`* file with the following column headers:
       `CH1_FREQ [Hz], CH1_AMPL [Vpp], CH2_FREQ [Hz], CH2_AMPL [Vpp], PHASE_DIFF [Deg], GAIN [dB]`
* Once measurements are made for all the frequencies in the provided range, plot the frequency (magnitude) response and the phase response of the given circuit.
* Display the Bode-plot and save the figure as a *`.png`* image.
* Exit the program.

## Getting Started

### Hardware Requirements
* Computer with USB-A port
* Agilent Infiniium DSO-X 3012A Oscilloscope
* Breadboard + Amplifier/Filter Circuit
* 3× BNC-Banana Cables
* USB A-to-B cable
* Jumper Cables

### Software Dependencies
* [Windows 10/11](https://www.microsoft.com/software-download/windows11)
* [NI-VISA Drivers](https://www.ni.com/en/support/downloads/drivers/download.ni-visa.html#494653)
* [NI Device Drivers](https://www.ni.com/en/support/downloads/drivers/download.ni-device-drivers.html#327643)
     OR
* [IVI Foundation Shared Components](https://www.ivifoundation.org/downloads/SharedComponents.htm)
* [Python 3](https://www.python.org/downloads/)

### Setting up Software
* Download the project *.zip* folder by clicking **'<> Code'** dropdown menu and selecting **'Download ZIP'**
* Download & install the NI Drivers listed under the dependencies
* Download the IVI Foundation Shared Components
     Ensure the presence of GlobMgr.dll at the following location: *`"C:/Program Files (x86)/IVI Foundation/VISA/VisaCom/GlobMgr.dll"`*
*  Install Python 3 and run the *`installPythonModules.bat`* to download the following modules required for this project using the requirements.txt file.
    * [comtypes](https://pypi.org/project/comtypes/)
    * [matplotlib](https://pypi.org/project/matplotlib/)
    * [pyVisa](https://pypi.org/project/PyVISA/)

### Cable Connections
![Illustration of Cable Connections for Testing Circuit](https://github.com/ayushchinmay/AutomatedFrequencyResponsePlotter/blob/main/readme_references/cable-connection-illustration.png)
* Connect the USB A-to-B Cable (USB-B) end behind the Oscilloscope, and the other end (USB-A) into the computer.
* Use 2×BNC-Banana cables to split the **`GEN-OUT`** port output such that one end is INPUT into the **`CIRCUIT`**, and another into **`CHANNEL-1`** of the oscilloscope. 
* Attach another BNC-Banana cable to connect the output of the **`CIRCUIT`** into **`CHANNEL-2`** of the oscilloscope.

### Executing program
* Run the *`FrequencyResponseCapture.py`* by double clicking on it.
    Alternatively, the following terminal command (while in the directory) can be used: ```python3 -m FrequencyResponseCapture.py```
* Enter the Start Frequency, Stop Frequency, and the number of Steps. Frequency inputs are floating-point numbers in Hertz.
    * If an invalid input is provided, the script will use the default values: 
        * Start Frequency: 100 Hz
        * Stop Frequency: 10000 Hz
        * Number of Steps: 25 steps
* Once the program execution is completed, a figure displaying the Frequency Response plot will appear
    * The file containing the captured data can be found in `../Data/` folder as a `.csv` file with the name-format: *'Bode_MM-DD-YYYY_HR-MIN-SEC.csv'* (see image below)
![Screenshot of a sample data-file](https://github.com/ayushchinmay/AutomatedFrequencyResponsePlotter/blob/main/readme_references/data-file-example.png)
    * The bode-plot figure is saved as a `.png` image in `../Img/` folder with the filename: *'Bode_MM-DD-YYYY_HR-MIN-SEC.png'* (see image below)
![Screenshot of a sample bode-plot](https://github.com/ayushchinmay/AutomatedFrequencyResponsePlotter/blob/main/readme_references/bode-plot-example.png)
* If any errors are encountered during the program execution, the error will be printed to the command-terminal, and the program will will terminate after saving the data captured until the error.
* Once the program execution ends, `interactivePlot.py` program can be used to open an interactive-plot in the browser.
    * The -3dB line is plotted based on the peak-value, which can be useful to find cut-off frequencies.

## Help
* Depending on the type of filter, the frequency range to sweep over may vary. Due to the **AUTOSCALE** function of the oscilloscope prioritizing **CHANNEL-1** waveform, if the output amplitude is too low, the measurements made may be erroneous; In such an event, the program will try to re-capture the measurements -- After 5 retries, the program will break out of the loop, save the data and exit.
    * In the case of low-pass filters, if the test-frequency is too high, the program will print the last *good-upper-frequency* that can be tested.
* Due to the large response-time for lower frequencies, it is recommended that the start-frequency be higher than **5 Hz**. 
    * In the case that the frequency is too low for the oscilloscope measurements to be captured efficiently, the program will increment the start frequency by 2 Hz each time it captures erroneous  data.
* Erroneous data can be detected in the following cases:
    * If the frequency of the INPUT Signal is more than 50% of the frequency of OUTPUT Signal
    * If the Frequency or Amplitude of the Output signal reads higher than 10 GHz or 10 GVpp. In these cases, the oscilloscope reads `98.99E36` (see screenshot below).
![Screenshot of Commandline during Program-Execution](https://github.com/ayushchinmay/AutomatedFrequencyResponsePlotter/blob/main/readme_references/program-execution-example.png)

## Authors
- [Ayush Chinmay](https://github.com/ayushchinmay)
- Andrew Morgan
- Jensen Dygert

## Version History
* **0.1.2**: Added Interactive Plot
* **0.1**: Initial Release

## Further Developments
* The current program only supports Aglient DSO-X 3012A oscilloscopes. In the future I intend on modifying the code to support a suite of instruments to configure/capture data.
* Feel free to make any modifications that may develop the code to support more instruments, or to upgrade the functionality fo the program.

## Resources
Some useful resources to keep in mind while utilizing this program...
* [Agilent Infiniivision 3012A DataSheet](https://www.keysight.com/nl/en/support/DSOX3012A/oscilloscope-100-mhz-2-channels.html)
* [Programming Manual: Keysight InfiniiVision 3000 X-Series Oscilloscope](https://www.keysight.com/us/en/assets/9018-06894/programming-guides/9018-06894.pdf)
* [MatPlotLib Documentation](https://matplotlib.org/stable/index.html)
* [Operational Amplifier Basics](https://www.electronics-tutorials.ws/opamp/opamp_1.html)
* [A Designer's Guide to Instrumentation Amplifiers](https://www.analog.com/media/en/training-seminars/design-handbooks/designers-guide-instrument-amps-complete.pdf)
* [Analog Devices Filter Design Tool](https://tools.analog.com/en/filterwizard/)
