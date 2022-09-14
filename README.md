# VevoLab_Data_Extraction_Helper

VevoLab Data Extraction Helper
tool to help with conversion of machine output from the VevoLab ultrasound software into an easier to use format for subsequent analysis. Additional features include an exploratory statistics and graphing feature.

written by Christopher Ward (C) 2021, christow@bcm.edu, ward.chris.s@gmail.com
version 4.1

manual available at
https://realchrisward.github.io/VevoLab_Data_Extraction_Helper/index.html

....................
Setting things up
....................
***PC and MAC versions (Precompiled)***
(not yet available)
* Extract the zipped folder into a convenient location on your computer
* To run the program Click on the "VevoLab Data Extraction Helper.exe" shortcut (or you can use the .exe file located in the "VevoLab Data * Extraction Helper" subfolder)

***Running from source code***
*open the VevoLab_Data_Extraction_Helper.py file using the python interpreter (note your environment will require the standard python libraries (tkinter, re, itertools, traceback) as well as additional libraries (pandas, numpy, scipy, pingouin)

....................
Running the Program
....................
* To run the program Click on the "VevoLab Data Extraction Helper.exe" shortcut (or you can use the .exe file located in the "VevoLab Data * Extraction Helper" subfolder
* The program will sequentially pop open 3 graphical interfaces to get inputs from the user
	1 Select the VevoLab Reports containing the data for extraction
	2 Select the settings/metadata file (or click cancel to generate a new file based on the current VevoLab Report)
	3 Select the output location
	4 [only if 'cancel' was selected in Step 2] Select the output location for the new settings/metadata template file
* The program will execute, some warnings in the console window are expected. However, if the expected outputs are not produced take note of any error messages provided in the console window

....................
Reporting Bugs
....................
to report any bugs or request new features, email Christopher Ward (christow@bcm.edu)
