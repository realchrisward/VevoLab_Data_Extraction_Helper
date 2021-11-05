# VevoLab_Data_Extraction_Helper

VevoLab Data Extraction Helper
tool to help with conversion of machine output from the VevoLab ultrasound software into an easier to use format for subsequent analysis. Additional features include an exploratory statistics and graphing feature.

written by Christopher Ward (C) 2021, christow@bcm.edu, ward.chris.s@gmail.com
version 3.0

....................
Setting things up
....................
***PC and MAC versions (Precompiled)***
* Extract the zipped folder into a convenient location on your computer
* To run the program Click on the "VevoLab Data Extraction Helper.exe" shortcut (or you can use the .exe file located in the "VevoLab Data * Extraction Helper" subfolder)
* This program uses a seperate file to indicate the desired settings for output, some examples are included 
	** weh metadata - provides full output with animal metadata, multiple output styles compatible with SPSS or Graphpad Prism, graphs, simple stats
	** weh columns list - provides a simple export with all data merged into a single file with one row for each series
	** komp column list - an alternative version of the simple export format with slightly modified column selections and column names
	***(note - please make sure that you confirm the outcome measures that you plan to extract are indicated in the settings file, and were analyzed in the vevolab exported file. if an outcome measure is listed in the settings file but not in the vevolab file then no output will be produced

***Running from source code***
*open the VevoLab_Data_Extraction_Helper_v03.py file using the python interpreter (note your environment will require the standard python libraries (tkinter, re, itertools, traceback) as well as additional libraries (pandas, numpy, scipy, pingouin)

....................
Running the Program
....................
* To run the program Click on the "VevoLab Data Extraction Helper.exe" shortcut (or you can use the .exe file located in the "VevoLab Data * Extraction Helper" subfolder
* The program will sequentially pop open 3 graphical interfaces to get inputs from the user
	1 Select the VevoLab Reports containing the data for extraction
	2 Select the settings/metadata file
	3 Select the output location
* The program will execute, some warnings in the console window are expected. However, if the expected outputs are not produced take note of any error messages provided in the console window

....................
Reporting Bugs
....................
to report any bugs or request new features, email Christopher Ward (christow@bcm.edu)