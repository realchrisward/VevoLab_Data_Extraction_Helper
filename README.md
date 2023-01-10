# VevoLab_Data_Extraction_Helper

## VevoLab Data Extraction Helper
tool to help with conversion of machine output from the VevoLab ultrasound software into an easier to use format for subsequent analysis. Additional features include an exploratory statistics and graphing feature.

written by Christopher Ward (C) 2021, christow@bcm.edu, ward.chris.s@gmail.com



## Setting things up

### PC and MAC versions (Precompiled)
* Extract the zipped folder into a convenient location on your computer
* In the root level of the unzipped folder, run the program Click on the "VDEH" executable (or make a shortcut pointing to this file)
* Use the program

### Running from source code
* Create a python environment
* install dependencies (requirements.txt)
* clone the repository
* launch the program using vdeh.py


## Running the Program
* You may run the program as a command line tool using the following command line arguments
    * -i, --input : path to VevoLab Report, may combine by declaring multiple times
    * -o, --output : path to use for the output xlsx file
    * -s, --settings : path to the metadata/settings file
    * -d, --dev : [FLAG] run using developer mode (creates a log file to aid with debugging)
    * -l, --loglevel : modify the logging level used by the gui console [DEBUG, INFO, WARNING, ERROR, ...] default is INFO
    * -x, --express : [FLAG] run in expless mode (command line only)
* You may run the program as a gui operated tool
    * Select the VevoLab Reports containing the data for extraction
    * Select the metadata/settings file
        * adjust the settings/metadata as needed
    * Select the output location
    * Run the Extraction/Analysis

## Reporting Bugs
to report any bugs or request new features, email Christopher Ward (christow@bcm.edu)
