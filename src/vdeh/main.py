# -*- coding: utf-8 -*-
"""
vdeh_main

VevoLab Data Extraction Helper
@author: Chris Ward (C) 2021


This program provides and interface for the user to select exported data from 
VevoLab as well as create metadata and analysis settings information to 
generate easily read reports

*inputs - used if run from the command line

report_path : a list of the paths to all VevoLab files that contain data
    (GUI tool helps user to select these files)

metadata_path : path to the excel file (.xlsx) that contains metadata and
    settings - the file should contain at least 2 sheets ...
    ...[DerivedData,ColumnNames]...
    (GUI tool helps user to select this file)
    *Animal_Data - table with columns for between subjects metadata - MUST
        CONTAIN 'Animal ID' column as this is the key used to link this data 
        with the VevoLab Data
    *Timepoint_Data - table allowing the user to categorize the timepoints in 
        their study with a column for the Timepoint and a column for the data
        (date values need to match the dates found in the VevoLab data in order 
        to link with this table)
    *Derived_Data - table indicating if any calculations of derived data such
        as age, post treatment time, or time in study ar needed. If the 
        derived data are calculated, they may be used in the stats/graphs. If
        neccessary data for their calculation are missing an error will be 
        raised when attempting to calculate the values, and the stats/graphs
        may not be created
    *Column_Names - table indicating preferred name for columns in the output
        data and the name to expect in the VevoLab data. Typos will create
        problems.
    *Model - single column table indicating the factors to use for statistical
        analysis and for clustering data on plots
        (default sorting order of clustering variables is ascending, which 
        populates the graph from bottom up)

output_path : path to folder and filename for the excel summary that will be 
    produced - this will also be a prefix for the graph files that are produced
    (GUI tool helps user to select/name this file)
    
    
!Warnings! - it is possible to create errors if columns in the data have names
that include reserved 'patterns' of characters
Currently reserved patterns:
    * '__F{}__'

!Warnings! - measurement names that end in a number are assumed to be 
technical replicates of a measurement name that precedes the number. If 
manually naming measurements, keep this in mind. (extraction of AutoLV data 
requires this behavior)

"""


__version__ = "5.4"
__license__ = "MIT License"
__license_text__ = """
MIT License

Copyright (c) 2021 Christopher S Ward

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

# %% import modules/libraries
# import gui
try:
    from gui import vdeh_controller, vdeh_model, vdeh_subgui_controller
except:
    from .gui import vdeh_controller, vdeh_model, vdeh_subgui_controller
# import gui.vdeh_controller as vdeh_controller
# import gui.vdeh_model as vdeh_model
# import gui.vdeh_subgui_controller as vdeh_subgui_controller
import os
import sys
import argparse
from PySide6 import QtWidgets
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile

# %% define functions/classes


# %% define main
def main():

    # parse command line arguments if provided
    parser = argparse.ArgumentParser(description="VevoLab Data Extraction Helper")
    parser.add_argument(
        "-i",
        "--input",
        action="append",
        help="path to VevoLab Report, may combine by declaring multiple times",
    )
    parser.add_argument("-s", "--settings", help="path to settings file")
    parser.add_argument(
        "-o",
        "--output",
        help=(
            "path for output file, "
            + "if no settings file provided then only extracted data is "
            + "created alongside a template 'settings.xlsx'"
        ),
    )
    parser.add_argument(
        "-d", "--dev", help="enable developer mode, specify path for log file"
    )
    parser.add_argument(
        "-l",
        "--loglevel",
        help=(
            "adjust the logging level to use for the gui console "
            + "[DEBUG,INFO,WARNING,ERROR,...] default is INFO"
        ),
    )
    parser.add_argument(
        "-x",
        "--express",
        action="store_true",
        help=(
            "run in express mode, "
            + "use command line arguments and run extraction/analysis "
            + "without launching gui"
        ),
    )

    args, others = parser.parse_known_args()

    if args.express:
        pass
    else:
        # create the application
        print("setting up loader")
        loader = QUiLoader()
        print("startup")
        app = QtWidgets.QApplication(sys.argv)
        print("app started")

        print("preparing ui file")
        ui_file = QFile(
            os.path.join(os.path.dirname(__file__), "gui/vdeh_form_lite.ui")
        )
        window_ui = loader.load(ui_file)
        print("ui loaded")
        window_ui.show()
        print("ui test")

        ui = vdeh_controller.vdeh_main_window(window_ui, vdeh_model.vdeh_model, loader)
        ui.model.version_info = {
            "VevoLab Data Extraction Helper": __version__,
            "vdeh model": vdeh_model.__component_version__,
            "vdeh gui": vdeh_controller.__component_version__,
            "vdeh subguis": vdeh_subgui_controller.__component_version__,
        }
        print("ui test - main")

        # if user specified --dev or --loglevel update model
        if args.dev:
            ui.model.log_file_path = args.dev
        if args.loglevel:
            ui.model.log_level = args.loglevel

        # show the gui
        # MainWindow.show()
        # ui.ui.show()

        sys.exit(app.exec_())


# %% run main
if __name__ == "__main__":
    main()
