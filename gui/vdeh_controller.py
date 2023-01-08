# -*- coding: utf-8 -*-
"""
VDEH_controller

"""

__component_version__ = "1.0"
__license__ = "MIT License"


#%% import modules/libraries
from .vdeh_form import Ui_MainWindow
from PyQt5.QtWidgets import QFileDialog, QListWidgetItem, QMessageBox
import pandas

#%% define functions


#%% define classes
class vdeh_main_window(Ui_MainWindow):
    def __init__(self, MainWindow, model):
        self.setupUi(MainWindow)
        self.model = model

        # connect the buttons
        # buttons for selecting i/o
        self.pushButton_load_vevolab_files.clicked.connect(
            self.action_load_vevolab_files
        )
        self.menu_Load_VevoLab_File_s.triggered.connect(
            self.action_load_vevolab_files
        )
        self.pushButton_clear_vevolab_files.clicked.connect(
            self.action_clear_vevolab_files
        )
        self.pushButton_load_metadata_settings_file.clicked.connect(
            self.action_load_metadata_settings_file
        )
        self.menu_Load_Metadata_Settings_File.triggered.connect(
            self.action_load_metadata_settings_file
        )
        self.pushButton_clear_metadata_settings_file.clicked.connect(
            self.action_clear_metadata_settings_file
        )
        self.pushButton_set_output_path.clicked.connect(
            self.action_set_output_path
        )
        self.menu_Set_Output_Path.triggered.connect(
            self.action_set_output_path
        )
        self.pushButton_clear_output_path.clicked.connect(
            self.action_clear_output_path
        )
        self.pushButton_reset_form.clicked.connect(
            self.action_reset_form    
        )
        self.menu_Reset.triggered.connect(
            self.action_reset_form
        )

        # buttons for settings (subgui launchers)
        
        
        # buttons for launching extractor and analyzer
        
        # buttons for the help section
        self.menu_User_Manual.triggered.connect(
            self.action_user_manual
        )
        self.menu_About.triggered.connect(
            self.action_about
        )
        
        # set initial state of gui
        self.pushButton_clear_vevolab_files.setHidden(True)
        self.pushButton_clear_metadata_settings_file.setHidden(True)
        self.pushButton_clear_output_path.setHidden(True)

    def action_load_vevolab_files(self):
        self.model.input_paths = QFileDialog.getOpenFileNames(
            None,
            "Select VevoLab Reports",
            "",
            "All Files (*);;Text Files (*.txt);;CSV Files (*.csv)",
        )
        # print(self.model.input_paths)
        self.listWidget_vevolab_files.clear()
        item = None
        for i in self.model.input_paths[0]:
            item = QListWidgetItem(i)
            self.listWidget_vevolab_files.addItem(item)
        self.pushButton_clear_vevolab_files.setHidden(False)    

    def action_clear_vevolab_files(self):
        self.model.input_paths = []
        self.listWidget_vevolab_files.clear()
        self.pushButton_clear_vevolab_files.setHidden(True)

    def action_load_metadata_settings_file(self):
        self.model.settings_path = QFileDialog.getOpenFileName(
            None,
            "Select Metadata/Settings File",
            "",
            "All Files (*);;Excel Files (*.xlsx)",
        )[0]
        self.label_metadata_settings_file.setText(
            f"Metadata/Settings File: {self.model.settings_path}"
        )
        # self.model.load_settings_from_file(self.model)
        self.pushButton_clear_metadata_settings_file.setHidden(False)
        self.pushButton_load_metadata_settings_file.setHidden(True)

    def action_clear_metadata_settings_file(self):
        self.model.settings_path = str()
        
        # settings
        self.animal_data = pandas.DataFrame()
        self.model.timepoint_data = pandas.DataFrame()
        self.model.derived_data = pandas.DataFrame()
        self.model.column_names = pandas.DataFrame()
        self.model.model = pandas.DataFrame()
        
        self.model.settings_changed = False
        
        self.label_metadata_settings_file.setText(
            "Metadata/Settings File: _____"
        )
        self.pushButton_clear_metadata_settings_file.setHidden(True)
        self.pushButton_load_metadata_settings_file.setHidden(False)
        
    def action_save_metadata_settings_file(self):
        new_output_path = QFileDialog.getSaveFileName(
            None,
            "Select filename for saved Metadata/Settings file",
            "",
            "Excel File (*.xlsx)",
            )[0]
        self.model.save_settings_to_file(self.model,new_output_path)
        self.model.settings_path = new_output_path
        self.label_metadata_settings_file.setText(
            f"Metadata/Settings File: {self.model.settings_path}"
        )
    
    def action_set_output_path(self):
        self.model.output_path = QFileDialog.getSaveFileName(
            None,
            "Select filename for output",
            "",
            "Excel File (*.xlsx)",
            )[0]
        self.label_output_path.setText(f"Output Path: {self.model.output_path}")
        self.pushButton_clear_output_path.setHidden(False)
        self.pushButton_set_output_path.setHidden(True)
        
    def action_clear_output_path(self):
        self.model.output_path = str()
        self.label_output_path.setText("Output Path: ____")
        self.pushButton_clear_output_path.setHidden(True)
        self.pushButton_set_output_path.setHidden(False)
        
        
    def action_reset_form(self):
        self.action_clear_vevolab_files()
        self.action_clear_metadata_settings_file()
        self.action_clear_output_path()
        
    def action_user_manual(self):
        pass
    
    def action_about(self):
        QMessageBox.information(
            None,
            'About VDEH', 
            '\n'.join(
                [f'{k} : {v}' for k,v in self.model.version_info.items()]
                )
            )
