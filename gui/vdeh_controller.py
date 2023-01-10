# -*- coding: utf-8 -*-
"""
VDEH_controller

"""

__component_version__ = "1.0"
__license__ = "MIT License"


#%% import modules/libraries
from .vdeh_form import Ui_MainWindow
from PyQt5.QtWidgets import QFileDialog, QListWidgetItem, QMessageBox
from PyQt5.QtWidgets import QTextEdit
import pandas
import logging
import sys



#%%
class VDEH_Logger():
    def __init__(
            self,
            gui_loglevel: int = logging.INFO,
            console_loglevel: int = logging.ERROR,
            file_loglevel: int = logging.DEBUG,
            log_file_path: str = None,
            gui_handler: QTextEdit = None,
            logname: str = __name__
            ):
        
        self.log_levels = {
            'notset':0,
            'debug':10,
            'info':20,
            'warning':30,
            'error':40,
            'critical':50
            }

        self.logger = logging.getLogger(logname)
        self.logger.setLevel(logging.DEBUG)
        
        self.gui_handler = gui_handler
        self.gui_loglevel = self.fix_level(gui_loglevel)

        # create format for log and apply to handlers
        log_format = logging.Formatter(
                '%(asctime)s | %(name)s | %(levelname)s | %(message)s'
                )
        
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(console_loglevel)
        console_handler.setFormatter(log_format)
        self.logger.addHandler(console_handler)
        
        if log_file_path:
            file_handler = logging.FileHandler(log_file_path)
            file_handler.setLevel(file_loglevel)
            file_handler.setFormatter(log_format)
            self.logger.addHandler(file_handler)
        
        # log initial inputs
        self.log('info','VDEH Logger Started')
        
    def fix_level(self,level):
        if type(level) is not int:
            if type(level) is not str:
                level = 40
            elif level.lower() in self.log_levels:
                level = self.log_levels[level.lower()]
            else:
                level = 40
        elif level not in self.log_levels.values():
            level = 40
        
        return level
        
    def log(
            self,
            level,
            message,
            gui_message: str = None, 
            gui_color: str = None,
            gui_style: str = None
            ):
        print(level)
        if type(level) is not int:
            if type(level) is not str:
                old_level = level
                message += f' |Abnormal log level provided: {old_level}|'
                level = 40
            elif level.lower() in self.log_levels:
                level = self.log_levels[level.lower()]
            else:
                old_level = level
                message += f' |Abnormal log level provided: {old_level}|'
                level = 40
        elif level not in self.log_levels.values():
            old_level = level
            message += f' |Abnormal log level provided: {old_level}|'
            level = 40
        print(level)
        self.logger.log(level,message)
        if level >= self.gui_loglevel and self.gui_handler:
            if not gui_color:
                if level < 20:
                    gui_color = 'green'
                elif level > 20:
                    gui_color = 'red'
                else:
                    gui_color = 'black'
            if not gui_style:
                if level >= 40:
                    gui_style = 'strong'
            
            
            self.gui_handler.insertHtml(
                (
                    f'<span style="color:{gui_color}"><{gui_style}>'
                    +f'{message}'
                    +f'</{gui_style}></span><br>'
                    )
                )
        






#%% define classes
class vdeh_main_window(Ui_MainWindow):
    def __init__(self, MainWindow, model):
        self.setupUi(MainWindow)
        self.model = model

        self.logger = VDEH_Logger(
            gui_loglevel=self.model.log_level,
            log_file_path=self.model.log_file_path,
            gui_handler=self.textEdit_status
            )
        self.model.logger = self.logger



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
        self.pushButton_generate_metadata_settings_template.clicked.connect(
            self.action_extract_data
        )
        
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
        )[0]
        # print(self.model.input_paths)
        self.listWidget_vevolab_files.clear()
        item = None
        for i in self.model.input_paths:
            item = QListWidgetItem(i)
            self.listWidget_vevolab_files.addItem(item)
        self.pushButton_clear_vevolab_files.setHidden(False)
        self.logger.log(
            'info',
            f'VevoLab Report files selected: {",".join([f for f in self.model.input_paths])}'
            )

    def action_clear_vevolab_files(self):
        self.model.input_paths = []
        self.model.model_data = pandas.DataFrame()
        self.listWidget_vevolab_files.clear()
        self.pushButton_clear_vevolab_files.setHidden(True)
        self.logger.log(
            'info',
            'VevoLab Report files cleared'
            )

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
        self.model.load_settings_from_file(self.model)
        self.pushButton_clear_metadata_settings_file.setHidden(False)
        self.pushButton_load_metadata_settings_file.setHidden(True)
        self.logger.log(
            'info',
            f'Metadata/Settings file selected: {self.model.settings_path}'
            )

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
        self.logger.log(
            'info',
            'Metadata/Settings file cleared'
            )
        
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
        self.logger.log(
            'info',
            f'Metadata/Settings File saved - and set as current settings: {self.model.settings_path}'
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
        self.logger.log(
            'info',
            f'Output location set: {self.model.output_path}')
        
    def action_clear_output_path(self):
        self.model.output_path = str()
        self.label_output_path.setText("Output Path: ____")
        self.pushButton_clear_output_path.setHidden(True)
        self.pushButton_set_output_path.setHidden(False)
        self.logger.log(
            'info',
            'Output location cleared'
            )
        
    def action_extract_data(self):
        self.model.check_data(self.model)
        print(self.model.column_names)
        print(self.model.model_data)
        
    def action_reset_form(self):
        self.action_clear_vevolab_files()
        self.action_clear_metadata_settings_file()
        self.action_clear_output_path()
        
    def action_user_manual(self):
        self.logger.log(
            'info',
            'Help -> User Manual')
    
    def action_about(self):
        QMessageBox.information(
            None,
            'About VDEH', 
            '\n'.join(
                [f'{k} : {v}' for k,v in self.model.version_info.items()]
                )
            )
