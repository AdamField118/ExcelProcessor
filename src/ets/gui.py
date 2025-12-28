"""
ETS Processor - Main GUI Window

This module contains the PyQt5 GUI for the ETS file processing application.
This is a variant of the Excel processor with different fields and processing logic.
"""

import sys
import os

# Add paths for both development and frozen executable
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    base_path = sys._MEIPASS
    sys.path.insert(0, base_path)
    sys.path.insert(0, os.path.join(base_path, 'src'))
else:
    # Running as normal Python script
    base_path = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, os.path.abspath(os.path.join(base_path, '..', '..')))

from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QWidget, QPushButton, QLabel, QLineEdit, QTextEdit, 
                             QFileDialog, QMessageBox, QProgressBar, QGroupBox,
                             QGridLayout, QScrollArea, QSizePolicy)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon

# Import components - works in both dev and frozen modes
try:
    from src.gui.components import FileDropArea
except ImportError:
    from gui.components import FileDropArea

try:
    from src.ets.processor import ETSProcessor
except ImportError:
    from ets.processor import ETSProcessor

try:
    from src.utils.validators import FileValidator
    from src.utils.logger import get_simple_logger
    from src.utils.resource_manager import ResourceManager
except ImportError:
    from utils.validators import FileValidator
    from utils.logger import get_simple_logger
    from utils.resource_manager import ResourceManager


class ProcessingThread(QThread):
    """Background thread for processing files"""
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    processing_complete = pyqtSignal(str)
    processing_error = pyqtSignal(str)
    
    def __init__(self, input_files, config_values, output_path):
        super().__init__()
        self.input_files = input_files
        self.config_values = config_values
        self.output_path = output_path
        self.logger = get_simple_logger(__name__)
    
    def run(self):
        try:
            self.logger.info(f"Starting ETS processing of {len(self.input_files)} files")
            
            # Validation phase (10% of progress)
            total_files = len(self.input_files)
            self.status_updated.emit("Validating files...")
            
            for i, file_path in enumerate(self.input_files):
                self.logger.info(f"Validating file: {file_path}")
                
                if not FileValidator.validate_file_accessibility(file_path):
                    raise Exception(f"Cannot access file: {file_path}")
                
                progress = int((i + 1) / total_files * 10)
                self.progress_updated.emit(progress)
            
            # Loading phase (20% of progress)
            self.status_updated.emit("Loading input files...")
            loaded_data = []
            
            for i, file_path in enumerate(self.input_files):
                self.logger.info(f"Loading file: {file_path}")
                file_type = FileValidator.get_file_type(file_path)
                self.status_updated.emit(f"Loading file {i+1}/{total_files}: {os.path.basename(file_path)} ({file_type})")
                
                from src.core.file_handler import FileHandler
                file_data = FileHandler.load_excel_file(file_path)
                loaded_data.append(file_data)
                
                progress = 10 + int((i + 1) / total_files * 20)
                self.progress_updated.emit(progress)
            
            # Processing phase (70% of progress)
            self.status_updated.emit("Processing data...")
            self.progress_updated.emit(30)
            
            # Create progress callback
            def progress_callback(message):
                self.status_updated.emit(message)
                if "%" in message:
                    try:
                        percent_part = message.split("(")[1].split("%")[0]
                        percent = int(percent_part)
                        gui_progress = 30 + int(percent * 0.65)
                        self.progress_updated.emit(gui_progress)
                    except:
                        pass
            
            # Use ETS processor
            processor = ETSProcessor()
            result = processor.process_data(
                loaded_data,
                self.config_values,
                self.output_path,
                progress_callback
            )
            
            self.progress_updated.emit(100)
            self.status_updated.emit("Processing complete!")
            
            self.logger.info("ETS processing completed successfully")
            success_message = (f"Successfully processed {result['total_files_processed']} files "
                             f"and saved to {os.path.basename(self.output_path)}!")
            self.processing_complete.emit(success_message)
            
        except Exception as e:
            error_msg = f"Error during processing: {str(e)}"
            self.logger.error(error_msg)
            self.status_updated.emit(f"ERROR: {error_msg}")
            self.processing_error.emit(error_msg)


class ETSMainWindow(QMainWindow):
    """
    Main window for ETS Processor
    
    TODO: Customize the input fields based on ETS requirements
    """
    
    def __init__(self):
        super().__init__()
        self.input_files = []
        self.output_path = ""
        
        # Use ETS-specific config file
        self.resource_manager = ResourceManager(config_file="ets_app_config.json")
        self.logger = get_simple_logger(__name__)
        
        # Get window configuration
        window_config = self.resource_manager.get_config("window")
        
        window_title = "ETS Processor"
        if window_config:
            window_title = window_config.get("title", window_title)
        
        self.setWindowTitle(window_title)
        self.setMinimumSize(450, 400)
        self.resize(650, 650)
        
        # Set icon
        app_icon = self.resource_manager.get_icon("ets_logo.png")
        if app_icon:
            self.setWindowIcon(QIcon(app_icon))
        else:
            app_icon = self.style().standardIcon(self.style().SP_ComputerIcon)
            self.setWindowIcon(app_icon)
        
        self.setup_ui()
        self.load_main_stylesheet()
        
        self.logger.info("ETS Application started")
    
    def load_main_stylesheet(self):
        """Load main window stylesheet"""
        main_stylesheet = self.resource_manager.get_stylesheet("main_window.qss")
        if main_stylesheet:
            self.setStyleSheet(main_stylesheet)
        
    def setup_ui(self):
        # Create scroll area
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        content_widget = QWidget()
        scroll_area.setWidget(content_widget)
        self.setCentralWidget(scroll_area)
        
        main_layout = QVBoxLayout()
        content_widget.setLayout(main_layout)
        content_widget.setMinimumSize(450, 540)
        
        # Title - from config
        title_config = self.resource_manager.get_config("ui.title")
        
        title_text = "ETS Processor"
        title_font_size = 16
        title_bold = True
        
        if title_config:
            title_text = title_config.get("text", title_text)
            title_font_size = title_config.get("font_size", title_font_size)
            title_bold = title_config.get("bold", title_bold)
        
        title = QLabel(title_text)
        title.setAlignment(Qt.AlignCenter)
        title.setObjectName("titleLabel")
        title.setWordWrap(True)
        
        font = QFont()
        font.setPointSize(title_font_size)
        font.setBold(title_bold)
        title.setFont(font)
        main_layout.addWidget(title)
        
        # Get labels from config
        labels_config = self.resource_manager.get_config("ui.labels")
        
        input_files_label = "Input Files"
        configuration_label = "Configuration"
        output_label = "Output Location"
        
        if labels_config:
            input_files_label = labels_config.get("input_files", input_files_label)
            configuration_label = labels_config.get("configuration", configuration_label)
            output_label = labels_config.get("output", output_label)
        
        # File drop area - REUSED component
        file_group = QGroupBox(input_files_label)
        file_group.setObjectName("inputFilesGroup")
        file_layout = QVBoxLayout()
        
        self.drop_area = FileDropArea()
        self.drop_area.setMinimumHeight(100)
        self.drop_area.setMaximumHeight(200)
        self.drop_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.drop_area.files_dropped.connect(self.on_files_selected)
        
        file_layout.addWidget(self.drop_area)
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)
        
        # Configuration fields - ETS SPECIFIC
        config_group = QGroupBox(configuration_label)
        config_group.setObjectName("configurationGroup")
        config_layout = QGridLayout()
        
        self.config_fields = {}
        
        # ETS-specific fields
        field_info = [
            ('part_name', 'Part Name:'),
            ('revision_number', 'Revision Number:'),
            ('customer_pn', 'Customer P/N:'),
            ('customer_po', 'Customer PO:'),
            ('ocv_min', 'OCV Min:'),
            ('ccv_min', 'CCV Min:'),
        ]
        
        placeholders = {
            'part_name': 'Enter part name (e.g., 3B0036-AN-03)...',
            'revision_number': 'Enter revision number (e.g., 002)...',
            'customer_pn': 'Enter customer P/N (e.g., 01-5032-03)...',
            'customer_po': 'Enter customer PO (e.g., PO71906)...',
            'ocv_min': 'Enter OCV minimum value (e.g., 3.92)...',
            'ccv_min': 'Enter CCV minimum value (e.g., 3)...',
        }
        
        for i, (field_key, label_text) in enumerate(field_info):
            label_widget = QLabel(label_text)
            label_widget.setObjectName(f"fieldLabel{i}")
            
            line_edit = QLineEdit()
            line_edit.setObjectName(f"fieldInput{i}")
            line_edit.setPlaceholderText(placeholders.get(field_key, f"Enter {label_text.lower().replace(':', '')}..."))
            
            # Set default values for OCV/CCV min
            if field_key == 'ocv_min':
                line_edit.setText('3.92')
            elif field_key == 'ccv_min':
                line_edit.setText('3')
            
            config_layout.addWidget(label_widget, i, 0)
            config_layout.addWidget(line_edit, i, 1)
            
            self.config_fields[field_key] = line_edit
        
        config_layout.setColumnStretch(1, 1)
        config_group.setLayout(config_layout)
        main_layout.addWidget(config_group)
        
        # Output section - from config
        output_config = self.resource_manager.get_config("ui.output")
        
        no_location_message = "No output location selected"
        save_button_text = "Choose Save Location"
        
        if output_config:
            no_location_message = output_config.get("no_location_message", no_location_message)
            save_button_text = output_config.get("button_text", save_button_text)
        
        output_group = QGroupBox(output_label)
        output_group.setObjectName("outputGroup")
        output_layout = QVBoxLayout()
        
        path_layout = QHBoxLayout()
        
        self.output_path_label = QLabel(no_location_message)
        self.output_path_label.setObjectName("outputPathLabel")
        self.output_path_label.setStyleSheet("color: #666; font-style: italic;")
        self.output_path_label.setWordWrap(True)
        
        self.save_as_btn = QPushButton(save_button_text)
        self.save_as_btn.setObjectName("saveAsButton")
        self.save_as_btn.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        
        save_icon = self.style().standardIcon(self.style().SP_DialogSaveButton)
        self.save_as_btn.setIcon(save_icon)
        self.save_as_btn.clicked.connect(self.choose_output_location)
        
        path_layout.addWidget(QLabel("Output File:"))
        path_layout.addWidget(self.output_path_label, 1)
        path_layout.addWidget(self.save_as_btn)
        output_layout.addLayout(path_layout)
        
        output_group.setLayout(output_layout)
        main_layout.addWidget(output_group)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # Status label
        status_config = self.resource_manager.get_config("ui.status_messages")
        ready_message = "Ready"
        
        if status_config:
            ready_message = status_config.get("ready", ready_message)
        
        self.status_label = QLabel(ready_message)
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        main_layout.addWidget(self.status_label)
        
        main_layout.addStretch()
        
        # Process button - from config
        button_config = self.resource_manager.get_config("ui.buttons.process")
        
        process_button_text = "Process Files"
        min_height = 35
        
        if button_config:
            process_button_text = button_config.get("text", process_button_text)
            min_height = button_config.get("min_height", min_height)
        
        self.process_btn = QPushButton(process_button_text)
        self.process_btn.setObjectName("processButton")
        self.process_btn.setMinimumHeight(min_height)
        self.process_btn.setEnabled(False)
        self.process_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        process_icon = self.style().standardIcon(self.style().SP_DialogApplyButton)
        self.process_btn.setIcon(process_icon)
        self.process_btn.clicked.connect(self.process_files)
        
        main_layout.addWidget(self.process_btn)
        
        # Status bar - from config
        if status_config:
            ready_message = status_config.get("ready", "Ready - Select input files and configure settings")
        else:
            ready_message = "Ready - Select input files and configure settings"
        
        self.statusBar().showMessage(ready_message)
    
    def on_files_selected(self, files):
        self.input_files = files
        self.logger.info(f"Selected {len(files)} input files")
        self.update_process_button_state()
        self.statusBar().showMessage(f"Selected {len(files)} file(s)")
    
    def choose_output_location(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Processed File As",
            "ets_processed_data.xlsx",
            "Excel Files (*.xlsx);;All Files (*)"
        )
        
        if file_path:
            if not FileValidator.validate_output_path(file_path):
                QMessageBox.warning(
                    self,
                    "Invalid Output Path",
                    "The selected output location is not accessible or the filename is invalid."
                )
                return
            
            self.output_path = file_path
            self.output_path_label.setText(os.path.basename(file_path))
            self.output_path_label.setStyleSheet("color: #000;")
            self.logger.info(f"Output path set to: {file_path}")
            self.update_process_button_state()
    
    def update_process_button_state(self):
        can_process = bool(self.input_files and self.output_path)
        self.process_btn.setEnabled(can_process)
        
        if can_process:
            active_button_style = self.resource_manager.get_stylesheet("buttons_active.qss")
            if active_button_style:
                self.process_btn.setStyleSheet(active_button_style)
        else:
            disabled_button_style = self.resource_manager.get_stylesheet("buttons_disabled.qss")
            if disabled_button_style:
                self.process_btn.setStyleSheet(disabled_button_style)
    
    def get_config_values(self):
        """Get all configuration field values as a dictionary"""
        values = {}
        for key, field in self.config_fields.items():
            values[key] = field.text().strip()
        return values
    
    def process_files(self):
        if not self.input_files:
            QMessageBox.warning(self, "No Files", "Please select input files first.")
            return
            
        if not self.output_path:
            QMessageBox.warning(self, "No Output Location",
                              "Please choose where to save the output file.")
            return
        
        config_values = self.get_config_values()
        
        # Validate ETS-specific fields
        validation_errors = ETSProcessor.validate_config(config_values)
        if validation_errors:
            QMessageBox.warning(self, "Invalid Configuration",
                              "Please check the following fields:\n" + "\n".join(validation_errors))
            return
        
        self.logger.info("Starting ETS file processing...")
        
        # Disable UI during processing
        self.process_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.statusBar().showMessage("Processing files...")
        
        # Start processing thread
        self.processing_thread = ProcessingThread(
            self.input_files,
            config_values,
            self.output_path
        )
        self.processing_thread.progress_updated.connect(self.progress_bar.setValue)
        self.processing_thread.status_updated.connect(self.status_label.setText)
        self.processing_thread.processing_complete.connect(self.on_processing_complete)
        self.processing_thread.processing_error.connect(self.on_processing_error)
        self.processing_thread.start()
    
    def on_processing_complete(self, message):
        self.progress_bar.setVisible(False)
        self.status_label.setText("Processing complete!")
        self.process_btn.setEnabled(True)
        self.statusBar().showMessage("Processing complete!")
        
        self.logger.info("ETS processing completed successfully")
        QMessageBox.information(self, "Success", message)
        
        # Ask if user wants to open the output file
        reply = QMessageBox.question(
            self,
            "Open File?",
            "Would you like to open the generated file?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                from src.core.file_handler import FileHandler
                if FileHandler.open_file(self.output_path):
                    self.logger.info(f"Opened output file: {self.output_path}")
                else:
                    raise Exception("File opening failed")
            except Exception as e:
                self.logger.error(f"Failed to open output file: {e}")
                QMessageBox.warning(
                    self,
                    "Cannot Open File",
                    f"Could not open the output file:\n{str(e)}"
                )
    
    def on_processing_error(self, error_message):
        self.progress_bar.setVisible(False)
        self.status_label.setText("Processing failed!")
        self.process_btn.setEnabled(True)
        self.statusBar().showMessage("Processing failed!")
        
        self.logger.error(f"ETS processing failed: {error_message}")
        QMessageBox.critical(self, "Processing Error", error_message)


def main():
    logger = get_simple_logger("ets_processor")
    logger.info("Starting ETS Processor application")
    
    app = QApplication(sys.argv)
    
    # Set application properties from config
    # Use ETS-specific config file
    resource_manager = ResourceManager(config_file="ets_app_config.json")
    app_config = resource_manager.get_config("application")
    
    app_name = "ETS Processor"
    app_version = "1.0"
    
    if app_config:
        app_name = app_config.get("name", app_name)
        app_version = app_config.get("version", app_version)
    
    app.setApplicationName(app_name)
    app.setApplicationVersion(app_version)
    
    try:
        # Load global stylesheet
        global_stylesheet = resource_manager.get_stylesheet("app_global.qss")
        if global_stylesheet:
            app.setStyleSheet(global_stylesheet)
            logger.info("Loaded global application stylesheet")
        
        # Set application icon
        app_icon_path = resource_manager.get_icon_path("pixel_logo.png")
        if app_icon_path and os.path.exists(app_icon_path):
            app.setWindowIcon(QIcon(app_icon_path))
            logger.info("Set application icon")
        
        window = ETSMainWindow()
        window.show()
        
        logger.info("ETS Main window displayed successfully")
        exit_code = app.exec_()
        logger.info(f"Application exited with code: {exit_code}")
        sys.exit(exit_code)
        
    except Exception as e:
        logger.error(f"Critical error in ETS application: {e}")
        QMessageBox.critical(None, "Critical Error",
                           f"A critical error occurred:\n{str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()