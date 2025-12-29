"""
Unified File Processor - Main GUI Window

This module contains the PyQt5 GUI for the unified file processing application
that can handle both Measurement and ETS data.
"""

import sys
import os

# Add paths for both development and frozen executable
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
    sys.path.insert(0, base_path)
    sys.path.insert(0, os.path.join(base_path, 'src'))
else:
    base_path = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, os.path.abspath(os.path.join(base_path, '..', '..')))

from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QWidget, QPushButton, QLabel, QLineEdit, QTextEdit, 
                             QFileDialog, QMessageBox, QProgressBar, QGroupBox,
                             QGridLayout, QScrollArea, QSizePolicy, QComboBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon

try:
    from src.gui.components import FileDropArea
    from src.core.unified_processor import UnifiedProcessor
    from src.utils.validators import FileValidator
    from src.utils.logger import get_simple_logger
    from src.utils.resource_manager import ResourceManager
except ImportError:
    from gui.components import FileDropArea
    from core.unified_processor import UnifiedProcessor
    from utils.validators import FileValidator
    from utils.logger import get_simple_logger
    from utils.resource_manager import ResourceManager


class ProcessingThread(QThread):
    """Background thread for processing files"""
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    processing_complete = pyqtSignal(str)
    processing_error = pyqtSignal(str)
    
    def __init__(self, mode, input_files, measurement_config, ets_config, 
                 measurement_output, ets_output):
        super().__init__()
        self.mode = mode
        self.input_files = input_files
        self.measurement_config = measurement_config
        self.ets_config = ets_config
        self.measurement_output = measurement_output
        self.ets_output = ets_output
        self.logger = get_simple_logger(__name__)
    
    def run(self):
        try:
            self.logger.info(f"Starting unified processing in {self.mode} mode")
            self.status_updated.emit(f"Starting {self.mode} processing...")
            
            # Progress callback
            def progress_callback(message):
                self.status_updated.emit(message)
                if "%" in message:
                    try:
                        percent_part = message.split("(")[1].split("%")[0]
                        percent = int(percent_part)
                        gui_progress = int(percent * 0.9)  # 0-90%
                        self.progress_updated.emit(gui_progress)
                    except:
                        pass
            
            # Process files
            results = UnifiedProcessor.process_unified(
                self.mode,
                self.input_files,
                self.measurement_config,
                self.ets_config,
                self.measurement_output,
                self.ets_output,
                progress_callback
            )
            
            self.progress_updated.emit(100)
            self.status_updated.emit("Processing complete!")
            
            # Build success message
            success_parts = []
            if results.get('measurement'):
                m_result = results['measurement']
                success_parts.append(
                    f"Measurement: {m_result.get('total_parts_found', 0)} parts processed"
                )
            if results.get('ets'):
                e_result = results['ets']
                success_parts.append(
                    f"ETS: {e_result.get('total_records', 0)} records processed"
                )
            
            success_message = "Successfully processed:\n" + "\n".join(success_parts)
            self.processing_complete.emit(success_message)
            
        except Exception as e:
            error_msg = f"Error during processing: {str(e)}"
            self.logger.error(error_msg)
            self.status_updated.emit(f"ERROR: {error_msg}")
            self.processing_error.emit(error_msg)


class UnifiedMainWindow(QMainWindow):
    """
    Main window for Unified File Processor
    Handles both Measurement and ETS data
    """
    
    def __init__(self):
        super().__init__()
        self.input_files = []
        self.measurement_output_path = ""
        self.ets_output_path = ""
        
        self.resource_manager = ResourceManager()
        self.logger = get_simple_logger(__name__)
        
        self.setWindowTitle("File Processor")
        self.setMinimumSize(500, 400)
        self.resize(700, 750)
        
        # Set icon
        app_icon = self.resource_manager.get_icon("pixel_logo.png")
        if app_icon:
            self.setWindowIcon(QIcon(app_icon))
        else:
            app_icon = self.style().standardIcon(self.style().SP_ComputerIcon)
            self.setWindowIcon(app_icon)
        
        self.setup_ui()
        self.load_main_stylesheet()
        
        self.logger.info("Application started")
    
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
        content_widget.setMinimumSize(480, 600)
        
        # Title
        title = QLabel("File Processor")
        title.setAlignment(Qt.AlignCenter)
        title.setObjectName("titleLabel")
        title.setWordWrap(True)
        
        font = QFont()
        font.setPointSize(18)
        font.setBold(True)
        title.setFont(font)
        main_layout.addWidget(title)
        
        # MODE SELECTOR - NEW!
        mode_group = QGroupBox("Processing Mode")
        mode_group.setObjectName("modeGroup")
        mode_layout = QVBoxLayout()
        
        mode_label = QLabel("Select the type of data you will be processing:")
        mode_layout.addWidget(mode_label)
        
        self.mode_selector = QComboBox()
        self.mode_selector.addItem("Measurement Data Only", UnifiedProcessor.MEASUREMENT_ONLY)
        self.mode_selector.addItem("ETS Data Only", UnifiedProcessor.ETS_ONLY)
        self.mode_selector.addItem("Measurement and ETS Data", UnifiedProcessor.BOTH)
        self.mode_selector.currentIndexChanged.connect(self.on_mode_changed)
        self.mode_selector.setMinimumHeight(35)
        
        mode_layout.addWidget(self.mode_selector)
        mode_group.setLayout(mode_layout)
        main_layout.addWidget(mode_group)
        
        # File drop area
        file_group = QGroupBox("Input Files")
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
        
        # Configuration fields - DYNAMIC!
        self.config_group = QGroupBox("Configuration")
        self.config_group.setObjectName("configurationGroup")
        self.config_layout = QVBoxLayout()
        
        # Measurement fields container
        self.measurement_fields_widget = QWidget()
        self.measurement_fields_layout = QGridLayout()
        self.measurement_fields = {}
        
        measurement_field_info = [
            ('part_name', 'Part Name:'),
            ('revision_number', 'Revision Number:'),
            ('lot_number', 'Lot Number:'),
            ('customer_p/n', 'Customer P/N:'),
            ('customer_po', 'Customer PO:'),
            ('measurement_units', 'Measurement Units:')
        ]
        
        measurement_placeholders = {
            'part_name': 'Enter part name...',
            'revision_number': 'Enter revision number...',
            'lot_number': 'Enter lot number...',
            'customer_p/n': 'Enter customer P/N...',
            'customer_po': 'Enter customer PO...',
            'measurement_units': 'Enter units (e.g., Inches, Millimeters)...'
        }
        
        for i, (field_key, label_text) in enumerate(measurement_field_info):
            label_widget = QLabel(label_text)
            label_widget.setObjectName(f"measurementLabel{i}")
            
            line_edit = QLineEdit()
            line_edit.setObjectName(f"measurementInput{i}")
            line_edit.setPlaceholderText(measurement_placeholders.get(field_key, ''))
            
            self.measurement_fields_layout.addWidget(label_widget, i, 0)
            self.measurement_fields_layout.addWidget(line_edit, i, 1)
            self.measurement_fields[field_key] = line_edit
        
        self.measurement_fields_layout.setColumnStretch(1, 1)
        self.measurement_fields_widget.setLayout(self.measurement_fields_layout)
        
        # ETS fields container
        self.ets_fields_widget = QWidget()
        self.ets_fields_layout = QGridLayout()
        self.ets_fields = {}
        
        ets_field_info = [
            ('part_name', 'Part Name:'),
            ('revision_number', 'Revision Number:'),
            ('customer_pn', 'Customer P/N:'),
            ('customer_po', 'Customer PO:'),
            ('ocv_min', 'OCV Min:'),
            ('ccv_min', 'CCV Min:'),
        ]
        
        ets_placeholders = {
            'part_name': 'Enter part name (e.g., 3B0036-AN-03)...',
            'revision_number': 'Enter revision number (e.g., 002)...',
            'customer_pn': 'Enter customer P/N (e.g., 01-5032-03)...',
            'customer_po': 'Enter customer PO (e.g., PO71906)...',
            'ocv_min': 'Enter OCV minimum value (e.g., 3.92)...',
            'ccv_min': 'Enter CCV minimum value (e.g., 3)...',
        }
        
        for i, (field_key, label_text) in enumerate(ets_field_info):
            label_widget = QLabel(label_text)
            label_widget.setObjectName(f"etsLabel{i}")
            
            line_edit = QLineEdit()
            line_edit.setObjectName(f"etsInput{i}")
            line_edit.setPlaceholderText(ets_placeholders.get(field_key, ''))
            
            # Set default values for OCV/CCV min
            if field_key == 'ocv_min':
                line_edit.setText('3.92')
            elif field_key == 'ccv_min':
                line_edit.setText('3')
            
            self.ets_fields_layout.addWidget(label_widget, i, 0)
            self.ets_fields_layout.addWidget(line_edit, i, 1)
            self.ets_fields[field_key] = line_edit
        
        self.ets_fields_layout.setColumnStretch(1, 1)
        self.ets_fields_widget.setLayout(self.ets_fields_layout)
        
        # Add both to config layout
        self.config_layout.addWidget(self.measurement_fields_widget)
        self.config_layout.addWidget(self.ets_fields_widget)
        self.config_group.setLayout(self.config_layout)
        main_layout.addWidget(self.config_group)
        
        # Output section - DYNAMIC!
        self.output_group = QGroupBox("Output Location")
        self.output_group.setObjectName("outputGroup")
        self.output_layout = QVBoxLayout()
        
        # Measurement output
        self.measurement_output_widget = QWidget()
        measurement_output_layout = QHBoxLayout()
        
        self.measurement_output_label = QLabel("No output location selected")
        self.measurement_output_label.setObjectName("outputPathLabel")
        self.measurement_output_label.setStyleSheet("color: #666; font-style: italic;")
        self.measurement_output_label.setWordWrap(True)
        
        self.measurement_save_btn = QPushButton("Choose Measurement Output")
        self.measurement_save_btn.setObjectName("saveAsButton")
        self.measurement_save_btn.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        save_icon = self.style().standardIcon(self.style().SP_DialogSaveButton)
        self.measurement_save_btn.setIcon(save_icon)
        self.measurement_save_btn.clicked.connect(self.choose_measurement_output)
        
        measurement_output_layout.addWidget(QLabel("Measurement Output:"))
        measurement_output_layout.addWidget(self.measurement_output_label, 1)
        measurement_output_layout.addWidget(self.measurement_save_btn)
        self.measurement_output_widget.setLayout(measurement_output_layout)
        
        # ETS output
        self.ets_output_widget = QWidget()
        ets_output_layout = QHBoxLayout()
        
        self.ets_output_label = QLabel("No output location selected")
        self.ets_output_label.setObjectName("outputPathLabel")
        self.ets_output_label.setStyleSheet("color: #666; font-style: italic;")
        self.ets_output_label.setWordWrap(True)
        
        self.ets_save_btn = QPushButton("Choose ETS Output")
        self.ets_save_btn.setObjectName("saveAsButton")
        self.ets_save_btn.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.ets_save_btn.setIcon(save_icon)
        self.ets_save_btn.clicked.connect(self.choose_ets_output)
        
        ets_output_layout.addWidget(QLabel("ETS Output:"))
        ets_output_layout.addWidget(self.ets_output_label, 1)
        ets_output_layout.addWidget(self.ets_save_btn)
        self.ets_output_widget.setLayout(ets_output_layout)
        
        # Add both outputs to output layout
        self.output_layout.addWidget(self.measurement_output_widget)
        self.output_layout.addWidget(self.ets_output_widget)
        self.output_group.setLayout(self.output_layout)
        main_layout.addWidget(self.output_group)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # Status label
        self.status_label = QLabel("Ready - Select processing mode and input files")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        main_layout.addWidget(self.status_label)
        
        main_layout.addStretch()
        
        # Process button
        self.process_btn = QPushButton("Process Files")
        self.process_btn.setObjectName("processButton")
        self.process_btn.setMinimumHeight(45)
        self.process_btn.setEnabled(False)
        self.process_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        process_icon = self.style().standardIcon(self.style().SP_DialogApplyButton)
        self.process_btn.setIcon(process_icon)
        self.process_btn.clicked.connect(self.process_files)
        
        main_layout.addWidget(self.process_btn)
        
        # Status bar
        self.statusBar().showMessage("Ready - Select processing mode and configure settings")
        
        # Initialize UI state
        self.on_mode_changed(0)
    
    def on_mode_changed(self, index):
        """Handle mode selector change"""
        mode = self.mode_selector.currentData()
        self.logger.info(f"Mode changed to: {mode}")
        
        # Show/hide appropriate configuration fields
        if mode == UnifiedProcessor.MEASUREMENT_ONLY:
            self.measurement_fields_widget.setVisible(True)
            self.ets_fields_widget.setVisible(False)
            self.measurement_output_widget.setVisible(True)
            self.ets_output_widget.setVisible(False)
            
        elif mode == UnifiedProcessor.ETS_ONLY:
            self.measurement_fields_widget.setVisible(False)
            self.ets_fields_widget.setVisible(True)
            self.measurement_output_widget.setVisible(False)
            self.ets_output_widget.setVisible(True)
            
        else:  # BOTH
            self.measurement_fields_widget.setVisible(True)
            self.ets_fields_widget.setVisible(True)
            self.measurement_output_widget.setVisible(True)
            self.ets_output_widget.setVisible(True)
        
        # Reset output paths when mode changes
        self.measurement_output_path = ""
        self.ets_output_path = ""
        self.measurement_output_label.setText("No output location selected")
        self.measurement_output_label.setStyleSheet("color: #666; font-style: italic;")
        self.ets_output_label.setText("No output location selected")
        self.ets_output_label.setStyleSheet("color: #666; font-style: italic;")
        
        self.update_process_button_state()
    
    def on_files_selected(self, files):
        self.input_files = files
        self.logger.info(f"Selected {len(files)} input files")
        self.update_process_button_state()
        self.statusBar().showMessage(f"Selected {len(files)} file(s)")
    
    def choose_measurement_output(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Measurement Output As",
            "measurement_processed.xlsx",
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
            
            self.measurement_output_path = file_path
            self.measurement_output_label.setText(os.path.basename(file_path))
            self.measurement_output_label.setStyleSheet("color: #000;")
            self.logger.info(f"Measurement output path set to: {file_path}")
            self.update_process_button_state()
    
    def choose_ets_output(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save ETS Output As",
            "ets_processed.xlsx",
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
            
            self.ets_output_path = file_path
            self.ets_output_label.setText(os.path.basename(file_path))
            self.ets_output_label.setStyleSheet("color: #000;")
            self.logger.info(f"ETS output path set to: {file_path}")
            self.update_process_button_state()
    
    def update_process_button_state(self):
        """Enable process button based on mode and required inputs"""
        mode = self.mode_selector.currentData()
        
        can_process = bool(self.input_files)
        
        if mode == UnifiedProcessor.MEASUREMENT_ONLY:
            can_process = can_process and bool(self.measurement_output_path)
        elif mode == UnifiedProcessor.ETS_ONLY:
            can_process = can_process and bool(self.ets_output_path)
        else:  # BOTH
            can_process = can_process and bool(self.measurement_output_path) and bool(self.ets_output_path)
        
        self.process_btn.setEnabled(can_process)
        
        if can_process:
            active_button_style = self.resource_manager.get_stylesheet("buttons_active.qss")
            if active_button_style:
                self.process_btn.setStyleSheet(active_button_style)
        else:
            disabled_button_style = self.resource_manager.get_stylesheet("buttons_disabled.qss")
            if disabled_button_style:
                self.process_btn.setStyleSheet(disabled_button_style)
    
    def get_measurement_config(self):
        """Get measurement configuration values"""
        values = {}
        for key, field in self.measurement_fields.items():
            values[key] = field.text().strip()
        return values
    
    def get_ets_config(self):
        """Get ETS configuration values"""
        values = {}
        for key, field in self.ets_fields.items():
            values[key] = field.text().strip()
        return values
    
    def process_files(self):
        if not self.input_files:
            QMessageBox.warning(self, "No Files", "Please select input files first.")
            return
        
        mode = self.mode_selector.currentData()
        measurement_config = self.get_measurement_config()
        ets_config = self.get_ets_config()
        
        # Validate based on mode
        validation_errors = UnifiedProcessor.validate_config(mode, measurement_config, ets_config)
        if validation_errors:
            QMessageBox.warning(self, "Invalid Configuration",
                              "Please check the following fields:\n" + "\n".join(validation_errors))
            return
        
        # Check outputs
        if mode == UnifiedProcessor.MEASUREMENT_ONLY and not self.measurement_output_path:
            QMessageBox.warning(self, "No Output", "Please choose measurement output location.")
            return
        if mode == UnifiedProcessor.ETS_ONLY and not self.ets_output_path:
            QMessageBox.warning(self, "No Output", "Please choose ETS output location.")
            return
        if mode == UnifiedProcessor.BOTH:
            if not self.measurement_output_path or not self.ets_output_path:
                QMessageBox.warning(self, "Missing Outputs",
                                  "Please choose both measurement and ETS output locations.")
                return
        
        self.logger.info(f"Starting processing in {mode} mode...")
        
        # Disable UI during processing
        self.process_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.statusBar().showMessage("Processing files...")
        
        # Start processing thread
        self.processing_thread = ProcessingThread(
            mode,
            self.input_files,
            measurement_config,
            ets_config,
            self.measurement_output_path,
            self.ets_output_path
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
        
        self.logger.info("Processing completed successfully")
        QMessageBox.information(self, "Success", message)
        
        # Ask if user wants to open the output files
        mode = self.mode_selector.currentData()
        
        if mode == UnifiedProcessor.MEASUREMENT_ONLY:
            self.offer_to_open_file(self.measurement_output_path, "measurement")
        elif mode == UnifiedProcessor.ETS_ONLY:
            self.offer_to_open_file(self.ets_output_path, "ETS")
        else:  # BOTH
            reply = QMessageBox.question(
                self,
                "Open Files?",
                "Would you like to open the generated files?",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.open_file(self.measurement_output_path)
                self.open_file(self.ets_output_path)
    
    def offer_to_open_file(self, file_path, file_type):
        """Offer to open a single file"""
        reply = QMessageBox.question(
            self,
            "Open File?",
            f"Would you like to open the {file_type} output file?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.open_file(file_path)
    
    def open_file(self, file_path):
        """Open a file in the default application"""
        try:
            from src.core.file_handler import FileHandler
            if FileHandler.open_file(file_path):
                self.logger.info(f"Opened output file: {file_path}")
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
        
        self.logger.error(f"Processing failed: {error_message}")
        QMessageBox.critical(self, "Processing Error", error_message)


def main():
    logger = get_simple_logger("unified_processor")
    logger.info("Starting  File Processor application")
    
    app = QApplication(sys.argv)
    
    app.setApplicationName("File Processor")
    app.setApplicationVersion("2.0")
    
    try:
        resource_manager = ResourceManager()
        
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
        
        window = UnifiedMainWindow()
        window.show()
        
        logger.info("Main window displayed successfully")
        exit_code = app.exec_()
        logger.info(f"Application exited with code: {exit_code}")
        sys.exit(exit_code)
        
    except Exception as e:
        logger.error(f"Critical error in application: {e}")
        QMessageBox.critical(None, "Critical Error",
                           f"A critical error occurred:\n{str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()