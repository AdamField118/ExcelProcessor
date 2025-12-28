"""
Excel File Processor - Main GUI Window

This module contains the PyQt5 GUI for the Excel file processing application.
"""

import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QWidget, QPushButton, QLabel, QLineEdit, QTextEdit, 
                             QFileDialog, QMessageBox, QProgressBar, QGroupBox,
                             QGridLayout, QFrame, QScrollArea, QSizePolicy)
from PyQt5.QtCore import Qt, QMimeData, QThread, pyqtSignal
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QFont, QPalette, QIcon, QPixmap

from core.base_processor import BaseProcessor
from src.core.file_handler import FileHandler
from src.utils.validators import FileValidator
from src.utils.logger import setup_logger, get_simple_logger
from src.utils.resource_manager import ResourceManager


class FileDropArea(QFrame):
    """Custom widget for drag and drop file functionality"""
    files_dropped = pyqtSignal(list)
    
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setMinimumHeight(150)
        self.setFrameStyle(QFrame.StyledPanel)
        self.setLineWidth(2)
        
        self.resource_manager = ResourceManager()
        
        self.setup_ui()
        self.load_styles()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Main label with icon
        label_layout = QHBoxLayout()
        
        upload_icon = QLabel()
        # Use Qt's built-in up arrow icon
        up_icon = self.style().standardIcon(self.style().SP_ArrowUp)
        upload_icon.setPixmap(up_icon.pixmap(32, 32))
        
        # Mention both Excel and CSV files
        self.main_label = QLabel("Drop Excel or CSV files here or click Browse")
        self.main_label.setAlignment(Qt.AlignCenter)
        
        font_config = self.resource_manager.get_config("fonts.main_label")
        font = QFont()
        font.setPointSize(font_config.get("size", 12) if font_config else 12)
        font.setBold(font_config.get("bold", False) if font_config else False)
        self.main_label.setFont(font)
        
        label_layout.addWidget(upload_icon)
        label_layout.addWidget(self.main_label)
        label_layout.setAlignment(Qt.AlignCenter)
        
        # Browse button with Qt's built-in folder icon
        self.browse_btn = QPushButton("Browse Files")
        folder_icon = self.style().standardIcon(self.style().SP_DirIcon)
        self.browse_btn.setIcon(folder_icon)
        
        self.browse_btn.setMaximumWidth(150)
        self.browse_btn.setMinimumHeight(45)
        self.browse_btn.clicked.connect(self.browse_files)
        
        # File list display - KEEP ORIGINAL SIZING
        self.file_list = QTextEdit()
        self.browse_btn.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.file_list.setMinimumHeight(50)
        self.file_list.setMaximumHeight(100)
        placeholder_text = self.resource_manager.get_config("ui.placeholders.file_list", "Selected files will appear here...")
        self.file_list.setPlaceholderText(placeholder_text)
        self.file_list.setReadOnly(True)
        
        layout.addLayout(label_layout)
        layout.addWidget(self.browse_btn, alignment=Qt.AlignCenter)
        layout.addWidget(QLabel("Selected Files:"))
        layout.addWidget(self.file_list)
        
        self.setLayout(layout)
    
    def load_styles(self):
        """Load stylesheet from external file"""
        stylesheet = self.resource_manager.get_stylesheet("drop_area.qss")
        if stylesheet:
            self.setStyleSheet(stylesheet)
        else:
            # Fallback inline styles (temporary)
            self.setStyleSheet("""
                FileDropArea {
                    border: 2px dashed #aaa;
                    border-radius: 10px;
                    background-color: #f9f9f9;
                }
                FileDropArea:hover {
                    border-color: #2196F3;
                    background-color: #e3f2fd;
                }
            """)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            active_stylesheet = self.resource_manager.get_stylesheet("drop_area_active.qss")
            if active_stylesheet:
                self.setStyleSheet(active_stylesheet)
            else:
                # Fallback inline style
                self.setStyleSheet("""
                    FileDropArea {
                        border: 2px dashed #2196F3;
                        border-radius: 10px;
                        background-color: #e3f2fd;
                    }
                """)
    
    def dragLeaveEvent(self, event):
        self.load_styles()
    
    def dropEvent(self, event: QDropEvent):
        files = []
        invalid_files = []
        
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            # UPDATED: Use new validation method that supports both Excel and CSV
            if FileValidator.is_valid_data_file(file_path):
                files.append(file_path)
            else:
                invalid_files.append(os.path.basename(file_path))
        
        if files:
            self.files_dropped.emit(files)
            self.update_file_list(files)
            
            if invalid_files:
                # UPDATED: Error message mentions both file types
                QMessageBox.warning(
                    self, 
                    "Some Invalid Files", 
                    f"The following files were skipped (not valid Excel or CSV files):\n" + 
                    "\n".join(invalid_files)
                )
        else:
            # UPDATED: Error message mentions both file types
            QMessageBox.warning(self, "Invalid Files", 
                              "Please drop only valid Excel (.xls, .xlsx) or CSV (.csv) files")
        
        # Reset style
        self.dragLeaveEvent(None)
    
    def browse_files(self):
        # UPDATED: File dialog includes CSV files
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "Select Excel or CSV Files",
            "",
            "Data Files (*.xlsx *.xls *.csv);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;All Files (*)"
        )
        
        if files:
            valid_files = []
            invalid_files = []
            
            for file_path in files:
                # UPDATED: Use new validation method
                if FileValidator.is_valid_data_file(file_path):
                    valid_files.append(file_path)
                else:
                    invalid_files.append(os.path.basename(file_path))
            
            if valid_files:
                self.files_dropped.emit(valid_files)
                self.update_file_list(valid_files)
                
                if invalid_files:
                    QMessageBox.warning(
                        self, 
                        "Some Invalid Files", 
                        f"The following files were skipped:\n" + 
                        "\n".join(invalid_files)
                    )
            else:
                # UPDATED: Error message mentions both file types
                QMessageBox.warning(self, "No Valid Files", 
                                  "No valid Excel or CSV files were selected.")
    
    def update_file_list(self, files):
        # UPDATED: Show file type next to filename
        file_names_with_types = []
        for f in files:
            file_name = os.path.basename(f)
            file_type = FileValidator.get_file_type(f)
            file_names_with_types.append(f"{file_name} ({file_type})")
        
        self.file_list.setText('\n'.join(file_names_with_types))


class ProcessingThread(QThread):
    """Optimized background thread for processing files with detailed progress"""
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)  # New signal for status text
    processing_complete = pyqtSignal(str)  # Success message
    processing_error = pyqtSignal(str)     # Error message
    
    def __init__(self, input_files, string_values, output_path):
        super().__init__()
        self.input_files = input_files
        self.string_values = string_values
        self.output_path = output_path
        # Use simple logger that won't create files in production
        self.logger = get_simple_logger(__name__)
    
    def run(self):
        try:
            self.logger.info(f"Starting processing of {len(self.input_files)} files")
            
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
                # UPDATED: Show file type in status
                file_type = FileValidator.get_file_type(file_path)
                self.status_updated.emit(f"Loading file {i+1}/{total_files}: {os.path.basename(file_path)} ({file_type})")
                
                # ORIGINAL: Using original load_excel_file method (now supports CSV internally)
                file_data = FileHandler.load_excel_file(file_path)
                loaded_data.append(file_data)
                
                progress = 10 + int((i + 1) / total_files * 20)
                self.progress_updated.emit(progress)
            
            # Processing phase (70% of progress) - ORIGINAL LOGIC UNCHANGED
            self.status_updated.emit("Processing data...")
            self.progress_updated.emit(30)
            
            # Convert GUI field names to match FileHandler expectations
            processed_string_values = self.convert_gui_fields_to_filehandler_format(self.string_values)
            
            # Create progress callback that updates GUI
            def progress_callback(message):
                self.status_updated.emit(message)
                # Extract percentage if available and update progress
                if "%" in message:
                    try:
                        # Extract percentage from message like "Processed 50/100 parts (50%)"
                        percent_part = message.split("(")[1].split("%")[0]
                        percent = int(percent_part)
                        # Map to 30-95% range (save 5% for completion)
                        gui_progress = 30 + int(percent * 0.65)
                        self.progress_updated.emit(gui_progress)
                    except:
                        pass  # If parsing fails, just continue
            
            # Use optimized ExcelProcessor with progress callback
            result = ExcelProcessor.process_files(
                processed_string_values, 
                loaded_data, 
                self.output_path,
                progress_callback
            )
            
            self.progress_updated.emit(100)
            self.status_updated.emit("Processing complete!")
            
            self.logger.info("Processing completed successfully")
            success_message = (f"Successfully processed {result['total_files_processed']} files "
                             f"containing {result['total_parts_found']} parts and saved to "
                             f"{os.path.basename(self.output_path)}!")
            self.processing_complete.emit(success_message)
            
        except Exception as e:
            error_msg = f"Error during processing: {str(e)}"
            self.logger.error(error_msg)
            self.status_updated.emit(f"ERROR: {error_msg}")
            self.processing_error.emit(error_msg)
    
    def convert_gui_fields_to_filehandler_format(self, gui_fields):
        """
        Convert GUI field names to match FileHandler expected format
        """
        field_mapping = {
            'part_name': 'part_name',
            'revision_number': 'revision_number', 
            'lot_number': 'lot_number',
            'customer_p/n': 'customer_p/n',
            'customer_po': 'customer_po',
            'measurement_units': 'measurement_units'
        }
        
        converted = {}
        for gui_key, value in gui_fields.items():
            # Map GUI field names to FileHandler expected names
            filehandler_key = field_mapping.get(gui_key, gui_key)
            converted[filehandler_key] = value
        
        return converted


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.input_files = []
        self.output_path = ""
        
        self.resource_manager = ResourceManager()
        # Use simple logger that won't create files in production
        self.logger = get_simple_logger(__name__)
        
        window_config = self.resource_manager.get_config("window")
        
        # UPDATED: Window title mentions CSV support
        window_title = "Excel & CSV File Processor"
        if window_config:
            window_title = window_config.get("title", window_title)
        
        self.setWindowTitle(window_title)
        
        self.setMinimumSize(450, 400)
        
        # Set a reasonable default size
        self.resize(650, 650)

        # Use Qt's built-in application icon or keep the custom one
        app_icon = self.resource_manager.get_icon("pixel_logo.png")
        if app_icon:
            self.setWindowIcon(QIcon(app_icon))
        else:
            # Fallback to Qt's built-in application icon
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
        # Create scroll area as the main container
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        # Create the actual content widget
        content_widget = QWidget()
        scroll_area.setWidget(content_widget)
        
        # Set scroll area as central widget
        self.setCentralWidget(scroll_area)
        
        # Use content_widget for the layout
        main_layout = QVBoxLayout()
        content_widget.setLayout(main_layout)
        
        # Set a minimum size for the content widget (this determines when scrolling kicks in)
        content_widget.setMinimumSize(450, 540)  # Minimum size before scrolling
        
        # Title with styling from config
        title_config = self.resource_manager.get_config("ui.title")
        
        # UPDATED: Title mentions CSV support
        title_text = "Excel & CSV File Processor"
        if title_config:
            title_text = title_config.get("text", title_text)
        
        title = QLabel(title_text)
        title.setAlignment(Qt.AlignCenter)
        title.setObjectName("titleLabel")  # For CSS styling
        title.setWordWrap(True)  # Allow text wrapping
        
        font = QFont()
        if title_config:
            font.setPointSize(title_config.get("font_size", 16))
            font.setBold(title_config.get("bold", True))
        else:
            font.setPointSize(16)
            font.setBold(True)
        title.setFont(font)
        main_layout.addWidget(title)
        
        # File drop area
        labels_config = self.resource_manager.get_config("ui.labels")
        
        # UPDATED: Group box title mentions CSV support
        input_files_label = "Input Files (Excel & CSV)"
        if labels_config:
            input_files_label = labels_config.get("input_files", input_files_label)
        
        file_group = QGroupBox(input_files_label)
        file_group.setObjectName("inputFilesGroup")  # For CSS styling
        file_layout = QVBoxLayout()
        
        self.drop_area = FileDropArea()
        # Make drop area responsive but with reasonable bounds
        self.drop_area.setMinimumHeight(100)
        self.drop_area.setMaximumHeight(200)  # Cap the maximum height
        self.drop_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        
        self.drop_area.files_dropped.connect(self.on_files_selected)
        file_layout.addWidget(self.drop_area)
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)
        
        # String input fields
        configuration_label = "Part Information"
        if labels_config:
            configuration_label = labels_config.get("configuration", configuration_label)
        
        string_group = QGroupBox(configuration_label)
        string_group.setObjectName("configurationGroup")
        string_layout = QGridLayout()
        
        self.string_fields = {}
        
        # Updated field setup with proper keys
        field_info = [
            ('part_name', 'Part Name:'),
            ('revision_number', 'Revision Number:'),
            ('lot_number', 'Lot Number:'),
            ('customer_p/n', 'Customer P/N:'),
            ('customer_po', 'Customer PO:'),
            ('measurement_units', 'Measurement Units:')
        ]
        
        for i, (field_key, label_text) in enumerate(field_info):
            label_widget = QLabel(label_text)
            label_widget.setObjectName(f"fieldLabel{i}")
            
            line_edit = QLineEdit()
            line_edit.setObjectName(f"fieldInput{i}")
            
            # Set appropriate placeholder text
            placeholders = {
                'part_name': 'Enter part name...',
                'revision_number': 'Enter revision number...',
                'lot_number': 'Enter lot number...',
                'customer_p/n': 'Enter customer part number...',
                'customer_po': 'Enter PO number...',
                'measurement_units': 'Enter measurement units (e.g., mm, inches)...'
            }
            
            line_edit.setPlaceholderText(placeholders.get(field_key, f"Enter {label_text.lower().replace(':', '')}..."))
            
            string_layout.addWidget(label_widget, i, 0)
            string_layout.addWidget(line_edit, i, 1)
            
            # Store reference to the field
            self.string_fields[field_key] = line_edit
        
        # Make form layout more responsive
        string_layout.setColumnStretch(1, 1)  # Make input column expandable
        
        string_group.setLayout(string_layout)
        main_layout.addWidget(string_group)
        
        # Output section
        output_label = "Output"
        if labels_config:
            output_label = labels_config.get("output", output_label)
        
        output_group = QGroupBox(output_label)
        output_group.setObjectName("outputGroup")  # For CSS styling
        output_layout = QVBoxLayout()
        
        # Output path display and selection
        path_layout = QHBoxLayout()
        
        output_config = self.resource_manager.get_config("ui.output")
        default_message = "No output location selected"
        if output_config:
            default_message = output_config.get("no_location_message", default_message)
        
        self.output_path_label = QLabel(default_message)
        self.output_path_label.setObjectName("outputPathLabel")  # For CSS styling
        self.output_path_label.setStyleSheet("color: #666; font-style: italic;")
        self.output_path_label.setWordWrap(True)  # Allow text wrapping
        
        button_text = "Choose Save Location"
        if output_config:
            button_text = output_config.get("button_text", button_text)
        
        self.save_as_btn = QPushButton(button_text)
        self.save_as_btn.setObjectName("saveAsButton")  # For CSS styling
        self.save_as_btn.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        
        # Use Qt's built-in save icon
        save_icon = self.style().standardIcon(self.style().SP_DialogSaveButton)
        self.save_as_btn.setIcon(save_icon)
        
        self.save_as_btn.clicked.connect(self.choose_output_location)
        
        path_layout.addWidget(QLabel("Output File:"))
        path_layout.addWidget(self.output_path_label, 1)
        path_layout.addWidget(self.save_as_btn)
        output_layout.addLayout(path_layout)
        
        output_group.setLayout(output_layout)
        main_layout.addWidget(output_group)
        
        # Progress bar - ORIGINAL LOGIC UNCHANGED
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")  # For CSS styling
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # Status label for detailed progress updates
        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)  # Allow text wrapping
        main_layout.addWidget(self.status_label)
        
        # Add stretch to push process button to bottom when window is large
        main_layout.addStretch()
        
        # Process button
        button_config = self.resource_manager.get_config("ui.buttons.process")
        
        process_button_text = "Process Files"
        if button_config:
            process_button_text = button_config.get("text", process_button_text)
        
        self.process_btn = QPushButton(process_button_text)
        self.process_btn.setObjectName("processButton")  # For CSS styling
        
        min_height = 35  # Reduce from 40
        if button_config:
            min_height = button_config.get("min_height", min_height)
        
        self.process_btn.setMinimumHeight(min_height)
        self.process_btn.setEnabled(False)
        self.process_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        # Use Qt's built-in apply/process icon
        process_icon = self.style().standardIcon(self.style().SP_DialogApplyButton)
        self.process_btn.setIcon(process_icon)
        
        self.process_btn.clicked.connect(self.process_files)
        main_layout.addWidget(self.process_btn)
        
        # Status bar
        status_config = self.resource_manager.get_config("ui.status_messages")
        ready_message = "Ready - Select input files and configure settings"
        if status_config:
            ready_message = status_config.get("ready", ready_message)
        
        self.statusBar().showMessage(ready_message)
    
    # ALL METHODS BELOW ARE ORIGINAL LOGIC - UNCHANGED
    def on_files_selected(self, files):
        self.input_files = files
        self.logger.info(f"Selected {len(files)} input files")  
        self.update_process_button_state()
        self.statusBar().showMessage(f"Selected {len(files)} file(s)")
    
    def choose_output_location(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Processed File As",
            "processed_data.xlsx",
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
        # Enable process button only if we have files and output location
        can_process = bool(self.input_files and self.output_path)
        self.process_btn.setEnabled(can_process)
        
        if can_process:
            active_button_style = self.resource_manager.get_stylesheet("buttons_active.qss")
            if active_button_style:
                self.process_btn.setStyleSheet(active_button_style)
            else:
                # Fallback inline style
                self.process_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #4CAF50;
                        color: white;
                        font-weight: bold;
                        border: none;
                        border-radius: 5px;
                    }
                    QPushButton:hover {
                        background-color: #45a049;
                    }
                """)
        else:
            disabled_button_style = self.resource_manager.get_stylesheet("buttons_disabled.qss")
            if disabled_button_style:
                self.process_btn.setStyleSheet(disabled_button_style)
            else:
                self.process_btn.setStyleSheet("")  # Reset to default
    
    def get_string_values(self):
        """Get all string field values as a dictionary"""
        values = {}
        
        # Map GUI field keys to the actual field names expected by FileHandler
        field_mapping = {
            'part_name': 'part_name',
            'revision_number': 'revision_number',
            'lot_number': 'lot_number', 
            'customer_p/n': 'customer_p/n',
            'customer_po': 'customer_po',  # This will be handled specially
            'measurement_units': 'measurement_units'
        }
        
        for gui_key, field in self.string_fields.items():
            if gui_key in field_mapping:
                filehandler_key = field_mapping[gui_key]
                field_value = field.text().strip()
                values[filehandler_key] = field_value
        
        return values
    
    def process_files(self):
        # Validate that we have all required inputs
        if not self.input_files:
            QMessageBox.warning(self, "No Files", "Please select input files first.")
            return
            
        if not self.output_path:
            QMessageBox.warning(self, "No Output Location", 
                              "Please choose where to save the output file.")
            return
        
        # Get and validate string values
        string_values = self.get_string_values()
        
        validation_errors = FileValidator.validate_string_inputs(string_values)
        if validation_errors:
            QMessageBox.warning(
                self, 
                "Invalid Configuration", 
                "Please check the following fields:\n" + "\n".join(validation_errors)
            )
            return
        
        self.logger.info("Starting file processing...")  
        
        # Disable UI during processing
        self.process_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.statusBar().showMessage("Processing files...")
        
        # Start processing in background thread
        self.processing_thread = ProcessingThread(
            self.input_files, 
            string_values, 
            self.output_path
        )
        self.processing_thread.progress_updated.connect(self.progress_bar.setValue)
        self.processing_thread.status_updated.connect(self.status_label.setText)  # Connect the new status signal
        self.processing_thread.processing_complete.connect(self.on_processing_complete)
        self.processing_thread.processing_error.connect(self.on_processing_error)
        self.processing_thread.start()
    
    def on_processing_complete(self, message):
        self.progress_bar.setVisible(False)
        self.status_label.setText("Processing complete!")  # Update status label
        self.process_btn.setEnabled(True)
        self.statusBar().showMessage("Processing complete!")
        
        self.logger.info("Processing completed successfully")
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
        self.status_label.setText("Processing failed!")  # Update status label
        self.process_btn.setEnabled(True)
        self.statusBar().showMessage("Processing failed!")
        
        self.logger.error(f"Processing failed: {error_message}")
        QMessageBox.critical(self, "Processing Error", error_message)


def main():
    # Use simple logger that won't create files in production
    logger = get_simple_logger("excel_processor")
    logger.info("Starting Excel File Processor application")
    
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("Excel File Processor")
    app.setApplicationVersion("1.0")
    
    try:
        resource_manager = ResourceManager()
        
        global_stylesheet = resource_manager.get_stylesheet("app_global.qss")
        if global_stylesheet:
            app.setStyleSheet(global_stylesheet)
            logger.info("Loaded global application stylesheet")
        
        app_icon_path = resource_manager.get_icon_path("pixel_logo.png")
        if app_icon_path and os.path.exists(app_icon_path):
            app.setWindowIcon(QIcon(app_icon_path))
            logger.info("Set application icon")
        
        app_config = resource_manager.get_config("application")
        if app_config:
            logger.info("Loaded application configuration")
        
        window = MainWindow()
        window.show()
        
        logger.info("Main window displayed successfully")
        exit_code = app.exec_()
        logger.info(f"Application exited with code: {exit_code}")
        sys.exit(exit_code)
        
    except Exception as e:
        logger.error(f"Critical error in main application: {e}")
        QMessageBox.critical(None, "Critical Error", 
                           f"A critical error occurred:\n{str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()