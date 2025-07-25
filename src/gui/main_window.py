"""
Excel File Processor - Main GUI Window

This module contains the PyQt5 GUI for the Excel file processing application.

TODO: The following modules need to be implemented for this GUI to work:
- src/core/excel_processor.py (ExcelProcessor class)
- src/core/file_handler.py (FileHandler class)  
- src/utils/validators.py (FileValidator class)
- src/utils/logger.py (setup_logger function)
- src/utils/resource_manager.py (ResourceManager class)

TODO: The following resource files need to be created:
- resources/styles/main_window.qss (Main window stylesheet)
- resources/styles/drop_area.qss (File drop area stylesheet)
- resources/styles/buttons.qss (Button styles)
- resources/icons/app_icon.ico (Application icon)
- resources/icons/folder.png (Folder icon for buttons)
- resources/icons/process.png (Process button icon)
- resources/config/app_config.json (Application configuration)

During development, you can comment out the imports and TODO function calls
to test the GUI layout and interactions.
"""

import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QWidget, QPushButton, QLabel, QLineEdit, QTextEdit, 
                             QFileDialog, QMessageBox, QProgressBar, QGroupBox,
                             QGridLayout, QFrame)
from PyQt5.QtCore import Qt, QMimeData, QThread, pyqtSignal
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QFont, QPalette, QIcon, QPixmap

# Import project modules
from core.excel_processor import ExcelProcessor  # TODO: Create this class
from core.file_handler import FileHandler  # TODO: Create this class
from utils.validators import FileValidator  # TODO: Create this class
from utils.logger import setup_logger  # TODO: Create this function
from utils.resource_manager import ResourceManager  # TODO: Create this class


class FileDropArea(QFrame):
    """Custom widget for drag and drop file functionality"""
    files_dropped = pyqtSignal(list)
    
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setMinimumHeight(150)
        self.setFrameStyle(QFrame.StyledPanel)
        self.setLineWidth(2)
        
        # TODO: Initialize ResourceManager for accessing styles and icons
        self.resource_manager = ResourceManager()
        
        self.setup_ui()
        self.load_styles()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Main label with icon
        label_layout = QHBoxLayout()
        
        # TODO: Add icon from resources/icons/upload.png
        upload_icon = QLabel()
        upload_pixmap = self.resource_manager.get_icon("upload.png")  # TODO: Create upload.png icon
        if upload_pixmap:
            upload_icon.setPixmap(upload_pixmap.scaled(32, 32, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        
        self.main_label = QLabel("Drop Excel files here or click Browse")
        self.main_label.setAlignment(Qt.AlignCenter)
        
        # TODO: Get font settings from resources/config/app_config.json
        font_config = self.resource_manager.get_config("fonts.main_label")  # TODO: Create app_config.json
        font = QFont()
        font.setPointSize(font_config.get("size", 12))
        font.setBold(font_config.get("bold", False))
        self.main_label.setFont(font)
        
        label_layout.addWidget(upload_icon)
        label_layout.addWidget(self.main_label)
        label_layout.setAlignment(Qt.AlignCenter)
        
        # Browse button with icon
        self.browse_btn = QPushButton("Browse Files")
        # TODO: Add icon from resources/icons/folder.png
        folder_icon = self.resource_manager.get_icon("folder.png")  # TODO: Create folder.png icon
        if folder_icon:
            self.browse_btn.setIcon(QIcon(folder_icon))
        
        self.browse_btn.setMaximumWidth(150)
        self.browse_btn.clicked.connect(self.browse_files)
        
        # File list display
        self.file_list = QTextEdit()
        self.file_list.setMaximumHeight(80)
        # TODO: Get placeholder text from resources/config/app_config.json
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
        # TODO: Load stylesheet from resources/styles/drop_area.qss
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
            # TODO: Load drag-active stylesheet from resources/styles/drop_area_active.qss
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
        # TODO: Reload normal stylesheet
        self.load_styles()
    
    def dropEvent(self, event: QDropEvent):
        files = []
        invalid_files = []
        
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            # TODO: Use FileValidator.is_valid_excel_file(file_path)
            if FileValidator.is_valid_excel_file(file_path):
                files.append(file_path)
            else:
                invalid_files.append(os.path.basename(file_path))
        
        if files:
            self.files_dropped.emit(files)
            self.update_file_list(files)
            
            if invalid_files:
                QMessageBox.warning(
                    self, 
                    "Some Invalid Files", 
                    f"The following files were skipped (not valid Excel files):\n" + 
                    "\n".join(invalid_files)
                )
        else:
            QMessageBox.warning(self, "Invalid Files", 
                              "Please drop only valid Excel files (.xls, .xlsx)")
        
        # Reset style
        self.dragLeaveEvent(None)
    
    def browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "Select Excel Files",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if files:
            # TODO: Use FileValidator to validate selected files
            valid_files = []
            invalid_files = []
            
            for file_path in files:
                if FileValidator.is_valid_excel_file(file_path):
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
                QMessageBox.warning(self, "No Valid Files", 
                                  "No valid Excel files were selected.")
    
    def update_file_list(self, files):
        file_names = [os.path.basename(f) for f in files]
        self.file_list.setText('\n'.join(file_names))


class ProcessingThread(QThread):
    """Background thread for processing files"""
    progress_updated = pyqtSignal(int)
    processing_complete = pyqtSignal(str)  # Success message
    processing_error = pyqtSignal(str)     # Error message
    
    def __init__(self, input_files, string_values, output_path):
        super().__init__()
        self.input_files = input_files
        self.string_values = string_values
        self.output_path = output_path
        self.logger = setup_logger(__name__)  # TODO: Implement setup_logger function
    
    def run(self):
        try:
            self.logger.info(f"Starting processing of {len(self.input_files)} files")  # TODO: Logger
            
            # TODO: Initialize FileHandler and ExcelProcessor
            file_handler = FileHandler()
            excel_processor = ExcelProcessor()
            
            # TODO: Validate all input files before processing
            total_files = len(self.input_files)
            for i, file_path in enumerate(self.input_files):
                self.logger.info(f"Validating file: {file_path}")  # TODO: Logger
                
                # TODO: Use FileValidator.validate_file_accessibility(file_path)
                if not FileValidator.validate_file_accessibility(file_path):
                    raise Exception(f"Cannot access file: {file_path}")
                
                # Update progress for validation phase (0-20%)
                progress = int((i + 1) / total_files * 20)
                self.progress_updated.emit(progress)
            
            # TODO: Load and process files using ExcelProcessor
            self.logger.info("Loading input files...")  # TODO: Logger
            loaded_data = []
            
            for i, file_path in enumerate(self.input_files):
                self.logger.info(f"Loading file: {file_path}")  # TODO: Logger
                
                # TODO: Use FileHandler.load_excel_file(file_path)
                file_data = file_handler.load_excel_file(file_path)
                loaded_data.append(file_data)
                
                # Update progress for loading phase (20-60%)
                progress = 20 + int((i + 1) / total_files * 40)
                self.progress_updated.emit(progress)
            
            # TODO: Process data using ExcelProcessor
            self.logger.info("Processing data...")  # TODO: Logger
            self.progress_updated.emit(70)
            
            # TODO: Use ExcelProcessor.process_files(loaded_data, self.string_values)
            processed_data = excel_processor.process_files(loaded_data, self.string_values)
            
            self.progress_updated.emit(85)
            
            # TODO: Save output using FileHandler
            self.logger.info(f"Saving output to: {self.output_path}")  # TODO: Logger
            
            # TODO: Use FileHandler.save_excel_file(processed_data, self.output_path)
            file_handler.save_excel_file(processed_data, self.output_path)
            
            self.progress_updated.emit(100)
            
            self.logger.info("Processing completed successfully")  # TODO: Logger
            self.processing_complete.emit(f"Successfully processed {total_files} files and saved to {os.path.basename(self.output_path)}!")
            
        except Exception as e:
            error_msg = f"Error during processing: {str(e)}"
            self.logger.error(error_msg)  # TODO: Logger
            self.processing_error.emit(error_msg)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.input_files = []
        self.output_path = ""
        
        # TODO: Initialize ResourceManager and Logger
        self.resource_manager = ResourceManager()
        self.logger = setup_logger(__name__)  # TODO: Implement setup_logger function
        
        # TODO: Load window configuration from resources/config/app_config.json
        window_config = self.resource_manager.get_config("window")
        
        self.setWindowTitle(window_config.get("title", "Excel File Processor"))
        min_size = window_config.get("min_size", {"width": 600, "height": 500})
        self.setMinimumSize(min_size["width"], min_size["height"])
        
        # TODO: Set application icon from resources/icons/app_icon.ico
        app_icon = self.resource_manager.get_icon("app_icon.ico")
        if app_icon:
            self.setWindowIcon(QIcon(app_icon))
        
        self.setup_ui()
        self.load_main_stylesheet()
        
        self.logger.info("Application started")  # TODO: Logger
    
    def load_main_stylesheet(self):
        """Load main window stylesheet"""
        # TODO: Load main stylesheet from resources/styles/main_window.qss
        main_stylesheet = self.resource_manager.get_stylesheet("main_window.qss")
        if main_stylesheet:
            self.setStyleSheet(main_stylesheet)
        
    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        
        # Title with styling from config
        # TODO: Get title configuration from resources/config/app_config.json
        title_config = self.resource_manager.get_config("ui.title")
        title_text = title_config.get("text", "Excel File Processor")
        
        title = QLabel(title_text)
        title.setAlignment(Qt.AlignCenter)
        title.setObjectName("titleLabel")  # For CSS styling
        
        # TODO: Apply title font from config
        font = QFont()
        font.setPointSize(title_config.get("font_size", 16))
        font.setBold(title_config.get("bold", True))
        title.setFont(font)
        main_layout.addWidget(title)
        
        # File drop area
        # TODO: Get group box labels from resources/config/app_config.json
        labels_config = self.resource_manager.get_config("ui.labels")
        
        file_group = QGroupBox(labels_config.get("input_files", "Input Files"))
        file_group.setObjectName("inputFilesGroup")  # For CSS styling
        file_layout = QVBoxLayout()
        
        self.drop_area = FileDropArea()
        self.drop_area.files_dropped.connect(self.on_files_selected)
        file_layout.addWidget(self.drop_area)
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)
        
        # String input fields
        string_group = QGroupBox(labels_config.get("configuration", "Configuration Values"))
        string_group.setObjectName("configurationGroup")  # For CSS styling
        string_layout = QGridLayout()
        
        self.string_fields = {}
        
        # TODO: Get field labels from resources/config/app_config.json
        field_config = self.resource_manager.get_config("ui.input_fields")
        field_labels = field_config.get("labels", [
            "Project Name:",
            "Department:",
            "Analyst:",
            "Report Date:",
            "Version:"
        ])
        
        for i, label in enumerate(field_labels):
            label_widget = QLabel(label)
            label_widget.setObjectName(f"fieldLabel{i}")  # For CSS styling
            
            line_edit = QLineEdit()
            line_edit.setObjectName(f"fieldInput{i}")  # For CSS styling
            
            # TODO: Get placeholder text from config
            field_key = label.replace(':', '').replace(' ', '_').lower()
            placeholder_config = self.resource_manager.get_config(f"ui.placeholders.{field_key}")
            placeholder_text = placeholder_config or f"Enter {label.replace(':', '').lower()}..."
            line_edit.setPlaceholderText(placeholder_text)
            
            string_layout.addWidget(label_widget, i, 0)
            string_layout.addWidget(line_edit, i, 1)
            
            # Store reference to the field
            self.string_fields[field_key] = line_edit
        
        string_group.setLayout(string_layout)
        main_layout.addWidget(string_group)
        
        # Output section
        output_group = QGroupBox(labels_config.get("output", "Output"))
        output_group.setObjectName("outputGroup")  # For CSS styling
        output_layout = QVBoxLayout()
        
        # Output path display and selection
        path_layout = QHBoxLayout()
        
        # TODO: Get default output messages from config
        output_config = self.resource_manager.get_config("ui.output")
        default_message = output_config.get("no_location_message", "No output location selected")
        
        self.output_path_label = QLabel(default_message)
        self.output_path_label.setObjectName("outputPathLabel")  # For CSS styling
        self.output_path_label.setStyleSheet("color: #666; font-style: italic;")
        
        self.save_as_btn = QPushButton(output_config.get("button_text", "Choose Save Location"))
        self.save_as_btn.setObjectName("saveAsButton")  # For CSS styling
        
        # TODO: Add icon from resources/icons/save.png
        save_icon = self.resource_manager.get_icon("save.png")
        if save_icon:
            self.save_as_btn.setIcon(QIcon(save_icon))
        
        self.save_as_btn.clicked.connect(self.choose_output_location)
        
        path_layout.addWidget(QLabel("Output File:"))
        path_layout.addWidget(self.output_path_label, 1)
        path_layout.addWidget(self.save_as_btn)
        output_layout.addLayout(path_layout)
        
        output_group.setLayout(output_layout)
        main_layout.addWidget(output_group)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")  # For CSS styling
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # Process button
        # TODO: Get button configuration from resources/config/app_config.json
        button_config = self.resource_manager.get_config("ui.buttons.process")
        
        self.process_btn = QPushButton(button_config.get("text", "Process Files"))
        self.process_btn.setObjectName("processButton")  # For CSS styling
        self.process_btn.setMinimumHeight(button_config.get("min_height", 40))
        self.process_btn.setEnabled(False)
        
        # TODO: Add icon from resources/icons/process.png
        process_icon = self.resource_manager.get_icon("process.png")
        if process_icon:
            self.process_btn.setIcon(QIcon(process_icon))
        
        self.process_btn.clicked.connect(self.process_files)
        main_layout.addWidget(self.process_btn)
        
        # Status bar
        # TODO: Get status messages from resources/config/app_config.json
        status_config = self.resource_manager.get_config("ui.status_messages")
        ready_message = status_config.get("ready", "Ready - Select input files and configure settings")
        self.statusBar().showMessage(ready_message)
    
    def on_files_selected(self, files):
        self.input_files = files
        self.logger.info(f"Selected {len(files)} input files")  # TODO: Logger
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
            # TODO: Use FileValidator.validate_output_path(file_path)
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
            self.logger.info(f"Output path set to: {file_path}")  # TODO: Logger
            self.update_process_button_state()
    
    def update_process_button_state(self):
        # Enable process button only if we have files and output location
        can_process = bool(self.input_files and self.output_path)
        self.process_btn.setEnabled(can_process)
        
        if can_process:
            # TODO: Load active button stylesheet from resources/styles/buttons_active.qss
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
            # TODO: Load disabled button stylesheet from resources/styles/buttons_disabled.qss
            disabled_button_style = self.resource_manager.get_stylesheet("buttons_disabled.qss")
            if disabled_button_style:
                self.process_btn.setStyleSheet(disabled_button_style)
            else:
                self.process_btn.setStyleSheet("")  # Reset to default
    
    def get_string_values(self):
        """Get all string field values as a dictionary"""
        values = {}
        for key, field in self.string_fields.items():
            values[key] = field.text().strip()
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
        
        # TODO: Use FileValidator.validate_string_inputs(string_values)
        validation_errors = FileValidator.validate_string_inputs(string_values)
        if validation_errors:
            QMessageBox.warning(
                self, 
                "Invalid Configuration", 
                "Please check the following fields:\n" + "\n".join(validation_errors)
            )
            return
        
        self.logger.info("Starting file processing...")  # TODO: Logger
        self.logger.info(f"Input files: {[os.path.basename(f) for f in self.input_files]}")  # TODO: Logger
        self.logger.info(f"Output path: {self.output_path}")  # TODO: Logger
        self.logger.info(f"Configuration values: {string_values}")  # TODO: Logger
        
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
        self.processing_thread.processing_complete.connect(self.on_processing_complete)
        self.processing_thread.processing_error.connect(self.on_processing_error)
        self.processing_thread.start()
    
    def on_processing_complete(self, message):
        self.progress_bar.setVisible(False)
        self.process_btn.setEnabled(True)
        self.statusBar().showMessage("Processing complete!")
        
        self.logger.info("Processing completed successfully")  # TODO: Logger
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
                # TODO: Use FileHandler.open_file(self.output_path) for cross-platform support
                os.startfile(self.output_path)  # Windows only
                # For cross-platform: subprocess.run(['open', self.output_path])  # macOS
                # For cross-platform: subprocess.run(['xdg-open', self.output_path])  # Linux
                self.logger.info(f"Opened output file: {self.output_path}")  # TODO: Logger
            except Exception as e:
                self.logger.error(f"Failed to open output file: {e}")  # TODO: Logger
                QMessageBox.warning(self, "Cannot Open File", 
                                  f"Could not open the output file:\n{str(e)}")
    
    def on_processing_error(self, error_message):
        self.progress_bar.setVisible(False)
        self.process_btn.setEnabled(True)
        self.statusBar().showMessage("Processing failed!")
        
        self.logger.error(f"Processing failed: {error_message}")  # TODO: Logger
        QMessageBox.critical(self, "Processing Error", error_message)


def main():
    # TODO: Set up logging before creating the application
    logger = setup_logger("excel_processor")
    logger.info("Starting Excel File Processor application")
    
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("Excel File Processor")
    app.setApplicationVersion("1.0")
    
    # TODO: Initialize ResourceManager and load application-wide styles
    try:
        resource_manager = ResourceManager()
        
        # TODO: Load global stylesheet from resources/styles/app_global.qss
        global_stylesheet = resource_manager.get_stylesheet("app_global.qss")
        if global_stylesheet:
            app.setStyleSheet(global_stylesheet)
            logger.info("Loaded global application stylesheet")
        
        # TODO: Set application icon from resources/icons/app_icon.ico
        app_icon_path = resource_manager.get_icon_path("app_icon.ico")
        if app_icon_path and os.path.exists(app_icon_path):
            app.setWindowIcon(QIcon(app_icon_path))
            logger.info("Set application icon")
        
        # TODO: Load application configuration
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