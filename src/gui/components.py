"""
Shared GUI Components - Reusable widgets for both Excel and ETS processors
"""

import os
import sys

# Add paths for both development and frozen executable
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
    sys.path.insert(0, base_path)
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

from PyQt5.QtWidgets import (QFrame, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QLabel, QTextEdit, QFileDialog, QMessageBox, QSizePolicy)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont

try:
    from src.utils.validators import FileValidator
    from src.utils.resource_manager import ResourceManager
except ImportError:
    from utils.validators import FileValidator
    from utils.resource_manager import ResourceManager


class FileDropArea(QFrame):
    """Custom widget for drag and drop file functionality - REUSABLE"""
    files_dropped = pyqtSignal(list)
    
    def __init__(self, file_types="Data Files (*.xlsx *.xls *.csv)", 
                 title_text="Drop Excel or CSV files here or click Browse"):
        super().__init__()
        self.file_types = file_types
        self.title_text = title_text
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
        up_icon = self.style().standardIcon(self.style().SP_ArrowUp)
        upload_icon.setPixmap(up_icon.pixmap(32, 32))
        
        self.main_label = QLabel(self.title_text)
        self.main_label.setAlignment(Qt.AlignCenter)
        
        font_config = self.resource_manager.get_config("fonts.main_label")
        font = QFont()
        font.setPointSize(font_config.get("size", 12) if font_config else 12)
        font.setBold(font_config.get("bold", False) if font_config else False)
        self.main_label.setFont(font)
        
        label_layout.addWidget(upload_icon)
        label_layout.addWidget(self.main_label)
        label_layout.setAlignment(Qt.AlignCenter)
        
        # Browse button
        self.browse_btn = QPushButton("Browse Files")
        folder_icon = self.style().standardIcon(self.style().SP_DirIcon)
        self.browse_btn.setIcon(folder_icon)
        
        self.browse_btn.setMaximumWidth(150)
        self.browse_btn.setMinimumHeight(45)
        self.browse_btn.clicked.connect(self.browse_files)
        
        # File list display
        self.file_list = QTextEdit()
        self.browse_btn.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.file_list.setMinimumHeight(50)
        self.file_list.setMaximumHeight(100)
        placeholder_text = self.resource_manager.get_config("ui.placeholders.file_list", 
                                                            "Selected files will appear here...")
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
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            active_stylesheet = self.resource_manager.get_stylesheet("drop_area_active.qss")
            if active_stylesheet:
                self.setStyleSheet(active_stylesheet)
    
    def dragLeaveEvent(self, event):
        self.load_styles()
    
    def dropEvent(self, event):
        files = []
        invalid_files = []
        
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if FileValidator.is_valid_data_file(file_path):
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
                    f"The following files were skipped (not valid data files):\n" + 
                    "\n".join(invalid_files)
                )
        else:
            QMessageBox.warning(self, "Invalid Files", 
                              "Please drop only valid data files")
        
        self.dragLeaveEvent(None)
    
    def browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "Select Data Files",
            "",
            self.file_types
        )
        
        if files:
            valid_files = []
            invalid_files = []
            
            for file_path in files:
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
                QMessageBox.warning(self, "No Valid Files", 
                                  "No valid data files were selected.")
    
    def update_file_list(self, files):
        file_names_with_types = []
        for f in files:
            file_name = os.path.basename(f)
            file_type = FileValidator.get_file_type(f)
            file_names_with_types.append(f"{file_name} ({file_type})")
        
        self.file_list.setText('\n'.join(file_names_with_types))