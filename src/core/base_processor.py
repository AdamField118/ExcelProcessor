"""
Base Processor - Abstract base class for all processor variants
"""

from abc import ABC, abstractmethod
import pandas as pd
from src.core.file_handler import FileHandler


class BaseProcessor(ABC):
    """
    Abstract base class for file processors.
    Subclasses must implement template_path and process_data methods.
    """
    
    @property
    @abstractmethod
    def template_path(self):
        """Return path to the output template file"""
        pass
    
    @abstractmethod
    def process_data(self, input_data, config_values, output_path, progress_callback=None):
        """
        Process input data and save to output
        
        Args:
            input_data: List of DataFrames from input files
            config_values: Dictionary of configuration values
            output_path: Path where to save the output file
            progress_callback: Optional callback for progress updates
            
        Returns:
            dict: Processing results summary
        """
        pass
    
    @staticmethod
    def load_files(file_paths, progress_callback=None):
        """
        Load multiple files into DataFrames
        
        Args:
            file_paths: List of file paths to load
            progress_callback: Optional callback for progress updates
            
        Returns:
            List of DataFrames
        """
        loaded_data = []
        total_files = len(file_paths)
        
        for i, file_path in enumerate(file_paths):
            if progress_callback:
                progress_callback(f"Loading file {i+1}/{total_files}...")
            
            file_data = FileHandler.load_excel_file(file_path)
            loaded_data.append(file_data)
        
        return loaded_data
    
    @staticmethod
    def validate_inputs(config_values, required_fields):
        """
        Validate that all required configuration fields are present
        
        Args:
            config_values: Dictionary of configuration values
            required_fields: List of required field names
            
        Returns:
            List of error messages (empty if valid)
        """
        errors = []
        
        for field in required_fields:
            value = config_values.get(field, '')
            if not isinstance(value, str) or not value.strip():
                errors.append(f"{field} is required")
        
        return errors