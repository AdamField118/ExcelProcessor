import pandas as pd
import os
from .file_handler import FileHandler, main as process_main

class ExcelProcessor:
    @staticmethod
    def process_files(string_values, loaded_data, output_path, progress_callback=None):
        """
        Process list of DataFrames using configuration values and save to output
        
        Args:
            string_values: Dictionary of configuration values from GUI
            loaded_data: List of DataFrames from input files  
            output_path: Path where to save the output file
            progress_callback: Optional callback function for progress updates
            
        Returns:
            dict: Processing results summary
        """
        try:
            if progress_callback:
                progress_callback("Setting up output file...")
            
            FileHandler.save_excel_file(string_values, loaded_data, output_path)
            
            if progress_callback:
                progress_callback("Processing parts...")
            
            process_main(loaded_data, output_path, progress_callback)
            
            total_parts = sum(len(FileHandler.find_header_loc1_positions(df)) for df in loaded_data)
            
            return {
                'status': 'success',
                'total_files_processed': len(loaded_data),
                'total_parts_found': total_parts,
                'output_file': output_path
            }
            
        except Exception as e:
            error_msg = f"Processing failed: {str(e)}"
            if progress_callback:
                progress_callback(f"ERROR: {error_msg}")
            raise Exception(error_msg)