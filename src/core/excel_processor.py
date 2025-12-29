"""
Excel/Measurement Data Processor
Handles processing of measurement data files (old and new formats)
Inherits from BaseProcessor for consistency with ETSProcessor
"""

from src.core.base_processor import BaseProcessor
from src.core.file_handler import FileHandler, main as process_main
from src.utils.logger import get_simple_logger

logger = get_simple_logger("excel_processor")


class ExcelProcessor(BaseProcessor):
    """
    Processor for measurement data (Excel/CSV files).
    Handles both old format (LOC1-based) and new format (OK-based).
    """
    
    @staticmethod
    def validate_config(config: dict) -> tuple[bool, str]:
        """
        Validate configuration for measurement data processing.
        
        Args:
            config: Dictionary with keys:
                - part_name: str
                - revision_number: str
                - lot_number: str
                - customer_p/n: str
                - customer_po: str
                - measurement_units: str
        
        Returns:
            tuple: (is_valid, error_message)
        """
        required_fields = [
            'part_name',
            'revision_number', 
            'lot_number',
            'customer_p/n',
            'customer_po',
            'measurement_units'
        ]
        
        for field in required_fields:
            if field not in config or not config[field]:
                return False, f"Missing required field: {field}"
        
        return True, ""
    
    @staticmethod
    def process_data(loaded_data: list, config: dict, output_path: str, progress_callback=None) -> dict:
        """
        Process measurement data from loaded DataFrames.
        Matches BaseProcessor interface and ETSProcessor signature.
        
        Args:
            loaded_data: List of loaded DataFrames (already loaded by caller)
            config: Configuration dictionary (validated before calling)
            output_path: Where to save the output file
            progress_callback: Optional callback for progress updates
        
        Returns:
            dict: Results with keys:
                - status: 'success' or 'error'
                - message: Status message
                - total_files_processed: Number of files processed
                - total_parts_found: Number of parts extracted
                - output_file: Path to output file
        """
        try:
            logger.info(f"Starting measurement data processing for {len(loaded_data)} DataFrames")
            
            if not loaded_data:
                return {
                    'status': 'error',
                    'message': 'No data provided',
                    'total_files_processed': 0,
                    'total_parts_found': 0,
                    'output_file': None
                }
            
            # Create output file with template and header info
            if progress_callback:
                progress_callback("Setting up output file...")
            
            FileHandler.save_excel_file(config, loaded_data, output_path)
            logger.info(f"Created output file: {output_path}")
            
            # Process all parts
            if progress_callback:
                progress_callback("Processing parts...")
            
            process_main(loaded_data, output_path, progress_callback)
            
            # Count total parts for return value
            total_parts = len(FileHandler.find_all_parts(loaded_data))
            
            logger.info(f"Processing complete. {total_parts} parts processed from {len(loaded_data)} files")
            
            return {
                'status': 'success',
                'message': f'Successfully processed {total_parts} parts from {len(loaded_data)} files',
                'total_files_processed': len(loaded_data),
                'total_parts_found': total_parts,
                'output_file': output_path
            }
            
        except Exception as e:
            error_msg = f"Processing failed: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return {
                'status': 'error',
                'message': error_msg,
                'total_files_processed': 0,
                'total_parts_found': 0,
                'output_file': None
            }
    
    @staticmethod
    def process_files(string_values, loaded_data, output_path, progress_callback=None):
        """
        Legacy method for backward compatibility with old GUI.
        This is what the old ProcessingThread called.
        
        Args:
            string_values: Dictionary of configuration values
            loaded_data: List of DataFrames (already loaded)
            output_path: Output file path
            progress_callback: Progress callback
            
        Returns:
            dict: Processing results
        """
        try:
            if progress_callback:
                progress_callback("Setting up output file...")
            
            # Create output file with header information
            FileHandler.save_excel_file(string_values, loaded_data, output_path)
            
            if progress_callback:
                progress_callback("Processing parts...")
            
            # Process all parts using the main() function
            process_main(loaded_data, output_path, progress_callback)
            
            # Count total parts for return value
            total_parts = len(FileHandler.find_all_parts(loaded_data))
            
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