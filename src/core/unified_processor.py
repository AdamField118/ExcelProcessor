"""
Unified Processor - Handles both Measurement and ETS data processing

This processor can:
1. Process measurement data only
2. Process ETS data only  
3. Process both types simultaneously
"""

import os
import sys
import pandas as pd

# Add paths for both development and frozen executable
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
    sys.path.insert(0, base_path)
else:
    base_path = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, os.path.abspath(os.path.join(base_path, '..', '..')))

try:
    from src.core.file_handler import FileHandler
    from src.ets.processor import ETSProcessor
    from src.utils.logger import get_simple_logger
except ImportError:
    from core.file_handler import FileHandler
    from ets.processor import ETSProcessor
    from utils.logger import get_simple_logger

logger = get_simple_logger("unified_processor")


class UnifiedProcessor:
    """
    Unified processor that handles both measurement and ETS data
    """
    
    MEASUREMENT_ONLY = "measurement"
    ETS_ONLY = "ets"
    BOTH = "both"
    
    @staticmethod
    def identify_file_type(file_path):
        """
        Identify if a file is ETS or Measurement data by checking A1 cell
        
        Args:
            file_path: Path to the file
            
        Returns:
            'ets' or 'measurement'
        """
        try:
            # Load just the first cell to check
            df = FileHandler.load_excel_file(file_path)
            
            # Check A1 cell (row 0, column 0)
            if len(df) > 0 and len(df.columns) > 0:
                a1_value = str(df.iloc[0, 0]).strip()
                
                if 'ETS Export' in a1_value or 'ETS' in a1_value:
                    logger.info(f"Identified {os.path.basename(file_path)} as ETS data")
                    return 'ets'
            
            logger.info(f"Identified {os.path.basename(file_path)} as Measurement data")
            return 'measurement'
            
        except Exception as e:
            logger.warning(f"Could not identify file type for {file_path}: {e}. Defaulting to measurement.")
            return 'measurement'
    
    @staticmethod
    def separate_files_by_type(file_paths):
        """
        Separate files into ETS and Measurement lists
        
        Args:
            file_paths: List of file paths
            
        Returns:
            dict with 'ets' and 'measurement' keys containing file lists
        """
        separated = {
            'ets': [],
            'measurement': []
        }
        
        for file_path in file_paths:
            file_type = UnifiedProcessor.identify_file_type(file_path)
            separated[file_type].append(file_path)
        
        logger.info(f"Separated files: {len(separated['ets'])} ETS, {len(separated['measurement'])} Measurement")
        return separated
    
    @staticmethod
    def validate_config(mode, measurement_config, ets_config):
        """
        Validate configuration based on processing mode
        
        Args:
            mode: Processing mode (measurement, ets, or both)
            measurement_config: Config dict for measurement processing
            ets_config: Config dict for ETS processing
            
        Returns:
            List of error messages (empty if valid)
        """
        errors = []
        
        if mode in [UnifiedProcessor.MEASUREMENT_ONLY, UnifiedProcessor.BOTH]:
            # Validate measurement fields
            required_measurement = ['part_name', 'revision_number', 'lot_number', 
                                   'customer_p/n', 'customer_po', 'measurement_units']
            for field in required_measurement:
                value = measurement_config.get(field, '')
                if not isinstance(value, str) or not value.strip():
                    field_display = field.replace('_', ' ').replace('/', ' / ').title()
                    errors.append(f"Measurement: {field_display} is required")
        
        if mode in [UnifiedProcessor.ETS_ONLY, UnifiedProcessor.BOTH]:
            # Validate ETS fields
            from src.ets.processor import ETSProcessor
            ets_errors = ETSProcessor.validate_config(ets_config)
            for error in ets_errors:
                errors.append(f"ETS: {error}")
        
        return errors
    
    @staticmethod
    def process_unified(mode, input_files, measurement_config, ets_config, 
                       measurement_output_path, ets_output_path, progress_callback=None):
        """
        Process files based on mode
        
        Args:
            mode: Processing mode (measurement, ets, or both)
            input_files: List of input file paths
            measurement_config: Config for measurement processing
            ets_config: Config for ETS processing
            measurement_output_path: Output path for measurement data
            ets_output_path: Output path for ETS data
            progress_callback: Optional callback for progress updates
            
        Returns:
            dict with processing results
        """
        results = {
            'status': 'success',
            'measurement': None,
            'ets': None
        }
        
        try:
            # Separate files by type
            if mode == UnifiedProcessor.BOTH:
                if progress_callback:
                    progress_callback("Separating files by type...")
                
                separated_files = UnifiedProcessor.separate_files_by_type(input_files)
                measurement_files = separated_files['measurement']
                ets_files = separated_files['ets']
                
                if not measurement_files and not ets_files:
                    raise Exception("No valid files found for processing")
                
                if not measurement_files:
                    logger.warning("No measurement files found, processing ETS only")
                    mode = UnifiedProcessor.ETS_ONLY
                elif not ets_files:
                    logger.warning("No ETS files found, processing measurement only")
                    mode = UnifiedProcessor.MEASUREMENT_ONLY
                    
            elif mode == UnifiedProcessor.MEASUREMENT_ONLY:
                measurement_files = input_files
                ets_files = []
            else:  # ETS_ONLY
                measurement_files = []
                ets_files = input_files
            
            # Process measurement data
            if mode in [UnifiedProcessor.MEASUREMENT_ONLY, UnifiedProcessor.BOTH]:
                if progress_callback:
                    progress_callback("Processing measurement data...")
                
                logger.info(f"Processing {len(measurement_files)} measurement files")
                
                # Load measurement files
                loaded_measurement = []
                for i, file_path in enumerate(measurement_files):
                    if progress_callback:
                        progress_callback(f"Loading measurement file {i+1}/{len(measurement_files)}...")
                    loaded_measurement.append(FileHandler.load_excel_file(file_path))
                
                # Process measurement data using ExcelProcessor
                try:
                    from src.core.excel_processor import ExcelProcessor
                except ImportError:
                    from core.excel_processor import ExcelProcessor
                
                measurement_result = ExcelProcessor.process_data(
                    loaded_measurement,
                    measurement_config,
                    measurement_output_path,
                    progress_callback
                )
                
                results['measurement'] = measurement_result
                logger.info("Measurement processing complete")
            
            # Process ETS data
            if mode in [UnifiedProcessor.ETS_ONLY, UnifiedProcessor.BOTH]:
                if progress_callback:
                    progress_callback("Processing ETS data...")
                
                logger.info(f"Processing {len(ets_files)} ETS files")
                
                # Load ETS files
                loaded_ets = []
                for i, file_path in enumerate(ets_files):
                    if progress_callback:
                        progress_callback(f"Loading ETS file {i+1}/{len(ets_files)}...")
                    loaded_ets.append(FileHandler.load_excel_file(file_path))
                
                # Process ETS data
                ets_processor = ETSProcessor()
                ets_result = ets_processor.process_data(
                    loaded_ets,
                    ets_config,
                    ets_output_path,
                    progress_callback
                )
                
                results['ets'] = ets_result
                logger.info("ETS processing complete")
            
            if progress_callback:
                progress_callback("Processing complete!")
            
            return results
            
        except Exception as e:
            error_msg = f"Unified processing failed: {str(e)}"
            logger.error(error_msg)
            results['status'] = 'failed'
            results['error'] = error_msg
            raise Exception(error_msg)