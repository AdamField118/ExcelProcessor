import pandas as pd
import openpyxl
import os
import shutil
import sys
from utils.logger import get_simple_logger
from contextlib import contextmanager
import math

logger = get_simple_logger("file_handler")

XLRD_AVAILABLE = None
XLRD_VERSION = None

def check_xlrd():
    """Check xlrd availability when needed"""
    global XLRD_AVAILABLE, XLRD_VERSION
    
    if XLRD_AVAILABLE is not None:
        return XLRD_AVAILABLE
    
    try:
        logger.info("Checking xlrd availability...")
        import xlrd
        XLRD_AVAILABLE = True
        XLRD_VERSION = xlrd.__version__
        logger.info(f"xlrd available, version: {XLRD_VERSION}")
        
        try:
            pd.io.excel._readers['xls'] = 'xlrd'
            logger.info("Registered xlrd with pandas")
        except Exception as reg_error:
            logger.error(f"Could not register xlrd with pandas: {reg_error}")
            
        return True
        
    except ImportError as e:
        XLRD_AVAILABLE = False
        XLRD_VERSION = None
        logger.error(f"xlrd not available: {e}")
        return False
    except Exception as e:
        XLRD_AVAILABLE = False
        XLRD_VERSION = None
        logger.error(f"xlrd error: {e}")
        return False

@contextmanager
def excel_batch_writer(output_path):
    """Context manager for batch Excel operations - keeps workbook open"""
    workbook = openpyxl.load_workbook(output_path)
    try:
        yield workbook
    finally:
        workbook.save(output_path)
        workbook.close()

class FileHandler:
    @staticmethod
    def truncate_to_three_decimals(value):
        """Truncate numeric value to 3 decimal places (do not round)"""
        if value == "" or value is None:
            return ""
        
        try:
            # Convert to float
            numeric_value = float(value)
            
            # Truncate to 3 decimal places by multiplying by 1000, 
            # taking integer part, then dividing by 1000
            truncated = math.trunc(numeric_value * 1000) / 1000
            
            # Format to exactly 3 decimal places
            return f"{truncated:.3f}"
            
        except (ValueError, TypeError):
            # If it's not a number, return as string
            return str(value)
    
    @staticmethod
    def load_excel_file(file_path):
        """Load Excel file into DataFrame"""
        logger.info(f"Loading: {os.path.basename(file_path)}")
        
        if file_path.lower().endswith('.csv'):
            logger.info("Loading .csv file")
            return FileHandler._load_csv_file(file_path)

        xlrd_works = check_xlrd()
        
        try:
            if file_path.lower().endswith('.xls'):
                logger.info("Loading .xls file")
                
                if not xlrd_works:
                    raise Exception("xlrd is not available for .xls files")
                
                import xlrd
                result = pd.read_excel(file_path, header=None, engine='xlrd')
                logger.info(f"Successfully loaded .xls file, shape: {result.shape}")
                result.attrs['source_file'] = os.path.basename(file_path)
                return result
            else:
                logger.info("Loading .xlsx file")
                result = pd.read_excel(file_path, header=None, engine='openpyxl')
                logger.info(f"Successfully loaded .xlsx file, shape: {result.shape}")
                result.attrs['source_file'] = os.path.basename(file_path)
                return result
        except Exception as e:
            logger.error(f"Failed to load {os.path.basename(file_path)}: {str(e)}")
            if file_path.lower().endswith('.xls'):
                try:
                    logger.info("Trying fallback method for .xls file")
                    result = pd.read_excel(file_path, header=None)
                    logger.info(f"Fallback successful, shape: {result.shape}")
                    result.attrs['source_file'] = os.path.basename(file_path)
                    return result
                except Exception as e2:
                    logger.error(f"Fallback also failed: {str(e2)}")
                    raise Exception(f"Cannot read .xls file. xlrd issue: {str(e)}")
            else:
                raise Exception(f"Failed to load {os.path.basename(file_path)}: {str(e)}")

    @staticmethod
    def save_excel_file(input_strings, data, output_path):
        """Save processed data to Excel with formatting"""
        template_path = os.path.join("resources", "config", "out_template.xlsx")
        
        if not os.path.exists(template_path):
            current_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(os.path.dirname(current_dir))
            template_path = os.path.join(project_root, "resources", "config", "out_template.xlsx")
        
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found. Expected at: {template_path}")
            
        shutil.copy2(template_path, output_path)
        
        workbook = openpyxl.load_workbook(output_path)
        
        customer_po = input_strings.get('customer_po', '')
        
        # Count parts correctly
        part_position_dict = FileHandler.find_header_loc1_positions(data)
        total_parts = len(part_position_dict)
        
        sheet = workbook.active
        sheet_name = f'PO{customer_po} Qty {total_parts}'
        sheet.title = sheet_name
        
        sheet['C1'] = input_strings.get('part_name', '')
        sheet['C2'] = input_strings.get('revision_number', '')
        sheet['C3'] = input_strings.get('lot_number', '')
        sheet['C4'] = input_strings.get('customer_p/n', '')
        sheet['C5'] = customer_po
        sheet['C6'] = input_strings.get('measurement_units', '')
        
        workbook.save(output_path)
        workbook.close()
    
    @staticmethod
    def find_header_loc1_positions(data):
        """
        Find only HEADER LOC1 positions (not data LOC1) in DataFrame(s)
        
        Header LOC1 instances are identified by having 'UNITS' in the adjacent column
        These are the reference points for measurement extraction logic
        
        Returns dictionary with part info including which DataFrame contains each part
        """
        if isinstance(data, list):
            all_positions = {}
            part_counter = 1
            
            for df_idx, df in enumerate(data):
                logger.info(f"Scanning DataFrame {df_idx + 1}/{len(data)} for header LOC1...")
                df_positions = FileHandler._find_header_loc1_in_single_df(df, part_counter, df_idx)
                all_positions.update(df_positions)
                part_counter += len(df_positions)
                logger.info(f"Found {len(df_positions)} header LOC1 positions in DataFrame {df_idx + 1}")
            
            logger.info(f"Total header LOC1 positions found: {len(all_positions)}")
            return all_positions
        else:
            return FileHandler._find_header_loc1_in_single_df(data, 1, 0)
    
    @staticmethod
    def _find_header_loc1_in_single_df(df, start_counter, df_index):
        """
        Helper function to find HEADER LOC1 positions in a single DataFrame
        
        Based on DataFrame output and sheet range B2:V392:
        - DataFrame row 0 = Excel row 2
        - Index 0: Empty (original Excel column A)
        - Index 1: LOC1 data (original Excel column B) 
        - Index 2: UNITS data (original Excel column C)
        
        Returns dictionary with part info including DataFrame index
        """
        positions = {}
        part_counter = start_counter
        
        df_str = df.astype(str)
        
        for row_idx in range(len(df)):
            if len(df.columns) > 2:  # Need at least 3 columns (0, 1, 2)
                cell_value = df_str.iloc[row_idx, 1].strip()  # Check index 1 for LOC1
                
                if cell_value == 'LOC1':
                    try:
                        # Check index 2 for UNITS
                        units_value = df_str.iloc[row_idx, 2].strip()
                        
                        if 'UNITS' in units_value:
                            # LOC1 is at DataFrame index 1 = Original Excel column B
                            # DataFrame row 0 = Excel row 2, so Excel row = row_idx + 2
                            excel_col = 'B'
                            excel_row = row_idx + 2
                            excel_ref = f'{excel_col}{excel_row}'
                            
                            # Store both position and DataFrame index
                            positions[f'part-{part_counter}'] = {
                                'position': excel_ref,
                                'df_index': df_index
                            }
                            part_counter += 1
                            logger.debug(f"Found header LOC1 at {excel_ref} with UNITS at index 2 in DataFrame {df_index}")
                        else:
                            logger.debug(f"Skipping data LOC1 at row {row_idx + 2} with index 2 value: '{units_value}'")
                            
                    except (IndexError, KeyError):
                        logger.warning(f"Could not check index 2 for LOC1 at row {row_idx + 2}")
        
        return positions
    
    @staticmethod
    def _number_to_excel_column(n):
        """Convert number to Excel column letter (1=A, 2=B, ..., 26=Z, 27=AA, etc.)"""
        result = ""
        while n > 0:
            n -= 1
            result = chr(n % 26 + ord('A')) + result
            n //= 26
        return result
    
    @staticmethod
    def _excel_column_to_number(col_letter):
        """Convert Excel column letter to number (A=1, B=2, ..., Z=26, AA=27, etc.)"""
        result = 0
        for char in col_letter:
            result = result * 26 + (ord(char.upper()) - ord('A') + 1)
        return result
    
    @staticmethod
    def extract_value(df, part_positions_dict, part_name, col_offset, row_offset):
        """Extract value from DataFrame relative to LOC1 position and return it"""
        if part_name not in part_positions_dict:
            raise ValueError(f"Part '{part_name}' not found in positions dictionary")
        
        part_info = part_positions_dict[part_name]
        
        # Handle both old format (string) and new format (dict)
        if isinstance(part_info, dict):
            loc1_position = part_info['position']
        else:
            loc1_position = part_info
        
        loc1_col_letter = ''.join(filter(str.isalpha, loc1_position))
        loc1_row_num = int(''.join(filter(str.isdigit, loc1_position)))
        loc1_col_num = FileHandler._excel_column_to_number(loc1_col_letter)
        
        target_col_num = loc1_col_num + col_offset
        target_row_num = loc1_row_num + row_offset
        
        # Convert Excel coordinates to DataFrame indices
        # DataFrame index 0 = Excel column A (1), index 1 = Excel column B (2), etc.
        # DataFrame row 0 = Excel row 2, row 1 = Excel row 3, etc.
        target_df_col = target_col_num - 1  # Excel col A(1) → DF index 0
        target_df_row = target_row_num - 2  # Excel row 2 → DF row 0
        
        try:
            value = df.iloc[target_df_row, target_df_col]
            if pd.notna(value):
                try:
                    return float(value)
                except (ValueError, TypeError):
                    print(f"*******COULDNT DO FLOAT: {value}")
                    return str(value)
            else:
                print("WHAT????")
                return ""
        except (IndexError, KeyError):
            filename = getattr(df, 'attrs', {}).get('source_file', 'Unknown file')
            print(f"HUHHHHHHHHHHHHHHHHH???? row {target_df_row} and column {target_df_col} in file: {filename}")
            return ""
    
    @staticmethod
    def extract_all_measurements_for_part(df, part_positions_dict, part_name):
        """Extract all measurements for a part in one pass"""
        measurements = {}
        
        # Base extractions for all parts
        base_extractions = [
            ('A', 6, 2),    ('B', 10, 2),   ('C', 11, 2),
            ('D', 6, 10),   ('E', 10, 10),  ('F', 11, 10),
            ('G', 6, 14),   ('H', 10, 14),  ('I', 11, 14),
            ('J', 5, 2),    ('K', 7, 2),    ('L', 8, 2),
        ]
        
        # Part-1 specific extractions (only for part-1)
        part1_extractions = [
            ('O', 5, 10),   ('P', 7, 10),   ('Q', 8, 10),
            ('T', 5, 14),   ('U', 7, 14),
        ]
        
        # Extract base measurements for all parts
        for col, col_offset, row_offset in base_extractions:
            value = FileHandler.extract_value(df, part_positions_dict, part_name, col_offset, row_offset)
            measurements[col] = FileHandler.truncate_to_three_decimals(value)
        
        # Only extract part-1 specific measurements for part-1
        if part_name == 'part-1':
            for col, col_offset, row_offset in part1_extractions:
                value = FileHandler.extract_value(df, part_positions_dict, part_name, col_offset, row_offset)
                measurements[col] = FileHandler.truncate_to_three_decimals(value)
            
            # Extract values for calculations
            val_5_2 = FileHandler.extract_value(df, part_positions_dict, part_name, 5, 2)
            val_7_2 = FileHandler.extract_value(df, part_positions_dict, part_name, 7, 2)
            val_8_2 = FileHandler.extract_value(df, part_positions_dict, part_name, 8, 2)
            val_5_10 = FileHandler.extract_value(df, part_positions_dict, part_name, 5, 10)
            val_7_10 = FileHandler.extract_value(df, part_positions_dict, part_name, 7, 10)
            val_8_10 = FileHandler.extract_value(df, part_positions_dict, part_name, 8, 10)
            val_5_14 = FileHandler.extract_value(df, part_positions_dict, part_name, 5, 14)
            val_7_14 = FileHandler.extract_value(df, part_positions_dict, part_name, 7, 14)
            
            # Calculate derived values (only for part-1) and truncate results
            if isinstance(val_5_2, (int, float)) and isinstance(val_7_2, (int, float)):
                measurements['M'] = FileHandler.truncate_to_three_decimals(val_5_2 + val_7_2)
            else:
                measurements['M'] = ""
            
            if isinstance(val_5_2, (int, float)) and isinstance(val_8_2, (int, float)):
                measurements['N'] = FileHandler.truncate_to_three_decimals(val_5_2 - val_8_2)
            else:
                measurements['N'] = ""
            
            if isinstance(val_5_10, (int, float)) and isinstance(val_7_10, (int, float)):
                measurements['R'] = FileHandler.truncate_to_three_decimals(val_5_10 + val_7_10)
            else:
                measurements['R'] = ""
            
            if isinstance(val_5_10, (int, float)) and isinstance(val_8_10, (int, float)):
                measurements['S'] = FileHandler.truncate_to_three_decimals(val_5_10 - val_8_10)
            else:
                measurements['S'] = ""
            
            if isinstance(val_5_14, (int, float)) and isinstance(val_7_14, (int, float)):
                measurements['V'] = FileHandler.truncate_to_three_decimals(val_5_14 + val_7_14)
            else:
                measurements['V'] = ""
        
        return measurements
    
    @staticmethod
    def write_measurements_to_row(sheet, row_num, measurements):
        """Write all measurements to a single row in one operation"""
        row_str = str(row_num)
        
        for col_letter, value in measurements.items():
            if value:
                sheet[f'{col_letter}{row_str}'] = value
    
    @staticmethod
    def process_part(data, part_positions_dict, part_name, workbook, start_row):
        """
        Process measurements for a part from the specific DataFrame where it was found
        Returns the next available row number
        """
        if not isinstance(data, list):
            data = [data]
        
        if part_name not in part_positions_dict:
            return start_row
        
        part_info = part_positions_dict[part_name]
        
        # Handle both old format (string) and new format (dict)
        if isinstance(part_info, dict):
            df_index = part_info['df_index']
        else:
            # Fallback for old format - process against all DataFrames (causes whitespace)
            df_index = 0
        
        # Only process the specific DataFrame that contains this part
        if df_index < len(data):
            df = data[df_index]
            
            # Extract measurements from the correct DataFrame
            measurements = FileHandler.extract_all_measurements_for_part(df, part_positions_dict, part_name)
            
            # Write to all sheets
            for sheet in workbook.worksheets:
                FileHandler.write_measurements_to_row(sheet, start_row, measurements)
            
            return start_row + 1
        else:
            logger.warning(f"DataFrame index {df_index} not available for part {part_name}")
            return start_row
    
    @staticmethod
    def open_file(file_path):
        """Windows-only file opening"""
        try:
            os.startfile(file_path)
            return True
        except Exception as e:
            logger.error(f"Failed to open file: {e}")
            return False
        
    @staticmethod
    def _load_csv_file(file_path):
        """Load CSV file with multiple encoding attempts"""
        # Try different encodings commonly used for CSV files
        encodings_to_try = ['utf-8', 'utf-8-sig', 'latin1', 'cp1252', 'iso-8859-1']
        
        def _is_first_column_numeric(df):
            """Check if the first column contains only numeric values"""
            if df.empty or len(df.columns) == 0:
                return False
            
            first_col = df.iloc[:, 0].dropna()
            
            if len(first_col) == 0:
                return False
            
            # Try to convert all values to numeric
            try:
                pd.to_numeric(first_col, errors='raise')
                return True
            except (ValueError, TypeError):
                return False
        
        for encoding in encodings_to_try:
            try:
                logger.info(f"Trying to load CSV with encoding: {encoding}")
                result = pd.read_csv(
                    file_path, 
                    header=None,  # No headers, same as Excel loading
                    encoding=encoding,
                    keep_default_na=True,  # Keep NaN values
                    na_values=[''],  # Treat empty strings as NaN
                    dtype=str  # Load everything as strings initially
                )

                # Only drop the first column if it appears to be numeric (like an index)
                if _is_first_column_numeric(result):
                    logger.info("First column appears to be numeric index, removing it")
                    result = result.drop(result.columns[0], axis=1)
                    result.columns = range(len(result.columns))
                else:
                    logger.info("First column contains non-numeric data, keeping it")

                logger.info(f"Successfully loaded CSV with encoding: {encoding}, shape: {result.shape}")
                result.attrs['source_file'] = os.path.basename(file_path)
                return result
            except UnicodeDecodeError:
                logger.warning(f"Failed to decode CSV with encoding: {encoding}")
                continue
            except Exception as e:
                logger.error(f"Error loading CSV with encoding {encoding}: {str(e)}")
                if encoding == encodings_to_try[-1]:  # Last encoding attempt
                    raise
                continue
        
        raise Exception("Could not load CSV file with any supported encoding")


def main(data, output_path, progress_callback=None):
    """Process all parts found in the data with optimized batch operations"""
    try:
        logger.info("Starting processing...")
        
        if progress_callback:
            progress_callback("Finding header LOC1 positions...")
        
        part_position_dict = FileHandler.find_header_loc1_positions(data)
        
        if not part_position_dict:
            raise ValueError("No header LOC1 positions found in the input data")
        
        total_parts = len(part_position_dict)
        logger.info(f"Found {total_parts} header parts to process")
        
        if progress_callback:
            progress_callback(f"Processing {total_parts} parts...")
        
        with excel_batch_writer(output_path) as workbook:
            current_row = 9
            processed_parts = 0
            
            for part_name in part_position_dict.keys():
                logger.info(f"Processing {part_name}... ({processed_parts + 1}/{total_parts})")
                
                current_row = FileHandler.process_part(
                    data, part_position_dict, part_name, workbook, current_row
                )
                
                processed_parts += 1
                
                if progress_callback:
                    progress = int((processed_parts / total_parts) * 100)
                    progress_callback(f"Processed {processed_parts}/{total_parts} parts ({progress}%)")
                
                if processed_parts % 10 == 0:
                    logger.info(f"Processed {processed_parts}/{total_parts} parts")
        
        logger.info(f"Successfully processed all {total_parts} parts. Output saved to: {output_path}")
        
        if progress_callback:
            progress_callback("Processing complete!")
        
    except Exception as e:
        logger.error(f"Error in processing: {e}")
        raise