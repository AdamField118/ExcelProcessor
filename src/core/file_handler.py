import pandas as pd
import openpyxl
import os
import shutil
import sys
from src.utils.logger import get_simple_logger
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
    def is_new_format(df):
        """
        Detect if DataFrame is in new format by checking if column A (index 0) contains
        only 'Design value', 'Upper Limit', and 'Lower Limit' (exactly one of each).
        This check is done AFTER any numerical first column has been popped.
        """
        if df.empty or len(df.columns) == 0:
            return False
        
        try:
            # Get first column (index 0) as strings, drop NaN/empty
            first_col = df.iloc[:, 0].astype(str).str.strip()
            # Filter out empty strings and 'nan' strings
            non_empty = first_col[(first_col != '') & (first_col != 'nan')].tolist()
            
            if len(non_empty) != 3:
                return False
            
            # Check if we have exactly one of each required value
            required_values = {'Design value', 'Upper Limit', 'Lower Limit'}
            return set(non_empty) == required_values
            
        except Exception as e:
            logger.error(f"Error checking new format: {e}")
            return False
    
    @staticmethod
    def extract_new_format_reference_values(df):
        """
        Extract reference values from new format file for first part only.
        Returns dict with nominal values and tolerances.
        """
        reference_values = {}
        
        # Find rows for each reference type
        first_col = df.iloc[:, 0].astype(str).str.strip()
        
        design_row = None
        upper_row = None
        lower_row = None
        
        for idx, val in first_col.items():
            if val == 'Design value':
                design_row = idx
            elif val == 'Upper Limit':
                upper_row = idx
            elif val == 'Lower Limit':
                lower_row = idx
        
        if design_row is None:
            raise ValueError("Could not find 'Design value' row in new format")
        
        # Extract values from columns G, H, I (indices 6, 7, 8)
        # Column G values → K, L, M
        reference_values['K'] = df.iloc[design_row, 6] if design_row is not None and len(df.columns) > 6 else ""
        reference_values['L'] = df.iloc[upper_row, 6] if upper_row is not None and len(df.columns) > 6 else ""
        reference_values['M'] = df.iloc[lower_row, 6] if lower_row is not None and len(df.columns) > 6 else ""
        
        # Column H values → U, V, (W is blank)
        reference_values['U'] = df.iloc[design_row, 7] if design_row is not None and len(df.columns) > 7 else ""
        reference_values['V'] = df.iloc[upper_row, 7] if upper_row is not None and len(df.columns) > 7 else ""
        reference_values['W'] = ""
        
        # Column I values → P, Q, R
        reference_values['P'] = df.iloc[design_row, 8] if design_row is not None and len(df.columns) > 8 else ""
        reference_values['Q'] = df.iloc[upper_row, 8] if upper_row is not None and len(df.columns) > 8 else ""
        reference_values['R'] = df.iloc[lower_row, 8] if lower_row is not None and len(df.columns) > 8 else ""
        
        # Store nominal values for calculations (as floats)
        try:
            reference_values['nominal_G'] = float(df.iloc[design_row, 6]) if design_row is not None and len(df.columns) > 6 and pd.notna(df.iloc[design_row, 6]) else 0
        except (ValueError, TypeError):
            reference_values['nominal_G'] = 0
            
        try:
            reference_values['nominal_H'] = float(df.iloc[design_row, 7]) if design_row is not None and len(df.columns) > 7 and pd.notna(df.iloc[design_row, 7]) else 0
        except (ValueError, TypeError):
            reference_values['nominal_H'] = 0
            
        try:
            reference_values['nominal_I'] = float(df.iloc[design_row, 8]) if design_row is not None and len(df.columns) > 8 and pd.notna(df.iloc[design_row, 8]) else 0
        except (ValueError, TypeError):
            reference_values['nominal_I'] = 0
        
        logger.info(f"Extracted reference values: K={reference_values['K']}, L={reference_values['L']}, M={reference_values['M']}")
        logger.info(f"Nominal values: G={reference_values['nominal_G']}, H={reference_values['nominal_H']}, I={reference_values['nominal_I']}")
        
        return reference_values
    
    @staticmethod
    def find_ok_positions_new_format(df, start_counter, df_index, reference_values):
        """
        Find all 'OK' positions in column E (index 4) for new format.
        Returns dict with part info similar to old format.
        """
        positions = {}
        part_counter = start_counter
        
        if len(df.columns) <= 4:
            logger.warning("DataFrame doesn't have enough columns for new format")
            return positions
        
        col_e = df.iloc[:, 4].astype(str).str.strip()
        
        for idx, val in col_e.items():
            if val == 'OK':
                positions[f'part-{part_counter}'] = {
                    'row_index': idx,  # Row where OK was found
                    'df_index': df_index,
                    'format': 'new',
                    'reference_values': reference_values
                }
                part_counter += 1
                logger.debug(f"Found OK at row {idx} in DataFrame {df_index}")
        
        logger.info(f"Found {len(positions)} OK positions in new format file")
        return positions
    
    @staticmethod
    def extract_measurements_new_format(df, part_info, is_first_part=False):
        """
        Extract measurements for a part in new format.
        part_info: dict containing row_index, reference_values, etc.
        """
        measurements = {}
        ok_row = part_info['row_index']
        reference_values = part_info['reference_values']
        
        # Column A: Lot Number (2 columns to the left of OK, which is at index 4)
        # So index 4 - 2 = index 2
        try:
            lot_number = df.iloc[ok_row, 2] if len(df.columns) > 2 else ""
            measurements['A'] = str(lot_number).strip() if pd.notna(lot_number) else ""
        except (IndexError, KeyError):
            measurements['A'] = ""
        
        # Get nominal values
        nominal_G = reference_values['nominal_G']
        nominal_H = reference_values['nominal_H']
        nominal_I = reference_values['nominal_I']
        
        # Column B: 2 columns to the right of OK (index 4 + 2 = index 6)
        try:
            actual_B_raw = df.iloc[ok_row, 6] if len(df.columns) > 6 else ""
            measurements['B'] = FileHandler.truncate_to_three_decimals(actual_B_raw)
            
            # Column C: truncated_B - nominal_G (use truncated value for calculation)
            if pd.notna(actual_B_raw) and actual_B_raw != "":
                try:
                    # Convert truncated string back to float for calculation
                    actual_B_truncated = float(measurements['B'])
                    measurements['C'] = FileHandler.truncate_to_three_decimals(actual_B_truncated - nominal_G)
                except (ValueError, TypeError):
                    measurements['C'] = ""
            else:
                measurements['C'] = ""
        except (IndexError, KeyError):
            measurements['B'] = ""
            measurements['C'] = ""
        
        # Column D: 0.000
        measurements['D'] = "0.000"
        
        # Column E: 4 columns to the right of OK (index 4 + 4 = index 8)
        try:
            actual_E_raw = df.iloc[ok_row, 8] if len(df.columns) > 8 else ""
            measurements['E'] = FileHandler.truncate_to_three_decimals(actual_E_raw)
            
            # Column F: truncated_E - nominal_I (use truncated value for calculation)
            if pd.notna(actual_E_raw) and actual_E_raw != "":
                try:
                    # Convert truncated string back to float for calculation
                    actual_E_truncated = float(measurements['E'])
                    measurements['F'] = FileHandler.truncate_to_three_decimals(actual_E_truncated - nominal_I)
                except (ValueError, TypeError):
                    measurements['F'] = ""
            else:
                measurements['F'] = ""
        except (IndexError, KeyError):
            measurements['E'] = ""
            measurements['F'] = ""
        
        # Column G: 0.000
        measurements['G'] = "0.000"
        
        # Column H: 3 columns to the right of OK (index 4 + 3 = index 7)
        try:
            actual_H_raw = df.iloc[ok_row, 7] if len(df.columns) > 7 else ""
            measurements['H'] = FileHandler.truncate_to_three_decimals(actual_H_raw)
            
            # Column I: truncated_H - nominal_H (use truncated value for calculation)
            if pd.notna(actual_H_raw) and actual_H_raw != "":
                try:
                    # Convert truncated string back to float for calculation
                    actual_H_truncated = float(measurements['H'])
                    measurements['I'] = FileHandler.truncate_to_three_decimals(actual_H_truncated - nominal_H)
                except (ValueError, TypeError):
                    measurements['I'] = ""
            else:
                measurements['I'] = ""
        except (IndexError, KeyError):
            measurements['H'] = ""
            measurements['I'] = ""
        
        # Column J: 0.000
        measurements['J'] = "0.000"
        
        # Add reference values only for first part across ALL files
        if is_first_part:
            measurements['K'] = FileHandler.truncate_to_three_decimals(reference_values['K'])
            measurements['L'] = FileHandler.truncate_to_three_decimals(reference_values['L'])
            measurements['M'] = FileHandler.truncate_to_three_decimals(reference_values['M'])
            measurements['P'] = FileHandler.truncate_to_three_decimals(reference_values['P'])
            measurements['Q'] = FileHandler.truncate_to_three_decimals(reference_values['Q'])
            measurements['R'] = FileHandler.truncate_to_three_decimals(reference_values['R'])
            measurements['U'] = FileHandler.truncate_to_three_decimals(reference_values['U'])
            measurements['V'] = FileHandler.truncate_to_three_decimals(reference_values['V'])
            measurements['W'] = ""  # Always blank
        
        # N, O, S, T should be blank for new format
        measurements['N'] = ""
        measurements['O'] = ""
        measurements['S'] = ""
        measurements['T'] = ""
        if not is_first_part:
            measurements['W'] = ""
        
        return measurements
    
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
        
        # Count parts correctly (works for both formats)
        part_position_dict = FileHandler.find_all_parts(data)
        total_parts = len(part_position_dict)
        
        sheet = workbook.active
        sheet_name = f'PO{customer_po} Qty {total_parts}'
        sheet.title = sheet_name
        
        sheet['B1'] = input_strings.get('part_name', '')
        sheet['B2'] = input_strings.get('revision_number', '')
        sheet['B3'] = input_strings.get('lot_number', '')
        sheet['B4'] = input_strings.get('customer_p/n', '')
        sheet['B5'] = customer_po
        sheet['B6'] = input_strings.get('measurement_units', '')
        
        workbook.save(output_path)
        workbook.close()
    
    @staticmethod
    def extract_serial_number_from_b5(df):
        """
        Extract serial number from B5 position (DataFrame index [4, 2])
        B5 corresponds to DataFrame row 4 (Excel row 5 - 1) and column 2 (Excel column C - 1)
        Note: After first column is popped, what was originally C5 is now at [4, 2]
        """
        try:
            if str(df.iloc[4, 2]) != "nan":
                serial_value = df.iloc[4, 2]
            else: 
                serial_value = str(df.iloc[9, 1]).split(':')[1]
            
            if pd.notna(serial_value) and str(serial_value).strip():
                serial_number = str(serial_value).strip()
                logger.info(f"Found serial number at B5: {serial_number}")
                return serial_number
            else:
                logger.warning("Serial number at B5 is empty or NaN")
                return ""
                
        except (IndexError, KeyError):
            logger.warning("Could not access B5 position for serial number")
            return ""
    
    @staticmethod
    def find_all_parts(data):
        """
        Find all parts in data, detecting format per file.
        Supports both old format (LOC1-based) and new format (OK-based).
        """
        if isinstance(data, list):
            all_positions = {}
            part_counter = 1
            is_very_first_part = True
            
            for df_idx, df in enumerate(data):
                logger.info(f"Scanning DataFrame {df_idx + 1}/{len(data)}...")
                
                # Detect format for this DataFrame
                if FileHandler.is_new_format(df):
                    logger.info(f"DataFrame {df_idx + 1} detected as NEW FORMAT")
                    # Extract reference values for new format
                    reference_values = FileHandler.extract_new_format_reference_values(df)
                    # Find all OK positions
                    df_positions = FileHandler.find_ok_positions_new_format(df, part_counter, df_idx, reference_values)
                else:
                    logger.info(f"DataFrame {df_idx + 1} detected as OLD FORMAT")
                    # Process as old format (existing logic)
                    serial_number = FileHandler.extract_serial_number_from_b5(df)
                    df_positions = FileHandler._find_header_loc1_in_single_df(df, part_counter, df_idx, serial_number)
                
                # Mark the very first part
                if is_very_first_part and len(df_positions) > 0:
                    first_part_key = list(df_positions.keys())[0]
                    df_positions[first_part_key]['is_very_first'] = True
                    is_very_first_part = False
                
                all_positions.update(df_positions)
                part_counter += len(df_positions)
                logger.info(f"Found {len(df_positions)} parts in DataFrame {df_idx + 1}")
            
            logger.info(f"Total parts found across all files: {len(all_positions)}")
            return all_positions
        else:
            # Single DataFrame
            if FileHandler.is_new_format(data):
                logger.info("Single DataFrame detected as NEW FORMAT")
                reference_values = FileHandler.extract_new_format_reference_values(data)
                positions = FileHandler.find_ok_positions_new_format(data, 1, 0, reference_values)
            else:
                logger.info("Single DataFrame detected as OLD FORMAT")
                serial_number = FileHandler.extract_serial_number_from_b5(data)
                positions = FileHandler._find_header_loc1_in_single_df(data, 1, 0, serial_number)
            
            # Mark first part
            if len(positions) > 0:
                first_part_key = list(positions.keys())[0]
                positions[first_part_key]['is_very_first'] = True
            
            return positions
    
    @staticmethod
    def find_header_loc1_positions(data):
        """
        DEPRECATED: Use find_all_parts instead.
        Kept for backward compatibility.
        """
        return FileHandler.find_all_parts(data)
    
    @staticmethod
    def _find_header_loc1_in_single_df(df, start_counter, df_index, serial_number=""):
        """
        Helper function to find HEADER LOC1 positions in a single DataFrame
        
        Based on DataFrame output and sheet range B2:V392:
        - DataFrame row 0 = Excel row 2
        - Index 0: Empty (original Excel column A)
        - Index 1: LOC1 data (original Excel column B) 
        - Index 2: UNITS data (original Excel column C)
        
        Returns dictionary with part info including DataFrame index and serial number
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
                            
                            # Store position, DataFrame index, serial number, and format
                            positions[f'part-{part_counter}'] = {
                                'position': excel_ref,
                                'df_index': df_index,
                                'serial_number': serial_number,
                                'format': 'old'
                            }
                            part_counter += 1
                            logger.debug(f"Found header LOC1 at {excel_ref} with UNITS at index 2 in DataFrame {df_index}, serial: {serial_number}")
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
    def extract_all_measurements_for_part(df, part_positions_dict, part_name, is_first_part=False):
        """Extract all measurements for a part in one pass - now includes serial number as first column"""
        measurements = {}
        
        # Get serial number for this part
        part_info = part_positions_dict[part_name]
        serial_number = part_info.get('serial_number', '') if isinstance(part_info, dict) else ''
        
        # Add serial number as column A
        measurements['A'] = serial_number
        
        # Base extractions for all parts (shifted right by 1 due to serial number)
        # For the first part only, include K, L, M (5 over 2 down, 7 over 2 down, 8 over 2 down)
        base_extractions = [
            ('B', 6, 2),    ('C', 10, 2),   ('D', 11, 2),
            ('E', 6, 10),   ('F', 10, 10),  ('G', 11, 10),
            ('H', 6, 14),   ('I', 10, 14),  ('J', 11, 14),
        ]
        
        # Only include K, L, M for the first part to avoid repetition
        if is_first_part:
            base_extractions.extend([
                ('K', 5, 2),    ('L', 7, 2),    ('M', 8, 2),
            ])
        
        # Part-1 specific extractions (only for part-1) - also shifted right
        # Original columns O,P,Q,T,U -> now P,Q,R,U,V
        part1_extractions = [
            ('P', 5, 10),   ('Q', 7, 10),   ('R', 8, 10),
            ('U', 5, 14),   ('V', 7, 14),
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
            
            # Calculate derived values (only for part-1) and truncate results (shifted right by 1)
            if isinstance(val_5_2, (int, float)) and isinstance(val_7_2, (int, float)):
                measurements['N'] = FileHandler.truncate_to_three_decimals(val_5_2 + val_7_2)
            else:
                measurements['N'] = ""
            
            if isinstance(val_5_2, (int, float)) and isinstance(val_8_2, (int, float)):
                measurements['O'] = FileHandler.truncate_to_three_decimals(val_5_2 - val_8_2)
            else:
                measurements['O'] = ""
            
            if isinstance(val_5_10, (int, float)) and isinstance(val_7_10, (int, float)):
                measurements['S'] = FileHandler.truncate_to_three_decimals(val_5_10 + val_7_10)
            else:
                measurements['S'] = ""
            
            if isinstance(val_5_10, (int, float)) and isinstance(val_8_10, (int, float)):
                measurements['T'] = FileHandler.truncate_to_three_decimals(val_5_10 - val_8_10)
            else:
                measurements['T'] = ""
            
            if isinstance(val_5_14, (int, float)) and isinstance(val_7_14, (int, float)):
                measurements['W'] = FileHandler.truncate_to_three_decimals(val_5_14 + val_7_14)
            else:
                measurements['W'] = ""
        
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
        Process measurements for a part from the specific DataFrame where it was found.
        Supports both old format (LOC1-based) and new format (OK-based).
        Returns the next available row number.
        """
        if not isinstance(data, list):
            data = [data]
        
        if part_name not in part_positions_dict:
            return start_row
        
        part_info = part_positions_dict[part_name]
        
        # Get format type and df_index
        if isinstance(part_info, dict):
            df_index = part_info.get('df_index', 0)
            format_type = part_info.get('format', 'old')
            is_very_first = part_info.get('is_very_first', False)
        else:
            # Fallback for old format without format field
            df_index = 0
            format_type = 'old'
            is_very_first = (part_name == 'part-1')
        
        # Only process the specific DataFrame that contains this part
        if df_index < len(data):
            df = data[df_index]
            
            # Extract measurements based on format
            if format_type == 'new':
                logger.info(f"Processing {part_name} using NEW FORMAT logic")
                measurements = FileHandler.extract_measurements_new_format(df, part_info, is_very_first)
            else:
                logger.info(f"Processing {part_name} using OLD FORMAT logic")
                measurements = FileHandler.extract_all_measurements_for_part(df, part_positions_dict, part_name, is_very_first)
            
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
            progress_callback("Finding parts (detecting format)...")
        
        # Use new unified method that detects format
        part_position_dict = FileHandler.find_all_parts(data)
        
        if not part_position_dict:
            raise ValueError("No parts found in the input data")
        
        total_parts = len(part_position_dict)
        logger.info(f"Found {total_parts} parts to process")
        
        if progress_callback:
            progress_callback(f"Processing {total_parts} parts...")
        
        with excel_batch_writer(output_path) as workbook:
            current_row = 9
            processed_parts = 0
            
            for part_name in part_position_dict.keys():
                logger.info(f"Processing {part_name}... ({processed_parts + 1}/{total_parts})")
                
                # Process part (format is automatically detected in process_part)
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