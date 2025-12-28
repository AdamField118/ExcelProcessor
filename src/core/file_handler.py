"""
File Handler - Core file operations
Simplified version for ETS processor
"""

import pandas as pd
import os
import sys
import logging

logger = logging.getLogger("file_handler")

def check_xlrd():
    """Check xlrd availability"""
    try:
        import xlrd
        return True
    except ImportError:
        return False


class FileHandler:
    @staticmethod
    def load_excel_file(file_path):
        """Load Excel or CSV file into DataFrame"""
        logger.info(f"Loading: {os.path.basename(file_path)}")
        
        if file_path.lower().endswith('.csv'):
            # Try different encodings for CSV
            encodings = ['utf-8', 'utf-8-sig', 'latin1', 'cp1252']
            for encoding in encodings:
                try:
                    result = pd.read_excel(file_path, header=None, encoding=encoding)
                    result.attrs['source_file'] = os.path.basename(file_path)
                    return result
                except:
                    continue
            raise Exception(f"Could not load CSV file: {file_path}")
        
        xlrd_available = check_xlrd()
        
        try:
            if file_path.lower().endswith('.xls'):
                if not xlrd_available:
                    raise Exception("xlrd is not available for .xls files")
                result = pd.read_excel(file_path, header=None, engine='xlrd')
            else:
                result = pd.read_excel(file_path, header=None, engine='openpyxl')
            
            result.attrs['source_file'] = os.path.basename(file_path)
            return result
            
        except Exception as e:
            raise Exception(f"Failed to load {os.path.basename(file_path)}: {str(e)}")
    
    @staticmethod
    def open_file(file_path):
        """Open file in default application (Windows only)"""
        try:
            os.startfile(file_path)
            return True
        except Exception as e:
            logger.error(f"Failed to open file: {e}")
            return False