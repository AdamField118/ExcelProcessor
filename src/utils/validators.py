import os

class FileValidator:
    @staticmethod
    def is_valid_excel_file(file_path):
        """Check if file is a valid Excel file"""
        return file_path.lower().endswith(('.xls', '.xlsx'))
    
    @staticmethod
    def is_valid_csv_file(file_path):
        """Check if file is a valid CSV file"""
        return file_path.lower().endswith('.csv')
    
    @staticmethod
    def is_valid_data_file(file_path):
        """Check if file is valid Excel OR CSV"""
        return FileValidator.is_valid_excel_file(file_path) or FileValidator.is_valid_csv_file(file_path)
    
    @staticmethod
    def get_file_type(file_path):
        """Get human-readable file type description"""
        if file_path.lower().endswith('.xls'):
            return 'Excel (XLS)'
        elif file_path.lower().endswith('.xlsx'):
            return 'Excel (XLSX)'
        elif file_path.lower().endswith('.csv'):
            return 'CSV'
        else:
            return 'Unknown'
    
    @staticmethod
    def validate_file_accessibility(file_path):
        return os.path.exists(file_path) and os.access(file_path, os.R_OK)
    
    @staticmethod
    def validate_output_path(file_path):
        dir_path = os.path.dirname(file_path) or '.'
        return os.access(dir_path, os.W_OK)