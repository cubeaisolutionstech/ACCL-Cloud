# utils/excel_processor.py
import pandas as pd
from io import BytesIO
import openpyxl

class ExcelProcessor:
    def __init__(self, file_source):
        """
        Initialize ExcelProcessor with either file path or BytesIO object
        
        Args:
            file_source: Either a file path (string) or BytesIO object
        """
        self.file_source = file_source
        self.is_buffer = isinstance(file_source, BytesIO)
        
    def get_sheet_names(self):
        """Get all sheet names from the Excel file"""
        try:
            if self.is_buffer:
                # Reset buffer position
                self.file_source.seek(0)
                workbook = openpyxl.load_workbook(self.file_source, read_only=True)
                sheet_names = workbook.sheetnames
                workbook.close()
                return sheet_names
            else:
                # File path
                return pd.ExcelFile(self.file_source).sheet_names
        except Exception as e:
            raise Exception(f"Error reading sheet names: {str(e)}")
    
    def read_sheet(self, sheet_name, header=0):
        """
        Read a specific sheet from the Excel file
        
        Args:
            sheet_name: Name of the sheet to read
            header: Row to use as header (default 0)
            
        Returns:
            pandas DataFrame
        """
        try:
            if self.is_buffer:
                # Reset buffer position
                self.file_source.seek(0)
                return pd.read_excel(self.file_source, sheet_name=sheet_name, header=header)
            else:
                # File path
                return pd.read_excel(self.file_source, sheet_name=sheet_name, header=header)
        except Exception as e:
            raise Exception(f"Error reading sheet '{sheet_name}': {str(e)}")
    
    def read_all_sheets(self, header=0):
        """
        Read all sheets from the Excel file
        
        Args:
            header: Row to use as header (default 0)
            
        Returns:
            Dictionary of DataFrames {sheet_name: DataFrame}
        """
        try:
            if self.is_buffer:
                # Reset buffer position
                self.file_source.seek(0)
                return pd.read_excel(self.file_source, sheet_name=None, header=header)
            else:
                # File path
                return pd.read_excel(self.file_source, sheet_name=None, header=header)
        except Exception as e:
            raise Exception(f"Error reading all sheets: {str(e)}")
    
    def get_sheet_info(self, sheet_name):
        """
        Get basic information about a sheet
        
        Args:
            sheet_name: Name of the sheet
            
        Returns:
            Dictionary with sheet information
        """
        try:
            df = self.read_sheet(sheet_name, header=None)
            return {
                'sheet_name': sheet_name,
                'rows': len(df),
                'columns': len(df.columns),
                'shape': df.shape,
                'first_few_rows': df.head().to_dict('records')
            }
        except Exception as e:
            raise Exception(f"Error getting sheet info for '{sheet_name}': {str(e)}")
    
    def close(self):
        """Close the file source if it's a buffer"""
        if self.is_buffer and hasattr(self.file_source, 'close'):
            self.file_source.close()