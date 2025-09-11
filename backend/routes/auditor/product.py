from flask import Blueprint, request, jsonify, current_app, session, send_file
import pandas as pd
import numpy as np
import re
from datetime import datetime
from io import BytesIO
import os
from werkzeug.utils import secure_filename
import base64

product_bp = Blueprint('product', __name__, url_prefix='/api/product')

class MergePreviewProcessor:
    """Handles the complete merge preview logic with EXACT column ordering"""
    
    def __init__(self):
        self.fiscal_info = self.calculate_fiscal_year()
        self.months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
        
        # Table headers for detection
        self.mt_table_headers = [
            "SALES in Tonage", "SALES IN TONAGE", "Tonage", "TONAGE",
            "Sales in MT", "SALES IN MT", "SALES in Ton", "Metric Tons", 
            "MT Sales", "Tonage Sales", "Sales Tonage"
        ]
        
        self.value_table_headers = [
            "SALES in Value", "SALES IN VALUE", "Sales in Rs", "SALES IN RS",
            "Value", "VALUE", "Sales Value"
        ]
        
        # Exclusion patterns for totals
        self.exclude_from_sort = ['TOTAL SALES', 'GRAND TOTAL', 'NORTH TOTAL', 'WEST SALES']
    
    def calculate_fiscal_year(self):
        """Calculate fiscal year information"""
        current_date = datetime.now()
        current_year = current_date.year
        
        if current_date.month >= 4:
            fiscal_year_start = current_year
            fiscal_year_end = current_year + 1
        else:
            fiscal_year_start = current_year - 1
            fiscal_year_end = current_year
            
        return {
            'fiscal_year_start': fiscal_year_start,
            'fiscal_year_end': fiscal_year_end,
            'fiscal_year_str': f"{str(fiscal_year_start)[-2:]}-{str(fiscal_year_end)[-2:]}",
            'last_fiscal_year_start': fiscal_year_start - 1,
            'last_fiscal_year_end': fiscal_year_end - 1,
            'last_fiscal_year_str': f"{str(fiscal_year_start - 1)[-2:]}-{str(fiscal_year_end - 1)[-2:]}"
        }
    
    def get_exact_column_positions(self):
        """Generate the EXACT column order as specified"""
        current_fy = self.fiscal_info['fiscal_year_str']  # e.g., "25-26"
        last_fy = self.fiscal_info['last_fiscal_year_str']  # e.g., "24-25"
        current_start_year = str(self.fiscal_info['fiscal_year_start'])[-2:]  # e.g., "25"
        current_end_year = str(self.fiscal_info['fiscal_year_end'])[-2:]  # e.g., "26"
        last_start_year = str(self.fiscal_info['last_fiscal_year_start'])[-2:]  # e.g., "24"
        last_end_year = str(self.fiscal_info['last_fiscal_year_end'])[-2:]  # e.g., "25"
        
        # EXACT ORDER as specified in your requirement
        exact_order = [
            # 1. Product column (position 0 - handled separately)
            
            # 2-6. April
            f'Budget-Apr-{current_start_year}',
            f'LY-Apr-{last_start_year}', 
            f'Act-Apr-{current_start_year}', 
            f'Gr-Apr-{current_start_year}', 
            f'Ach-Apr-{current_start_year}',
            
            # 7-11. May  
            f'Budget-May-{current_start_year}', 
            f'LY-May-{last_start_year}', 
            f'Act-May-{current_start_year}', 
            f'Gr-May-{current_start_year}', 
            f'Ach-May-{current_start_year}',
            
            # 12-16. June
            f'Budget-Jun-{current_start_year}', 
            f'LY-Jun-{last_start_year}', 
            f'Act-Jun-{current_start_year}', 
            f'Gr-Jun-{current_start_year}', 
            f'Ach-Jun-{current_start_year}',
            
            # 17-21. Q1 YTD (Apr to Jun)
            f'YTD-{current_fy} (Apr to Jun)Budget', 
            f'YTD-{last_fy} (Apr to Jun)LY', 
            f'Act-YTD-{current_fy} (Apr to Jun)', 
            f'Gr-YTD-{current_fy} (Apr to Jun)', 
            f'Ach-YTD-{current_fy} (Apr to Jun)',
            
            # 22-26. July
            f'Budget-Jul-{current_start_year}', 
            f'LY-Jul-{last_start_year}', 
            f'Act-Jul-{current_start_year}', 
            f'Gr-Jul-{current_start_year}', 
            f'Ach-Jul-{current_start_year}',
            
            # 27-31. August
            f'Budget-Aug-{current_start_year}', 
            f'LY-Aug-{last_start_year}', 
            f'Act-Aug-{current_start_year}', 
            f'Gr-Aug-{current_start_year}', 
            f'Ach-Aug-{current_start_year}',
            
            # 32-36. September
            f'Budget-Sep-{current_start_year}', 
            f'LY-Sep-{last_start_year}', 
            f'Act-Sep-{current_start_year}', 
            f'Gr-Sep-{current_start_year}', 
            f'Ach-Sep-{current_start_year}',
            
            # 37-41. H1 YTD (Apr to Sep)
            f'YTD-{current_fy} (Apr to Sep)Budget', 
            f'YTD-{last_fy} (Apr to Sep)LY', 
            f'Act-YTD-{current_fy} (Apr to Sep)', 
            f'Gr-YTD-{current_fy} (Apr to Sep)', 
            f'Ach-YTD-{current_fy} (Apr to Sep)',
            
            # 42-46. October
            f'Budget-Oct-{current_start_year}', 
            f'LY-Oct-{last_start_year}', 
            f'Act-Oct-{current_start_year}', 
            f'Gr-Oct-{current_start_year}', 
            f'Ach-Oct-{current_start_year}',
            
            # 47-51. November
            f'Budget-Nov-{current_start_year}', 
            f'LY-Nov-{last_start_year}', 
            f'Act-Nov-{current_start_year}', 
            f'Gr-Nov-{current_start_year}', 
            f'Ach-Nov-{current_start_year}',
            
            # 52-56. December
            f'Budget-Dec-{current_start_year}', 
            f'LY-Dec-{last_start_year}', 
            f'Act-Dec-{current_start_year}', 
            f'Gr-Dec-{current_start_year}', 
            f'Ach-Dec-{current_start_year}',
            
            # 57-61. 9M YTD (Apr to Dec)
            f'YTD-{current_fy} (Apr to Dec)Budget', 
            f'YTD-{last_fy} (Apr to Dec)LY', 
            f'Act-YTD-{current_fy} (Apr to Dec)', 
            f'Gr-YTD-{current_fy} (Apr to Dec)', 
            f'Ach-YTD-{current_fy} (Apr to Dec)',
            
            # 62-66. January
            f'Budget-Jan-{current_end_year}', 
            f'LY-Jan-{last_end_year}', 
            f'Act-Jan-{current_end_year}', 
            f'Gr-Jan-{current_end_year}', 
            f'Ach-Jan-{current_end_year}',
            
            # 67-71. February
            f'Budget-Feb-{current_end_year}', 
            f'LY-Feb-{last_end_year}', 
            f'Act-Feb-{current_end_year}', 
            f'Gr-Feb-{current_end_year}', 
            f'Ach-Feb-{current_end_year}',
            
            # 72-76. March
            f'Budget-Mar-{current_end_year}', 
            f'LY-Mar-{last_end_year}', 
            f'Act-Mar-{current_end_year}', 
            f'Gr-Mar-{current_end_year}', 
            f'Ach-Mar-{current_end_year}',
            
            # 77-82. Full Year YTD (Apr to Mar) - EXACTLY as specified with dash and extra Budget
            f'YTD-{current_fy} (Apr to Mar) Budget',  # Note the DASH before (Apr to Mar)
            f'YTD-{last_fy} (Apr to Mar)LY', 
            f'Act-YTD-{current_fy} (Apr to Mar)', 
            f'Gr-YTD-{current_fy} (Apr to Mar)', 
            f'Ach-YTD-{current_fy} (Apr to Mar)'
        
        ]
        
        return exact_order
    
    def get_exact_column_order_priority(self, col_name):
        """Get exact column order priority based on specified sequence"""
        if pd.isna(col_name) or col_name == '':
            return 99999
            
        col_str = str(col_name).strip()
        
        # First column (product name) gets highest priority
        if col_str in ['SALES in Tonage', 'SALES in Value']:
            return 0
        
        # Get the exact order list
        exact_order = self.get_exact_column_positions()
        
        # Find exact match in the order
        try:
            return exact_order.index(col_str) + 1
        except ValueError:
            # If not found in exact order, put at end
            return 99999
        
    def reorder_columns_exact_fiscal_year(self, df, product_col_name):
        """Reorder columns in EXACT fiscal year order as specified"""
        if df is None or df.empty:
            return df
            
        # Get all columns
        all_columns = df.columns.tolist()
        
        # Separate product column and other columns
        other_columns = [col for col in all_columns if col != product_col_name]
        
        # Sort other columns by exact fiscal year priority
        sorted_columns = sorted(other_columns, key=self.get_exact_column_order_priority)
        
        # Final column order: product column first, then sorted columns
        final_column_order = [product_col_name] + sorted_columns
        
        # Reorder the dataframe
        return df[final_column_order]

    def calculate_exact_ytd_periods(self):
        """Calculate YTD periods with EXACT column names as specified"""
        current_fy = self.fiscal_info['fiscal_year_str']
        last_fy = self.fiscal_info['last_fiscal_year_str']
        current_start_year = str(self.fiscal_info['fiscal_year_start'])[-2:]
        current_end_year = str(self.fiscal_info['fiscal_year_end'])[-2:]
        last_start_year = str(self.fiscal_info['last_fiscal_year_start'])[-2:]
        last_end_year = str(self.fiscal_info['last_fiscal_year_end'])[-2:]
        
        ytd_periods = {}
        
        # Q1 YTD (Apr to Jun) - EXACT column names
        ytd_periods[f'YTD-{current_fy} (Apr to Jun)Budget'] = [
            f'Budget-Apr-{current_start_year}', 
            f'Budget-May-{current_start_year}', 
            f'Budget-Jun-{current_start_year}'
        ]
        ytd_periods[f'YTD-{last_fy} (Apr to Jun)LY'] = [
            f'LY-Apr-{last_start_year}', 
            f'LY-May-{last_start_year}', 
            f'LY-Jun-{last_start_year}'
        ]
        ytd_periods[f'Act-YTD-{current_fy} (Apr to Jun)'] = [
            f'Act-Apr-{current_start_year}', 
            f'Act-May-{current_start_year}', 
            f'Act-Jun-{current_start_year}'
        ]
        
        # H1 YTD (Apr to Sep) - EXACT column names
        ytd_periods[f'YTD-{current_fy} (Apr to Sep)Budget'] = [
            f'Budget-Apr-{current_start_year}', f'Budget-May-{current_start_year}', f'Budget-Jun-{current_start_year}',
            f'Budget-Jul-{current_start_year}', f'Budget-Aug-{current_start_year}', f'Budget-Sep-{current_start_year}'
        ]
        ytd_periods[f'YTD-{last_fy} (Apr to Sep)LY'] = [
            f'LY-Apr-{last_start_year}', f'LY-May-{last_start_year}', f'LY-Jun-{last_start_year}',
            f'LY-Jul-{last_start_year}', f'LY-Aug-{last_start_year}', f'LY-Sep-{last_start_year}'
        ]
        ytd_periods[f'Act-YTD-{current_fy} (Apr to Sep)'] = [
            f'Act-Apr-{current_start_year}', f'Act-May-{current_start_year}', f'Act-Jun-{current_start_year}',
            f'Act-Jul-{current_start_year}', f'Act-Aug-{current_start_year}', f'Act-Sep-{current_start_year}'
        ]
        
        # 9M YTD (Apr to Dec) - EXACT column names
        ytd_periods[f'YTD-{current_fy} (Apr to Dec)Budget'] = [
            f'Budget-Apr-{current_start_year}', f'Budget-May-{current_start_year}', f'Budget-Jun-{current_start_year}',
            f'Budget-Jul-{current_start_year}', f'Budget-Aug-{current_start_year}', f'Budget-Sep-{current_start_year}',
            f'Budget-Oct-{current_start_year}', f'Budget-Nov-{current_start_year}', f'Budget-Dec-{current_start_year}'
        ]
        ytd_periods[f'YTD-{last_fy} (Apr to Dec)LY'] = [
            f'LY-Apr-{last_start_year}', f'LY-May-{last_start_year}', f'LY-Jun-{last_start_year}',
            f'LY-Jul-{last_start_year}', f'LY-Aug-{last_start_year}', f'LY-Sep-{last_start_year}',
            f'LY-Oct-{last_start_year}', f'LY-Nov-{last_start_year}', f'LY-Dec-{last_start_year}'
        ]
        ytd_periods[f'Act-YTD-{current_fy} (Apr to Dec)'] = [
            f'Act-Apr-{current_start_year}', f'Act-May-{current_start_year}', f'Act-Jun-{current_start_year}',
            f'Act-Jul-{current_start_year}', f'Act-Aug-{current_start_year}', f'Act-Sep-{current_start_year}',
            f'Act-Oct-{current_start_year}', f'Act-Nov-{current_start_year}', f'Act-Dec-{current_start_year}'
        ]
        
        # Full Year YTD (Apr to Mar) - EXACT column names with DASH and extra Budget
        ytd_periods[f'YTD-{current_fy} (Apr to Mar) Budget'] = [  # Note the DASH
            f'Budget-Apr-{current_start_year}', f'Budget-May-{current_start_year}', f'Budget-Jun-{current_start_year}',
            f'Budget-Jul-{current_start_year}', f'Budget-Aug-{current_start_year}', f'Budget-Sep-{current_start_year}',
            f'Budget-Oct-{current_start_year}', f'Budget-Nov-{current_start_year}', f'Budget-Dec-{current_start_year}',
            f'Budget-Jan-{current_end_year}', f'Budget-Feb-{current_end_year}', f'Budget-Mar-{current_end_year}'
        ]
        ytd_periods[f'YTD-{last_fy} (Apr to Mar)LY'] = [
            f'LY-Apr-{last_start_year}', f'LY-May-{last_start_year}', f'LY-Jun-{last_start_year}',
            f'LY-Jul-{last_start_year}', f'LY-Aug-{last_start_year}', f'LY-Sep-{last_start_year}',
            f'LY-Oct-{last_start_year}', f'LY-Nov-{last_start_year}', f'LY-Dec-{last_start_year}',
            f'LY-Jan-{last_end_year}', f'LY-Feb-{last_end_year}', f'LY-Mar-{last_end_year}'
        ]
        ytd_periods[f'Act-YTD-{current_fy} (Apr to Mar)'] = [
            f'Act-Apr-{current_start_year}', f'Act-May-{current_start_year}', f'Act-Jun-{current_start_year}',
            f'Act-Jul-{current_start_year}', f'Act-Aug-{current_start_year}', f'Act-Sep-{current_start_year}',
            f'Act-Oct-{current_start_year}', f'Act-Nov-{current_start_year}', f'Act-Dec-{current_start_year}',
            f'Act-Jan-{current_end_year}', f'Act-Feb-{current_end_year}', f'Act-Mar-{current_end_year}'
        ]
        
        return ytd_periods
    
    def calculate_ytd_growth_achievement(self, merged_data):
        """Calculate YTD Growth and Achievement with EXACT column names - ONLY when Act values exist"""
        current_fy = self.fiscal_info['fiscal_year_str']
        last_fy = self.fiscal_info['last_fiscal_year_str']
        
        ytd_pairs = [
            ('Apr to Jun', f'YTD-{current_fy} (Apr to Jun)Budget', 
             f'YTD-{last_fy} (Apr to Jun)LY', f'Act-YTD-{current_fy} (Apr to Jun)'),
            ('Apr to Sep', f'YTD-{current_fy} (Apr to Sep)Budget', 
             f'YTD-{last_fy} (Apr to Sep)LY', f'Act-YTD-{current_fy} (Apr to Sep)'),
            ('Apr to Dec', f'YTD-{current_fy} (Apr to Dec)Budget', 
             f'YTD-{last_fy} (Apr to Dec)LY', f'Act-YTD-{current_fy} (Apr to Dec)'),
            ('Apr to Mar', f'YTD-{current_fy} (Apr to Mar) Budget',  # Note the DASH
             f'YTD-{last_fy} (Apr to Mar)LY', f'Act-YTD-{current_fy} (Apr to Mar)')
        ]
        
        for period, budget_col, ly_col, act_col in ytd_pairs:
            # Check if Act column exists and has actual values
            if act_col in merged_data.columns:
                # Check if there are any non-zero Act values
                has_act_values = (merged_data[act_col] > 0.01).any()
                
                if has_act_values:
                    # Growth calculation (if both LY and Actual exist)
                    if ly_col in merged_data.columns:
                        gr_col = f'Gr-YTD-{current_fy} ({period})'
                        merged_data[gr_col] = np.where(
                            merged_data[ly_col] > 0.01,
                            ((merged_data[act_col] - merged_data[ly_col]) / merged_data[ly_col] * 100).round(2),
                            0
                        )
                        current_app.logger.info(f"Calculated Growth column: {gr_col} (Act values present)")
                    
                    # Achievement calculation (if both Budget and Actual exist)
                    if budget_col in merged_data.columns:
                        ach_col = f'Ach-YTD-{current_fy} ({period})'
                        merged_data[ach_col] = np.where(
                            merged_data[budget_col] > 0.01,
                            (merged_data[act_col] / merged_data[budget_col] * 100).round(2),
                            0
                        )
                        current_app.logger.info(f"Calculated Achievement column: {ach_col} (Act values present)")
                else:
                    # Initialize Growth and Achievement columns with zeros if no Act values
                    gr_col = f'Gr-YTD-{current_fy} ({period})'
                    ach_col = f'Ach-YTD-{current_fy} ({period})'
                    
                    if gr_col not in merged_data.columns:
                        merged_data[gr_col] = 0.0
                    if ach_col not in merged_data.columns:
                        merged_data[ach_col] = 0.0
                    
                    current_app.logger.info(f"Initialized Gr/Ach columns with zeros for {period} (no Act values)")
            else:
                current_app.logger.warning(f"Act column {act_col} not found for period {period}")
        
        return merged_data

    # Rest of the MergePreviewProcessor methods remain the same...
    def detect_product_sheet(self, auditor_filepath):
        """Detect product analysis sheet in auditor file"""
        try:
            xls_auditor = pd.ExcelFile(auditor_filepath)
            sheet_names = xls_auditor.sheet_names
            
            # Look for product-related sheet
            for sheet in sheet_names:
                if 'product' in sheet.lower():
                    return sheet, sheet_names
            
            # If no product sheet found, return first sheet
            return sheet_names[0] if sheet_names else None, sheet_names
            
        except Exception as e:
            current_app.logger.error(f"Error detecting product sheet: {str(e)}")
            return None, []
    
    def extract_tables(self, df_auditor, table_headers, is_product_analysis=True):
        """Extract table location from auditor data"""
        table_idx = None
        data_start_idx = None
        
        for idx, row in df_auditor.iterrows():
            # Convert row to string for searching
            row_values = []
            for val in row:
                if pd.notna(val):
                    row_values.append(str(val).upper())
            row_str = ' '.join(row_values)
            
            # Check if any header matches
            for header in table_headers:
                if header.upper() in row_str:
                    table_idx = idx
                    data_start_idx = idx + 1
                    break
            if table_idx is not None:
                break
        
        return table_idx, data_start_idx
    
    def smart_product_sorting(self, all_products):
        """Smart product sorting - regular products first, then totals"""
        regular_products = sorted([p for p in all_products if p not in self.exclude_from_sort])
        total_products = [p for p in all_products if p in self.exclude_from_sort]
        return regular_products + total_products
    
    
    def recalculate_totals(self, merged_data, product_col_name):
        """Recalculate total rows including ALL columns (Budget, LY, Act, Gr, Ach) - FIXED VERSION"""
        if 'TOTAL SALES' in merged_data[product_col_name].values:
            numeric_cols = merged_data.select_dtypes(include=[np.number]).columns
        
            # Calculate sum for ALL columns including Gr and Ach
            mask = ~merged_data[product_col_name].isin(self.exclude_from_sort)
            valid_products = merged_data[mask]
        
            for col in numeric_cols:
                # FIXED: For ALL numeric columns, calculate the sum (including Gr and Ach)
                sum_value = valid_products[col].sum()
                merged_data.loc[
                     merged_data[product_col_name] == 'TOTAL SALES', col
                ] = round(sum_value, 2)
                
                current_app.logger.info(f"Total calculated for {col}: {sum_value:.2f}")
    
        return merged_data
    

# Initialize the processor
merge_processor = MergePreviewProcessor()

def build_exact_columns_and_calculate_values(data_df, fiscal_info, analysis_type='mt'):
    """Build exact column structure and calculate all values - ONLY when Act values present"""
    
    # Get fiscal year components
    current_fy = fiscal_info['fiscal_year_str']
    last_fy = fiscal_info['last_fiscal_year_str'] 
    current_start_year = str(fiscal_info['fiscal_year_start'])[-2:]
    current_end_year = str(fiscal_info['fiscal_year_end'])[-2:]
    last_start_year = str(fiscal_info['last_fiscal_year_start'])[-2:]
    last_end_year = str(fiscal_info['last_fiscal_year_end'])[-2:]
    
    # Create all monthly columns in exact order
    months_data = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
    
    for i, month in enumerate(months_data):
        # Determine year based on month position in fiscal year
        if month in ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']:
            current_year = current_start_year
            last_year = last_start_year
        else:  # Jan, Feb, Mar
            current_year = current_end_year
            last_year = last_end_year
        
        # Create column names in exact order: Budget, LY, Act, Gr, Ach
        budget_col = f'Budget-{month}-{current_year}'
        ly_col = f'LY-{month}-{last_year}'
        act_col = f'Act-{month}-{current_year}'
        gr_col = f'Gr-{month}-{current_year}'
        ach_col = f'Ach-{month}-{current_year}'
        
        # Initialize columns if they don't exist
        for col in [budget_col, ly_col, act_col, gr_col, ach_col]:
            if col not in data_df.columns:
                data_df[col] = 0.0
        
        # Check if Act column has actual values before calculating Gr and Ach
        has_act_values = (data_df[act_col] > 0.01).any()
        
        if has_act_values:
            # Calculate Growth and Achievement for monthly data ONLY when Act values exist
            data_df[gr_col] = np.where(
                data_df[ly_col] > 0.01,
                ((data_df[act_col] - data_df[ly_col]) / data_df[ly_col] * 100).round(2),
                0
            )
            
            data_df[ach_col] = np.where(
                data_df[budget_col] > 0.01,
                (data_df[act_col] / data_df[budget_col] * 100).round(2),
                0
            )
            
            current_app.logger.info(f"Calculated Gr/Ach for {month}-{current_year} (Act values present)")
        else:
            # Keep Gr and Ach as 0 if no Act values
            current_app.logger.info(f"Skipped Gr/Ach calculation for {month}-{current_year} (no Act values)")
    
    # Calculate YTD columns in exact order and position
    
    # Q1 YTD (Apr to Jun) - Position after June columns
    q1_budget_col = f'YTD-{current_fy} (Apr to Jun)Budget'
    q1_ly_col = f'YTD-{last_fy} (Apr to Jun)LY'
    q1_act_col = f'Act-YTD-{current_fy} (Apr to Jun)'
    q1_gr_col = f'Gr-YTD-{current_fy} (Apr to Jun)'
    q1_ach_col = f'Ach-YTD-{current_fy} (Apr to Jun)'
    
    data_df[q1_budget_col] = (
        data_df.get(f'Budget-Apr-{current_start_year}', 0) +
        data_df.get(f'Budget-May-{current_start_year}', 0) +
        data_df.get(f'Budget-Jun-{current_start_year}', 0)
    )
    data_df[q1_ly_col] = (
        data_df.get(f'LY-Apr-{last_start_year}', 0) +
        data_df.get(f'LY-May-{last_start_year}', 0) +
        data_df.get(f'LY-Jun-{last_start_year}', 0)
    )
    data_df[q1_act_col] = (
        data_df.get(f'Act-Apr-{current_start_year}', 0) +
        data_df.get(f'Act-May-{current_start_year}', 0) +
        data_df.get(f'Act-Jun-{current_start_year}', 0)
    )
    
    # Check if Q1 Act YTD has values before calculating Gr and Ach
    has_q1_act_values = (data_df[q1_act_col] > 0.01).any()
    
    if has_q1_act_values:
        data_df[q1_gr_col] = np.where(
            data_df[q1_ly_col] > 0.01,
            ((data_df[q1_act_col] - data_df[q1_ly_col]) / data_df[q1_ly_col] * 100).round(2),
            0
        )
        data_df[q1_ach_col] = np.where(
            data_df[q1_budget_col] > 0.01,
            (data_df[q1_act_col] / data_df[q1_budget_col] * 100).round(2),
            0
        )
        current_app.logger.info("Calculated Q1 YTD Gr/Ach (Act values present)")
    else:
        data_df[q1_gr_col] = 0.0
        data_df[q1_ach_col] = 0.0
        current_app.logger.info("Skipped Q1 YTD Gr/Ach calculation (no Act values)")
    
    # H1 YTD (Apr to Sep) - Position after September columns
    h1_budget_col = f'YTD-{current_fy} (Apr to Sep)Budget'
    h1_ly_col = f'YTD-{last_fy} (Apr to Sep)LY'
    h1_act_col = f'Act-YTD-{current_fy} (Apr to Sep)'
    h1_gr_col = f'Gr-YTD-{current_fy} (Apr to Sep)'
    h1_ach_col = f'Ach-YTD-{current_fy} (Apr to Sep)'
    
    data_df[h1_budget_col] = (
        data_df.get(f'Budget-Apr-{current_start_year}', 0) +
        data_df.get(f'Budget-May-{current_start_year}', 0) +
        data_df.get(f'Budget-Jun-{current_start_year}', 0) +
        data_df.get(f'Budget-Jul-{current_start_year}', 0) +
        data_df.get(f'Budget-Aug-{current_start_year}', 0) +
        data_df.get(f'Budget-Sep-{current_start_year}', 0)
    )
    data_df[h1_ly_col] = (
        data_df.get(f'LY-Apr-{last_start_year}', 0) +
        data_df.get(f'LY-May-{last_start_year}', 0) +
        data_df.get(f'LY-Jun-{last_start_year}', 0) +
        data_df.get(f'LY-Jul-{last_start_year}', 0) +
        data_df.get(f'LY-Aug-{last_start_year}', 0) +
        data_df.get(f'LY-Sep-{last_start_year}', 0)
    )
    data_df[h1_act_col] = (
        data_df.get(f'Act-Apr-{current_start_year}', 0) +
        data_df.get(f'Act-May-{current_start_year}', 0) +
        data_df.get(f'Act-Jun-{current_start_year}', 0) +
        data_df.get(f'Act-Jul-{current_start_year}', 0) +
        data_df.get(f'Act-Aug-{current_start_year}', 0) +
        data_df.get(f'Act-Sep-{current_start_year}', 0)
    )
    
    # Check if H1 Act YTD has values before calculating Gr and Ach
    has_h1_act_values = (data_df[h1_act_col] > 0.01).any()
    
    if has_h1_act_values:
        data_df[h1_gr_col] = np.where(
            data_df[h1_ly_col] > 0.01,
            ((data_df[h1_act_col] - data_df[h1_ly_col]) / data_df[h1_ly_col] * 100).round(2),
            0
        )
        data_df[h1_ach_col] = np.where(
            data_df[h1_budget_col] > 0.01,
            (data_df[h1_act_col] / data_df[h1_budget_col] * 100).round(2),
            0
        )
        current_app.logger.info("Calculated H1 YTD Gr/Ach (Act values present)")
    else:
        data_df[h1_gr_col] = 0.0
        data_df[h1_ach_col] = 0.0
        current_app.logger.info("Skipped H1 YTD Gr/Ach calculation (no Act values)")
    
    # 9M YTD (Apr to Dec) - Position after December columns
    nm_budget_col = f'YTD-{current_fy} (Apr to Dec)Budget'
    nm_ly_col = f'YTD-{last_fy} (Apr to Dec)LY'
    nm_act_col = f'Act-YTD-{current_fy} (Apr to Dec)'
    nm_gr_col = f'Gr-YTD-{current_fy} (Apr to Dec)'
    nm_ach_col = f'Ach-YTD-{current_fy} (Apr to Dec)'
    
    data_df[nm_budget_col] = (
        data_df.get(f'Budget-Apr-{current_start_year}', 0) +
        data_df.get(f'Budget-May-{current_start_year}', 0) +
        data_df.get(f'Budget-Jun-{current_start_year}', 0) +
        data_df.get(f'Budget-Jul-{current_start_year}', 0) +
        data_df.get(f'Budget-Aug-{current_start_year}', 0) +
        data_df.get(f'Budget-Sep-{current_start_year}', 0) +
        data_df.get(f'Budget-Oct-{current_start_year}', 0) +
        data_df.get(f'Budget-Nov-{current_start_year}', 0) +
        data_df.get(f'Budget-Dec-{current_start_year}', 0)
    )
    data_df[nm_ly_col] = (
        data_df.get(f'LY-Apr-{last_start_year}', 0) +
        data_df.get(f'LY-May-{last_start_year}', 0) +
        data_df.get(f'LY-Jun-{last_start_year}', 0) +
        data_df.get(f'LY-Jul-{last_start_year}', 0) +
        data_df.get(f'LY-Aug-{last_start_year}', 0) +
        data_df.get(f'LY-Sep-{last_start_year}', 0) +
        data_df.get(f'LY-Oct-{last_start_year}', 0) +
        data_df.get(f'LY-Nov-{last_start_year}', 0) +
        data_df.get(f'LY-Dec-{last_start_year}', 0)
    )
    data_df[nm_act_col] = (
        data_df.get(f'Act-Apr-{current_start_year}', 0) +
        data_df.get(f'Act-May-{current_start_year}', 0) +
        data_df.get(f'Act-Jun-{current_start_year}', 0) +
        data_df.get(f'Act-Jul-{current_start_year}', 0) +
        data_df.get(f'Act-Aug-{current_start_year}', 0) +
        data_df.get(f'Act-Sep-{current_start_year}', 0) +
        data_df.get(f'Act-Oct-{current_start_year}', 0) +
        data_df.get(f'Act-Nov-{current_start_year}', 0) +
        data_df.get(f'Act-Dec-{current_start_year}', 0)
    )
    
    # Check if 9M Act YTD has values before calculating Gr and Ach
    has_nm_act_values = (data_df[nm_act_col] > 0.01).any()
    
    if has_nm_act_values:
        data_df[nm_gr_col] = np.where(
            data_df[nm_ly_col] > 0.01,
            ((data_df[nm_act_col] - data_df[nm_ly_col]) / data_df[nm_ly_col] * 100).round(2),
            0
        )
        data_df[nm_ach_col] = np.where(
            data_df[nm_budget_col] > 0.01,
            (data_df[nm_act_col] / data_df[nm_budget_col] * 100).round(2),
            0
        )
        current_app.logger.info("Calculated 9M YTD Gr/Ach (Act values present)")
    else:
        data_df[nm_gr_col] = 0.0
        data_df[nm_ach_col] = 0.0
        current_app.logger.info("Skipped 9M YTD Gr/Ach calculation (no Act values)")

    # Full Year YTD (Apr to Mar) - Position after March columns - EXACT as specified
    fy_budget_col = f'YTD-{current_fy} (Apr to Mar) Budget'  # Note the DASH
    fy_ly_col = f'YTD-{last_fy} (Apr to Mar)LY'
    fy_act_col = f'Act-YTD-{current_fy} (Apr to Mar)'
    fy_gr_col = f'Gr-YTD-{current_fy} (Apr to Mar)'
    fy_ach_col = f'Ach-YTD-{current_fy} (Apr to Mar)'
    
    data_df[fy_budget_col] = (
        data_df.get(f'Budget-Apr-{current_start_year}', 0) +
        data_df.get(f'Budget-May-{current_start_year}', 0) +
        data_df.get(f'Budget-Jun-{current_start_year}', 0) +
        data_df.get(f'Budget-Jul-{current_start_year}', 0) +
        data_df.get(f'Budget-Aug-{current_start_year}', 0) +
        data_df.get(f'Budget-Sep-{current_start_year}', 0) +
        data_df.get(f'Budget-Oct-{current_start_year}', 0) +
        data_df.get(f'Budget-Nov-{current_start_year}', 0) +
        data_df.get(f'Budget-Dec-{current_start_year}', 0) +
        data_df.get(f'Budget-Jan-{current_end_year}', 0) +
        data_df.get(f'Budget-Feb-{current_end_year}', 0) +
        data_df.get(f'Budget-Mar-{current_end_year}', 0)
    )
    data_df[fy_ly_col] = (
        data_df.get(f'LY-Apr-{last_start_year}', 0) +
        data_df.get(f'LY-May-{last_start_year}', 0) +
        data_df.get(f'LY-Jun-{last_start_year}', 0) +
        data_df.get(f'LY-Jul-{last_start_year}', 0) +
        data_df.get(f'LY-Aug-{last_start_year}', 0) +
        data_df.get(f'LY-Sep-{last_start_year}', 0) +
        data_df.get(f'LY-Oct-{last_start_year}', 0) +
        data_df.get(f'LY-Nov-{last_start_year}', 0) +
        data_df.get(f'LY-Dec-{last_start_year}', 0) +
        data_df.get(f'LY-Jan-{last_end_year}', 0) +
        data_df.get(f'LY-Feb-{last_end_year}', 0) +
        data_df.get(f'LY-Mar-{last_end_year}', 0)
    )
    data_df[fy_act_col] = (
        data_df.get(f'Act-Apr-{current_start_year}', 0) +
        data_df.get(f'Act-May-{current_start_year}', 0) +
        data_df.get(f'Act-Jun-{current_start_year}', 0) +
        data_df.get(f'Act-Jul-{current_start_year}', 0) +
        data_df.get(f'Act-Aug-{current_start_year}', 0) +
        data_df.get(f'Act-Sep-{current_start_year}', 0) +
        data_df.get(f'Act-Oct-{current_start_year}', 0) +
        data_df.get(f'Act-Nov-{current_start_year}', 0) +
        data_df.get(f'Act-Dec-{current_start_year}', 0) +
        data_df.get(f'Act-Jan-{current_end_year}', 0) +
        data_df.get(f'Act-Feb-{current_end_year}', 0) +
        data_df.get(f'Act-Mar-{current_end_year}', 0)
    )
    
    # Check if Full Year Act YTD has values before calculating Gr and Ach
    has_fy_act_values = (data_df[fy_act_col] > 0.01).any()
    
    if has_fy_act_values:
        data_df[fy_gr_col] = np.where(
            data_df[fy_ly_col] > 0.01,
            ((data_df[fy_act_col] - data_df[fy_ly_col]) / data_df[fy_ly_col] * 100).round(2),
            0
        )
        data_df[fy_ach_col] = np.where(
            data_df[fy_budget_col] > 0.01,
            (data_df[fy_act_col] / data_df[fy_budget_col] * 100).round(2),
            0
        )
        current_app.logger.info("Calculated Full Year YTD Gr/Ach (Act values present)")
    else:
        data_df[fy_gr_col] = 0.0
        data_df[fy_ach_col] = 0.0
        current_app.logger.info("Skipped Full Year YTD Gr/Ach calculation (no Act values)")
    
    return data_df

# Session Storage Functions for Sales Monthwise Integration
def store_totals_in_session(data_dict, data_type='mt'):
    """
    Store product analysis totals in session for use by sales monthwise
    data_dict: Dictionary from product analysis (with 'data' array)
    data_type: 'mt' or 'value'
    """
    try:
        if not data_dict or not data_dict.get('data'):
            current_app.logger.warning(f"No data to store in session for {data_type}")
            return False
        
        # Find the TOTAL SALES row in the data
        total_sales_row = None
        first_column = data_dict['columns'][0] if data_dict.get('columns') else None
        
        if not first_column:
            current_app.logger.warning(f"No columns found in {data_type} data")
            return False
        
        for row in data_dict['data']:
            product_name = row.get(first_column, '')
            if isinstance(product_name, str) and 'TOTAL SALES' in product_name.upper():
                total_sales_row = row
                break
        
        if not total_sales_row:
            current_app.logger.warning(f"No TOTAL SALES row found in {data_type} data")
            return False
        
        # Initialize session storage if it doesn't exist
        if 'total_sales_data' not in session:
            session['total_sales_data'] = {}
        
        # Store the total sales row data
        session_key = 'tonnage' if data_type.lower() == 'mt' else 'value'
        session['total_sales_data'][session_key] = {
            'data': total_sales_row,
            'columns': data_dict['columns'],
            'timestamp': datetime.now().isoformat(),
            'source': 'product_analysis'
        }
        
        # Mark session as modified
        session.modified = True
        
        current_app.logger.info(f"Stored {data_type} totals in session with {len(total_sales_row)} columns")
        return True
        
    except Exception as e:
        current_app.logger.error(f"Error storing {data_type} totals in session: {str(e)}")
        return False

# Utility functions
def handle_duplicate_columns(df):
    """Handle duplicate column names by adding suffix"""
    cols = df.columns.tolist()
    seen = {}
    for i, col in enumerate(cols):
        if col in seen:
            seen[col] += 1
            cols[i] = f"{col}_{seen[col]}"
        else:
            seen[col] = 0
    df.columns = cols
    return df

def find_column(df, possible_names, threshold=80, case_sensitive=False):
    """Find column by fuzzy matching"""
    from difflib import SequenceMatcher
    
    def similarity(a, b):
        return SequenceMatcher(None, a, b).ratio() * 100
    
    best_match = None
    best_score = 0
    
    for col in df.columns:
        for name in possible_names:
            if not case_sensitive:
                score = similarity(str(col).lower(), str(name).lower())
            else:
                score = similarity(str(col), str(name))
            
            if score >= threshold and score > best_score:
                best_score = score
                best_match = col
    
    return best_match

def rename_columns(columns):
    """Rename columns to standard format"""
    renamed = []
    for col in columns:
        if pd.isna(col) or str(col).strip() == '':
            renamed.append('Unnamed')
        else:
            col_str = str(col).strip()
            renamed.append(col_str)
    return renamed

def clean_and_convert_numeric(df):
    """Clean and convert numeric columns"""
    for col in df.columns[1:]:  # Skip first column (usually product names)
        df[col] = pd.to_numeric(df[col], errors='coerce')
        df[col] = df[col].fillna(0)
    return df

def detect_analysis_type(sheet_name):
    """Detect the type of analysis based on sheet name"""
    sheet_lower = sheet_name.lower()
    
    is_region_analysis = 'region' in sheet_lower
    is_product_analysis = 'product' in sheet_lower or 'ts-pw' in sheet_lower or 'ero-pw' in sheet_lower
    is_sales_analysis_month_wise = bool(re.search(r'sales\s*analysis\s*month\s*wise', sheet_lower))
    
    return {
        'is_region_analysis': is_region_analysis,
        'is_product_analysis': is_product_analysis,
        'is_sales_analysis_month_wise': is_sales_analysis_month_wise
    }

def extract_auditor_tables(df_auditor, table_headers, is_product_analysis=True):
    """Extract tables from auditor data"""
    table_idx = None
    data_start_idx = None
    
    for idx, row in df_auditor.iterrows():
        row_values = []
        for val in row:
            if pd.notna(val):
                row_values.append(str(val).upper())
        row_str = ' '.join(row_values)
        
        for header in table_headers:
            if header.upper() in row_str:
                table_idx = idx
                data_start_idx = idx + 1
                break
        if table_idx is not None:
            break
    
    return table_idx, data_start_idx

def process_auditor_table(df_auditor, table_headers, end_row=None):
    """Process auditor table and return cleaned DataFrame"""
    try:
        table_idx, data_start_idx = extract_auditor_tables(df_auditor, table_headers)
        
        if table_idx is None:
            return None, f"No table found with headers: {table_headers}"
        
        if end_row is not None:
            table_df = df_auditor.iloc[data_start_idx:end_row].dropna(how='all').reset_index(drop=True)
        else:
            table_df = df_auditor.iloc[data_start_idx:].dropna(how='all').reset_index(drop=True)
        
        if table_df.empty:
            return None, "Extracted table is empty"
        
        original_headers = df_auditor.iloc[table_idx].tolist()
        table_df.columns = [str(col) for col in original_headers]
        
        # Processing pipeline
        new_column_names = rename_columns(table_df.columns.tolist())
        table_df.columns = new_column_names
        table_df = handle_duplicate_columns(table_df)
        table_df = clean_and_convert_numeric(table_df)
        
        # Filter out summary rows
        first_col = table_df.columns[0]
        table_df = table_df[~table_df[first_col].astype(str).str.upper().str.contains('TOTAL|GRAND|SUMMARY', na=False)]
        
        # Clean first column (product names)
        table_df[first_col] = table_df[first_col].replace([pd.NA, None, ''], 'Unknown').fillna('Unknown')
        table_df[first_col] = table_df[first_col].astype(str).str.strip().str.upper()
        
        return table_df, None
        
    except Exception as e:
        return None, f"Error processing auditor table: {str(e)}"

def calculate_fiscal_year():
    """Calculate current fiscal year"""
    current_date = datetime.now()
    current_year = current_date.year
    if current_date.month >= 4:
        fiscal_year_start = current_year
        fiscal_year_end = current_year + 1
    else:
        fiscal_year_start = current_year - 1
        fiscal_year_end = current_year

    fiscal_year_str = f"{str(fiscal_year_start)[-2:]}-{str(fiscal_year_end)[-2:]}"
    last_fiscal_year_start = fiscal_year_start - 1
    last_fiscal_year_end = fiscal_year_end - 1
    last_fiscal_year_str = f"{str(last_fiscal_year_start)[-2:]}-{str(last_fiscal_year_end)[-2:]}"
    
    return {
        'fiscal_year_start': fiscal_year_start,
        'fiscal_year_end': fiscal_year_end,
        'fiscal_year_str': fiscal_year_str,
        'last_fiscal_year_start': last_fiscal_year_start,
        'last_fiscal_year_end': last_fiscal_year_end,
        'last_fiscal_year_str': last_fiscal_year_str
    }

def process_budget_data_product(budget_df, group_type='product'):
    """Process budget data for product analysis - FIXED to exclude invalid columns"""
    budget_df = handle_duplicate_columns(budget_df.copy())
    budget_df.columns = budget_df.columns.str.strip()

    product_col = None
    product_names = ['Product', 'Product Group', 'PRODUCT NAME']
    
    for col in product_names:
        if col in budget_df.columns:
            product_col = col
            break
    if not product_col:
        product_col = find_column(budget_df, product_names[0], threshold=80)

    if not product_col:
        return None, "Could not find Product Group column in budget dataset."

    budget_cols = {'Qty': [], 'Value': []}
    
    # FIXED: More specific patterns to avoid incorrect range columns
    detailed_pattern = r'^(Qty|Value)\s*[-]\s*(\w{3,})\'?(\d{2,4})'
    range_pattern = r'^(Qty|Value)\s*(\w{3,})\'?(\d{2,4})[-]\s*(\w{3,})\'?(\d{2,4})'
    
    # Invalid patterns to exclude
    invalid_patterns = [
        r'.*april.*dec.*',  # Exclude "april...dec" patterns
        r'.*\w{4,}.*\w{4,}.*',  # Exclude columns with multiple long words
        r'.*\d{2,4}.*\d{2,4}.*[-].*\d{2,4}',  # Exclude malformed range patterns
    ]

    current_app.logger.info(f"Processing budget columns: {list(budget_df.columns)}")

    for col in budget_df.columns:
        col_lower = col.lower()
        
        # Check if column should be excluded
        should_exclude = False
        for invalid_pattern in invalid_patterns:
            if re.search(invalid_pattern, col_lower, re.IGNORECASE):
                current_app.logger.info(f"Excluding invalid budget column: {col}")
                should_exclude = True
                break
        
        if should_exclude:
            continue
        
        # Check for detailed pattern (single month)
        detailed_match = re.match(detailed_pattern, col, re.IGNORECASE)
        if detailed_match:
            qty_or_value, month, year = detailed_match.groups()
            month = month.capitalize()
            year = year[-2:] if len(year) > 2 else year
            month_year = f"{month}-{year}"
            
            current_app.logger.info(f"Found valid detailed budget column: {col} -> {qty_or_value}-{month_year}")
            
            if qty_or_value.lower() == 'qty':
                budget_cols['Qty'].append((col, month_year))
            elif qty_or_value.lower() == 'value':
                budget_cols['Value'].append((col, month_year))
            continue

        # Check for range pattern (ONLY if it's a valid short range)
        range_match = re.match(range_pattern, col, re.IGNORECASE)
        if range_match:
            qty_or_value, start_month, start_year, end_month, end_year = range_match.groups()
            
            # Additional validation for range columns
            start_month = start_month.capitalize()
            end_month = end_month.capitalize()
            
            # Only allow reasonable range columns (3-4 char month names)
            if len(start_month) <= 4 and len(end_month) <= 4:
                start_year = start_year[-2:] if len(start_year) > 2 else start_year
                end_year = end_year[-2:] if len(end_year) > 2 else end_year
                month_year = f"{start_month}{start_year}{end_month.lower()}-{end_year}"
                
                current_app.logger.info(f"Found valid range budget column: {col} -> {qty_or_value}-{month_year}")
                
                if qty_or_value.lower() == 'qty':
                    budget_cols['Qty'].append((col, month_year))
                elif qty_or_value.lower() == 'value':
                    budget_cols['Value'].append((col, month_year))
            else:
                current_app.logger.info(f"Excluding invalid range column (long month names): {col}")

    current_app.logger.info(f"Valid budget columns found - Qty: {len(budget_cols['Qty'])}, Value: {len(budget_cols['Value'])}")

    if not budget_cols['Qty'] and not budget_cols['Value']:
        return None, "No valid budget quantity or value columns found."

    for col, _ in budget_cols['Qty'] + budget_cols['Value']:
        budget_df[col] = pd.to_numeric(budget_df[col], errors='coerce')

    group_cols = [col for col, _ in budget_cols['Qty'] + budget_cols['Value']]
    budget_data = budget_df.groupby([product_col])[group_cols].sum().reset_index()

    rename_dict = {product_col: 'PRODUCT NAME'}
    for col, month_year in budget_cols['Qty']:
        rename_dict[col] = f'Budget-{month_year}_MT'
    for col, month_year in budget_cols['Value']:
        rename_dict[col] = f'Budget-{month_year}_Value'

    budget_data = budget_data.rename(columns=rename_dict)
    budget_data['PRODUCT NAME'] = budget_data['PRODUCT NAME'].str.strip().str.upper()

    current_app.logger.info(f"Final budget data columns: {list(budget_data.columns)}")

    return budget_data, None

def process_last_year_data_enhanced(last_year_file_info, fiscal_info, months):
    """Enhanced last year data processing with better error handling"""
    if not last_year_file_info:
        current_app.logger.info("No last year file provided")
        return {}, {}
    
    actual_ly_mt_data = {}
    actual_ly_value_data = {}
    
    try:
        current_app.logger.info(f"Processing last year file: {last_year_file_info}")
        
        # Read the last year file
        df_last_year = pd.read_excel(
            last_year_file_info['filepath'], 
            sheet_name=last_year_file_info['sheet_name'], 
            header=0
        )
        
        # Handle MultiIndex columns
        if isinstance(df_last_year.columns, pd.MultiIndex):
            df_last_year.columns = ['_'.join(col).strip() for col in df_last_year.columns.values]
        
        df_last_year = handle_duplicate_columns(df_last_year)
        
        current_app.logger.info(f"Last year data shape: {df_last_year.shape}")
        current_app.logger.info(f"Last year columns: {list(df_last_year.columns)}")
        
        # Find required columns with multiple possible names
        product_possible_names = [
            'Type (Make)', 'Type(Make)', 'Product Group' 
            
        ]
        date_possible_names = [
            'Month Format', 'Date', 'Month'
            
        ]
        qty_possible_names = [
            'Actual Quantity', 'Acutal Quantity', 'Qty', 
            'Actual Qty', 'Sales Qty',
        ]
        value_possible_names = [
            'Amount', 'Value', 'Sales Value',
            
        ]
        
        product_col = find_column(df_last_year, product_possible_names, threshold=70, case_sensitive=False)
        date_col = find_column(df_last_year, date_possible_names, threshold=70, case_sensitive=False)
        qty_col = find_column(df_last_year, qty_possible_names, threshold=70, case_sensitive=False)
        value_col = find_column(df_last_year, value_possible_names, threshold=70, case_sensitive=False)
        
        current_app.logger.info(f"Found columns - Product: {product_col}, Date: {date_col}, Qty: {qty_col}, Value: {value_col}")
        
        if not product_col:
            current_app.logger.error("Product column not found in last year data")
            return actual_ly_mt_data, actual_ly_value_data
            
        if not date_col:
            current_app.logger.error("Date column not found in last year data")
            return actual_ly_mt_data, actual_ly_value_data
        
        # Get last year info
        last_start_year = str(fiscal_info['last_fiscal_year_start'])[-2:]
        last_end_year = str(fiscal_info['last_fiscal_year_end'])[-2:]
        
        # Function to extract month from various date formats
        def extract_month_from_date(date_series):
            """Extract month abbreviation from various date formats"""
            month_series = pd.Series(index=date_series.index, dtype=str)
            
            for idx, date_val in date_series.items():
                try:
                    if pd.isna(date_val):
                        continue
                        
                    date_str = str(date_val).strip()
                    
                    # Try different parsing approaches
                    if pd.api.types.is_datetime64_any_dtype(pd.Series([date_val])):
                        # Already datetime
                        month_series.iloc[idx] = pd.to_datetime(date_val).strftime('%b')
                    elif '-' in date_str and len(date_str.split('-')) >= 2:
                        # Format like "2024-04" or "04-2024"
                        parts = date_str.split('-')
                        if len(parts[0]) == 4:  # Year first
                            month_num = int(parts[1])
                        else:  # Month first
                            month_num = int(parts[0])
                        month_series.iloc[idx] = pd.to_datetime(f'2024-{month_num:02d}-01').strftime('%b')
                    elif date_str.isdigit() and 1 <= int(date_str) <= 12:
                        # Just month number
                        month_series.iloc[idx] = pd.to_datetime(f'2024-{int(date_str):02d}-01').strftime('%b')
                    elif len(date_str) >= 3:
                        # Try to parse as month name
                        try:
                            month_series.iloc[idx] = pd.to_datetime(date_str, format='%B').strftime('%b')
                        except:
                            month_series.iloc[idx] = date_str[:3].capitalize()
                    else:
                        current_app.logger.warning(f"Could not parse date: {date_str}")
                        
                except Exception as e:
                    current_app.logger.warning(f"Error parsing date {date_val}: {str(e)}")
                    continue
            
            return month_series
        
        # Process LY Quantity data
        if qty_col:
            current_app.logger.info("Processing LY quantity data")
            df_qty = df_last_year[[product_col, date_col, qty_col]].copy()
            df_qty.columns = ['Product Group', 'Month Format', 'Actual Quantity']
            
            # Clean and convert data
            df_qty['Actual Quantity'] = pd.to_numeric(df_qty['Actual Quantity'], errors='coerce')
            df_qty['Product Group'] = df_qty['Product Group'].replace([pd.NA, None, ''], 'Unknown').fillna('Unknown')
            df_qty['Product Group'] = df_qty['Product Group'].astype(str).str.strip().str.upper()
            
            # Extract month using enhanced function
            df_qty['Month'] = extract_month_from_date(df_qty['Month Format'])
            
            # Clean and filter data
            df_qty = df_qty.dropna(subset=['Actual Quantity', 'Month'])
            df_qty = df_qty[df_qty['Actual Quantity'] != 0]
            df_qty = df_qty[df_qty['Month'] != '']
            
            if not df_qty.empty:
                # Group by product and month
                grouped_qty = df_qty.groupby(['Product Group', 'Month'])['Actual Quantity'].sum().reset_index()
                
                current_app.logger.info(f"Grouped LY quantity data: {len(grouped_qty)} records")
                
                # Convert to LY columns format
                for _, row in grouped_qty.iterrows():
                    product = row['Product Group']
                    month = row['Month']
                    qty_amount = row['Actual Quantity']
                    
                    # Determine year based on month position in fiscal year
                    if month in ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']:
                        year = last_start_year
                    else:  # Jan, Feb, Mar
                        year = last_end_year
                    
                    col_name = f'LY-{month}-{year}'
                    
                    if product not in actual_ly_mt_data:
                        actual_ly_mt_data[product] = {}
                    
                    if col_name in actual_ly_mt_data[product]:
                        actual_ly_mt_data[product][col_name] += qty_amount
                    else:
                        actual_ly_mt_data[product][col_name] = qty_amount
        
        # Process LY Value data (similar logic)
        if value_col:
            current_app.logger.info("Processing LY value data")
            df_value = df_last_year[[product_col, date_col, value_col]].copy()
            df_value.columns = ['Product Group', 'Month Format', 'Value']
            
            # Clean and convert data
            df_value['Value'] = pd.to_numeric(df_value['Value'], errors='coerce')
            df_value['Product Group'] = df_value['Product Group'].replace([pd.NA, None, ''], 'Unknown').fillna('Unknown')
            df_value['Product Group'] = df_value['Product Group'].astype(str).str.strip().str.upper()
            
            # Extract month using enhanced function
            df_value['Month'] = extract_month_from_date(df_value['Month Format'])
            
            # Clean and filter data
            df_value = df_value.dropna(subset=['Value', 'Month'])
            df_value = df_value[df_value['Value'] != 0]
            df_value = df_value[df_value['Month'] != '']
            
            if not df_value.empty:
                # Group by product and month
                grouped_value = df_value.groupby(['Product Group', 'Month'])['Value'].sum().reset_index()
                
                current_app.logger.info(f"Grouped LY value data: {len(grouped_value)} records")
                
                # Convert to LY columns format
                for _, row in grouped_value.iterrows():
                    product = row['Product Group']
                    month = row['Month']
                    value_amount = row['Value']
                    
                    # Determine year based on month position in fiscal year
                    if month in ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']:
                        year = last_start_year
                    else:  # Jan, Feb, Mar
                        year = last_end_year
                    
                    col_name = f'LY-{month}-{year}'
                    
                    if product not in actual_ly_value_data:
                        actual_ly_value_data[product] = {}
                    
                    if col_name in actual_ly_value_data[product]:
                        actual_ly_value_data[product][col_name] += value_amount
                    else:
                        actual_ly_value_data[product][col_name] = value_amount
        
        current_app.logger.info(f"LY MT data extracted for {len(actual_ly_mt_data)} products")
        current_app.logger.info(f"LY Value data extracted for {len(actual_ly_value_data)} products")
        
        return actual_ly_mt_data, actual_ly_value_data
        
    except Exception as e:
        current_app.logger.error(f"Error processing last year data: {str(e)}")
        import traceback
        current_app.logger.error(f"Full traceback: {traceback.format_exc()}")
        return {}, {}

def apply_exact_fiscal_year_column_ordering(df, product_col_name):
    """Apply exact fiscal year column ordering to any dataframe"""
    return merge_processor.reorder_columns_exact_fiscal_year(df, product_col_name)

def build_mt_analysis(budget_data, all_products, actual_mt_data, fiscal_info, months):
    """Build MT (tonnage) analysis with EXACT column ordering and FIXED total calculation"""
    try:
        mt_cols = [col for col in budget_data.columns if col.endswith('_MT') and not col.endswith('LY_MT')]
        if not mt_cols:
            return jsonify({'success': False, 'error': 'No budget tonnage columns found'})

        result_product_mt = pd.DataFrame({'PRODUCT NAME': sorted(list(all_products))})
        
        # Add budget data
        month_cols = sorted(set(col.replace('_MT', '') for col in mt_cols))
        for month_col in month_cols:
            if f'{month_col}_MT' in budget_data.columns:
                result_product_mt[month_col] = 0.0
                matching_data = budget_data[['PRODUCT NAME', f'{month_col}_MT']].dropna()
                matching_data['PRODUCT NAME'] = matching_data['PRODUCT NAME'].astype(str).str.strip().str.upper()
                for _, row in matching_data.iterrows():
                    idx = result_product_mt[result_product_mt['PRODUCT NAME'] == row['PRODUCT NAME']].index
                    if not idx.empty:
                        result_product_mt.loc[idx, month_col] = row[f'{month_col}_MT']

        # Add actual data
        for product, data in actual_mt_data.items():
            idx = result_product_mt[result_product_mt['PRODUCT NAME'] == product].index
            if idx.empty:
                new_row = pd.DataFrame({'PRODUCT NAME': [product]})
                result_product_mt = pd.concat([result_product_mt, new_row], ignore_index=True)
                idx = result_product_mt[result_product_mt['PRODUCT NAME'] == product].index
            for col, value in data.items():
                if col not in result_product_mt.columns:
                    result_product_mt[col] = 0.0
                result_product_mt.loc[idx, col] = value

        # Build EXACT column structure and calculate values
        result_product_mt = build_exact_columns_and_calculate_values(result_product_mt, fiscal_info, 'mt')

        # FIXED: Calculate totals including ALL columns (Budget, LY, Act, Gr, Ach)
        exclude_products = ['NORTH TOTAL', 'WEST SALES', 'GRAND TOTAL', 'TOTAL SALES']
        mask = ~result_product_mt['PRODUCT NAME'].isin(exclude_products)
        valid_products = result_product_mt[mask]
        
        numeric_cols = result_product_mt.select_dtypes(include=[np.number]).columns
        total_row = pd.DataFrame({'PRODUCT NAME': ['TOTAL SALES']})
        
        # FIXED: Include ALL numeric columns including Gr and Ach
        for col in numeric_cols:
            if col in valid_products.columns:
                # Calculate sum for ALL columns - no exclusion for Gr and Ach
                total_row[col] = [valid_products[col].sum(skipna=True).round(2)]
            else:
                total_row[col] = [0.0]

        result_product_mt = pd.concat([valid_products, total_row], ignore_index=True)
        
        # Rename to match specification
        result_product_mt = result_product_mt.rename(columns={'PRODUCT NAME': 'SALES in Tonage'})
        
        # Apply EXACT fiscal year column ordering
        result_product_mt = apply_exact_fiscal_year_column_ordering(result_product_mt, 'SALES in Tonage')
        
        result_dict = result_product_mt.fillna(0).to_dict('records')
        columns = result_product_mt.columns.tolist()
        
        return jsonify({
            'success': True,
            'data': result_dict,
            'columns': columns,
            'shape': [len(result_dict), len(columns)],
            'fiscal_info': fiscal_info,
            'type': 'mt_analysis'
        })
        
    except Exception as e:
        current_app.logger.error(f"Error in MT analysis: {str(e)}")
        return jsonify({'success': False, 'error': f'MT analysis error: {str(e)}'})

def build_value_analysis(budget_data, all_products, actual_value_data, fiscal_info, months):
    """Build Value analysis with EXACT column ordering and FIXED total calculation"""
    try:
        value_cols = [col for col in budget_data.columns if col.endswith('_Value') and not col.endswith('LY_Value')]
        if not value_cols:
            return jsonify({'success': False, 'error': 'No budget value columns found'})

        result_product_value = pd.DataFrame({'PRODUCT NAME': sorted(list(all_products))})
        
        # Add budget data
        month_cols = sorted(set(col.replace('_Value', '') for col in value_cols))
        for month_col in month_cols:
            if f'{month_col}_Value' in budget_data.columns:
                result_product_value[month_col] = 0.0
                matching_data = budget_data[['PRODUCT NAME', f'{month_col}_Value']].dropna()
                matching_data['PRODUCT NAME'] = matching_data['PRODUCT NAME'].astype(str).str.strip().str.upper()
                for _, row in matching_data.iterrows():
                    idx = result_product_value[result_product_value['PRODUCT NAME'] == row['PRODUCT NAME']].index
                    if not idx.empty:
                        result_product_value.loc[idx, month_col] = row[f'{month_col}_Value']

        # Add actual data
        for product, data in actual_value_data.items():
            idx = result_product_value[result_product_value['PRODUCT NAME'] == product].index
            if idx.empty:
                new_row = pd.DataFrame({'PRODUCT NAME': [product]})
                result_product_value = pd.concat([result_product_value, new_row], ignore_index=True)
                idx = result_product_value[result_product_value['PRODUCT NAME'] == product].index
            for col, value in data.items():
                if col not in result_product_value.columns:
                    result_product_value[col] = 0.0
                result_product_value.loc[idx, col] = value

        # Build EXACT column structure and calculate values
        result_product_value = build_exact_columns_and_calculate_values(result_product_value, fiscal_info, 'value')

        # FIXED: Calculate totals including ALL columns (Budget, LY, Act, Gr, Ach)
        exclude_products = ['NORTH TOTAL', 'WEST SALES', 'GRAND TOTAL', 'TOTAL SALES']
        mask = ~result_product_value['PRODUCT NAME'].isin(exclude_products)
        valid_products = result_product_value[mask]
        
        numeric_cols = result_product_value.select_dtypes(include=[np.number]).columns
        total_row = pd.DataFrame({'PRODUCT NAME': ['TOTAL SALES']})
        
        # FIXED: Include ALL numeric columns including Gr and Ach
        for col in numeric_cols:
            if col in valid_products.columns:
                # Calculate sum for ALL columns - no exclusion for Gr and Ach
                total_row[col] = [valid_products[col].sum(skipna=True).round(2)]
            else:
                total_row[col] = [0.0]

        result_product_value = pd.concat([valid_products, total_row], ignore_index=True)
        
        # Rename to match specification  
        result_product_value = result_product_value.rename(columns={'PRODUCT NAME': 'SALES in Value'})
        
        # Apply EXACT fiscal year column ordering
        result_product_value = apply_exact_fiscal_year_column_ordering(result_product_value, 'SALES in Value')
        
        result_dict = result_product_value.fillna(0).to_dict('records')
        columns = result_product_value.columns.tolist()
        
        return jsonify({
            'success': True,
            'data': result_dict,
            'columns': columns,
            'shape': [len(result_dict), len(columns)],
            'fiscal_info': fiscal_info,
            'type': 'value_analysis'
        })
        
    except Exception as e:
        current_app.logger.error(f"Error in Value analysis: {str(e)}")
        return jsonify({'success': False, 'error': f'Value analysis error: {str(e)}'})


def build_merge_analysis(budget_data, all_products, actual_mt_data, actual_value_data, fiscal_info, months):
    """Build merged analysis with both MT and Value data combined"""
    try:
        current_app.logger.info("Starting merge analysis build")
        
        # Build individual analyses
        mt_result = build_mt_analysis(budget_data, all_products, actual_mt_data, fiscal_info, months)
        value_result = build_value_analysis(budget_data, all_products, actual_value_data, fiscal_info, months)
        
        mt_data = None
        value_data = None
        merge_summary = {
            'total_products': len(all_products),
            'mt_products': 0,
            'value_products': 0,
            'fiscal_year': fiscal_info.get('fiscal_year_str', '25-26')
        }
        
        # Process MT result
        if mt_result.status_code == 200:
            mt_json = mt_result.get_json()
            current_app.logger.info(f"MT result success: {mt_json.get('success', False)}")
            if mt_json.get('success'):
                mt_data = {
                    'data': mt_json['data'],
                    'columns': mt_json['columns'],
                    'shape': mt_json.get('shape', [len(mt_json['data']), len(mt_json['columns'])]),
                    'type': 'mt_analysis'
                }
                merge_summary['mt_products'] = len(mt_json['data'])
                current_app.logger.info(f"MT data prepared: {len(mt_json['data'])} rows, {len(mt_json['columns'])} columns")
        else:
            current_app.logger.warning(f"MT analysis failed with status: {mt_result.status_code}")
        
        # Process Value result
        if value_result.status_code == 200:
            value_json = value_result.get_json()
            current_app.logger.info(f"Value result success: {value_json.get('success', False)}")
            if value_json.get('success'):
                value_data = {
                    'data': value_json['data'],
                    'columns': value_json['columns'],
                    'shape': value_json.get('shape', [len(value_json['data']), len(value_json['columns'])]),
                    'type': 'value_analysis'
                }
                merge_summary['value_products'] = len(value_json['data'])
                current_app.logger.info(f"Value data prepared: {len(value_json['data'])} rows, {len(value_json['columns'])} columns")
        else:
            current_app.logger.warning(f"Value analysis failed with status: {value_result.status_code}")
        
        # Check if we have any data
        if not mt_data and not value_data:
            current_app.logger.error("No data available for merge")
            return jsonify({
                'success': False,
                'error': 'No MT or Value data could be generated for merge'
            })
        
        result = {
            'success': True,
            'type': 'merge_analysis',
            'mt_data': mt_data,
            'value_data': value_data,
            'merge_summary': merge_summary,
            'fiscal_info': fiscal_info,
            'message': f'Merge analysis completed. MT: {"✓" if mt_data else "✗"}, Value: {"✓" if value_data else "✗"}'
        }
        
        current_app.logger.info(f"Merge analysis completed successfully: {result['message']}")
        return jsonify(result)
        
    except Exception as e:
        current_app.logger.error(f"Error in merge analysis: {str(e)}")
        import traceback
        current_app.logger.error(f"Full traceback: {traceback.format_exc()}")
        return jsonify({'success': False, 'error': f'Merge analysis error: {str(e)}'})

# MAIN ENDPOINT: Process product analysis with conditional Gr/Ach calculation
@product_bp.route('/process', methods=['POST'])
def process_product_analysis():
    """Process product analysis with budget, sales data, and LAST YEAR data - Only calculate Gr/Ach when Act values present"""
    try:
        data = request.get_json()
        budget_filepath = data.get('budget_filepath')
        budget_sheet = data.get('budget_sheet')
        sales_files = data.get('sales_files', [])
        last_year_file = data.get('last_year_file')
        analysis_type = data.get('analysis_type', 'mt')
        
        current_app.logger.info(f"Processing product analysis - Type: {analysis_type}")
        current_app.logger.info(f"Last year file provided: {last_year_file is not None}")
        
        if not budget_filepath or not budget_sheet:
            return jsonify({'success': False, 'error': 'Budget file and sheet are required'})

        # Process budget data
        budget_df = pd.read_excel(budget_filepath, sheet_name=budget_sheet)
        budget_df.columns = budget_df.columns.str.strip()
        budget_df = budget_df.dropna(how='all').reset_index(drop=True)

        budget_data, error = process_budget_data_product(budget_df, group_type='product')
        if budget_data is None:
            return jsonify({'success': False, 'error': error})

        budget_data['PRODUCT NAME'] = budget_data['PRODUCT NAME'].astype(str).str.strip().str.upper()

        fiscal_info = calculate_fiscal_year()
        months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']

        all_products = set(budget_data['PRODUCT NAME'].dropna().astype(str).str.strip().str.upper())
        actual_mt_data = {}
        actual_value_data = {}

        # Process current year sales files
        for sales_file_info in sales_files:
            try:
                df_sales = pd.read_excel(sales_file_info['filepath'], sheet_name=sales_file_info['sheet_name'], header=0)
                if isinstance(df_sales.columns, pd.MultiIndex):
                    df_sales.columns = ['_'.join(col).strip() for col in df_sales.columns.values]
                df_sales = handle_duplicate_columns(df_sales)
                
                product_col = find_column(df_sales, ['Type (Make)', 'Type(Make)'], case_sensitive=False)
                date_col = find_column(df_sales, ['Month Format', 'Date', 'Month'], case_sensitive=False)
                qty_col = find_column(df_sales, ['Actual Quantity', 'Acutal Quantity'], case_sensitive=False)
                value_col = find_column(df_sales, ['Amount', 'Value', 'Sales Value'], case_sensitive=False)
                
                if product_col and date_col:
                    unique_products = df_sales[product_col].dropna().astype(str).str.strip().str.upper()
                    all_products.update(unique_products)
                    
                    # Process current year quantity data (Act- columns)
                    if qty_col:
                        df_qty = df_sales[[product_col, date_col, qty_col]].copy()
                        df_qty.columns = ['Product Group', 'Month Format', 'Actual Quantity']
                        df_qty['Actual Quantity'] = pd.to_numeric(df_qty['Actual Quantity'], errors='coerce')
                        df_qty['Product Group'] = df_qty['Product Group'].replace([pd.NA, None, ''], 'Unknown').fillna('Unknown')
                        
                        if pd.api.types.is_datetime64_any_dtype(df_qty['Month Format']):
                            df_qty['Month'] = pd.to_datetime(df_qty['Month Format']).dt.strftime('%b')
                        else:
                            month_str = df_qty['Month Format'].astype(str).str.strip().str.title()
                            try:
                                df_qty['Month'] = pd.to_datetime(month_str, format='%B').dt.strftime('%b')
                            except ValueError:
                                df_qty['Month'] = month_str.str[:3]
                        
                        df_qty = df_qty.dropna(subset=['Actual Quantity', 'Month'])
                        df_qty = df_qty[df_qty['Actual Quantity'] != 0]
                        
                        grouped = df_qty.groupby(['Product Group', 'Month'])['Actual Quantity'].sum().reset_index()
                        
                        for _, row in grouped.iterrows():
                            product = row['Product Group']
                            month = row['Month']
                            qty = row['Actual Quantity']
                            year = str(fiscal_info['fiscal_year_start'])[-2:] if month in months[:9] else str(fiscal_info['fiscal_year_end'])[-2:]
                            col_name = f'Act-{month}-{year}'
                            
                            if product not in actual_mt_data:
                                actual_mt_data[product] = {}
                            if col_name in actual_mt_data[product]:
                                actual_mt_data[product][col_name] += qty
                            else:
                                actual_mt_data[product][col_name] = qty
                    
                    # Process current year value data (Act- columns) 
                    if value_col:
                        df_value = df_sales[[product_col, date_col, value_col]].copy()
                        df_value.columns = ['Product Group', 'Month Format', 'Value']
                        df_value['Value'] = pd.to_numeric(df_value['Value'], errors='coerce')
                        df_value['Product Group'] = df_value['Product Group'].replace([pd.NA, None, ''], 'Unknown').fillna('Unknown')
                        
                        if pd.api.types.is_datetime64_any_dtype(df_value['Month Format']):
                            df_value['Month'] = pd.to_datetime(df_value['Month Format']).dt.strftime('%b')
                        else:
                            month_str = df_value['Month Format'].astype(str).str.strip().str.title()
                            try:
                                df_value['Month'] = pd.to_datetime(month_str, format='%B').dt.strftime('%b')
                            except ValueError:
                                df_value['Month'] = month_str.str[:3]
                        
                        df_value = df_value.dropna(subset=['Value', 'Month'])
                        df_value = df_value[df_value['Value'] != 0]
                        
                        grouped = df_value.groupby(['Product Group', 'Month'])['Value'].sum().reset_index()
                        
                        for _, row in grouped.iterrows():
                            product = row['Product Group']
                            month = row['Month']
                            value = row['Value']
                            year = str(fiscal_info['fiscal_year_start'])[-2:] if month in months[:9] else str(fiscal_info['fiscal_year_end'])[-2:]
                            col_name = f'Act-{month}-{year}'
                            
                            if product not in actual_value_data:
                                actual_value_data[product] = {}
                            if col_name in actual_value_data[product]:
                                actual_value_data[product][col_name] += value
                            else:
                                actual_value_data[product][col_name] = value
                                
            except Exception as e:
                current_app.logger.warning(f"Error processing sales file: {str(e)}")
                continue

        # Process last year data to get LY- columns
        ly_mt_data, ly_value_data = process_last_year_data_enhanced(last_year_file, fiscal_info, months)
        
        # Merge LY data into actual data dictionaries
        for product, ly_data in ly_mt_data.items():
            if product not in actual_mt_data:
                actual_mt_data[product] = {}
            actual_mt_data[product].update(ly_data)
            all_products.add(product)
        
        for product, ly_data in ly_value_data.items():
            if product not in actual_value_data:
                actual_value_data[product] = {}
            actual_value_data[product].update(ly_data)
            all_products.add(product)
        
        current_app.logger.info(f"Total products after including LY: {len(all_products)}")
        current_app.logger.info(f"Products with MT data: {len(actual_mt_data)}")
        current_app.logger.info(f"Products with Value data: {len(actual_value_data)}")

        # Generate analysis based on type
        if analysis_type == 'mt':
            return build_mt_analysis(budget_data, all_products, actual_mt_data, fiscal_info, months)
        elif analysis_type == 'value':
            return build_value_analysis(budget_data, all_products, actual_value_data, fiscal_info, months)
        else:  # merge
            return build_merge_analysis(budget_data, all_products, actual_mt_data, actual_value_data, fiscal_info, months)
            
    except Exception as e:
        current_app.logger.error(f"Error in product analysis: {str(e)}")
        import traceback
        current_app.logger.error(f"Full traceback: {traceback.format_exc()}")
        return jsonify({'success': False, 'error': f'Processing error: {str(e)}'})

# Additional endpoint for session storage version
@product_bp.route('/process-with-session', methods=['POST'])
def process_product_analysis_with_session():
    """Process product analysis with session storage for sales monthwise integration"""
    try:
        # Call the main process function
        result = process_product_analysis()
        
        # If successful, store totals in session
        if result.status_code == 200:
            response_data = result.get_json()
            if response_data.get('success'):
                # Store totals based on analysis type
                analysis_type = response_data.get('type', 'unknown')
                
                if analysis_type == 'mt_analysis':
                    store_totals_in_session(response_data, 'mt')
                elif analysis_type == 'value_analysis':
                    store_totals_in_session(response_data, 'value')
                elif analysis_type == 'merge_analysis':
                    if response_data.get('mt_data'):
                        store_totals_in_session(response_data['mt_data'], 'mt')
                    if response_data.get('value_data'):
                        store_totals_in_session(response_data['value_data'], 'value')
                
                # Add session info to response
                response_data['session_info'] = {
                    'totals_stored': True,
                    'available_for_sales_monthwise': True,
                    'session_keys': list(session.get('total_sales_data', {}).keys()),
                    'timestamp': datetime.now().isoformat()
                }
                
                return jsonify(response_data)
        
        return result
        
    except Exception as e:
        current_app.logger.error(f"Error in product analysis with session: {str(e)}")
        return jsonify({'success': False, 'error': f'Processing error: {str(e)}'})


@product_bp.route('/sheets', methods=['POST'])
def get_sheet_info():
    """Get available sheets from uploaded files"""
    try:
        data = request.get_json()
        file_path = data.get('file_path')
        
        if not file_path or not os.path.exists(file_path):
            return jsonify({'success': False, 'error': 'File not found'})
        
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        
        return jsonify({
            'success': True,
            'sheet_names': sheet_names,
            'total_sheets': len(sheet_names)
        })
        
    except Exception as e:
        current_app.logger.error(f"Error getting sheet info: {str(e)}")
        return jsonify({'success': False, 'error': f'Error reading file: {str(e)}'})
    
# Add this new endpoint to your existing product.py file

@product_bp.route('/export/combined-single-sheet', methods=['POST'])
def export_combined_single_sheet():
    """Generate and download Excel file with both MT and Value data in single sheet with proper column ordering"""
    try:
        data = request.json
        mt_data = data.get('mt_data', [])
        value_data = data.get('value_data', [])
        mt_columns = data.get('mt_columns', [])
        value_columns = data.get('value_columns', [])
        fiscal_info = data.get('fiscal_info', {})
        include_both_tables = data.get('include_both_tables', True)
        
        fiscal_year = fiscal_info.get('fiscal_year_str', '25-26')
        current_start_year = str(fiscal_info.get('fiscal_year_start', 2025))[-2:]
        current_end_year = str(fiscal_info.get('fiscal_year_end', 2026))[-2:]
        last_start_year = str(fiscal_info.get('last_fiscal_year_start', 2024))[-2:]
        last_end_year = str(fiscal_info.get('last_fiscal_year_end', 2025))[-2:]
        
        months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
        
        def get_exact_column_order(first_col_name, columns):
            """Get exact column order based on fiscal year pattern"""
            if not columns or len(columns) <= 1:
                return columns
                
            # Separate first column and data columns
            first_col = columns[0] if columns[0] == first_col_name else first_col_name
            data_columns = [col for col in columns if col != first_col_name]
            
            ordered_columns = [first_col]
            
            # Monthly columns in exact order
            for month in months:
                # Determine year based on fiscal year
                if month in ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']:
                    year = current_start_year
                    ly_year = last_start_year
                else:  # Jan, Feb, Mar
                    year = current_end_year
                    ly_year = last_end_year
                
                # Order: Budget, LY, Act, Gr, Ach
                month_columns = [
                    f'Budget-{month}-{year}',
                    f'LY-{month}-{ly_year}',
                    f'Act-{month}-{year}',
                    f'Gr-{month}-{year}',
                    f'Ach-{month}-{year}'
                ]
                
                # Add existing columns in this order
                for col in month_columns:
                    if col in data_columns:
                        ordered_columns.append(col)
                
                # Add YTD columns after specific months
                if month == 'Jun':  # Q1 YTD
                    ytd_cols = [
                        f'YTD-{fiscal_year} (Apr to Jun)Budget',
                        f'YTD-{last_start_year}-{last_end_year} (Apr to Jun)LY',
                        f'Act-YTD-{fiscal_year} (Apr to Jun)',
                        f'Gr-YTD-{fiscal_year} (Apr to Jun)',
                        f'Ach-YTD-{fiscal_year} (Apr to Jun)'
                    ]
                    for col in ytd_cols:
                        if col in data_columns:
                            ordered_columns.append(col)
                
                elif month == 'Sep':  # H1 YTD
                    ytd_cols = [
                        f'YTD-{fiscal_year} (Apr to Sep)Budget',
                        f'YTD-{last_start_year}-{last_end_year} (Apr to Sep)LY',
                        f'Act-YTD-{fiscal_year} (Apr to Sep)',
                        f'Gr-YTD-{fiscal_year} (Apr to Sep)',
                        f'Ach-YTD-{fiscal_year} (Apr to Sep)'
                    ]
                    for col in ytd_cols:
                        if col in data_columns:
                            ordered_columns.append(col)
                
                elif month == 'Dec':  # 9M YTD
                    ytd_cols = [
                        f'YTD-{fiscal_year} (Apr to Dec)Budget',
                        f'YTD-{last_start_year}-{last_end_year} (Apr to Dec)LY',
                        f'Act-YTD-{fiscal_year} (Apr to Dec)',
                        f'Gr-YTD-{fiscal_year} (Apr to Dec)',
                        f'Ach-YTD-{fiscal_year} (Apr to Dec)'
                    ]
                    for col in ytd_cols:
                        if col in data_columns:
                            ordered_columns.append(col)
                
                elif month == 'Mar':  # Full Year YTD
                    ytd_cols = [
                        f'YTD-{fiscal_year} (Apr to Mar) Budget',
                        f'YTD-{last_start_year}-{last_end_year} (Apr to Mar)LY',
                        f'Act-YTD-{fiscal_year} (Apr to Mar)',
                        f'Gr-YTD-{fiscal_year} (Apr to Mar)',
                        f'Ach-YTD-{fiscal_year} (Apr to Mar)'
                    ]
                    for col in ytd_cols:
                        if col in data_columns:
                            ordered_columns.append(col)
            
            # Add any remaining columns that weren't matched
            for col in data_columns:
                if col not in ordered_columns:
                    ordered_columns.append(col)
            
            return ordered_columns
        
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define formats
            title_format = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 18, 'font_color': '#000000', 'bg_color': '#B4C6E7',
                'border': 2, 'border_color': '#4472C4'
            })
            
            subtitle_format = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 14, 'font_color': '#1F4E79', 'bg_color': '#D9E1F2',
                'border': 1
            })
            
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center',
                'fg_color': '#4472C4', 'font_color': 'white', 'border': 1,
                'font_size': 10
            })
            
            product_header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center',
                'fg_color': '#2E5F8A', 'font_color': 'white', 'border': 1,
                'font_size': 10
            })
            
            num_format = workbook.add_format({
                'num_format': '#,##0.00', 'border': 1, 'align': 'right'
            })
            
            text_format = workbook.add_format({
                'border': 1, 'valign': 'vcenter', 'align': 'left'
            })
            
            total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#E2EFDA', 
                'font_color': '#155724', 'border': 1, 'font_size': 11
            })
            
            # Create worksheet
            worksheet_name = 'Combined Product Analysis'
            worksheet = workbook.add_worksheet(worksheet_name)
            
            current_row = 0
            
            # Main title
            main_title = f"Product-wise Sales Analysis - Combined Report (FY {fiscal_year})"
            max_cols = max(len(mt_columns) if mt_columns else 0, len(value_columns) if value_columns else 0)
            if max_cols > 1:
                worksheet.merge_range(current_row, 0, current_row, max_cols - 1, main_title, title_format)
            else:
                worksheet.write(current_row, 0, main_title, title_format)
            current_row += 2
            
            def get_total_format_product(product_name):
                """Get appropriate format based on product name"""
                product_upper = str(product_name).upper()
                if 'TOTAL SALES' in product_upper or 'GRAND TOTAL' in product_upper:
                    return total_format
                else:
                    return None
            
            # Write MT Data Table with proper column ordering
            if mt_data and include_both_tables:
                # Reorder MT columns
                ordered_mt_columns = get_exact_column_order('SALES in Tonage', mt_columns)
                
                # MT Table Title
                mt_title = f"Product-wise SALES in Tonnage Analysis - FY {fiscal_year}"
                if len(ordered_mt_columns) > 1:
                    worksheet.merge_range(current_row, 0, current_row, len(ordered_mt_columns) - 1, mt_title, subtitle_format)
                else:
                    worksheet.write(current_row, 0, mt_title, subtitle_format)
                current_row += 2
                
                # MT Headers
                for col_num, column_name in enumerate(ordered_mt_columns):
                    if col_num == 0:  # First column (product names)
                        worksheet.write(current_row, col_num, column_name, product_header_format)
                        worksheet.set_column(col_num, col_num, 30)  # Wider for product names
                    else:
                        worksheet.write(current_row, col_num, column_name, header_format)
                        worksheet.set_column(col_num, col_num, 12)  # Standard width for data
                current_row += 1
                
                # MT Data
                for row_data in mt_data:
                    for col_num, column_name in enumerate(ordered_mt_columns):
                        cell_value = row_data.get(column_name, 0)
                        
                        if col_num == 0:  # Product column
                            product_name = str(cell_value)
                            worksheet.write(current_row, col_num, product_name, text_format)
                        else:  # Data columns
                            # Check if this is a total row
                            product_name = row_data.get(ordered_mt_columns[0], '')
                            total_fmt = get_total_format_product(product_name)
                            
                            if total_fmt:
                                worksheet.write(current_row, col_num, cell_value, total_fmt)
                            else:
                                worksheet.write(current_row, col_num, cell_value, num_format)
                    
                    current_row += 1
                
                # Add spacing between tables
                current_row += 3
            
            # Write Value Data Table with proper column ordering
            if value_data and include_both_tables:
                # Reorder Value columns
                ordered_value_columns = get_exact_column_order('SALES in Value', value_columns)
                
                # Value Table Title
                value_title = f"Product-wise SALES in Value Analysis - FY {fiscal_year}"
                if len(ordered_value_columns) > 1:
                    worksheet.merge_range(current_row, 0, current_row, len(ordered_value_columns) - 1, value_title, subtitle_format)
                else:
                    worksheet.write(current_row, 0, value_title, subtitle_format)
                current_row += 2
                
                # Value Headers
                for col_num, column_name in enumerate(ordered_value_columns):
                    if col_num == 0:  # First column (product names)
                        worksheet.write(current_row, col_num, column_name, product_header_format)
                        worksheet.set_column(col_num, col_num, 30)  # Wider for product names
                    else:
                        worksheet.write(current_row, col_num, column_name, header_format)
                        worksheet.set_column(col_num, col_num, 12)  # Standard width for data
                current_row += 1
                
                # Value Data
                for row_data in value_data:
                    for col_num, column_name in enumerate(ordered_value_columns):
                        cell_value = row_data.get(column_name, 0)
                        
                        if col_num == 0:  # Product column
                            product_name = str(cell_value)
                            worksheet.write(current_row, col_num, product_name, text_format)
                        else:  # Data columns
                            # Check if this is a total row
                            product_name = row_data.get(ordered_value_columns[0], '')
                            total_fmt = get_total_format_product(product_name)
                            
                            if total_fmt:
                                worksheet.write(current_row, col_num, cell_value, total_fmt)
                            else:
                                worksheet.write(current_row, col_num, cell_value, num_format)
                    
                    current_row += 1
            
            # Handle single table cases
            elif mt_data and not include_both_tables:
                # Only MT table with proper ordering
                ordered_mt_columns = get_exact_column_order('SALES in Tonage', mt_columns)
                mt_title = f"Product-wise SALES in Tonnage Analysis - FY {fiscal_year}"
                if len(ordered_mt_columns) > 1:
                    worksheet.merge_range(current_row, 0, current_row, len(ordered_mt_columns) - 1, mt_title, subtitle_format)
                current_row += 2
                
                # Write MT data with proper ordering
                for col_num, column_name in enumerate(ordered_mt_columns):
                    fmt = product_header_format if col_num == 0 else header_format
                    worksheet.write(current_row, col_num, column_name, fmt)
                    worksheet.set_column(col_num, col_num, 30 if col_num == 0 else 12)
                current_row += 1
                
                for row_data in mt_data:
                    for col_num, column_name in enumerate(ordered_mt_columns):
                        cell_value = row_data.get(column_name, 0)
                        if col_num == 0:
                            worksheet.write(current_row, col_num, str(cell_value), text_format)
                        else:
                            product_name = row_data.get(ordered_mt_columns[0], '')
                            total_fmt = get_total_format_product(product_name)
                            fmt = total_fmt if total_fmt else num_format
                            worksheet.write(current_row, col_num, cell_value, fmt)
                    current_row += 1
            
            # Add footer with generation info
            current_row += 3
            footer_info = [
                f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"Fiscal Year: {fiscal_year}",
                f"MT Records: {len(mt_data)}",
                f"Value Records: {len(value_data)}",
                f"Total Records: {len(mt_data) + len(value_data)}",
                f"Column Ordering: Fiscal Year Pattern (Budget → LY → Act → Gr → Ach)"
            ]
            
            footer_format = workbook.add_format({
                'font_size': 9, 'italic': True, 'font_color': '#666666'
            })
            
            for info_line in footer_info:
                worksheet.write(current_row, 0, info_line, footer_format)
                current_row += 1
            
            # Freeze panes at the first data row
            worksheet.freeze_panes(4, 1)  # Freeze after titles and headers
            
            # Set print settings
            worksheet.set_landscape()
            worksheet.set_paper(9)  # A4
            worksheet.fit_to_pages(1, 0)  # Fit to 1 page wide, unlimited pages tall
        
        excel_data = output.getvalue()
        
        return send_file(
            BytesIO(excel_data),
            as_attachment=True,
            download_name=f"product_combined_single_sheet_{fiscal_year}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        current_app.logger.error(f"Error in product single sheet generation: {str(e)}")
        import traceback
        current_app.logger.error(f"Full traceback: {traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': f'Product single sheet Excel generation error: {str(e)}'
        }), 500


# Optional: Add an endpoint to get single sheet preview/metadata for products
@product_bp.route('/get-single-sheet-preview', methods=['POST'])
def get_product_single_sheet_preview():
    """Get preview information for product single sheet generation"""
    try:
        data = request.json
        mt_data = data.get('mt_data', [])
        value_data = data.get('value_data', [])
        mt_columns = data.get('mt_columns', [])
        value_columns = data.get('value_columns', [])
        fiscal_year = data.get('fiscal_year', '')
        
        preview_info = {
            'sheet_name': 'Combined Product Analysis',
            'layout': {
                'main_title': f"Product-wise Sales Analysis - Combined Report (FY {fiscal_year})",
                'tables': []
            },
            'estimated_rows': 0,
            'estimated_columns': 0,
            'column_ordering': {
                'pattern': 'Budget → LY → Act → Gr → Ach',
                'fiscal_year_months': ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar'],
                'ytd_periods': ['Q1 (Apr-Jun)', 'H1 (Apr-Sep)', '9M (Apr-Dec)', 'FY (Apr-Mar)']
            }
        }
        
        current_row = 3  # Start after main title and spacing
        
        if mt_data:
            mt_table_info = {
                'title': f"Product-wise SALES in Tonnage Analysis - FY {fiscal_year}",
                'start_row': current_row,
                'end_row': current_row + 2 + len(mt_data),
                'columns': len(mt_columns),
                'rows': len(mt_data),
                'type': 'MT',
                'first_column': 'SALES in Tonage'
            }
            preview_info['layout']['tables'].append(mt_table_info)
            current_row = mt_table_info['end_row'] + 3
        
        if value_data:
            value_table_info = {
                'title': f"Product-wise SALES in Value Analysis - FY {fiscal_year}",
                'start_row': current_row,
                'end_row': current_row + 2 + len(value_data),
                'columns': len(value_columns),
                'rows': len(value_data),
                'type': 'Value',
                'first_column': 'SALES in Value'
            }
            preview_info['layout']['tables'].append(value_table_info)
            current_row = value_table_info['end_row']
        
        preview_info['estimated_rows'] = current_row + 10  # Add footer space
        preview_info['estimated_columns'] = max(len(mt_columns) if mt_columns else 0, 
                                               len(value_columns) if value_columns else 0)
        
        # Product classifications info
        preview_info['product_classifications'] = {
            'individual_products': 'Sorted alphabetically',
            'totals': ['TOTAL SALES'],
            'formatting': {
                'total_rows': 'Green background with bold formatting',
                'data_columns': 'Right-aligned numbers with comma separators',
                'product_column': 'Left-aligned text, wider column'
            }
        }
        
        return jsonify({
            'success': True,
            'preview': preview_info,
            'can_generate': len(mt_data) > 0 or len(value_data) > 0
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'Preview generation error: {str(e)}'
        }), 500


# Helper function to validate column ordering
def validate_fiscal_year_ordering(columns, fiscal_year):
    """Validate that columns follow proper fiscal year ordering"""
    if not columns or len(columns) <= 1:
        return True, "No data columns to validate"
    
    months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
    fiscal_years = fiscal_year.split('-') if '-' in fiscal_year else ['25', '26']
    
    expected_pattern = ['Budget', 'LY', 'Act', 'Gr', 'Ach']
    issues = []
    
    current_month_index = 0
    for i, col in enumerate(columns[1:], 1):  # Skip first column (product names)
        # Check if column follows expected pattern
        found_pattern = False
        for pattern in expected_pattern:
            if pattern in col:
                found_pattern = True
                break
        
        if not found_pattern:
            issues.append(f"Column {i}: '{col}' doesn't match expected pattern")
    
    is_valid = len(issues) == 0
    message = "Column ordering is valid" if is_valid else f"Issues found: {'; '.join(issues)}"
    
    return is_valid, message        

# Add these endpoints to your existing product.py file

@product_bp.route('/download-combined-single-sheet', methods=['POST'])
def download_combined_single_sheet():
    """Download combined single sheet Excel file directly"""
    try:
        data = request.json
        mt_data = data.get('mt_data', [])
        value_data = data.get('value_data', [])
        mt_columns = data.get('mt_columns', [])
        value_columns = data.get('value_columns', [])
        fiscal_year = data.get('fiscal_year', '')
        include_both_tables = data.get('include_both_tables', True)
        
        current_app.logger.info(f"=== DEBUG: Product Download Single Sheet ===")
        current_app.logger.info(f"MT data records: {len(mt_data)}")
        current_app.logger.info(f"Value data records: {len(value_data)}")
        current_app.logger.info(f"Include both tables: {include_both_tables}")
        
        if not mt_data and not value_data:
            return jsonify({
                'success': False,
                'error': 'No data provided for download'
            }), 400
        
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define enhanced formats
            title_format = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 18, 'font_color': '#000000', 'bg_color': '#B4C6E7',
                'border': 2, 'border_color': '#4472C4'
            })
            
            subtitle_format = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 14, 'font_color': '#1F4E79', 'bg_color': '#D9E1F2',
                'border': 1
            })
            
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center',
                'fg_color': '#4472C4', 'font_color': 'white', 'border': 1,
                'font_size': 10
            })
            
            product_header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center',
                'fg_color': '#2E5F8A', 'font_color': 'white', 'border': 1,
                'font_size': 10
            })
            
            num_format = workbook.add_format({
                'num_format': '#,##0.00', 'border': 1, 'align': 'right'
            })
            
            text_format = workbook.add_format({
                'border': 1, 'valign': 'vcenter', 'align': 'left'
            })
            
            total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#E2EFDA', 
                'font_color': '#155724', 'border': 1, 'font_size': 11
            })
            
            # Create single worksheet
            worksheet_name = 'Combined Product Analysis'
            worksheet = workbook.add_worksheet(worksheet_name)
            
            current_row = 0
            
            # Main title
            main_title = f"Product-wise Sales Analysis - Combined Report (FY {fiscal_year})"
            max_cols = max(len(mt_columns) if mt_columns else 0, len(value_columns) if value_columns else 0)
            if max_cols > 1:
                worksheet.merge_range(current_row, 0, current_row, max_cols - 1, main_title, title_format)
            else:
                worksheet.write(current_row, 0, main_title, title_format)
            current_row += 2
            
            def get_total_format_product(product_name):
                """Get appropriate format based on product name"""
                product_upper = str(product_name).upper()
                if 'TOTAL SALES' in product_upper or 'GRAND TOTAL' in product_upper:
                    return total_format
                else:
                    return None
            
            # Function to reorder columns according to fiscal year pattern
            def reorder_product_columns(columns, first_col_name):
                """Reorder columns to match fiscal year pattern: Budget, LY, Act, Gr, Ach"""
                if not columns or len(columns) <= 1:
                    return columns
                
                # Separate first column and data columns
                first_col = columns[0] if columns[0] == first_col_name else first_col_name
                data_columns = [col for col in columns if col != first_col_name]
                
                # Define month order and fiscal year pattern
                months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
                fiscal_years = [fiscal_year.split('-')[0], fiscal_year.split('-')[1]] if '-' in fiscal_year else ['25', '26']
                
                ordered_columns = [first_col]
                
                # Process individual months first
                for month in months:
                    # Determine year based on fiscal year
                    if month in ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']:
                        year = fiscal_years[0]
                        ly_year = str(int(fiscal_years[0]) - 1).zfill(2)
                    else:  # Jan, Feb, Mar
                        year = fiscal_years[1]
                        ly_year = str(int(fiscal_years[1]) - 1).zfill(2)
                    
                    # Order: Budget, LY, Act, Gr, Ach
                    month_columns = [
                        f'Budget-{month}-{year}',
                        f'LY-{month}-{ly_year}',
                        f'Act-{month}-{year}',
                        f'Gr-{month}-{year}',
                        f'Ach-{month}-{year}'
                    ]
                    
                    # Add existing columns in this order
                    for col in month_columns:
                        if col in data_columns:
                            ordered_columns.append(col)
                    
                    # Add YTD columns after specific months
                    if month == 'Jun':  # Q1 YTD
                        ytd_cols = [
                            f'YTD-{fiscal_year} (Apr to Jun)Budget',
                            f'YTD-{ly_year}-{str(int(ly_year)+1).zfill(2)} (Apr to Jun)LY',
                            f'Act-YTD-{fiscal_year} (Apr to Jun)',
                            f'Gr-YTD-{fiscal_year} (Apr to Jun)',
                            f'Ach-YTD-{fiscal_year} (Apr to Jun)'
                        ]
                        for col in ytd_cols:
                            if col in data_columns:
                                ordered_columns.append(col)
                    
                    elif month == 'Sep':  # H1 YTD
                        ytd_cols = [
                            f'YTD-{fiscal_year} (Apr to Sep)Budget',
                            f'YTD-{ly_year}-{str(int(ly_year)+1).zfill(2)} (Apr to Sep)LY',
                            f'Act-YTD-{fiscal_year} (Apr to Sep)',
                            f'Gr-YTD-{fiscal_year} (Apr to Sep)',
                            f'Ach-YTD-{fiscal_year} (Apr to Sep)'
                        ]
                        for col in ytd_cols:
                            if col in data_columns:
                                ordered_columns.append(col)
                    
                    elif month == 'Dec':  # 9M YTD
                        ytd_cols = [
                            f'YTD-{fiscal_year} (Apr to Dec)Budget',
                            f'YTD-{ly_year}-{str(int(ly_year)+1).zfill(2)} (Apr to Dec)LY',
                            f'Act-YTD-{fiscal_year} (Apr to Dec)',
                            f'Gr-YTD-{fiscal_year} (Apr to Dec)',
                            f'Ach-YTD-{fiscal_year} (Apr to Dec)'
                        ]
                        for col in ytd_cols:
                            if col in data_columns:
                                ordered_columns.append(col)
                    
                    elif month == 'Mar':  # Full Year YTD
                        ytd_cols = [
                            f'YTD-{fiscal_year} (Apr to Mar) Budget',  # Note the space
                            f'YTD-{ly_year}-{str(int(ly_year)+1).zfill(2)} (Apr to Mar)LY',
                            f'Act-YTD-{fiscal_year} (Apr to Mar)',
                            f'Gr-YTD-{fiscal_year} (Apr to Mar)',
                            f'Ach-YTD-{fiscal_year} (Apr to Mar)'
                        ]
                        for col in ytd_cols:
                            if col in data_columns:
                                ordered_columns.append(col)
                
                # Add any remaining columns that weren't matched
                for col in data_columns:
                    if col not in ordered_columns:
                        ordered_columns.append(col)
                
                current_app.logger.info(f"Reordered columns from {len(columns)} to {len(ordered_columns)}")
                return ordered_columns
            
            # Write MT Data Table with proper column ordering
            if mt_data and include_both_tables:
                # Reorder MT columns
                ordered_mt_columns = reorder_product_columns(mt_columns, 'SALES in Tonage')
                
                # MT Table Title
                mt_title = f"Product-wise SALES in Tonnage Analysis - FY {fiscal_year}"
                if len(ordered_mt_columns) > 1:
                    worksheet.merge_range(current_row, 0, current_row, len(ordered_mt_columns) - 1, mt_title, subtitle_format)
                else:
                    worksheet.write(current_row, 0, mt_title, subtitle_format)
                current_row += 2
                
                # MT Headers
                for col_num, column_name in enumerate(ordered_mt_columns):
                    if col_num == 0:  # First column (product names)
                        worksheet.write(current_row, col_num, column_name, product_header_format)
                        worksheet.set_column(col_num, col_num, 30)  # Wider for product names
                    else:
                        worksheet.write(current_row, col_num, column_name, header_format)
                        worksheet.set_column(col_num, col_num, 12)  # Standard width for data
                current_row += 1
                
                # MT Data
                for row_num, row_data in enumerate(mt_data):
                    for col_num, column_name in enumerate(ordered_mt_columns):
                        cell_value = row_data.get(column_name, 0)
                        
                        if col_num == 0:  # Product column
                            product_name = str(cell_value)
                            worksheet.write(current_row, col_num, product_name, text_format)
                        else:  # Data columns
                            # Check if this is a total row
                            product_name = row_data.get(ordered_mt_columns[0], '')
                            total_fmt = get_total_format_product(product_name)
                            
                            if total_fmt:
                                worksheet.write(current_row, col_num, cell_value, total_fmt)
                            else:
                                worksheet.write(current_row, col_num, cell_value, num_format)
                    
                    current_row += 1
                
                # Add spacing between tables
                current_row += 3
            
            # Write Value Data Table with proper column ordering
            if value_data and include_both_tables:
                # Reorder Value columns
                ordered_value_columns = reorder_product_columns(value_columns, 'SALES in Value')
                
                # Value Table Title
                value_title = f"Product-wise SALES in Value Analysis - FY {fiscal_year}"
                if len(ordered_value_columns) > 1:
                    worksheet.merge_range(current_row, 0, current_row, len(ordered_value_columns) - 1, value_title, subtitle_format)
                else:
                    worksheet.write(current_row, 0, value_title, subtitle_format)
                current_row += 2
                
                # Value Headers
                for col_num, column_name in enumerate(ordered_value_columns):
                    if col_num == 0:  # First column (product names)
                        worksheet.write(current_row, col_num, column_name, product_header_format)
                        worksheet.set_column(col_num, col_num, 30)  # Wider for product names
                    else:
                        worksheet.write(current_row, col_num, column_name, header_format)
                        worksheet.set_column(col_num, col_num, 12)  # Standard width for data
                current_row += 1
                
                # Value Data
                for row_num, row_data in enumerate(value_data):
                    for col_num, column_name in enumerate(ordered_value_columns):
                        cell_value = row_data.get(column_name, 0)
                        
                        if col_num == 0:  # Product column
                            product_name = str(cell_value)
                            worksheet.write(current_row, col_num, product_name, text_format)
                        else:  # Data columns
                            # Check if this is a total row
                            product_name = row_data.get(ordered_value_columns[0], '')
                            total_fmt = get_total_format_product(product_name)
                            
                            if total_fmt:
                                worksheet.write(current_row, col_num, cell_value, total_fmt)
                            else:
                                worksheet.write(current_row, col_num, cell_value, num_format)
                    
                    current_row += 1
            
            # Handle single table cases
            elif mt_data and not include_both_tables:
                # Only MT table with proper ordering
                ordered_mt_columns = reorder_product_columns(mt_columns, 'SALES in Tonage')
                mt_title = f"Product-wise SALES in Tonnage Analysis - FY {fiscal_year}"
                if len(ordered_mt_columns) > 1:
                    worksheet.merge_range(current_row, 0, current_row, len(ordered_mt_columns) - 1, mt_title, subtitle_format)
                current_row += 2
                
                # Write MT data with proper ordering
                for col_num, column_name in enumerate(ordered_mt_columns):
                    fmt = product_header_format if col_num == 0 else header_format
                    worksheet.write(current_row, col_num, column_name, fmt)
                    worksheet.set_column(col_num, col_num, 30 if col_num == 0 else 12)
                current_row += 1
                
                for row_data in mt_data:
                    for col_num, column_name in enumerate(ordered_mt_columns):
                        cell_value = row_data.get(column_name, 0)
                        if col_num == 0:
                            worksheet.write(current_row, col_num, str(cell_value), text_format)
                        else:
                            product_name = row_data.get(ordered_mt_columns[0], '')
                            total_fmt = get_total_format_product(product_name)
                            fmt = total_fmt if total_fmt else num_format
                            worksheet.write(current_row, col_num, cell_value, fmt)
                    current_row += 1
            
            # Add footer with generation info
            current_row += 3
            footer_info = [
                f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"Fiscal Year: {fiscal_year}",
                f"MT Records: {len(mt_data)}",
                f"Value Records: {len(value_data)}",
                f"Total Records: {len(mt_data) + len(value_data)}",
                f"Column Ordering: Fiscal Year Pattern (Budget → LY → Act → Gr → Ach)"
            ]
            
            footer_format = workbook.add_format({
                'font_size': 9, 'italic': True, 'font_color': '#666666'
            })
            
            for info_line in footer_info:
                worksheet.write(current_row, 0, info_line, footer_format)
                current_row += 1
            
            # Freeze panes at the first data row
            worksheet.freeze_panes(4, 1)  # Freeze after titles and headers
            
            # Set print settings
            worksheet.set_landscape()
            worksheet.set_paper(9)  # A4
            worksheet.fit_to_pages(1, 0)  # Fit to 1 page wide, unlimited pages tall
        
        excel_data = output.getvalue()
        
        current_app.logger.info(f"=== DEBUG: Product download file generated successfully ===")
        current_app.logger.info(f"File size: {len(excel_data)} bytes")
        
        # Return as binary data for download
        from flask import send_file
        return send_file(
            BytesIO(excel_data),
            as_attachment=True,
            download_name=f"product_combined_download_{fiscal_year}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        current_app.logger.error(f"=== DEBUG: Error in product download generation ===")
        current_app.logger.error(f"Error: {str(e)}")
        import traceback
        current_app.logger.error(f"Full traceback: {traceback.format_exc()}")
        
        return jsonify({
            'success': False,
            'error': f'Product download Excel generation error: {str(e)}'
        }), 500




# Export the blueprint
__all__ = ['product_bp', 'store_totals_in_session']