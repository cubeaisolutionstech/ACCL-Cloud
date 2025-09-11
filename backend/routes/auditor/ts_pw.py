from flask import Blueprint, request, jsonify, current_app
import pandas as pd
import numpy as np
import re
from datetime import datetime
from io import BytesIO
import os
from werkzeug.utils import secure_filename
import base64

ts_pw_bp = Blueprint('ts_pw', __name__, url_prefix='/api/ts-pw')

class TSPWMergePreviewProcessor:
    """Handles the complete TS-PW merge preview logic with EXACT column ordering"""
    
    def __init__(self):
        self.fiscal_info = self.calculate_fiscal_year()
        self.months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
        
        # Table headers for detection
        self.mt_table_headers = [
            "SALES in Tonage", "SALES IN TONAGE", "Tonage", "TONAGE",
            "Sales in MT", "SALES IN MT", "SALES in Ton", "Metric Tons", 
            "MT Sales", "Tonage Sales", "Sales Tonage", "SALES in Tonnage"
        ]
        
        self.value_table_headers = [
            "SALES in Value", "SALES IN VALUE", "Sales in Rs", "SALES IN RS",
            "Value", "VALUE", "Sales Value", "SALES VALUE"
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
        
        # EXACT ORDER as specified in requirements
        exact_order = [
            # 1-5. April
            f'Budget-Apr-{current_start_year}',
            f'LY-Apr-{last_start_year}', 
            f'Act-Apr-{current_start_year}', 
            f'Gr-Apr-{current_start_year}', 
            f'Ach-Apr-{current_start_year}',
            
            # 6-10. May  
            f'Budget-May-{current_start_year}', 
            f'LY-May-{last_start_year}', 
            f'Act-May-{current_start_year}', 
            f'Gr-May-{current_start_year}', 
            f'Ach-May-{current_start_year}',
            
            # 11-15. June
            f'Budget-Jun-{current_start_year}', 
            f'LY-Jun-{last_start_year}', 
            f'Act-Jun-{current_start_year}', 
            f'Gr-Jun-{current_start_year}', 
            f'Ach-Jun-{current_start_year}',
            
            # 16-20. Q1 YTD (Apr to Jun)
            f'YTD-{current_fy} (Apr to Jun)Budget', 
            f'YTD-{last_fy} (Apr to Jun)LY', 
            f'Act-YTD-{current_fy} (Apr to Jun)', 
            f'Gr-YTD-{current_fy} (Apr to Jun)', 
            f'Ach-YTD-{current_fy} (Apr to Jun)',
            
            # 21-25. July
            f'Budget-Jul-{current_start_year}', 
            f'LY-Jul-{last_start_year}', 
            f'Act-Jul-{current_start_year}', 
            f'Gr-Jul-{current_start_year}', 
            f'Ach-Jul-{current_start_year}',
            
            # 26-30. August
            f'Budget-Aug-{current_start_year}', 
            f'LY-Aug-{last_start_year}', 
            f'Act-Aug-{current_start_year}', 
            f'Gr-Aug-{current_start_year}', 
            f'Ach-Aug-{current_start_year}',
            
            # 31-35. September
            f'Budget-Sep-{current_start_year}', 
            f'LY-Sep-{last_start_year}', 
            f'Act-Sep-{current_start_year}', 
            f'Gr-Sep-{current_start_year}', 
            f'Ach-Sep-{current_start_year}',
            
            # 36-40. H1 YTD (Apr to Sep)
            f'YTD-{current_fy} (Apr to Sep)Budget', 
            f'YTD-{last_fy} (Apr to Sep)LY', 
            f'Act-YTD-{current_fy} (Apr to Sep)', 
            f'Gr-YTD-{current_fy} (Apr to Sep)', 
            f'Ach-YTD-{current_fy} (Apr to Sep)',
            
            # 41-45. October
            f'Budget-Oct-{current_start_year}', 
            f'LY-Oct-{last_start_year}', 
            f'Act-Oct-{current_start_year}', 
            f'Gr-Oct-{current_start_year}', 
            f'Ach-Oct-{current_start_year}',
            
            # 46-50. November
            f'Budget-Nov-{current_start_year}', 
            f'LY-Nov-{last_start_year}', 
            f'Act-Nov-{current_start_year}', 
            f'Gr-Nov-{current_start_year}', 
            f'Ach-Nov-{current_start_year}',
            
            # 51-55. December
            f'Budget-Dec-{current_start_year}', 
            f'LY-Dec-{last_start_year}', 
            f'Act-Dec-{current_start_year}', 
            f'Gr-Dec-{current_start_year}', 
            f'Ach-Dec-{current_start_year}',
            
            # 56-60. 9M YTD (Apr to Dec)
            f'YTD-{current_fy} (Apr to Dec)Budget', 
            f'YTD-{last_fy} (Apr to Dec)LY', 
            f'Act-YTD-{current_fy} (Apr to Dec)', 
            f'Gr-YTD-{current_fy} (Apr to Dec)', 
            f'Ach-YTD-{current_fy} (Apr to Dec)',
            
            # 61-65. January
            f'Budget-Jan-{current_end_year}', 
            f'LY-Jan-{last_end_year}', 
            f'Act-Jan-{current_end_year}', 
            f'Gr-Jan-{current_end_year}', 
            f'Ach-Jan-{current_end_year}',
            
            # 66-70. February
            f'Budget-Feb-{current_end_year}', 
            f'LY-Feb-{last_end_year}', 
            f'Act-Feb-{current_end_year}', 
            f'Gr-Feb-{current_end_year}', 
            f'Ach-Feb-{current_end_year}',
            
            # 71-75. March
            f'Budget-Mar-{current_end_year}', 
            f'LY-Mar-{last_end_year}', 
            f'Act-Mar-{current_end_year}', 
            f'Gr-Mar-{current_end_year}', 
            f'Ach-Mar-{current_end_year}',
            
            # 76-81. Full Year YTD (Apr to Mar) - EXACTLY as specified with dash and extra Budget
            f'YTD-{current_fy} - (Apr to Mar) Budget',  # Note the DASH before (Apr to Mar)
            f'YTD-{last_fy} (Apr to Mar)LY', 
            f'Act-YTD-{current_fy} (Apr to Mar)', 
            f'Gr-YTD-{current_fy} (Apr to Mar)', 
            f'Ach-YTD-{current_fy} (Apr to Mar)',
            f'YTD-{current_fy} (Apr to Mar)Budget'  # Extra Budget column at the end
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
        """Reorder columns according to the EXACT fiscal year structure"""
        # Get the exact column order
        exact_order = self.get_exact_column_positions()
    
        # Ensure product column is first
        if product_col_name not in df.columns:
            raise ValueError(f"Product column '{product_col_name}' not found in dataframe")
    
        # Get all existing columns that are in our exact order
        existing_columns = [col for col in exact_order if col in df.columns]
    
    # Get remaining columns (excluding product column)
        other_columns = [col for col in df.columns 
                    if col not in existing_columns and col != product_col_name]
    
    # Create final column order: product column + exact order columns + other columns
        final_order = [product_col_name] + existing_columns + other_columns
    
    # Reorder the dataframe
        df = df[final_order]
    
        return df
        
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
                
                current_app.logger.info(f"TS-PW Total calculated for {col}: {sum_value:.2f}")
    
        return merged_data
    

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
        ytd_periods[f'YTD-{current_fy} - (Apr to Mar) Budget'] = [  # Note the DASH
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
        ytd_periods[f'YTD-{current_fy} (Apr to Mar)Budget'] = [  # Extra Budget column
            f'Budget-Apr-{current_start_year}', f'Budget-May-{current_start_year}', f'Budget-Jun-{current_start_year}',
            f'Budget-Jul-{current_start_year}', f'Budget-Aug-{current_start_year}', f'Budget-Sep-{current_start_year}',
            f'Budget-Oct-{current_start_year}', f'Budget-Nov-{current_start_year}', f'Budget-Dec-{current_start_year}',
            f'Budget-Jan-{current_end_year}', f'Budget-Feb-{current_end_year}', f'Budget-Mar-{current_end_year}'
        ]
        
        return ytd_periods
    

# Initialize the processor
ts_pw_merge_processor = TSPWMergePreviewProcessor()

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

def find_column(df, search_terms, threshold=80, case_sensitive=False):
    """Find column by fuzzy matching"""
    from difflib import SequenceMatcher
    
    def similarity(a, b):
        return SequenceMatcher(None, a.lower() if not case_sensitive else a, 
                             b.lower() if not case_sensitive else b).ratio() * 100
    
    for term in search_terms:
        for col in df.columns:
            if similarity(str(col), term) >= threshold:
                return col
    return None

def process_budget_data_product_region(budget_df, group_type='product_region'):
    """Process budget data for TS-PW analysis"""
    budget_df = handle_duplicate_columns(budget_df.copy())
    budget_df.columns = budget_df.columns.str.strip()

    product_col = None
    region_col = None
    product_names = ['Product', 'Product Group', 'PRODUCT NAME']
    region_names = ['Region', 'Branch', 'REGIONS']

    for col in product_names:
        if col in budget_df.columns:
            product_col = col
            break
    if not product_col:
        product_col = find_column(budget_df, product_names, threshold=80)

    for col in region_names:
        if col in budget_df.columns:
            region_col = col
            break
    if not region_col:
        region_col = find_column(budget_df, region_names, threshold=80)

    if not product_col or not region_col:
        return {'error': 'Could not find Product Group or Region column in budget dataset.'}

    budget_cols = {'Qty': [], 'Value': []}
    detailed_pattern = r'(Qty|Value)\s*[-]\s*(\w{3,})\'?(\d{2,4})'
    range_pattern = r'(Qty|Value)\s*(\w{3,})\'?(\d{2,4})[-]\s*(\w{3,})\'?(\d{2,4})'

    for col in budget_df.columns:
        detailed_match = re.match(detailed_pattern, col, re.IGNORECASE)
        if detailed_match:
            qty_or_value, month, year = detailed_match.groups()
            month = month.capitalize()
            year = year[-2:] if len(year) > 2 else year
            month_year = f"{month}-{year}"
            if qty_or_value.lower() == 'qty':
                budget_cols['Qty'].append((col, month_year))
            elif qty_or_value.lower() == 'value':
                budget_cols['Value'].append((col, month_year))
            continue

        range_match = re.match(range_pattern, col, re.IGNORECASE)
        if range_match:
            qty_or_value, start_month, start_year, end_month, end_year = range_match.groups()
            start_month = start_month.capitalize()
            start_year = start_year[-2:] if len(start_year) > 2 else start_year
            end_year = end_year[-2:] if len(end_year) > 2 else end_year
            month_year = f"{start_month}{start_year}{end_month.lower()}-{end_year}"
            if qty_or_value.lower() == 'qty':
                budget_cols['Qty'].append((col, month_year))
            elif qty_or_value.lower() == 'value':
                budget_cols['Value'].append((col, month_year))

    if not budget_cols['Qty'] and not budget_cols['Value']:
        return {'error': 'No budget quantity or value columns found.'}

    for col, _ in budget_cols['Qty'] + budget_cols['Value']:
        budget_df[col] = pd.to_numeric(budget_df[col], errors='coerce')

    group_cols = [col for col, _ in budget_cols['Qty'] + budget_cols['Value']]
    budget_data = budget_df.groupby([product_col, region_col])[group_cols].sum().reset_index()

    rename_dict = {
        product_col: 'PRODUCT NAME',
        region_col: 'Region'
    }
    for col, month_year in budget_cols['Qty']:
        rename_dict[col] = f'Budget-{month_year}_MT'
    for col, month_year in budget_cols['Value']:
        rename_dict[col] = f'Budget-{month_year}_Value'

    budget_data = budget_data.rename(columns=rename_dict)
    budget_data['PRODUCT NAME'] = budget_data['PRODUCT NAME'].str.strip().str.upper()
    budget_data['Region'] = budget_data['Region'].str.strip().str.upper()

    return budget_data

def process_sales_data(df_sales, fiscal_year_start, months):
    """
    Process current year sales data for TS-PW analysis
    Returns: tuple of (quantity_data, value_data) or (None, None) on error
    """
    try:
        # Handle multi-index columns if present
        if isinstance(df_sales.columns, pd.MultiIndex):
            df_sales.columns = ['_'.join(col).strip() for col in df_sales.columns]
        
        # Handle duplicate column names
        df_sales = handle_duplicate_columns(df_sales)
        
        # Find required columns with flexible matching
        region_col = find_column(df_sales, ['Region', 'Area', 'Zone'], case_sensitive=False)
        product_col = find_column(df_sales, ['Type (Make)', 'Type(Make)', 'Product', 'Product Group'], case_sensitive=False)
        date_col = find_column(df_sales, ['Date', 'Month Format', 'Month'], case_sensitive=False)
        qty_col = find_column(df_sales, ['Actual Quantity', 'Acutal Qty', 'Qty'], case_sensitive=False)
        value_col = find_column(df_sales, ['Value', 'Amount', 'Sales Value'], case_sensitive=False)
        
        # Rename columns to standard names for consistency
        rename_dict = {}
        if date_col: rename_dict[date_col] = 'Month_Format'
        if product_col: rename_dict[product_col] = 'Product_Group'
        if qty_col: rename_dict[qty_col] = 'Actual_Quantity'
        if value_col: rename_dict[value_col] = 'Value'
        if region_col: rename_dict[region_col] = 'Region'
        
        df_sales = df_sales.rename(columns=rename_dict)
        
        # Process quantity data if available
        qty_data = None
        if 'Actual_Quantity' in df_sales.columns and 'Product_Group' in df_sales.columns and 'Month_Format' in df_sales.columns:
            df_qty = df_sales.copy()
            
            # Filter for NORTH region if region column exists
            if 'Region' in df_qty.columns:
                df_qty = df_qty[df_qty['Region'].str.strip().str.upper() == 'NORTH']
            
            if not df_qty.empty:
                # Convert quantity to numeric
                df_qty['Actual_Quantity'] = pd.to_numeric(df_qty['Actual_Quantity'], errors='coerce')
                
                # Process month column - handle both datetime and string formats
                if pd.api.types.is_datetime64_any_dtype(df_qty['Month_Format']):
                    df_qty['Month'] = pd.to_datetime(df_qty['Month_Format']).dt.strftime('%b')
                else:
                    month_str = df_qty['Month_Format'].astype(str).str.strip()
                    try:
                        df_qty['Month'] = pd.to_datetime(month_str).dt.strftime('%b')
                    except ValueError:
                        df_qty['Month'] = month_str.str[:3]  # Take first 3 chars as month abbreviation
                
                # Group by product and month, summing quantities
                qty_agg = df_qty.groupby(['Product_Group', 'Month'])['Actual_Quantity'].sum().reset_index()
                qty_agg.columns = ['PRODUCT_NAME', 'Month', 'Actual']
                
                # Create Month_Year column with proper fiscal year tagging
                # April-December use current fiscal year (e.g., 23)
                # January-March use next fiscal year (e.g., 24)
                qty_agg['Month_Year'] = qty_agg['Month'].apply(
                    lambda x: f'Act-{x}-{str(fiscal_year_start)[-2:]}' if x in months[:9] 
                    else f'Act-{x}-{str(fiscal_year_start + 1)[-2:]}'
                )
                
                # Pivot to wide format with months as columns
                qty_data = qty_agg.pivot_table(
                    index='PRODUCT_NAME',
                    columns='Month_Year',
                    values='Actual',
                    aggfunc='sum'
                ).reset_index().fillna(0)
                qty_data['Region'] = 'NORTH'
        
        # Process value data if available
        value_data = None
        if 'Value' in df_sales.columns and 'Product_Group' in df_sales.columns and 'Month_Format' in df_sales.columns:
            df_value = df_sales.copy()
            
            # Filter for NORTH region if region column exists
            if 'Region' in df_value.columns:
                df_value = df_value[df_value['Region'].str.strip().str.upper() == 'NORTH']
            
            if not df_value.empty:
                # Convert value to numeric
                df_value['Value'] = pd.to_numeric(df_value['Value'], errors='coerce')
                
                # Process month column - handle both datetime and string formats
                if pd.api.types.is_datetime64_any_dtype(df_value['Month_Format']):
                    df_value['Month'] = pd.to_datetime(df_value['Month_Format']).dt.strftime('%b')
                else:
                    month_str = df_value['Month_Format'].astype(str).str.strip()
                    try:
                        df_value['Month'] = pd.to_datetime(month_str).dt.strftime('%b')
                    except ValueError:
                        df_value['Month'] = month_str.str[:3]  # Take first 3 chars as month abbreviation
                
                # Group by product and month, summing values
                value_agg = df_value.groupby(['Product_Group', 'Month'])['Value'].sum().reset_index()
                value_agg.columns = ['PRODUCT_NAME', 'Month', 'Actual']
                
                # Create Month_Year column with proper fiscal year tagging
                # April-December use current fiscal year (e.g., 23)
                # January-March use next fiscal year (e.g., 24)
                value_agg['Month_Year'] = value_agg['Month'].apply(
                    lambda x: f'Act-{x}-{str(fiscal_year_start)[-2:]}' if x in months[:9] 
                    else f'Act-{x}-{str(fiscal_year_start + 1)[-2:]}'
                )
                
                # Pivot to wide format with months as columns
                value_data = value_agg.pivot_table(
                    index='PRODUCT_NAME',
                    columns='Month_Year',
                    values='Actual',
                    aggfunc='sum'
                ).reset_index().fillna(0)
                value_data['Region'] = 'NORTH'
        
        return qty_data, value_data
    
    except Exception as e:
        print(f"Error processing sales data: {str(e)}")
        return None, None
    
def build_exact_columns_and_calculate_values(data_df, fiscal_info, analysis_type='mt'):
    """Build exact column structure and calculate all values for TS-PW"""
    
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
        
        # Check if actual column has any non-zero values before calculating Gr and Ach
        has_actual_data = (data_df[act_col] > 0.01).any()
        
        if has_actual_data:
            # Calculate Growth only if we have actual data
            data_df[gr_col] = np.where(
                data_df[ly_col] > 0.01,
                ((data_df[act_col] - data_df[ly_col]) / data_df[ly_col] * 100).round(2),
                0
            )
            
            # Calculate Achievement only if we have actual data
            data_df[ach_col] = np.where(
                data_df[budget_col] > 0.01,
                (data_df[act_col] / data_df[budget_col] * 100).round(2),
                0
            )
        else:
            # If no actual data, set Gr and Ach to 0
            data_df[gr_col] = 0.0
            data_df[ach_col] = 0.0
    
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
    
    # Check if Q1 YTD actual has any non-zero values
    has_q1_actual_data = (data_df[q1_act_col] > 0.01).any()
    
    if has_q1_actual_data:
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
    else:
        data_df[q1_gr_col] = 0.0
        data_df[q1_ach_col] = 0.0
    
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
    
    # Check if H1 YTD actual has any non-zero values
    has_h1_actual_data = (data_df[h1_act_col] > 0.01).any()
    
    if has_h1_actual_data:
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
    else:
        data_df[h1_gr_col] = 0.0
        data_df[h1_ach_col] = 0.0
    
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
    
    # Check if 9M YTD actual has any non-zero values
    has_nm_actual_data = (data_df[nm_act_col] > 0.01).any()
    
    if has_nm_actual_data:
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
    else:
        data_df[nm_gr_col] = 0.0
        data_df[nm_ach_col] = 0.0

    # Full Year YTD (Apr to Mar) - Position after March columns - EXACT as specified
    fy_budget_col = f'YTD-{current_fy} - (Apr to Mar) Budget'  # Note the DASH
    fy_ly_col = f'YTD-{last_fy} (Apr to Mar)LY'
    fy_act_col = f'Act-YTD-{current_fy} (Apr to Mar)'
    fy_gr_col = f'Gr-YTD-{current_fy} (Apr to Mar)'
    fy_ach_col = f'Ach-YTD-{current_fy} (Apr to Mar)'
    fy_extra_budget_col = f'YTD-{current_fy} (Apr to Mar)Budget'  # Extra Budget column
    
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
    
    # Check if Full Year YTD actual has any non-zero values
    has_fy_actual_data = (data_df[fy_act_col] > 0.01).any()
    
    if has_fy_actual_data:
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
    else:
        data_df[fy_gr_col] = 0.0
        data_df[fy_ach_col] = 0.0
    
    # Extra Budget column (same as main budget column)
    data_df[fy_extra_budget_col] = data_df[fy_budget_col]
    
    return data_df


def calculate_ytd_growth_achievement(self, merged_data):
    """Calculate YTD Growth and Achievement with EXACT column names - UPDATED VERSION"""
    current_fy = self.fiscal_info['fiscal_year_str']
    last_fy = self.fiscal_info['last_fiscal_year_str']
    
    ytd_pairs = [
        ('Apr to Jun', f'YTD-{current_fy} (Apr to Jun)Budget', 
         f'YTD-{last_fy} (Apr to Jun)LY', f'Act-YTD-{current_fy} (Apr to Jun)'),
        ('Apr to Sep', f'YTD-{current_fy} (Apr to Sep)Budget', 
         f'YTD-{last_fy} (Apr to Sep)LY', f'Act-YTD-{current_fy} (Apr to Sep)'),
        ('Apr to Dec', f'YTD-{current_fy} (Apr to Dec)Budget', 
         f'YTD-{last_fy} (Apr to Dec)LY', f'Act-YTD-{current_fy} (Apr to Dec)'),
        ('Apr to Mar', f'YTD-{current_fy} - (Apr to Mar) Budget',  # Note the DASH
         f'YTD-{last_fy} (Apr to Mar)LY', f'Act-YTD-{current_fy} (Apr to Mar)')
    ]
    
    for period, budget_col, ly_col, act_col in ytd_pairs:
        # Check if all required columns exist
        required_cols = [budget_col, ly_col, act_col]
        existing_cols = [col for col in required_cols if col in merged_data.columns]
        
        if len(existing_cols) >= 2:  # At least 2 out of 3 columns needed for calculations
            # Check if actual column has any non-zero values before calculating Gr and Ach
            has_actual_data = False
            if act_col in merged_data.columns:
                has_actual_data = (merged_data[act_col] > 0.01).any()
            
            if has_actual_data:
                # Growth calculation (if both LY and Actual exist)
                if ly_col in merged_data.columns and act_col in merged_data.columns:
                    gr_col = f'Gr-YTD-{current_fy} ({period})'
                    merged_data[gr_col] = np.where(
                        merged_data[ly_col] > 0.01,
                        ((merged_data[act_col] - merged_data[ly_col]) / merged_data[ly_col] * 100).round(2),
                        0
                    )
                    current_app.logger.info(f"Calculated Growth column: {gr_col}")
                
                # Achievement calculation (if both Budget and Actual exist)
                if budget_col in merged_data.columns and act_col in merged_data.columns:
                    ach_col = f'Ach-YTD-{current_fy} ({period})'
                    merged_data[ach_col] = np.where(
                        merged_data[budget_col] > 0.01,
                        (merged_data[act_col] / merged_data[budget_col] * 100).round(2),
                        0
                    )
                    current_app.logger.info(f"Calculated Achievement column: {ach_col}")
            else:
                # If no actual data, set Gr and Ach to 0
                gr_col = f'Gr-YTD-{current_fy} ({period})'
                ach_col = f'Ach-YTD-{current_fy} ({period})'
                
                if gr_col not in merged_data.columns:
                    merged_data[gr_col] = 0.0
                if ach_col not in merged_data.columns:
                    merged_data[ach_col] = 0.0
                    
                current_app.logger.info(f"No actual data found for {period}, set Gr and Ach to 0")
        else:
            current_app.logger.warning(f"Insufficient columns for YTD calculations in period {period}. Found: {existing_cols}")
    
    return merged_data
def remove_specific_unwanted_columns(df, product_col_name):
    """Remove ONLY Region column and budget range columns like Budget-April24dec-24"""
    
    # Start with all columns
    columns_to_keep = []
    
    for col in df.columns:
        col_lower = str(col).lower().strip()
        should_remove = False
        
        # Remove Region column
        if col_lower == 'region':
            current_app.logger.info(f"Removing Region column: {col}")
            should_remove = True
        
        # Remove budget range columns like "Budget-April24dec-24"
        elif re.search(r'budget.*april.*dec', col_lower, re.IGNORECASE):
            current_app.logger.info(f"Removing budget range column: {col}")
            should_remove = True
        
        # Remove any column with format like "Budget-[Month][Year][Month]-[Year]"
        elif re.search(r'budget.*\w{3,}\d{2}.*\w{3,}.*\d{2}', col_lower, re.IGNORECASE):
            current_app.logger.info(f"Removing budget range column: {col}")
            should_remove = True
        
        # Keep all other columns
        if not should_remove:
            columns_to_keep.append(col)
    
    # Return filtered dataframe
    filtered_df = df[columns_to_keep].copy()
    current_app.logger.info(f"Removed {len(df.columns) - len(filtered_df.columns)} unwanted columns")
    current_app.logger.info(f"Columns removed: {set(df.columns) - set(filtered_df.columns)}")
    
    return filtered_df

# Main TS-PW Analysis Processing 


@ts_pw_bp.route('/process-ts-pw', methods=['POST'])
def process_ts_pw_analysis():
    """Process TS-PW data analysis - UPDATED with improved LY data logic"""
    try:
        data = request.json
        budget_filepath = data.get('budget_filepath')
        budget_sheet = data.get('budget_sheet')
        sales_filepath = data.get('sales_filepath')
        sales_sheets = data.get('sales_sheets', [])
        last_year_filepath = data.get('last_year_filepath')
        last_year_sheet = data.get('last_year_sheet')
        
        if not budget_filepath or not budget_sheet:
            return jsonify({
                'success': False,
                'error': 'Budget file and sheet are required'
            }), 400
        
        # Get current fiscal year
        fiscal_info = ts_pw_merge_processor.fiscal_info
        months = ts_pw_merge_processor.months
        
        # Process budget data (using original function - no changes to budget processing)
        xls_budget = pd.ExcelFile(budget_filepath)
        df_budget = pd.read_excel(xls_budget, sheet_name=budget_sheet)
        df_budget.columns = df_budget.columns.str.strip()
        df_budget = df_budget.dropna(how='all').reset_index(drop=True)
        
        budget_data = process_budget_data_product_region(df_budget, group_type='product_region')
        
        if isinstance(budget_data, dict) and 'error' in budget_data:
            return jsonify({
                'success': False,
                'error': budget_data['error']
            }), 400
        
        # Filter for NORTH region (but will remove Region column later)
        if 'Region' in budget_data.columns:
            budget_data = budget_data[budget_data['Region'].str.strip().str.upper() == 'NORTH']
            if budget_data.empty:
                return jsonify({
                    'success': False,
                    'error': 'No data found for NORTH region.'
                }), 400
        
        # Extract all products
        all_products = set(budget_data['PRODUCT NAME'].dropna().astype(str).str.strip().str.upper())
        
        # Process current year sales data
        actual_mt_current = {}
        actual_value_current = {}
        
        if sales_filepath and sales_sheets:
            xls_sales = pd.ExcelFile(sales_filepath)
            
            for sheet_name in sales_sheets:
                try:
                    df_sales = pd.read_excel(xls_sales, sheet_name=sheet_name, header=0)
                    
                    if isinstance(df_sales.columns, pd.MultiIndex):
                        df_sales.columns = ['_'.join(col).strip() for col in df_sales.columns.values]
                    df_sales = handle_duplicate_columns(df_sales)
                    
                    region_col = find_column(df_sales, ['Region', 'Area', 'Zone'], case_sensitive=False)
                    product_col = find_column(df_sales, ['Type (Make)', 'Type(Make)', 'Product'], case_sensitive=False)
                    date_col = find_column(df_sales, ['Date', 'Month Format', 'Month'], case_sensitive=False)
                    qty_col = find_column(df_sales, ['Actual Quantity', 'Quantity', 'Qty'], case_sensitive=False)
                    value_col = find_column(df_sales, ['Value', 'Amount', 'Sales Value'], case_sensitive=False)
                    
                    if product_col:
                        # Filter for NORTH region
                        if region_col:
                            df_sales = df_sales[df_sales[region_col].str.strip().str.upper() == 'NORTH']
                        
                        # Add products to the set
                        unique_products = df_sales[product_col].dropna().astype(str).str.strip().str.upper()
                        all_products.update(unique_products)
                        
                        # Process actual sales data if we have required columns
                        if date_col and not df_sales.empty:
                            # Process date column
                            if pd.api.types.is_datetime64_any_dtype(df_sales[date_col]):
                                df_sales['Month'] = pd.to_datetime(df_sales[date_col]).dt.strftime('%b')
                            else:
                                month_str = df_sales[date_col].astype(str).str.strip()
                                try:
                                    df_sales['Month'] = pd.to_datetime(month_str, format='%Y-%m', errors='coerce').dt.strftime('%b')
                                except ValueError:
                                    df_sales['Month'] = month_str.str[:3]
                            
                            # Process quantity data
                            if qty_col:
                                df_qty = df_sales[[product_col, 'Month', qty_col]].copy()
                                df_qty[qty_col] = pd.to_numeric(df_qty[qty_col], errors='coerce')
                                df_qty = df_qty.dropna(subset=[qty_col, 'Month'])
                                df_qty = df_qty[df_qty[qty_col] != 0]
                                
                                grouped = df_qty.groupby([product_col, 'Month'])[qty_col].sum().reset_index()
                                
                                for _, row in grouped.iterrows():
                                    product = str(row[product_col]).strip().upper()
                                    month = row['Month']
                                    qty = row[qty_col]
                                    
                                    year = str(fiscal_info['fiscal_year_start'])[-2:] if month in months[:9] else str(fiscal_info['fiscal_year_end'])[-2:]
                                    col_name = f'Act-{month}-{year}'
                                    
                                    if product not in actual_mt_current:
                                        actual_mt_current[product] = {}
                                    
                                    if col_name in actual_mt_current[product]:
                                        actual_mt_current[product][col_name] += qty
                                    else:
                                        actual_mt_current[product][col_name] = qty
                            
                            # Process value data
                            if value_col:
                                df_value = df_sales[[product_col, 'Month', value_col]].copy()
                                df_value[value_col] = pd.to_numeric(df_value[value_col], errors='coerce')
                                df_value = df_value.dropna(subset=[value_col, 'Month'])
                                df_value = df_value[df_value[value_col] != 0]
                                
                                grouped = df_value.groupby([product_col, 'Month'])[value_col].sum().reset_index()
                                
                                for _, row in grouped.iterrows():
                                    product = str(row[product_col]).strip().upper()
                                    month = row['Month']
                                    value = row[value_col]
                                    
                                    year = str(fiscal_info['fiscal_year_start'])[-2:] if month in months[:9] else str(fiscal_info['fiscal_year_end'])[-2:]
                                    col_name = f'Act-{month}-{year}'
                                    
                                    if product not in actual_value_current:
                                        actual_value_current[product] = {}
                                    
                                    if col_name in actual_value_current[product]:
                                        actual_value_current[product][col_name] += value
                                    else:
                                        actual_value_current[product][col_name] = value
                                        
                except Exception as e:
                    current_app.logger.warning(f"Error processing sales sheet {sheet_name}: {str(e)}")
                    continue
        
        # ========== IMPROVED LAST YEAR DATA PROCESSING LOGIC ==========        
        actual_mt_last = {}
        actual_value_last = {}
        ly_processing_method = "none"
        
        # FIRST ATTEMPT: Try original last year file logic
        if last_year_filepath and last_year_sheet:
            try:
                current_app.logger.info("Attempting FIRST method: Original last year file logic")
                xls_last_year = pd.ExcelFile(last_year_filepath)
                df_last_year = pd.read_excel(xls_last_year, sheet_name=last_year_sheet)
                
                if isinstance(df_last_year.columns, pd.MultiIndex):
                    df_last_year.columns = ['_'.join(col).strip() for col in df_last_year.columns.values]
                df_last_year = handle_duplicate_columns(df_last_year)
                
                region_col = find_column(df_last_year, ['Region'], case_sensitive=False)
                product_col = find_column(df_last_year, ['Type (Make)', 'Type(Make)', 'Product Group'], case_sensitive=False)
                date_col = find_column(df_last_year, ['Date', 'Month Format', 'Month'], case_sensitive=False)
                qty_col = find_column(df_last_year, ['Actual Quantity', 'Acutal Quantity', 'Qty'], case_sensitive=False)
                amount_col = find_column(df_last_year, ['Amount', 'Value', 'Sales Value'], case_sensitive=False)
                
                if product_col and date_col:
                    # Filter for NORTH region
                    if region_col:
                        df_last_year = df_last_year[df_last_year[region_col].str.strip().str.upper() == 'NORTH']
                    
                    # Add products to the set
                    unique_products = df_last_year[product_col].dropna().astype(str).str.strip().str.upper()
                    all_products.update(unique_products)
                    
                    if not df_last_year.empty:
                        # Process date column
                        if pd.api.types.is_datetime64_any_dtype(df_last_year[date_col]):
                            df_last_year['Month'] = pd.to_datetime(df_last_year[date_col]).dt.strftime('%b')
                        else:
                            month_str = df_last_year[date_col].astype(str).str.strip()
                            try:
                                df_last_year['Month'] = pd.to_datetime(month_str, format='%Y-%m', errors='coerce').dt.strftime('%b')
                            except ValueError:
                                df_last_year['Month'] = month_str.str[:3]
                        
                        # Process quantity data for last year
                        if qty_col:
                            df_last_year_qty = df_last_year.copy()
                            df_last_year_qty[qty_col] = pd.to_numeric(df_last_year_qty[qty_col], errors='coerce')
                            df_last_year_qty = df_last_year_qty.dropna(subset=[qty_col, 'Month'])
                            df_last_year_qty = df_last_year_qty[df_last_year_qty[qty_col] != 0]
                            
                            if not df_last_year_qty.empty:
                                grouped = df_last_year_qty.groupby([product_col, 'Month'])[qty_col].sum().reset_index()
                                
                                for _, row in grouped.iterrows():
                                    product = str(row[product_col]).strip().upper()
                                    month = row['Month']
                                    qty = row[qty_col]
                                    
                                    year = str(fiscal_info['last_fiscal_year_start'])[-2:] if month in months[:9] else str(fiscal_info['last_fiscal_year_end'])[-2:]
                                    col_name = f'LY-{month}-{year}'
                                    
                                    if product not in actual_mt_last:
                                        actual_mt_last[product] = {}
                                    
                                    if col_name in actual_mt_last[product]:
                                        actual_mt_last[product][col_name] += qty
                                    else:
                                        actual_mt_last[product][col_name] = qty
                        
                        # Process value data for last year
                        if amount_col:
                            df_last_year_val = df_last_year.copy()
                            df_last_year_val[amount_col] = pd.to_numeric(df_last_year_val[amount_col], errors='coerce')
                            df_last_year_val = df_last_year_val.dropna(subset=[amount_col, 'Month'])
                            df_last_year_val = df_last_year_val[df_last_year_val[amount_col] != 0]
                            
                            if not df_last_year_val.empty:
                                grouped = df_last_year_val.groupby([product_col, 'Month'])[amount_col].sum().reset_index()
                                
                                for _, row in grouped.iterrows():
                                    product = str(row[product_col]).strip().upper()
                                    month = row['Month']
                                    value = row[amount_col]
                                    
                                    year = str(fiscal_info['last_fiscal_year_start'])[-2:] if month in months[:9] else str(fiscal_info['last_fiscal_year_end'])[-2:]
                                    col_name = f'LY-{month}-{year}'
                                    
                                    if product not in actual_value_last:
                                        actual_value_last[product] = {}
                                    
                                    if col_name in actual_value_last[product]:
                                        actual_value_last[product][col_name] += value
                                    else:
                                        actual_value_last[product][col_name] = value
                
                # Check if we got meaningful data from the first method
                total_ly_records = len(actual_mt_last) + len(actual_value_last)
                if total_ly_records > 0:
                    ly_processing_method = "original_last_year_file"
                    current_app.logger.info(f"SUCCESS: First method worked - Got {total_ly_records} LY records")
                else:
                    raise ValueError("No meaningful data found in last year file")
                            
            except Exception as e:
                current_app.logger.warning(f"FIRST method failed: {str(e)}")
                # Reset the data since first method failed
                actual_mt_last = {}
                actual_value_last = {}
        
        # SECOND ATTEMPT: If first method failed or no last year file, use sales sheet logic
        if ly_processing_method == "none" and sales_filepath and sales_sheets:
            try:
                current_app.logger.info("Attempting SECOND method: Sales sheet logic for LY data")
                xls_sales = pd.ExcelFile(sales_filepath)
                
                for sheet_name in sales_sheets:
                    try:
                        df_sales_ly = pd.read_excel(xls_sales, sheet_name=sheet_name, header=0)
                        
                        if isinstance(df_sales_ly.columns, pd.MultiIndex):
                            df_sales_ly.columns = ['_'.join(col).strip() for col in df_sales_ly.columns.values]
                        df_sales_ly = handle_duplicate_columns(df_sales_ly)
                        
                        region_col = find_column(df_sales_ly, ['Region', 'Area', 'Zone'], case_sensitive=False)
                        product_col = find_column(df_sales_ly, ['Type (Make)', 'Type(Make)', 'Product'], case_sensitive=False)
                        date_col = find_column(df_sales_ly, ['Date', 'Month Format', 'Month'], case_sensitive=False)
                        qty_col = find_column(df_sales_ly, ['Actual Quantity', 'Quantity', 'Qty'], case_sensitive=False)
                        value_col = find_column(df_sales_ly, ['Value', 'Amount', 'Sales Value'], case_sensitive=False)
                        
                        if product_col and date_col:
                            # Filter for NORTH region
                            if region_col:
                                df_sales_ly = df_sales_ly[df_sales_ly[region_col].str.strip().str.upper() == 'NORTH']
                            
                            if not df_sales_ly.empty:
                                # Process date column to identify last year data
                                if pd.api.types.is_datetime64_any_dtype(df_sales_ly[date_col]):
                                    df_sales_ly['Year'] = pd.to_datetime(df_sales_ly[date_col]).dt.year
                                    df_sales_ly['Month'] = pd.to_datetime(df_sales_ly[date_col]).dt.strftime('%b')
                                else:
                                    month_str = df_sales_ly[date_col].astype(str).str.strip()
                                    try:
                                        df_sales_ly['Year'] = pd.to_datetime(month_str, format='%Y-%m', errors='coerce').dt.year
                                        df_sales_ly['Month'] = pd.to_datetime(month_str, format='%Y-%m', errors='coerce').dt.strftime('%b')
                                    except ValueError:
                                        # Try to extract year and month from string
                                        df_sales_ly['Month'] = month_str.str[:3]
                                        df_sales_ly['Year'] = fiscal_info['last_fiscal_year_start']  # Default to last fiscal year
                                
                                # Filter for last fiscal year data
                                ly_start_year = fiscal_info['last_fiscal_year_start']
                                ly_end_year = fiscal_info['last_fiscal_year_end']
                                
                                # Create a mask for last year fiscal year data
                                ly_mask = (
                                    ((df_sales_ly['Year'] == ly_start_year) & (df_sales_ly['Month'].isin(['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']))) |
                                    ((df_sales_ly['Year'] == ly_end_year) & (df_sales_ly['Month'].isin(['Jan', 'Feb', 'Mar'])))
                                )
                                
                                df_sales_ly_filtered = df_sales_ly[ly_mask]
                                
                                if not df_sales_ly_filtered.empty:
                                    # Process quantity data for last year
                                    if qty_col:
                                        df_qty = df_sales_ly_filtered[[product_col, 'Month', qty_col]].copy()
                                        df_qty[qty_col] = pd.to_numeric(df_qty[qty_col], errors='coerce')
                                        df_qty = df_qty.dropna(subset=[qty_col, 'Month'])
                                        df_qty = df_qty[df_qty[qty_col] != 0]
                                        
                                        grouped = df_qty.groupby([product_col, 'Month'])[qty_col].sum().reset_index()
                                        
                                        for _, row in grouped.iterrows():
                                            product = str(row[product_col]).strip().upper()
                                            month = row['Month']
                                            qty = row[qty_col]
                                            
                                            year = str(fiscal_info['last_fiscal_year_start'])[-2:] if month in months[:9] else str(fiscal_info['last_fiscal_year_end'])[-2:]
                                            col_name = f'LY-{month}-{year}'
                                            
                                            if product not in actual_mt_last:
                                                actual_mt_last[product] = {}
                                            
                                            if col_name in actual_mt_last[product]:
                                                actual_mt_last[product][col_name] += qty
                                            else:
                                                actual_mt_last[product][col_name] = qty
                                    
                                    # Process value data for last year
                                    if value_col:
                                        df_value = df_sales_ly_filtered[[product_col, 'Month', value_col]].copy()
                                        df_value[value_col] = pd.to_numeric(df_value[value_col], errors='coerce')
                                        df_value = df_value.dropna(subset=[value_col, 'Month'])
                                        df_value = df_value[df_value[value_col] != 0]
                                        
                                        grouped = df_value.groupby([product_col, 'Month'])[value_col].sum().reset_index()
                                        
                                        for _, row in grouped.iterrows():
                                            product = str(row[product_col]).strip().upper()
                                            month = row['Month']
                                            value = row[value_col]
                                            
                                            year = str(fiscal_info['last_fiscal_year_start'])[-2:] if month in months[:9] else str(fiscal_info['last_fiscal_year_end'])[-2:]
                                            col_name = f'LY-{month}-{year}'
                                            
                                            if product not in actual_value_last:
                                                actual_value_last[product] = {}
                                            
                                            if col_name in actual_value_last[product]:
                                                actual_value_last[product][col_name] += value
                                            else:
                                                actual_value_last[product][col_name] = value
                    
                    except Exception as sheet_error:
                        current_app.logger.warning(f"Error processing sales sheet {sheet_name} for LY data: {str(sheet_error)}")
                        continue
                
                # Check if we got meaningful data from the second method
                total_ly_records = len(actual_mt_last) + len(actual_value_last)
                if total_ly_records > 0:
                    ly_processing_method = "sales_sheet_logic"
                    current_app.logger.info(f"SUCCESS: Second method worked - Got {total_ly_records} LY records from sales sheets")
                else:
                    current_app.logger.warning("Second method also failed - No LY data found")
                    
            except Exception as e:
                current_app.logger.warning(f"SECOND method failed: {str(e)}")
        
        # Log the final LY processing result
        total_ly_mt_records = sum(len(data) for data in actual_mt_last.values())
        total_ly_value_records = sum(len(data) for data in actual_value_last.values())
        current_app.logger.info(f"Final LY processing result: Method={ly_processing_method}, MT records={total_ly_mt_records}, Value records={total_ly_value_records}")
        
        # ========== CONTINUE WITH ORIGINAL PROCESSING LOGIC ==========
        
        # Process MT (Tonnage) data
        mt_cols = [col for col in budget_data.columns if col.endswith('_MT')]
        mt_result = None
        mt_columns = []
        
        if mt_cols:
            # Initialize result with all products
            result_ts_pw_mt = pd.DataFrame({
                'PRODUCT NAME': sorted(list(all_products)),
                'Region': 'NORTH'
            })
            
            # Populate budget data
            month_cols = sorted(set(col.replace('_MT', '') for col in mt_cols if 'Budget-' in col))
            for month_col in month_cols:
                if f'{month_col}_MT' in budget_data.columns:
                    result_ts_pw_mt[month_col] = 0.0
                    for _, row in budget_data.iterrows():
                        product = row['PRODUCT NAME']
                        value = row[f'{month_col}_MT']
                        if pd.notna(value):
                            idx = result_ts_pw_mt[result_ts_pw_mt['PRODUCT NAME'] == product].index
                            if not idx.empty:
                                result_ts_pw_mt.loc[idx, month_col] = value
            
            # Merge current year actual MT data
            for product, data in actual_mt_current.items():
                idx = result_ts_pw_mt[result_ts_pw_mt['PRODUCT NAME'] == product].index
                if idx.empty:
                    new_row = pd.DataFrame({'PRODUCT NAME': [product], 'Region': ['NORTH']})
                    result_ts_pw_mt = pd.concat([result_ts_pw_mt, new_row], ignore_index=True)
                    idx = result_ts_pw_mt[result_ts_pw_mt['PRODUCT NAME'] == product].index
                
                for col, value in data.items():
                    if col not in result_ts_pw_mt.columns:
                        result_ts_pw_mt[col] = 0.0
                    result_ts_pw_mt.loc[idx, col] = value
            
            # Merge last year actual MT data
            for product, data in actual_mt_last.items():
                idx = result_ts_pw_mt[result_ts_pw_mt['PRODUCT NAME'] == product].index
                if idx.empty:
                    new_row = pd.DataFrame({'PRODUCT NAME': [product], 'Region': ['NORTH']})
                    result_ts_pw_mt = pd.concat([result_ts_pw_mt, new_row], ignore_index=True)
                    idx = result_ts_pw_mt[result_ts_pw_mt['PRODUCT NAME'] == product].index
                
                for col, value in data.items():
                    if col not in result_ts_pw_mt.columns:
                        result_ts_pw_mt[col] = 0.0
                    result_ts_pw_mt.loc[idx, col] = value
            
            # Build EXACT column structure and calculate values
            result_ts_pw_mt = build_exact_columns_and_calculate_values(result_ts_pw_mt, fiscal_info, 'mt')
            
            # Calculate totals - SUM ALL COLUMNS INCLUDING Gr and Ach
            exclude_products = ['NORTH TOTAL', 'WEST SALES', 'GRAND TOTAL', 'TOTAL SALES']
            valid_products = result_ts_pw_mt[~result_ts_pw_mt['PRODUCT NAME'].isin(exclude_products)]
            
            total_row = {'PRODUCT NAME': 'TOTAL SALES'}
            numeric_cols = result_ts_pw_mt.select_dtypes(include=[np.number]).columns
            
            # Sum ALL numeric columns - no exceptions for Gr and Ach
            for col in numeric_cols:
                total_value = valid_products[col].sum()
                total_row[col] = round(total_value, 2)
                current_app.logger.info(f"TS-PW MT Total calculated for {col}: {total_value:.2f}")
            
            result_ts_pw_mt = pd.concat([result_ts_pw_mt, pd.DataFrame([total_row])], ignore_index=True)
            result_ts_pw_mt = result_ts_pw_mt.rename(columns={'PRODUCT NAME': 'SALES in Tonage'})
            
            # Remove specific unwanted columns
            result_ts_pw_mt = remove_specific_unwanted_columns_fixed(result_ts_pw_mt, 'SALES in Tonage')
            
            # Apply EXACT fiscal year column ordering
            result_ts_pw_mt = ts_pw_merge_processor.reorder_columns_exact_fiscal_year(result_ts_pw_mt, 'SALES in Tonage')
            
            mt_result = result_ts_pw_mt.fillna(0).round(2).to_dict('records')
            mt_columns = list(result_ts_pw_mt.columns)
        
        # Process Value data
        value_cols = [col for col in budget_data.columns if col.endswith('_Value')]
        value_result = None
        value_columns = []
        
        if value_cols:
            # Initialize result with all products
            result_ts_pw_value = pd.DataFrame({
                'PRODUCT NAME': sorted(list(all_products)),
                'Region': 'NORTH'
            })
            
            # Populate budget data
            month_cols = sorted(set(col.replace('_Value', '') for col in value_cols if 'Budget-' in col))
            for month_col in month_cols:
                if f'{month_col}_Value' in budget_data.columns:
                    result_ts_pw_value[month_col] = 0.0
                    for _, row in budget_data.iterrows():
                        product = row['PRODUCT NAME']
                        value = row[f'{month_col}_Value']
                        if pd.notna(value):
                            idx = result_ts_pw_value[result_ts_pw_value['PRODUCT NAME'] == product].index
                            if not idx.empty:
                                result_ts_pw_value.loc[idx, month_col] = value
            
            # Merge current year actual Value data
            for product, data in actual_value_current.items():
                idx = result_ts_pw_value[result_ts_pw_value['PRODUCT NAME'] == product].index
                if idx.empty:
                    new_row = pd.DataFrame({'PRODUCT NAME': [product], 'Region': ['NORTH']})
                    result_ts_pw_value = pd.concat([result_ts_pw_value, new_row], ignore_index=True)
                    idx = result_ts_pw_value[result_ts_pw_value['PRODUCT NAME'] == product].index
                
                for col, value in data.items():
                    if col not in result_ts_pw_value.columns:
                        result_ts_pw_value[col] = 0.0
                    result_ts_pw_value.loc[idx, col] = value
            
            # Merge last year actual Value data
            for product, data in actual_value_last.items():
                idx = result_ts_pw_value[result_ts_pw_value['PRODUCT NAME'] == product].index
                if idx.empty:
                    new_row = pd.DataFrame({'PRODUCT NAME': [product], 'Region': ['NORTH']})
                    result_ts_pw_value = pd.concat([result_ts_pw_value, new_row], ignore_index=True)
                    idx = result_ts_pw_value[result_ts_pw_value['PRODUCT NAME'] == product].index
                
                for col, value in data.items():
                    if col not in result_ts_pw_value.columns:
                        result_ts_pw_value[col] = 0.0
                    result_ts_pw_value.loc[idx, col] = value
            
            # Build EXACT column structure and calculate values
            result_ts_pw_value = build_exact_columns_and_calculate_values(result_ts_pw_value, fiscal_info, 'value')
            
            # Calculate totals - SUM ALL COLUMNS INCLUDING Gr and Ach
            valid_products = result_ts_pw_value[~result_ts_pw_value['PRODUCT NAME'].isin(exclude_products)]
            
            total_row = {'PRODUCT NAME': 'TOTAL SALES'}
            numeric_cols = result_ts_pw_value.select_dtypes(include=[np.number]).columns
            
            # Sum ALL numeric columns - no exceptions for Gr and Ach
            for col in numeric_cols:
                total_value = valid_products[col].sum()
                total_row[col] = round(total_value, 2)
                current_app.logger.info(f"TS-PW Value Total calculated for {col}: {total_value:.2f}")
            
            result_ts_pw_value = pd.concat([result_ts_pw_value, pd.DataFrame([total_row])], ignore_index=True)
            result_ts_pw_value = result_ts_pw_value.rename(columns={'PRODUCT NAME': 'SALES in Value'})
            
            # Remove specific unwanted columns
            result_ts_pw_value = remove_specific_unwanted_columns_fixed(result_ts_pw_value, 'SALES in Value')
            
            # Apply EXACT fiscal year column ordering
            result_ts_pw_value = ts_pw_merge_processor.reorder_columns_exact_fiscal_year(result_ts_pw_value, 'SALES in Value')
            
            value_result = result_ts_pw_value.fillna(0).round(2).to_dict('records')
            value_columns = list(result_ts_pw_value.columns)
        
        return jsonify({
            'success': True,
            'data': {
                'mt_data': mt_result,
                'value_data': value_result,
                'fiscal_year': fiscal_info['fiscal_year_str'],
                'shape': {
                    'mt': [len(mt_result), len(mt_columns)] if mt_result else [0, 0],
                    'value': [len(value_result), len(value_columns)] if value_result else [0, 0]
                },
                'columns': {
                    'mt': mt_columns,
                    'value': value_columns
                },
                'ly_processing_info': {
                    'method_used': ly_processing_method,
                    'mt_ly_records': total_ly_mt_records,
                    'value_ly_records': total_ly_value_records,
                    'total_ly_products': len(set(list(actual_mt_last.keys()) + list(actual_value_last.keys()))),
                    'description': {
                        'original_last_year_file': 'Used dedicated last year file with original logic',
                        'sales_sheet_logic': 'Used sales sheet data filtered for last fiscal year',
                        'none': 'No last year data processed'
                    }.get(ly_processing_method, 'Unknown method')
                },
                'removed_columns_info': {
                    'region_column_removed': True,
                    'budget_range_columns_removed': True,
                    'duplicate_ytd_budget_removed': True,
                    'other_budget_columns_kept': True
                }
            }
        })
    
    except Exception as e:
        current_app.logger.error(f"Error in TS-PW analysis: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

def remove_specific_unwanted_columns_fixed(df, product_col_name):
    """Remove Region column, budget range columns like Budget-April24dec-24, AND duplicate YTD Budget column"""
    
    # Start with all columns
    columns_to_keep = []
    
    for col in df.columns:
        col_lower = str(col).lower().strip()
        should_remove = False
        
        # Remove Region column
        if col_lower == 'region':
            current_app.logger.info(f"Removing Region column: {col}")
            should_remove = True
        
        # Remove budget range columns like "Budget-April24dec-24"
        elif re.search(r'budget.*april.*dec', col_lower, re.IGNORECASE):
            current_app.logger.info(f"Removing budget range column: {col}")
            should_remove = True
        
        # Remove any column with format like "Budget-[Month][Year][Month]-[Year]"
        elif re.search(r'budget.*\w{3,}\d{2}.*\w{3,}.*\d{2}', col_lower, re.IGNORECASE):
            current_app.logger.info(f"Removing budget range column: {col}")
            should_remove = True
        
        # ✅ NEW: Remove duplicate YTD Budget column (without dash)
        elif col.endswith('(Apr to Mar)Budget') and not '- (Apr to Mar) Budget' in col:
            current_app.logger.info(f"Removing duplicate YTD Budget column: {col}")
            should_remove = True
        
        # Keep all other columns
        if not should_remove:
            columns_to_keep.append(col)
    
    # Return filtered dataframe
    filtered_df = df[columns_to_keep].copy()
    current_app.logger.info(f"Removed {len(df.columns) - len(filtered_df.columns)} unwanted columns")
    current_app.logger.info(f"Columns removed: {set(df.columns) - set(filtered_df.columns)}")
    
    return filtered_df

@ts_pw_bp.route('/export-ts-pw-excel', methods=['POST'])
def export_ts_pw_excel():
    """Export TS-PW data to Excel with enhanced formatting matching region code style"""
    try:
        data = request.json
        data_type = data.get('data_type', 'mt')  # 'mt' or 'value'
        table_data = data.get('data')
        columns = data.get('columns', [])
        fiscal_year = data.get('fiscal_year', '25-26')
        
        if not table_data:
            return jsonify({
                'success': False,
                'error': 'No data provided for export'
            }), 400
        
        df = pd.DataFrame(table_data)
        
        if data_type == 'mt':
            sheet_name = 'TS_PW_MT_Analysis'
            filename = f"ts_pw_monthly_budget_tonage_north_{fiscal_year}.xlsx"
            table_title = f"TS-PW Monthly Budget and Actual Tonnage (NORTH) [FY {fiscal_year}]"
        else:
            sheet_name = 'TS_PW_Value_Analysis'
            filename = f"ts_pw_monthly_budget_value_north_{fiscal_year}.xlsx"
            table_title = f"TS-PW Monthly Budget and Actual Value (NORTH) [FY {fiscal_year}]"
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define enhanced formats matching region code
            title_format = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 16, 'font_color': '#000000', 'bg_color': '#D9E1F2',
                'border': 2, 'border_color': '#4472C4'
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
            
            # Data formats with color coding
            budget_format = workbook.add_format({
                'num_format': '#,##0.00', 'border': 1, 'align': 'right',
                
            })
            
            ly_format = workbook.add_format({
                'num_format': '#,##0.00', 'border': 1, 'align': 'right',
                  
            })
            
            actual_format = workbook.add_format({
                'num_format': '#,##0.00', 'border': 1, 'align': 'right',
            
            })
            
            growth_format = workbook.add_format({
                'num_format': '0.00', 'border': 1, 'align': 'right',
        
            })
            
            achievement_format = workbook.add_format({
                'num_format': '0.00', 'border': 1, 'align': 'right',
                
            })
            
            ytd_format = workbook.add_format({
                'num_format': '#,##0.00', 'border': 1, 'align': 'right',
                
            })
            
            # Total row format - enhanced
            total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#C3E6CB', 
                'font_color': '#155724', 'border': 1, 'font_size': 11
            })
            
            text_format = workbook.add_format({
                'border': 1, 'valign': 'vcenter', 'align': 'left'
            })
            
            footer_format = workbook.add_format({
                'font_size': 9, 'italic': True, 'font_color': '#666666'
            })
            
            # Create worksheet
            worksheet = workbook.add_worksheet(sheet_name)
            
            # Write title
            if len(df.columns) > 1:
                worksheet.merge_range(0, 0, 0, len(df.columns) - 1, table_title, title_format)
            else:
                worksheet.write(0, 0, table_title, title_format)
            
            # Write headers with different formats for first column vs data columns
            for col_num, column_name in enumerate(df.columns):
                if col_num == 0:  # Product/first column
                    worksheet.write(3, col_num, column_name, product_header_format)
                    worksheet.set_column(col_num, col_num, 25)  # Wider for product names
                else:
                    worksheet.write(3, col_num, column_name, header_format)
                    worksheet.set_column(col_num, col_num, 12)  # Standard width for data
            
            # Write data with intelligent formatting based on column content
            for row_num, (index, row) in enumerate(df.iterrows()):
                for col_num, column_name in enumerate(df.columns):
                    cell_value = row[column_name]
                    
                    if col_num == 0:  # Product column
                        # Check if this is TOTAL SALES row
                        if str(cell_value).upper() == 'TOTAL SALES':
                            worksheet.write(row_num + 4, col_num, cell_value, 
                                          workbook.add_format({'bold': True, 'bg_color': '#C3E6CB', 
                                                              'font_color': '#155724', 'border': 1}))
                        else:
                            worksheet.write(row_num + 4, col_num, cell_value, text_format)
                    else:  # Data columns
                        # Determine format based on column name and row type
                        is_total_row = str(row[df.columns[0]]).upper() == 'TOTAL SALES'
                        
                        if is_total_row:
                            fmt = total_format
                        else:
                            # Intelligent column formatting based on content
                            col_name_upper = str(column_name).upper()
                            if 'BUDGET-' in col_name_upper and 'YTD' not in col_name_upper:
                                fmt = budget_format
                            elif 'LY-' in col_name_upper and 'YTD' not in col_name_upper:
                                fmt = ly_format
                            elif 'ACT-' in col_name_upper and 'YTD' not in col_name_upper:
                                fmt = actual_format
                            elif 'GR-' in col_name_upper:
                                fmt = growth_format
                            elif 'ACH-' in col_name_upper:
                                fmt = achievement_format
                            elif 'YTD' in col_name_upper:
                                fmt = ytd_format
                            else:
                                fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
                        
                        if isinstance(cell_value, (int, float)):
                            worksheet.write(row_num + 4, col_num, cell_value, fmt)
                        else:
                            worksheet.write(row_num + 4, col_num, cell_value, text_format)
            
            # Add footer information
            footer_row = len(df) + 6
            footer_info = [
                f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"Fiscal Year: {fiscal_year}",
                f"Data Type: {'Tonnage (MT)' if data_type == 'mt' else 'Value (Rs)'}",
                f"Total Records: {len(df)}",
                f"Region: NORTH"
            ]
            
            for i, info_line in enumerate(footer_info):
                worksheet.write(footer_row + i, 0, info_line, footer_format)
            
            # Freeze panes for better navigation
            worksheet.freeze_panes(4, 1)  # Freeze after headers, first column
            
            # Set print settings
            worksheet.set_landscape()
            worksheet.set_paper(9)  # A4
            worksheet.fit_to_pages(1, 0)  # Fit to 1 page wide
        
        excel_data = output.getvalue()
        excel_b64 = base64.b64encode(excel_data).decode('utf-8')
        
        return jsonify({
            'success': True,
            'filename': filename,
            'data': excel_b64
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

@ts_pw_bp.route('/export-merged-excel', methods=['POST'])
def export_merged_excel():
    """Export merged TS-PW data with auditor format using enhanced formatting"""
    try:
        data = request.json
        merged_data = data.get('merged_data')
        data_type = data.get('data_type', 'value')
        fiscal_year = data.get('fiscal_year', '25-26')
        
        if not merged_data:
            return jsonify({
                'success': False,
                'error': 'No merged data provided for export'
            }), 400
        
        df = pd.DataFrame(merged_data)
        
        if data_type == 'mt':
            sheet_name = 'Merged_TS_PW_MT_Analysis'
            filename = f"merged_ts_pw_auditor_tonage_north_{fiscal_year}.xlsx"
            table_title = f"Merged TS-PW Auditor Tonnage Analysis (NORTH) [FY {fiscal_year}]"
        else:
            sheet_name = 'Merged_TS_PW_Value_Analysis'
            filename = f"merged_ts_pw_auditor_value_north_{fiscal_year}.xlsx"
            table_title = f"Merged TS-PW Auditor Value Analysis (NORTH) [FY {fiscal_year}]"
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet(sheet_name)
            
            # Use the same enhanced formats as other functions
            title_format = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 16, 'font_color': '#000000', 'bg_color': '#D9E1F2',
                'border': 2, 'border_color': '#4472C4'
            })
            
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center',
                'fg_color': '#4472C4', 'font_color': 'white', 'border': 1
            })
            
            product_header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center',
                'fg_color': '#2E5F8A', 'font_color': 'white', 'border': 1
            })
            
            total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#C3E6CB', 
                'font_color': '#155724', 'border': 1
            })
            
            # Column-specific formats
            budget_format = workbook.add_format({
                'num_format': '#,##0.00',  'border': 1
            })
            actual_format = workbook.add_format({
                'num_format': '#,##0.00', 'border': 1
            })
            ly_format = workbook.add_format({
                'num_format': '#,##0.00',  'border': 1
            })
            growth_format = workbook.add_format({
                'num_format': '0.00',  'border': 1
            })
            achievement_format = workbook.add_format({
                'num_format': '0.00',  'border': 1
            })
            ytd_format = workbook.add_format({
                'num_format': '#,##0.00',  'border': 1
            })
            
            text_format = workbook.add_format({'border': 1})
            
            # Write title
            worksheet.merge_range(0, 0, 0, len(df.columns) - 1, table_title, title_format)
            
            # Write headers
            for col_num, value in enumerate(df.columns):
                if col_num == 0:
                    worksheet.write(3, col_num, value, product_header_format)
                    worksheet.set_column(col_num, col_num, 25)
                else:
                    worksheet.write(3, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, 12)
            
            # Write data with intelligent formatting
            for row_num in range(len(df)):
                for col_num, col_name in enumerate(df.columns):
                    value = df.iloc[row_num, col_num]
                    
                    if col_num == 0:  # Product column
                        if df.iloc[row_num, 0] == 'TOTAL SALES':
                            worksheet.write(row_num + 4, col_num, value, 
                                          workbook.add_format({'bold': True, 'border': 1, 
                                                              'bg_color': '#C3E6CB', 'font_color': '#155724'}))
                        else:
                            worksheet.write(row_num + 4, col_num, value, text_format)
                    else:  # Data columns
                        is_total = df.iloc[row_num, 0] == 'TOTAL SALES'
                        
                        if is_total:
                            fmt = total_format
                        else:
                            # Intelligent formatting based on column name
                            col_upper = str(col_name).upper()
                            if 'BUDGET-' in col_upper and 'YTD' not in col_upper:
                                fmt = budget_format
                            elif 'ACT-' in col_upper and 'YTD' not in col_upper:
                                fmt = actual_format
                            elif 'LY-' in col_upper and 'YTD' not in col_upper:
                                fmt = ly_format
                            elif 'GR-' in col_upper:
                                fmt = growth_format
                            elif 'ACH-' in col_upper:
                                fmt = achievement_format
                            elif 'YTD' in col_upper:
                                fmt = ytd_format
                            else:
                                fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
                        
                        if isinstance(value, (int, float)):
                            worksheet.write(row_num + 4, col_num, value, fmt)
                        else:
                            worksheet.write(row_num + 4, col_num, value, text_format)
            
            # Add footer
            footer_row = len(df) + 6
            footer_format = workbook.add_format({
                'font_size': 9, 'italic': True, 'font_color': '#666666'
            })
            
            footer_info = [
                f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"Type: Merged TS-PW with Auditor Format",
                f"Data Type: {'Tonnage' if data_type == 'mt' else 'Value'}",
                f"Records: {len(df)}",
                f"Fiscal Year: {fiscal_year}"
            ]
            
            for i, info in enumerate(footer_info):
                worksheet.write(footer_row + i, 0, info, footer_format)
            
            # Freeze panes and print settings
            worksheet.freeze_panes(4, 1)
            worksheet.set_landscape()
            worksheet.set_paper(9)
        
        excel_data = output.getvalue()
        excel_b64 = base64.b64encode(excel_data).decode('utf-8')
        
        return jsonify({
            'success': True,
            'filename': filename,
            'data': excel_b64
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@ts_pw_bp.route('/export-combined-ts-pw-excel', methods=['POST'])
def export_combined_ts_pw_excel():
    """Export both MT and Value data to Excel with enhanced formatting matching region style"""
    try:
        data = request.json
        mt_data = data.get('mt_data', [])
        value_data = data.get('value_data', [])
        mt_columns = data.get('mt_columns', [])
        value_columns = data.get('value_columns', [])
        fiscal_year = data.get('fiscal_year', '25-26')
        
        if not mt_data and not value_data:
            return jsonify({
                'success': False,
                'error': 'No data provided for export'
            }), 400
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"ts_pw_combined_analysis_north_{fiscal_year}_{timestamp}.xlsx"
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Enhanced formats matching region code exactly
            def create_enhanced_formats():
                formats = {}
                
                formats['main_title'] = workbook.add_format({
                    'bold': True, 'align': 'center', 'valign': 'vcenter',
                    'font_size': 18, 'font_color': '#000000', 'bg_color': '#B4C6E7',
                    'border': 2, 'border_color': '#4472C4'
                })
                
                formats['table_title'] = workbook.add_format({
                    'bold': True, 'align': 'center', 'valign': 'vcenter',
                    'font_size': 14, 'font_color': '#1F4E79', 'bg_color': '#D9E1F2',
                    'border': 1
                })
                
                formats['data_header'] = workbook.add_format({
                    'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center',
                    'fg_color': '#4472C4', 'font_color': 'white', 'border': 1,
                    'font_size': 10
                })
                
                formats['product_header'] = workbook.add_format({
                    'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center',
                    'fg_color': '#2E5F8A', 'font_color': 'white', 'border': 1,
                    'font_size': 10
                })
                
                # Column-specific data formats
                formats['budget'] = workbook.add_format({
                    'num_format': '#,##0.00', 'border': 1, 'align': 'right'
                })
                
                formats['ly'] = workbook.add_format({
                    'num_format': '#,##0.00', 'border': 1, 'align': 'right'
                })
                
                formats['actual'] = workbook.add_format({
                    'num_format': '#,##0.00', 'border': 1, 'align': 'right'
                })
                
                formats['growth'] = workbook.add_format({
                    'num_format': '0.00', 'border': 1, 'align': 'right'
                })
                
                formats['achievement'] = workbook.add_format({
                    'num_format': '0.00', 'border': 1, 'align': 'right'
                })
                
                formats['ytd'] = workbook.add_format({
                    'num_format': '#,##0.00', 'border': 1, 'align': 'right'
                })
                
                formats['total'] = workbook.add_format({
                    'bold': True, 'num_format': '#,##0.00', 'bg_color': '#C3E6CB', 
                    'font_color': '#155724', 'border': 1, 'font_size': 11
                })
                
                formats['text'] = workbook.add_format({
                    'border': 1, 'valign': 'vcenter', 'align': 'left'
                })
                
                formats['footer'] = workbook.add_format({
                    'font_size': 9, 'italic': True, 'font_color': '#666666'
                })
                
                return formats
            
            def get_column_format(column_name, formats, is_total_row=False):
                """Get appropriate format based on column name"""
                if is_total_row:
                    return formats['total']
                
                col_name_upper = str(column_name).upper()
                if 'BUDGET-' in col_name_upper and 'YTD' not in col_name_upper:
                    return formats['budget']
                elif 'LY-' in col_name_upper and 'YTD' not in col_name_upper:
                    return formats['ly']
                elif 'ACT-' in col_name_upper and 'YTD' not in col_name_upper:
                    return formats['actual']
                elif 'GR-' in col_name_upper:
                    return formats['growth']
                elif 'ACH-' in col_name_upper:
                    return formats['achievement']
                elif 'YTD' in col_name_upper:
                    return formats['ytd']
                else:
                    return workbook.add_format({'num_format': '#,##0.00', 'border': 1})
            
            def write_table_data(worksheet, data, columns, start_row, table_title, formats):
                """Write table data with enhanced formatting"""
                current_row = start_row
                
                # Table title
                if len(columns) > 1:
                    worksheet.merge_range(current_row, 0, current_row, len(columns) - 1, 
                                        table_title, formats['table_title'])
                else:
                    worksheet.write(current_row, 0, table_title, formats['table_title'])
                current_row += 2
                
                # Headers
                for col_num, column_name in enumerate(columns):
                    if col_num == 0:  # Product column
                        worksheet.write(current_row, col_num, column_name, formats['product_header'])
                        worksheet.set_column(col_num, col_num, 25)
                    else:
                        worksheet.write(current_row, col_num, column_name, formats['data_header'])
                        worksheet.set_column(col_num, col_num, 12)
                current_row += 1
                
                # Data rows
                for row_data in data:
                    for col_num, column_name in enumerate(columns):
                        cell_value = row_data.get(column_name, 0)
                        
                        if col_num == 0:  # Product column
                            if str(cell_value).upper() == 'TOTAL SALES':
                                fmt = workbook.add_format({
                                    'bold': True, 'bg_color': '#C3E6CB', 
                                    'font_color': '#155724', 'border': 1
                                })
                            else:
                                fmt = formats['text']
                            worksheet.write(current_row, col_num, str(cell_value), fmt)
                        else:  # Data columns
                            is_total_row = str(row_data.get(columns[0], '')).upper() == 'TOTAL SALES'
                            fmt = get_column_format(column_name, formats, is_total_row)
                            
                            if isinstance(cell_value, (int, float)):
                                worksheet.write(current_row, col_num, cell_value, fmt)
                            else:
                                worksheet.write(current_row, col_num, cell_value, formats['text'])
                    
                    current_row += 1
                
                return current_row
            
            formats = create_enhanced_formats()
            
            # Single sheet with both tables
            worksheet = workbook.add_worksheet('TS-PW')
            current_row = 0
            
            # Main title
            main_title = f"TS-PW Combined Analysis Report - NORTH Region (FY {fiscal_year})"
            max_cols = max(len(mt_columns) if mt_columns else 0, len(value_columns) if value_columns else 0)
            if max_cols > 1:
                worksheet.merge_range(current_row, 0, current_row, max_cols - 1, main_title, formats['main_title'])
            current_row += 2
            
            # Write MT table
            if mt_data and mt_columns:
                mt_title = f"TS-PW Monthly Budget and Actual Tonnage (NORTH) [FY {fiscal_year}]"
                current_row = write_table_data(worksheet, mt_data, mt_columns, current_row, mt_title, formats)
                current_row += 3  # Add 3 rows spacing between tables
            
            # Write Value table
            if value_data and value_columns:
                value_title = f"TS-PW Monthly Budget and Actual Value (NORTH) [FY {fiscal_year}]"
                current_row = write_table_data(worksheet, value_data, value_columns, current_row, value_title, formats)
            
            # Footer
            current_row += 3
            footer_info = [
                f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"Fiscal Year: {fiscal_year}",
                f"MT Records: {len(mt_data)}",
                f"Value Records: {len(value_data)}",
                f"Region: NORTH",
                f"Analysis Type: Combined TS-PW"
            ]
            
            for i, info_line in enumerate(footer_info):
                worksheet.write(current_row + i, 0, info_line, formats['footer'])
            
            # Freeze panes and print settings
            worksheet.freeze_panes(4, 1)
            worksheet.set_landscape()
            worksheet.set_paper(9)
            worksheet.fit_to_pages(1, 0)
        
        excel_data = output.getvalue()
        excel_b64 = base64.b64encode(excel_data).decode('utf-8')
        
        return jsonify({
            'success': True,
            'filename': filename,
            'data': excel_b64,
            'file_info': {
                'type': 'ts-pw-enhanced',
                'fiscal_year': fiscal_year,
                'mt_records': len(mt_data) if mt_data else 0,
                'value_records': len(value_data) if value_data else 0,
                'enhanced_formatting': True
            }
        })
    
    except Exception as e:
        current_app.logger.error(f"Error in enhanced TS-PW export: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


# NEW ENDPOINT: Store TS-PW Session Data
@ts_pw_bp.route('/store-session-data', methods=['POST'])
def store_tspw_session_data():
    """Store TS-PW analysis data in session for Combined Data Manager"""
    try:
        data = request.json
        session_data = data.get('session_data', {})
        session_id = data.get('session_id', f"tspw_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        
        # Store in a simple in-memory cache or database
        # For now, we'll return success assuming frontend handles storage
        
        return jsonify({
            'success': True,
            'session_id': session_id,
            'message': 'TS-PW session data stored successfully',
            'stored_data': {
                'mt_records': len(session_data.get('mt_data', [])),
                'value_records': len(session_data.get('value_data', [])),
                'fiscal_year': session_data.get('fiscal_year', ''),
                'timestamp': datetime.now().isoformat()
            }
        })
    
    except Exception as e:
        current_app.logger.error(f"Error storing TS-PW session data: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500









































































































































