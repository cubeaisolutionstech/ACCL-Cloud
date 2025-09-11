import pandas as pd
import numpy as np
import re
import logging
import gc
from flask import current_app

# Helper Functions
def allowed_file(filename):
    allowed = current_app.config.get("ALLOWED_EXTENSIONS", {"xlsx", "xls"})
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed

def optimize_memory():
    gc.collect()

def safe_convert_value(x):
    try:
        if x is None or (hasattr(x, 'isna') and pd.isna(x)) or pd.isna(x):
            return None
        str_val = str(x)
        if str_val.lower() in ['nan', 'none', 'nat', '', 'null']:
            return None
        return str_val.strip()
    except:
        return None
    
def extract_tables(df, table1_header, table2_header):
    table1_idx = table2_idx = None
    for i in range(len(df)):
        row_text = ' '.join(str(cell) for cell in df.iloc[i].values if pd.notna(cell)).upper()
        
        # Skip unwanted tables (TS-PW & ERO-PW Monthly Budget tables)
        if "MONTHLY BUDGET AND ACTUAL VALUES" in row_text:
            continue
            
        # Handle table1 header (SALES IN MT / TONNAGE)
        if table1_idx is None:
            if (re.search(r'\bSALES\s*IN\s*MT\b', row_text, re.IGNORECASE) or
                re.search(r'\bSALES\s*IN\s*TON(?:NAGE|AGE)\b', row_text, re.IGNORECASE)):
                table1_idx = i
                logging.debug(f"Table 1 header found at index {i}: {row_text}")
                
        # Handle table2 header (SALES IN VALUE)
        if table2_idx is None and re.search(r'\bSALES\s*IN\s*VALUE\b', row_text, re.IGNORECASE):
            table2_idx = i
            logging.debug(f"Table 2 header found at index {i}: {row_text}")
            
    if table1_idx is None:
        logging.warning("Table 1 header ('SALES IN MT' or 'SALES IN TONNAGE') not found")
    if table2_idx is None:
        logging.warning("Table 2 header 'SALES IN VALUE' not found")
        
    return table1_idx, table2_idx

def rename_columns(columns, is_monthwise=False, is_product=False):
    """Clean column names with special handling for product sheets"""
    renamed = []
    for i, col in enumerate(columns):
        if pd.isna(col) or str(col).strip() == '':
            renamed.append(f'Unnamed_{i}')
            continue
            
        col_str = str(col).strip()
        
        # Skip processing for monthwise sheets
        if is_monthwise:
            renamed.append(col_str)
            continue
            
        # Special handling for product sheets
        if is_product:
            # Clean but preserve original names more for product sheets
            col_str = re.sub(r'\s+', ' ', col_str)  # Normalize spaces
            renamed.append(col_str)
        else:
            # Process YTD columns for other sheets
            renamed.append(clean_ytd_column_name(col_str))
    return renamed

def find_table_end(df, start_idx, is_branch_analysis=False, is_product_analysis=False):
    """Improved table end detection with specific handling for different sheet types"""
    for i in range(start_idx, len(df)):
        row_text = ' '.join(str(cell) for cell in df.iloc[i].values if pd.notna(cell)).upper()
        
        if is_branch_analysis:
            # For branch analysis
            if re.search(r'\b(?:GRAND TOTAL|TOTAL SALES\s*(?:IN\s*(?:MT|TONNAGE|TONAGE)?)?)\b', row_text):
                logging.debug(f"Branch analysis table end found at index {i}: {row_text}")
                return i + 1
        
        elif is_product_analysis:
            # For product analysis
            if re.search(r'\b(?:TOTAL SALES\s*(?:IN\s*(?:MT|TONNAGE|TONAGE)?)?|OVERALL TOTAL|SUMMARY)\b', row_text):
                logging.debug(f"Product analysis table end found at index {i}: {row_text}")
                return i + 1
        
        else:
            # For other sheets
            if re.search(r'\b(?:TOTAL SALES|GRAND TOTAL|OVERALL TOTAL)\b', row_text):
                logging.debug(f"Table end found at index {i}: {row_text}")
                return i + 1
    
    logging.debug(f"No table end marker found, using end of sheet: {len(df)}")
    return len(df)

def clean_ytd_column_name(col_name):
    """
    Standardizes YTD column names while preserving metric type (Budget/LY/Act/Gr/Ach)
    """
    try:
        if pd.isna(col_name):
            return "Unnamed"
            
        col_str = str(col_name).strip().replace("–", "-")  # Normalize dashes
        
        # Match all YTD patterns with metric types
        ytd_match = re.search(
            r'(?P<prefix>Budget|LY|Act|Gr|Ach)?[-\s]*(?P<ytd>YTD[-\s]*(?P<start>\d{2,4})[-\s]*(?P<end>\d{2,4})\s*\((?P<range>.*?)\)',
            col_str,
            re.IGNORECASE
        )
        
        if ytd_match:
            groups = ytd_match.groupdict()
            prefix = groups.get('prefix', '').strip()
            ytd_str = f"YTD-{groups['start']}-{groups['end']}({groups['range']})"
            
            # Only add prefix if it exists and isn't already part of the YTD string
            if prefix and prefix not in ytd_str:
                return f"{prefix} {ytd_str}"
            return ytd_str
            
        return col_str  # Return original if no YTD pattern matched
        
    except Exception as e:
        logging.warning(f"Could not clean YTD column '{col_name}': {str(e)}")
        return str(col_name)
    
def is_monthwise_sheet(sheet_name):
    """Check if the sheet is a Sales Analysis Month-wise sheet"""
    sheet_lower = sheet_name.lower().strip()
    return any(
        sheet_lower == term.lower() 
        for term in [
            "Sales Analysis Month wise",  # Exact match for your sheet name
            "Sales Analysis Month-wise",
            "Month-wise Sales"
        ]
    )    
    
def make_jsonly_serializable(df):
    if df.empty:
        return df
    df = df.copy()
    for col in df.columns:
        try:
            if pd.api.types.is_numeric_dtype(df[col]):
                if pd.api.types.is_integer_dtype(df[col]):
                    df[col] = df[col].astype('Int64').replace([pd.NA, np.nan], None)
                else:
                    df[col] = df[col].astype(float).replace([np.nan, np.inf, -np.inf], None)
            else:
                df[col] = [safe_convert_value(val) for val in df[col]]
        except Exception as e:
            logging.warning(f"Error processing column {col}: {e}")
            df[col] = [str(val) if val is not None and not pd.isna(val) else None for val in df[col]]
    return df.reset_index(drop=True)


def extract_performance_column(df, is_product=False):
    """
    Extracts the performance column from a DataFrame with proper number formatting handling.
    
    Args:
        df (pd.DataFrame): Input dataframe with formatted numbers (with commas)
        is_product (bool): Whether processing product data
    
    Returns:
        str: Name of the best performance column
    """
    def clean_and_convert(value):
        """Convert formatted strings to float (handles commas)"""
        if pd.isna(value):
            return 0.0
        try:
            if isinstance(value, str):
                return float(value.replace(',', ''))
            return float(value)
        except:
            return 0.0
    
    def extract_years(col_name):
        """Extract years from column name"""
        col_str = str(col_name).upper()
        match = re.search(r'YTD[-\s]*(\d{2})[-\s]*(\d{2})', col_str)
        if match:
            return int(match.group(1)), int(match.group(2))
        return (0, 0)
    
    # First look for Act-YTD columns with highest year
    act_ytd_cols = []
    for col in df.columns:
        col_str = str(col).upper()
        if 'ACT' in col_str and 'YTD' in col_str:
            start_year, end_year = extract_years(col)
            # Calculate sum of valid numbers in this column
            col_sum = sum(clean_and_convert(x) for x in df[col])
            act_ytd_cols.append((end_year, start_year, col_sum, col))
    
    if act_ytd_cols:
        # Sort by year (desc) then by sum (desc)
        act_ytd_cols.sort(reverse=True)
        
        # For products, prefer "Apr to Mar" fiscal year
        if is_product:
            for item in act_ytd_cols:
                if '(APR TO MAR)' in str(item[3]).upper():
                    return item[3]
        
        return act_ytd_cols[0][3]  # Return column name with highest year
    
    # Fallback to any YTD column
    ytd_cols = []
    for col in df.columns:
        if 'YTD' in str(col).upper():
            start_year, end_year = extract_years(col)
            col_sum = sum(clean_and_convert(x) for x in df[col])
            ytd_cols.append((end_year, start_year, col_sum, col))
    
    if ytd_cols:
        ytd_cols.sort(reverse=True)
        return ytd_cols[0][3]
    
    # Final fallback - find first numeric column with highest sum
    numeric_cols = []
    for col in df.columns[1:]:  # Skip first column (names)
        try:
            col_sum = sum(clean_and_convert(x) for x in df[col])
            if col_sum > 0:  # Only consider columns with positive values
                numeric_cols.append((col_sum, col))
        except:
            continue
    
    if numeric_cols:
        numeric_cols.sort(reverse=True)
        return numeric_cols[0][1]
    
    return None


def ensure_numeric_data(data, y_col):
    """Ensure column is numeric with comprehensive checks"""
    if y_col not in data.columns:
        logging.error(f"Column {y_col} not found")
        return False
    
    try:
        # First try direct conversion
        data[y_col] = pd.to_numeric(data[y_col], errors='coerce')
        
        # If that fails, try cleaning strings
        if data[y_col].isna().any():
            data[y_col] = pd.to_numeric(
                data[y_col].astype(str)
                .str.replace(r'[^\d.-]', '', regex=True)
                .str.replace(r'\((.*?)\)', '-\\1', regex=True),  # Handle negative values in parentheses
                errors='coerce'
            )
        
        data.dropna(subset=[y_col], inplace=True)
        
        if data.empty:
            logging.warning(f"No valid numeric data in {y_col} after conversion")
            return False
            
        # Verify we have at least some non-zero values
        if (data[y_col] == 0).all():
            logging.warning(f"All values in {y_col} are zero")
            return False
            
        return True
        
    except Exception as e:
        logging.error(f"Failed to convert {y_col} to numeric: {str(e)}")
        return False

def extract_month_year(col_name):
    col_str = str(col_name).strip()
    col_str = re.sub(r'^(Gr[-\s]*|Ach[-\s]*|Act[-\s]*|Budget[-\s]*|LY[-\s]*)', '', col_str, flags=re.IGNORECASE)
    month_year_match = re.search(r'(\w{3,})[-–\s]*(\d{2})', col_str, re.IGNORECASE)
    if month_year_match:
        month, year = month_year_match.groups()
        return f"{month.capitalize()}-{year}"
    return col_str

def safe_sum(series):
    try:
        if hasattr(series, 'sum'):
            result = series.sum()
            return float(result.item() if hasattr(result, 'item') else result)
        return 0.0
    except:
        return 0.0

def safe_mean(series):
    try:
        if hasattr(series, 'mean'):
            result = series.mean()
            return float(result.item() if hasattr(result, 'item') else result)
        return 0.0
    except:
        return 0.0

def convert_to_numeric(series):
    try:
        return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce')
    except:
        return series

def column_filter(col, selected_month=None, selected_year=None):
    """Determine if column should be included based on month/year filters"""
    col_str = str(col).replace(" ", "").lower().replace(",", "").replace("–", "-")
    
    # Handle month filtering
    month_ok = True
    if selected_month and selected_month != "Select All":
        month_ok = any(
            selected_month.lower()[:3] in col_str 
            for month_format in [
                selected_month.lower(),  # Full month name
                selected_month.lower()[:3]  # 3-letter abbreviation
            ]
        )
    
    # Handle year filtering
    year_ok = True
    if selected_year and selected_year != "Select All":
        # Extract last 2 digits of year
        year_short = str(selected_year)[-2:]
        year_ok = year_short in col_str
    
    return month_ok and year_ok

def sort_ytd_periods(periods):
    """Sort YTD periods in fiscal year order (Apr-Mar)"""
    month_order = {'Apr':1, 'May':2, 'Jun':3, 'Jul':4, 'Aug':5, 'Sep':6,
                  'Oct':7, 'Nov':8, 'Dec':9, 'Jan':10, 'Feb':11, 'Mar':12}
    
    def sort_key(period):
        match = re.search(r'\((\w{3}).*?(\w{3})\)', period)
        if match:
            start_month = match.group(1)
            return month_order.get(start_month, 99)
        return 99
    
    return sorted(periods, key=sort_key)