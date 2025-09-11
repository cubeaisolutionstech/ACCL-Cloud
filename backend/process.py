import pandas as pd
import numpy as np
import re
import warnings
from datetime import datetime
from io import BytesIO
import traceback
from fuzzywuzzy import process
import Levenshtein

# Configuration
warnings.filterwarnings('ignore')
pd.set_option("styler.render.max_elements", 500000)

# Memory-efficient merge functions
def safe_merge_dataframes(left_df, right_df, on_column, how='left', max_rows_threshold=100000):
    """
    Safely merge dataframes with memory checks and deduplication to prevent memory explosion.
    
    Args:
        left_df: Left dataframe to merge
        right_df: Right dataframe to merge  
        on_column: Column to merge on
        how: Type of merge ('left', 'right', 'inner', 'outer')
        max_rows_threshold: Maximum expected rows after merge
    
    Returns:
        Tuple: (status_dict, merged_dataframe)
    """
    if left_df is None or right_df is None:
        return {"error": "One of the dataframes is None"}, left_df if right_df is None else right_df
    
    if left_df.empty or right_df.empty:
        return {"warning": "One of the dataframes is empty"}, left_df if right_df.empty else right_df
    
    if on_column not in left_df.columns:
        return {"error": f"Column '{on_column}' not found in left dataframe"}, left_df
    
    if on_column not in right_df.columns:
        return {"error": f"Column '{on_column}' not found in right dataframe"}, left_df

    # Clean and deduplicate the merge column
    left_df = left_df.copy()
    right_df = right_df.copy()
    
    # Ensure merge column is string and cleaned
    left_df[on_column] = left_df[on_column].astype(str).str.strip().str.upper()
    right_df[on_column] = right_df[on_column].astype(str).str.strip().str.upper()
    
    # Remove empty/invalid keys
    left_df = left_df[~left_df[on_column].isin(['', 'NAN', 'NONE', 'NULL'])]
    right_df = right_df[~right_df[on_column].isin(['', 'NAN', 'NONE', 'NULL'])]
    
    # Check for duplicates in merge column and warn
    left_dups = left_df[on_column].duplicated().sum()
    right_dups = right_df[on_column].duplicated().sum()
    
    messages = []
    
    if left_dups > 0:
        # Keep only the first occurrence to prevent explosion
        left_df = left_df.drop_duplicates(subset=[on_column], keep='first')
        messages.append(f"Removed {left_dups} duplicates from left dataframe")
    
    if right_dups > 0:
        # For right dataframe, aggregate numeric columns by sum to prevent loss
        numeric_cols = right_df.select_dtypes(include=[np.number]).columns
        non_numeric_cols = [col for col in right_df.columns if col not in numeric_cols and col != on_column]
        
        if len(numeric_cols) > 0:
            # Aggregate numeric columns by sum, keep first value for non-numeric
            agg_dict = {col: 'sum' for col in numeric_cols}
            if non_numeric_cols:
                agg_dict.update({col: 'first' for col in non_numeric_cols})
            
            right_df = right_df.groupby(on_column).agg(agg_dict).reset_index()
        else:
            right_df = right_df.drop_duplicates(subset=[on_column], keep='first')
        
        messages.append(f"Processed {right_dups} duplicates in right dataframe")
    
    # Estimate merge result size
    estimated_rows = len(left_df) * right_df.groupby(on_column).size().max() if how == 'left' else len(left_df) + len(right_df)
    estimated_cols = len(left_df.columns) + len(right_df.columns) - 1  # -1 for merge column overlap
    
    if estimated_rows > max_rows_threshold:
        return {"error": f"Estimated merge size ({estimated_rows} rows) exceeds threshold ({max_rows_threshold}). Aborting to prevent memory issues."}, left_df
    
    # Perform the merge with error handling
    try:
        # Use categorical data type for merge column to save memory
        if left_df[on_column].nunique() < len(left_df) * 0.5:  # If less than 50% unique values
            left_df[on_column] = left_df[on_column].astype('category')
            right_df[on_column] = right_df[on_column].astype('category')
        
        merged_df = left_df.merge(right_df, on=on_column, how=how, suffixes=('', '_right'))
        
        # Remove duplicate columns that may have been created
        duplicate_cols = [col for col in merged_df.columns if col.endswith('_right')]
        if duplicate_cols:
            merged_df = merged_df.drop(columns=duplicate_cols)
            messages.append("Removed duplicate columns after merge")
        
        return {"success": True, "messages": messages}, merged_df
        
    except MemoryError as e:
        return {"error": f"Memory error during merge: {str(e)}"}, left_df
        
    except Exception as e:
        return {"error": f"Error during merge: {str(e)}"}, left_df

def chunk_based_merge(left_df, right_df, on_column, how='left', chunk_size=10000):
    """
    Perform merge in chunks to handle large datasets with limited memory.
    """
    try:
        merged_chunks = []
        total_chunks = (len(left_df) // chunk_size) + 1
        
        for i in range(0, len(left_df), chunk_size):
            chunk = left_df.iloc[i:i+chunk_size]
            merged_chunk = chunk.merge(right_df, on=on_column, how=how, suffixes=('', '_right'))
            
            # Remove duplicate columns
            duplicate_cols = [col for col in merged_chunk.columns if col.endswith('_right')]
            if duplicate_cols:
                merged_chunk = merged_chunk.drop(columns=duplicate_cols)
            
            merged_chunks.append(merged_chunk)
        
        # Combine all chunks
        final_merged = pd.concat(merged_chunks, ignore_index=True)
        return {"success": True, "message": f"Chunk-based merge completed! Final shape: {final_merged.shape}"}, final_merged
        
    except Exception as e:
        return {"error": f"Chunk-based merge also failed: {str(e)}"}, left_df

def optimize_dataframe_memory(df):
    """
    Optimize dataframe memory usage by converting data types.
    """
    if df is None or df.empty:
        return df
    
    original_memory = df.memory_usage(deep=True).sum() / 1024**2  # MB
    
    # Optimize numeric columns
    for col in df.select_dtypes(include=['int64']).columns:
        if df[col].min() >= -32768 and df[col].max() <= 32767:
            df[col] = df[col].astype('int16')
        elif df[col].min() >= -2147483648 and df[col].max() <= 2147483647:
            df[col] = df[col].astype('int32')
    
    for col in df.select_dtypes(include=['float64']).columns:
        if df[col].min() >= -3.4e38 and df[col].max() <= 3.4e38:
            df[col] = df[col].astype('float32')
    
    # Optimize string columns with low cardinality
    for col in df.select_dtypes(include=['object']).columns:
        if df[col].nunique() / len(df) < 0.5:  # Less than 50% unique values
            df[col] = df[col].astype('category')
    
    new_memory = df.memory_usage(deep=True).sum() / 1024**2  # MB
    
    memory_info = {
        'original_memory': original_memory,
        'new_memory': new_memory,
        'reduction_percentage': (1 - new_memory/original_memory) * 100 if original_memory > 0 else 0
    }
    
    return df, memory_info

# Helper Functions
def handle_duplicate_columns(df):
    """Handle duplicate column names in dataframe"""
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols[cols == dup].index.values.tolist()] = [
            f"{dup}_{i}" if i != 0 else dup 
            for i in range(sum(cols == dup))
        ]
    df.columns = cols
    return df

def clean_header_row(header_row):
    """Clean header row by removing None values and empty strings"""
    return [
        str(col).strip() if pd.notna(col) and str(col).strip() != '' else None
        for col in header_row
    ]

def extract_tables(df, possible_headers, is_product_analysis=False):
    """Extract tables from dataframe based on possible headers"""
    for i in range(len(df)):
        row_text = ' '.join(df.iloc[i].astype(str).str.lower().tolist())
        for header in possible_headers:
            if header.lower() in row_text:
                # First check for budget/actual style headers
                potential_header = df.iloc[i]
                if any(str(col).strip().lower().startswith(('budget-', 'act-', 'ly-', 'gr.', 'ach.')) for col in potential_header[1:]):
                    data_start = i + 2 if i + 2 < len(df) else i + 1
                    if data_start < len(df):
                        first_col = str(df.iloc[data_start, 0]).strip().upper()
                        identifier_cols = ['REGIONS', 'REGION', 'BRANCH', 'ORGANIZATION', 'ORGANIZATION NAME'] if not is_product_analysis else ['PRODUCT', 'PRODUCT GROUP', 'PRODUCT NAME', 'ACETIC ACID', 'AUXILARIES', 'CSF', 'TOTAL']
                        if any(r in first_col for r in identifier_cols) or first_col in ['ACCLLP', 'TOTAL SALES']:
                            return i, data_start
                    
                    data_start = i + 1
                    if data_start < len(df):
                        first_col = str(df.iloc[data_start, 0]).strip().upper()
                        if any(r in first_col for r in identifier_cols) or first_col in ['ACCLLP', 'TOTAL SALES']:
                            return i, data_start
                
                # Check next row for budget/actual style headers
                if i + 1 < len(df):
                    potential_header = df.iloc[i + 1]
                    if any(str(col).strip().lower().startswith(('budget-', 'act-', 'ly-', 'gr.', 'ach.')) for col in potential_header[1:]):
                        data_start = i + 2 if i + 2 < len(df) else i + 1
                        if data_start < len(df):
                            first_col = str(df.iloc[data_start, 0]).strip().upper()
                            identifier_cols = ['REGIONS', 'REGION', 'BRANCH', 'ORGANIZATION', 'ORGANIZATION NAME'] if not is_product_analysis else ['PRODUCT', 'PRODUCT GROUP', 'PRODUCT NAME', 'ACETIC ACID', 'AUXILARIES', 'CSF', 'TOTAL']
                            if any(r in first_col for r in identifier_cols) or first_col in ['ACCLLP', 'TOTAL SALES']:
                                return i + 1, data_start
                
                # If budget/actual style not found, try the first approach
                for j in range(i + 1, min(i + 5, len(df))):
                    row = df.iloc[j]
                    first_col = str(row.iloc[0]).strip().upper()
                    identifier_cols = ['REGIONS', 'REGION', 'BRANCH'] if not is_product_analysis else ['PRODUCT', 'PRODUCT GROUP', 'PRODUCT NAME', 'ACETIC ACID', 'AUXILARIES', 'CSF', 'TOTAL']
                    if any(r in first_col for r in identifier_cols):
                        header_row = j - 1 if j > 0 else j
                        potential_header = df.iloc[header_row]
                        if not all(str(col).strip().upper() in ['MT', 'RS', ''] for col in potential_header[1:]):
                            return header_row, j
                        else:
                            header_row = j - 2 if j > 1 else j
                            potential_header = df.iloc[header_row]
                            if not all(str(col).strip().upper() in ['MT', 'RS', ''] for col in potential_header[1:]):
                                return header_row, j
                            else:
                                return None, None
                
                return None, None
    
    return None, None

def find_column(df, possible_names, case_sensitive=False, threshold=80):
    """Enhanced fuzzy matching for column names"""
    if isinstance(possible_names, str):
        possible_names = [possible_names]
    
    for name in possible_names:
        if case_sensitive:
            if name in df.columns:
                return name
        else:
            for col in df.columns:
                if col.lower() == name.lower():
                    return col
    
    # If exact match not found, try fuzzy matching
    for name in possible_names:
        matches = process.extractOne(name, df.columns, score_cutoff=threshold)
        if matches:
            return matches[0]
    
    return None

def standardize_column_names(df, is_auditor=False):
    """Standardize column names for consistency"""
    df = df.copy()
    original_columns = df.columns.tolist()
    new_columns = []
    last_act_column = None
    column_mapping = {}

    for i, col in enumerate(df.columns):
        col = str(col).upper().strip()
        if is_auditor:
            column_mapping[col] = original_columns[i]

        if any(x in col for x in ['PRODUCT', 'REGION', 'BRANCH']):
            new_columns.append('PRODUCT NAME' if 'PRODUCT' in col else 'REGIONS')
            continue

        if col in ['GR.', 'ACH.']:
            if last_act_column:
                prefix = 'Gr' if col == 'GR.' else 'Ach'
                if 'YTD' in last_act_column:
                    ytd_base = last_act_column.replace('ACT-', '').replace('ACT.', '').strip()
                    new_columns.append(f'{prefix}-{ytd_base}')
                else:
                    month_year = last_act_column.replace('ACT-', '')
                    new_columns.append(f'{prefix}-{month_year}')
            else:
                new_columns.append(col)
            continue

        col_clean = col.replace('BUDGET', '').replace('ACTUAL', '').replace('ACT', '').replace('LY', '').strip()
        month_year = None
        for fmt in ['%b-%y', '%B-%y', '%b-%Y', '%B-%Y']:
            try:
                parsed_date = datetime.strptime(col_clean, fmt)
                month_year = parsed_date.strftime('%b-%y')
                break
            except ValueError:
                continue

        if month_year:
            if 'BUDGET' in col:
                new_columns.append(f'Budget-{month_year}')
            elif 'ACT' in col or 'ACTUAL' in col:
                new_columns.append(f'Act-{month_year}')
                last_act_column = f'Act-{month_year}'
            elif 'LY' in col:
                new_columns.append(f'LY-{month_year}')
            else:
                new_columns.append(col)
        else:
            ytd_match = re.match(r'(?:ACT[- ]*)?YTD[-–\s]*(\d{2})[-–\s]*(\d{2})\s*\((.*?)\)', col, re.IGNORECASE)
            if ytd_match:
                start_year, end_year, period = ytd_match.groups()
                period = period.replace('June', 'Jun').replace('July', 'Jul').replace('August', 'Aug').replace('September', 'Sep').replace('October', 'Oct').replace('November', 'Nov').replace('December', 'Dec')
                ytd_base = f'YTD-{start_year}-{end_year} ({period})'
                if 'ACT' in col:
                    new_columns.append(f'Act-{ytd_base}')
                    last_act_column = f'Act-{ytd_base}'
                elif 'L,Y' in col:
                    new_columns.append(f'{ytd_base}L,Y')
                else:
                    new_columns.append(ytd_base)
            else:
                new_columns.append(col)

    df.columns = new_columns
    return df, column_mapping

def clean_and_convert_numeric(df):
    """
    Clean and convert DataFrame columns to appropriate data types.
    Handles mixed data types that cause serialization issues.
    """
    if df is None or df.empty:
        return df
        
    df = df.copy()
    
    # Get the first column (identifier column)
    identifier_col = df.columns[0]
    
    # Process each column
    for col in df.columns:
        if col == identifier_col:
            # Keep identifier column as string, ensure no numeric conversion
            df[col] = df[col].astype(str).str.strip()
            # Replace common problematic values
            df[col] = df[col].replace({'nan': '', 'NaN': '', 'None': '', 'null': ''})
            # Don't convert to numeric for identifier columns
            continue
            
        # Try to convert to numeric for other columns
        try:
            # First, convert to string and clean
            series_str = df[col].astype(str).str.strip()
            
            # Replace common non-numeric values
            series_str = series_str.replace({
                'nan': '0',
                'NaN': '0', 
                'None': '0',
                '': '0',
                '-': '0',
                'null': '0'
            })
            
            # Try to convert to numeric
            df[col] = pd.to_numeric(series_str, errors='coerce')
            
            # Fill any remaining NaN values with 0
            df[col] = df[col].fillna(0.0)
            
            # Ensure it's float64 for consistency
            df[col] = df[col].astype('float64')
            
        except Exception as e:
            # If conversion fails, keep as string but clean it
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({'nan': '', 'NaN': '', 'None': '', 'null': ''})
    
    return df

def validate_dataframe(df, name, required_cols=['REGIONS', 'PRODUCT NAME']):
    """Validate dataframe structure and content"""
    if df is None or df.empty:
        return False, f"{name} is None or empty."
    
    present_cols = [col for col in required_cols if col in df.columns]
    if not present_cols:
        return False, f"None of the required columns {', '.join(required_cols)} found in {name}."
    
    identifier_col = present_cols[0]
    if df[identifier_col].isna().all():
        return False, f"All '{identifier_col}' values in {name} are NaN."
    
    return True, "Validation passed"

def process_budget_data(budget_df, group_type='region'):
    """Process budget data based on group type"""
    try:
        budget_df = handle_duplicate_columns(budget_df.copy())
        budget_df.columns = budget_df.columns.str.strip()
        
        identifier_col = None
        identifier_names = ['Branch', 'Region', 'REGIONS'] if group_type == 'region' else ['Product', 'Product Group', 'PRODUCT NAME']
        for col in identifier_names:
            if col in budget_df.columns:
                identifier_col = col
                break
        
        if not identifier_col:
            identifier_col = find_column(budget_df, identifier_names[0], threshold=80)
            if not identifier_col:
                return None, f"Could not find {group_type.capitalize()} column in budget dataset."
        
        budget_cols = {'Qty': [], 'Value': []}
        detailed_pattern = r'(Qty|Value)\s*[-]\s*(\w{3,})\'?(\d{2,4})'
        range_pattern = r'(Qty|Value)\s*(\w{3,})\'?(\d{2,4})[-]\s*(\w{3,})\'?(\d{2,4})'
        
        for col in budget_df.columns:
            col_clean = col.lower().replace("'", "").replace(" ", "").replace("-", "")
            
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
            return None, f"No budget quantity or value columns found: Qty={budget_cols['Qty']}, Value={budget_cols['Value']}"
        
        for col, _ in budget_cols['Qty'] + budget_cols['Value']:
            budget_df[col] = pd.to_numeric(budget_df[col], errors='coerce')
        
        group_cols = [col for col, _ in budget_cols['Qty'] + budget_cols['Value']]
        budget_data = budget_df.groupby(identifier_col)[group_cols].sum().reset_index()
        
        rename_dict = {identifier_col: 'REGIONS' if group_type == 'region' else 'PRODUCT NAME'}
        for col, month_year in budget_cols['Qty']:
            rename_dict[col] = f'Budget-{month_year}_MT'
        for col, month_year in budget_cols['Value']:
            rename_dict[col] = f'Budget-{month_year}_Value'
        
        budget_data = budget_data.rename(columns=rename_dict)
        budget_data[rename_dict[identifier_col]] = budget_data[rename_dict[identifier_col]].str.strip().str.upper()
        
        return budget_data, "Budget data processed successfully"
        
    except Exception as e:
        return None, f"Error processing budget data: {str(e)}"

def process_last_year_data(last_year_df, group_type='region'):
    """Process last year data based on group type"""
    try:
        last_year_df = handle_duplicate_columns(last_year_df.copy())
        last_year_df.columns = last_year_df.columns.str.strip()
        
        identifier_col = None
        identifier_names = ['Branch', 'Region', 'REGIONS'] if group_type == 'region' else ['Product', 'Product Group', 'PRODUCT NAME']
        for col in identifier_names:
            if col in last_year_df.columns:
                identifier_col = col
                break
        
        if not identifier_col:
            identifier_col = find_column(last_year_df, identifier_names[0], threshold=80)
            if not identifier_col:
                return None, f"Could not find {group_type.capitalize()} column in last year dataset."
        
        ly_cols = {'Qty': [], 'Value': []}
        detailed_pattern = r'(Qty|Value)\s*[-]\s*(\w{3,})\'?(\d{2,4})'
        range_pattern = r'(Qty|Value)\s*(\w{3,})\'?(\d{2,4})[-]\s*(\w{3,})\'?(\d{2,4})'
        
        for col in last_year_df.columns:
            col_clean = col.lower().replace("'", "").replace(" ", "").replace("-", "")
            
            detailed_match = re.match(detailed_pattern, col, re.IGNORECASE)
            if detailed_match:
                qty_or_value, month, year = detailed_match.groups()
                month = month.capitalize()
                year = year[-2:] if len(year) > 2 else year
                month_year = f"{month}-{year}"
                if qty_or_value.lower() == 'qty':
                    ly_cols['Qty'].append((col, month_year))
                elif qty_or_value.lower() == 'value':
                    ly_cols['Value'].append((col, month_year))
                continue
            
            range_match = re.match(range_pattern, col, re.IGNORECASE)
            if range_match:
                qty_or_value, start_month, start_year, end_month, end_year = range_match.groups()
                start_month = start_month.capitalize()
                start_year = start_year[-2:] if len(start_year) > 2 else start_year
                end_year = end_year[-2:] if len(end_year) > 2 else end_year
                month_year = f"{start_month}{start_year}{end_month.lower()}-{end_year}"
                if qty_or_value.lower() == 'qty':
                    ly_cols['Qty'].append((col, month_year))
                elif qty_or_value.lower() == 'value':
                    ly_cols['Value'].append((col, month_year))
        
        if not ly_cols['Qty'] and not ly_cols['Value']:
            return None, f"No last year quantity or value columns found: Qty={ly_cols['Qty']}, Value={ly_cols['Value']}"
        
        for col, _ in ly_cols['Qty'] + ly_cols['Value']:
            last_year_df[col] = pd.to_numeric(last_year_df[col], errors='coerce')
        
        group_cols = [col for col, _ in ly_cols['Qty'] + ly_cols['Value']]
        last_year_data = last_year_df.groupby(identifier_col)[group_cols].sum().reset_index()
        
        rename_dict = {identifier_col: 'REGIONS' if group_type == 'region' else 'PRODUCT NAME'}
        for col, month_year in ly_cols['Qty']:
            rename_dict[col] = f'LY-{month_year}_MT'
        for col, month_year in ly_cols['Value']:
            rename_dict[col] = f'LY-{month_year}_Value'
        
        last_year_data = last_year_data.rename(columns=rename_dict)
        last_year_data[rename_dict[identifier_col]] = last_year_data[rename_dict[identifier_col]].str.strip().str.upper()
        
        return last_year_data, "Last year data processed successfully"
        
    except Exception as e:
        return None, f"Error processing last year data: {str(e)}"