from flask import Blueprint, request, jsonify, send_file
import pandas as pd
import numpy as np
import re
from datetime import datetime
from io import BytesIO
from fuzzywuzzy import process
from werkzeug.utils import secure_filename
import os
import xlsxwriter

region_bp = Blueprint('region', __name__)

# Enhanced column aliases for comprehensive branch identification
column_aliases = {
    'Date': ['Month Format', 'Month', 'Date', 'Period', 'Time Period', 'Month Year'],
    'Product Name': ['Type(Make)', 'Product Group'],
    'Value': ['Value', 'Amount'],
    'Amount': ['Amount', 'Value'],
    'Branch': [
        'Branch', 'Branch.1', 'Branch',
        'Region', 'REGIONS', 'Region',
        
    ],
    'Actual Quantity': [
        'Actual Quantity', 'Acutal Quantity', 'Quantity', 
        
    ]
}

def create_branch_name_mappings():
    """Create comprehensive branch name mappings for standardization"""
    
    branch_mappings = {
        # Bangalore variations
        'BANGALORE': 'BGLR',
        'BENGALURU': 'BGLR', 
        'BANG': 'BGLR',
        'BLR': 'BGLR',
        'BGLR': 'BGLR',
        
        # Chennai variations
        'CHENNAI': 'CHENNAI',
        'MADRAS': 'CHENNAI',
        'CHE': 'CHENNAI',
        'CHEN': 'CHENNAI',
        'MAS': 'CHENNAI',
        
        # Pondicherry variations
        'PONDICHERRY': 'PONDY',
        'PUDUCHERRY': 'PONDY',
        'PONDY': 'PONDY',
        'PON': 'PONDY',
        'PDY': 'PONDY',
        
        # Coimbatore variations
        'COIMBATORE': 'COVAI',
        'COVAI': 'COVAI',
        'COI': 'COVAI',
        'CBE': 'COVAI',
        'KOVAI': 'COVAI',
        
        # Erode variations
        'ERODE': 'ERODE',
        'ERD': 'ERODE',
        'ERODE CITY': 'ERODE',
        
        # Madurai variations
        'MADURAI': 'MADURAI',
        'MDU': 'MADURAI',
        'MAD': 'MADURAI',
        
        # Poultry variations
        'POULTRY': 'POULTRY',
        'POULTRY DIVISION': 'POULTRY',
        'POULTRY DEPT': 'POULTRY',
        'POULTRY UNIT': 'POULTRY',
        'POUL': 'POULTRY',
        
        # Karur variations
        'KARUR': 'KARUR',
        'KAR': 'KARUR',
        'KARUR CITY': 'KARUR',
        
        # Salem variations
        'SALEM': 'SALEM',
        'SAL': 'SALEM',
        'SALEM CITY': 'SALEM',
        
        # Tirupur variations
        'TIRUPUR': 'TIRUPUR',
        'TIRRUPUR': 'TIRUPUR',
        'TUP': 'TIRUPUR',
        'TPR': 'TIRUPUR',
        'TIRUPPUR': 'TIRUPUR',
        
        # Group company variations
        'GROUP': 'GROUP',
        'GROUP COMPANY': 'GROUP',
        'GROUP COMPANIES': 'GROUP',
        'HEAD OFFICE': 'GROUP',
        'HO': 'GROUP',
        'CORPORATE': 'GROUP',
        'GROUP SALES': 'GROUP'
    }
    
    return branch_mappings

def normalize_branch_name(branch_name, branch_mappings=None):
    """
    Normalize branch name using various mapping strategies
    """
    if pd.isna(branch_name) or branch_name == '':
        return ''
    
    # Convert to string and clean
    clean_name = str(branch_name).strip().upper()
    
    # Remove common prefixes/suffixes
    prefixes_to_remove = ['BRANCH', 'REGION', 'OFFICE', 'DEPOT', 'UNIT', 'DIVISION', 'CENTRE', 'CENTER']
    suffixes_to_remove = ['BRANCH', 'REGION', 'OFFICE', 'DEPOT', 'UNIT', 'DIVISION', 'CENTRE', 'CENTER']
    
    for prefix in prefixes_to_remove:
        if clean_name.startswith(prefix + ' '):
            clean_name = clean_name[len(prefix):].strip()
    
    for suffix in suffixes_to_remove:
        if clean_name.endswith(' ' + suffix):
            clean_name = clean_name[:-len(suffix)].strip()
    
    # Use mappings if provided, otherwise use default
    if branch_mappings is None:
        branch_mappings = create_branch_name_mappings()
    
    # Direct mapping
    if clean_name in branch_mappings:
        return branch_mappings[clean_name]
    
    # Fuzzy matching for close matches
    best_match = process.extractOne(clean_name, list(branch_mappings.keys()), score_cutoff=85)
    if best_match:
        return branch_mappings[best_match[0]]
    
    # Return cleaned name if no mapping found
    return clean_name

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
    
    for name in possible_names:
        matches = process.extractOne(name, df.columns, score_cutoff=threshold)
        if matches:
            return matches[0]
    
    return None

def enhanced_find_column(df, possible_names, case_sensitive=False, threshold=80, exact_match_priority=True):
    """
    Enhanced fuzzy matching for column names with priority for exact matches
    """
    if isinstance(possible_names, str):
        possible_names = [possible_names]
    
    # First try exact matches
    if exact_match_priority:
        for name in possible_names:
            if case_sensitive:
                if name in df.columns:
                    return name
            else:
                for col in df.columns:
                    if col.lower() == name.lower():
                        return col
    
    # Then try fuzzy matching
    for name in possible_names:
        matches = process.extractOne(name, df.columns, score_cutoff=threshold)
        if matches:
            return matches[0]
    
    return None

def handle_duplicate_columns(df):
    """Handle duplicate column names by renaming them"""
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols
    return df

def detect_budget_sheet_header_row(filepath, sheet_name):
    """
    Detect the correct header row for budget sheet by checking for expected columns
    Returns the appropriate header row number (0 or 1)
    """
    try:
        # Expected column patterns for budget data
        expected_patterns = [
            # Region identifiers
            ['Branch', 'Region', 'REGIONS'],
            # Budget column patterns
            ['Qty', 'Value', 'Budget'],
            # Month patterns
            ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
        ]
        
        # Try header=0 first
        try:
            df_test_0 = pd.read_excel(filepath, sheet_name=sheet_name, header=0, nrows=5)
            df_test_0.columns = df_test_0.columns.str.strip()
            
            # Check if we can find expected columns in row 0
            found_patterns_0 = 0
            for pattern_group in expected_patterns:
                for pattern in pattern_group:
                    if find_column(df_test_0, pattern, threshold=70):
                        found_patterns_0 += 1
                        break
            
            # If we found at least 2 pattern groups, row 0 is likely correct
            if found_patterns_0 >= 2:
                print(f"Budget sheet: Using header row 0 (found {found_patterns_0} pattern matches)")
                return 0
                
        except Exception as e:
            print(f"Error testing header=0: {e}")
        
        # Try header=1 if row 0 didn't work well
        try:
            df_test_1 = pd.read_excel(filepath, sheet_name=sheet_name, header=1, nrows=5)
            df_test_1.columns = df_test_1.columns.str.strip()
            
            # Check if we can find expected columns in row 1
            found_patterns_1 = 0
            for pattern_group in expected_patterns:
                for pattern in pattern_group:
                    if find_column(df_test_1, pattern, threshold=70):
                        found_patterns_1 += 1
                        break
            
            print(f"Budget sheet: Comparing row 0 ({found_patterns_0} matches) vs row 1 ({found_patterns_1} matches)")
            
            # Use row 1 if it has more matches, otherwise default to row 0
            if found_patterns_1 > found_patterns_0:
                print(f"Budget sheet: Using header row 1 (found {found_patterns_1} pattern matches)")
                return 1
                
        except Exception as e:
            print(f"Error testing header=1: {e}")
        
        # Default to row 0 if both fail or row 0 is equal/better
        print("Budget sheet: Defaulting to header row 0")
        return 0
        
    except Exception as e:
        print(f"Error in header detection: {e}")
        return 0

def extract_tables_from_auditor(df, headers):
    """Extract table data from auditor format"""
    table_idx = None
    data_start = None
    
    for idx, row in df.iterrows():
        row_str = ' '.join([str(cell).strip().upper() for cell in row if pd.notna(cell)])
        for header in headers:
            if header.upper() in row_str:
                table_idx = idx
                data_start = idx + 1
                break
        if table_idx is not None:
            break
    
    return table_idx, data_start

def rename_columns(columns):
    """Rename columns to standard format"""
    new_columns = []
    for col in columns:
        if pd.isna(col):
            new_columns.append('Unnamed')
        else:
            new_columns.append(str(col).strip())
    return new_columns

def process_budget_data(budget_df, group_type='region'):
    """Process budget data for region analysis"""
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
            return None
    
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
        return None
    
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
    
    # Enhanced branch name normalization for budget data
    budget_data[rename_dict[identifier_col]] = budget_data[rename_dict[identifier_col]].apply(
        lambda x: normalize_branch_name(x) if pd.notna(x) else ''
    )
    
    return budget_data

def add_regional_totals(df, data_type='MT', fiscal_year_start=None, fiscal_year_end=None, 
                       last_fiscal_year_start=None, last_fiscal_year_end=None):
    """Add regional totals with all required fiscal year parameters"""
    if df.empty:
        return df
    
    # Define regional classifications
    north_regions = ['BGLR', 'CHENNAI', 'PONDY']
    west_regions = ['COVAI', 'ERODE', 'MADURAI', 'POULTRY', 'KARUR', 'SALEM', 'TIRUPUR']
    group_companies = ['GROUP']
    
    id_col = 'SALES in MT' if data_type == 'MT' else 'SALES in Value'
    
    # Remove existing totals
    df_clean = df[~df[id_col].isin(['NORTH TOTAL', 'WEST SALES', 'GROUP COMPANIES', 'GRAND TOTAL'])].copy()
    
    def calculate_summed_totals(group_data, group_name):
        """Sums ALL values including Gr and Ach percentages"""
        if group_data.empty:
            return None
            
        total_row = {id_col: group_name}
        
        # Sum ALL numeric columns (including percentages)
        numeric_cols = group_data.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            total_row[col] = group_data[col].sum()
        
        return total_row
    
    # Calculate totals
    totals = [
        ('NORTH TOTAL', north_regions),
        ('WEST SALES', west_regions),
        ('GROUP COMPANIES', group_companies),
        ('GRAND TOTAL', None)  # All remaining regions
    ]
    
    for total_name, regions in totals:
        if regions:
            group_data = df_clean[df_clean[id_col].isin(regions)]
        else:
            group_data = df_clean[~df_clean[id_col].isin(['NORTH TOTAL', 'WEST SALES', 'GROUP COMPANIES'])]
        
        total_row = calculate_summed_totals(group_data, total_name)
        if total_row:
            df_clean = pd.concat([df_clean, pd.DataFrame([total_row])], ignore_index=True)
    
    # Maintain original ordering
    result_df = pd.concat([
        df_clean[df_clean[id_col].isin(north_regions)].sort_values(id_col),
        df_clean[df_clean[id_col] == 'NORTH TOTAL'],
        df_clean[df_clean[id_col].isin(west_regions)].sort_values(id_col),
        df_clean[df_clean[id_col] == 'WEST SALES'],
        df_clean[~df_clean[id_col].isin(north_regions + west_regions + group_companies + 
                                      ['NORTH TOTAL', 'WEST SALES', 'GROUP COMPANIES', 'GRAND TOTAL'])],
        df_clean[df_clean[id_col] == 'GROUP COMPANIES'],
        df_clean[df_clean[id_col] == 'GRAND TOTAL']
    ], ignore_index=True)
    
    return result_df

def add_ytd_calculations_auditor_format(df, data_type='MT', fiscal_year_start=None, fiscal_year_end=None,
                                      last_fiscal_year_start=None, last_fiscal_year_end=None):
    """Add YTD calculations in auditor format with proper column ordering"""
    if df.empty:
        return df
    
    id_col = 'SALES in MT' if data_type == 'MT' else 'SALES in Value'
    months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
    
    # Define YTD periods for auditor format
    ytd_periods = [
        ('Apr', 'Jun', 'Apr to Jun'),    # Q1
        ('Apr', 'Sep', 'Apr to Sep'),    # H1  
        ('Apr', 'Dec', 'Apr to Dec'),    # 9M
        ('Apr', 'Mar', 'Apr to Mar')     # Full Year
    ]
    
    # Define regional groupings
    north_regions = ['BGLR', 'CHENNAI', 'PONDY']
    west_regions = ['COVAI', 'ERODE', 'MADURAI', 'POULTRY', 'KARUR', 'SALEM', 'TIRUPUR']
    group_companies = ['GROUP']
    
    # Create ordered column list starting with region column
    ordered_columns = [id_col]
    
    # Process each month and add YTD after quarters
    for i, month in enumerate(months):
        # Determine fiscal year for this month
        budget_year = str(fiscal_year_start)[-2:] if month in months[:9] else str(fiscal_year_end)[-2:]
        actual_year = str(fiscal_year_start)[-2:] if month in months[:9] else str(fiscal_year_end)[-2:]
        ly_year = str(last_fiscal_year_start)[-2:] if month in months[:9] else str(last_fiscal_year_end)[-2:]
        
        # Add individual month columns
        budget_col = f'Budget-{month}-{budget_year}'
        ly_col = f'LY-{month}-{ly_year}'
        actual_col = f'Act-{month}-{actual_year}'
        gr_col = f'Gr-{month}-{actual_year}'
        ach_col = f'Ach-{month}-{actual_year}'
        
        # Ensure columns exist in dataframe
        for col in [budget_col, ly_col, actual_col, gr_col, ach_col]:
            if col not in df.columns:
                df[col] = 0.0
        
        # Add to ordered columns
        ordered_columns.extend([budget_col, ly_col, actual_col, gr_col, ach_col])
        
        # Check if we need to add YTD after this month
        for start_month, end_month, period_name in ytd_periods:
            if month == end_month:
                # Calculate YTD for this period
                ytd_budget_col = f'YTD-{str(fiscal_year_start)[-2:]}-{str(fiscal_year_end)[-2:]} ({period_name})Budget'
                ytd_ly_col = f'YTD-{str(last_fiscal_year_start)[-2:]}-{str(last_fiscal_year_end)[-2:]} ({period_name})LY'
                ytd_actual_col = f'Act-YTD-{str(fiscal_year_start)[-2:]}-{str(fiscal_year_end)[-2:]} ({period_name})'
                ytd_gr_col = f'Gr-YTD-{str(fiscal_year_start)[-2:]}-{str(fiscal_year_end)[-2:]} ({period_name})'
                ytd_ach_col = f'Ach-YTD-{str(fiscal_year_start)[-2:]}-{str(fiscal_year_end)[-2:]} ({period_name})'
                
                # Initialize YTD columns
                df[ytd_budget_col] = 0.0
                df[ytd_ly_col] = 0.0
                df[ytd_actual_col] = 0.0
                df[ytd_gr_col] = 0.0
                df[ytd_ach_col] = 0.0
                
                # Calculate YTD values for all regions first
                start_idx = months.index(start_month)
                end_idx = months.index(end_month) + 1
                ytd_months = months[start_idx:end_idx]
                
                for _, row in df.iterrows():
                    row_idx = row.name
                    ytd_budget_sum = 0
                    ytd_ly_sum = 0
                    ytd_actual_sum = 0
                    
                    # Sum up individual months for YTD
                    for ytd_month in ytd_months:
                        month_budget_year = str(fiscal_year_start)[-2:] if ytd_month in months[:9] else str(fiscal_year_end)[-2:]
                        month_ly_year = str(last_fiscal_year_start)[-2:] if ytd_month in months[:9] else str(last_fiscal_year_end)[-2:]
                        month_actual_year = str(fiscal_year_start)[-2:] if ytd_month in months[:9] else str(fiscal_year_end)[-2:]
                        
                        month_budget_col = f'Budget-{ytd_month}-{month_budget_year}'
                        month_ly_col = f'LY-{ytd_month}-{month_ly_year}'
                        month_actual_col = f'Act-{ytd_month}-{month_actual_year}'
                        
                        if month_budget_col in df.columns:
                            ytd_budget_sum += df.loc[row_idx, month_budget_col]
                        if month_ly_col in df.columns:
                            ytd_ly_sum += df.loc[row_idx, month_ly_col]
                        if month_actual_col in df.columns:
                            ytd_actual_sum += df.loc[row_idx, month_actual_col]
                    
                    # Set YTD values
                    df.loc[row_idx, ytd_budget_col] = ytd_budget_sum
                    df.loc[row_idx, ytd_ly_col] = ytd_ly_sum
                    df.loc[row_idx, ytd_actual_col] = ytd_actual_sum
                    
                    # Calculate YTD Growth Rate (for non-total rows)
                    region_name = str(df.loc[row_idx, id_col]).upper()
                    if region_name not in ['NORTH TOTAL', 'WEST SALES', 'GROUP COMPANIES', 'GRAND TOTAL']:
                        if ytd_ly_sum != 0:
                            df.loc[row_idx, ytd_gr_col] = round(((ytd_actual_sum - ytd_ly_sum) / ytd_ly_sum * 100), 2)
                        else:
                            df.loc[row_idx, ytd_gr_col] = 0.0
                        
                        # Calculate YTD Achievement (skip for GROUP companies)
                        if 'GROUP' not in region_name:
                            if ytd_budget_sum > 0:
                                df.loc[row_idx, ytd_ach_col] = round((ytd_actual_sum / ytd_budget_sum * 100), 2)
                            else:
                                df.loc[row_idx, ytd_ach_col] = 0.0
                        else:
                            df.loc[row_idx, ytd_ach_col] = 0.0
                
                # Now handle totals by summing individual regions' YTD Gr/Ach values
                # North Total
                north_mask = df[id_col].isin(north_regions)
                if north_mask.any():
                    df.loc[df[id_col] == 'NORTH TOTAL', ytd_gr_col] = df[north_mask][ytd_gr_col].sum()
                    df.loc[df[id_col] == 'NORTH TOTAL', ytd_ach_col] = df[north_mask][ytd_ach_col].sum()
                
                # West Sales
                west_mask = df[id_col].isin(west_regions)
                if west_mask.any():
                    df.loc[df[id_col] == 'WEST SALES', ytd_gr_col] = df[west_mask][ytd_gr_col].sum()
                    df.loc[df[id_col] == 'WEST SALES', ytd_ach_col] = df[west_mask][ytd_ach_col].sum()
                
                # Group Companies (set Ach to 0)
                group_mask = df[id_col].str.upper().isin([g.upper() for g in group_companies])
                if group_mask.any():
                    df.loc[df[id_col] == 'GROUP COMPANIES', ytd_gr_col] = df[group_mask][ytd_gr_col].sum()
                    df.loc[df[id_col] == 'GROUP COMPANIES', ytd_ach_col] = 0.0
                
                # Grand Total (sum of all regions except other totals)
                non_total_mask = ~df[id_col].isin(['NORTH TOTAL', 'WEST SALES', 'GROUP COMPANIES'])
                if non_total_mask.any():
                    df.loc[df[id_col] == 'GRAND TOTAL', ytd_gr_col] = df[non_total_mask][ytd_gr_col].sum()
                    df.loc[df[id_col] == 'GRAND TOTAL', ytd_ach_col] = df[non_total_mask][ytd_ach_col].sum()
                
                # Add YTD columns to ordered list
                ordered_columns.extend([ytd_budget_col, ytd_ly_col, ytd_actual_col, ytd_gr_col, ytd_ach_col])
                break

    # Reorder dataframe columns according to auditor format
    existing_columns = [col for col in ordered_columns if col in df.columns]
    df_reordered = df[existing_columns].copy()
    
    return df_reordered

def process_sales_data_for_year(filepath, sheet_name, is_last_year=False, data_type='MT', 
                              fiscal_year_start=None, fiscal_year_end=None,
                              last_fiscal_year_start=None, last_fiscal_year_end=None,
                              custom_branch_mappings=None):
    """Enhanced process sales data with comprehensive branch naming"""
    try:
        print(f"=== Processing {'Last Year' if is_last_year else 'Current Year'} {data_type} Data: {sheet_name} ===")
        
        # Read and prepare the data
        df_sales = pd.read_excel(filepath, sheet_name=sheet_name, header=0)
        
        if isinstance(df_sales.columns, pd.MultiIndex):
            df_sales.columns = ['_'.join(col).strip() for col in df_sales.columns.values]
        
        df_sales = handle_duplicate_columns(df_sales)
        
        print(f"Initial data shape: {df_sales.shape}")
        print(f"Available columns: {df_sales.columns.tolist()[:10]}...")
        
        # Find required columns using enhanced search
        branch_col = enhanced_find_column(df_sales, column_aliases['Branch'], case_sensitive=False, threshold=75)
        date_col = enhanced_find_column(df_sales, column_aliases['Date'], case_sensitive=False, threshold=75)
        
        if data_type == 'MT':
            value_col = enhanced_find_column(df_sales, column_aliases['Actual Quantity'], case_sensitive=False, threshold=75)
            value_column_name = 'Actual Quantity'
        else:  # Value
            value_col = enhanced_find_column(df_sales, column_aliases['Value'] + column_aliases['Amount'], case_sensitive=False, threshold=75)
            value_column_name = 'Value' if data_type == 'Value' else 'Amount'
        
        print(f"Detected columns - Branch: {branch_col}, Date: {date_col}, Value: {value_col}")
        
        if not all([branch_col, date_col, value_col]):
            missing = []
            if not branch_col: missing.append(f'Branch (tried: {column_aliases["Branch"]})')
            if not date_col: missing.append(f'Date (tried: {column_aliases["Date"]})')
            if not value_col: missing.append(f'{value_column_name}')
            print(f"Missing required columns: {missing}")
            return {}
        
        # Process the data - filter out rows with blank/empty branch names
        df_sales = df_sales[[branch_col, date_col, value_col]].copy()
        df_sales.columns = ['Branch', 'Month Format', value_column_name]
        
        print(f"Data before branch processing: {df_sales.shape[0]} rows")
        
        # Enhanced branch name processing with debug info
        original_branches = df_sales['Branch'].unique()[:5]
        print(f"Sample original branch names: {[str(b) for b in original_branches]}")
        
        # Apply branch name normalization
        branch_mappings = custom_branch_mappings if custom_branch_mappings else create_branch_name_mappings()
        df_sales['Branch_Original'] = df_sales['Branch'].copy()
        df_sales['Branch'] = df_sales['Branch'].apply(
            lambda x: normalize_branch_name(x, branch_mappings)
        )
        
        # Show branch mapping results (sample)
        branch_mapping_sample = df_sales[['Branch_Original', 'Branch']].drop_duplicates().head(5)
        print("Branch mapping sample:")
        for _, row in branch_mapping_sample.iterrows():
            print(f"  '{row['Branch_Original']}' -> '{row['Branch']}'")
        
        # Filter out rows with blank/empty branch names
        df_sales = df_sales[
            df_sales['Branch'].notna() & 
            (df_sales['Branch'].astype(str).str.strip() != '') &
            (df_sales[value_column_name].notna())
        ]
        
        print(f"Data after branch normalization: {df_sales.shape[0]} rows")
        
        # Convert and clean data
        df_sales[value_column_name] = pd.to_numeric(df_sales[value_column_name], errors='coerce').fillna(0)
        
        # Enhanced month processing
        if pd.api.types.is_datetime64_any_dtype(df_sales['Month Format']):
            df_sales['Month'] = pd.to_datetime(df_sales['Month Format']).dt.strftime('%b')
        else:
            # Handle various month formats
            month_str = df_sales['Month Format'].astype(str).str.strip().str.title()
            
            # Try different month parsing strategies
            df_sales['Month'] = None
            
            # Strategy 1: Full month names
            try:
                mask_full_month = pd.to_datetime(month_str, format='%B', errors='coerce').notna()
                df_sales.loc[mask_full_month, 'Month'] = pd.to_datetime(month_str[mask_full_month], format='%B').dt.strftime('%b')
            except:
                pass
            
            # Strategy 2: Short month names (3 letters)
            mask_short_month = df_sales['Month'].isna() & month_str.str.len() == 3
            if mask_short_month.any():
                df_sales.loc[mask_short_month, 'Month'] = month_str[mask_short_month].str[:3]
            
            # Strategy 3: Extract from longer strings
            month_names = ['January', 'February', 'March', 'April', 'May', 'June',
                          'July', 'August', 'September', 'October', 'November', 'December']
            month_abbr = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                         'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            
            mask_remaining = df_sales['Month'].isna()
            for i, month_name in enumerate(month_names):
                month_mask = mask_remaining & month_str.str.contains(month_name, case=False, na=False)
                df_sales.loc[month_mask, 'Month'] = month_abbr[i]
            
            # Final fallback: take first 3 characters
            mask_still_empty = df_sales['Month'].isna()
            df_sales.loc[mask_still_empty, 'Month'] = month_str[mask_still_empty].str[:3]
        
        # Filter valid months and non-zero values
        valid_months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                       'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        df_sales = df_sales[
            (df_sales['Month'].isin(valid_months)) &
            (df_sales[value_column_name] != 0)
        ]
        
        print(f"Data after month filtering: {df_sales.shape[0]} rows")
        print(f"Months found: {sorted(df_sales['Month'].unique())}")
        
        if df_sales.empty:
            print("No valid data found after processing")
            return {}
        
        # Group by branch and month
        grouped = df_sales.groupby(['Branch', 'Month'])[value_column_name].sum().reset_index()
        print(f"Grouped data: {grouped.shape[0]} branch-month combinations")
        
        # Create the result dictionary with proper fiscal year logic
        result_data = {}
        
        # Define fiscal year months correctly
        # Fiscal year: Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec, Jan, Feb, Mar
        first_half_months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']  # Use fiscal_year_start
        second_half_months = ['Jan', 'Feb', 'Mar']  # Use fiscal_year_end
        
        for _, row in grouped.iterrows():
            branch = row['Branch']
            month = row['Month']
            amount = row[value_column_name]
            
            # Determine the year suffix based on fiscal year
            if is_last_year:
                # For last year data
                if month in first_half_months:  # Apr-Dec of last fiscal year
                    year = str(last_fiscal_year_start)[-2:]
                else:  # Jan-Mar of last fiscal year
                    year = str(last_fiscal_year_end)[-2:]
                col_name = f'LY-{month}-{year}'
            else:
                # For current year data
                if month in first_half_months:  # Apr-Dec of current fiscal year
                    year = str(fiscal_year_start)[-2:]
                else:  # Jan-Mar of current fiscal year
                    year = str(fiscal_year_end)[-2:]
                col_name = f'Act-{month}-{year}'
            
            if branch not in result_data:
                result_data[branch] = {}
            
            # Sum amounts for the same branch and period
            result_data[branch][col_name] = result_data[branch].get(col_name, 0) + amount
        
        print(f"Final result: {len(result_data)} branches with data")
        sample_branches = list(result_data.keys())[:3]
        for branch in sample_branches:
            print(f"Sample data for {branch}: {len(result_data[branch])} months")
        
        return result_data
        
    except Exception as e:
        print(f"Error processing sales data: {str(e)}")
        import traceback
        traceback.print_exc()
        return {}

def merge_last_year_data_with_current(current_data, last_year_data):
    """
    Enhanced function to merge last year data with current year data
    """
    print(f"=== Merging Last Year Data ===")
    print(f"Current data branches: {len(current_data)}")
    print(f"Last year data branches: {len(last_year_data)}")
    
    # Create merged result
    merged_data = current_data.copy()
    
    # Track branches found in last year but not current
    new_branches = []
    updated_branches = []
    
    for ly_branch, ly_data in last_year_data.items():
        if ly_branch in merged_data:
            # Update existing branch with last year data
            merged_data[ly_branch].update(ly_data)
            updated_branches.append(ly_branch)
        else:
            # Add new branch from last year data
            merged_data[ly_branch] = ly_data.copy()
            new_branches.append(ly_branch)
    
    print(f"Updated existing branches: {len(updated_branches)}")
    print(f"Added new branches from LY data: {len(new_branches)}")
    if new_branches:
        print(f"New branches: {new_branches[:5]}...")  # Show first 5
    
    return merged_data

# Main processing route
@region_bp.route('/process-region-analysis', methods=['POST'])
def process_region_analysis():
    """Process region analysis using existing uploaded files from main app"""
    try:
        print("=== DEBUG: Starting enhanced region analysis ===")
        data = request.json
        print(f"Request data keys: {list(data.keys()) if data else 'No data'}")
        
        # Get file paths from the main app's upload structure
        sales_filepath = data.get('sales_filepath')
        budget_filepath = data.get('budget_filepath')
        total_sales_filepath = data.get('total_sales_filepath')
        auditor_filepath = data.get('auditor_filepath')
        
        # Get selected sheets
        selected_sales_sheet = data.get('selected_sales_sheet')
        selected_budget_sheet = data.get('selected_budget_sheet')
        selected_total_sales_sheet = data.get('selected_total_sales_sheet')
        
        # Get custom branch mappings if provided
        custom_branch_mappings = data.get('custom_branch_mappings')
        
        if not sales_filepath or not budget_filepath:
            return jsonify({
                'success': False,
                'error': 'Sales and Budget files are required'
            }), 400
        
        if not selected_sales_sheet or not selected_budget_sheet:
            return jsonify({
                'success': False,
                'error': 'Sales and Budget sheet selection is required'
            }), 400
        
        # Check if files exist
        if not os.path.exists(sales_filepath):
            return jsonify({
                'success': False,
                'error': f'Sales file not found: {sales_filepath}'
            }), 404
            
        if not os.path.exists(budget_filepath):
            return jsonify({
                'success': False,
                'error': f'Budget file not found: {budget_filepath}'
            }), 404
        
        # Get current date and determine fiscal year
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
        
        months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar']
        
        # ENHANCED BUDGET PROCESSING
        print("=== DEBUG: Processing budget data with enhanced header detection ===")
        
        # Detect correct header row for budget sheet
        budget_header_row = detect_budget_sheet_header_row(budget_filepath, selected_budget_sheet)
        print(f"Using header row {budget_header_row} for budget sheet")
        
        # Process budget data with detected header row
        xls_budget = pd.ExcelFile(budget_filepath)
        df_budget = pd.read_excel(xls_budget, sheet_name=selected_budget_sheet, header=budget_header_row)
        df_budget.columns = df_budget.columns.str.strip()
        df_budget = df_budget.dropna(how='all').reset_index(drop=True)
        
        # Debug: Print first few columns to verify header detection
        print(f"Budget sheet columns after header detection: {df_budget.columns.tolist()[:10]}")
        print(f"Budget sheet shape: {df_budget.shape}")
        
        budget_data = process_budget_data(df_budget, group_type='region')
        if budget_data is None:
            return jsonify({
                'success': False,
                'error': 'Failed to process budget data for regions - no valid budget columns found'
            }), 400
        
        print(f"Budget data processed successfully. Columns: {budget_data.columns.tolist()}")
        
        # Process MT data with enhanced processing
        mt_cols = [col for col in budget_data.columns if col.endswith('_MT')]
        result_mt = pd.DataFrame()
        
        if mt_cols:
            result_mt['SALES in MT'] = budget_data['REGIONS'].copy()
            result_mt['SALES in MT'] = result_mt['SALES in MT'].replace([pd.NA, np.nan, None], '').apply(
                lambda x: str(x).strip().upper() if pd.notna(x) else ''
            )
            
            # Add budget columns
            for col in sorted(mt_cols):
                month_col = col.replace('_MT', '')
                result_mt[month_col] = budget_data[col]
            
            # Process current year MT data with enhanced processing
            actual_mt_data = {}
            if sales_filepath and os.path.exists(sales_filepath) and selected_sales_sheet:
                actual_mt_data = process_sales_data_for_year(
                    sales_filepath, 
                    selected_sales_sheet, 
                    is_last_year=False, 
                    data_type='MT',
                    fiscal_year_start=fiscal_year_start,
                    fiscal_year_end=fiscal_year_end,
                    last_fiscal_year_start=last_fiscal_year_start,
                    last_fiscal_year_end=last_fiscal_year_end,
                    custom_branch_mappings=custom_branch_mappings
                )
            
            # Enhanced last year MT data processing
            last_year_mt_data = {}
            use_sales_sheet_for_ly = False
            
            if total_sales_filepath and os.path.exists(total_sales_filepath) and selected_total_sales_sheet:
                last_year_mt_data = process_sales_data_for_year(
                    total_sales_filepath, 
                    selected_total_sales_sheet, 
                    is_last_year=True, 
                    data_type='MT',
                    fiscal_year_start=fiscal_year_start,
                    fiscal_year_end=fiscal_year_end,
                    last_fiscal_year_start=last_fiscal_year_start,
                    last_fiscal_year_end=last_fiscal_year_end,
                    custom_branch_mappings=custom_branch_mappings
                )
                if not last_year_mt_data:
                    use_sales_sheet_for_ly = True
            else:
                use_sales_sheet_for_ly = True
            
            if use_sales_sheet_for_ly and sales_filepath and os.path.exists(sales_filepath) and selected_sales_sheet:
                last_year_mt_data = process_sales_data_for_year(
                    sales_filepath, 
                    selected_sales_sheet, 
                    is_last_year=True, 
                    data_type='MT',
                    fiscal_year_start=fiscal_year_start,
                    fiscal_year_end=fiscal_year_end,
                    last_fiscal_year_start=last_fiscal_year_start,
                    last_fiscal_year_end=last_fiscal_year_end,
                    custom_branch_mappings=custom_branch_mappings
                )
            
            # Merge last year data with actual data using enhanced merge function
            actual_mt_data = merge_last_year_data_with_current(actual_mt_data, last_year_mt_data)
            
            # Add actual columns
            all_mt_cols = set()
            for branch_data in actual_mt_data.values():
                all_mt_cols.update(branch_data.keys())
            
            for col_name in sorted(all_mt_cols):
                if col_name not in result_mt.columns:
                    result_mt[col_name] = 0.0
            
            # Merge actual MT data
            for branch, data in actual_mt_data.items():
                matching_rows = result_mt[result_mt['SALES in MT'] == branch]
                
                if not matching_rows.empty:
                    if len(matching_rows) > 1:
                        # Aggregate duplicate rows
                        branch_data = result_mt[result_mt['SALES in MT'] == branch]
                        numeric_cols_branch = branch_data.select_dtypes(include=[np.number]).columns
                        
                        aggregated_data = {'SALES in MT': branch}
                        for col in result_mt.columns[1:]:
                            if col in numeric_cols_branch:
                                aggregated_data[col] = branch_data[col].sum()
                            else:
                                aggregated_data[col] = 0
                        
                        result_mt = result_mt[result_mt['SALES in MT'] != branch]
                        result_mt = pd.concat([result_mt, pd.DataFrame([aggregated_data])], ignore_index=True)
                        
                        idx = result_mt[result_mt['SALES in MT'] == branch].index[0]
                    else:
                        idx = matching_rows.index[0]
                    
                    for col, value in data.items():
                        if pd.notna(value) and value != 0:
                            result_mt.loc[idx, col] = value
                else:
                    # Add new branch
                    new_row_data = {'SALES in MT': branch}
                    for col in result_mt.columns[1:]:
                        new_row_data[col] = data.get(col, 0)
                    
                    result_mt = pd.concat([result_mt, pd.DataFrame([new_row_data])], ignore_index=True)
            
            # Final deduplication
            if result_mt['SALES in MT'].duplicated().any():
                numeric_cols_final = result_mt.select_dtypes(include=[np.number]).columns
                agg_dict = {col: 'sum' for col in numeric_cols_final}
                for col in result_mt.columns:
                    if col not in numeric_cols_final and col != 'SALES in MT':
                        agg_dict[col] = 'first'
                result_mt = result_mt.groupby('SALES in MT', as_index=False).agg(agg_dict)
            
            # Fill NaN values
            numeric_cols = result_mt.select_dtypes(include=[np.number]).columns
            result_mt[numeric_cols] = result_mt[numeric_cols].fillna(0)
            
            # Calculate Growth Rate and Achievement
            group_companies = ['GROUP']
            
            for month in months:
                budget_year = str(fiscal_year_start)[-2:] if month in months[:9] else str(fiscal_year_end)[-2:]
                actual_year = str(fiscal_year_start)[-2:] if month in months[:9] else str(fiscal_year_end)[-2:]
                ly_year = str(last_fiscal_year_start)[-2:] if month in months[:9] else str(last_fiscal_year_end)[-2:]
                
                budget_col = f'Budget-{month}-{budget_year}'
                actual_col = f'Act-{month}-{actual_year}'
                ly_col = f'LY-{month}-{ly_year}'
                gr_col = f'Gr-{month}-{actual_year}'
                ach_col = f'Ach-{month}-{actual_year}'
                
                # Add Gr and Ach columns
                if gr_col not in result_mt.columns:
                    result_mt[gr_col] = 0.0
                if ach_col not in result_mt.columns:
                    result_mt[ach_col] = 0.0
                
                # Growth Rate calculation
                if ly_col in result_mt.columns and actual_col in result_mt.columns:
                    result_mt[gr_col] = np.where(
                        (result_mt[ly_col] != 0) & (pd.notna(result_mt[ly_col])) & (pd.notna(result_mt[actual_col])),
                        ((result_mt[actual_col] - result_mt[ly_col]) / result_mt[ly_col] * 100).round(2),
                        0
                    )
                
                # Achievement calculation
                if budget_col in result_mt.columns and actual_col in result_mt.columns:
                    group_mask = result_mt['SALES in MT'].str.upper().isin([g.upper() for g in group_companies])
                    result_mt[budget_col] = result_mt[budget_col].round(2)
                    
                    result_mt[ach_col] = np.where(
                        (~group_mask) & (result_mt[budget_col].abs() > 0.01) &
                        (pd.notna(result_mt[budget_col])) & 
                        (pd.notna(result_mt[actual_col])),
                        (result_mt[actual_col] / result_mt[budget_col] * 100).round(2),
                        0.0
                    )
                    result_mt[ach_col] = np.where(group_mask, 0.0, result_mt[ach_col])
            
            # Add regional totals
            result_mt = add_regional_totals(result_mt, 'MT', fiscal_year_start, fiscal_year_end,
                                          last_fiscal_year_start, last_fiscal_year_end)
            
            # Apply auditor format
            result_mt = add_ytd_calculations_auditor_format(result_mt, 'MT', fiscal_year_start, fiscal_year_end,
                                                          last_fiscal_year_start, last_fiscal_year_end)
        
        # Process Value data (similar enhanced structure as MT data)
        value_cols = [col for col in budget_data.columns if col.endswith('_Value')]
        result_value = pd.DataFrame()
        
        if value_cols:
            result_value['SALES in Value'] = budget_data['REGIONS'].copy()
            result_value['SALES in Value'] = result_value['SALES in Value'].replace([pd.NA, np.nan, None], '').apply(
                lambda x: str(x).strip().upper() if pd.notna(x) else ''
            )
            
            # Add budget columns
            for col in sorted(value_cols):
                month_col = col.replace('_Value', '')
                result_value[month_col] = budget_data[col]
            
            # Process current year Value data with enhanced processing
            actual_value_data = {}
            if sales_filepath and os.path.exists(sales_filepath) and selected_sales_sheet:
                actual_value_data = process_sales_data_for_year(
                    sales_filepath, 
                    selected_sales_sheet, 
                    is_last_year=False, 
                    data_type='Value',
                    fiscal_year_start=fiscal_year_start,
                    fiscal_year_end=fiscal_year_end,
                    last_fiscal_year_start=last_fiscal_year_start,
                    last_fiscal_year_end=last_fiscal_year_end,
                    custom_branch_mappings=custom_branch_mappings
                )
            
            # Process last year Value data with enhanced processing
            last_year_value_data = {}
            use_sales_sheet_for_ly_value = False
            
            if total_sales_filepath and os.path.exists(total_sales_filepath) and selected_total_sales_sheet:
                last_year_value_data = process_sales_data_for_year(
                    total_sales_filepath, 
                    selected_total_sales_sheet, 
                    is_last_year=True, 
                    data_type='Value',
                    fiscal_year_start=fiscal_year_start,
                    fiscal_year_end=fiscal_year_end,
                    last_fiscal_year_start=last_fiscal_year_start,
                    last_fiscal_year_end=last_fiscal_year_end,
                    custom_branch_mappings=custom_branch_mappings
                )
                if not last_year_value_data:
                    use_sales_sheet_for_ly_value = True
            else:
                use_sales_sheet_for_ly_value = True
            
            if use_sales_sheet_for_ly_value and sales_filepath and os.path.exists(sales_filepath) and selected_sales_sheet:
                last_year_value_data = process_sales_data_for_year(
                    sales_filepath, 
                    selected_sales_sheet, 
                    is_last_year=True, 
                    data_type='Value',
                    fiscal_year_start=fiscal_year_start,
                    fiscal_year_end=fiscal_year_end,
                    last_fiscal_year_start=last_fiscal_year_start,
                    last_fiscal_year_end=last_fiscal_year_end,
                    custom_branch_mappings=custom_branch_mappings
                )
            
            # Merge last year data with actual data using enhanced merge function
            actual_value_data = merge_last_year_data_with_current(actual_value_data, last_year_value_data)
            
            # Add actual columns
            all_value_cols = set()
            for branch_data in actual_value_data.values():
                all_value_cols.update(branch_data.keys())
            
            for col_name in sorted(all_value_cols):
                if col_name not in result_value.columns:
                    result_value[col_name] = 0.0
            
            # Merge actual Value data (similar logic as MT)
            for branch, data in actual_value_data.items():
                matching_rows = result_value[result_value['SALES in Value'] == branch]
                
                if not matching_rows.empty:
                    if len(matching_rows) > 1:
                        # Aggregate duplicate rows
                        branch_data = result_value[result_value['SALES in Value'] == branch]
                        numeric_cols_branch = branch_data.select_dtypes(include=[np.number]).columns
                        
                        aggregated_data = {'SALES in Value': branch}
                        for col in result_value.columns[1:]:
                            if col in numeric_cols_branch:
                                aggregated_data[col] = branch_data[col].sum()
                            else:
                                aggregated_data[col] = 0
                        
                        result_value = result_value[result_value['SALES in Value'] != branch]
                        result_value = pd.concat([result_value, pd.DataFrame([aggregated_data])], ignore_index=True)
                        
                        idx = result_value[result_value['SALES in Value'] == branch].index[0]
                    else:
                        idx = matching_rows.index[0]
                    
                    for col, value in data.items():
                        if pd.notna(value) and value != 0:
                            result_value.loc[idx, col] = value
                else:
                    # Add new branch
                    new_row_data = {'SALES in Value': branch}
                    for col in result_value.columns[1:]:
                        new_row_data[col] = data.get(col, 0)
                    
                    result_value = pd.concat([result_value, pd.DataFrame([new_row_data])], ignore_index=True)
            
            # Final deduplication
            if result_value['SALES in Value'].duplicated().any():
                numeric_cols_final = result_value.select_dtypes(include=[np.number]).columns
                agg_dict = {col: 'sum' for col in numeric_cols_final}
                for col in result_value.columns:
                    if col not in numeric_cols_final and col != 'SALES in Value':
                        agg_dict[col] = 'first'
                result_value = result_value.groupby('SALES in Value', as_index=False).agg(agg_dict)
            
            # Fill NaN values
            numeric_cols_value = result_value.select_dtypes(include=[np.number]).columns
            result_value[numeric_cols_value] = result_value[numeric_cols_value].fillna(0)
            
            # Calculate Growth Rate and Achievement (same logic as MT)
            for month in months:
                budget_year = str(fiscal_year_start)[-2:] if month in months[:9] else str(fiscal_year_end)[-2:]
                actual_year = str(fiscal_year_start)[-2:] if month in months[:9] else str(fiscal_year_end)[-2:]
                ly_year = str(last_fiscal_year_start)[-2:] if month in months[:9] else str(last_fiscal_year_end)[-2:]
                
                budget_col = f'Budget-{month}-{budget_year}'
                actual_col = f'Act-{month}-{actual_year}'
                ly_col = f'LY-{month}-{ly_year}'
                gr_col = f'Gr-{month}-{actual_year}'
                ach_col = f'Ach-{month}-{actual_year}'
                
                # Add Gr and Ach columns
                if gr_col not in result_value.columns:
                    result_value[gr_col] = 0.0
                if ach_col not in result_value.columns:
                    result_value[ach_col] = 0.0
                
                # Growth Rate calculation
                if ly_col in result_value.columns and actual_col in result_value.columns:
                    result_value[gr_col] = np.where(
                        (result_value[ly_col] != 0) & (pd.notna(result_value[ly_col])) & (pd.notna(result_value[actual_col])),
                        ((result_value[actual_col] - result_value[ly_col]) / result_value[ly_col] * 100).round(2),
                        0
                    )
                
                # Achievement calculation
                if budget_col in result_value.columns and actual_col in result_value.columns:
                    group_mask = result_value['SALES in Value'].str.upper().isin([g.upper() for g in group_companies])
                    result_value[budget_col] = result_value[budget_col].round(2)
                    
                    result_value[ach_col] = np.where(
                        (~group_mask) & (result_value[budget_col].abs() > 0.01) &
                        (pd.notna(result_value[budget_col])) & 
                        (pd.notna(result_value[actual_col])),
                        (result_value[actual_col] / result_value[budget_col] * 100).round(2),
                        0.0
                    )
                    result_value[ach_col] = np.where(group_mask, 0.0, result_value[ach_col])
            
            # Add regional totals
            result_value = add_regional_totals(result_value, 'Value', fiscal_year_start, fiscal_year_end,
                                             last_fiscal_year_start, last_fiscal_year_end)
            
            # Apply auditor format
            result_value = add_ytd_calculations_auditor_format(result_value, 'Value', fiscal_year_start, fiscal_year_end,
                                                             last_fiscal_year_start, last_fiscal_year_end)
        
        # Prepare response
        result_data = {
            'mt_data': result_mt.to_dict('records') if not result_mt.empty else [],
            'value_data': result_value.to_dict('records') if not result_value.empty else [],
            'fiscal_year': fiscal_year_str,
            'columns': {
                'mt_columns': result_mt.columns.tolist() if not result_mt.empty else [],
                'value_columns': result_value.columns.tolist() if not result_value.empty else []
            },
            'fallback_info': {
                'mt_used_fallback': use_sales_sheet_for_ly,
                'value_used_fallback': use_sales_sheet_for_ly_value,
                'total_sales_file_available': total_sales_filepath and os.path.exists(total_sales_filepath),
                'total_sales_sheet_selected': bool(selected_total_sales_sheet),
                'custom_mappings_used': bool(custom_branch_mappings)
            }
        }
        
        print("=== DEBUG: Enhanced processing completed successfully ===")
        print(f"MT data rows: {len(result_data['mt_data'])}")
        print(f"Value data rows: {len(result_data['value_data'])}")
        
        return jsonify({
            'success': True,
            'data': result_data
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        
        return jsonify({
            'success': False,
            'error': f'Enhanced processing error: {str(e)}'
        }), 500

@region_bp.route('/download-combined-single-sheet', methods=['POST'])
def download_combined_single_sheet():
    """Generate and download Excel file with both MT and Value data in single sheet"""
    try:
        data = request.json
        mt_data = data.get('mt_data', [])
        value_data = data.get('value_data', [])
        mt_columns = data.get('mt_columns', [])
        value_columns = data.get('value_columns', [])
        fiscal_year = data.get('fiscal_year', '')
        include_both_tables = data.get('include_both_tables', True)
        single_sheet = data.get('single_sheet', True)
        
        if not mt_data and not value_data:
            return jsonify({
                'success': False,
                'error': 'No data provided for Excel generation'
            }), 400
        
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define enhanced formats for single sheet
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
            
            region_header_format = workbook.add_format({
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
            
            # Regional total formats
            north_total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#D4EDDA', 
                'font_color': '#155724', 'border': 1
            })
            
            west_total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#FFF3CD', 
                'font_color': '#856404', 'border': 1
            })
            
            group_total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#CCE5FF', 
                'font_color': '#004085', 'border': 1
            })
            
            grand_total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#C3E6CB', 
                'font_color': '#155724', 'border': 1, 'font_size': 11
            })
            
            # Create single worksheet
            worksheet_name = 'Region wise analysis'
            worksheet = workbook.add_worksheet(worksheet_name)
            
            current_row = 0
            
            # Main title
            main_title = f"Region-wise Sales Analysis - Enhanced Report (FY {fiscal_year})"
            max_cols = max(len(mt_columns) if mt_columns else 0, len(value_columns) if value_columns else 0)
            if max_cols > 1:
                worksheet.merge_range(current_row, 0, current_row, max_cols - 1, main_title, title_format)
            else:
                worksheet.write(current_row, 0, main_title, title_format)
            current_row += 2
            
            def get_total_format(region_name):
                """Get appropriate format based on region name"""
                region_upper = str(region_name).upper()
                if 'NORTH TOTAL' in region_upper:
                    return north_total_format
                elif 'WEST SALES' in region_upper:
                    return west_total_format
                elif 'GROUP COMPANIES' in region_upper:
                    return group_total_format
                elif 'GRAND TOTAL' in region_upper:
                    return grand_total_format
                else:
                    return None
            
            # Write MT Data Table
            if mt_data and include_both_tables:
                # MT Table Title
                mt_title = f"Region-wise SALES in Tonnage Analysis - FY {fiscal_year}"
                if len(mt_columns) > 1:
                    worksheet.merge_range(current_row, 0, current_row, len(mt_columns) - 1, mt_title, subtitle_format)
                else:
                    worksheet.write(current_row, 0, mt_title, subtitle_format)
                current_row += 2
                
                # MT Headers
                for col_num, column_name in enumerate(mt_columns):
                    if col_num == 0:  # First column (region names)
                        worksheet.write(current_row, col_num, column_name, region_header_format)
                        worksheet.set_column(col_num, col_num, 25)  # Wider for region names
                    else:
                        worksheet.write(current_row, col_num, column_name, header_format)
                        worksheet.set_column(col_num, col_num, 12)  # Standard width for data
                current_row += 1
                
                # MT Data
                df_mt = pd.DataFrame(mt_data)
                for row_num, row_data in enumerate(mt_data):
                    for col_num, column_name in enumerate(mt_columns):
                        cell_value = row_data.get(column_name, 0)
                        
                        if col_num == 0:  # Region column
                            region_name = str(cell_value)
                            worksheet.write(current_row, col_num, region_name, text_format)
                        else:  # Data columns
                            # Check if this is a total row
                            region_name = row_data.get(mt_columns[0], '')
                            total_fmt = get_total_format(region_name)
                            
                            if total_fmt:
                                worksheet.write(current_row, col_num, cell_value, total_fmt)
                            else:
                                worksheet.write(current_row, col_num, cell_value, num_format)
                    
                    current_row += 1
                
                # Add spacing between tables
                current_row += 3
            
            # Write Value Data Table
            if value_data and include_both_tables:
                # Value Table Title
                value_title = f"Region-wise SALES in Value Analysis - FY {fiscal_year}"
                if len(value_columns) > 1:
                    worksheet.merge_range(current_row, 0, current_row, len(value_columns) - 1, value_title, subtitle_format)
                else:
                    worksheet.write(current_row, 0, value_title, subtitle_format)
                current_row += 2
                
                # Value Headers
                for col_num, column_name in enumerate(value_columns):
                    if col_num == 0:  # First column (region names)
                        worksheet.write(current_row, col_num, column_name, region_header_format)
                        worksheet.set_column(col_num, col_num, 25)  # Wider for region names
                    else:
                        worksheet.write(current_row, col_num, column_name, header_format)
                        worksheet.set_column(col_num, col_num, 12)  # Standard width for data
                current_row += 1
                
                # Value Data
                df_value = pd.DataFrame(value_data)
                for row_num, row_data in enumerate(value_data):
                    for col_num, column_name in enumerate(value_columns):
                        cell_value = row_data.get(column_name, 0)
                        
                        if col_num == 0:  # Region column
                            region_name = str(cell_value)
                            worksheet.write(current_row, col_num, region_name, text_format)
                        else:  # Data columns
                            # Check if this is a total row
                            region_name = row_data.get(value_columns[0], '')
                            total_fmt = get_total_format(region_name)
                            
                            if total_fmt:
                                worksheet.write(current_row, col_num, cell_value, total_fmt)
                            else:
                                worksheet.write(current_row, col_num, cell_value, num_format)
                    
                    current_row += 1
            
            # If only one table is requested, write just that table
            elif mt_data and not include_both_tables:
                # Only MT table
                mt_title = f"Region-wise SALES in MT Analysis - FY {fiscal_year}"
                if len(mt_columns) > 1:
                    worksheet.merge_range(current_row, 0, current_row, len(mt_columns) - 1, mt_title, subtitle_format)
                current_row += 2
                
                # Write MT data (same logic as above)
                for col_num, column_name in enumerate(mt_columns):
                    fmt = region_header_format if col_num == 0 else header_format
                    worksheet.write(current_row, col_num, column_name, fmt)
                    worksheet.set_column(col_num, col_num, 25 if col_num == 0 else 12)
                current_row += 1
                
                for row_data in mt_data:
                    for col_num, column_name in enumerate(mt_columns):
                        cell_value = row_data.get(column_name, 0)
                        if col_num == 0:
                            worksheet.write(current_row, col_num, str(cell_value), text_format)
                        else:
                            region_name = row_data.get(mt_columns[0], '')
                            total_fmt = get_total_format(region_name)
                            fmt = total_fmt if total_fmt else num_format
                            worksheet.write(current_row, col_num, cell_value, fmt)
                    current_row += 1
                    
            elif value_data and not include_both_tables:
                # Only Value table
                value_title = f"Region-wise SALES in Value Analysis - FY {fiscal_year}"
                if len(value_columns) > 1:
                    worksheet.merge_range(current_row, 0, current_row, len(value_columns) - 1, value_title, subtitle_format)
                current_row += 2
                
                # Write Value data (same logic as above)
                for col_num, column_name in enumerate(value_columns):
                    fmt = region_header_format if col_num == 0 else header_format
                    worksheet.write(current_row, col_num, column_name, fmt)
                    worksheet.set_column(col_num, col_num, 25 if col_num == 0 else 12)
                current_row += 1
                
                for row_data in value_data:
                    for col_num, column_name in enumerate(value_columns):
                        cell_value = row_data.get(column_name, 0)
                        if col_num == 0:
                            worksheet.write(current_row, col_num, str(cell_value), text_format)
                        else:
                            region_name = row_data.get(value_columns[0], '')
                            total_fmt = get_total_format(region_name)
                            fmt = total_fmt if total_fmt else num_format
                            worksheet.write(current_row, col_num, cell_value, fmt)
                    current_row += 1
            
            # Add footer with generation info
            current_row += 3
            footer_info = [
                f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                f"Fiscal Year: {fiscal_year}",
                f"Enhanced Branch Processing: Enabled",
                f"MT Records: {len(mt_data)}",
                f"Value Records: {len(value_data)}",
                f"Total Records: {len(mt_data) + len(value_data)}"
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

        # Generate filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"region_enhanced_single_sheet_{fiscal_year}_{timestamp}.xlsx"
        
        return send_file(
            BytesIO(excel_data),
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        print(f"=== DEBUG: Error in enhanced single sheet generation ===")
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        
        return jsonify({
            'success': False,
            'error': f'Enhanced single sheet Excel generation error: {str(e)}'
        }), 500

@region_bp.route('/generate-region-report', methods=['POST'])
def generate_region_report():
    """Generate comprehensive enhanced region analysis report with multiple sheets"""
    try:
        data = request.json
        mt_data = data.get('mt_data', [])
        value_data = data.get('value_data', [])
        mt_columns = data.get('mt_columns', [])
        value_columns = data.get('value_columns', [])
        fiscal_year = data.get('fiscal_year', '')
        uploaded_files = data.get('uploaded_files', {})
        selected_sheets = data.get('selected_sheets', {})
        
        if not mt_data and not value_data:
            return jsonify({
                'success': False,
                'error': 'No data provided for report generation'
            }), 400
        
        # Generate timestamp for filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"region_enhanced_comprehensive_report_{fiscal_year}_{timestamp}.xlsx"
        
        # Create Excel report with multiple sheets
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define comprehensive formats
            title_format = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 20, 'font_color': '#FFFFFF', 'bg_color': '#2E5F8A',
                'border': 2, 'border_color': '#1F4E79'
            })
            
            subtitle_format = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 16, 'font_color': '#1F4E79', 'bg_color': '#D9E1F2',
                'border': 1
            })
            
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center',
                'fg_color': '#4472C4', 'font_color': 'white', 'border': 1,
                'font_size': 11
            })
            
            region_header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top', 'align': 'center',
                'fg_color': '#2E5F8A', 'font_color': 'white', 'border': 1,
                'font_size': 11
            })
            
            num_format = workbook.add_format({
                'num_format': '#,##0.00', 'border': 1, 'align': 'right'
            })
            
            text_format = workbook.add_format({
                'border': 1, 'valign': 'vcenter', 'align': 'left'
            })
            
            # Regional total formats
            north_total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#D4EDDA', 
                'font_color': '#155724', 'border': 1
            })
            
            west_total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#FFF3CD', 
                'font_color': '#856404', 'border': 1
            })
            
            group_total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#CCE5FF', 
                'font_color': '#004085', 'border': 1
            })
            
            grand_total_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#C3E6CB', 
                'font_color': '#155724', 'border': 1, 'font_size': 12
            })
            
            def get_total_format(region_name):
                """Get appropriate format based on region name"""
                region_upper = str(region_name).upper()
                if 'NORTH TOTAL' in region_upper:
                    return north_total_format
                elif 'WEST SALES' in region_upper:
                    return west_total_format
                elif 'GROUP COMPANIES' in region_upper:
                    return group_total_format
                elif 'GRAND TOTAL' in region_upper:
                    return grand_total_format
                else:
                    return None
            
            # Enhanced Combined Analysis Sheet
            if mt_data or value_data:
                combined_sheet = workbook.add_worksheet('Enhanced Region Analysis')
                current_row = 0
                
                # Main title
                main_title = f"Enhanced Region-wise Analysis - FY {fiscal_year}"
                max_cols = max(len(mt_columns), len(value_columns))
                if max_cols > 1:
                    combined_sheet.merge_range(current_row, 0, current_row, max_cols - 1, main_title, title_format)
                else:
                    combined_sheet.write(current_row, 0, main_title, title_format)
                current_row += 2
                
                # MT Section (only if MT data exists)
                if mt_data and mt_columns:
                    mt_section_title = "SALES in MT Analysis (Enhanced Branch Processing)"
                    if len(mt_columns) > 1:
                        combined_sheet.merge_range(current_row, 0, current_row, len(mt_columns) - 1, mt_section_title, subtitle_format)
                    else:
                        combined_sheet.write(current_row, 0, mt_section_title, subtitle_format)
                    current_row += 2
                    
                    # MT Headers
                    for col_num, column_name in enumerate(mt_columns):
                        fmt = region_header_format if col_num == 0 else header_format
                        combined_sheet.write(current_row, col_num, column_name, fmt)
                        combined_sheet.set_column(col_num, col_num, 25 if col_num == 0 else 12)
                    current_row += 1
                    
                    # MT Data
                    for row_data in mt_data:
                        for col_num, column_name in enumerate(mt_columns):
                            cell_value = row_data.get(column_name, 0)
                            if col_num == 0:
                                combined_sheet.write(current_row, col_num, str(cell_value), text_format)
                            else:
                                region_name = row_data.get(mt_columns[0], '')
                                total_fmt = get_total_format(region_name)
                                fmt = total_fmt if total_fmt else num_format
                                combined_sheet.write(current_row, col_num, cell_value, fmt)
                        current_row += 1
                
                current_row += 3  # Space between sections
                
                # Value Section
                if value_data and value_columns:
                    value_section_title = "SALES in Value Analysis (Enhanced Branch Processing)"
                    if len(value_columns) > 1:
                        combined_sheet.merge_range(current_row, 0, current_row, len(value_columns) - 1, value_section_title, subtitle_format)
                    else:
                        combined_sheet.write(current_row, 0, value_section_title, subtitle_format)
                    current_row += 2
                    
                    # Value Headers
                    for col_num, column_name in enumerate(value_columns):
                        fmt = region_header_format if col_num == 0 else header_format
                        combined_sheet.write(current_row, col_num, column_name, fmt)
                        combined_sheet.set_column(col_num, col_num, 25 if col_num == 0 else 12)
                    current_row += 1
                    
                    # Value Data
                    for row_data in value_data:
                        for col_num, column_name in enumerate(value_columns):
                            cell_value = row_data.get(column_name, 0)
                            if col_num == 0:
                                combined_sheet.write(current_row, col_num, str(cell_value), text_format)
                            else:
                                region_name = row_data.get(value_columns[0], '')
                                total_fmt = get_total_format(region_name)
                                fmt = total_fmt if total_fmt else num_format
                                combined_sheet.write(current_row, col_num, cell_value, fmt)
                        current_row += 1
                
                # Add footer with enhanced processing info
                current_row += 3
                footer_info = [
                    f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                    f"Enhanced Branch Processing: ENABLED",
                    f"Branch Name Standardization: ACTIVE",
                    f"Fuzzy Match Threshold: 85%",
                    f"Total Branches Processed: {len(set([row.get(mt_columns[0] if mt_columns else value_columns[0], '') for row in (mt_data + value_data)]))}",
                    f"Regional Classifications: Auto-Detected"
                ]
                
                footer_format = workbook.add_format({
                    'font_size': 9, 'italic': True, 'font_color': '#666666'
                })
                
                for info_line in footer_info:
                    combined_sheet.write(current_row, 0, info_line, footer_format)
                    current_row += 1
                
                # Freeze panes
                combined_sheet.freeze_panes(4, 1)
            
            # Set print settings for the sheet
            combined_sheet.set_landscape()
            combined_sheet.set_paper(9)  # A4
            combined_sheet.fit_to_pages(1, 0)  # Fit to 1 page wide
        
        # Convert to base64 for JSON response
        excel_data = output.getvalue()
        import base64
        excel_b64 = base64.b64encode(excel_data).decode('utf-8')
        
        return jsonify({
            'success': True,
            'data': excel_b64,
            'filename': filename,
            'content_type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })
        
    except Exception as e:
        print(f"=== DEBUG: Error in enhanced region report generation ===")
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        
        return jsonify({
            'success': False,
            'error': f'Enhanced region report generation error: {str(e)}'
        }), 500
