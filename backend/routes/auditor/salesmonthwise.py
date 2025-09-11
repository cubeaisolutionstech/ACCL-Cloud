from flask import Blueprint, request, jsonify, send_file, current_app, session
import pandas as pd
import numpy as np
import re
from io import BytesIO
import os
from werkzeug.utils import secure_filename
from datetime import datetime

# Create Blueprint
salesmonthwise_bp = Blueprint('salesmonthwise', __name__, url_prefix='/api/salesmonthwise')

def get_fiscal_years():
    """Get current and last fiscal year strings based on current date"""
    today = datetime.now()
    current_year = today.year
    current_month = today.month
    
    # Fiscal year runs April to March (April is start of new fiscal year)
    if current_month >= 4:
        fiscal_current_year_short = str(current_year)[-2:]
        fiscal_next_year_short = str(current_year + 1)[-2:]
        fiscal_last_year_short = str(current_year - 1)[-2:]
        fiscal_last_last_year_short = str(current_year - 2)[-2:]
    else:
        fiscal_current_year_short = str(current_year - 1)[-2:]
        fiscal_next_year_short = str(current_year)[-2:]
        fiscal_last_year_short = str(current_year - 2)[-2:]
        fiscal_last_last_year_short = str(current_year - 3)[-2:]
    
    current_fiscal = f"{fiscal_current_year_short}-{fiscal_next_year_short}"
    last_fiscal = f"{fiscal_last_year_short}-{fiscal_current_year_short}"
    last_last_fiscal = f"{fiscal_last_last_year_short}-{fiscal_last_year_short}"
    
    return current_fiscal, last_fiscal, last_last_fiscal

def generate_month_columns():
    """Generate dynamic month columns based on fiscal year in the exact order specified"""
    current_fiscal, last_fiscal, _ = get_fiscal_years()
    
    # Generate individual month columns in the specified order
    months = [
        ("Apr", f"Budget-Apr-{current_fiscal[:2]}", f"LY-Apr-{last_fiscal[:2]}", f"Act-Apr-{current_fiscal[:2]}", f"Gr-Apr-{current_fiscal[:2]}", f"Ach-Apr-{current_fiscal[:2]}"),
        ("May", f"Budget-May-{current_fiscal[:2]}", f"LY-May-{last_fiscal[:2]}", f"Act-May-{current_fiscal[:2]}", f"Gr-May-{current_fiscal[:2]}", f"Ach-May-{current_fiscal[:2]}"),
        ("Jun", f"Budget-Jun-{current_fiscal[:2]}", f"LY-Jun-{last_fiscal[:2]}", f"Act-Jun-{current_fiscal[:2]}", f"Gr-Jun-{current_fiscal[:2]}", f"Ach-Jun-{current_fiscal[:2]}"),
        ("Jul", f"Budget-Jul-{current_fiscal[:2]}", f"LY-Jul-{last_fiscal[:2]}", f"Act-Jul-{current_fiscal[:2]}", f"Gr-Jul-{current_fiscal[:2]}", f"Ach-Jul-{current_fiscal[:2]}"),
        ("Aug", f"Budget-Aug-{current_fiscal[:2]}", f"LY-Aug-{last_fiscal[:2]}", f"Act-Aug-{current_fiscal[:2]}", f"Gr-Aug-{current_fiscal[:2]}", f"Ach-Aug-{current_fiscal[:2]}"),
        ("Sep", f"Budget-Sep-{current_fiscal[:2]}", f"LY-Sep-{last_fiscal[:2]}", f"Act-Sep-{current_fiscal[:2]}", f"Gr-Sep-{current_fiscal[:2]}", f"Ach-Sep-{current_fiscal[:2]}"),
        ("Oct", f"Budget-Oct-{current_fiscal[:2]}", f"LY-Oct-{last_fiscal[:2]}", f"Act-Oct-{current_fiscal[:2]}", f"Gr-Oct-{current_fiscal[:2]}", f"Ach-Oct-{current_fiscal[:2]}"),
        ("Nov", f"Budget-Nov-{current_fiscal[:2]}", f"LY-Nov-{last_fiscal[:2]}", f"Act-Nov-{current_fiscal[:2]}", f"Gr-Nov-{current_fiscal[:2]}", f"Ach-Nov-{current_fiscal[:2]}"),
        ("Dec", f"Budget-Dec-{current_fiscal[:2]}", f"LY-Dec-{last_fiscal[:2]}", f"Act-Dec-{current_fiscal[:2]}", f"Gr-Dec-{current_fiscal[:2]}", f"Ach-Dec-{current_fiscal[:2]}"),
        ("Jan", f"Budget-Jan-{current_fiscal[3:]}", f"LY-Jan-{current_fiscal[:2]}", f"Act-Jan-{current_fiscal[3:]}", f"Gr-Jan-{current_fiscal[3:]}", f"Ach-Jan-{current_fiscal[3:]}"),
        ("Feb", f"Budget-Feb-{current_fiscal[3:]}", f"LY-Feb-{current_fiscal[:2]}", f"Act-Feb-{current_fiscal[3:]}", f"Gr-Feb-{current_fiscal[3:]}", f"Ach-Feb-{current_fiscal[3:]}"),
        ("Mar", f"Budget-Mar-{current_fiscal[3:]}", f"LY-Mar-{current_fiscal[:2]}", f"Act-Mar-{current_fiscal[3:]}", f"Gr-Mar-{current_fiscal[3:]}", f"Ach-Mar-{current_fiscal[3:]}")
    ]
    
    # Generate YTD periods in the specified order
    ytd_periods = [
        ("Apr-Jun", f"YTD-{current_fiscal} (Apr to Jun)Budget", f"YTD-{last_fiscal} (Apr to Jun)LY", f"Act-YTD-{current_fiscal} (Apr to Jun)", f"Gr-YTD-{current_fiscal} (Apr to Jun)", f"Ach-YTD-{current_fiscal} (Apr to Jun)"),
        ("Apr-Sep", f"YTD-{current_fiscal} (Apr to Sep)Budget", f"YTD-{last_fiscal} (Apr to Sep)LY", f"Act-YTD-{current_fiscal} (Apr to Sep)", f"Gr-YTD-{current_fiscal} (Apr to Sep)", f"Ach-YTD-{current_fiscal} (Apr to Sep)"),
        ("Apr-Dec", f"YTD-{current_fiscal} (Apr to Dec)Budget", f"YTD-{last_fiscal} (Apr to Dec)LY", f"Act-YTD-{current_fiscal} (Apr to Dec)", f"Gr-YTD-{current_fiscal} (Apr to Dec)", f"Ach-YTD-{current_fiscal} (Apr to Dec)"),
        ("Apr-Mar", f"YTD-{current_fiscal} (Apr to Mar) Budget", f"YTD-{last_fiscal} (Apr to Mar)LY", f"Act-YTD-{current_fiscal} (Apr to Mar)", f"Gr-YTD-{current_fiscal} (Apr to Mar)", f"Ach-YTD-{current_fiscal} (Apr to Mar)")
    ]
    
    return months, ytd_periods

def build_custom_headers(table_type):
    """Build custom headers in the exact order specified"""
    current_fiscal, last_fiscal, _ = get_fiscal_years()
    months, ytd_periods = generate_month_columns()
    
    # Start with the table type header
    headers = [f"SALES in {table_type}"]
    
    # Add Apr-Jun individual months
    for month in months[:3]:  # Apr, May, Jun
        headers.extend(month[1:])  # Skip month name, add all other columns
    
    # Add Apr-Jun YTD
    headers.extend(ytd_periods[0][1:])  # Apr-Jun YTD
    
    # Add Jul-Sep individual months
    for month in months[3:6]:  # Jul, Aug, Sep
        headers.extend(month[1:])  # Skip month name, add all other columns
    
    # Add Apr-Sep YTD
    headers.extend(ytd_periods[1][1:])  # Apr-Sep YTD
    
    # Add Oct-Dec individual months
    for month in months[6:9]:  # Oct, Nov, Dec
        headers.extend(month[1:])  # Skip month name, add all other columns
    
    # Add Apr-Dec YTD
    headers.extend(ytd_periods[2][1:])  # Apr-Dec YTD
    
    # Add Jan-Mar individual months
    for month in months[9:12]:  # Jan, Feb, Mar
        headers.extend(month[1:])  # Skip month name, add all other columns
    
    # Add Apr-Mar YTD (Full Year)
    headers.extend(ytd_periods[3][1:])  # Apr-Mar YTD
    
    return headers

def extract_tables(df_sheet, possible_headers):
    """Extract table positions based on possible headers"""
    for idx, row in df_sheet.iterrows():
        row_str = ' '.join(str(cell) for cell in row if pd.notna(cell)).strip()
        for header in possible_headers:
            if re.search(re.escape(header), row_str, re.IGNORECASE):
                data_start = idx + 1
                return idx, data_start
    return None, None

def rename_columns(columns):
    """Rename columns to handle duplicates and nan values"""
    renamed = []
    for col in columns:
        if pd.isna(col) or str(col).lower() in ['nan', 'none']:
            renamed.append('Unnamed')
        else:
            renamed.append(str(col))
    return renamed

def handle_duplicate_columns(df):
    """Handle duplicate column names"""
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

def process_sales_analysis_core(filepath, sheet_name):
    """
    Core logic for processing sales analysis data with correct column order
    """
    current_fiscal, last_fiscal, _ = get_fiscal_years()
    
    # Build custom headers for both tables
    custom_headers_mt = build_custom_headers('MT')
    custom_headers_value = build_custom_headers('Value')
    
    xls = pd.ExcelFile(filepath)
    
    # Find sales analysis sheet
    sales_analysis_sheet = None
    for sheet in xls.sheet_names:
        if re.search(r'sales\s*analysis\s*month\s*wise', sheet.lower(), re.IGNORECASE):
            sales_analysis_sheet = sheet
            break
    
    if not sales_analysis_sheet:
        raise ValueError(f'No Sales Analysis Month Wise sheet found. Available sheets: {xls.sheet_names}')
    
    df_sheet = pd.read_excel(xls, sheet_name=sales_analysis_sheet, header=None, dtype=str)
    
    # Define possible headers
    table1_possible_headers = [
        "SALES in MT", "SALES IN MT", "Sales in MT", "SALES IN TONNAGE", "SALES IN TON",
        "Tonnage", "TONNAGE", "Tonnage Sales", "Sales Tonnage", "Metric Tons", "MT Sales"
    ]
    table2_possible_headers = [
        "SALES in Value", "SALES IN VALUE", "Sales in Value", "SALES IN RS", "VALUE SALES",
        "Value", "VALUE", "Sales Value"
    ]
    
    # Extract table positions
    idx1, data_start1 = extract_tables(df_sheet, table1_possible_headers)
    idx2, data_start2 = extract_tables(df_sheet, table2_possible_headers)
    
    if idx1 is None:
        raise ValueError('Could not locate SALES in MT table header')
    
    # Process Table 1 (MT)
    table1_end = idx2 if idx2 is not None and idx2 > idx1 else len(df_sheet)
    table1 = df_sheet.iloc[data_start1:table1_end].dropna(how='all')
    table1.columns = df_sheet.iloc[idx1]
    table1.columns = table1.columns.map(str)
    table1.columns = rename_columns(table1.columns)
    table1 = handle_duplicate_columns(table1)
    
    # Apply custom headers with correct column count
    available_cols = len(table1.columns)
    table1.columns = custom_headers_mt[:available_cols]
    
    # Convert numeric columns
    for col in table1.columns[1:]:  
        table1[col] = pd.to_numeric(table1[col], errors='coerce').fillna(0).astype(float)
    
    # Process Table 2 (Value) if it exists
    table2 = None
    if idx2 is not None and idx2 > idx1:
        table2 = df_sheet.iloc[data_start2:].dropna(how='all')
        table2.columns = df_sheet.iloc[idx2]
        table2.columns = table2.columns.map(str)
        table2.columns = rename_columns(table2.columns)
        table2 = handle_duplicate_columns(table2)
        
        # Apply custom headers with correct column count
        available_cols = len(table2.columns)
        table2.columns = custom_headers_value[:available_cols]
        
        for col in table2.columns[1:]:
            table2[col] = pd.to_numeric(table2[col], errors='coerce').fillna(0).astype(float)
    
    return table1, table2, current_fiscal, last_fiscal

def clean_column_name(col_name):
    """
    Clean column name for comparison
    """
    import re
    # Convert to lowercase and remove special characters, keeping numbers and letters
    cleaned = re.sub(r'[^a-z0-9]', '', str(col_name).lower())
    return cleaned

def calculate_column_similarity(col1, col2):
    """
    Calculate similarity between two column names
    """
    # Simple similarity based on common substrings
    if col1 == col2:
        return 1.0
    
    if col1 in col2 or col2 in col1:
        return 0.8
    
    # Check for common patterns
    common_patterns = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 
                      'jul', 'aug', 'sep', 'oct', 'nov', 'dec',
                      'ytd', 'budget', 'actual', 'act', 'ly', 'gr', 'ach']
    
    col1_patterns = [p for p in common_patterns if p in col1]
    col2_patterns = [p for p in common_patterns if p in col2]
    
    common_count = len(set(col1_patterns) & set(col2_patterns))
    total_patterns = len(set(col1_patterns) | set(col2_patterns))
    
    if total_patterns > 0:
        return common_count / total_patterns
    
    return 0.0

def find_best_column_match(target_col, available_cols):
    """
    Find the best matching column name using fuzzy matching
    """
    target_clean = clean_column_name(target_col)
    
    best_match = None
    best_score = 0
    
    for available_col in available_cols:
        available_clean = clean_column_name(available_col)
        score = calculate_column_similarity(target_clean, available_clean)
        
        if score > best_score and score > 0.6:  # Minimum 60% similarity
            best_score = score
            best_match = available_col
    
    return best_match

def add_accllp_row_with_fixed_mapping(table, table_type, session_totals):
    """
    Add ACCLLP row with proper session totals mapping
    """
    if table is None or table.empty:
        return table
    
    updated_table = table.copy()
    first_col = updated_table.columns[0]
    
    # Create ACCLLP row
    accllp_row = {col: 0.0 for col in updated_table.columns}
    accllp_row[first_col] = 'ACCLLP'
    
    # Create TOTAL SALES row
    total_sales_row = {col: 0.0 for col in updated_table.columns}
    total_sales_row[first_col] = 'TOTAL SALES'
    
    # Map session totals to table columns
    if session_totals:
        # Get the appropriate totals data
        if table_type == 'MT':
            totals_data = session_totals.get('tonnage', {})
        else:
            totals_data = session_totals.get('value', {})
        
        if totals_data:
            # Strategy 1: Direct exact match
            for sales_col in updated_table.columns[1:]:  # Skip first column (organization names)
                if sales_col in totals_data:
                    try:
                        value = totals_data[sales_col]
                        numeric_value = pd.to_numeric(value, errors='coerce')
                        if pd.notna(numeric_value) and not np.isinf(numeric_value) and numeric_value != 0:
                            rounded_value = round(float(numeric_value), 2)
                            accllp_row[sales_col] = rounded_value
                            total_sales_row[sales_col] = rounded_value
                    except (ValueError, TypeError):
                        continue
            
            # Strategy 2: Fuzzy matching for remaining columns
            for sales_col in updated_table.columns[1:]:
                if accllp_row[sales_col] == 0.0:  # Only map unmapped columns
                    best_match = find_best_column_match(sales_col, list(totals_data.keys()))
                    if best_match:
                        try:
                            value = totals_data[best_match]
                            numeric_value = pd.to_numeric(value, errors='coerce')
                            if pd.notna(numeric_value) and not np.isinf(numeric_value) and numeric_value != 0:
                                rounded_value = round(float(numeric_value), 2)
                                accllp_row[sales_col] = rounded_value
                                total_sales_row[sales_col] = rounded_value
                        except (ValueError, TypeError):
                            continue
    
    # Remove existing ACCLLP/TOTAL SALES rows
    updated_table = updated_table[
        ~updated_table[first_col].astype(str).str.upper().isin(['ACCLLP', 'TOTAL SALES'])
    ].reset_index(drop=True)
    
    # Add new rows
    accllp_df = pd.DataFrame([accllp_row])
    total_sales_df = pd.DataFrame([total_sales_row])
    accllp_df = accllp_df[updated_table.columns]
    total_sales_df = total_sales_df[updated_table.columns]
    updated_table = pd.concat([updated_table, accllp_df, total_sales_df], ignore_index=True)
    
    return updated_table

def add_standard_accllp_row(table):
    """Add ACCLLP row with standard calculations (no session integration)"""
    if table is None or table.empty:
        return table
    
    updated_table = table.copy()
    first_col = updated_table.columns[0]
    
    # Create ACCLLP row with zeros
    accllp_row = {col: 0.0 for col in updated_table.columns}
    accllp_row[first_col] = 'ACCLLP'
    
    # Create TOTAL SALES row with sum of all other rows
    total_sales_row = {col: 0.0 for col in updated_table.columns}
    total_sales_row[first_col] = 'TOTAL SALES'
    
    # Calculate totals from existing data
    for col in updated_table.columns[1:]:
        try:
            col_sum = updated_table[col].sum()
            if pd.notna(col_sum) and not np.isinf(col_sum):
                total_sales_row[col] = round(float(col_sum), 2)
        except:
            total_sales_row[col] = 0.0
    
    # Remove existing ACCLLP/TOTAL SALES rows
    updated_table = updated_table[
        ~updated_table[first_col].astype(str).str.upper().isin(['ACCLLP', 'TOTAL SALES'])
    ].reset_index(drop=True)
    
    # Add new rows
    accllp_df = pd.DataFrame([accllp_row])
    total_sales_df = pd.DataFrame([total_sales_row])
    accllp_df = accllp_df[updated_table.columns]
    total_sales_df = total_sales_df[updated_table.columns]
    updated_table = pd.concat([updated_table, accllp_df, total_sales_df], ignore_index=True)
    
    return updated_table

@salesmonthwise_bp.route('/process-sales-monthwise-with-session', methods=['POST'])
def process_sales_monthwise_with_session():
    """Process Sales Analysis with session totals passed directly from frontend"""
    try:
        data = request.get_json()
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        session_totals = data.get('session_totals', {})  # Get totals from request
        
        if not filepath or not sheet_name:
            return jsonify({
                'success': False,
                'error': 'Filepath and sheet_name are required'
            }), 400
        
        if not os.path.exists(filepath):
            return jsonify({
                'success': False,
                'error': 'File not found'
            }), 404
        
        # Process core sales analysis
        table1, table2, current_fiscal, last_fiscal = process_sales_analysis_core(filepath, sheet_name)
        
        # Add ACCLLP rows WITH direct session integration
        table1_with_accllp = add_accllp_row_with_fixed_mapping(table1, 'MT', session_totals)
        table2_with_accllp = add_accllp_row_with_fixed_mapping(table2, 'Value', session_totals)
        
        has_session_data = bool(session_totals)
        
        response_data = {
            'success': True,
            'fiscal_years': {
                'current': current_fiscal,
                'last': last_fiscal
            },
            'sales_mt_table': {
                'data': table1_with_accllp.to_dict('records') if table1_with_accllp is not None else [],
                'columns': table1_with_accllp.columns.tolist() if table1_with_accllp is not None else [],
                'shape': table1_with_accllp.shape if table1_with_accllp is not None else [0, 0]
            },
            'sales_value_table': {
                'data': table2_with_accllp.to_dict('records') if table2_with_accllp is not None else [],
                'columns': table2_with_accllp.columns.tolist() if table2_with_accllp is not None else [],
                'shape': table2_with_accllp.shape if table2_with_accllp is not None else [0, 0]
            },
            'session_data_used': has_session_data,
            'message': f'Sales Analysis processed {"with direct session totals" if has_session_data else "without session data"}'
        }
        
        return jsonify(response_data)
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'Error processing Sales Analysis: {str(e)}'
        }), 500

@salesmonthwise_bp.route('/download-sales-monthwise-excel', methods=['POST'])
def download_sales_monthwise_excel():
    """Download Sales Analysis Month-wise data as Excel file"""
    try:
        data = request.get_json()
        sales_mt_data = data.get('sales_mt_data')
        sales_value_data = data.get('sales_value_data')
        
        if not sales_mt_data and not sales_value_data:
            return jsonify({
                'success': False,
                'error': 'No data provided for export'
            }), 400
        
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Sales Analysis Month wise')
            writer.sheets['Sales_Analysis_Monthly'] = worksheet
            
            # Define formats
            title_format = workbook.add_format({
                'bold': True, 'align': 'center', 'valign': 'vcenter',
                'font_size': 14, 'font_color': '#000000'
            })
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top',
                'fg_color': '#4472C4', 'font_color': 'white', 'border': 1
            })
            num_format = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
            accllp_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#FFE6CC', 'border': 1
            })
            session_integrated_format = workbook.add_format({
                'bold': True, 'num_format': '#,##0.00', 'bg_color': '#D4EDDA', 'border': 1
            })
            text_format = workbook.add_format({'border': 1})
            
            # Start writing from row 4 to leave space for headers
            start_row = 4
            
            # Write MT table
            if sales_mt_data and sales_mt_data.get('data'):
                table1_df = pd.DataFrame(sales_mt_data['data'])
                if not table1_df.empty:
                    # Clean numeric data
                    numeric_cols = table1_df.select_dtypes(include=[np.number]).columns
                    table1_df[numeric_cols] = table1_df[numeric_cols].replace([np.inf, -np.inf], 0)
                    table1_df[numeric_cols] = table1_df[numeric_cols].fillna(0)
                    
                    # Write title
                    worksheet.merge_range(
                        start_row, 0, start_row, len(table1_df.columns) - 1,
                        "SALES IN MT - MONTH WISE (with Session Totals Integration)", title_format
                    )
                    
                    # Write headers
                    for col_num, value in enumerate(table1_df.columns.values):
                        worksheet.write(start_row + 3, col_num, value, header_format)
                    
                    # Write data
                    first_col = table1_df.columns[0]
                    for row_num in range(len(table1_df)):
                        is_accllp = str(table1_df.iloc[row_num, 0]).upper() in ['ACCLLP', 'TOTAL SALES']
                        
                        # Check if ACCLLP row has session data (non-zero values)
                        has_session_data = False
                        if is_accllp:
                            row_data = table1_df.iloc[row_num, 1:].values
                            has_session_data = any(pd.to_numeric(val, errors='coerce') > 0 for val in row_data if pd.notna(val))
                        
                        for col_num in range(len(table1_df.columns)):
                            value = table1_df.iloc[row_num, col_num]
                            if col_num > 0 and pd.api.types.is_numeric_dtype(table1_df.iloc[:, col_num]):
                                if pd.isna(value) or np.isinf(value):
                                    value = 0.0
                                else:
                                    value = float(value)
                            
                            # Choose format based on row type and data presence
                            if is_accllp and has_session_data:
                                fmt = session_integrated_format if col_num > 0 else text_format
                            elif is_accllp:
                                fmt = accllp_format if col_num > 0 else text_format
                            else:
                                fmt = text_format if col_num == 0 else num_format
                            
                            worksheet.write(start_row + 4 + row_num, col_num, value if value != 0.00 else "", fmt)
                    
                    # Set column widths
                    for i, col in enumerate(table1_df.columns):
                        max_len = max((table1_df[col].astype(str).str.len().max(), len(str(col)))) + 2
                        worksheet.set_column(i, i, min(max_len, 25))
                    
                    start_row += len(table1_df) + 7
            
            # Write Value table
            if sales_value_data and sales_value_data.get('data'):
                table2_df = pd.DataFrame(sales_value_data['data'])
                if not table2_df.empty:
                    # Clean numeric data
                    numeric_cols = table2_df.select_dtypes(include=[np.number]).columns
                    table2_df[numeric_cols] = table2_df[numeric_cols].replace([np.inf, -np.inf], 0)
                    table2_df[numeric_cols] = table2_df[numeric_cols].fillna(0)
                    
                    # Write title
                    worksheet.merge_range(
                        start_row, 0, start_row, len(table2_df.columns) - 1,
                        "SALES IN VALUE - MONTH WISE (with Session Totals Integration)", title_format
                    )
                    
                    # Write headers
                    for col_num, value in enumerate(table2_df.columns.values):
                        worksheet.write(start_row + 3, col_num, value, header_format)
                    
                    # Write data
                    first_col = table2_df.columns[0]
                    for row_num in range(len(table2_df)):
                        is_accllp = str(table2_df.iloc[row_num, 0]).upper() in ['ACCLLP', 'TOTAL SALES']
                        
                        # Check if ACCLLP row has session data (non-zero values)
                        has_session_data = False
                        if is_accllp:
                            row_data = table2_df.iloc[row_num, 1:].values
                            has_session_data = any(pd.to_numeric(val, errors='coerce') > 0 for val in row_data if pd.notna(val))
                        
                        for col_num in range(len(table2_df.columns)):
                            value = table2_df.iloc[row_num, col_num]
                            if col_num > 0 and pd.api.types.is_numeric_dtype(table2_df.iloc[:, col_num]):
                                if pd.isna(value) or np.isinf(value):
                                    value = 0.0
                                else:
                                    value = float(value)
                            
                            # Choose format based on row type and data presence
                            if is_accllp and has_session_data:
                                fmt = session_integrated_format if col_num > 0 else text_format
                            elif is_accllp:
                                fmt = accllp_format if col_num > 0 else text_format
                            else:
                                fmt = text_format if col_num == 0 else num_format
                            
                            worksheet.write(start_row + 4 + row_num, col_num, value if value != 0.00 else "", fmt)
                    
                    # Set column widths
                    for i, col in enumerate(table2_df.columns):
                        max_len = max((table2_df[col].astype(str).str.len().max(), len(str(col)))) + 2
                        worksheet.set_column(i, i, min(max_len, 25))
        
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='sales_analysis_monthly_with_session_totals.xlsx'
        )
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'Error generating Excel file: {str(e)}'
        }), 500

@salesmonthwise_bp.route('/process-sales-monthwise', methods=['POST'])
def process_sales_monthwise():
    """Process Sales Analysis without session totals (standard version)"""
    try:
        data = request.get_json()
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        
        if not filepath or not sheet_name:
            return jsonify({
                'success': False,
                'error': 'Filepath and sheet_name are required'
            }), 400
        
        if not os.path.exists(filepath):
            return jsonify({
                'success': False,
                'error': 'File not found'
            }), 404
        
        # Process core sales analysis
        table1, table2, current_fiscal, last_fiscal = process_sales_analysis_core(filepath, sheet_name)
        
        # Add ACCLLP rows WITHOUT session integration
        table1_with_accllp = add_standard_accllp_row(table1)
        table2_with_accllp = add_standard_accllp_row(table2)
        
        response_data = {
            'success': True,
            'fiscal_years': {
                'current': current_fiscal,
                'last': last_fiscal
            },
            'sales_mt_table': {
                'data': table1_with_accllp.to_dict('records') if table1_with_accllp is not None else [],
                'columns': table1_with_accllp.columns.tolist() if table1_with_accllp is not None else [],
                'shape': table1_with_accllp.shape if table1_with_accllp is not None else [0, 0]
            },
            'sales_value_table': {
                'data': table2_with_accllp.to_dict('records') if table2_with_accllp is not None else [],
                'columns': table2_with_accllp.columns.tolist() if table2_with_accllp is not None else [],
                'shape': table2_with_accllp.shape if table2_with_accllp is not None else [0, 0]
            },
            'session_data_used': False,
            'message': 'Sales Analysis processed without session totals'
        }
        
        return jsonify(response_data)
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'Error processing Sales Analysis: {str(e)}'
        }), 500

__all__ = ['salesmonthwise_bp']