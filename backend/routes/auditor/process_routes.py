from flask import Blueprint, request, jsonify
import pandas as pd
import traceback
from process import (
    handle_duplicate_columns,
    clean_and_convert_numeric,
    process_budget_data,
    process_last_year_data,
    extract_tables,
    standardize_column_names,
    validate_dataframe,
    optimize_dataframe_memory
)

# Create the blueprint - this line is CRITICAL
process_bp = Blueprint('process', __name__)

@process_bp.route('/process-sheet', methods=['POST'])
def process_sheet():
    """Process a specific sheet from uploaded file"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        processing_type = data.get('processing_type', 'general')
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        # Read the specific sheet
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        
        # Process based on type
        if processing_type == 'budget':
            result = process_budget_sheet(df)
        elif processing_type == 'sales':
            result = process_sales_sheet(df)
        elif processing_type == 'auditor':
            result = process_auditor_sheet(df)
        elif processing_type == 'totalSales':
            result = process_total_sales_sheet(df)
        else:
            # General processing
            result = process_general_sheet(df)
        
        return jsonify(result)
        
    except Exception as e:
        return jsonify({'error': f'Processing failed', 'traceback': traceback.format_exc()}), 500

def process_general_sheet(df):
    """General sheet processing"""
    try:
        # Basic cleaning
        processed_df = handle_duplicate_columns(df.copy())
        
        # Try to detect header row
        header_row = 0
        for i in range(min(10, len(df))):
            row = df.iloc[i]
            if row.notna().sum() > len(df.columns) * 0.5:  # More than 50% non-null values
                header_row = i
                break
        
        # Re-read with detected header
        if header_row > 0:
            processed_df = df.iloc[header_row:].reset_index(drop=True)
            processed_df.columns = df.iloc[header_row-1] if header_row > 0 else df.iloc[0]
        
        # Clean and convert
        processed_df = clean_and_convert_numeric(processed_df)
        
        # Optimize memory
        processed_df, memory_info = optimize_dataframe_memory(processed_df)
        
        return {
            'success': True,
            'data': processed_df.head(100).to_dict('records'),
            'columns': processed_df.columns.tolist(),
            'shape': processed_df.shape,
            'type': 'general',
            'header_row': header_row,
            'memory_info': memory_info
        }
    except Exception as e:
        return 

def process_budget_sheet(df):
    """Process budget data sheet"""
    try:
        # Detect if this is region or product budget
        sheet_text = ' '.join(df.astype(str).values.flatten()).lower()
        group_type = 'product' if any(term in sheet_text for term in ['product', 'acetic', 'auxil']) else 'region'
        
        processed_data, message = process_budget_data(df, group_type)
        
        if processed_data is not None:
            # Optimize memory
            processed_data, memory_info = optimize_dataframe_memory(processed_data)
            
            return {
                'success': True,
                'data': processed_data.head(100).to_dict('records'),
                'columns': processed_data.columns.tolist(),
                'shape': processed_data.shape,
                'type': 'budget',
                'group_type': group_type,
                'message': message,
                'memory_info': memory_info
            }
        else:
            return {'error': message}
            
    except Exception as e:
        return 

def process_sales_sheet(df):
    """Process sales data sheet"""
    try:
        # Look for sales analysis tables
        possible_headers = [
            'Region wise Sales Analysis Month wise',
            'Product wise Sales Analysis Month wise', 
            'Sales Analysis Month wise',
            'Region Analysis',
            'Product Analysis'
        ]
        
        header_row, data_start = extract_tables(df, possible_headers)
        
        if header_row is not None and data_start is not None:
            # Extract the table
            table_df = df.iloc[data_start:].copy()
            table_df.columns = df.iloc[header_row]
            
            # Clean the dataframe
            table_df = table_df.dropna(how='all').reset_index(drop=True)
            table_df = handle_duplicate_columns(table_df)
            
            # Standardize column names
            table_df, column_mapping = standardize_column_names(table_df)
            
            # Clean and convert numeric
            table_df = clean_and_convert_numeric(table_df)
            
            # Optimize memory
            table_df, memory_info = optimize_dataframe_memory(table_df)
            
            return {
                'success': True,
                'data': table_df.head(100).to_dict('records'),
                'columns': table_df.columns.tolist(),
                'shape': table_df.shape,
                'type': 'sales',
                'header_row': header_row,
                'data_start': data_start,
                'memory_info': memory_info
            }
        else:
            return {'error': 'Could not locate sales analysis table in the sheet'}
            
    except Exception as e:
        return 

def process_auditor_sheet(df):
    """Process auditor format sheet"""
    try:
        # Determine analysis type based on content
        sheet_text = ' '.join(df.astype(str).values.flatten()).lower()
        
        is_region_analysis = 'region' in sheet_text
        is_product_analysis = any(term in sheet_text for term in ['product', 'ts-pw', 'ero-pw'])
        
        possible_headers = []
        if is_region_analysis:
            possible_headers = ['Region wise Sales Analysis Month wise', 'Region Analysis']
        elif is_product_analysis:
            possible_headers = ['Product wise Sales Analysis Month wise', 'Product Analysis', 'TS-PW', 'ERO-PW']
        else:
            possible_headers = ['Sales Analysis Month wise', 'Analysis']
        
        header_row, data_start = extract_tables(df, possible_headers, is_product_analysis)
        
        if header_row is not None and data_start is not None:
            # Extract the table
            table_df = df.iloc[data_start:].copy()
            table_df.columns = df.iloc[header_row]
            
            # Clean the dataframe
            table_df = table_df.dropna(how='all').reset_index(drop=True)
            table_df = handle_duplicate_columns(table_df)
            
            # Standardize column names
            table_df, column_mapping = standardize_column_names(table_df, is_auditor=True)
            
            # Clean and convert numeric
            table_df = clean_and_convert_numeric(table_df)
            
            # Optimize memory
            table_df, memory_info = optimize_dataframe_memory(table_df)
            
            analysis_type = 'product' if is_product_analysis else 'region' if is_region_analysis else 'general'
            
            return {
                'success': True,
                'data': table_df.head(100).to_dict('records'),
                'columns': table_df.columns.tolist(),
                'shape': table_df.shape,
                'type': 'auditor',
                'analysis_type': analysis_type,
                'header_row': header_row,
                'data_start': data_start,
                'column_mapping': column_mapping,
                'memory_info': memory_info
            }
        else:
            return 
            
    except Exception as e:
        return 

def process_total_sales_sheet(df):
    """Process total sales (last year) data sheet"""
    try:
        # Detect if this is region or product data
        sheet_text = ' '.join(df.astype(str).values.flatten()).lower()
        group_type = 'product' if any(term in sheet_text for term in ['product', 'acetic', 'auxil']) else 'region'
        
        processed_data, message = process_last_year_data(df, group_type)
        
        if processed_data is not None:
            # Optimize memory
            processed_data, memory_info = optimize_dataframe_memory(processed_data)
            
            return {
                'success': True,
                'data': processed_data.head(100).to_dict('records'),
                'columns': processed_data.columns.tolist(),
                'shape': processed_data.shape,
                'type': 'totalSales',
                'group_type': group_type,
                'message': message,
                'memory_info': memory_info
            }
        else:
            return {'error': message}
            
    except Exception as e:
        return {'error': f'Total sales processing failed: {str(e)}'}

@process_bp.route('/validate-data', methods=['POST'])
def validate_data():
    """Validate processed data"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        required_columns = data.get('required_columns', ['REGIONS', 'PRODUCT NAME'])
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        df = pd.read_excel(filepath, sheet_name=sheet_name)
        is_valid, message = validate_dataframe(df, f"{sheet_name}", required_columns)
        
        return jsonify({
            'valid': is_valid,
            'message': message,
            'shape': df.shape,
            'columns': df.columns.tolist()
        })
        
    except Exception as e:
        return jsonify({'error': f'Validation failed: {str(e)}'}), 500

@process_bp.route('/optimize-memory', methods=['POST'])
def optimize_memory():
    """Optimize dataframe memory usage"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        df = pd.read_excel(filepath, sheet_name=sheet_name)
        optimized_df, memory_info = optimize_dataframe_memory(df)
        
        return jsonify({
            'success': True,
            'memory_info': memory_info,
            'shape': optimized_df.shape,
            'data_types': optimized_df.dtypes.astype(str).to_dict()
        })
        
    except Exception as e:
        return jsonify({'error': f'Memory optimization failed: {str(e)}'}), 500

# Debug print to confirm blueprint creation
print("📝 process_bp blueprint created successfully")