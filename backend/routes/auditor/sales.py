from flask import Blueprint, request, jsonify
import pandas as pd
import traceback
import numpy as np
from datetime import datetime
from utils.auditor.data_processor import DataProcessor
from process import process_budget_data,process_last_year_data,  handle_duplicate_columns

# Create the blueprint
sales1_bp = Blueprint('sales1', __name__)

def clean_dataframe_for_json(df):
    """Clean DataFrame to make it JSON serializable"""
    df_clean = df.copy()
    
    for col in df_clean.columns:
        # Handle datetime columns
        if df_clean[col].dtype == 'datetime64[ns]':
            # Convert datetime to string, handling NaT values
            df_clean[col] = df_clean[col].dt.strftime('%Y-%m-%d %H:%M:%S').fillna('')
        
        # Handle timedelta columns
        elif df_clean[col].dtype == 'timedelta64[ns]':
            # Convert timedelta to string representation
            df_clean[col] = df_clean[col].astype(str).replace('NaT', '')
        
        # Handle numeric columns with NaN
        elif df_clean[col].dtype in ['int64', 'float64']:
            df_clean[col] = df_clean[col].fillna(0)
            df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)
        
        # Handle object columns that might contain datetime-like objects
        elif df_clean[col].dtype == 'object':
            # Check if any values are datetime-like
            sample_values = df_clean[col].dropna().head(5)
            is_datetime_like = False
            
            for val in sample_values:
                if isinstance(val, (datetime, pd.Timestamp)) or pd.api.types.is_datetime64_any_dtype(pd.Series([val])):
                    is_datetime_like = True
                    break
            
            if is_datetime_like:
                # Convert datetime-like objects to strings
                df_clean[col] = df_clean[col].apply(
                    lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notna(x) and hasattr(x, 'strftime') else str(x) if pd.notna(x) else ''
                )
            else:
                # Fill NaN values in object columns
                df_clean[col] = df_clean[col].fillna('')
        
        # Handle any remaining NaT or NaN values
        df_clean[col] = df_clean[col].replace({pd.NaT: '', np.nan: ''})
    
    return df_clean

def fallback_process_budget_data(df, group_type='region'):
    """Simple fallback for budget data - just clean and return"""
    try:
        df_processed = clean_dataframe_for_json(df.copy())
        message = f"Budget data processed (fallback mode) - {df_processed.shape[0]} rows"
        return df_processed, message
    except Exception as e:
        return None, f"Fallback budget processing failed: {str(e)}"

def fallback_process_last_year_data(df, group_type='region'):
    """Simple fallback for last year data - just clean and return"""
    try:
        df_processed = clean_dataframe_for_json(df.copy())
        message = f"Last year data processed (fallback mode) - {df_processed.shape[0]} rows"
        return df_processed, message
    except Exception as e:
        return None, f"Fallback last year processing failed: {str(e)}"

@sales1_bp.route('/process-sales-sheet', methods=['POST'])
def process_sales_sheet():
    """Process sales sheet data"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        # Read Excel with header in first row (row index 0)
        df_sales = pd.read_excel(filepath, sheet_name=sheet_name, header=0)
        df_sales = handle_duplicate_columns(df_sales)
        
        # Clean the data
        df_sales.columns = df_sales.columns.str.strip()
        df_sales = df_sales.dropna(how='all').reset_index(drop=True)
        
        # Clean DataFrame for JSON serialization
        df_sales_clean = clean_dataframe_for_json(df_sales)
        
        return jsonify({
            'success': True,
            'data': df_sales_clean.to_dict('records'),
            'columns': df_sales_clean.columns.tolist(),
            'shape': [int(df_sales_clean.shape[0]), int(df_sales_clean.shape[1])],
            'sheet_name': sheet_name,
            'data_type': 'sales',
            'message': f'Sales sheet "{sheet_name}" processed successfully'
        })
        
    except Exception as e:
        return jsonify({
            'error': f'Error processing sales sheet: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@sales1_bp.route('/process-budget-sheet', methods=['POST'])
def process_budget_sheet():
    """Process budget sheet data"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        group_type = data.get('group_type', 'region')
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        # Read Excel with header in first row (row index 0)
        df_budget = pd.read_excel(filepath, sheet_name=sheet_name, header=0)
        df_budget.columns = df_budget.columns.str.strip()
        df_budget = df_budget.dropna(how='all').reset_index(drop=True)
        df_budget = handle_duplicate_columns(df_budget)
        
        # Try original processing first
        try:
            budget_data, message = process_budget_data(df_budget, group_type=group_type)
            
            if budget_data is not None:
                # Clean DataFrame for JSON serialization
                budget_data_clean = clean_dataframe_for_json(budget_data)
                df_budget_clean = clean_dataframe_for_json(df_budget)
                
                return jsonify({
                    'success': True,
                    'data': budget_data_clean.to_dict('records'),
                    'columns': budget_data_clean.columns.tolist(),
                    'shape': [int(budget_data_clean.shape[0]), int(budget_data_clean.shape[1])],
                    'sheet_name': sheet_name,
                    'data_type': 'budget',
                    'group_type': group_type,
                    'message': message,
                    'raw_data': df_budget_clean.to_dict('records'),
                    'raw_columns': df_budget_clean.columns.tolist()
                })
            else:
                # Original processing returned None, try fallback
                budget_data, message = fallback_process_budget_data(df_budget, group_type)
                
                if budget_data is not None:
                    return jsonify({
                        'success': True,
                        'data': budget_data.to_dict('records'),
                        'columns': budget_data.columns.tolist(),
                        'shape': [int(budget_data.shape[0]), int(budget_data.shape[1])],
                        'sheet_name': sheet_name,
                        'data_type': 'budget',
                        'group_type': group_type,
                        'message': f"Fallback: {message}",
                        'raw_data': budget_data.to_dict('records'),
                        'raw_columns': budget_data.columns.tolist()
                    })
                else:
                    df_budget_clean = clean_dataframe_for_json(df_budget)
                    return jsonify({
                        'error': message,
                        'raw_data': df_budget_clean.to_dict('records'),
                        'raw_columns': df_budget_clean.columns.tolist()
                    }), 400
                
        except Exception as processing_error:
            # Original processing failed, try fallback
            print(f"Original budget processing failed: {str(processing_error)}")
            budget_data, message = fallback_process_budget_data(df_budget, group_type)
            
            if budget_data is not None:
                return jsonify({
                    'success': True,
                    'data': budget_data.to_dict('records'),
                    'columns': budget_data.columns.tolist(),
                    'shape': [int(budget_data.shape[0]), int(budget_data.shape[1])],
                    'sheet_name': sheet_name,
                    'data_type': 'budget',
                    'group_type': group_type,
                    'message': f"Fallback: {message}",
                    'raw_data': budget_data.to_dict('records'),
                    'raw_columns': budget_data.columns.tolist()
                })
            else:
                return jsonify({
                    'error': f'Both original and fallback processing failed. Original error: {str(processing_error)}, Fallback error: {message}',
                    'traceback': traceback.format_exc()
                }), 500
        
    except Exception as e:
        return jsonify({
            'error': f'Error processing budget sheet: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@sales1_bp.route('/process-last-year-sheet', methods=['POST'])
def process_last_year_sheet():
    """Process last year sheet data"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        group_type = data.get('group_type', 'region')
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        # Read Excel with header in first row (row index 0)
        df_last_year = pd.read_excel(filepath, sheet_name=sheet_name, header=0)
        df_last_year.columns = df_last_year.columns.str.strip()
        df_last_year = df_last_year.dropna(how='all').reset_index(drop=True)
        df_last_year = handle_duplicate_columns(df_last_year)
        
        # Try original processing first
        try:
            last_year_data, message = process_last_year_data(df_last_year, group_type=group_type)
            
            if last_year_data is not None:
                # Clean DataFrame for JSON serialization
                last_year_data_clean = clean_dataframe_for_json(last_year_data)
                df_last_year_clean = clean_dataframe_for_json(df_last_year)
                
                return jsonify({
                    'success': True,
                    'data': last_year_data_clean.to_dict('records'),
                    'columns': last_year_data_clean.columns.tolist(),
                    'shape': [int(last_year_data_clean.shape[0]), int(last_year_data_clean.shape[1])],
                    'sheet_name': sheet_name,
                    'data_type': 'last_year',
                    'group_type': group_type,
                    'message': message,
                    'raw_data': df_last_year_clean.to_dict('records'),
                    'raw_columns': df_last_year_clean.columns.tolist()
                })
            else:
                # Original processing returned None, try fallback
                last_year_data, message = fallback_process_last_year_data(df_last_year, group_type)
                
                if last_year_data is not None:
                    return jsonify({
                        'success': True,
                        'data': last_year_data.to_dict('records'),
                        'columns': last_year_data.columns.tolist(),
                        'shape': [int(last_year_data.shape[0]), int(last_year_data.shape[1])],
                        'sheet_name': sheet_name,
                        'data_type': 'last_year',
                        'group_type': group_type,
                        'message': f"Fallback: {message}",
                        'raw_data': last_year_data.to_dict('records'),
                        'raw_columns': last_year_data.columns.tolist()
                    })
                else:
                    df_last_year_clean = clean_dataframe_for_json(df_last_year)
                    return jsonify({
                        'error': message,
                        'raw_data': df_last_year_clean.to_dict('records'),
                        'raw_columns': df_last_year_clean.columns.tolist()
                    }), 400
                
        except Exception as processing_error:
            # Original processing failed, try fallback
            print(f"Original last year processing failed: {str(processing_error)}")
            last_year_data, message = fallback_process_last_year_data(df_last_year, group_type)
            
            if last_year_data is not None:
                return jsonify({
                    'success': True,
                    'data': last_year_data.to_dict('records'),
                    'columns': last_year_data.columns.tolist(),
                    'shape': [int(last_year_data.shape[0]), int(last_year_data.shape[1])],
                    'sheet_name': sheet_name,
                    'data_type': 'last_year',
                    'group_type': group_type,
                    'message': f"Fallback: {message}",
                    'raw_data': last_year_data.to_dict('records'),
                    'raw_columns': last_year_data.columns.tolist()
                })
            else:
                return jsonify({
                    'error': f'Both original and fallback processing failed. Original error: {str(processing_error)}, Fallback error: {message}',
                    'traceback': traceback.format_exc()
                }), 500
        
    except Exception as e:
        return jsonify({
            'error': f'Error processing last year sheet: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@sales1_bp.route('/get-sheet-preview', methods=['POST'])
def get_sheet_preview():
    """Get a preview of any sheet without processing"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        preview_rows = data.get('preview_rows', 10)
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        # Read Excel sheet
        df = pd.read_excel(filepath, sheet_name=sheet_name, header=0)
        df = handle_duplicate_columns(df)
        
        # Clean basic formatting
        df.columns = df.columns.str.strip()
        df = df.dropna(how='all').reset_index(drop=True)
        
        # Get preview data
        preview_df = df.head(preview_rows)
        
        # Clean DataFrame for JSON serialization
        preview_df_clean = clean_dataframe_for_json(preview_df)
        
        return jsonify({
            'success': True,
            'data': preview_df_clean.to_dict('records'),
            'columns': preview_df_clean.columns.tolist(),
            'total_rows': int(len(df)),
            'preview_rows': int(len(preview_df_clean)),
            'shape': [int(df.shape[0]), int(df.shape[1])],
            'sheet_name': sheet_name
        })
        
    except Exception as e:
        return jsonify({
            'error': f'Error getting sheet preview: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@sales1_bp.route('/export-sales-data', methods=['POST'])
def export_sales_data():
    """Export processed sales data"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        export_format = data.get('format', 'csv')
        data_type = data.get('data_type', 'sales')
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        # Process the data based on type
        if data_type == 'sales':
            response = process_sales_sheet()
        elif data_type == 'budget':
            response = process_budget_sheet()
        elif data_type == 'last_year':
            response = process_last_year_sheet()
        else:
            return jsonify({'error': 'Invalid data_type'}), 400
        
        if isinstance(response, tuple):
            result = response[0].get_json()
        else:
            result = response.get_json()
        
        if not result.get('success'):
            return jsonify(result), 400
        
        # Create DataFrame from processed data
        df = pd.DataFrame(result['data'])
        
        # Generate filename
        filename = f"{data_type}_{sheet_name}_{export_format}"
        
        if export_format == 'csv':
            from io import StringIO
            output = StringIO()
            df.to_csv(output, index=False)
            csv_data = output.getvalue()
            
            return jsonify({
                'success': True,
                'data': csv_data,
                'filename': f"{filename}.csv",
                'format': 'csv'
            })
        else:
            return jsonify({'error': 'Only CSV export is currently supported'}), 400
            
    except Exception as e:
        return jsonify({
            'error': f'Error exporting data: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

# Debug print to confirm blueprint creation
print("📝 sales_bp blueprint created successfully")