from flask import Blueprint, request, jsonify, send_file
import pandas as pd
from io import BytesIO
import traceback
from process import safe_merge_dataframes, clean_and_convert_numeric

# Create the blueprint - this line is CRITICAL
data_bp = Blueprint('data', __name__)

@data_bp.route('/merge-data', methods=['POST'])
def merge_data():
    """Merge multiple datasets"""
    try:
        data = request.json
        datasets = data.get('datasets', [])
        merge_column = data.get('merge_column')
        merge_type = data.get('merge_type', 'left')
        
        if len(datasets) < 2:
            return jsonify({'error': 'At least 2 datasets required for merge'}), 400
        
        if not merge_column:
            return jsonify({'error': 'Merge column is required'}), 400
        
        # Load datasets
        dfs = []
        dataset_info = []
        
        for dataset in datasets:
            try:
                df = pd.read_excel(dataset['filepath'], sheet_name=dataset['sheet_name'])
                dfs.append(df)
                dataset_info.append({
                    'name': f"{dataset.get('name', 'Unknown')} - {dataset['sheet_name']}",
                    'shape': df.shape,
                    'columns': df.columns.tolist()
                })
            except Exception as e:
                return jsonify({'error': f"Failed to load dataset {dataset.get('name', 'Unknown')}: {str(e)}"}), 400
        
        # Validate merge column exists in all datasets
        for i, df in enumerate(dfs):
            if merge_column not in df.columns:
                return jsonify({'error': f"Column '{merge_column}' not found in dataset {dataset_info[i]['name']}"}), 400
        
        # Perform merge
        result_df = dfs[0]
        messages = []
        
        for i, df in enumerate(dfs[1:], 1):
            merge_result, result_df = safe_merge_dataframes(
                result_df, df, merge_column, merge_type
            )
            
            if 'error' in merge_result:
                return jsonify(merge_result), 400
                
            if 'messages' in merge_result:
                messages.extend(merge_result['messages'])
        
        # Clean the final result
        result_df = clean_and_convert_numeric(result_df)
        
        return jsonify({
            'success': True,
            'data': result_df.head(100).to_dict('records'),
            'columns': result_df.columns.tolist(),
            'shape': result_df.shape,
            'messages': messages,
            'dataset_info': dataset_info,
            'merge_column': merge_column,
            'merge_type': merge_type
        })
        
    except Exception as e:
        return jsonify({'error': f'Merge failed: {str(e)}', 'traceback': traceback.format_exc()}), 500

@data_bp.route('/export-csv', methods=['POST'])
def export_csv():
    """Export processed data as CSV"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        filename = data.get('filename')
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        df = pd.read_excel(filepath, sheet_name=sheet_name)
        
        # Clean the data for export
        df = clean_and_convert_numeric(df)
        
        # Create CSV in memory
        output = BytesIO()
        df.to_csv(output, index=False)
        output.seek(0)
        
        # Generate filename if not provided
        if not filename:
            filename = f'{sheet_name}_export.csv'
        elif not filename.endswith('.csv'):
            filename += '.csv'
        
        return send_file(
            output,
            mimetype='text/csv',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        return jsonify({'error': f'Export failed: {str(e)}'}), 500

@data_bp.route('/export-excel', methods=['POST'])
def export_excel():
    """Export processed data as Excel"""
    try:
        data = request.json
        datasets = data.get('datasets', [])
        filename = data.get('filename', 'export.xlsx')
        
        if not datasets:
            return jsonify({'error': 'No datasets provided for export'}), 400
        
        # Create Excel writer
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for dataset in datasets:
                try:
                    df = pd.read_excel(dataset['filepath'], sheet_name=dataset['sheet_name'])
                    df = clean_and_convert_numeric(df)
                    
                    # Use provided sheet name or generate one
                    export_sheet_name = dataset.get('export_name', dataset['sheet_name'])
                    
                    # Excel sheet names have a 31 character limit
                    if len(export_sheet_name) > 31:
                        export_sheet_name = export_sheet_name[:31]
                    
                    df.to_excel(writer, sheet_name=export_sheet_name, index=False)
                    
                except Exception as e:
                    return jsonify({'error': f"Failed to process dataset {dataset.get('name', 'Unknown')}: {str(e)}"}), 400
        
        output.seek(0)
        
        # Ensure filename has .xlsx extension
        if not filename.endswith('.xlsx'):
            filename += '.xlsx'
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        return jsonify({'error': f'Excel export failed: {str(e)}'}), 500

@data_bp.route('/get-sheet-info', methods=['POST'])
def get_sheet_info():
    """Get information about a specific sheet"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        df = pd.read_excel(filepath, sheet_name=sheet_name, nrows=0)  # Just get columns
        full_df = pd.read_excel(filepath, sheet_name=sheet_name)
        
        # Get basic statistics
        numeric_columns = full_df.select_dtypes(include=['number']).columns.tolist()
        categorical_columns = full_df.select_dtypes(include=['object', 'category']).columns.tolist()
        
        stats = {}
        if numeric_columns:
            stats['numeric'] = full_df[numeric_columns].describe().to_dict()
        
        return jsonify({
            'success': True,
            'shape': full_df.shape,
            'columns': full_df.columns.tolist(),
            'data_types': full_df.dtypes.astype(str).to_dict(),
            'numeric_columns': numeric_columns,
            'categorical_columns': categorical_columns,
            'memory_usage': full_df.memory_usage(deep=True).sum(),
            'null_counts': full_df.isnull().sum().to_dict(),
            'statistics': stats
        })
        
    except Exception as e:
        return jsonify({'error': f'Failed to get sheet info: {str(e)}'}), 500

@data_bp.route('/preview-data', methods=['POST'])
def preview_data():
    """Get a preview of the data"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        num_rows = data.get('num_rows', 50)
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        df = pd.read_excel(filepath, sheet_name=sheet_name)
        
        # Get preview data
        preview_df = df.head(num_rows)
        
        return jsonify({
            'success': True,
            'data': preview_df.to_dict('records'),
            'columns': df.columns.tolist(),
            'total_rows': len(df),
            'preview_rows': len(preview_df),
            'shape': df.shape
        })
        
    except Exception as e:
        return jsonify({'error': f'Preview failed: {str(e)}'}), 500

@data_bp.route('/search-data', methods=['POST'])
def search_data():
    """Search for specific data in the sheet"""
    try:
        data = request.json
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        search_term = data.get('search_term')
        search_column = data.get('search_column')
        
        if not filepath or not sheet_name or not search_term:
            return jsonify({'error': 'Missing required parameters'}), 400
        
        df = pd.read_excel(filepath, sheet_name=sheet_name)
        
        if search_column and search_column in df.columns:
            # Search in specific column
            mask = df[search_column].astype(str).str.contains(search_term, case=False, na=False)
        else:
            # Search in all columns
            mask = df.astype(str).apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)
        
        results = df[mask]
        
        return jsonify({
            'success': True,
            'data': results.to_dict('records'),
            'columns': df.columns.tolist(),
            'total_matches': len(results),
            'search_term': search_term,
            'search_column': search_column
        })
        
    except Exception as e:
        return jsonify({'error': f'Search failed: {str(e)}'}), 500

# Debug print to confirm blueprint creation
print("📝 data_bp blueprint created successfully")