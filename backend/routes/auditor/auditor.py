from flask import Blueprint, request, jsonify, session
import pandas as pd
import re
import traceback
import logging
from utils.auditor.data_processor import DataProcessor

# Create the blueprint
auditor_bp = Blueprint('auditor', __name__)

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def detect_analysis_type(sheet_name):
    """Detect the type of analysis based on sheet name"""
    try:
        sheet_lower = sheet_name.lower()
        
        is_region_analysis = 'region' in sheet_lower
        is_product_analysis = 'product' in sheet_lower or 'ts-pw' in sheet_lower or 'ero-pw' in sheet_lower
        is_sales_analysis_month_wise = bool(re.search(r'sales\s*analysis\s*month\s*wise', sheet_lower))
        
        logger.debug(f"Analysis type detection for '{sheet_name}': region={is_region_analysis}, product={is_product_analysis}, sales_month_wise={is_sales_analysis_month_wise}")
        
        return {
            'is_region_analysis': is_region_analysis,
            'is_product_analysis': is_product_analysis,
            'is_sales_analysis_month_wise': is_sales_analysis_month_wise
        }
    except Exception as e:
        logger.error(f"Error in detect_analysis_type: {str(e)}")
        return {
            'is_region_analysis': False,
            'is_product_analysis': False,
            'is_sales_analysis_month_wise': False
        }

def get_session_key(filepath, sheet_name, table_type):
    """Generate a unique session key for each table"""
    try:
        # Create a clean key from filepath and sheet name
        file_key = filepath.split('/')[-1].replace('.xlsx', '').replace('.xls', '')
        # Clean sheet name for session key
        clean_sheet_name = re.sub(r'[^\w\-_]', '_', sheet_name)
        session_key = f"auditor_{file_key}_{clean_sheet_name}_{table_type}"
        logger.debug(f"Generated session key: {session_key}")
        return session_key
    except Exception as e:
        logger.error(f"Error generating session key: {str(e)}")
        return f"auditor_default_{table_type}"

@auditor_bp.route('/process-auditor-auto', methods=['POST'])
def process_auditor_auto():
    """Auto-process auditor table and store both tables in session"""
    logger.info("Starting process-auditor-auto endpoint")
    
    try:
        # Get request data
        data = request.json
        if not data:
            logger.error("No JSON data received in request")
            return jsonify({'error': 'No JSON data received'}), 400
            
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        table_choice = data.get('table_choice', 'Table 2: SALES in Value')
        force_refresh = data.get('force_refresh', False)
        
        logger.info(f"Request parameters: filepath={filepath}, sheet_name={sheet_name}, table_choice={table_choice}, force_refresh={force_refresh}")
        
        if not filepath or not sheet_name:
            logger.error(f"Missing required parameters: filepath={filepath}, sheet_name={sheet_name}")
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        # Check if both tables are already in session (unless force refresh)
        table1_key = get_session_key(filepath, sheet_name, 'Table1')
        table2_key = get_session_key(filepath, sheet_name, 'Table2')
        
        logger.debug(f"Session keys: table1={table1_key}, table2={table2_key}")
        logger.debug(f"Session contains table1: {'table1_key' in session}, table2: {'table2_key' in session}")
        
        if not force_refresh and table1_key in session and table2_key in session:
            logger.info("Both tables found in session, returning cached data")
            # Return the requested table from session
            if table_choice == "Table 1: SALES in MT/Tonage" and table1_key in session:
                stored_data = session[table1_key]
                logger.info("Returning cached Table 1 data")
                return jsonify(stored_data)
            elif table_choice == "Table 2: SALES in Value" and table2_key in session:
                stored_data = session[table2_key]
                logger.info("Returning cached Table 2 data")
                return jsonify(stored_data)
        
        # Initialize DataProcessor
        try:
            logger.info("Initializing DataProcessor")
            processor = DataProcessor()
        except Exception as e:
            logger.error(f"Failed to initialize DataProcessor: {str(e)}")
            return jsonify({
                'error': f'Failed to initialize DataProcessor: {str(e)}',
                'traceback': traceback.format_exc()
            }), 500
        
        # Read the Excel sheet
        try:
            logger.info(f"Reading Excel file: {filepath}, sheet: {sheet_name}")
            df_sheet = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
            logger.info(f"Excel sheet loaded successfully. Shape: {df_sheet.shape}")
            logger.debug(f"First few rows:\n{df_sheet.head()}")
        except FileNotFoundError:
            logger.error(f"File not found: {filepath}")
            return jsonify({'error': f'File not found: {filepath}'}), 404
        except ValueError as e:
            logger.error(f"Invalid sheet name '{sheet_name}': {str(e)}")
            return jsonify({'error': f'Invalid sheet name "{sheet_name}": {str(e)}'}), 400
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}")
            return jsonify({
                'error': f'Error reading Excel file: {str(e)}',
                'traceback': traceback.format_exc()
            }), 500
        
        # Detect analysis type
        try:
            logger.info("Detecting analysis type")
            analysis_info = detect_analysis_type(sheet_name)
            logger.info(f"Analysis type detected: {analysis_info}")
        except Exception as e:
            logger.error(f"Error detecting analysis type: {str(e)}")
            analysis_info = {
                'is_region_analysis': False,
                'is_product_analysis': False,
                'is_sales_analysis_month_wise': False
            }
        
        # Define possible headers for both tables
        table1_possible_headers = [
            "SALES in MT", "SALES IN MT", "Sales in MT", "SALES IN TONNAGE", "SALES IN TON",
            "Tonnage", "TONNAGE", "Tonnage Sales", "Sales Tonnage", "Metric Tons", "MT Sales"
        ]
        
        table2_possible_headers = [
            "SALES in Value", "SALES IN VALUE", "Sales in Value", "SALES IN RS", "VALUE SALES",
            "Value", "VALUE", "Sales Value"
        ]
        
        logger.debug(f"Searching for table1 headers: {table1_possible_headers}")
        logger.debug(f"Searching for table2 headers: {table2_possible_headers}")
        
        try:
            # Extract both tables using DataProcessor
            logger.info("Extracting Table 1 (MT/Tonnage)")
            try:
                idx1, data_start1 = processor.extract_tables(
                    df_sheet, 
                    table1_possible_headers, 
                    is_product_analysis=analysis_info['is_product_analysis']
                )
                logger.info(f"Table 1 extraction result: idx={idx1}, data_start={data_start1}")
            except Exception as e:
                logger.error(f"Error extracting Table 1: {str(e)}")
                idx1, data_start1 = None, None
            
            logger.info("Extracting Table 2 (Value)")
            try:
                idx2, data_start2 = processor.extract_tables(
                    df_sheet, 
                    table2_possible_headers, 
                    is_product_analysis=analysis_info['is_product_analysis']
                )
                logger.info(f"Table 2 extraction result: idx={idx2}, data_start={data_start2}")
            except Exception as e:
                logger.error(f"Error extracting Table 2: {str(e)}")
                idx2, data_start2 = None, None
            
            # Process and store Table 1 if found
            table1_data = None
            if idx1 is not None:
                logger.info(f"Processing Table 1 starting from row {data_start1}")
                try:
                    table_end1 = idx2 if idx2 is not None and idx2 > idx1 else len(df_sheet)
                    logger.debug(f"Table 1 data range: {data_start1}:{table_end1}")
                    
                    table1_df = df_sheet.iloc[data_start1:table_end1].dropna(how='all').reset_index(drop=True)
                    logger.info(f"Table 1 dataframe shape after cleanup: {table1_df.shape}")
                    
                    if not table1_df.empty:
                        # Set headers and process data
                        original_headers1 = df_sheet.iloc[idx1].tolist()
                        logger.debug(f"Table 1 original headers: {original_headers1}")
                        
                        table1_df.columns = [str(col) for col in original_headers1]
                        
                        # Process with DataProcessor
                        new_column_names1 = processor.rename_columns(table1_df.columns.tolist())
                        logger.debug(f"Table 1 renamed columns: {new_column_names1}")
                        
                        table1_df.columns = new_column_names1
                        table1_df = processor.handle_duplicate_columns(table1_df)
                        table1_df = processor.clean_and_convert_numeric(table1_df)
                        
                        # Get column information
                        act_columns1 = [col for col in table1_df.columns if col.startswith('Act-')]
                        gr_columns1 = [col for col in table1_df.columns if col.startswith('Gr-')]
                        ach_columns1 = [col for col in table1_df.columns if col.startswith('Ach-')]
                        budget_columns1 = [col for col in table1_df.columns if col.startswith('Budget-')]
                        ly_columns1 = [col for col in table1_df.columns if col.startswith('LY-')]
                        
                        logger.info(f"Table 1 column analysis: Act={len(act_columns1)}, Gr={len(gr_columns1)}, Ach={len(ach_columns1)}, Budget={len(budget_columns1)}, LY={len(ly_columns1)}")
                        
                        # Prepare Table 1 data for session storage
                        table1_data = {
                            'success': True,
                            'table_name': 'Table 1: SALES in MT/Tonage',
                            'data': table1_df.to_dict('records'),
                            'columns': table1_df.columns.tolist(),
                            'shape': [int(table1_df.shape[0]), int(table1_df.shape[1])],
                            'analysis_type': analysis_info,
                            'column_info': {
                                'act_columns': act_columns1,
                                'gr_columns': gr_columns1,
                                'ach_columns': ach_columns1,
                                'budget_columns': budget_columns1,
                                'ly_columns': ly_columns1,
                                'total_columns': int(len(table1_df.columns))
                            },
                            'original_headers': [str(h) for h in original_headers1],
                            'processing_info': {
                                'original_column_count': len(original_headers1),
                                'final_column_count': len(table1_df.columns),
                                'processing_successful': True
                            }
                        }
                        
                        # Store in session
                        session[table1_key] = table1_data
                        logger.info(f"Table 1 stored in session with key: {table1_key}")
                    else:
                        logger.warning("Table 1 dataframe is empty after processing")
                        
                except Exception as e:
                    logger.error(f"Error processing Table 1: {str(e)}")
                    logger.error(f"Table 1 processing traceback: {traceback.format_exc()}")
            else:
                logger.info("Table 1 not found in sheet")
            
            # Process and store Table 2 if found
            table2_data = None
            if idx2 is not None:
                logger.info(f"Processing Table 2 starting from row {data_start2}")
                try:
                    table2_df = df_sheet.iloc[data_start2:len(df_sheet)].dropna(how='all').reset_index(drop=True)
                    logger.info(f"Table 2 dataframe shape after cleanup: {table2_df.shape}")
                    
                    if not table2_df.empty:
                        # Set headers and process data
                        original_headers2 = df_sheet.iloc[idx2].tolist()
                        logger.debug(f"Table 2 original headers: {original_headers2}")
                        
                        table2_df.columns = [str(col) for col in original_headers2]
                        
                        # Process with DataProcessor
                        new_column_names2 = processor.rename_columns(table2_df.columns.tolist())
                        logger.debug(f"Table 2 renamed columns: {new_column_names2}")
                        
                        table2_df.columns = new_column_names2
                        table2_df = processor.handle_duplicate_columns(table2_df)
                        table2_df = processor.clean_and_convert_numeric(table2_df)
                        
                        # Get column information
                        act_columns2 = [col for col in table2_df.columns if col.startswith('Act-')]
                        gr_columns2 = [col for col in table2_df.columns if col.startswith('Gr-')]
                        ach_columns2 = [col for col in table2_df.columns if col.startswith('Ach-')]
                        budget_columns2 = [col for col in table2_df.columns if col.startswith('Budget-')]
                        ly_columns2 = [col for col in table2_df.columns if col.startswith('LY-')]
                        
                        logger.info(f"Table 2 column analysis: Act={len(act_columns2)}, Gr={len(gr_columns2)}, Ach={len(ach_columns2)}, Budget={len(budget_columns2)}, LY={len(ly_columns2)}")
                        
                        # Prepare Table 2 data for session storage
                        table2_data = {
                            'success': True,
                            'table_name': 'Table 2: SALES in Value',
                            'data': table2_df.to_dict('records'),
                            'columns': table2_df.columns.tolist(),
                            'shape': [int(table2_df.shape[0]), int(table2_df.shape[1])],
                            'analysis_type': analysis_info,
                            'column_info': {
                                'act_columns': act_columns2,
                                'gr_columns': gr_columns2,
                                'ach_columns': ach_columns2,
                                'budget_columns': budget_columns2,
                                'ly_columns': ly_columns2,
                                'total_columns': int(len(table2_df.columns))
                            },
                            'original_headers': [str(h) for h in original_headers2],
                            'processing_info': {
                                'original_column_count': len(original_headers2),
                                'final_column_count': len(table2_df.columns),
                                'processing_successful': True
                            }
                        }
                        
                        # Store in session
                        session[table2_key] = table2_data
                        logger.info(f"Table 2 stored in session with key: {table2_key}")
                    else:
                        logger.warning("Table 2 dataframe is empty after processing")
                        
                except Exception as e:
                    logger.error(f"Error processing Table 2: {str(e)}")
                    logger.error(f"Table 2 processing traceback: {traceback.format_exc()}")
            else:
                logger.info("Table 2 not found in sheet")
            
            # Return the requested table
            logger.info(f"Determining which table to return for choice: {table_choice}")
            
            if table_choice == "Table 1: SALES in MT/Tonage" and table1_data:
                logger.info("Returning Table 1 data")
                return jsonify(table1_data)
            elif table_choice == "Table 2: SALES in Value" and table2_data:
                logger.info("Returning Table 2 data")
                return jsonify(table2_data)
            else:
                # Fallback: return any available table
                logger.info("Using fallback logic to return available table")
                if table2_data:
                    logger.info("Returning Table 2 data as fallback")
                    return jsonify(table2_data)
                elif table1_data:
                    logger.info("Returning Table 1 data as fallback")
                    return jsonify(table1_data)
                else:
                    logger.error("No valid tables found in the sheet")
                    # Provide debug information
                    debug_info = {
                        'tried_headers_table1': table1_possible_headers,
                        'tried_headers_table2': table2_possible_headers,
                        'sheet_shape': df_sheet.shape,
                        'analysis_type': analysis_info,
                        'extraction_results': {
                            'table1_idx': idx1,
                            'table1_data_start': data_start1,
                            'table2_idx': idx2,
                            'table2_data_start': data_start2
                        }
                    }
                    
                    # Add sheet preview (first 10 rows) for debugging
                    try:
                        debug_info['sheet_preview'] = df_sheet.head(10).to_dict('records')
                    except Exception as e:
                        logger.error(f"Error creating sheet preview: {str(e)}")
                        debug_info['sheet_preview_error'] = str(e)
                    
                    return jsonify({
                        'error': 'No valid tables found in the sheet',
                        'debug_info': debug_info
                    }), 400
            
        except Exception as e:
            logger.error(f"Error in table extraction and processing: {str(e)}")
            logger.error(f"Full traceback: {traceback.format_exc()}")
            return jsonify({
                'error': f'Error extracting tables: {str(e)}',
                'traceback': traceback.format_exc(),
                'debug_info': {
                    'filepath': filepath,
                    'sheet_name': sheet_name,
                    'sheet_shape': df_sheet.shape if 'df_sheet' in locals() else 'N/A',
                    'analysis_type': analysis_info if 'analysis_info' in locals() else 'N/A'
                }
            }), 500
        
    except Exception as e:
        logger.error(f"Unexpected error in process-auditor-auto: {str(e)}")
        logger.error(f"Full traceback: {traceback.format_exc()}")
        return jsonify({
            'error': f'Error processing auditor file: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@auditor_bp.route('/get-available-tables', methods=['POST'])
def get_available_tables():
    """Get list of available tables in the auditor sheet and check session cache"""
    logger.info("Starting get-available-tables endpoint")
    
    try:
        data = request.json
        if not data:
            logger.error("No JSON data received")
            return jsonify({'error': 'No JSON data received'}), 400
            
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        
        logger.info(f"Parameters: filepath={filepath}, sheet_name={sheet_name}")
        
        if not filepath or not sheet_name:
            logger.error(f"Missing required parameters")
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        # Check session first
        table1_key = get_session_key(filepath, sheet_name, 'Table1')
        table2_key = get_session_key(filepath, sheet_name, 'Table2')
        
        available_tables = []
        table_info = {}
        
        # Check if tables are in session
        if table1_key in session:
            available_tables.append("Table 1: SALES in MT/Tonage")
            table_info["Table 1: SALES in MT/Tonage"] = {
                'source': 'session_cache',
                'estimated_rows': session[table1_key]['shape'][0]
            }
            logger.info("Found Table 1 in session cache")
        
        if table2_key in session:
            available_tables.append("Table 2: SALES in Value")
            table_info["Table 2: SALES in Value"] = {
                'source': 'session_cache',
                'estimated_rows': session[table2_key]['shape'][0]
            }
            logger.info("Found Table 2 in session cache")
        
        # If both tables are in session, return cached info
        if len(available_tables) == 2:
            logger.info("Both tables found in cache, returning cached information")
            return jsonify({
                'success': True,
                'available_tables': available_tables,
                'table_info': table_info,
                'analysis_type': session[table2_key]['analysis_type'] if table2_key in session else session[table1_key]['analysis_type'],
                'default_table': "Table 2: SALES in Value" if "Table 2: SALES in Value" in available_tables else available_tables[0],
                'from_cache': True
            })
        
        # If not in session, process the file to detect tables
        logger.info("Tables not in cache, processing file to detect available tables")
        
        # Initialize DataProcessor
        try:
            processor = DataProcessor()
        except Exception as e:
            logger.error(f"Failed to initialize DataProcessor: {str(e)}")
            return jsonify({'error': f'Failed to initialize DataProcessor: {str(e)}'}), 500
        
        # Read the Excel sheet
        try:
            df_sheet = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
            logger.info(f"Sheet loaded for detection. Shape: {df_sheet.shape}")
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}")
            return jsonify({'error': f'Error reading Excel file: {str(e)}'}), 500
        
        # Detect analysis type
        analysis_info = detect_analysis_type(sheet_name)
        
        # Define possible headers
        table1_possible_headers = [
            "SALES in MT", "SALES IN MT", "Sales in MT", "SALES IN TONNAGE", "SALES IN TON",
            "Tonnage", "TONNAGE", "Tonnage Sales", "Sales Tonnage", "Metric Tons", "MT Sales"
        ]
        
        table2_possible_headers = [
            "SALES in Value", "SALES IN VALUE", "Sales in Value", "SALES IN RS", "VALUE SALES",
            "Value", "VALUE", "Sales Value"
        ]
        
        available_tables = []
        table_info = {}
        
        # Check for Table 1 (MT/Tonnage) using DataProcessor
        try:
            idx1, data_start1 = processor.extract_tables(
                df_sheet, 
                table1_possible_headers, 
                is_product_analysis=analysis_info['is_product_analysis']
            )
            if idx1 is not None:
                available_tables.append("Table 1: SALES in MT/Tonage")
                table_info["Table 1: SALES in MT/Tonage"] = {
                    'header_row': idx1,
                    'data_start': data_start1,
                    'estimated_rows': len(df_sheet) - data_start1,
                    'source': 'file_detection'
                }
                logger.info(f"Table 1 detected at row {idx1}")
        except Exception as e:
            logger.error(f"Could not find Table 1: {e}")
        
        # Check for Table 2 (Value) using DataProcessor
        try:
            idx2, data_start2 = processor.extract_tables(
                df_sheet, 
                table2_possible_headers, 
                is_product_analysis=analysis_info['is_product_analysis']
            )
            if idx2 is not None:
                available_tables.append("Table 2: SALES in Value")
                table_info["Table 2: SALES in Value"] = {
                    'header_row': idx2,
                    'data_start': data_start2,
                    'estimated_rows': len(df_sheet) - data_start2,
                    'source': 'file_detection'
                }
                logger.info(f"Table 2 detected at row {idx2}")
        except Exception as e:
            logger.error(f"Could not find Table 2: {e}")
        
        logger.info(f"Detection complete. Available tables: {available_tables}")
        
        return jsonify({
            'success': True,
            'available_tables': available_tables,
            'table_info': table_info,
            'analysis_type': analysis_info,
            'default_table': "Table 2: SALES in Value" if "Table 2: SALES in Value" in available_tables else (available_tables[0] if available_tables else None),
            'sheet_info': {
                'total_rows': len(df_sheet),
                'total_columns': len(df_sheet.columns)
            },
            'from_cache': False
        })
        
    except Exception as e:
        logger.error(f"Unexpected error in get-available-tables: {str(e)}")
        logger.error(f"Full traceback: {traceback.format_exc()}")
        return jsonify({
            'error': f'Error getting available tables: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@auditor_bp.route('/get-cached-table', methods=['POST'])
def get_cached_table():
    """Get a specific table from session cache"""
    logger.info("Starting get-cached-table endpoint")
    
    try:
        data = request.json
        if not data:
            return jsonify({'error': 'No JSON data received'}), 400
            
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        table_choice = data.get('table_choice')
        
        logger.info(f"Parameters: filepath={filepath}, sheet_name={sheet_name}, table_choice={table_choice}")
        
        if not all([filepath, sheet_name, table_choice]):
            return jsonify({'error': 'Missing required parameters'}), 400
        
        # Determine which table to get
        if table_choice == "Table 1: SALES in MT/Tonage":
            table_key = get_session_key(filepath, sheet_name, 'Table1')
        elif table_choice == "Table 2: SALES in Value":
            table_key = get_session_key(filepath, sheet_name, 'Table2')
        else:
            return jsonify({'error': 'Invalid table choice'}), 400
        
        # Check if table exists in session
        if table_key in session:
            logger.info(f"Returning cached table: {table_choice}")
            return jsonify(session[table_key])
        else:
            logger.warning(f"Table not found in cache: {table_key}")
            return jsonify({'error': 'Table not found in cache. Please refresh.'}), 404
            
    except Exception as e:
        logger.error(f"Error in get-cached-table: {str(e)}")
        return jsonify({
            'error': f'Error getting cached table: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@auditor_bp.route('/clear-session-cache', methods=['POST'])
def clear_session_cache():
    """Clear session cache for specific file/sheet or all"""
    logger.info("Starting clear-session-cache endpoint")
    
    try:
        data = request.json
        filepath = data.get('filepath') if data else None
        sheet_name = data.get('sheet_name') if data else None
        
        if filepath and sheet_name:
            # Clear specific file/sheet cache
            table1_key = get_session_key(filepath, sheet_name, 'Table1')
            table2_key = get_session_key(filepath, sheet_name, 'Table2')
            
            removed_count = 0
            if table1_key in session:
                session.pop(table1_key, None)
                removed_count += 1
            if table2_key in session:
                session.pop(table2_key, None)
                removed_count += 1
            
            logger.info(f"Cleared {removed_count} tables for sheet '{sheet_name}'")
            
            return jsonify({
                'success': True,
                'message': f'Cache cleared for {sheet_name} ({removed_count} tables removed)'
            })
        else:
            # Clear all auditor cache
            keys_to_remove = [key for key in session.keys() if key.startswith('auditor_')]
            for key in keys_to_remove:
                session.pop(key, None)
            
            logger.info(f"Cleared {len(keys_to_remove)} cached tables")
            
            return jsonify({
                'success': True,
                'message': f'Cleared {len(keys_to_remove)} cached tables'
            })
            
    except Exception as e:
        logger.error(f"Error clearing cache: {str(e)}")
        return jsonify({
            'error': f'Error clearing cache: {str(e)}'
        }), 500

@auditor_bp.route('/export-auditor-data', methods=['POST'])
def export_auditor_data():
    """Export processed auditor data from session cache"""
    logger.info("Starting export-auditor-data endpoint")
    
    try:
        data = request.json
        if not data:
            return jsonify({'error': 'No JSON data received'}), 400
            
        filepath = data.get('filepath')
        sheet_name = data.get('sheet_name')
        table_choice = data.get('table_choice', 'Table 2: SALES in Value')
        export_format = data.get('format', 'csv')
        
        logger.info(f"Export parameters: filepath={filepath}, sheet_name={sheet_name}, table_choice={table_choice}, format={export_format}")
        
        if not filepath or not sheet_name:
            return jsonify({'error': 'Missing filepath or sheet_name'}), 400
        
        # Try to get data from session cache first
        if table_choice == "Table 1: SALES in MT/Tonage":
            table_key = get_session_key(filepath, sheet_name, 'Table1')
        else:
            table_key = get_session_key(filepath, sheet_name, 'Table2')
        
        df = None
        
        if table_key in session:
            # Get data from session
            logger.info("Getting data from session cache")
            cached_data = session[table_key]
            df = pd.DataFrame(cached_data['data'])
        else:
            # If not in cache, process fresh (fallback)
            logger.info("Data not in cache, processing fresh")
            try:
                # Call the main processing function
                from flask import current_app
                with current_app.test_request_context(
                    '/process-auditor-auto',
                    method='POST',
                    json={
                        'filepath': filepath,
                        'sheet_name': sheet_name,
                        'table_choice': table_choice
                    }
                ):
                    process_response = process_auditor_auto()
                    
                    if isinstance(process_response, tuple):
                        process_data = process_response[0].get_json()
                        status_code = process_response[1]
                    else:
                        process_data = process_response.get_json()
                        status_code = 200
                    
                    if status_code != 200 or not process_data.get('success'):
                        logger.error(f"Failed to process data: {process_data}")
                        return jsonify(process_data), status_code if status_code != 200 else 400
                    
                    df = pd.DataFrame(process_data['data'])
                    
            except Exception as e:
                logger.error(f"Error in fallback processing: {str(e)}")
                return jsonify({
                    'error': f'Data not in cache and fallback processing failed: {str(e)}',
                    'traceback': traceback.format_exc()
                }), 500
        
        if df is None or df.empty:
            return jsonify({'error': 'No data available for export'}), 400
        
        # Generate filename
        clean_table_name = table_choice.replace(':', '').replace(' ', '_').lower()
        clean_sheet_name = re.sub(r'[^\w\-_]', '_', sheet_name)
        filename = f"auditor_{clean_table_name}_{clean_sheet_name}"
        
        if export_format == 'csv':
            try:
                from io import StringIO
                output = StringIO()
                df.to_csv(output, index=False)
                csv_data = output.getvalue()
                
                logger.info(f"CSV export successful. Data size: {len(csv_data)} characters")
                
                return jsonify({
                    'success': True,
                    'data': csv_data,
                    'filename': f"{filename}.csv",
                    'format': 'csv',
                    'rows_exported': len(df),
                    'columns_exported': len(df.columns)
                })
            except Exception as e:
                logger.error(f"Error creating CSV: {str(e)}")
                return jsonify({
                    'error': f'Error creating CSV: {str(e)}',
                    'traceback': traceback.format_exc()
                }), 500
        else:
            return jsonify({'error': 'Only CSV export is currently supported'}), 400
                
    except Exception as e:
        logger.error(f"Unexpected error in export-auditor-data: {str(e)}")
        logger.error(f"Full traceback: {traceback.format_exc()}")
        return jsonify({
            'error': f'Error exporting auditor data: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

# Additional utility endpoints for debugging

@auditor_bp.route('/debug-session', methods=['GET'])
def debug_session():
    """Debug endpoint to check session contents"""
    logger.info("Debug session endpoint called")
    
    try:
        auditor_keys = [key for key in session.keys() if key.startswith('auditor_')]
        session_info = {}
        
        for key in auditor_keys:
            try:
                data = session[key]
                session_info[key] = {
                    'table_name': data.get('table_name', 'Unknown'),
                    'shape': data.get('shape', [0, 0]),
                    'columns_count': len(data.get('columns', [])),
                    'data_rows': len(data.get('data', [])),
                    'success': data.get('success', False)
                }
            except Exception as e:
                session_info[key] = {'error': str(e)}
        
        return jsonify({
            'success': True,
            'total_auditor_keys': len(auditor_keys),
            'session_info': session_info,
            'all_session_keys': list(session.keys())
        })
        
    except Exception as e:
        logger.error(f"Error in debug-session: {str(e)}")
        return jsonify({
            'error': f'Error debugging session: {str(e)}',
            'traceback': traceback.format_exc()
        }), 500

@auditor_bp.route('/health-check', methods=['GET'])
def health_check():
    """Health check endpoint"""
    try:
        # Test DataProcessor initialization
        processor = DataProcessor()
        
        return jsonify({
            'success': True,
            'message': 'Auditor blueprint is healthy',
            'data_processor_available': True,
            'session_keys_count': len(session.keys()),
            'auditor_session_keys': len([k for k in session.keys() if k.startswith('auditor_')])
        })
    except Exception as e:
        logger.error(f"Health check failed: {str(e)}")
        return jsonify({
            'success': False,
            'error': f'Health check failed: {str(e)}',
            'data_processor_available': False
        }), 500