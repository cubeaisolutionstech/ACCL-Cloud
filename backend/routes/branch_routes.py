from flask import Blueprint, request, jsonify, send_file
import os
import pandas as pd
import datetime
import logging
import json 
from werkzeug.utils import secure_filename
from utils.budget_vs_billed import calculate_budget_vs_billed, auto_map_budget_columns
from utils.ppt_generator import create_budget_ppt, create_od_ppt_updated, create_product_growth_ppt, create_nbc_individual_ppt, create_od_individual_ppt, create_consolidated_ppt, format_budget_dataframes, format_od_target_vs_collection_dataframes, format_product_growth_dataframes
from utils.od_target import auto_map_od_columns,calculate_od_values_updated,create_region_branch_mapping,create_dynamic_regional_summary,get_cumulative_branches,get_cumulative_regions
from utils.product_growth import calculate_product_growth,auto_map_product_growth_columns,standardize_name,format_product_growth_dataframes_backend
from utils.nbc_od_utils import auto_map_nbc_columns,auto_map_od_target_columns,create_customer_table,filter_os_qty,nbc_branch_mapping
from utils.branch_proof import (
    create_proof_of_calculation_excel,
    get_required_columns
)

branch_bp = Blueprint('branch', __name__)
logger = logging.getLogger(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs("static", exist_ok=True)


@branch_bp.route('/upload', methods=['POST'])
def upload_file():
    file = request.files.get('file')
    if not file:
        return jsonify({'error': 'No file uploaded'}), 400
    filename = secure_filename(file.filename)
    path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(path)
    return jsonify({'message': 'File uploaded successfully', 'filename': filename})

@branch_bp.route('/sheets', methods=['POST'])
def get_sheet_names():
    file = request.json.get('filename')
    path = os.path.join(UPLOAD_FOLDER, file)
    try:
        xl = pd.ExcelFile(path)
        return jsonify({'sheets': xl.sheet_names})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@branch_bp.route('/get_columns', methods=['POST'])
def get_columns_and_auto_map():
    data = request.json
    path = os.path.join('uploads', data['filename'])
    header_row = int(data['header']) - 1  # user sends 1-based
    sheet_name = data['sheet_name']

    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
        columns = df.columns.tolist()
        return jsonify({'columns': columns})
    except Exception as e:
        import traceback
        logger.error(f"❌ Error in get_columns: {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500
 
@branch_bp.route('/auto_map_columns', methods=['POST'])
def auto_map_columns():
    try:
        data = request.json
        sales_cols = data.get('sales_columns', [])
        budget_cols = data.get('budget_columns', [])
        sales_mapping, budget_mapping = auto_map_budget_columns(sales_cols, budget_cols)
        return jsonify({
            'sales_mapping': sales_mapping,
            'budget_mapping': budget_mapping
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@branch_bp.route('/get_exec_branch_options', methods=['POST'])
def get_exec_branch_options():
    try:
        data = request.json
        sales_path = os.path.join('uploads', data['sales_filename'])
        budget_path = os.path.join('uploads', data['budget_filename'])

        sales_df = pd.read_excel(sales_path, sheet_name=data['sales_sheet'], header=data['sales_header'] - 1)
        budget_df = pd.read_excel(budget_path, sheet_name=data['budget_sheet'], header=data['budget_header'] - 1)

        exec_sales_col = data['sales_exec_col']
        exec_budget_col = data['budget_exec_col']
        area_sales_col = data['sales_area_col']
        area_budget_col = data['budget_area_col']

        from utils.budget_vs_billed import map_branch

        sales_execs = sorted(sales_df[exec_sales_col].dropna().unique().tolist())
        budget_execs = sorted(budget_df[exec_budget_col].dropna().unique().tolist())

        combined = pd.concat([sales_df[area_sales_col], budget_df[area_budget_col]], ignore_index=True).dropna()
        branches = sorted(set(map(map_branch, combined)))

        return jsonify({
            'sales_executives': sales_execs,
            'budget_executives': budget_execs,
            'branches': branches
        })

    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@branch_bp.route('/extract_months', methods=['POST'])
def extract_months():
    try:
        data = request.json
        sales_path = os.path.join('uploads', data['sales_filename'])
        df = pd.read_excel(sales_path, sheet_name=data['sales_sheet'], header=data['sales_header'] - 1)
        date_col = data['sales_date_col']

        months = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce').dt.strftime('%b %y').dropna().unique().tolist()
        return jsonify({'months': months})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@branch_bp.route('/calculate_budget_vs_billed', methods=['POST'])
def budget_vs_billed():
    data = request.json
    try:
        results = calculate_budget_vs_billed(data)
        return jsonify(results)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

# NEW ROUTES FOR PROOF OF CALCULATION

@branch_bp.route('/generate_proof_of_calculation', methods=['POST'])
def generate_proof_of_calculation():
    """Generate proof of calculation Excel file"""
    try:
        data = request.json
        
        # Load the DataFrames
        budget_path = os.path.join('uploads', data['budget_filename'])
        sales_path = os.path.join('uploads', data['sales_filename'])
        
        budget_df = pd.read_excel(budget_path, sheet_name=data['budget_sheet'], header=data['budget_header'] - 1)
        sales_df = pd.read_excel(sales_path, sheet_name=data['sales_sheet'], header=data['sales_header'] - 1)
        
        # Get parameters
        selected_month = data['selected_month']
        selected_executives = data.get('selected_executives', [])
        selected_branches = data.get('selected_branches', [])
        
        # Generate the Excel file
        excel_buffer = create_proof_of_calculation_excel(
            budget_df=budget_df,
            sales_df=sales_df,
            selected_month=selected_month,
            budget_exec_col=data['budget_exec_col'],
            budget_area_col=data['budget_area_col'],
            budget_sl_code_col=data['budget_sl_code_col'],
            budget_product_group_col=data['budget_product_group_col'],
            budget_qty_col=data['budget_qty_col'],
            budget_value_col=data['budget_value_col'],
            sales_exec_col=data['sales_exec_col'],
            sales_date_col=data['sales_date_col'],
            sales_area_col=data['sales_area_col'],
            sales_sl_code_col=data['sales_sl_code_col'],
            sales_product_group_col=data['sales_product_group_col'],
            sales_qty_col=data['sales_qty_col'],
            sales_value_col=data['sales_value_col'],
            selected_executives=selected_executives,
            selected_branches=selected_branches
        )
        
        if excel_buffer is None:
            return jsonify({'error': 'Failed to generate proof of calculation'}), 500
        
        filename = f"Proof_of_Calculation_{selected_month.replace(' ', '_')}.xlsx"
        
        return send_file(
            excel_buffer,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        import traceback
        logger.error(f"❌ Error in generate_proof_of_calculation: {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500

@branch_bp.route('/get_proof_summary', methods=['POST'])
def get_proof_summary():
    """Get summary statistics for proof of calculation"""
    try:
        data = request.json
        
        # Load the DataFrames
        budget_path = os.path.join('uploads', data['budget_filename'])
        sales_path = os.path.join('uploads', data['sales_filename'])
        
        budget_df = pd.read_excel(budget_path, sheet_name=data['budget_sheet'], header=data['budget_header'] - 1)
        sales_df = pd.read_excel(sales_path, sheet_name=data['sales_sheet'], header=data['sales_header'] - 1)
        
        # Get parameters
        selected_month = data['selected_month']
        selected_executives = data.get('selected_executives', [])
        selected_branches = data.get('selected_branches', [])
        
        # Create proof DataFrame without generating Excel (for summary only)
        from utils.branch_proof import BRANCH_MAPPING
        import copy
        
        # Create copies to avoid modifying original DataFrames
        sales_df_copy = sales_df.copy()
        budget_df_copy = budget_df.copy()
        
        # Apply the same filtering logic as in the main function
        sales_df_copy[data['sales_date_col']] = pd.to_datetime(sales_df_copy[data['sales_date_col']], dayfirst=True, errors='coerce')
        filtered_sales_df = sales_df_copy[sales_df_copy[data['sales_date_col']].dt.strftime('%b %y') == selected_month].copy()
        
        if filtered_sales_df.empty:
            return jsonify({'error': f'No sales data found for {selected_month}'}), 400
        
        # Apply executive filtering
        if selected_executives:
            filtered_sales_df = filtered_sales_df[filtered_sales_df[data['sales_exec_col']].isin(selected_executives)]
            budget_df_copy = budget_df_copy[budget_df_copy[data['budget_exec_col']].isin(selected_executives)]
        
        # Apply branch filtering and mapping (simplified version)
        if selected_branches:
            # Apply branch standardization
            budget_df_copy[data['budget_area_col']] = budget_df_copy[data['budget_area_col']].astype(str).str.strip()
            budget_df_copy[data['budget_area_col']] = budget_df_copy[data['budget_area_col']].str.split(' - ').str[-1].str.upper()
            budget_df_copy[data['budget_area_col']] = budget_df_copy[data['budget_area_col']].replace(BRANCH_MAPPING)
            
            filtered_sales_df[data['sales_area_col']] = filtered_sales_df[data['sales_area_col']].astype(str).str.upper().replace(BRANCH_MAPPING)
            
            filtered_sales_df = filtered_sales_df[filtered_sales_df[data['sales_area_col']].isin(selected_branches)]
            budget_df_copy = budget_df_copy[budget_df_copy[data['budget_area_col']].isin(selected_branches)]
        
        # Generate summary statistics
        budget_records = len(budget_df_copy)
        sales_records = len(filtered_sales_df)
        
        # Count valid budget records
        budget_df_copy[data['budget_qty_col']] = pd.to_numeric(budget_df_copy[data['budget_qty_col']], errors='coerce').fillna(0)
        budget_df_copy[data['budget_value_col']] = pd.to_numeric(budget_df_copy[data['budget_value_col']], errors='coerce').fillna(0)
        valid_budget_records = len(budget_df_copy[
            (budget_df_copy[data['budget_qty_col']] > 0) & 
            (budget_df_copy[data['budget_value_col']] > 0)
        ])
        
        summary = {
            'total_budget_records': budget_records,
            'total_sales_records': sales_records,
            'valid_budget_records': valid_budget_records,
            'selected_month': selected_month,
            'selected_executives_count': len(selected_executives) if selected_executives else 0,
            'selected_branches_count': len(selected_branches) if selected_branches else 0
        }
        
        return jsonify(summary)
        
    except Exception as e:
        logger.error(f"❌ Error in get_proof_summary: {e}")
        return jsonify({'error': str(e)}), 500

@branch_bp.route('/get_required_proof_columns', methods=['GET'])
def get_required_proof_columns():
    """Get the standard list of required columns for proof of calculation"""
    try:
        budget_columns, sales_columns = get_required_columns()
        return jsonify({
            'budget_columns': budget_columns,
            'sales_columns': sales_columns
        })
    except Exception as e:
        logger.error(f"❌ Error in get_required_proof_columns: {e}")
        return jsonify({'error': str(e)}), 500

# EXISTING ROUTES CONTINUE...

@branch_bp.route('/generate_consolidated_branch_ppt', methods=['POST'])
def generate_consolidated_branch_ppt():
    try:
        data = request.get_json()
        report_title = data.get('reportTitle', 'ACCLLP Consolidated Report')
        all_dfs_info = data.get('allDfsInfo', [])
        
        parsed_sections = []
        for item in all_dfs_info:
            df_data = item.get("df")
            title = item.get("title", "Untitled Section")
            percent_cols = item.get("percent_cols", [])
            columns = item.get("columns", [])  # Get column order
            
            if df_data:
                df = pd.DataFrame(df_data)
                
                # IMPORTANT: Apply column ordering if provided
                if columns:
                    # Reorder columns according to frontend specification
                    available_cols = [col for col in columns if col in df.columns]
                    if available_cols:
                        df = df[available_cols]
                        logger.info(f"✅ Applied column order for {title}: {available_cols}")
                
                parsed_sections.append({
                    "title": title,
                    "df": df,
                    "percent_cols": percent_cols
                })

        # Use the branch consolidated PPT function
        ppt_stream = create_consolidated_ppt(
            all_dfs_with_titles=parsed_sections,  # Use the correct parameter name
            title=report_title,
            logo_file=None
        )

        return send_file(
            ppt_stream,
            as_attachment=True,
            download_name=f"{report_title.replace(' ', '_')}.pptx",
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        logger.error(f"❌ Error in generate_consolidated_branch_ppt: {e}")
        return jsonify({"error": str(e)}), 500
    
@branch_bp.route('/download_ppt', methods=['POST'])
def download_ppt():
    try:
        payload = request.get_json()
        print("📦 PPT Payload Keys:", list(payload.keys()))
        
        # Extract the data from payload and ensure proper formatting
        budget_vs_billed_qty_data = payload.get('budget_vs_billed_qty', [])
        budget_vs_billed_value_data = payload.get('budget_vs_billed_value', [])
        overall_sales_qty_data = payload.get('overall_sales_qty', [])
        overall_sales_value_data = payload.get('overall_sales_value', [])
        
        # Create DataFrames
        budget_vs_billed_qty_df = pd.DataFrame(budget_vs_billed_qty_data)
        budget_vs_billed_value_df = pd.DataFrame(budget_vs_billed_value_data)
        overall_sales_qty_df = pd.DataFrame(overall_sales_qty_data)
        overall_sales_value_df = pd.DataFrame(overall_sales_value_data)
        
        # Apply proper column formatting and ordering
        budget_vs_billed_qty_df, budget_vs_billed_value_df, overall_sales_qty_df, overall_sales_value_df = format_budget_dataframes(
            budget_vs_billed_qty_df, budget_vs_billed_value_df, overall_sales_qty_df, overall_sales_value_df
        )
        
        month_title = payload.get('month_title', 'Monthly Report')
        
        ppt_buffer = create_budget_ppt(
            budget_vs_billed_qty_df=budget_vs_billed_qty_df,
            budget_vs_billed_value_df=budget_vs_billed_value_df,
            overall_sales_qty_df=overall_sales_qty_df,
            overall_sales_value_df=overall_sales_value_df,
            month_title=month_title,
            logo_file=payload.get('logo_file')
        )
        
        if ppt_buffer:
            return send_file(
                ppt_buffer,
                as_attachment=True,
                download_name=f"Budget_Report_{month_title.replace(' ', '_')}.pptx",
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
        else:
            return jsonify({'error': 'PPT generation failed'}), 500
            
    except Exception as e:
        print("❌ PPT generation failed:", str(e))
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
    
@branch_bp.route('/get_od_columns', methods=['POST'])
def get_od_columns():
    data = request.get_json()
    os_prev_cols = data.get('os_prev_columns', [])
    os_curr_cols = data.get('os_curr_columns', [])
    sales_cols = data.get('sales_columns', [])

    try:
        os_jan_map, os_feb_map, sales_map = auto_map_od_columns(os_prev_cols, os_curr_cols, sales_cols)
        return jsonify({
            'os_jan_mapping': os_jan_map,
            'os_feb_mapping': os_feb_map,
            'sales_mapping': sales_map
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@branch_bp.route('/get_od_filter_options', methods=['POST'])
def get_od_filter_options():
    data = request.get_json()
    try:
        os_prev = pd.read_excel(f"uploads/{data['os_prev_filename']}", sheet_name=data['os_prev_sheet'], header=data['os_prev_header'] - 1)
        os_curr = pd.read_excel(f"uploads/{data['os_curr_filename']}", sheet_name=data['os_curr_sheet'], header=data['os_curr_header'] - 1)
        sales = pd.read_excel(f"uploads/{data['sales_filename']}", sheet_name=data['sales_sheet'], header=data['sales_header'] - 1)

        # Pull mapped column names
        os_prev_exec = data['os_prev_mapping'].get('executive')
        os_curr_exec = data['os_curr_mapping'].get('executive')
        sales_exec = data['sales_mapping'].get('executive')

        os_prev_branch = data['os_prev_mapping'].get('branch')
        os_curr_branch = data['os_curr_mapping'].get('branch')
        sales_branch = data['sales_mapping'].get('branch')

        os_prev_region = data['os_prev_mapping'].get('region')
        os_curr_region = data['os_curr_mapping'].get('region')
        sales_region = data['sales_mapping'].get('region')

        execs = set()
        branches = set()
        regions = set()

        if os_prev_exec in os_prev.columns:
            execs.update(os_prev[os_prev_exec].dropna().unique())
        if os_curr_exec in os_curr.columns:
            execs.update(os_curr[os_curr_exec].dropna().unique())
        if sales_exec in sales.columns:
            execs.update(sales[sales_exec].dropna().unique())

        if os_prev_branch in os_prev.columns:
            branches.update(os_prev[os_prev_branch].dropna().unique())
        if os_curr_branch in os_curr.columns:
            branches.update(os_curr[os_curr_branch].dropna().unique())
        if sales_branch in sales.columns:
            branches.update(sales[sales_branch].dropna().unique())

        if os_prev_region and os_prev_region in os_prev.columns:
            regions.update(os_prev[os_prev_region].dropna().unique())
        if os_curr_region and os_curr_region in os_curr.columns:
            regions.update(os_curr[os_curr_region].dropna().unique())
        if sales_region and sales_region in sales.columns:
            regions.update(sales[sales_region].dropna().unique())

        return jsonify({
            "executives": sorted(list(execs)),
            "branches": sorted(list(branches)),
            "regions": sorted(list(regions))
        })

    except Exception as e:
        return jsonify({'error': f'Failed to fetch filters: {str(e)}'}), 500

@branch_bp.route('/calculate_od_target', methods=['POST'])
def calculate_od_target():
    try:
        data = request.get_json()

        print("🔧 OD Calculation Payload Keys:", list(data.keys()))
        print("🧩 Mappings Preview (os_prev):", data.get("os_prev_mapping", {}))
        print("🧩 Mappings Preview (sales):", data.get("sales_mapping", {}))
        print("📄 Filenames:", data['os_prev_filename'], data['sales_filename'])

        # Load the 3 Excel files
        os_prev = pd.read_excel(f"uploads/{data['os_prev_filename']}", sheet_name=data['os_prev_sheet'], header=data['os_prev_header'] - 1)
        os_curr = pd.read_excel(f"uploads/{data['os_curr_filename']}", sheet_name=data['os_curr_sheet'], header=data['os_curr_header'] - 1)
        sales = pd.read_excel(f"uploads/{data['sales_filename']}", sheet_name=data['sales_sheet'], header=data['sales_header'] - 1)

        # === Apply filters and calculate final output ===
        final, regional, region_map = calculate_od_values_updated(
            os_prev, os_curr, sales,
            data['selected_month'],
            data['os_prev_mapping']['due_date'],
            data['os_prev_mapping']['ref_date'],
            data['os_prev_mapping']['branch'],
            data['os_prev_mapping']['net_value'],
            data['os_prev_mapping']['executive'],
            data['os_prev_mapping'].get('region'),
            data['os_curr_mapping']['due_date'],
            data['os_curr_mapping']['ref_date'],
            data['os_curr_mapping']['branch'],
            data['os_curr_mapping']['net_value'],
            data['os_curr_mapping']['executive'],
            data['os_curr_mapping'].get('region'),
            data['sales_mapping']['bill_date'],
            data['sales_mapping']['due_date'],
            data['sales_mapping']['branch'],
            data['sales_mapping']['value'],
            data['sales_mapping']['executive'],
            data['sales_mapping'].get('region'),
            data.get('selected_executives', []),
            data.get('selected_branches', []),
            data.get('selected_regions', [])
        )

        return jsonify({
            "branch_summary": final.to_dict(orient='records') if final is not None else [],
            "regional_summary": regional.to_dict(orient='records') if regional is not None else [],
            "region_mapping": region_map
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'OD Target calculation failed: {str(e)}'}), 500

@branch_bp.route('/download_od_ppt', methods=['POST'])
def download_od_ppt():
    try:
        data = request.get_json()

        branch_data = data.get('branch_summary', {}).get('data', [])
        branch_columns = data.get('branch_summary', {}).get('columns', [])

        regional_data = data.get('regional_summary', {}).get('data', [])
        regional_columns = data.get('regional_summary', {}).get('columns', [])

        title = data.get('title', 'OD Target vs Collection')

        def to_ordered_df(data, columns):
            df = pd.DataFrame(data)
            if columns and not df.empty:
                # Ensure we only use columns that exist in the DataFrame
                available_columns = [col for col in columns if col in df.columns]
                if available_columns:
                    df = df[available_columns]
            
            # CRITICAL FIX: Ensure percentage columns are properly formatted
            for col in df.columns:
                if 'achieved' in col.lower() and '%' in col.lower():
                    # Ensure percentage values are numeric
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).round(2)
                elif pd.api.types.is_numeric_dtype(df[col]):
                    df[col] = df[col].round(2)
            
            return df
        
        branch_df = to_ordered_df(branch_data, branch_columns)
        regional_df = to_ordered_df(regional_data, regional_columns)

        # CRITICAL FIX: Debug regional DataFrame before formatting
        logger.info(f"Regional DF columns before formatting: {regional_df.columns.tolist()}")
        if not regional_df.empty:
            logger.info(f"Sample regional row: {regional_df.iloc[0].to_dict()}")

        # Apply formatting before creating PPT
        branch_df, regional_df = format_od_target_vs_collection_dataframes(branch_df, regional_df)
        
        # CRITICAL FIX: Debug after formatting
        logger.info(f"Regional DF columns after formatting: {regional_df.columns.tolist()}")

        ppt_buffer = create_od_ppt_updated(branch_df, regional_df, title)

        if ppt_buffer:
            return send_file(
                ppt_buffer,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                as_attachment=True,
                download_name=f"OD_Target_vs_Collection_{title.replace(' ', '_')}.pptx"
            )
        else:
            return jsonify({'error': 'PPT generation failed'}), 500
    except Exception as e:
        logger.error(f"Error in download_od_ppt: {e}")
        return jsonify({'error': str(e)}), 500

# Add these routes to your branch_routes.py if they're missing or replace existing ones

@branch_bp.route('/auto_map_product_growth', methods=['POST'])
def auto_map_pg_columns():
    data = request.get_json()
    ly_mapping, cy_mapping, budget_mapping = auto_map_product_growth_columns(
        data['ly_columns'], data['cy_columns'], data['budget_columns']
    )
    return jsonify({
        'ly_mapping': ly_mapping,
        'cy_mapping': cy_mapping,
        'budget_mapping': budget_mapping
    })

@branch_bp.route("/get_product_growth_filters", methods=["POST"])
def get_product_growth_filters():
    try:
        data = request.get_json()

        ly_filename = data["ly_filename"]
        cy_filename = data["cy_filename"]
        budget_filename = data["budget_filename"]

        ly_sheet = data["ly_sheet"]
        cy_sheet = data["cy_sheet"]
        budget_sheet = data["budget_sheet"]

        ly_header = data["ly_header"] - 1
        cy_header = data["cy_header"] - 1
        budget_header = data["budget_header"] - 1

        # Column names provided from frontend
        ly_date_col = data["ly_date_col"]
        cy_date_col = data["cy_date_col"]

        ly_exec_col = data["ly_exec_col"]
        cy_exec_col = data["cy_exec_col"]
        budget_exec_col = data["budget_exec_col"]

        ly_group_col = data["ly_group_col"]
        cy_group_col = data["cy_group_col"]
        budget_group_col = data["budget_group_col"]

        # Load DataFrames
        ly_df = pd.read_excel(f"uploads/{ly_filename}", sheet_name=ly_sheet, header=ly_header)
        cy_df = pd.read_excel(f"uploads/{cy_filename}", sheet_name=cy_sheet, header=cy_header)
        budget_df = pd.read_excel(f"uploads/{budget_filename}", sheet_name=budget_sheet, header=budget_header)

        # Extract unique months from LY and CY
        ly_months = pd.to_datetime(ly_df[ly_date_col], dayfirst=True, errors="coerce").dt.strftime("%b %y").dropna().unique().tolist()
        cy_months = pd.to_datetime(cy_df[cy_date_col], dayfirst=True, errors="coerce").dt.strftime("%b %y").dropna().unique().tolist()

        # Get unique executives
        executives = pd.concat([
            ly_df[ly_exec_col].dropna(),
            cy_df[cy_exec_col].dropna(),
            budget_df[budget_exec_col].dropna()
        ]).dropna().astype(str).unique().tolist()

        # Get unique company groups
        from utils.product_growth import standardize_name
        all_groups = pd.concat([
            ly_df[ly_group_col].dropna(),
            cy_df[cy_group_col].dropna(),
            budget_df[budget_group_col].dropna()
        ]).dropna().astype(str).map(standardize_name)

        company_groups = sorted(set(all_groups))

        return jsonify({
            "ly_months": sorted(ly_months),
            "cy_months": sorted(cy_months),
            "executives": sorted(executives),
            "company_groups": sorted(company_groups)
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@branch_bp.route('/calculate_product_growth', methods=['POST'])
def calculate_product_growth_api():
    try:
        data = request.get_json()
        print(f"Product Growth API called with payload keys: {list(data.keys())}")

        # Load Excel files directly from 'uploads/'
        ly_df = pd.read_excel(f"uploads/{data['ly_filename']}", sheet_name=data['ly_sheet'], header=data['ly_header'] - 1)
        cy_df = pd.read_excel(f"uploads/{data['cy_filename']}", sheet_name=data['cy_sheet'], header=data['cy_header'] - 1)
        budget_df = pd.read_excel(f"uploads/{data['budget_filename']}", sheet_name=data['budget_sheet'], header=data['budget_header'] - 1)

        print(f"Loaded dataframes - LY: {ly_df.shape}, CY: {cy_df.shape}, Budget: {budget_df.shape}")

        # Import the calculate_product_growth function
        from utils.product_growth import calculate_product_growth

        result = calculate_product_growth(
            ly_df, cy_df, budget_df,
            data['ly_months'], data['cy_months'],
            data['ly_date_col'], data['cy_date_col'],
            data['ly_qty_col'], data['cy_qty_col'],
            data['ly_value_col'], data['cy_value_col'],
            data['budget_qty_col'], data['budget_value_col'],
            data['ly_product_col'], data['cy_product_col'],
            data['ly_company_group_col'], data['cy_company_group_col'],
            data['budget_company_group_col'], data['budget_product_group_col'],
            data['ly_exec_col'], data['cy_exec_col'], data['budget_exec_col'],
            data.get('selected_executives', []),
            data.get('selected_company_groups', [])
        )
        
        if not result:
            return jsonify({'status': 'error', 'error': 'Calculation failed or no data returned'}), 400

        print(f"Raw calculation result keys: {list(result.keys())}")
        for group, group_data in result.items():
            print(f"Group {group} - qty_df shape: {group_data['qty_df'].shape}, value_df shape: {group_data['value_df'].shape}")

        # Apply formatting to ensure consistent column structure
        formatted_results = format_product_growth_dataframes_backend(result)

        # Convert DataFrames to JSON with proper formatting
        response = {}
        for group, df_pair in formatted_results.items():
            qty_records = df_pair['qty_df'].to_dict(orient='records')
            value_records = df_pair['value_df'].to_dict(orient='records')
            
            print(f"Group {group} - qty records: {len(qty_records)}, value records: {len(value_records)}")
            if qty_records:
                print(f"Sample qty record: {qty_records[0]}")
            if value_records:
                print(f"Sample value record: {value_records[0]}")
            
            response[group] = {
                'qty_df': qty_records,
                'value_df': value_records
            }

        print(f"Final response keys: {list(response.keys())}")
        return jsonify({'status': 'success', 'results': response})
        
    except Exception as e:
        import traceback
        print("Error in calculate_product_growth:", e)
        print(traceback.format_exc())
        return jsonify({'status': 'error', 'error': str(e)}), 500

@branch_bp.route('/download_product_growth_ppt', methods=['POST'])
def download_product_growth_ppt():
    try:
        data = request.get_json()
        group_results = data.get("group_results", {})
        month_title = data.get("month_title", "Product Growth")

        parsed_results = {}

        for group, content in group_results.items():
            # If content is a JSON string, parse it
            if isinstance(content, str):
                try:
                    content = json.loads(content)
                except Exception as parse_err:
                    logger.warning(f"Failed to parse JSON for group '{group}': {parse_err}")
                    continue

            # Safely extract DataFrames
            qty_df = pd.DataFrame(content.get("qty_df", []))
            value_df = pd.DataFrame(content.get("value_df", []))

            if qty_df.empty or value_df.empty:
                logger.warning(f"Empty DataFrame for group: {group}")
                continue

            parsed_results[group] = {
                "qty_df": qty_df,
                "value_df": value_df
            }

        if not parsed_results:
            return jsonify({"error": "No data available to generate PPT"}), 400

        # Apply formatting before creating PPT
        formatted_results = format_product_growth_dataframes(parsed_results)

        # Create clean title without LY/CY month details
        clean_title = "Product Growth"
        ppt_buffer = create_product_growth_ppt(formatted_results, clean_title)

        if ppt_buffer:
            return send_file(
                ppt_buffer,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                as_attachment=True,
                download_name=f"Product_Growth_Report.pptx"
            )
        else:
            return jsonify({"error": "PPT generation failed"}), 500

    except Exception as e:
        import traceback
        logger.error("Error in download_product_growth_ppt: %s", traceback.format_exc())
        return jsonify({"error": str(e)}), 500

@branch_bp.route('/get_nbc_columns', methods=['POST'])
def get_nbc_columns():
    try:
        data = request.get_json()
        filename = data['filename']
        sheet_name = data['sheet_name']
        header = data['header'] - 1

        df = pd.read_excel(f"uploads/{filename}", sheet_name=sheet_name, header=header)
        columns = df.columns.tolist()
        mapped = auto_map_nbc_columns(columns)

        return jsonify({"columns": columns, "mapping": mapped})
    except Exception as e:
        import traceback
        print("❌ Error in get_nbc_columns:", traceback.format_exc())
        return jsonify({"error": str(e)}), 500
    
@branch_bp.route("/get_nbc_filters", methods=["POST"])
def get_nbc_filters():
    try:
        data = request.get_json()
        filename = data["filename"]
        sheet_name = data["sheet_name"]
        header_row = data["header"] - 1

        date_col = data["date_col"]
        branch_col = data["branch_col"]
        executive_col = data["executive_col"]

        df = pd.read_excel(f"uploads/{filename}", sheet_name=sheet_name, header=header_row)

        # Clean column names
        df.columns = [str(col).strip() for col in df.columns]

        # Basic validation
        for col in [date_col, branch_col, executive_col]:
            if col not in df.columns:
                return jsonify({"error": f"Column '{col}' not found in data"}), 400

        # Normalize branch names using the updated mapping
        raw_branches = df[branch_col].dropna().astype(str).str.upper().unique().tolist()
        all_branches = sorted(set([
            nbc_branch_mapping.get(branch.split(" - ")[-1], branch.split(" - ")[-1])
            for branch in raw_branches
        ]))

        all_executives = sorted(df[executive_col].dropna().astype(str).unique().tolist())

        return jsonify({
            "branches": all_branches,
            "executives": all_executives
        })

    except Exception as e:
        import traceback
        logger.error(f"❌ Error in get_nbc_filters: {traceback.format_exc()}")
        return jsonify({"error": str(e)}), 500

@branch_bp.route('/calculate_nbc_table', methods=['POST'])
def calculate_nbc_table():
    try:
        data = request.get_json()
        
        filename = data["filename"]
        sheet_name = data["sheet_name"]
        header = data["header"] - 1

        date_col = data["date_col"]
        customer_id_col = data["customer_id_col"]
        branch_col = data["branch_col"]
        executive_col = data["executive_col"]

        selected_branches = data.get("selected_branches", [])
        selected_executives = data.get("selected_executives", [])

        # Load and validate data
        try:
            df = pd.read_excel(f"uploads/{filename}", sheet_name=sheet_name, header=header)
            logger.info(f"Loaded DataFrame with shape: {df.shape}")
            logger.info(f"Columns: {df.columns.tolist()}")
        except Exception as e:
            logger.error(f"Error loading Excel file: {e}")
            return jsonify({"error": f"Failed to load file: {str(e)}"}), 500

        # Validate columns exist
        missing_cols = []
        for col_name, col_value in [("date", date_col), ("branch", branch_col), ("customer_id", customer_id_col), ("executive", executive_col)]:
            if not col_value or col_value not in df.columns:
                missing_cols.append(f"{col_name}: '{col_value}'")
        
        if missing_cols:
            return jsonify({"error": f"Missing columns: {', '.join(missing_cols)}"}), 400

        results = create_customer_table(
            df, date_col, branch_col, customer_id_col, executive_col,
            selected_branches=selected_branches,
            selected_executives=selected_executives
        )

        if not results:
            return jsonify({"error": "No results generated. Check your data and filters."}), 400

        final_output = {}
        for fy, fy_data in results.items():
            result_df = fy_data['dataframe']
            sorted_months = fy_data['months']
            recent_month = fy_data['recent_month']
            
            final_output[fy] = {
                "data": result_df.to_dict(orient="records"),
                "months": sorted_months,
                "recent_month": recent_month
            }

        logger.info(f"Generated results for {len(final_output)} financial years")
        return jsonify({"results": final_output})
        
    except Exception as e:
        import traceback
        logger.error(f"Error in calculate_nbc_table: {traceback.format_exc()}")
        return jsonify({"error": str(e)}), 500

@branch_bp.route('/download_nbc_ppt', methods=['POST'])
def download_nbc_ppt():
    try:
        data = request.get_json()
        df_data = data.get("data", [])
        title = data.get("title", "NBC Report")
        months = data.get("months", [])
        fy = data.get("financial_year", "")
        logo_filename = data.get("logo_file")

        # Validate input data
        if not df_data:
            logger.error("No data provided for PPT generation")
            return jsonify({"error": "No data provided for PPT generation"}), 400

        # Create DataFrame and validate
        df = pd.DataFrame(df_data)
        logger.info(f"PPT DataFrame shape: {df.shape}, columns: {df.columns.tolist()}")

        if df.empty:
            logger.error("DataFrame is empty")
            return jsonify({"error": "DataFrame is empty"}), 400

        # Handle logo path
        logo_path = None
        if logo_filename:
            logo_path = f"uploads/{logo_filename}"
            if not os.path.exists(logo_path):
                logger.warning(f"Logo file not found: {logo_path}")
                logo_path = None

        # Generate PPT with error handling
        try:
            ppt_buffer = create_nbc_individual_ppt(df, title, months, fy, logo_path)
        except Exception as ppt_error:
            logger.error(f"PPT generation failed: {ppt_error}")
            return jsonify({"error": f"PPT generation failed: {str(ppt_error)}"}), 500

        if not ppt_buffer:
            logger.error("PPT generation returned None")
            return jsonify({"error": "Failed to generate PPT - no content returned"}), 500

        # Validate buffer has content
        try:
            buffer_size = len(ppt_buffer.getvalue())
            logger.info(f"PPT buffer size: {buffer_size} bytes")
            if buffer_size == 0:
                return jsonify({"error": "Generated PPT is empty"}), 500
        except Exception as buffer_error:
            logger.error(f"Buffer validation failed: {buffer_error}")
            return jsonify({"error": "Invalid PPT buffer generated"}), 500

        # Create filename
        month_part = title.split(' - ')[-1] if ' - ' in title else fy
        filename = f"NBC_Report_{month_part.replace(' ', '_')}.pptx"
        
        logger.info(f"Sending PPT file: {filename}")
        
        # Reset buffer position before sending
        ppt_buffer.seek(0)
        
        # Create response with explicit headers to avoid CORS conflicts
        response = send_file(
            ppt_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=filename
        )
        
        # Explicitly handle CORS to prevent duplication
        # Remove any existing CORS headers that might cause conflicts
        response.headers.pop('Access-Control-Allow-Origin', None)
        response.headers.pop('Access-Control-Allow-Methods', None)
        response.headers.pop('Access-Control-Allow-Headers', None)
        
        return response
            
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        logger.error(f"Error in download_nbc_ppt: {error_trace}")
        return jsonify({"error": f"Server error: {str(e)}"}), 500

@branch_bp.route("/get_od_target_columns", methods=["POST"])
def get_od_target_columns():
    try:
        data = request.get_json()
        columns = data.get("columns", [])
        mapping = auto_map_od_target_columns(columns)
        return jsonify({"mapping": mapping})
    except Exception as e:
        logger.error(f"❌ Error in get_od_target_columns: {e}")
        return jsonify({"error": str(e)}), 500

@branch_bp.route('/get_column_unique_values', methods=['POST'])
def get_column_unique_values():
    try:
        data = request.get_json()
        filename = data["filename"]
        sheet_name = data["sheet_name"]
        header = data["header"] - 1
        column_names = data.get("column_names", [])

        df = pd.read_excel(f"uploads/{filename}", sheet_name=sheet_name, header=header)

        response = {}

        for col in column_names:
            if col not in df.columns:
                continue

            col_data = df[col].dropna()

            # Detect if this column is due date, area, or executive
            lower_col = col.lower()
            if "date" in lower_col:
                col_data = pd.to_datetime(col_data, errors='coerce')
                valid_years = sorted(col_data.dt.year.dropna().astype(int).unique())
                response[col] = {"years": [str(y) for y in valid_years]}
            elif "area" in lower_col or "branch" in lower_col:
                from utils.nbc_od_utils import extract_area_name
                cleaned = sorted(set(filter(None, col_data.map(extract_area_name))))
                response[col] = {"values": cleaned}
            elif "executive" in lower_col:
                from utils.nbc_od_utils import extract_executive_name
                cleaned = sorted(set(filter(None, col_data.map(extract_executive_name))))
                response[col] = {"values": cleaned}
            else:
                cleaned = sorted(set(col_data.astype(str).str.strip().unique()))
                response[col] = {"values": cleaned}

        return jsonify(response)

    except Exception as e:
        logger.error(f"❌ Error in get_column_unique_values: {e}")
        return jsonify({"error": str(e)}), 500

@branch_bp.route('/calculate_od_target_table', methods=['POST'])
def calculate_od_target_table():
    try:
        data = request.get_json()
        filename = data["filename"]
        sheet_name = data["sheet_name"]
        header = data["header"] - 1

        area_col = data["area_col"]
        due_date_col = data["due_date_col"]
        qty_col = data["qty_col"]
        executive_col = data["executive_col"]

        selected_branches = data.get("selected_branches", [])
        selected_executives = data.get("selected_executives", [])
        selected_years = data.get("selected_years", [])
        till_month = data.get("till_month")

        df = pd.read_excel(f"uploads/{filename}", sheet_name=sheet_name, header=header)

        result_df, start_date, end_date = filter_os_qty(
            df,
            os_area_col=area_col,
            os_qty_col=qty_col,
            os_due_date_col=due_date_col,
            os_exec_col=executive_col,
            selected_branches=selected_branches,
            selected_years=selected_years,
            till_month=till_month,
            selected_executives=selected_executives
        )

        if result_df is None:
            return jsonify({"error": "Failed to generate OD Target table"}), 400

        start_str = start_date.strftime('%b %Y') if start_date else "Start"
        end_str = end_date.strftime('%b %Y') if end_date else "End"

        return jsonify({
            "table": result_df.to_dict(orient="records"),
            "start": start_str,
            "end": end_str
        })

    except Exception as e:
        logger.error(f"❌ Error in calculate_od_target_table: {e}")
        return jsonify({"error": str(e)}), 500

@branch_bp.route('/download_od_target_ppt', methods=['POST'])
def download_od_target_ppt():
    try:
        data = request.get_json()

        df_data = data.get("result", [])
        title = data.get("title", "OD Target")
        logo_filename = data.get("logo_file")

        if not df_data:
            return jsonify({"error": "No data provided"}), 400

        df = pd.DataFrame(df_data)
        
        # FIXED: Clean title formatting - remove any "(Value in Lakhs)" text
        clean_title = title.replace(" (Value in Lakhs)", "").replace("(Value in Lakhs)", "")
        clean_title = clean_title.replace(" - Value in Lakhs", "")
        
        logo_path = f"uploads/{logo_filename}" if logo_filename else None
        ppt_buffer = create_od_individual_ppt(df, clean_title.upper(), logo_path)

        if ppt_buffer:
            return send_file(
                ppt_buffer,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                as_attachment=True,
                download_name=f"{clean_title.replace(' ', '_')}.pptx"
            )
        else:
            return jsonify({"error": "Failed to generate PPT"}), 500
    except Exception as e:
        import traceback
        logger.error("❌ Error in download_od_target_ppt: %s", traceback.format_exc())
        return jsonify({"error": str(e)}), 500
    
