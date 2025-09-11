from flask import Blueprint, request, jsonify, send_file
import pandas as pd
from io import BytesIO
import base64
import re
from services.mapping_service import process_budget_file

budget_bp = Blueprint("budget", __name__)

def ensure_numeric_executive_code(df, exec_code_col):
    """Ensure executive code column is properly numeric"""
    if exec_code_col not in df.columns:
        return df
    
    print(f"Original {exec_code_col} data type: {df[exec_code_col].dtype}")
    print(f"Sample original values: {df[exec_code_col].head().tolist()}")
    
    # Create a copy to avoid modifying original
    df_copy = df.copy()
    
    # Convert to string first to handle mixed data types
    df_copy[exec_code_col] = df_copy[exec_code_col].astype(str)
    
    # Handle NaN/None/empty values
    df_copy[exec_code_col] = df_copy[exec_code_col].replace(['nan', 'None', 'NaN', ''], '0')
    
    # Extract only digits from each value
    df_copy[exec_code_col] = df_copy[exec_code_col].apply(
        lambda x: re.sub(r'\D', '', str(x)) if pd.notna(x) else '0'
    )
    
    # Handle empty strings after digit extraction
    df_copy[exec_code_col] = df_copy[exec_code_col].apply(lambda x: '0' if x == '' else x)
    
    # Convert to numeric, coerce errors to NaN, then fill NaN with 0
    df_copy[exec_code_col] = pd.to_numeric(df_copy[exec_code_col], errors='coerce').fillna(0)
    
    # Ensure integer type
    df_copy[exec_code_col] = df_copy[exec_code_col].astype('int64')
    
    print(f"Cleaned {exec_code_col} data type: {df_copy[exec_code_col].dtype}")
    print(f"Sample cleaned values: {df_copy[exec_code_col].head().tolist()}")
    
    return df_copy

@budget_bp.route("/upload-budget-file", methods=["POST"])
def upload_budget_file():
    try:
        file = request.files["file"]
        sheet_name = request.form["sheet_name"]
        header_row = int(request.form["header_row"])

        # Read Excel file
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)

        # Get mapped columns from request
        customer_col = request.form["customer_col"]
        exec_code_col = request.form["exec_code_col"]
        exec_name_col = request.form["exec_name_col"]
        branch_col = request.form["branch_col"]
        region_col = request.form["region_col"]
        cust_name_col = request.form.get("cust_name_col") or None

        # --- Enhanced Executive Code column cleaning ---
        df = ensure_numeric_executive_code(df, exec_code_col)

        # Process file
        processed = process_budget_file(
            df, customer_col, exec_code_col, exec_name_col, branch_col, region_col, cust_name_col
        )

        # Ensure the processed dataframe also has numeric executive codes
        processed = ensure_numeric_executive_code(processed, exec_code_col)

        # --- Enhanced Export with better number formatting ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Ensure numeric columns are properly typed before writing
            if exec_code_col in processed.columns:
                processed[exec_code_col] = pd.to_numeric(
                    processed[exec_code_col], errors='coerce'
                ).fillna(0).astype('int64')
            
            processed.to_excel(writer, index=False, sheet_name="Sheet1")
            workbook = writer.book
            worksheet = writer.sheets["Sheet1"]

            # Apply numeric format to Executive Code column
            if exec_code_col in processed.columns:
                col_idx = processed.columns.get_loc(exec_code_col)
                col_letter = chr(ord("A") + col_idx)
                
                # Create number format (no decimals, right aligned)
                num_format = workbook.add_format({
                    'num_format': '0',      # Integer format
                    'align': 'right',       # Right align numbers
                    'valign': 'vcenter'     # Vertical center alignment
                })
                
                # Apply format to the entire column (starting from row 2, after header)
                worksheet.set_column(f"{col_letter}:{col_letter}", 15, num_format)

        output.seek(0)
        encoded_excel = base64.b64encode(output.read()).decode("utf-8")

        return {
            "preview": processed.head(10).fillna("").to_dict(orient="records"),
            "counts": {
                "total": int(len(processed)),
                "execs": int(processed["Executive Name"].astype(bool).sum()) if "Executive Name" in processed.columns else 0,
                "branches": int((processed["Branch"] != "").sum()) if "Branch" in processed.columns else 0,
                "regions": int((processed["Region"] != "").sum()) if "Region" in processed.columns else 0,
            },
            "columns": list(processed.columns),
            "file_data": encoded_excel,
        }
        
    except Exception as e:
        print(f"Error in upload_budget_file: {str(e)}")
        return jsonify({"error": f"Failed to process file: {str(e)}"}), 500


@budget_bp.route("/save-budget-file", methods=["POST"])
def save_budget_file():
    try:
        data = request.get_json()
        base64_excel = data.get("file_data")
        filename = data.get("filename", "Processed_Budget.xlsx")

        if not filename.lower().endswith(".xlsx"):
            filename += ".xlsx"

        if not base64_excel:
            return jsonify({"error": "No file data provided"}), 400

        from base64 import b64decode
        from models.schema import BudgetFile
        from extensions import db

        binary_data = b64decode(base64_excel)
        budget_file = BudgetFile(filename=filename, file_data=binary_data)
        db.session.add(budget_file)
        db.session.commit()

        return jsonify({"message": "File saved successfully", "id": budget_file.id})
        
    except Exception as e:
        print(f"Error in save_budget_file: {str(e)}")
        return jsonify({"error": f"Failed to save file: {str(e)}"}), 500


@budget_bp.route("/budget-files", methods=["GET"])
def list_budget_files():
    try:
        from models.schema import BudgetFile
        files = BudgetFile.query.order_by(BudgetFile.uploaded_at.desc()).all()
        return jsonify([
            {
                "id": f.id, 
                "filename": f.filename, 
                "uploaded_at": f.uploaded_at.isoformat()
            }
            for f in files
        ])
    except Exception as e:
        print(f"Error in list_budget_files: {str(e)}")
        return jsonify({"error": f"Failed to list files: {str(e)}"}), 500


@budget_bp.route("/budget-files/<int:file_id>/download", methods=["GET"])
def download_budget_file(file_id):
    try:
        from models.schema import BudgetFile
        import io
        
        file = BudgetFile.query.get_or_404(file_id)

        return send_file(
            io.BytesIO(file.file_data),
            download_name=file.filename,
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        print(f"Error in download_budget_file: {str(e)}")
        return jsonify({"error": f"Failed to download file: {str(e)}"}), 500


@budget_bp.route("/budget-files/<int:file_id>", methods=["DELETE"])
def delete_budget_file(file_id):
    try:
        from models.schema import BudgetFile
        from extensions import db
        
        budget_file = BudgetFile.query.get(file_id)
        if not budget_file:
            return jsonify({"error": "File not found"}), 404

        db.session.delete(budget_file)
        db.session.commit()

        return jsonify({"message": "File deleted successfully"})
        
    except Exception as e:
        print(f"Error in delete_budget_file: {str(e)}")
        return jsonify({"error": f"Failed to delete file: {str(e)}"}), 500


@budget_bp.route("/budget-files/<int:file_id>/data", methods=["GET"])
def get_budget_file_data(file_id):
    try:
        from models.schema import BudgetFile
        from utils.file_utils import read_excel_from_binary  # Ensure this exists
        
        file = BudgetFile.query.get(file_id)
        if not file:
            return jsonify({"error": "Budget file not found"}), 404

        df = read_excel_from_binary(file.file_data)
        
        # Convert data to records, handling any numeric formatting issues
        records = df.head(100).fillna("").to_dict(orient="records")
        
        return jsonify(records)
        
    except Exception as e:
        print(f"Error in get_budget_file_data: {str(e)}")
        return jsonify({"error": f"Failed to read Excel: {str(e)}"}), 500


# Additional utility function for data validation
@budget_bp.route("/validate-budget-columns", methods=["POST"])
def validate_budget_columns():
    """Validate that the required columns exist and have appropriate data types"""
    try:
        file = request.files["file"]
        sheet_name = request.form["sheet_name"]
        header_row = int(request.form["header_row"])
        
        # Read Excel file
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
        
        # Get column information
        column_info = {}
        for col in df.columns:
            column_info[col] = {
                "dtype": str(df[col].dtype),
                "non_null_count": int(df[col].notna().sum()),
                "total_count": len(df),
                "sample_values": df[col].dropna().head(3).tolist()
            }
        
        return jsonify({
            "columns": list(df.columns),
            "column_info": column_info,
            "total_rows": len(df)
        })
        
    except Exception as e:
        print(f"Error in validate_budget_columns: {str(e)}")
        return jsonify({"error": f"Failed to validate columns: {str(e)}"}), 500
