# routes/cumulative_routes.py - Updated to process files fully in-memory

import traceback
import logging
from flask import Blueprint, request, jsonify, send_file
import pandas as pd
from io import BytesIO
from werkzeug.utils import secure_filename

# Create the blueprint - make sure this matches your existing code
api_bp = Blueprint('cumulative_api', __name__)

logger = logging.getLogger(__name__)

def read_excel_file(file_stream, month, filename, skip_first_row=False):
    """Read Excel file directly from memory with improved error handling"""
    try:
        logger.info(f"Reading Excel file: {filename}")

        # Try reading with different engines
        try:
            df = pd.read_excel(file_stream, header=None if skip_first_row else 0, engine="openpyxl")
        except Exception as e1:
            logger.warning(f"openpyxl failed, trying xlrd: {e1}")
            try:
                file_stream.seek(0)
                df = pd.read_excel(file_stream, header=None if skip_first_row else 0, engine="xlrd")
            except Exception as e2:
                logger.warning(f"xlrd failed, trying calamine: {e2}")
                file_stream.seek(0)
                df = pd.read_excel(file_stream, header=None if skip_first_row else 0, engine="calamine")

        if skip_first_row and not df.empty:
            df.columns = df.iloc[0]
            df = df.iloc[1:]

        if not df.empty:
            # Clean column names
            df.columns = [str(col).strip() for col in df.columns]

            # Replace problematic values
            df = df.replace([float('inf'), -float('inf')], None)

            # Add metadata
            if "Month" not in df.columns:
                df["Month"] = month
            df["Source File"] = filename

            logger.info(f"Processed {filename}: {len(df)} rows, {len(df.columns)} columns")
            return df
        else:
            logger.warning(f"Empty dataframe from {filename}")
            return None

    except Exception as e:
        logger.error(f"Error reading {filename}: {str(e)}")
        logger.error(traceback.format_exc())
        return None


@api_bp.route("/process", methods=["POST", "OPTIONS"])
def process_files():
    """Process uploaded Excel files in memory and return preview"""
    logger.info("=== PROCESS ENDPOINT CALLED ===")

    if request.method == "OPTIONS":
        response = jsonify({"status": "ok"})
        response.headers.add('Access-Control-Allow-Origin', 'http://13.201.218.4')
        response.headers.add('Access-Control-Allow-Headers', 'Content-Type')
        response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
        response.headers.add('Access-Control-Allow-Credentials', 'true')
        return response

    try:
        skip_first_row = request.form.get('skipFirstRow', 'false').lower() == 'true'
        files = request.files

        if not files:
            return jsonify({"success": False, "message": "No files uploaded"}), 400

        all_dfs, errors, uploaded_months = [], [], []

        for month, file in files.items():
            logger.info(f"Processing month={month}, file={file.filename if file else 'None'}")

            if file and file.filename.lower().endswith(('.xlsx', '.xls')):
                try:
                    filename = secure_filename(file.filename)
                    file.stream.seek(0)
                    df = read_excel_file(file.stream, month, filename, skip_first_row)

                    if df is not None and not df.empty:
                        all_dfs.append(df)
                        uploaded_months.append(month)
                    else:
                        errors.append(f"No data found in {month} file")

                except Exception as e:
                    errors.append(f"Error processing {month}: {str(e)}")
                    logger.error(traceback.format_exc())
            else:
                errors.append(f"Invalid file type for {month}: {file.filename}")

        if not all_dfs:
            return jsonify({
                "success": False,
                "message": "No valid files could be processed",
                "errors": errors
            }), 400

        # Combine all dataframes
        final_df = pd.concat(all_dfs, ignore_index=True)

        # Create preview (first 5 rows)
        preview_data = []
        for _, row in final_df.head(5).iterrows():
            row_dict = {}
            for col, value in row.items():
                if pd.isna(value) or value in [None, float('inf'), -float('inf')]:
                    row_dict[col] = ""
                else:
                    row_dict[col] = str(value) if not isinstance(value, (int, float)) else value
            preview_data.append(row_dict)

        response_data = {
            "success": True,
            "message": f"Successfully processed {len(all_dfs)} files",
            "preview": preview_data,
            "uploaded_months": uploaded_months,
            "total_rows": len(final_df),
            "total_columns": len(final_df.columns),
        }

        if errors:
            response_data["warnings"] = errors

        return jsonify(response_data)

    except Exception as e:
        logger.error(traceback.format_exc())
        return jsonify({"success": False, "message": f"Server error: {str(e)}"}), 500


@api_bp.route("/download", methods=["POST", "OPTIONS"])
def download_file():
    """Download combined Excel file in memory"""
    logger.info("=== DOWNLOAD ENDPOINT CALLED ===")

    if request.method == "OPTIONS":
        response = jsonify({"status": "ok"})
        response.headers.add('Access-Control-Allow-Origin', 'http://13.201.218.4')
        response.headers.add('Access-Control-Allow-Headers', 'Content-Type')
        response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
        response.headers.add('Access-Control-Allow-Credentials', 'true')
        return response

    try:
        skip_first_row = request.form.get('skipFirstRow', 'false').lower() == 'true'
        files = request.files

        if not files:
            return jsonify({"success": False, "message": "No files uploaded"}), 400

        all_dfs, errors = [], []

        for month, file in files.items():
            if file and file.filename.lower().endswith(('.xlsx', '.xls')):
                try:
                    filename = secure_filename(file.filename)
                    file.stream.seek(0)
                    df = read_excel_file(file.stream, month, filename, skip_first_row)

                    if df is not None and not df.empty:
                        all_dfs.append(df)
                except Exception as e:
                    errors.append(f"Error processing {month}: {str(e)}")
                    logger.error(traceback.format_exc())
            else:
                errors.append(f"Invalid file type for {month}: {file.filename}")

        if not all_dfs:
            return jsonify({
                "success": False,
                "message": "No valid files to process for download",
                "errors": errors
            }), 400

        combined_df = pd.concat(all_dfs, ignore_index=True)

        # Write combined Excel to memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            combined_df.to_excel(writer, index=False, sheet_name="Combined Sales")

            workbook = writer.book
            worksheet = writer.sheets["Combined Sales"]

            header_format = workbook.add_format({
                "bold": True,
                "bg_color": "#D3D3D3",
                "border": 1
            })
            for col_num, value in enumerate(combined_df.columns.values):
                worksheet.write(0, col_num, value, header_format)

        output.seek(0)

        return send_file(
            output,
            download_name="Combined_Sales_Report.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True
        )

    except Exception as e:
        logger.error(traceback.format_exc())
        return jsonify({
            "success": False,
            "message": f"Error generating download file: {str(e)}"
        }), 500


@api_bp.route("/cumulative/health", methods=["GET"])
def cumulative_health():
    return jsonify({
        "status": "healthy",
        "service": "Cumulative Sales Data Processor",
        "endpoints": ["/api/process", "/api/download"]
    })


@api_bp.route("/cumulative/debug", methods=["GET"])
def cumulative_debug():
    return jsonify({
        "blueprint": "cumulative_api",
        "available_routes": [
            "POST /api/process - Process Excel files",
            "POST /api/download - Download combined Excel",
            "GET /api/cumulative/health - Health check",
            "GET /api/cumulative/debug - This debug endpoint"
        ],
        "status": "ready"
    })
