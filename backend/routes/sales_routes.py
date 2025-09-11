from flask import Blueprint, request, jsonify, send_file
import pandas as pd
from io import BytesIO
import base64
from services.mapping_service import process_sales_file
from models.schema import *
import io

sales_bp = Blueprint("sales", __name__)

@sales_bp.route("/upload-sales-file", methods=["POST"])
def upload_sales_file():
    file = request.files["file"]
    sheet_name = request.form["sheet_name"]
    header_row = int(request.form["header_row"])

    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)

    # Required mappings
    exec_code_col = request.form["exec_code_col"]
    exec_name_col = request.form.get("exec_name_col") or None
    product_col = request.form.get("product_col") or None
    unit_col = request.form.get("unit_col") or None
    quantity_col = request.form.get("quantity_col") or None
    value_col = request.form.get("value_col") or None

    processed = process_sales_file(
        df, exec_code_col, product_col, exec_name_col,
        unit_col, quantity_col, value_col
    )

    output = BytesIO()
    processed.to_excel(output, index=False)
    output.seek(0)
    encoded_excel = base64.b64encode(output.read()).decode("utf-8")

    return {
        "preview": processed.head(10).fillna("").to_dict(orient="records"),
        "columns": list(processed.columns),
        "file_data": encoded_excel
    }

@sales_bp.route("/save-sales-file", methods=["POST"])
def save_sales_file():
    data = request.get_json()
    base64_excel = data.get("file_data")
    filename = data.get("filename", "Processed_Sales.xlsx")

    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"


    if not base64_excel:
        return jsonify({"error": "No file data provided"}), 400

    from base64 import b64decode
    from models.schema import SalesFile
    from extensions import db

    binary_data = b64decode(base64_excel)
    sales_file = SalesFile(filename=filename, file_data=binary_data)
    db.session.add(sales_file)
    db.session.commit()

    return jsonify({"message": "File saved", "id": sales_file.id})

@sales_bp.route("/sales-files", methods=["GET"])
def list_sales_files():
    files = SalesFile.query.order_by(SalesFile.uploaded_at.desc()).all()
    return jsonify([
        {"id": f.id, "filename": f.filename, "uploaded_at": f.uploaded_at.isoformat()}
        for f in files
    ])

@sales_bp.route("/sales-files/<int:file_id>/download", methods=["GET"])
def download_sales_file(file_id):
    file = SalesFile.query.get_or_404(file_id)
    return send_file(
        io.BytesIO(file.file_data),
        download_name=file.filename,
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@sales_bp.route("/sales-files/<int:file_id>", methods=["DELETE"])
def delete_sales_file(file_id):
    from models.schema import SalesFile
    os_file = SalesFile.query.get(file_id)
    if not os_file:
        return jsonify({"error": "File not found"}), 404

    from extensions import db
    db.session.delete(os_file)
    db.session.commit()

    return jsonify({"message": "File deleted"})

@sales_bp.route("/sales-files/<int:file_id>/data", methods=["GET"])
def get_sales_file_data(file_id):
    from utils.file_utils import read_excel_from_binary
    file = SalesFile.query.get(file_id)
    if not file:
        return jsonify({"error": "Sales file not found"}), 404

    try:
        df = read_excel_from_binary(file.file_data)
        return jsonify(df.head(100).to_dict(orient="records"))
    except Exception as e:
        return jsonify({"error": f"Failed to parse file: {str(e)}"}), 500
