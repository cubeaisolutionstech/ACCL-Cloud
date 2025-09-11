from flask import Blueprint, request, send_file, jsonify
from models.schema import BudgetFile, SalesFile, OsFile
import io
import pandas as pd
from services.ppt_generator import generate_branch_ppt

ppt_bp = Blueprint("ppt", __name__)

@ppt_bp.route("/generate-ppt", methods=["POST"])
def generate_ppt():
    data = request.get_json()
    budget_id = data.get("budget_id")
    sales_id = data.get("sales_id")
    os_prev_id = data.get("os_prev_id")
    os_curr_id = data.get("os_curr_id")

    if not all([budget_id, sales_id, os_prev_id, os_curr_id]):
        return jsonify({"error": "Missing file selections"}), 400

    # Get files from DB
    budget_file = BudgetFile.query.get(budget_id)
    sales_file = SalesFile.query.get(sales_id)
    os_prev_file = OsFile.query.get(os_prev_id)
    os_curr_file = OsFile.query.get(os_curr_id)

    if not all([budget_file, sales_file, os_prev_file, os_curr_file]):
        return jsonify({"error": "Invalid file IDs"}), 404

    # Convert to DataFrames
    df_budget = pd.read_excel(io.BytesIO(budget_file.file_data))
    df_sales = pd.read_excel(io.BytesIO(sales_file.file_data))
    df_os_prev = pd.read_excel(io.BytesIO(os_prev_file.file_data))
    df_os_curr = pd.read_excel(io.BytesIO(os_curr_file.file_data))

    # Generate PPT (you'll define this function)
    pptx_buffer = generate_branch_ppt(df_budget, df_sales, df_os_prev, df_os_curr)

    return send_file(
        pptx_buffer,
        download_name="branch_dashboard.pptx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
