# routes/file_routes.py

from flask import Blueprint, request, jsonify, send_file
from werkzeug.utils import secure_filename
from config import Config
import os
import pandas as pd
from services import mapping_service as svc

file_bp = Blueprint("file", __name__)

ALLOWED_EXTENSIONS = {"xls", "xlsx"}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@file_bp.route("/upload/budget", methods=["POST"])
def upload_budget():
    file = request.files.get("file")
    if not file or not allowed_file(file.filename):
        return jsonify({"error": "Invalid or missing file"}), 400

    filename = secure_filename(file.filename)
    path = os.path.join(Config.UPLOAD_FOLDER, filename)
    file.save(path)

    sheet = request.form.get("sheet") or 0
    header = int(request.form.get("header", 0))

    try:
        df = pd.read_excel(path, sheet_name=sheet, header=header)
        processed_df = svc.process_budget_file(df)

        output_path = os.path.join(Config.PROCESSED_FOLDER, f"processed_{filename}")
        processed_df.to_excel(output_path, index=False)

        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@file_bp.route("/upload/sales", methods=["POST"])
def upload_sales():
    file = request.files.get("file")
    if not file or not allowed_file(file.filename):
        return jsonify({"error": "Invalid or missing file"}), 400

    filename = secure_filename(file.filename)
    path = os.path.join(Config.UPLOAD_FOLDER, filename)
    file.save(path)

    sheet = request.form.get("sheet") or 0
    header = int(request.form.get("header", 0))

    try:
        df = pd.read_excel(path, sheet_name=sheet, header=header)
        processed_df = svc.process_sales_file(df)

        output_path = os.path.join(Config.PROCESSED_FOLDER, f"processed_{filename}")
        processed_df.to_excel(output_path, index=False)

        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@file_bp.route("/upload/os", methods=["POST"])
def upload_os():
    file = request.files.get("file")
    if not file or not allowed_file(file.filename):
        return jsonify({"error": "Invalid or missing file"}), 400

    filename = secure_filename(file.filename)
    path = os.path.join(Config.UPLOAD_FOLDER, filename)
    file.save(path)

    sheet = request.form.get("sheet") or 0
    header = int(request.form.get("header", 0))

    try:
        df = pd.read_excel(path, sheet_name=sheet, header=header)
        processed_df = svc.process_os_file(df)

        output_path = os.path.join(Config.PROCESSED_FOLDER, f"processed_{filename}")
        processed_df.to_excel(output_path, index=False)

        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

