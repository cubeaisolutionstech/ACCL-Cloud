from flask import Blueprint, request, jsonify, send_file
from services.mapping_service import process_os_file
import pandas as pd
import io

os_bp = Blueprint("os_bp", __name__)

@os_bp.route("/process-os-file", methods=["POST"])
def api_process_os_file():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    sheet_name = request.form.get("sheet_name")
    header_row = int(request.form.get("header_row", 0))
    exec_code_col = request.form.get("exec_code_col")

    if not sheet_name or not exec_code_col:
        return jsonify({"error": "Missing parameters"}), 400

    try:
        # Read sheet into DataFrame
        df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)
    except Exception as e:
        return jsonify({"error": f"Failed to read Excel: {str(e)}"}), 500

    processed_df = process_os_file(df, exec_code_col)

    # Save to in-memory Excel file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        processed_df.to_excel(writer, index=False, sheet_name="Processed")
    output.seek(0)

    # Return as downloadable file
    return send_file(
        output,
        as_attachment=True,
        download_name="processed_os.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@os_bp.route("/save-os-file", methods=["POST"])
def save_os_file():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    filename = request.form.get("filename") or file.filename
    binary_data = file.read()

    from models.schema import OsFile
    from extensions import db

    # Check if already exists
    existing = OsFile.query.filter_by(filename=filename).first()
    if existing:
        db.session.delete(existing)
        db.session.commit()

    new_file = OsFile(filename=filename, file_data=binary_data)
    db.session.add(new_file)
    db.session.commit()

    return jsonify({"message": "OS file saved to database."})

@os_bp.route("/os-files", methods=["GET"])
def list_os_files():
    from models.schema import OsFile
    files = OsFile.query.order_by(OsFile.uploaded_at.desc()).all()
    return jsonify([
        {"id": f.id, "filename": f.filename, "uploaded_at": f.uploaded_at.isoformat()}
        for f in files
    ])

@os_bp.route("/os-files/<int:file_id>/download", methods=["GET"])
def download_os_file(file_id):
    from models.schema import OsFile
    import io
    file = OsFile.query.get_or_404(file_id)

    return send_file(
        io.BytesIO(file.file_data),
        download_name=file.filename if file.filename.endswith(".xlsx") else f"{file.filename}.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@os_bp.route("/os-files/<int:file_id>", methods=["DELETE"])
def delete_os_file(file_id):
    from models.schema import OsFile
    os_file = OsFile.query.get(file_id)
    if not os_file:
        return jsonify({"error": "File not found"}), 404

    from extensions import db
    db.session.delete(os_file)
    db.session.commit()

    return jsonify({"message": "File deleted"})
