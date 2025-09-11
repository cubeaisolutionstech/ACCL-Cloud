from flask import Blueprint, request, jsonify
import pandas as pd

upload_tools_bp = Blueprint("upload_tools", __name__)

@upload_tools_bp.route("/upload-tools/sheet-names", methods=["POST"])
def get_sheet_names():
    file = request.files["file"]
    excel = pd.ExcelFile(file)
    return jsonify({"sheet_names": excel.sheet_names})


# In routes/upload_tools.py

@upload_tools_bp.route("/upload-tools/preview", methods=["POST"])
def preview_sheet():
    try:
        file = request.files["file"]
        sheet_name = request.form["sheet_name"]
        header_row = int(request.form.get("header_row", 0))
        
        # ✅ REWIND THE FILE STREAM TO THE BEGINNING
        file.seek(0) 

        df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)

        return jsonify({
            "columns": list(df.columns),
            "preview": df.fillna("").astype(str).head(10).to_dict(orient="records")
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
