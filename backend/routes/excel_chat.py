from flask import Blueprint, request, jsonify
import pandas as pd
import google.generativeai as genai
import numpy as np
from werkzeug.utils import secure_filename
import tempfile
import os

excel_chat = Blueprint("excel_chat", __name__)

genai.configure(api_key="AIzaSyDrvURpHhrOkNKxlunjnN7pDs8tfjCLXdU")
model = genai.GenerativeModel("gemini-1.5-flash")

@excel_chat.route('/upload_excel', methods=['POST'])
def upload_excel():
    try:
        file = request.files['file']
        sheet_name = request.form.get("sheet")
        header_index = request.form.get("header", type=int)
        remove_unnamed = request.form.get("removeUnnamed", "true") == "true"
        remove_last = request.form.get("removeLastRow", "true") == "true"

        filename = secure_filename(file.filename)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            file.save(tmp.name)
            xls = pd.ExcelFile(tmp.name)
            if sheet_name not in xls.sheet_names:
                return jsonify({"error": "Invalid sheet name"}), 400

            df = pd.read_excel(xls, sheet_name=sheet_name, header=header_index)

        if remove_unnamed:
            df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

        if remove_last:
            df = df.iloc[:-1]

        # Save DataFrame in memory (naive approach - can use Redis or file cache)
        request_id = os.urandom(8).hex()
        df.to_pickle(f"/tmp/{request_id}.pkl")

        return jsonify({
            "requestId": request_id,
            "columns": df.columns.tolist(),
            "shape": df.shape,
            "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()},
            "preview": df.head(10).to_dict(orient="records"),
            "sheetNames": xls.sheet_names
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@excel_chat.route("/ask_question", methods=["POST"])
def ask_question():
    try:
        data = request.json
        request_id = data["requestId"]
        question = data["question"]

        df = pd.read_pickle(f"/tmp/{request_id}.pkl")

        system_prompt = f"""
You are a Python data analyst working with a Pandas DataFrame called `df`.

DataFrame Info:
- Columns: {df.columns.tolist()}
- Shape: {df.shape}
- Data types: {dict(df.dtypes)}

IMPORTANT RULES:
1. Write ONLY ONE line of Python code that assigns the result to `result`
2. Always check if filtered DataFrames are empty before using .iloc[]
3. Handle potential errors using try/except logic
4. If no data found or error occurs, set result = 'Not found'
5. Use pandas methods like .sum(), .mean(), .nlargest(), .str.contains()

Question: {question}

Output only the Python code line starting with 'result =':
"""
        response = model.generate_content(system_prompt)
        python_code = response.text.strip()
        if "```python" in python_code:
            python_code = python_code.split("```python")[1].split("```")[0].strip()
        elif "```" in python_code:
            python_code = python_code.strip("```").strip()
        if not python_code.startswith("result ="):
            python_code = "result = " + python_code

        local_vars = {"df": df, "pd": pd, "np": np}
        exec(python_code, {}, local_vars)
        result = local_vars.get("result", "No result")

        result_type = type(result).__name__
        if isinstance(result, pd.DataFrame):
            result_json = result.to_dict(orient="records")
        elif isinstance(result, pd.Series):
            result_json = result.to_frame().to_dict(orient="records")
        else:
            result_json = str(result)

        return jsonify({
            "result": result_json,
            "resultType": result_type,
            "code": python_code
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500
