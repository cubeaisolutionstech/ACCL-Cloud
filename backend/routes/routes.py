from flask import request, jsonify, current_app ,Blueprint
import pandas as pd
import google.generativeai as genai
import numpy as np
import tempfile
import os
import shutil
import atexit

api1_bp = Blueprint('api1', __name__)

# Configure Gemini - you could load your API key from environment or config instead of hardcoding
genai.configure(api_key="AIzaSyDrvURpHhrOkNKxlunjnN7pDs8tfjCLXdU")
model = genai.GenerativeModel("gemini-1.5-flash")

# Keep track of temp dirs for cleanup
temp_dirs = []

def cleanup_temp_dirs():
    for temp_dir in temp_dirs:
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception as e:
            print(f"Error deleting temp directory {temp_dir}: {e}")

atexit.register(cleanup_temp_dirs)

@api1_bp.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    temp_dir = tempfile.mkdtemp()
    temp_dirs.append(temp_dir)
    temp_path = os.path.join(temp_dir, file.filename)
    
    try:
        file.save(temp_path)
        xls = pd.ExcelFile(temp_path)
        sheet_names = xls.sheet_names
        xls.close()
        
        return jsonify({
            'success': True,
            'sheet_names': sheet_names,
            'filename': file.filename
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        try:
            if os.path.exists(temp_path):
                os.unlink(temp_path)
        except Exception as e:
            print(f"Error deleting temp file {temp_path}: {e}")

@api1_bp.route('/analyze', methods=['POST'])
def analyze_data():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    temp_dir = tempfile.mkdtemp()
    temp_dirs.append(temp_dir)
    temp_path = os.path.join(temp_dir, file.filename)
    
    try:
        file.save(temp_path)
        xls = pd.ExcelFile(temp_path)
        
        header_option = request.form.get('header_option', 'Row 1 (0)')
        selected_sheet = request.form.get('selected_sheet')
        remove_unnamed = request.form.get('remove_unnamed', 'true') == 'true'
        remove_last_row = request.form.get('remove_last_row', 'true') == 'true'
        
        df = pd.read_excel(
            xls, 
            sheet_name=selected_sheet,
            header=None if header_option == 'No header' else int(header_option.split('(')[1].split(')')[0])
        )
        xls.close()
        
        if remove_unnamed:
            df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
        
        if remove_last_row:
            df = df.iloc[:-1]
        
        preview_rows = min(10, len(df))
        data_info = {
            'columns': [str(col) for col in df.columns.tolist()],
            'shape': df.shape,
            'dtypes': {str(col): str(dtype) for col, dtype in df.dtypes.items()},
            'preview': df.head(preview_rows).fillna('').to_dict('records')
        }
        
        return jsonify({
            'success': True,
            'data_info': data_info,
            'message': f"Uploaded: {file.filename} | Sheet: {selected_sheet} | Shape: {df.shape}"
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        try:
            if os.path.exists(temp_path):
                os.unlink(temp_path)
        except Exception as e:
            print(f"Error deleting temp file {temp_path}: {e}")

@api1_bp.route('/ask', methods=['POST'])
def ask_question():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    temp_dir = tempfile.mkdtemp()
    temp_dirs.append(temp_dir)
    temp_path = os.path.join(temp_dir, file.filename)
    
    try:
        file.save(temp_path)
        xls = pd.ExcelFile(temp_path)
        
        user_question = request.form.get('user_question')
        header_option = request.form.get('header_option', 'Row 1 (0)')
        selected_sheet = request.form.get('selected_sheet')
        remove_unnamed = request.form.get('remove_unnamed', 'true') == 'true'
        remove_last_row = request.form.get('remove_last_row', 'true') == 'true'
        
        df = pd.read_excel(
            xls, 
            sheet_name=selected_sheet,
            header=None if header_option == 'No header' else int(header_option.split('(')[1].split(')')[0])
        )
        xls.close()
        
        if remove_unnamed:
            df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed", na=False)]
        
        if remove_last_row:
            df = df.iloc[:-1]
        
        df.columns = df.columns.astype(str)
        
        system_prompt = f"""
You are a Python data analyst working with a Pandas DataFrame called `df`.

DataFrame Info:
- Columns: {df.columns.tolist()}
- Shape: {df.shape}
- Data types: {dict(df.dtypes)}

IMPORTANT RULES:
1. Write ONLY ONE line of Python code that assigns the result to `result`
2. Always check if filtered DataFrames are empty before using .iloc[]
3. Handle potential errors (use try/except logic in a single line if needed)
4. If no data found or error occurs, set result = 'Not found' or appropriate message
5. Use proper pandas methods (.sum(), .mean(), .max(), .min(), .count(), etc.)
6. For text matching, use .str.contains() with case=False and na=False for flexibility
7. Output format: result = [your_code_here]

Examples:
- result = df['column'].sum() if 'column' in df.columns else 'Column not found'
- result = df.nlargest(5, 'column') if 'column' in df.columns and not df.empty else 'No data'
- result = df[df['column'] > 100].head() if 'column' in df.columns else 'Column not found'

Question: {user_question}

Output only the Python code line starting with 'result =':
"""
        gemini_response = model.generate_content(system_prompt)
        python_code = gemini_response.text.strip()

        if "```python" in python_code:
            python_code = python_code.split("```python")[1].split("```")[0].strip()
        elif "```" in python_code:
            python_code = python_code.strip("```").strip()

        if not python_code.startswith("result ="):
            python_code = "result = " + python_code

        local_vars = {"df": df, "pd": pd, "np": np}
        exec(python_code, {}, local_vars)
        result = local_vars.get("result", "No result found.")

        if isinstance(result, (pd.DataFrame, pd.Series)):
            result_data = {
                'type': 'dataframe',
                'data': result.fillna('').to_dict('records') if isinstance(result, pd.DataFrame) else result.to_dict(),
                'shape': result.shape if hasattr(result, 'shape') else None
            }
        else:
            result_data = {
                'type': 'text',
                'data': str(result),
                'shape': None
            }

        return jsonify({
            'success': True,
            'result': result_data,
            'python_code': python_code,
            'question': user_question
        })
    except Exception as e:
        return jsonify({'error': str(e), 'python_code': python_code if 'python_code' in locals() else None}), 500
