from flask import Blueprint, request, jsonify
import pandas as pd
import os
from datetime import datetime
from werkzeug.utils import secure_filename

# Create the blueprint - this line is CRITICAL
upload_bp = Blueprint('upload', __name__)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Ensure upload folder exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@upload_bp.route('/upload', methods=['POST'])
def upload_file():
    """Handle file uploads"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        file_type = request.form.get('type', 'general')
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            # Add timestamp to avoid conflicts
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{timestamp}_{filename}"
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)
            
            # Read Excel file and get sheet names
            try:
                xl_file = pd.ExcelFile(filepath)
                sheet_names = xl_file.sheet_names
                xl_file.close()
                
                return jsonify({
                    'success': True,
                    'filename': filename,
                    'filepath': filepath,
                    'sheet_names': sheet_names,
                    'file_type': file_type,
                    'message': 'File uploaded successfully'
                })
            except Exception as e:
                # Clean up the uploaded file if Excel reading fails
                if os.path.exists(filepath):
                    os.remove(filepath)
                return jsonify({'error': f'Error reading Excel file: {str(e)}'}), 400
        
        return jsonify({'error': 'Invalid file type. Please upload .xlsx or .xls files only.'}), 400
        
    except Exception as e:
        return jsonify({'error': f'Upload failed: {str(e)}'}), 500

@upload_bp.route('/delete-file', methods=['POST'])
def delete_file():
    """Delete uploaded file"""
    try:
        data = request.json
        filepath = data.get('filepath')
        
        if not filepath:
            return jsonify({'error': 'No filepath provided'}), 400
        
        if os.path.exists(filepath):
            os.remove(filepath)
            return jsonify({'success': True, 'message': 'File deleted successfully'})
        else:
            return jsonify({'error': 'File not found'}), 404
            
    except Exception as e:
        return jsonify({'error': f'Delete failed: {str(e)}'}), 500

@upload_bp.route('/list-files', methods=['GET'])
def list_files():
    """List all uploaded files"""
    try:
        files = []
        if os.path.exists(UPLOAD_FOLDER):
            for filename in os.listdir(UPLOAD_FOLDER):
                if allowed_file(filename):
                    filepath = os.path.join(UPLOAD_FOLDER, filename)
                    file_stats = os.stat(filepath)
                    files.append({
                        'filename': filename,
                        'filepath': filepath,
                        'size': file_stats.st_size,
                        'created': datetime.fromtimestamp(file_stats.st_ctime).isoformat(),
                        'modified': datetime.fromtimestamp(file_stats.st_mtime).isoformat()
                    })
        
        return jsonify({'files': files})
        
    except Exception as e:
        return jsonify({'error': f'Failed to list files: {str(e)}'}), 500

# Debug print to confirm blueprint creation
print("📝 upload_bp blueprint created successfully")