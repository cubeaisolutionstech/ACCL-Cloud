# Updated main Flask app file with direct file processing (no uploads folder needed)
from flask import Flask, jsonify, request
from flask_cors import CORS
from config import Config
from extensions import db
import logging
import pandas as pd
from io import StringIO, BytesIO

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Your existing imports...
from routes.mapping_routes import mapping_bp
from routes.bulk_assign_customers import bulk_bp
from routes.file_routes import file_bp
from routes.branch_region_routes import branch_region_bp
from routes.company_product_routes import company_product_bp
from routes.budget_routes import budget_bp
from routes.upload_tools import upload_tools_bp
from routes.sales_routes import sales_bp
from routes.os_processing_routes import os_bp
from routes.branch_routes import branch_bp
from routes.executive_routes import executive_bp

# This is your cumulative routes - make sure this file exists and has the right exports
from routes.cumulative_routes import api_bp

# All your auditor routes...
from routes.auditor.auditor import auditor_bp
from routes.auditor.combined_data import combined_bp
from routes.auditor.data_routes import data_bp
from routes.auditor.ero_pw import ero_pw_bp
from routes.auditor.process_routes import process_bp
from routes.auditor.product import product_bp
from routes.auditor.region import region_bp
from routes.auditor.sales import sales1_bp
from routes.auditor.salesmonthwise import salesmonthwise_bp
from routes.auditor.ts_pw import ts_pw_bp
from routes.auditor.upload_routes import upload_bp

from routes.dashboard.main_routes import main_bp
from routes.routes import api1_bp

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)
    
    # Remove upload folder config since we're processing files directly
    # No longer needed for direct file processing
    app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
    
    db.init_app(app)
    
    # FIXED CORS configuration - Allow both 3000 and 3001 ports and any local development
    CORS(app, resources={
        r"/api/*": {
            "origins": [
                "http://localhost:3000", 
                "http://localhost:3001",  # Added this line
                "http://127.0.0.1:3000", 
                "http://127.0.0.1:3001",  # Added this line
                "http://localhost:*",     # Allow any localhost port
                "http://127.0.0.1:*",      # Allow any 127.0.0.1 port
                "http://13.201.218.4"
            ],
            "methods": ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
            "allow_headers": ["Content-Type", "Authorization", "Accept"],
            "supports_credentials": True,
            "expose_headers": ["Content-Disposition"]  # For file downloads
        }
    }, supports_credentials=True)

    # Add CORS headers manually for additional safety
    @app.after_request
    def after_request(response):
        origin = request.headers.get('Origin')
        if origin and (
            origin.startswith('http://localhost:') or 
            origin.startswith('http://127.0.0.1:') or
            origin == "http://13.201.218.4"
        ):
            response.headers.add('Access-Control-Allow-Origin', origin)
            response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization,Accept')
            response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
            response.headers.add('Access-Control-Allow-Credentials', 'true')
            response.headers.add('Access-Control-Expose-Headers', 'Content-Disposition')
        return response

    # Direct file processing routes
    @app.route('/api/process-file-direct', methods=['POST'])
    def process_file_direct():
        """Process uploaded files directly in memory without saving to disk"""
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400
        
        try:
            # Reset file pointer to beginning
            file.seek(0)
            
            # For CSV files
            if file.filename.lower().endswith('.csv'):
                df = pd.read_csv(file)
                logger.info(f"Processed CSV file: {file.filename}, Rows: {len(df)}")
                return jsonify({
                    "status": "success",
                    "filename": file.filename,
                    "rows": len(df),
                    "columns": list(df.columns),
                    "data_preview": df.head(5).to_dict('records') if len(df) > 0 else [],
                    "message": "CSV file processed successfully"
                })
            
            # For Excel files
            elif file.filename.lower().endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file)
                logger.info(f"Processed Excel file: {file.filename}, Rows: {len(df)}")
                return jsonify({
                    "status": "success",
                    "filename": file.filename,
                    "rows": len(df),
                    "columns": list(df.columns),
                    "data_preview": df.head(5).to_dict('records') if len(df) > 0 else [],
                    "message": "Excel file processed successfully"
                })
            
            # For JSON files
            elif file.filename.lower().endswith('.json'):
                import json
                content = file.read().decode('utf-8')
                data = json.loads(content)
                logger.info(f"Processed JSON file: {file.filename}")
                return jsonify({
                    "status": "success",
                    "filename": file.filename,
                    "data": data,
                    "message": "JSON file processed successfully"
                })
            
            # For text files
            elif file.filename.lower().endswith('.txt'):
                content = file.read().decode('utf-8')
                lines = content.split('\n')
                logger.info(f"Processed text file: {file.filename}, Lines: {len(lines)}")
                return jsonify({
                    "status": "success",
                    "filename": file.filename,
                    "content": content,
                    "lines": len(lines),
                    "preview": lines[:10] if len(lines) > 10 else lines,
                    "message": "Text file processed successfully"
                })
            
            else:
                return jsonify({
                    "error": "Unsupported file type", 
                    "supported_types": ["csv", "xlsx", "xls", "json", "txt"]
                }), 400
                
        except Exception as e:
            logger.error(f"Error processing file {file.filename}: {str(e)}")
            return jsonify({"error": f"Error processing file: {str(e)}"}), 500

    @app.route('/api/upload-stream', methods=['POST'])
    def upload_stream():
        """Stream process files line by line for large files"""
        file = request.files.get('file')
        if not file:
            return jsonify({"error": "No file provided"}), 400
        
        try:
            file.seek(0)
            line_count = 0
            processed_data = []
            
            # Process file line by line
            for line in file:
                line_str = line.decode('utf-8').strip()
                if line_str:  # Skip empty lines
                    processed_data.append(line_str)
                    line_count += 1
                    
                    # Optional: limit processing for demo (remove this in production)
                    if line_count > 1000:  # Process first 1000 lines
                        break
            
            logger.info(f"Stream processed file: {file.filename}, Lines: {line_count}")
            return jsonify({
                "status": "success",
                "filename": file.filename,
                "lines_processed": line_count,
                "sample_data": processed_data[:5],
                "message": f"Successfully streamed and processed {line_count} lines"
            })
            
        except Exception as e:
            logger.error(f"Error streaming file {file.filename}: {str(e)}")
            return jsonify({"error": str(e)}), 500

    @app.route('/api/process-csv-direct', methods=['POST'])
    def process_csv_direct():
        """Specialized CSV processing endpoint"""
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
        
        file = request.files['file']
        if not file.filename.lower().endswith('.csv'):
            return jsonify({"error": "Only CSV files are supported"}), 400
        
        try:
            file.seek(0)
            
            # Read CSV with various options
            df = pd.read_csv(file, 
                            encoding='utf-8',
                            skipinitialspace=True,
                            na_values=['', 'NULL', 'null', 'N/A', 'n/a'])
            
            # Basic data analysis
            analysis = {
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "columns": list(df.columns),
                "data_types": df.dtypes.astype(str).to_dict(),
                "missing_values": df.isnull().sum().to_dict(),
                "numeric_columns": df.select_dtypes(include=['number']).columns.tolist(),
                "text_columns": df.select_dtypes(include=['object']).columns.tolist()
            }
            
            # Add summary statistics for numeric columns
            if analysis['numeric_columns']:
                analysis['numeric_summary'] = df[analysis['numeric_columns']].describe().to_dict()
            
            logger.info(f"CSV analysis completed for {file.filename}")
            return jsonify({
                "status": "success",
                "filename": file.filename,
                "analysis": analysis,
                "data_sample": df.head(10).to_dict('records'),
                "message": "CSV file analyzed successfully"
            })
            
        except Exception as e:
            logger.error(f"Error analyzing CSV {file.filename}: {str(e)}")
            return jsonify({"error": f"Error analyzing CSV: {str(e)}"}), 500

    # Register all your existing blueprints
    app.register_blueprint(mapping_bp, url_prefix="/api")
    app.register_blueprint(bulk_bp, url_prefix="/api")
    app.register_blueprint(file_bp, url_prefix="/api")
    app.register_blueprint(branch_region_bp, url_prefix="/api")
    app.register_blueprint(company_product_bp, url_prefix="/api")
    app.register_blueprint(budget_bp, url_prefix="/api")
    app.register_blueprint(upload_tools_bp, url_prefix="/api")
    app.register_blueprint(sales_bp, url_prefix="/api")
    app.register_blueprint(os_bp, url_prefix="/api")
    app.register_blueprint(branch_bp, url_prefix='/api/branch')
    app.register_blueprint(executive_bp, url_prefix='/api/executive')
    
    # This is your cumulative routes blueprint
    app.register_blueprint(api_bp, url_prefix="/api")
    logger.info(f"Registered cumulative API blueprint at /api")
    
    # All auditor blueprints
    app.register_blueprint(upload_bp, url_prefix='/api')
    app.register_blueprint(process_bp, url_prefix='/api')
    app.register_blueprint(data_bp, url_prefix='/api')
    app.register_blueprint(auditor_bp, url_prefix='/api')        
    app.register_blueprint(sales1_bp, url_prefix='/api')
    app.register_blueprint(region_bp, url_prefix='/api/region')
    app.register_blueprint(product_bp, url_prefix='/api/product')
    app.register_blueprint(ts_pw_bp, url_prefix='/api/ts-pw')
    app.register_blueprint(combined_bp, url_prefix='/api/combined')  
    app.register_blueprint(salesmonthwise_bp, url_prefix='/api')
    app.register_blueprint(ero_pw_bp, url_prefix='/api/ero-pw')
    app.register_blueprint(main_bp)
    app.register_blueprint(api1_bp, url_prefix='/api')
    
    # Add debugging routes
    @app.route('/health')
    def health_check():
        return jsonify({"status": "healthy", "message": "Flask server is running"})
    
    @app.route('/api/debug/routes')
    def list_routes():
        """Debug endpoint to list all registered routes"""
        routes = []
        for rule in app.url_map.iter_rules():
            methods = ','.join(sorted(rule.methods - {'HEAD', 'OPTIONS'}))
            routes.append({
                'endpoint': rule.endpoint,
                'methods': methods,
                'rule': str(rule)
            })
        return jsonify(routes)
    
    # Test endpoint for CORS
    @app.route('/api/test-cors', methods=['GET', 'POST', 'OPTIONS'])
    def test_cors():
        if request.method == 'OPTIONS':
            return jsonify({"status": "CORS preflight OK"})
        return jsonify({
            "status": "CORS test successful",
            "origin": request.headers.get('Origin', 'No origin header'),
            "method": request.method
        })
    
    # File upload test endpoint
    @app.route('/api/test-upload', methods=['POST'])
    def test_upload():
        """Test endpoint for file uploads"""
        if 'file' not in request.files:
            return jsonify({"error": "No file in request"}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400
        
        return jsonify({
            "message": "File upload test successful",
            "filename": file.filename,
            "content_type": file.content_type,
            "file_size": len(file.read())
        })
    
    # Add error handlers with CORS headers
    @app.errorhandler(404)
    def not_found(error):
        response = jsonify({
            "error": "Endpoint not found", 
            "message": str(error),
            "available_endpoints": "Visit /api/debug/routes to see all available endpoints"
        })
        response.status_code = 404
        return response
    
    @app.errorhandler(500)
    def internal_error(error):
        logger.error(f"Internal server error: {str(error)}")
        response = jsonify({"error": "Internal server error", "message": str(error)})
        response.status_code = 500
        return response
    
    @app.errorhandler(413)
    def file_too_large(error):
        response = jsonify({
            "error": "File too large", 
            "message": "File size exceeds the maximum allowed limit (16MB)"
        })
        response.status_code = 413
        return response
    
    return app

if __name__ == "__main__":
    app = create_app()
    
    # Print all registered routes for debugging
    print("\n" + "="*70)
    print("REGISTERED ROUTES:")
    print("="*70)
    cumulative_routes_found = False
    direct_processing_routes = 0
    
    for rule in app.url_map.iter_rules():
        methods = ','.join(sorted(rule.methods - {'HEAD', 'OPTIONS'}))
        print(f"{rule.endpoint:35} {methods:20} {rule}")
        
        if 'process' in str(rule) or 'download' in str(rule):
            cumulative_routes_found = True
            print(f"  --> CUMULATIVE ROUTE FOUND: {rule}")
        
        if 'direct' in str(rule) or 'stream' in str(rule):
            direct_processing_routes += 1
            print(f"  --> DIRECT PROCESSING ROUTE: {rule}")
    
    if not cumulative_routes_found:
        print("\n⚠️  WARNING: No cumulative routes found! Check cumulative_routes.py import")
    
    print(f"\n✅ Direct file processing routes added: {direct_processing_routes}")
    print("="*70)
    print(f"🚀 Server starting on http://0.0.0.0:5000")
    print(f"🌐 CORS enabled for localhost ports 3000, 3001, and others")
    print(f"🔍 Debug routes available at /api/debug/routes")
    print(f"💊 Health check at /health")
    print(f"🧪 CORS test at /api/test-cors")
    print(f"📁 Direct file processing at:")
    print(f"   - /api/process-file-direct (general file processing)")
    print(f"   - /api/upload-stream (streaming for large files)")
    print(f"   - /api/process-csv-direct (specialized CSV processing)")
    print(f"   - /api/test-upload (upload testing)")
    print("="*70 + "\n")
    
    app.app_context().push()
    app.run(debug=True, port=5000, host='0.0.0.0')
