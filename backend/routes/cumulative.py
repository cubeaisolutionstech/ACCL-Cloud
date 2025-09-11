import os
import traceback
import logging
from flask import Blueprint, request, jsonify, send_file
import pandas as pd
from io import BytesIO
from werkzeug.utils import secure_filename
import mysql.connector
from datetime import datetime
import uuid

cumulative_bp = Blueprint('cumulative', __name__)

# Enhanced logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Database configuration
# db_config = {
#     'host': 'localhost',
#     'user': 'root',
#     'password': 'Nav@1234',
#     'database': 'sales_data'
# }
print("✅ Connected to DB")

months = [
    "April", "May", "June", "July",
    "August", "September", "October", "November",
    "December", "January", "February", "March"
]

UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_db_connection():
    print("🔌 Attempting DB connection")
    try:
        conn = mysql.connector.connect(
            host="accl-db.cxu048yyut0c.ap-south-1.rds.amazonaws.com",
            user="admin",
            password="accl1234accl",
            database="sales_data",
            port=3306,
            connection_timeout= 10 
        )
        if conn.is_connected():
            print("✅ DB is connected")
            return conn
        else:
            print("❌ DB not connected (is_connected returned False)")
            return None
    except mysql.connector.Error as err:
        print(f"❌ Connection error: {err}")
        return None
    
print("1")

def init_db():
    print("🛠️ init_db() called")
    conn = get_db_connection()
    if not conn:
        print("❌ No DB connection in init_db")
        return False
    
    try:
        cursor = conn.cursor()
        print("🧱 Creating table")
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS monthly_files (
            id VARCHAR(36) PRIMARY KEY,
            month VARCHAR(20) NOT NULL,
            filename VARCHAR(255) NOT NULL,
            upload_date DATETIME NOT NULL,
            file_data LONGBLOB NOT NULL,
            skip_first_row BOOLEAN NOT NULL
        )
        """)
        cursor.execute("CREATE INDEX idx_month ON monthly_files(month)")
        cursor.execute("CREATE INDEX idx_upload_date ON monthly_files(upload_date)")
        conn.commit()
        print("✅ Table created or already exists")
        return True
    except mysql.connector.Error as err:
        logger.error(f"Database initialization error: {err}")
        import traceback
        traceback.print_exc()
        print(f"❌ DB init error: {err}")
        return False
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()
print("2")

def check_db_tables():
    print("🔍 check_db_tables() called")
    conn = get_db_connection()
    if not conn:
        print("❌ No DB connection in check_db_tables")
        return False
    
    try:
        cursor = conn.cursor()
        cursor.execute("SHOW TABLES LIKE 'monthly_files'")
        result = cursor.fetchone()
        print(f"👀 Table exists? {result}")
        if not result:
            print("📁 Table doesn't exist, calling init_db()")
            return init_db()
        print("✅ Table exists")
        return True
    except mysql.connector.Error as err:
        logger.error(f"Table check error: {err}")
        return False
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()
print ("3")
if not check_db_tables():
    logger.error("Failed to initialize database tables")
print("4")
def read_excel_file(file_path, month, filename, skip_first_row=False):
    try:
        try:
            df = pd.read_excel(file_path, header=None if skip_first_row else 0, engine='openpyxl')
        except:
            df = pd.read_excel(file_path, header=None if skip_first_row else 0, engine='xlrd')

        if skip_first_row and not df.empty:
            df = df.iloc[1:]
            if not df.empty:
                df.columns = df.iloc[0]
                df = df.iloc[1:]
                df.reset_index(drop=True, inplace=True)

        if not df.empty:
            df.columns = [str(col).strip() for col in df.columns]
            df["Month"] = month
            df["Source File"] = filename
            return df
        return None
    except Exception as e:
        logger.error(f"Error reading {file_path}: {str(e)}")
        return None
print("5")
def store_file_in_db(month, filename, file_data, skip_first_row):
    conn = get_db_connection()
    if not conn:
        return None
    
    try:
        cursor = conn.cursor()
        file_id = str(uuid.uuid4())
        upload_date = datetime.now()
        
        if isinstance(file_data, str):
            file_data = file_data.encode('utf-8')
            
        cursor.execute(
            "INSERT INTO monthly_files (id, month, filename, upload_date, file_data, skip_first_row) "
            "VALUES (%s, %s, %s, %s, %s, %s)",
            (file_id, month, filename, upload_date, file_data, skip_first_row)
        )
        
        conn.commit()
        return file_id
    except mysql.connector.Error as err:
        logger.error(f"Error storing file in DB: {err}")
        return None
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()
print(6)
@cumulative_bp.route("/api/files", methods=["GET"])
def get_available_files():
    conn = None
    cursor = None
    try:
        conn = get_db_connection()
        if not conn:
            return jsonify({"success": False, "message": "Database connection failed"}), 500
            
        cursor = conn.cursor(dictionary=True)
        placeholders = ', '.join(['%s'] * len(months))
        query = f"""
            SELECT 
                m1.id, 
                m1.month, 
                m1.filename, 
                m1.upload_date, 
                m1.skip_first_row 
            FROM monthly_files m1
            JOIN (
                SELECT month, MAX(upload_date) as latest_date 
                FROM monthly_files 
                GROUP BY month
            ) m2 ON m1.month = m2.month AND m1.upload_date = m2.latest_date
            ORDER BY FIELD(m1.month, {placeholders})
        """
        cursor.execute(query, tuple(months))
        
        files = []
        for file in cursor.fetchall():
            file['upload_date'] = file['upload_date'].isoformat()
            files.append(file)
        
        return jsonify({
            "success": True,
            "files": files
        })
    except Exception as e:
        logger.error(f"Error retrieving files: {str(e)}")
        return jsonify({
            "success": False,
            "message": "Error retrieving files",
            "error": str(e)
        }), 500
    finally:
        if cursor:
            cursor.close()
        if conn and conn.is_connected():
            conn.close()
print("7")
@cumulative_bp.route("/api/process", methods=["POST"])
def process_files():
    conn = None
    try:
        skip_first_row = request.form.get('skipFirstRow', 'false').lower() == 'true'
        files = request.files
        all_dfs = []
        errors = []
        uploaded_months = []

        for month in months:
            if month in files:
                file = files[month]
                if file and file.filename.lower().endswith(('.xlsx', '.xls')):
                    try:
                        filename = secure_filename(file.filename)
                        file_path = os.path.join(UPLOAD_FOLDER, filename)
                        file.save(file_path)
                        
                        with open(file_path, 'rb') as f:
                            file_data = f.read()
                        
                        file_id = store_file_in_db(month, filename, file_data, skip_first_row)
                        if not file_id:
                            errors.append(f"Failed to store {month} file in database")
                            continue
                            
                        uploaded_months.append(month)
                        df = read_excel_file(file_path, month, filename, skip_first_row)
                        
                        if df is not None and not df.empty:
                            all_dfs.append(df)
                        else:
                            errors.append(f"File {filename} was empty or could not be read")

                        if os.path.exists(file_path):
                            os.remove(file_path)
                    except Exception as e:
                        errors.append(f"Error processing {month}: {str(e)}")
                        continue
                else:
                    errors.append(f"Invalid file for {month}: must be .xlsx or .xls")

        selected_months = request.form.getlist('months[]')
        if selected_months:
            conn = get_db_connection()
            if not conn:
                errors.append("Database connection failed")
            else:
                cursor = conn.cursor(dictionary=True)
                placeholders = ', '.join(['%s'] * len(selected_months))
                query = f"""
                    SELECT m1.id, m1.month, m1.filename, m1.upload_date, m1.file_data, m1.skip_first_row
                    FROM monthly_files m1
                    JOIN (
                        SELECT month, MAX(upload_date) as latest_date 
                        FROM monthly_files 
                        WHERE month IN ({placeholders})
                        GROUP BY month
                    ) m2 ON m1.month = m2.month AND m1.upload_date = m2.latest_date
                """
                cursor.execute(query, tuple(selected_months))
                db_files = cursor.fetchall()
                
                for file in db_files:
                    try:
                        file_path = os.path.join(UPLOAD_FOLDER, file['filename'])
                        with open(file_path, 'wb') as f:
                            f.write(file['file_data'])
                        
                        df = read_excel_file(file_path, file['month'], file['filename'], file['skip_first_row'])
                        if df is not None and not df.empty:
                            all_dfs.append(df)
                        else:
                            errors.append(f"Failed to read {file['month']} file")

                        if os.path.exists(file_path):
                            os.remove(file_path)
                    except Exception as e:
                        errors.append(f"Error processing {file['month']}: {str(e)}")
                        continue

        if not all_dfs:
            return jsonify({
                "success": False,
                "message": "No valid files could be processed.",
                "errors": errors
            }), 400

        combined_df = pd.concat(all_dfs, ignore_index=True)
        preview_data = combined_df.head().replace({pd.NA: None}).to_dict(orient="records")

        return jsonify({
            "success": True,
            "message": f"Successfully processed {len(all_dfs)} files",
            "preview": preview_data,
            "uploaded_months": uploaded_months,
            "warnings": errors if errors else None
        })
    except Exception as e:
        logger.error(f"Processing error: {str(e)}")
        return jsonify({
            "success": False,
            "message": f"Error processing files: {str(e)}",
            "errors": errors
        }), 500
    finally:
        if conn and conn.is_connected():
            conn.close()

@cumulative_bp.route("/api/download", methods=["POST"])
def download_file():
    conn = None
    try:
        skip_first_row = request.form.get('skipFirstRow', 'false').lower() == 'true'
        selected_months = request.form.getlist('months[]')
        all_dfs = []
        errors = []

        files = request.files
        for month in months:
            if month in files:
                file = files[month]
                if file and file.filename.lower().endswith(('.xlsx', '.xls')):
                    try:
                        filename = secure_filename(file.filename)
                        file_path = os.path.join(UPLOAD_FOLDER, filename)
                        file.save(file_path)
                        
                        df = read_excel_file(file_path, month, filename, skip_first_row)
                        if df is not None and not df.empty:
                            all_dfs.append(df)
                        else:
                            errors.append(f"File {filename} was empty or could not be read")

                        if os.path.exists(file_path):
                            os.remove(file_path)
                    except Exception as e:
                        errors.append(f"Error processing {month}: {str(e)}")
                        continue

        if selected_months:
            conn = get_db_connection()
            if not conn:
                errors.append("Database connection failed")
            else:
                cursor = conn.cursor(dictionary=True)
                placeholders = ', '.join(['%s'] * len(selected_months))
                query = f"""
                    SELECT m1.id, m1.month, m1.filename, m1.upload_date, m1.file_data, m1.skip_first_row
                    FROM monthly_files m1
                    JOIN (
                        SELECT month, MAX(upload_date) as latest_date 
                        FROM monthly_files 
                        WHERE month IN ({placeholders})
                        GROUP BY month
                    ) m2 ON m1.month = m2.month AND m1.upload_date = m2.latest_date
                """
                cursor.execute(query, tuple(selected_months))
                db_files = cursor.fetchall()
                
                for file in db_files:
                    try:
                        file_path = os.path.join(UPLOAD_FOLDER, file['filename'])
                        with open(file_path, 'wb') as f:
                            f.write(file['file_data'])
                        
                        df = read_excel_file(file_path, file['month'], file['filename'], file['skip_first_row'])
                        if df is not None and not df.empty:
                            all_dfs.append(df)
                        else:
                            errors.append(f"Failed to read {file['month']} file")

                        if os.path.exists(file_path):
                            os.remove(file_path)
                    except Exception as e:
                        errors.append(f"Error processing {file['month']}: {str(e)}")
                        continue

        if not all_dfs:
            return jsonify({
                "success": False,
                "message": "No valid files to process.",
                "errors": errors
            }), 400

        combined_df = pd.concat(all_dfs, ignore_index=True)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            combined_df.to_excel(writer, index=False, sheet_name="Combined Sales")
        excel_data = output.getvalue()

        return send_file(
            BytesIO(excel_data),
            download_name="Combined_Sales_Report.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True
        )
    except Exception as e:
        logger.error(f"Download error: {str(e)}")
        return jsonify({
            "success": False,
            "message": f"Error generating file: {str(e)}",
            "errors": errors
        }), 500
    finally:
        if conn and conn.is_connected():
            conn.close() 