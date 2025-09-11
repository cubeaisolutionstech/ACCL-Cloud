import mysql.connector

print("🔌 Trying manual connection...")
try:
    conn = mysql.connector.connect(
        host="accl-db.cxu048yyut0c.ap-south-1.rds.amazonaws.com",
        user="admin",
        password="accl1234accl",
        database="sales_data",
        port=3306,
        connection_timeout=5
    )
    print("✅ Connected successfully!")
    conn.close()
except Exception as e:
    print(f"❌ Failed: {e}")
