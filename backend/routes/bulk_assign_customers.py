from flask import Blueprint, request, jsonify
from models.schema import *
from extensions import db

bulk_bp = Blueprint("bulk", __name__)

@bulk_bp.route("/bulk-assign-customers", methods=["POST"])
def bulk_assign_customers():
    try:
        payload = request.get_json()
        data = payload.get("data", [])
        exec_name_col = payload.get("execNameCol")
        exec_code_col = payload.get("execCodeCol")
        cust_code_col = payload.get("custCodeCol")
        cust_name_col = payload.get("custNameCol")

        if not exec_name_col or not cust_code_col:
            return jsonify({"error": "Missing required column mappings"}), 400

        exec_map = {}
        for row in data:
            exec_name = str(row.get(exec_name_col, "")).strip()
            exec_code = str(row.get(exec_code_col, "")).strip() if exec_code_col else None
            cust_code = str(row.get(cust_code_col, "")).strip()
            cust_name = str(row.get(cust_name_col, "")).strip() if cust_name_col else None

            if not exec_name or not cust_code:
                continue

            # Find or create executive
            if exec_name not in exec_map:
                exec_obj = Executive.query.filter_by(name=exec_name).first()
                if not exec_obj:
                    exec_obj = Executive(name=exec_name, code=exec_code)
                    db.session.add(exec_obj)
                    db.session.flush()
                exec_map[exec_name] = exec_obj
            else:
                exec_obj = exec_map[exec_name]

            # Create or update customer
            cust = Customer.query.filter_by(code=cust_code).first()
            if not cust:
                cust = Customer(code=cust_code, name=cust_name or "", executive_id=exec_obj.id)
                db.session.add(cust)
            else:
                cust.executive_id = exec_obj.id
                if cust_name:
                    cust.name = cust_name

        db.session.commit()
        return jsonify({"message": "Bulk customer assignment completed."})

    except Exception as e:
        return jsonify({"error": str(e)}), 500
