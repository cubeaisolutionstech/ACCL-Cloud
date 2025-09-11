# routes/mapping_routes.py

from flask import Blueprint, request, jsonify
from services import mapping_service as svc
from models.schema import *
import logging  # <-- ADD THIS IMPORT

# Get a logger instance
logger = logging.getLogger(__name__)

mapping_bp = Blueprint("mapping", __name__)

### --------- Executive APIs ----------

@mapping_bp.route("/executives", methods=["GET"])
def get_all_executives():
    execs = svc.get_executives()
    return jsonify([{"id": e.id, "name": e.name, "code": e.code} for e in execs])

@mapping_bp.route("/executive", methods=["POST"])
def add_executive():
    data = request.json
    name = data.get("name")
    code = data.get("code")
    success, msg = svc.add_executive(name, code)
    return jsonify({"success": success, "message": msg})

@mapping_bp.route("/executive/<string:name>", methods=["DELETE"])
def remove_executive(name):
    success, count = svc.remove_executive(name)
    return jsonify({"success": success, "removed_customers": count})

@mapping_bp.route("/executive/customers", methods=["POST"])
def assign_customers():
    data = request.json
    exec_name = data.get("executive")
    customers = data.get("customers", [])
    count = svc.assign_customers_to_exec(exec_name, customers)
    return jsonify({"assigned": count})

@mapping_bp.route("/executive/<string:name>/customers", methods=["GET"])
def get_exec_customers(name):
    customers = svc.get_exec_customers(name)
    return jsonify([{"code": c.code, "name": c.name} for c in customers])

# ===============================================
# === THIS IS THE FUNCTION THAT HAS BEEN FIXED ===
# ===============================================
@mapping_bp.route("/executives-with-counts", methods=["GET"])
def get_executives_with_counts():
    try:
        execs = Executive.query.all()
        result = []

        for e in execs:
            customer_count = Customer.query.filter_by(executive_id=e.id).count()
            branch_count = BranchExecutiveMap.query.filter_by(executive_id=e.id).count()

            result.append({
                "name": e.name,
                "code": e.code,
                "customers": customer_count,
                "branches": branch_count
            })

        return jsonify(result)
    except Exception as e:
        # If any error occurs, log it and return a safe, empty array.
        logger.error(f"Error in /executives-with-counts: {str(e)}")
        return jsonify([]) # This prevents the frontend from crashing.
# ===============================================
# ===============================================

@mapping_bp.route("/customers")
def get_customers_by_executive():
    exec_name = request.args.get("executive")
    exec_obj = Executive.query.filter_by(name=exec_name).first()
    if not exec_obj:
        return jsonify([])

    customers = Customer.query.filter_by(executive_id=exec_obj.id).all()
    return jsonify([
        {"code": c.code, "name": c.name} for c in customers
    ])

@mapping_bp.route("/customers/unmapped")
def get_unmapped_customers():
    customers = Customer.query.filter_by(executive_id=None).all()
    return jsonify([
        {"code": c.code, "name": c.name} for c in customers
    ])

@mapping_bp.route("/remove-customer", methods=["POST"])
def remove_customers():
    data = request.get_json()
    exec_name = data.get("executive")
    customer_codes = data.get("customers", [])

    exec_obj = Executive.query.filter_by(name=exec_name).first()
    if not exec_obj:
        return jsonify({"error": "Executive not found"}), 404

    count = 0
    for code in customer_codes:
        cust = Customer.query.filter_by(code=code, executive_id=exec_obj.id).first()
        if cust:
            cust.executive_id = None
            count += 1

    db.session.commit()
    return jsonify({"removed": count})

@mapping_bp.route("/assign-customer", methods=["POST"])
def assign_customers_manually():
    data = request.get_json()
    exec_name = data.get("executive")
    customer_codes = data.get("customers", [])

    exec_obj = Executive.query.filter_by(name=exec_name).first()
    if not exec_obj:
        return jsonify({"error": "Executive not found"}), 404

    count = 0
    for code in customer_codes:
        cust = Customer.query.filter_by(code=code).first()
        if cust:
            cust.executive_id = exec_obj.id
            count += 1
        else:
            # create if not exist
            cust = Customer(code=code, name="", executive_id=exec_obj.id)
            db.session.add(cust)
            count += 1

    db.session.commit()
    return jsonify({"assigned": count})


### --------- Branch & Region APIs ----------

@mapping_bp.route("/branch", methods=["POST"])
def add_branch():
    data = request.json
    svc.add_branch(data.get("name"))
    return jsonify({"message": "Branch added"})

@mapping_bp.route("/region", methods=["POST"])
def add_region():
    data = request.json
    svc.add_region(data.get("name"))
    return jsonify({"message": "Region added"})

@mapping_bp.route("/map-executive-branch", methods=["POST"])
def map_exec_branch():
    data = request.json
    svc.map_exec_to_branch(data.get("executive"), data.get("branch"))
    return jsonify({"message": "Mapping added"})

@mapping_bp.route("/map-branch-region", methods=["POST"])
def map_branch_region():
    data = request.json
    svc.map_branch_to_region(data.get("branch"), data.get("region"))
    return jsonify({"message": "Mapping added"})

### --------- Company & Product APIs ----------

@mapping_bp.route("/company", methods=["POST"])
def add_company():
    data = request.json
    svc.add_company(data.get("name"))
    return jsonify({"message": "Company added"})

@mapping_bp.route("/product", methods=["POST"])
def add_product():
    data = request.json
    svc.add_product(data.get("name"))
    return jsonify({"message": "Product added"})

@mapping_bp.route("/map-company-product", methods=["POST"])
def map_product_company():
    data = request.json
    svc.map_product_to_company(data.get("product"), data.get("company"))
    return jsonify({"message": "Mapping added"})

### --------- Export/Backup API ----------

@mapping_bp.route("/export", methods=["GET"])
def export_all():
    return jsonify(svc.export_all_mappings())
