from flask import Blueprint, request, jsonify
from services import mapping_service as svc
from models.schema import Company, Product
from extensions import db
import pandas as pd
from io import BytesIO

company_product_bp = Blueprint("company_product", __name__)

# Add new company
@company_product_bp.route("/company", methods=["POST"])
def add_company():
    data = request.json
    name = data.get("name")
    svc.add_company(name)
    return jsonify({"success": True})

# Add new product
@company_product_bp.route("/product", methods=["POST"])
def add_product():
    data = request.json
    name = data.get("name")
    svc.add_product(name)
    return jsonify({"success": True})

# Get all companies
@company_product_bp.route("/companies", methods=["GET"])
def get_companies():
    companies = svc.get_all_companies()
    return jsonify([{"id": c.id, "name": c.name} for c in companies])

# Get all products
@company_product_bp.route("/products", methods=["GET"])
def get_products():
    products = svc.get_all_products()
    return jsonify([{"id": p.id, "name": p.name} for p in products])

# Get products mapped to a company
@company_product_bp.route("/company/<string:name>/products", methods=["GET"])
def get_products_for_company(name):
    products = svc.get_company_products(name)
    return jsonify(products)

# Map multiple products to a company (overwrite)
@company_product_bp.route("/map-company-products", methods=["POST"])
def map_products_to_company():
    data = request.get_json()
    company = data.get("company")
    products = data.get("products", [])
    count = svc.map_company_products(company, products)
    return jsonify({"mapped": count})

# Get all current company-product mappings
@company_product_bp.route("/company-product-mappings", methods=["GET"])
def get_all_company_product_mappings():
    mappings = svc.get_all_company_product_mappings()
    return jsonify(mappings)

# ------------------ Deletion APIs ------------------

@company_product_bp.route("/product/<string:name>", methods=["DELETE"])
def delete_product(name):
    product = Product.query.filter_by(name=name).first()
    if not product:
        return jsonify({"success": False, "message": "Product not found"}), 404

    # Remove all mappings first
    db.session.query(svc.CompanyProductMap).filter_by(product_id=product.id).delete()
    db.session.delete(product)
    db.session.commit()
    return jsonify({"success": True})

@company_product_bp.route("/company/<string:name>", methods=["DELETE"])
def delete_company(name):
    company = Company.query.filter_by(name=name).first()
    if not company:
        return jsonify({"success": False, "message": "Company not found"}), 404

    db.session.query(svc.CompanyProductMap).filter_by(company_id=company.id).delete()
    db.session.delete(company)
    db.session.commit()
    return jsonify({"success": True})

# Upload Excel file with company-product data
@company_product_bp.route("/upload-company-product-file", methods=["POST"])
def upload_company_product_file():
    file = request.files.get("file")
    sheet_name = request.form.get("sheet_name")
    header_row = int(request.form.get("header_row", 0))
    company_col = request.form.get("company_col")
    product_col = request.form.get("product_col")

    df = pd.read_excel(BytesIO(file.read()), sheet_name=sheet_name, header=header_row)

    count = svc.upload_company_product_from_df(df, company_col, product_col)
    return jsonify({"message": f"Processed {count} new mappings."})


