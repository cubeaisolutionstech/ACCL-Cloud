from extensions import db
from datetime import datetime
from sqlalchemy import LargeBinary

# ========================
# Executive-related Models
# ========================

class Executive(db.Model):
    __tablename__ = 'executives'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)
    code = db.Column(db.String(50), nullable=True)

    customers = db.relationship("Customer", backref="executive", lazy=True)
    branches = db.relationship("BranchExecutiveMap", backref="executive", lazy=True)

# ========================
# Customer-related Models
# ========================

class Customer(db.Model):
    __tablename__ = 'customers'
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(50), nullable=False)
    name = db.Column(db.String(100), nullable=True)

    executive_id = db.Column(db.Integer, db.ForeignKey('executives.id'))

# ========================
# Branch and Region
# ========================

class Branch(db.Model):
    __tablename__ = 'branches'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)

    exec_mappings = db.relationship("BranchExecutiveMap", backref="branch", lazy=True)
    region_mappings = db.relationship("RegionBranchMap", backref="branch", lazy=True)

class Region(db.Model):
    __tablename__ = 'regions'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)

    branch_mappings = db.relationship("RegionBranchMap", backref="region", lazy=True)

class BranchExecutiveMap(db.Model):
    __tablename__ = 'branch_exec_map'
    id = db.Column(db.Integer, primary_key=True)
    branch_id = db.Column(db.Integer, db.ForeignKey('branches.id'), nullable=False)
    executive_id = db.Column(db.Integer, db.ForeignKey('executives.id'), nullable=False)

class RegionBranchMap(db.Model):
    __tablename__ = 'region_branch_map'
    id = db.Column(db.Integer, primary_key=True)
    region_id = db.Column(db.Integer, db.ForeignKey('regions.id'), nullable=False)
    branch_id = db.Column(db.Integer, db.ForeignKey('branches.id'), nullable=False)

# ========================
# Company and Product
# ========================

class Company(db.Model):
    __tablename__ = 'companies'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)

    product_mappings = db.relationship("CompanyProductMap", backref="company", lazy=True)

class Product(db.Model):
    __tablename__ = 'products'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)

    company_mappings = db.relationship("CompanyProductMap", backref="product", lazy=True)

class CompanyProductMap(db.Model):
    __tablename__ = 'company_product_map'
    id = db.Column(db.Integer, primary_key=True)
    company_id = db.Column(db.Integer, db.ForeignKey('companies.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('products.id'), nullable=False)

# ========================
# File Processing
# ========================

class BudgetFile(db.Model):
    __tablename__ = "budget_files"
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    file_data = db.Column(LargeBinary, nullable=False)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)

class SalesFile(db.Model):
    __tablename__ = "sales_files"
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    file_data = db.Column(db.LargeBinary, nullable=False)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)

class OsFile(db.Model):
    __tablename__ = "os_files"
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    file_data = db.Column(db.LargeBinary, nullable=False)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)

class LastYearSalesFile(db.Model):
    __tablename__ = 'last_year_sales_files'
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    file_data = db.Column(db.LargeBinary, nullable=False)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)