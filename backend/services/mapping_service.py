# services/mapping_service.py

from models.schema import *
from extensions import db
from sqlalchemy.orm import joinedload
import pandas as pd
import io

### --------- Executive & Customer Logic ----------

def add_executive(name, code=None):
    if Executive.query.filter_by(name=name).first():
        return False, "Executive already exists"
    exec = Executive(name=name, code=code)
    db.session.add(exec)
    db.session.commit()
    return True, "Executive added"

def remove_executive(name):
    exec = Executive.query.filter_by(name=name).first()
    if not exec:
        return False, 0

    # Unlink customers
    count = Customer.query.filter_by(executive_id=exec.id).delete()

    # Remove branch mappings
    BranchExecutiveMap.query.filter_by(executive_id=exec.id).delete()

    db.session.delete(exec)
    db.session.commit()
    return True, count

def assign_customers_to_exec(exec_name, customer_codes):
    exec = Executive.query.filter_by(name=exec_name).first()
    if not exec:
        return 0
    count = 0
    for code in customer_codes:
        cust = Customer.query.filter_by(code=code).first()
        if cust:
            cust.executive_id = exec.id
        else:
            db.session.add(Customer(code=code, executive_id=exec.id))
        count += 1
    db.session.commit()
    return count

def get_executives():
    return Executive.query.all()

def get_exec_customers(exec_name):
    exec = Executive.query.filter_by(name=exec_name).first()
    return exec.customers if exec else []


### --------- Branch & Region Logic ----------

def add_branch(name):
    if not Branch.query.filter_by(name=name).first():
        db.session.add(Branch(name=name))
        db.session.commit()

def add_region(name):
    if not Region.query.filter_by(name=name).first():
        db.session.add(Region(name=name))
        db.session.commit()

def map_exec_to_branch(exec_name, branch_name):
    exec = Executive.query.filter_by(name=exec_name).first()
    branch = Branch.query.filter_by(name=branch_name).first()
    if exec and branch:
        if not BranchExecutiveMap.query.filter_by(executive_id=exec.id, branch_id=branch.id).first():
            db.session.add(BranchExecutiveMap(executive_id=exec.id, branch_id=branch.id))
            db.session.commit()

def map_branch_to_region(branch_name, region_name):
    branch = Branch.query.filter_by(name=branch_name).first()
    region = Region.query.filter_by(name=region_name).first()
    if branch and region:
        if not RegionBranchMap.query.filter_by(branch_id=branch.id, region_id=region.id).first():
            db.session.add(RegionBranchMap(branch_id=branch.id, region_id=region.id))
            db.session.commit()

### --------- Company & Product Logic ----------

def add_company(name):
    if not Company.query.filter_by(name=name).first():
        db.session.add(Company(name=name))
        db.session.commit()

def add_product(name):
    if not Product.query.filter_by(name=name).first():
        db.session.add(Product(name=name))
        db.session.commit()

def get_all_products():
    return Product.query.all()

def get_all_companies():
    return Company.query.all()

def get_company_products(company_name):
    company = Company.query.filter_by(name=company_name).first()
    if not company:
        return []
    mappings = CompanyProductMap.query.filter_by(company_id=company.id).all()
    product_ids = [m.product_id for m in mappings]
    products = Product.query.filter(Product.id.in_(product_ids)).all()
    return [p.name for p in products]

def map_company_products(company_name, product_names):
    company = Company.query.filter_by(name=company_name).first()
    if not company:
        return 0

    # remove previous mappings
    CompanyProductMap.query.filter_by(company_id=company.id).delete()

    count = 0
    for product_name in product_names:
        product = Product.query.filter_by(name=product_name).first()
        if product:
            db.session.add(CompanyProductMap(company_id=company.id, product_id=product.id))
            count += 1

    db.session.commit()
    return count

def get_all_company_product_mappings():
    companies = Company.query.all()
    result = []
    for company in companies:
        mappings = CompanyProductMap.query.filter_by(company_id=company.id).all()
        product_ids = [m.product_id for m in mappings]
        product_names = [p.name for p in Product.query.filter(Product.id.in_(product_ids)).all()]
        result.append({
            "company": company.name,
            "products": product_names,
            "count": len(product_names)
        })
    return result

def upload_company_product_from_df(df, company_col, product_col):
    count = 0
    for _, row in df.iterrows():
        company_name = str(row.get(company_col)).strip()
        product_name = str(row.get(product_col)).strip()

        if not company_name or not product_name:
            continue

        company = Company.query.filter_by(name=company_name).first()
        if not company:
            company = Company(name=company_name)
            db.session.add(company)
            db.session.flush()

        product = Product.query.filter_by(name=product_name).first()
        if not product:
            product = Product(name=product_name)
            db.session.add(product)
            db.session.flush()

        exists = CompanyProductMap.query.filter_by(
            company_id=company.id, product_id=product.id
        ).first()

        if not exists:
            db.session.add(CompanyProductMap(company_id=company.id, product_id=product.id))
            count += 1

    db.session.commit()
    return count

def count_companies_using_product(product_name):
    product = Product.query.filter_by(name=product_name).first()
    if not product:
        return 0
    return CompanyProductMap.query.filter_by(product_id=product.id).count()

def count_products_for_company(company_name):
    company = Company.query.filter_by(name=company_name).first()
    if not company:
        return 0
    return CompanyProductMap.query.filter_by(company_id=company.id).count()



### --------- Utility: Export Mappings ----------

def export_all_mappings():
    data = {
        "executives": [{"name": e.name, "code": e.code} for e in Executive.query.all()],
        "customers": [{"code": c.code, "name": c.name, "executive": c.executive.name if c.executive else None} for c in Customer.query.options(joinedload(Customer.executive)).all()],
        "branches": [b.name for b in Branch.query.all()],
        "regions": [r.name for r in Region.query.all()],
        "branch_exec_map": [{"branch": b.branch.name, "executive": b.executive.name} for b in BranchExecutiveMap.query.options(joinedload(BranchExecutiveMap.branch), joinedload(BranchExecutiveMap.executive)).all()],
        "region_branch_map": [{"region": r.region.name, "branch": r.branch.name} for r in RegionBranchMap.query.options(joinedload(RegionBranchMap.region), joinedload(RegionBranchMap.branch)).all()],
        "companies": [c.name for c in Company.query.all()],
        "products": [p.name for p in Product.query.all()],
        "company_product_map": [{"company": c.company.name, "product": c.product.name} for c in CompanyProductMap.query.options(joinedload(CompanyProductMap.company), joinedload(CompanyProductMap.product)).all()],
    }
    return data


### --------- Excel Processing ----------

def normalize_customer_code(code):
    import pandas as pd  # ensure pandas is imported

    if pd.isna(code):
        return ""
    
    code_str = str(code).strip()
    try:
        if '.' in code_str and code_str.replace('.', '').replace('-', '').isdigit():
            float_val = float(code_str)
            if float_val == int(float_val):
                return str(int(float_val))
    except (ValueError, OverflowError):
        pass
    
    return code_str

def get_exec_by_customer_code(customer_code):
    cust = Customer.query.filter_by(code=customer_code).first()
    if cust and cust.executive:
        return cust.executive.name
    return None

def get_exec_code(exec_name):
    exec_obj = Executive.query.filter_by(name=exec_name).first()
    return exec_obj.code if exec_obj else ""

def get_branches_for_executive(exec_name):
    exec_obj = Executive.query.filter_by(name=exec_name).first()
    if not exec_obj:
        return ""
    
    branch_maps = BranchExecutiveMap.query.filter_by(executive_id=exec_obj.id).all()
    branches = [b.branch.name for b in branch_maps if b.branch]
    return ", ".join(sorted(set(branches)))

def get_region_for_branch(branch_name):
    branch_obj = Branch.query.filter_by(name=branch_name).first()
    if not branch_obj:
        return ""
    
    region_map = RegionBranchMap.query.filter_by(branch_id=branch_obj.id).first()
    if region_map and region_map.region:
        return region_map.region.name
    
    return ""


def process_budget_file(df, customer_col, exec_code_col, exec_name_col, branch_col, region_col, cust_name_col=None):
    from models.schema import Customer, Executive
    import pandas as pd

    processed_df = df.copy()

    if "Branch" not in processed_df.columns:
        processed_df["Branch"] = ""
    if "Region" not in processed_df.columns:
        processed_df["Region"] = ""
    if "Company Group" not in processed_df.columns:
        processed_df["Company Group"] = ""

    # Load customer code → executive name mapping from DB
    customers = Customer.query.all()
    customer_code_map = {normalize_customer_code(c.code): c.executive.name for c in customers if c.executive}

    # Load executive name → code mapping
    executives = Executive.query.all()
    exec_code_map = {e.name: e.code for e in executives}

    for idx, row in processed_df.iterrows():
        if pd.notna(row[customer_col]):
            customer_code = normalize_customer_code(str(row[customer_col]).strip())

            if customer_code in customer_code_map:
                exec_name = customer_code_map[customer_code]

                # Assign executive code
                if exec_code_col in processed_df.columns and exec_name in exec_code_map:
                    processed_df.at[idx, exec_code_col] = exec_code_map[exec_name]

                # Assign executive name
                if exec_name_col in processed_df.columns:
                    processed_df.at[idx, exec_name_col] = exec_name

                # Assign branch
                branch = get_branches_for_executive(exec_name)
                if branch_col in processed_df.columns:
                    processed_df.at[idx, branch_col] = branch
                processed_df.at[idx, "Branch"] = branch

                # Assign region
                if branch and "," not in branch:
                    region = get_region_for_branch(branch)
                    if region_col in processed_df.columns:
                        processed_df.at[idx, region_col] = region
                    processed_df.at[idx, "Region"] = region

    return processed_df


#####################
# Sales Processing
#####################      

def get_exec_name_by_code(exec_code):
    exec_obj = Executive.query.filter_by(code=exec_code).first()
    return exec_obj.name if exec_obj else None

def get_branches_for_exec(exec_name):
    exec_obj = Executive.query.filter_by(name=exec_name).first()
    if not exec_obj:
        return ""
    branches = BranchExecutiveMap.query.filter_by(executive_id=exec_obj.id).all()
    return ", ".join(sorted(set(b.branch.name for b in branches))) if branches else ""

def get_region_for_branch(branch_name):
    branch = Branch.query.filter_by(name=branch_name).first()
    if not branch:
        return ""
    region_map = RegionBranchMap.query.filter_by(branch_id=branch.id).first()
    return region_map.region.name if region_map else ""

def get_company_for_product(product_name):
    normalized = ' '.join(product_name.lower().strip().split())
    for mapping in CompanyProductMap.query.all():
        mapped_product = mapping.product.name.lower().strip()
        mapped_norm = ' '.join(mapped_product.split())
        if normalized == mapped_norm:
            return mapping.company.name
    for mapping in CompanyProductMap.query.all():
        mapped_product = mapping.product.name.lower().strip()
        mapped_norm = ' '.join(mapped_product.split())
        if normalized in mapped_norm or mapped_norm in normalized:
            return mapping.company.name
    return ""

def process_sales_file(df, exec_code_col, product_col=None, exec_name_col=None, unit_col=None, quantity_col=None, value_col=None):
    df = df.copy()

    # Add standard columns
    df["Branch"] = ""
    df["Region"] = ""
    df["Company Group"] = ""
    if unit_col and quantity_col:
        df["Actual Quantity"] = ""
    if value_col:
        df["Value"] = ""

    exec_found, branch_found, region_found, product_mapped = 0, 0, 0, 0

    for idx, row in df.iterrows():
        exec_name = None

        # Try by code
        exec_code = str(row[exec_code_col]).strip() if pd.notna(row[exec_code_col]) else None
        if exec_code:
            exec_name = get_exec_name_by_code(exec_code)
            if exec_name:
                exec_found += 1

        # Fallback by name
        if not exec_name and exec_name_col and pd.notna(row[exec_name_col]):
            potential_name = str(row[exec_name_col]).strip()
            if Executive.query.filter_by(name=potential_name).first():
                exec_name = potential_name
                exec_found += 1

        # Update executive name
        if exec_name and exec_name_col and exec_name_col in df.columns:
            df.at[idx, exec_name_col] = exec_name

        # Get branch
        if exec_name:
            branch = get_branches_for_exec(exec_name)
            if branch:
                df.at[idx, "Branch"] = branch
                branch_found += 1

                if "," not in branch:
                    region = get_region_for_branch(branch)
                    if region:
                        df.at[idx, "Region"] = region
                        region_found += 1

        # Product → Company
        if product_col and pd.notna(row[product_col]):
            product_name = str(row[product_col]).strip()
            company = get_company_for_product(product_name)
            if company:
                df.at[idx, "Company Group"] = company
                product_mapped += 1

        # Quantity normalization
        if unit_col and quantity_col and pd.notna(row[unit_col]) and pd.notna(row[quantity_col]):
            try:
                unit = str(row[unit_col]).strip().upper()
                quantity = float(row[quantity_col])
                if unit == "MT":
                    actual = quantity
                elif unit in ["KGS", "NOS"]:
                    actual = quantity / 1000
                else:
                    actual = quantity
                df.at[idx, "Actual Quantity"] = actual
            except:
                df.at[idx, "Actual Quantity"] = ""

        # Value in lakhs
        if value_col and pd.notna(row[value_col]):
            try:
                value = float(row[value_col])
                df.at[idx, "Value"] = round(value / 100000, 2)
            except:
                df.at[idx, "Value"] = 0

    print(f"[Sales File] Total: {len(df)}, Execs: {exec_found}, Branches: {branch_found}, Regions: {region_found}, Products: {product_mapped}")
    return df

#################
# OS Processing
#################

def get_branch_for_exec(exec_name):
    exec_obj = Executive.query.filter_by(name=exec_name).first()
    if not exec_obj:
        return ""
    branches = BranchExecutiveMap.query.filter_by(executive_id=exec_obj.id).all()
    return ", ".join(sorted(set(b.branch.name for b in branches))) if branches else ""

def get_region_for_branch(branch_name):
    branch = Branch.query.filter_by(name=branch_name).first()
    if not branch:
        return ""
    region_map = RegionBranchMap.query.filter_by(branch_id=branch.id).first()
    return region_map.region.name if region_map else ""

def match_exec_name(exec_code_from_file):
    exec_code_from_file = str(exec_code_from_file).strip().lower()
    for exec in Executive.query.all():
        db_code = exec.code.strip().lower() if exec.code else ""
        if exec_code_from_file == db_code:
            return exec.name
    for exec in Executive.query.all():
        db_code = exec.code.strip().lower() if exec.code else ""
        if exec_code_from_file in db_code or db_code in exec_code_from_file:
            return exec.name
    return None

def process_os_file(df, exec_code_col):
    df = df.copy()
    if "Branch" not in df.columns:
        df["Branch"] = ""
    if "Region" not in df.columns:
        df["Region"] = ""

    total_rows = len(df)
    exec_found_count = 0
    branch_mapped_count = 0
    region_mapped_count = 0
    exec_not_found = []
    branch_not_found = []

    for idx, row in df.iterrows():
        exec_name = None
        if pd.notna(row[exec_code_col]):
            exec_code = str(row[exec_code_col]).strip()
            exec_name = match_exec_name(exec_code)

            if exec_name:
                exec_found_count += 1
            else:
                exec_not_found.append(exec_code)

        if exec_name:
            branch = get_branch_for_exec(exec_name)
            if branch:
                df.at[idx, "Branch"] = branch
                branch_mapped_count += 1

                if "," not in branch:
                    region = get_region_for_branch(branch)
                    if region:
                        df.at[idx, "Region"] = region
                        region_mapped_count += 1
                    else:
                        branch_not_found.append(branch)
                else:
                    df.at[idx, "Region"] = "Multiple Branches"
            else:
                branch_not_found.append(f"Executive: {exec_name}")

    # Optional: add debug logging
    print(f"[OS FILE] Total Rows: {total_rows}")
    print(f"Executives Found: {exec_found_count}")
    print(f"Branches Mapped: {branch_mapped_count}")
    print(f"Regions Mapped: {region_mapped_count}")
    print(f"Execs not found: {exec_not_found[:10]}")
    print(f"Branch issues: {branch_not_found[:10]}")

    return df
