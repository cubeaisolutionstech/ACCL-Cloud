import pandas as pd
import numpy as np
from io import BytesIO
import logging
from datetime import datetime
from dateutil.relativedelta import relativedelta

logger = logging.getLogger(__name__)

nbc_branch_mapping = {
    'PONDY': 'PONDY', 'PDY': 'PONDY',
    'COVAI': 'ERODE_CBE_KRR',
    'ERODE': 'ERODE_CBE_KRR',
    'KARUR': 'ERODE_CBE_KRR',
    'SLM002': 'SALEM', 'SLMTD1': 'SALEM',
    'BHV1': 'BHAVANI',
    'MDU': 'MADURAI',
    'CHENNAI': 'CHENNAI', 'CHN': 'CHENNAI',
    'TIRUPUR': 'TIRUPUR', 'TPR': 'TIRUPUR',
    'POULTRY': 'POULTRY'
}

def determine_financial_year(date):
    """Determine the financial year for a given date (April to March)."""
    year = date.year
    month = date.month
    if month >= 4:
        return f"{year % 100}-{year % 100 + 1}"
    else:
        return f"{year % 100 - 1}-{year % 100}"

def find_column_by_names(columns, target_names):
    """Find a column by checking multiple possible names (case-insensitive)."""
    columns_upper = [col.upper() for col in columns]
    target_names_upper = [name.upper() for name in target_names]
    
    for target in target_names_upper:
        if target in columns_upper:
            return columns[columns_upper.index(target)]
    return None

def auto_map_nbc_columns(sales_columns):
    """Auto-map Number of Billed Customers columns."""
    return {
        'date': find_column_by_names(sales_columns, ['Date', 'Bill Date', 'Invoice Date']),
        'customer_id': find_column_by_names(sales_columns, ['Customer Code', 'SL Code', 'Customer ID']),
        'branch': find_column_by_names(sales_columns, ['Branch', 'Area', 'Location']),
        'executive': find_column_by_names(sales_columns, ['Executive', 'Sales Executive', 'Executive Name'])
    }

def get_most_recent_month(sorted_months):
    """Extract the most recent month from sorted months list"""
    if not sorted_months or len(sorted_months) == 0:
        return None
    recent_month = sorted_months[-1] if isinstance(sorted_months, list) else str(sorted_months)
    return recent_month.upper()

def create_customer_table(sales_df, date_col, branch_col, customer_id_col, executive_col, selected_branches=None, selected_executives=None):
    """Create customer table WITHOUT S.No column and with proper Branch column"""
    sales_df = sales_df.copy()

    # Validate essential columns
    for col in [date_col, branch_col, customer_id_col, executive_col]:
        if col not in sales_df.columns:
            logger.warning(f"Missing column in sales data: {col}")
            return None

    # Convert to datetime
    sales_df[date_col] = pd.to_datetime(sales_df[date_col], errors='coerce')
    sales_df = sales_df[sales_df[date_col].notna()]

    if sales_df.empty:
        logger.warning("No valid date rows in sales_df")
        return None

    sales_df['Financial_Year'] = sales_df[date_col].apply(determine_financial_year)
    sales_df['Month_Name'] = sales_df[date_col].dt.strftime('%b %Y').str.upper()

    # FIXED: Proper branch processing
    sales_df['Raw_Branch'] = sales_df[branch_col].astype(str).str.strip().str.upper()
    
    # Apply branch mapping - handle cases where mapping doesn't exist
    def map_branch(raw_branch):
        if pd.isna(raw_branch) or raw_branch == '' or raw_branch == 'NAN':
            return 'UNKNOWN'
        # Try direct mapping first
        if raw_branch in nbc_branch_mapping:
            return nbc_branch_mapping[raw_branch]
        # Try extracting last part after ' - '
        if ' - ' in raw_branch:
            branch_part = raw_branch.split(' - ')[-1].strip()
            if branch_part in nbc_branch_mapping:
                return nbc_branch_mapping[branch_part]
        # Return cleaned branch name if no mapping found
        return raw_branch

    sales_df['Branch'] = sales_df['Raw_Branch'].apply(map_branch)
    
    # Debug logging
    logger.info(f"Unique raw branches: {sales_df['Raw_Branch'].unique()[:10]}")
    logger.info(f"Unique mapped branches: {sales_df['Branch'].unique()}")

    # Filters
    if selected_branches:
        sales_df = sales_df[sales_df['Branch'].isin(selected_branches)]
    if selected_executives:
        sales_df = sales_df[sales_df[executive_col].isin(selected_executives)]

    if sales_df.empty:
        logger.warning("No data after applying filters")
        return None

    result_dict = {}
    for fy in sorted(sales_df['Financial_Year'].dropna().unique()):
        fy_df = sales_df[sales_df['Financial_Year'] == fy].copy()
        if fy_df.empty:
            continue

        available_months = fy_df.sort_values(date_col)['Month_Name'].unique().tolist()
        
        # Group by Branch and Month
        grouped = fy_df.groupby(['Branch', 'Month_Name'])[customer_id_col].nunique().reset_index(name='Count')
        
        # Create pivot table
        pivot_df = grouped.pivot_table(index='Branch', columns='Month_Name', values='Count', aggfunc='sum', fill_value=0).reset_index()

        # Ensure we have data
        if pivot_df.empty:
            logger.warning(f"No pivot data for FY {fy}")
            continue

        result_df = pivot_df.copy()

        # Add Grand Total
        grand_total = {'Branch': 'GRAND TOTAL'}
        for month in available_months:
            if month in result_df.columns:
                grand_total[month] = result_df[month].sum()
            else:
                grand_total[month] = 0
        
        result_df = pd.concat([result_df, pd.DataFrame([grand_total])], ignore_index=True)

        # Final columns
        final_cols = ['Branch'] + available_months
        result_df = result_df[final_cols]
        result_df.attrs["columns"] = final_cols

        most_recent_month = get_most_recent_month(available_months)
        result_dict[fy] = {
            'dataframe': result_df,
            'months': available_months,
            'recent_month': most_recent_month
        }

    return result_dict
# -------------------- OD Target --------------------

def auto_map_od_target_columns(os_columns):
    """Auto-map OD Target columns."""
    return {
        'area': find_column_by_names(os_columns, ['Branch', 'Area', 'Location', 'Unit']),
        'due_date': find_column_by_names(os_columns, ['Due Date', 'Due_Date', 'DueDate']),
        'net_value': find_column_by_names(os_columns, ['Net Value', 'NetValue', 'Amount', 'Value']),
        'executive': find_column_by_names(os_columns, ['Executive Name', 'Executive', 'Sales Executive'])
    }

def extract_area_name(area):
    """Extract and standardize branch names."""
    if pd.isna(area) or str(area).strip() == '':
        return None  # Return None for empty/null values  
    area = str(area).strip().upper()
    if area == 'HO' or area.endswith('-HO'):
        return None
    branch_variations = {
        'PUDUCHERRY': ['PUDUCHERRY', 'PONDY', 'PUDUCHERRY - PONDY', 'PONDICHERRY', 'PUDUCHERI', 'aaaa - PUDUCHERRY', 'AAAA - PUDUCHERRY'],
        'COIMBATORE': ['COIMBATORE', 'CBE', 'COIMBATORE - CBE', 'COIMBATURE', 'aaaa - COIMBATORE', 'AAAA - COIMBATORE'],
        'KARUR': ['KARUR', 'KRR', 'KARUR - KRR', 'aaaa - KARUR', 'AAAA - KARUR'],
        'MADURAI': ['MADURAI', 'MDU', 'MADURAI - MDU', 'MADURA', 'aaaa - MADURAI', 'AAAA - MADURAI'],
        'CHENNAI': ['CHENNAI', 'CHN', 'aaaa - CHENNAI', 'AAAA - CHENNAI'],
    }
    
    for standard_name, variations in branch_variations.items():
        for variation in variations:
            if variation in area:
                return standard_name
    prefixes = ['AAAA - ', 'aaaa - ', 'BBB - ', 'bbb - ', 'ASIA CRYSTAL COMMODITY LLP - ']
    for prefix in prefixes:
        if area.startswith(prefix.upper()):
            return area[len(prefix):].strip()
    separators = [' - ', '-', ':']
    for sep in separators:
        if sep in area:
            return area.split(sep)[-1].strip()    
    return area

def extract_executive_name(executive):
    """Normalize executive names, treating null/empty as 'BLANK'."""
    if pd.isna(executive) or str(executive).strip() == '':
        return 'BLANK'  # Treat null/empty as 'BLANK'
    return str(executive).strip().upper()  # Standardize to uppercase

def filter_os_qty(
    os_df,
    os_area_col,
    os_qty_col,
    os_due_date_col,
    os_exec_col,
    selected_branches=None,
    selected_years=None,
    till_month=None,
    selected_executives=None
):
    """Backend-compatible version of OD Target filter and aggregation - FIXED column naming."""
    from utils.nbc_od_utils import extract_area_name, extract_executive_name  # If not in same file

    required_columns = [os_area_col, os_qty_col, os_due_date_col, os_exec_col]
    for col in required_columns:
        if col not in os_df.columns:
            logger.warning(f"Column '{col}' not found in OS data.")
            return None, None, None

    os_df = os_df.copy()
    os_df[os_area_col] = os_df[os_area_col].apply(extract_area_name)
    os_df[os_exec_col] = os_df[os_exec_col].apply(extract_executive_name)

    try:
        os_df[os_due_date_col] = pd.to_datetime(os_df[os_due_date_col], errors='coerce')
    except Exception as e:
        logger.warning(f"Error converting due date column to datetime: {e}")
        return None, None, None

    start_date, end_date = None, None
    if selected_years and till_month:
        month_map = {
            'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6,
            'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12
        }
        till_month_num = month_map.get(till_month)
        if not till_month_num:
            logger.warning("Invalid month selected for OD Target.")
            return None, None, None

        selected_years = [int(year) for year in selected_years]
        earliest_year = min(selected_years)
        latest_year = max(selected_years)
        start_date = datetime(earliest_year, 1, 1)
        end_date = datetime(latest_year, till_month_num, 1) + relativedelta(months=1) - relativedelta(days=1)

        os_df = os_df[
            (os_df[os_due_date_col].notna()) &
            (os_df[os_due_date_col] >= start_date) &
            (os_df[os_due_date_col] <= end_date)
        ]

        if os_df.empty:
            logger.warning(f"No data between Jan {earliest_year} and {end_date.strftime('%b %Y')}.")
            return None, None, None

    all_branches = sorted(os_df[os_area_col].dropna().unique())
    if not selected_branches:
        selected_branches = all_branches

    if sorted(selected_branches) != all_branches:
        os_df = os_df[os_df[os_area_col].isin(selected_branches)]
        if os_df.empty:
            logger.warning("No data matches selected branches.")
            return None, None, None

    all_executives = sorted(os_df[os_exec_col].dropna().unique())
    if selected_executives and sorted(selected_executives) != all_executives:
        os_df = os_df[os_df[os_exec_col].isin(selected_executives)]
        if os_df.empty:
            logger.warning("No data matches selected executives.")
            return None, None, None

    os_df[os_qty_col] = pd.to_numeric(os_df[os_qty_col], errors='coerce').fillna(0)
    os_df_positive = os_df[os_df[os_qty_col] > 0].copy()

    if os_df_positive.empty:
        logger.warning("No positive net values found in filtered OS data.")
        return None, None, None

    # FIXED: Use 'TARGET' instead of 'TARGET/L' for consistency with Streamlit
    os_grouped_qty = (
        os_df_positive.groupby(os_area_col)[os_qty_col]
        .sum()
        .reset_index()
        .rename(columns={os_area_col: 'Area', os_qty_col: 'TARGET'})  # Changed to 'TARGET'
    )

    os_grouped_qty['TARGET'] = os_grouped_qty['TARGET'] / 100000  # Convert to lakhs

    branches_to_display = selected_branches or all_branches
    result_df = pd.DataFrame({'Area': branches_to_display})
    result_df = pd.merge(result_df, os_grouped_qty, on='Area', how='left').fillna({'TARGET': 0})

    total_row = pd.DataFrame([{'Area': 'TOTAL', 'TARGET': result_df['TARGET'].sum()}])
    result_df = pd.concat([result_df, total_row], ignore_index=True)

    result_df['TARGET'] = result_df['TARGET'].round(2)
    return result_df, start_date, end_date