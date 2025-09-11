import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta
import logging
import os
from pathlib import Path

logger = logging.getLogger(__name__)

def determine_financial_year(date):
    """Determine financial year from date"""
    if pd.isna(date):
        return None
    year = date.year
    month = date.month
    if month >= 4:
        return f"{year}-{year+1}"
    else:
        return f"{year-1}-{year}"

def extract_area_name(area):
    """Extract and standardize area/branch names"""
    if pd.isna(area) or not str(area).strip():
        return None
    
    area = str(area).strip()
    area_upper = area.upper()
    
    # Skip HO entries
    if area_upper == 'HO' or area_upper.endswith('-HO'):
        return None
    
    # Branch variations mapping
    branch_variations = {
        'PUDUCHERRY': ['PUDUCHERRY', 'PONDY', 'PUDUCHERRY - PONDY', 'PONDICHERRY', 'PUDUCHERI', 'aaaa - PUDUCHERRY', 'AAAA - PUDUCHERRY'],
        'COIMBATORE': ['COIMBATORE', 'CBE', 'COIMBATORE - CBE', 'COIMBATURE', 'aaaa - COIMBATORE', 'AAAA - COIMBATORE'],
        'KARUR': ['KARUR', 'KRR', 'KARUR - KRR', 'aaaa - KARUR', 'AAAA - KARUR'],
        'MADURAI': ['MADURAI', 'MDU', 'MADURAI - MDU', 'MADURA', 'aaaa - MADURAI', 'AAAA - MADURAI'],
        'CHENNAI': ['CHENNAI', 'CHN', 'aaaa - CHENNAI', 'AAAA - CHENNAI'],
    }
    
    # Check for branch variations
    for standard_name, variations in branch_variations.items():
        for variation in variations:
            if variation in area_upper:
                return standard_name
    
    # Remove common prefixes
    prefixes = ['AAAA - ', 'aaaa - ', 'BBB - ', 'bbb - ', 'ASIA CRYSTAL COMMODITY LLP - ']
    for prefix in prefixes:
        if area_upper.startswith(prefix.upper()):
            return area[len(prefix):].strip().upper()
    
    # Handle separators
    separators = [' - ', '-', ':']
    for sep in separators:
        if sep in area_upper:
            return area_upper.split(sep)[-1].strip()
    
    return area_upper

# ==================== BILLED CUSTOMERS FUNCTIONS ====================

def auto_map_customer_columns(sales_df):
    """Auto-map columns for customer analysis"""
    try:
        def find_column(columns, target_names):
            for target in target_names:
                for col in columns:
                    if col.lower() == target.lower():
                        return col
            return None
        
        column_mappings = {
            'date': ['Date', 'Bill Date', 'Invoice Date'],
            'branch': ['Branch', 'Area', 'Location'],
            'customer_id': ['Customer Code', 'SL Code', 'Customer ID'],
            'executive': ['Executive Name', 'Executive', 'Sales Executive']
        }
        
        mapping = {}
        for key, targets in column_mappings.items():
            mapping[key] = find_column(sales_df.columns.tolist(), targets)
        
        return {
            'success': True,
            'mapping': mapping,
            'columns': list(sales_df.columns)
        }
        
    except Exception as e:
        logger.error(f"Error auto-mapping customer columns: {e}")
        return {
            'success': False,
            'mapping': {},
            'columns': []
        }

def get_customer_options(sales_df, date_col, branch_col, executive_col):
    """Get available options for customer analysis filters"""
    try:
        sales_df = sales_df.copy()
        
        # Convert date column
        sales_df[date_col] = pd.to_datetime(sales_df[date_col], errors='coerce', dayfirst=True)
        
        # Get available months
        available_months = sorted(sales_df[date_col].dt.strftime('%b %Y').dropna().unique().tolist())
        
        # Get branches (extract raw branches directly from branch column)
        raw_branches = sales_df[branch_col].dropna().astype(str).str.strip().str.upper().unique().tolist()
        all_branches = sorted(raw_branches)
        
        # Get executives
        all_executives = sorted(sales_df[executive_col].dropna().astype(str).str.strip().str.upper().unique().tolist())
        
        return {
            'success': True,
            'available_months': available_months,
            'branches': all_branches,
            'executives': all_executives
        }
        
    except Exception as e:
        logger.error(f"Error getting customer options: {e}")
        return {
            'success': False,
            'error': str(e),
            'available_months': [],
            'branches': [],
            'executives': []
        }

def create_customer_table(sales_df, date_col, branch_col, customer_id_col, executive_col, 
                         selected_months=None, selected_branches=None, selected_executives=None):
    """
    Creates a table of unique customer counts per executive for selected months, branches, and executives.
    """
    try:
        sales_df = sales_df.copy()
        
        # Validate columns
        required_columns = [date_col, branch_col, customer_id_col, executive_col]
        for col in required_columns:
            if col not in sales_df.columns:
                return {
                    'success': False,
                    'error': f"Column '{col}' not found in sales data."
                }
        
        # Convert date column to datetime
        try:
            sales_df[date_col] = pd.to_datetime(sales_df[date_col], errors='coerce', dayfirst=True)
        except Exception as e:
            return {
                'success': False,
                'error': f"Error converting '{date_col}' to datetime: {e}. Ensure dates are in a valid format."
            }
        
        # Check for valid dates
        valid_dates = sales_df[date_col].notna()
        if not valid_dates.any():
            return {
                'success': False,
                'error': f"Column '{date_col}' contains no valid dates."
            }
        
        # Extract month-year for filtering
        sales_df['Month_Year'] = sales_df[date_col].dt.strftime('%b %Y')
        
        # Filter by selected months if provided
        if selected_months:
            sales_df = sales_df[sales_df['Month_Year'].isin(selected_months)]
            if sales_df.empty:
                return {
                    'success': False,
                    'error': f"No data found for selected months: {', '.join(selected_months)}"
                }
        
        # Determine financial year
        sales_df['Financial_Year'] = sales_df[date_col].apply(determine_financial_year)
        available_financial_years = sales_df['Financial_Year'].dropna().unique()
        
        if len(available_financial_years) == 0:
            return {
                'success': False,
                'error': "No valid financial years found in the data."
            }
        
        result_dict = {}
        
        for fin_year in sorted(available_financial_years):
            fy_df = sales_df[sales_df['Financial_Year'] == fin_year].copy()
            if fy_df.empty:
                continue
            
            # Extract unique months in chronological order
            fy_df['Month_Year_Period'] = fy_df[date_col].dt.to_period('M')
            available_months_periods = fy_df['Month_Year_Period'].unique()
            month_names = [pd.to_datetime(str(m) + '-01').strftime('%b %Y') for m in sorted(available_months_periods)]
            
            if not month_names:
                continue
            
            # Filter by selected months if provided
            if selected_months:
                month_names = [m for m in month_names if m in selected_months]
                if not month_names:
                    continue
            
            # Standardize branch and executive names
            try:
                fy_df['Branch'] = fy_df[branch_col].astype(str).str.strip().str.upper()
                fy_df['Executive_Upper'] = fy_df[executive_col].astype(str).str.strip().str.upper()
            except Exception as e:
                logger.error(f"Error processing columns: {e}")
                continue
            
            # Apply branch filter if provided
            if selected_branches:
                fy_df = fy_df[fy_df['Branch'].isin([b.upper() for b in selected_branches])]
                if fy_df.empty:
                    continue
            
            # Determine executives to display based on both branch and executive selections
            if selected_branches:
                # If branches are selected, get executives associated with those branches
                branch_df = fy_df[fy_df['Branch'].isin([b.upper() for b in selected_branches])]
                branch_executives = sorted(branch_df['Executive_Upper'].dropna().unique())
                
                # If specific executives are also selected, use intersection
                if selected_executives:
                    selected_execs_upper = [str(e).upper() for e in selected_executives]
                    executives_to_display = [exec for exec in branch_executives if exec in selected_execs_upper]
                else:
                    executives_to_display = branch_executives
            else:
                # Use provided selected_executives or all executives in filtered data
                executives_to_display = [str(e).upper() for e in selected_executives] if selected_executives else sorted(fy_df['Executive_Upper'].dropna().unique())
            
            # Apply executive filter
            if executives_to_display:
                fy_df = fy_df[fy_df['Executive_Upper'].isin(executives_to_display)]
                if fy_df.empty:
                    continue
            
            if not executives_to_display:
                continue
            
            # Group by executive and month to count unique customer codes
            grouped_df = fy_df.groupby(['Executive_Upper', 'Month_Year'])[customer_id_col].nunique().reset_index(name='Customer_Count')
            
            # Pivot to create table with months as columns
            pivot_df = grouped_df.pivot_table(
                values='Customer_Count',
                index='Executive_Upper',
                columns='Month_Year',
                aggfunc='sum',
                fill_value=0
            ).reset_index()
            
            # Rename index column
            pivot_df = pivot_df.rename(columns={'Executive_Upper': 'Executive Name'})
            
            # Create result dataframe with all executives to display
            result_df = pd.DataFrame({'Executive Name': executives_to_display})
            result_df = pd.merge(
                result_df,
                pivot_df,
                on='Executive Name',
                how='left'
            ).fillna(0)
            
            # Keep only selected months
            columns_to_keep = ['Executive Name'] + month_names
            result_df = result_df[[col for col in columns_to_keep if col in result_df.columns]]
            
            # Convert counts to integers
            for col in result_df.columns[1:]:
                result_df[col] = result_df[col].astype(int)
            
            # Add S.No column
            result_df.insert(0, 'S.No', [str(i) for i in range(1, len(result_df) + 1)])
            
            # Add total row
            total_row = {'S.No': '0', 'Executive Name': 'GRAND TOTAL'}
            for col in month_names:
                if col in result_df.columns:
                    total_row[col] = result_df[col].sum()
            
            result_df = pd.concat([result_df, pd.DataFrame([total_row])], ignore_index=True)
            
            result_dict[fin_year] = {
                'data': result_df.to_dict('records'),
                'columns': list(result_df.columns),
                'sorted_months': month_names
            }
        
        return {
            'success': True,
            'results': result_dict
        }
        
    except Exception as e:
        logger.error(f"Error creating customer table: {e}")
        return {
            'success': False,
            'error': f"Error creating customer table: {str(e)}"
        }

# ==================== OD TARGET FUNCTIONS ====================

def auto_map_od_columns(os_df):
    """Auto-map columns for OD target analysis"""
    try:
        def find_column(columns, target_names):
            for target in target_names:
                for col in columns:
                    if col.lower() == target.lower():
                        return col
            return None
        
        column_mappings = {
            'area': ['Branch', 'Area', 'Location'],
            'net_value': ['Net Value', 'Value', 'Amount'],
            'due_date': ['Due Date', 'Date'],
            'executive': ['Executive Name', 'Executive', 'Sales Executive']
        }
        
        mapping = {}
        for key, targets in column_mappings.items():
            mapping[key] = find_column(os_df.columns.tolist(), targets)
        
        return {
            'success': True,
            'mapping': mapping,
            'columns': list(os_df.columns)
        }
        
    except Exception as e:
        logger.error(f"Error auto-mapping OD columns: {e}")
        return {
            'success': False,
            'mapping': {},
            'columns': []
        }

def get_od_options(os_df, due_date_col, area_col, executive_col):
    """Get available options for OD target analysis"""
    try:
        os_df = os_df.copy()
        
        # Convert date column
        os_df[due_date_col] = pd.to_datetime(os_df[due_date_col], errors='coerce')
        
        # Get available years
        years = sorted(os_df[due_date_col].dt.year.dropna().astype(int).unique().tolist())
        
        # Get branches using extract_area_name
        os_branches = sorted(set([b for b in os_df[area_col].apply(extract_area_name).dropna().unique() if b]))
        
        # Get executives
        os_executives = sorted(os_df[executive_col].dropna().astype(str).unique().tolist())
        
        return {
            'success': True,
            'years': [str(year) for year in years],
            'branches': os_branches,
            'executives': os_executives
        }
        
    except Exception as e:
        logger.error(f"Error getting OD options: {e}")
        return {
            'success': False,
            'error': str(e),
            'years': [],
            'branches': [],
            'executives': []
        }

def filter_os_qty(os_df, os_area_col, os_qty_col, os_due_date_col, os_exec_col, 
                  selected_branches=None, selected_years=None, till_month=None, selected_executives=None):
    """Filter OS quantity data based on selections"""
    try:
        required_columns = [os_area_col, os_qty_col, os_due_date_col, os_exec_col]
        for col in required_columns:
            if col not in os_df.columns:
                return {
                    'success': False,
                    'error': f"Column '{col}' not found in OS data."
                }
        
        os_df = os_df.copy()
        os_df[os_area_col] = os_df[os_area_col].apply(extract_area_name).astype(str).str.strip().str.upper()
        os_df[os_exec_col] = os_df[os_exec_col].apply(lambda x: 'BLANK' if pd.isna(x) or str(x).strip() == '' else str(x).strip().upper())
        
        try:
            os_df[os_due_date_col] = pd.to_datetime(os_df[os_due_date_col], errors='coerce')
        except Exception as e:
            return {
                'success': False,
                'error': f"Error converting '{os_due_date_col}' to datetime: {e}. Ensure dates are in 'YYYY-MM-DD' format."
            }
        
        # Convert negative values to 0 BEFORE division
        os_df[os_qty_col] = pd.to_numeric(os_df[os_qty_col], errors='coerce').fillna(0)
        os_df[os_qty_col] = os_df[os_qty_col].clip(lower=0)  # Convert negative values to 0
        os_df[os_qty_col] = os_df[os_qty_col] / 100000  # Then divide by 100000
        
        start_date, end_date = None, None
        if selected_years and till_month:
            month_map = {
                'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6,
                'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12
            }
            if till_month == "November Brodha":
                till_month = "November"
            till_month_num = month_map.get(till_month)
            if not till_month_num:
                return {
                    'success': False,
                    'error': f"Invalid month selected: {till_month}"
                }
            selected_years = [int(year) for year in selected_years]
            earliest_year = min(selected_years)
            latest_year = max(selected_years)
            start_date = datetime(earliest_year, 1, 1)
            end_date = (datetime(latest_year, till_month_num, 1) + relativedelta(months=1) - relativedelta(days=1))
            os_df = os_df[
                (os_df[os_due_date_col].notna()) &
                (os_df[os_due_date_col] >= start_date) &
                (os_df[os_due_date_col] <= end_date)
            ]
            if os_df.empty:
                return {
                    'success': False,
                    'error': f"No data matches the period from Jan {earliest_year} to {end_date.strftime('%b %Y')}."
                }
        
        # Apply branch filter if provided
        if selected_branches:
            os_df = os_df[os_df[os_area_col].isin([b.upper() for b in selected_branches])]
            if os_df.empty:
                return {
                    'success': False,
                    'error': "No data matches the selected branches."
                }
        
        # Determine executives to display based on both branch and executive selections
        if selected_branches:
            # If branches are selected, get executives associated with those branches
            branch_df = os_df[os_df[os_area_col].isin([b.upper() for b in selected_branches])]
            branch_executives = sorted(branch_df[os_exec_col].dropna().unique())
            
            # If specific executives are also selected, use intersection
            if selected_executives:
                selected_execs_upper = [str(e).upper() for e in selected_executives]
                executives_to_display = [exec for exec in branch_executives if exec in selected_execs_upper]
            else:
                executives_to_display = branch_executives
        else:
            # Use provided selected_executives or all executives in filtered data
            executives_to_display = [str(e).upper() for e in selected_executives] if selected_executives else sorted(os_df[os_exec_col].dropna().unique())
        
        # Filter data by selected executives
        if executives_to_display:
            os_df = os_df[os_df[os_exec_col].isin(executives_to_display)]
            if os_df.empty:
                return {
                    'success': False,
                    'error': "No data matches the selected executives."
                }
        
        # Filter to only positive values (after negative conversion to 0)
        os_df = os_df[os_df[os_qty_col] > 0]
        if os_df.empty:
            return {
                'success': False,
                'error': "No positive net values found in the filtered data."
            }
        
        # Group and aggregate data
        os_grouped_qty = (os_df.groupby(os_exec_col)
                         .agg({os_qty_col: 'sum'})
                         .reset_index()
                         .rename(columns={os_exec_col: 'Executive', os_qty_col: 'TARGET'}))

        # Ensure all executives_to_display are included
        result_df = pd.DataFrame({'Executive': executives_to_display})
        result_df = pd.merge(result_df, os_grouped_qty, on='Executive', how='left').fillna({'TARGET': 0})
        
        # Add total row
        total_row = pd.DataFrame([{'Executive': 'TOTAL', 'TARGET': result_df['TARGET'].sum()}])
        result_df = pd.concat([result_df, total_row], ignore_index=True)
        result_df['TARGET'] = result_df['TARGET'].round(2)
        
        return {
            'success': True,
            'data': result_df.to_dict('records'),
            'columns': list(result_df.columns),
            'start_date': start_date.strftime('%b %Y') if start_date else None,
            'end_date': end_date.strftime('%b %Y') if end_date else None
        }
        
    except Exception as e:
        logger.error(f"Error filtering OS quantity: {e}")
        return {
            'success': False,
            'error': f"Error filtering OS quantity: {str(e)}"
        }

def load_data(file_path):
    """Load data from file"""
    try:
        if not os.path.isabs(file_path):
            file_path = os.path.abspath(file_path)
        
        if file_path.endswith('.csv'):
            return pd.read_csv(file_path)
        elif file_path.endswith(('.xlsx', '.xls')):
            return pd.read_excel(file_path)
        else:
            raise ValueError(f"Unsupported file format: {file_path}")
    except Exception as e:
        logger.error(f"Error loading file {file_path}: {str(e)}")
        raise