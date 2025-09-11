import pandas as pd
import numpy as np
import os
from pathlib import Path
import logging
from datetime import datetime
from dateutil.relativedelta import relativedelta

logger = logging.getLogger(__name__)

def extract_area_name(area):
    """Extract and standardize area/branch name"""
    if pd.isna(area) or not str(area).strip():
        return None
    area = str(area).strip()
    area_upper = area.upper()
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
    
    for standard_name, variations in branch_variations.items():
        for variation in variations:
            if variation in area_upper:
                return standard_name
    
    # Remove common prefixes
    prefixes = ['AAAA - ', 'aaaa - ', 'BBB - ', 'bbb - ', 'ASIA CRYSTAL COMMODITY LLP - ']
    for prefix in prefixes:
        if area_upper.startswith(prefix.upper()):
            return area[len(prefix):].strip().upper()
    
    # Split by common separators
    separators = [' - ', '-', ':']
    for sep in separators:
        if sep in area_upper:
            return area_upper.split(sep)[-1].strip()
    
    return area_upper

def get_available_months_od(os_jan, os_feb, total_sale,
                           os_jan_due_date_col, os_jan_ref_date_col,
                           os_feb_due_date_col, os_feb_ref_date_col,
                           sale_bill_date_col, sale_due_date_col):
    """Get available months from all OD-related files"""
    months = set()
    
    for df, date_cols in [
        (os_jan, [os_jan_due_date_col, os_jan_ref_date_col]),
        (os_feb, [os_feb_due_date_col, os_feb_ref_date_col]),
        (total_sale, [sale_bill_date_col, sale_due_date_col])
    ]:
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                valid_dates = df[col].dropna()
                month_years = valid_dates.dt.strftime('%b-%y').unique()
                months.update(month_years)
    
    months = sorted(list(months), key=lambda x: pd.to_datetime("01-" + x, format="%d-%b-%y"))
    return months

def get_od_executives_and_branches(os_jan, os_feb, total_sale,
                                  os_jan_exec_col, os_feb_exec_col, sale_exec_col,
                                  os_jan_area_col, os_feb_area_col, sale_area_col):
    """Get available executives and branches from OD data"""
    try:
        # Get executives
        all_executives = set()
        for df, exec_col in [
            (os_jan, os_jan_exec_col),
            (os_feb, os_feb_exec_col),
            (total_sale, sale_exec_col)
        ]:
            if exec_col in df.columns:
                execs = df[exec_col].dropna().astype(str).str.strip().str.upper().unique().tolist()
                all_executives.update(execs)
        
        # Get branches
        all_branches = set()
        for df, area_col in [
            (os_jan, os_jan_area_col),
            (os_feb, os_feb_area_col),
            (total_sale, sale_area_col)
        ]:
            if area_col in df.columns:
                branches = df[area_col].apply(extract_area_name).dropna().astype(str).str.upper().unique().tolist()
                all_branches.update([b for b in branches if b])
        
        return {
            'executives': sorted(list(all_executives)),
            'branches': sorted(list(all_branches))
        }
        
    except Exception as e:
        print(f"Error getting OD executives and branches: {str(e)}")
        return {
            'executives': [],
            'branches': []
        }

def auto_map_od_columns(os_jan_df, os_feb_df, sales_df):
    """Auto-map columns for OD analysis with improved error handling"""
    try:
        def find_column(columns, target_names, default_index=0):
            """Find matching column with fallback options"""
            if not columns:
                return None
                
            # First try exact matches (case-insensitive)
            for target in target_names:
                for col in columns:
                    if str(col).lower().strip() == str(target).lower().strip():
                        return col
            
            # Then try partial matches
            for target in target_names:
                for col in columns:
                    if str(target).lower() in str(col).lower():
                        return col
            
            # Fallback to first available column if exists
            return columns[default_index] if len(columns) > default_index else None

        # Enhanced column mappings for OD analysis with more variations
        os_jan_mappings = {
            'due_date': ['Due Date', 'DueDate', 'Due_Date', 'date_due', 'due', 'maturity_date'],
            'ref_date': ['Ref. Date', 'Reference Date', 'Ref Date', 'RefDate', 'ref_date', 'reference_date'],
            'net_value': ['Net Value', 'NetValue', 'Net_Value', 'net_val', 'value', 'amount', 'net_amount'],
            'executive': ['Executive Name', 'Executive', 'Exec Name', 'exec', 'salesperson', 'sales_exec'],
            'sl_code': ['Party Code', 'SL Code', 'Customer Code', 'SLCode', 'sl_code', 'party_code', 'customer_code', 'cust_code'],
            'area': ['Branch', 'Area', 'Location', 'Region', 'Office', 'Territory']
        }

        os_feb_mappings = {
            'due_date': ['Due Date', 'DueDate', 'Due_Date', 'date_due', 'due', 'maturity_date'],
            'ref_date': ['Ref. Date', 'Reference Date', 'Ref Date', 'RefDate', 'ref_date', 'reference_date'],
            'net_value': ['Net Value', 'NetValue', 'Net_Value', 'net_val', 'value', 'amount', 'net_amount'],
            'executive': ['Executive Name', 'Executive', 'Exec Name', 'exec', 'salesperson', 'sales_exec'],
            'sl_code': ['Party Code', 'SL Code', 'Customer Code', 'SLCode', 'sl_code', 'party_code', 'customer_code', 'cust_code'],
            'area': ['Branch', 'Area', 'Location', 'Region', 'Office', 'Territory']
        }

        sales_mappings = {
            'bill_date': ['Date', 'Bill Date', 'BillDate', 'Bill_Date', 'invoice_date', 'sale_date', 'transaction_date'],
            'due_date': ['Due Date', 'DueDate', 'Due_Date', 'date_due', 'due', 'maturity_date'],
            'value': ['Invoice Value', 'Value', 'InvoiceValue', 'Invoice_Value', 'amount', 'sale_value', 'bill_value'],
            'executive': ['Executive Name', 'Executive', 'Exec Name', 'exec', 'salesperson', 'sales_exec'],
            'sl_code': ['Customer Code', 'SL Code', 'Party Code', 'SLCode', 'sl_code', 'party_code', 'customer_code', 'cust_code'],
            'area': ['Branch', 'Area', 'Location', 'Region', 'Office', 'Territory']
        }

        # Initialize mappings
        os_jan_mapping = {}
        os_feb_mapping = {}
        sales_mapping = {}

        # Get column lists (handle None DataFrames)
        os_jan_columns = list(os_jan_df.columns) if os_jan_df is not None and not os_jan_df.empty else []
        os_feb_columns = list(os_feb_df.columns) if os_feb_df is not None and not os_feb_df.empty else []
        sales_columns = list(sales_df.columns) if sales_df is not None and not sales_df.empty else []

        # Auto-map OS Jan columns
        if os_jan_columns:
            for key, targets in os_jan_mappings.items():
                mapped_col = find_column(os_jan_columns, targets)
                os_jan_mapping[key] = mapped_col
        else:
            # Set all mappings to None if no columns available
            for key in os_jan_mappings.keys():
                os_jan_mapping[key] = None

        # Auto-map OS Feb columns
        if os_feb_columns:
            for key, targets in os_feb_mappings.items():
                mapped_col = find_column(os_feb_columns, targets)
                os_feb_mapping[key] = mapped_col
        else:
            # Set all mappings to None if no columns available
            for key in os_feb_mappings.keys():
                os_feb_mapping[key] = None

        # Auto-map Sales columns
        if sales_columns:
            for key, targets in sales_mappings.items():
                mapped_col = find_column(sales_columns, targets)
                sales_mapping[key] = mapped_col
        else:
            # Set all mappings to None if no columns available
            for key in sales_mappings.keys():
                sales_mapping[key] = None

        # Validation: Check for missing critical mappings
        missing_mappings = []
        
        critical_os_jan_fields = ['due_date', 'net_value', 'executive']
        critical_os_feb_fields = ['due_date', 'net_value', 'executive']
        critical_sales_fields = ['bill_date', 'value', 'executive']
        
        for field in critical_os_jan_fields:
            if not os_jan_mapping.get(field):
                missing_mappings.append(f"OS Jan: {field}")
                
        for field in critical_os_feb_fields:
            if not os_feb_mapping.get(field):
                missing_mappings.append(f"OS Feb: {field}")
                
        for field in critical_sales_fields:
            if not sales_mapping.get(field):
                missing_mappings.append(f"Sales: {field}")

        result = {
            'os_jan_mapping': os_jan_mapping,
            'os_feb_mapping': os_feb_mapping,
            'sales_mapping': sales_mapping,
            'os_jan_columns': os_jan_columns,
            'os_feb_columns': os_feb_columns,
            'sales_columns': sales_columns,
            'success': len(missing_mappings) == 0,
            'missing_mappings': missing_mappings,
            'error': None if len(missing_mappings) == 0 else f"Missing required field mappings: {', '.join(missing_mappings)}"
        }
        
        return result
        
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print(f"Error in auto_map_od_columns: {str(e)}")
        print(f"Full traceback: {error_details}")
        
        return {
            'os_jan_mapping': {},
            'os_feb_mapping': {},
            'sales_mapping': {},
            'os_jan_columns': [],
            'os_feb_columns': [],
            'sales_columns': [],
            'success': False,
            'missing_mappings': [],
            'error': f"Auto-mapping failed: {str(e)}"
        }

def calculate_od_values(os_jan, os_feb, total_sale, selected_month_str,
                        os_jan_due_date_col, os_jan_ref_date_col, os_jan_net_value_col, os_jan_exec_col, os_jan_sl_code_col, os_jan_area_col,
                        os_feb_due_date_col, os_feb_ref_date_col, os_feb_net_value_col, os_feb_exec_col, os_feb_sl_code_col, os_feb_area_col,
                        sale_bill_date_col, sale_due_date_col, sale_value_col, sale_exec_col, sale_sl_code_col, sale_area_col,
                        selected_executives, selected_branches=None):
    """
    Flask version of OD Target vs Collection calculation with improved individual executive report logic
    """
    try:
        # Convert and validate numeric cols
        for df, col, file in [
            (os_jan, os_jan_net_value_col, "OS Jan"),
            (os_feb, os_feb_net_value_col, "OS Feb"),
            (total_sale, sale_value_col, "Total Sale")
        ]:
            try:
                df[col] = pd.to_numeric(df[col], errors='coerce')
                if df[col].isna().all():
                    return {"success": False, "error": f"Column '{col}' in {file} contains no valid numeric data."}
            except Exception as e:
                return {"success": False, "error": f"Error processing column '{col}' in {file}: {e}"}

        # Clip negative
        os_jan[os_jan_net_value_col] = os_jan[os_jan_net_value_col].clip(lower=0)
        os_feb[os_feb_net_value_col] = os_feb[os_feb_net_value_col].clip(lower=0)

        # Standardize branch
        os_jan[os_jan_area_col] = os_jan[os_jan_area_col].apply(extract_area_name).astype(str).str.strip().str.upper()
        os_feb[os_feb_area_col] = os_feb[os_feb_area_col].apply(extract_area_name).astype(str).str.strip().str.upper()
        total_sale[sale_area_col] = total_sale[sale_area_col].apply(extract_area_name).astype(str).str.strip().str.upper()

        # Branch filter
        if selected_branches:
            os_jan = os_jan[os_jan[os_jan_area_col].isin([b.upper() for b in selected_branches])]
            os_feb = os_feb[os_feb[os_feb_area_col].isin([b.upper() for b in selected_branches])]
            total_sale = total_sale[total_sale[sale_area_col].isin([b.upper() for b in selected_branches])]

        # Date conversion + exec normalization (matching original backend naming)
        os_jan[os_jan_due_date_col] = pd.to_datetime(os_jan[os_jan_due_date_col], errors='coerce')
        os_jan[os_jan_ref_date_col] = pd.to_datetime(os_jan.get(os_jan_ref_date_col), errors='coerce')
        os_jan["SL Code"] = os_jan[os_jan_sl_code_col].astype(str)
        os_jan["Executive"] = os_jan[os_jan_exec_col].astype(str).str.strip().str.upper()

        os_feb[os_feb_due_date_col] = pd.to_datetime(os_feb[os_feb_due_date_col], errors='coerce')
        os_feb[os_feb_ref_date_col] = pd.to_datetime(os_feb.get(os_feb_ref_date_col), errors='coerce')
        os_feb["SL Code"] = os_feb[os_feb_sl_code_col].astype(str)
        os_feb["Executive"] = os_feb[os_feb_exec_col].astype(str).str.strip().str.upper()

        total_sale[sale_bill_date_col] = pd.to_datetime(total_sale[sale_bill_date_col], errors='coerce')
        total_sale[sale_due_date_col] = pd.to_datetime(total_sale[sale_due_date_col], errors='coerce')
        total_sale["SL Code"] = total_sale[sale_sl_code_col].astype(str)
        total_sale["Executive"] = total_sale[sale_exec_col].astype(str).str.strip().str.upper()

        # FIXED: Get all executives from all sources, not just selected branches
        if selected_executives:
            selected_execs_upper = [str(e).strip().upper() for e in selected_executives]
            executives_to_display = selected_execs_upper
        else:
            # Get all executives from all datasets
            all_executives = set()
            all_executives.update(os_jan["Executive"].dropna().unique())
            all_executives.update(os_feb["Executive"].dropna().unique()) 
            all_executives.update(total_sale["Executive"].dropna().unique())
            executives_to_display = sorted(list(all_executives))

        # Filter data by selected executives - but don't require all datasets to have data
        os_jan_filtered = os_jan[os_jan["Executive"].isin(executives_to_display)]
        os_feb_filtered = os_feb[os_feb["Executive"].isin(executives_to_display)]
        total_sale_filtered = total_sale[total_sale["Executive"].isin(executives_to_display)]

        # CRITICAL FIX: Check if we have ANY data, not if ALL datasets have data
        if total_sale_filtered.empty and os_jan_filtered.empty and os_feb_filtered.empty:
            return {"success": False, "error": "No data found for selected executives."}
        
        # Check for partial data and prepare warning message
        executives_in_sale = set(total_sale_filtered["Executive"].unique())
        executives_in_os_jan = set(os_jan_filtered["Executive"].unique())
        executives_in_os_feb = set(os_feb_filtered["Executive"].unique())
        
        partial_data_execs = []
        warning_message = ""
        for exec_name in executives_to_display:
            if exec_name in executives_in_sale and (exec_name not in executives_in_os_jan or exec_name not in executives_in_os_feb):
                partial_data_execs.append(exec_name)
        
        if partial_data_execs:
            warning_message = f"Note: {', '.join(partial_data_execs)} found only in sales data, not in OS files. OS calculations will show as 0."

        specified_date = pd.to_datetime("01-" + selected_month_str, format="%d-%b-%y")
        specified_month_end = specified_date + pd.offsets.MonthEnd(0)

        # Due Target - Handle empty os_jan_filtered
        if not os_jan_filtered.empty:
            due_target_sum = os_jan_filtered[os_jan_filtered[os_jan_due_date_col] <= specified_month_end] \
                .groupby("Executive")[os_jan_net_value_col].sum().reset_index()
            due_target_sum.columns = ["Executive", "Due Target"]
        else:
            due_target_sum = pd.DataFrame(columns=["Executive", "Due Target"])

        # OS Jan Coll - Handle empty os_jan_filtered
        if not os_jan_filtered.empty:
            os_jan_coll_sum = os_jan_filtered[os_jan_filtered[os_jan_due_date_col] <= specified_month_end] \
                .groupby("Executive")[os_jan_net_value_col].sum().reset_index()
            os_jan_coll_sum.columns = ["Executive", "OS Jan Coll"]
        else:
            os_jan_coll_sum = pd.DataFrame(columns=["Executive", "OS Jan Coll"])

        # OS Feb Coll - Handle empty os_feb_filtered
        if not os_feb_filtered.empty:
            os_feb_coll_sum = os_feb_filtered[(os_feb_filtered[os_feb_ref_date_col] < specified_date) &
                                     (os_feb_filtered[os_feb_due_date_col] <= specified_month_end)] \
                .groupby("Executive")[os_feb_net_value_col].sum().reset_index()
            os_feb_coll_sum.columns = ["Executive", "OS Feb Coll"]
        else:
            os_feb_coll_sum = pd.DataFrame(columns=["Executive", "OS Feb Coll"])

        # Collection + early Overall %
        collection = os_jan_coll_sum.merge(os_feb_coll_sum, on="Executive", how="outer").fillna(0)
        collection["Collection Achieved"] = collection["OS Jan Coll"] - collection["OS Feb Coll"]
        collection = collection.merge(due_target_sum, on="Executive", how="outer").fillna(0)
        collection["Overall % Achieved"] = np.where(
            collection["Due Target"] > 0,
            (collection["Collection Achieved"] / collection["Due Target"]) * 100,
            0
        )

        # Overdue - Handle empty total_sale_filtered
        if not total_sale_filtered.empty:
            overdue_sum = total_sale_filtered[
                total_sale_filtered[sale_bill_date_col].between(specified_date, specified_month_end) &
                total_sale_filtered[sale_due_date_col].between(specified_date, specified_month_end)
            ].groupby("Executive")[sale_value_col].sum().reset_index()
            overdue_sum.columns = ["Executive", "For the month Overdue"]
        else:
            overdue_sum = pd.DataFrame(columns=["Executive", "For the month Overdue"])

        # Month collection
        if not overdue_sum.empty:
            sale_value_sum = overdue_sum.rename(columns={"For the month Overdue": "Sale Value"})
        else:
            sale_value_sum = pd.DataFrame(columns=["Executive", "Sale Value"])
        
        if not os_feb_filtered.empty:
            os_feb_month_sum = os_feb_filtered[
                os_feb_filtered[os_feb_ref_date_col].between(specified_date, specified_month_end) &
                os_feb_filtered[os_feb_due_date_col].between(specified_date, specified_month_end)
            ].groupby("Executive")[os_feb_net_value_col].sum().reset_index()
            os_feb_month_sum.columns = ["Executive", "OS Month Collection"]
        else:
            os_feb_month_sum = pd.DataFrame(columns=["Executive", "OS Month Collection"])

        month_collection = sale_value_sum.merge(os_feb_month_sum, on="Executive", how="outer").fillna(0)
        month_collection["For the month Collection"] = month_collection["Sale Value"] - month_collection["OS Month Collection"]
        month_collection["% Achieved (Selected Month)"] = np.where(
            month_collection["Sale Value"] > 0,
            (month_collection["For the month Collection"] / month_collection["Sale Value"]) * 100,
            0
        )

        # Merge all
        final = collection.drop(columns=["OS Jan Coll", "OS Feb Coll"], errors='ignore') \
            .merge(overdue_sum, on="Executive", how="outer") \
            .merge(month_collection[["Executive", "For the month Collection", "% Achieved (Selected Month)"]],
                   on="Executive", how="outer").fillna(0)

        # Ensure all selected executives are included
        final = pd.DataFrame({'Executive': executives_to_display}).merge(final, on='Executive', how='left').fillna(0)

        # Remove HO/HEAD OFFICE
        final = final[~final["Executive"].str.upper().isin(["HO", "HEAD OFFICE"])]

        # Scale + rounding
        val_cols = ["Due Target", "Collection Achieved", "For the month Overdue", "For the month Collection"]
        final[val_cols] = final[val_cols].div(100000).round(2)
        final[["Overall % Achieved", "% Achieved (Selected Month)"]] = final[["Overall % Achieved", "% Achieved (Selected Month)"]].round(2)

        # Sort + TOTAL
        final.sort_values("Executive", inplace=True)
        total_row = {'Executive': 'TOTAL'}
        for col in final.columns[1:]:
            if col in ["Overall % Achieved", "% Achieved (Selected Month)"]:
                # Handle weighted average for percentages
                if col == "Overall % Achieved":
                    weights = final["Due Target"]
                else:
                    weights = final["For the month Overdue"]
                
                if weights.sum() > 0:
                    total_row[col] = round(np.average(final[col], weights=weights), 2)
                else:
                    total_row[col] = 0
            else:
                total_row[col] = round(final[col].sum(), 2)
        final = pd.concat([final, pd.DataFrame([total_row])], ignore_index=True)

        # Prepare return data (matching original backend format)
        result = {
            "success": True, 
            "od_results": final.to_dict("records")
        }
        
        # Add warning message if there are executives with partial data
        if warning_message:
            result["warning"] = warning_message
            
        return result

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"success": False, "error": str(e)}


def load_data(file_path):
    """Load data from CSV or Excel file"""
    try:
        # Convert to absolute path
        if not os.path.isabs(file_path):
            file_path = os.path.abspath(file_path)
        
        if file_path.endswith('.csv'):
            return pd.read_csv(file_path)
        elif file_path.endswith(('.xlsx', '.xls')):
            return pd.read_excel(file_path)
        else:
            raise ValueError(f"Unsupported file format: {file_path}")
    except Exception as e:
        print(f"Error loading file {file_path}: {str(e)}")
        raise
