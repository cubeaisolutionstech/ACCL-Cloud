import pandas as pd
import numpy as np
import os
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Mappings from branch.py
branch_mapping = {
    'PONDY': 'PUDUCHERRY', 'PDY': 'PUDUCHERRY', 'Puducherry - PDY': 'PUDUCHERRY',
    'COVAI': 'COIMBATORE', 'CBE': 'COIMBATORE', 'Coimbatore - CBE': 'COIMBATORE',
    'ERD': 'ERODE', 'Erode - ERD': 'ERODE', 'ERD002': 'ERODE', 'ERD001': 'ERODE',
    'ERDTD1': 'ERODE', 'ERD003': 'ERODE', 'ERD004': 'ERODE', 'ERD005': 'ERODE',
    'ERD007': 'ERODE',
    'KRR': 'KARUR',
    'Chennai - CHN': 'CHENNAI', 'CHN': 'CHENNAI',
    'Tirupur - TPR': 'TIRUPUR', 'TPR': 'TIRUPUR',
    'Madurai - MDU': 'MADURAI', 'MDU': 'MADURAI',
    'POULTRY': 'POULTRY', 'Poultry - PLT': 'POULTRY',
    'SALEM': 'SALEM', 'Salem - SLM': 'SALEM',
    'HO': 'HO',
    'SLM002': 'SALEM', 'SLMTD1': 'SALEM',
    'BHV1': 'BHAVANI',
    'CBU': 'COIMBATORE',
    'VLR': 'VELLORE',
    'TRZ': 'TRICHY',
    'TVL': 'TIRUNELVELI',
    'NGS': 'NAGERCOIL',
    'PONDICHERRY': 'PUDUCHERRY',
    'BLR': 'BANGALORE', 'BANGALORE': 'BANGALORE', 'BGLR': 'BANGALORE'
}

def map_branch(branch_name):
    if pd.isna(branch_name):
        return 'Unknown'
    branch_str = str(branch_name).strip().upper()
    if ' - ' in branch_str:
        branch_str = branch_str.split(' - ')[-1].strip()
    return branch_mapping.get(branch_str, branch_str)

def find_column_by_names(columns, target_names):
    columns_upper = [col.upper() for col in columns]
    target_names_upper = [name.upper() for name in target_names]
    for target in target_names_upper:
        if target in columns_upper:
            return columns[columns_upper.index(target)]
    return None

def auto_map_budget_columns(sales_columns, budget_columns):
    sales_mapping = {
        'date': find_column_by_names(sales_columns, ['Date', 'Bill Date', 'Invoice Date']),
        'value': find_column_by_names(sales_columns, ['Value', 'Invoice Value', 'Amount']),
        'product_group': find_column_by_names(sales_columns, ['Type (Make)', 'Product Group', 'Type', 'Make']),
        'area': find_column_by_names(sales_columns, ['Branch', 'Area', 'Location']),
        'quantity': find_column_by_names(sales_columns, ['Actual Quantity', 'Quantity', 'Qty']),
        'sl_code': find_column_by_names(sales_columns, ['Customer Code', 'SL Code', 'Customer ID']),
        'executive': find_column_by_names(sales_columns, ['Executive', 'Sales Executive', 'Executive Name'])
    }
    budget_mapping = {
        'area': find_column_by_names(budget_columns, ['Branch', 'Area', 'Location']),
        'quantity': find_column_by_names(budget_columns, ["Qty – Apr'25", 'Quantity', 'Qty', 'Budget Qty']),
        'sl_code': find_column_by_names(budget_columns, ['SL Code', 'Customer Code', 'Customer ID']),
        'value': find_column_by_names(budget_columns, ["Value – Apr'25", 'Value', 'Budget Value', 'Amount']),
        'product_group': find_column_by_names(budget_columns, ['Product Group', 'Type(Make)', 'Product', 'Type']),
        'executive': find_column_by_names(budget_columns, ['Executive Name', 'Executive', 'Sales Executive'])
    }
    return sales_mapping, budget_mapping

def calculate_budget_vs_billed(data):
    try:
        # Create copies to avoid modifying original DataFrames
        sales_df = pd.read_excel(
            os.path.join('uploads', data['sales_filename']), 
            sheet_name=data['sales_sheet'], 
            header=data['sales_header'] - 1, 
            dtype={'SL Code': str}
        ).copy()
        
        budget_df = pd.read_excel(
            os.path.join('uploads', data['budget_filename']), 
            sheet_name=data['budget_sheet'], 
            header=data['budget_header'] - 1, 
            dtype={'SL Code': str}
        ).copy()
        
        print(f"🔍 Initial data loaded - Sales: {len(sales_df)} rows, Budget: {len(budget_df)} rows")
        
        # Convert and clean numeric and date columns
        sales_df[data['sales_date_col']] = pd.to_datetime(sales_df[data['sales_date_col']], dayfirst=True, errors='coerce')
        sales_df[data['sales_value_col']] = pd.to_numeric(sales_df[data['sales_value_col']], errors='coerce').fillna(0)
        sales_df[data['sales_qty_col']] = pd.to_numeric(sales_df[data['sales_qty_col']], errors='coerce').fillna(0)

        budget_df[data['budget_value_col']] = pd.to_numeric(budget_df[data['budget_value_col']], errors='coerce').fillna(0)
        budget_df[data['budget_qty_col']] = pd.to_numeric(budget_df[data['budget_qty_col']], errors='coerce').fillna(0)
        
        # Filter by executives
        if data['selected_sales_execs']:
            sales_df = sales_df[sales_df[data['sales_exec_col']].isin(data['selected_sales_execs'])].copy()
            print(f"🔍 After sales exec filter: {len(sales_df)} rows")
        if data['selected_budget_execs']:
            budget_df = budget_df[budget_df[data['budget_exec_col']].isin(data['selected_budget_execs'])].copy()
            print(f"🔍 After budget exec filter: {len(budget_df)} rows")

        if sales_df.empty or budget_df.empty:
            print("❌ No data found for selected executives in one or both files.")
            return None
            
        # Filter sales data for the selected month
        filtered_sales_df = sales_df[sales_df[data['sales_date_col']].dt.strftime('%b %y') == data['selected_month']].copy()
        print(f"🔍 After month filter ({data['selected_month']}): {len(filtered_sales_df)} sales rows")
        
        if filtered_sales_df.empty:
            print(f"❌ No sales data found for {data['selected_month']}")
            return None

        # Debug: Show sample data BEFORE processing
        print("🔍 BEFORE PROCESSING:")
        print("Sales sample (branch, sl_code, product):")
        print(filtered_sales_df[[data['sales_area_col'], data['sales_sl_code_col'], data['sales_product_group_col']]].head(3).to_string())
        print("Budget sample (branch, sl_code, product):")
        print(budget_df[[data['budget_area_col'], data['budget_sl_code_col'], data['budget_product_group_col']]].head(3).to_string())

        # Standardize string columns
        filtered_sales_df[data['sales_area_col']] = filtered_sales_df[data['sales_area_col']].astype(str).str.strip()
        budget_df[data['budget_area_col']] = budget_df[data['budget_area_col']].astype(str).str.strip()
        filtered_sales_df[data['sales_product_group_col']] = filtered_sales_df[data['sales_product_group_col']].astype(str).str.strip()
        filtered_sales_df[data['sales_sl_code_col']] = filtered_sales_df[data['sales_sl_code_col']].astype(str).str.strip().str.replace('\.0$', '', regex=True)
        budget_df[data['budget_product_group_col']] = budget_df[data['budget_product_group_col']].astype(str).str.strip()
        budget_df[data['budget_sl_code_col']] = budget_df[data['budget_sl_code_col']].astype(str).str.strip().str.replace('\.0$', '', regex=True)

        # Apply branch mapping exactly like Streamlit
        budget_df[data['budget_area_col']] = budget_df[data['budget_area_col']].str.split(' - ').str[-1].str.upper()
        budget_df[data['budget_area_col']] = budget_df[data['budget_area_col']].str.replace('AAAA - ', '', regex=False).str.upper()
        budget_df[data['budget_area_col']] = budget_df[data['budget_area_col']].replace(branch_mapping)
        filtered_sales_df[data['sales_area_col']] = filtered_sales_df[data['sales_area_col']].str.upper().replace(branch_mapping)
        
        # Apply branch filtering if specified
        if data['selected_branches']:
            filtered_sales_df = filtered_sales_df[filtered_sales_df[data['sales_area_col']].isin(data['selected_branches'])].copy()
            budget_df = budget_df[budget_df[data['budget_area_col']].isin(data['selected_branches'])].copy()
            print(f"🔍 After branch filter: Sales {len(filtered_sales_df)}, Budget {len(budget_df)}")

        # Standardize product groups and SL codes exactly like Streamlit
        filtered_sales_df[data['sales_product_group_col']] = filtered_sales_df[data['sales_product_group_col']].str.upper()
        filtered_sales_df[data['sales_sl_code_col']] = filtered_sales_df[data['sales_sl_code_col']].str.upper()
        budget_df[data['budget_product_group_col']] = budget_df[data['budget_product_group_col']].str.upper()
        budget_df[data['budget_sl_code_col']] = budget_df[data['budget_sl_code_col']].str.upper()

        # Debug: Show sample data AFTER processing
        print("🔍 AFTER PROCESSING:")
        print("Sales sample (branch, sl_code, product):")
        print(filtered_sales_df[[data['sales_area_col'], data['sales_sl_code_col'], data['sales_product_group_col']]].head(3).to_string())
        print("Budget sample (branch, sl_code, product):")
        print(budget_df[[data['budget_area_col'], data['budget_sl_code_col'], data['budget_product_group_col']]].head(3).to_string())

        # Group by Branch + SL Code + Product Group and sum quantities/values
        budget_grouped = budget_df.groupby([
            data['budget_area_col'],
            data['budget_sl_code_col'], 
            data['budget_product_group_col']
        ]).agg({
            data['budget_qty_col']: 'sum',
            data['budget_value_col']: 'sum'
        }).reset_index()
        
        # Filter: Only include rows where BOTH qty > 0 AND value > 0
        budget_valid = budget_grouped[
            (budget_grouped[data['budget_qty_col']] > 0) & 
            (budget_grouped[data['budget_value_col']] > 0)
        ].copy()
        
        print(f"🔍 Valid budget combinations: {len(budget_valid)}")
        
        if budget_valid.empty:
            print("❌ No valid budget data found (with qty > 0 and value > 0).")
            return None
        
        # Debug: Show what we're trying to match
        print("🔍 Budget combinations to match:")
        print(budget_valid[[data['budget_area_col'], data['budget_sl_code_col'], data['budget_product_group_col']]].head(5).to_string())
        
        # Initialize results storage
        final_results = []
        matches_found = 0
        
        # For each valid budget record, find matching sales
        for i, budget_row in budget_valid.iterrows():
            branch = budget_row[data['budget_area_col']]
            sl_code = budget_row[data['budget_sl_code_col']]
            product = budget_row[data['budget_product_group_col']]
            budget_qty = budget_row[data['budget_qty_col']]
            budget_value = budget_row[data['budget_value_col']]
            
            # Find matching sales records (same branch + sl_code + product)
            matching_sales = filtered_sales_df[
                (filtered_sales_df[data['sales_area_col']] == branch) &
                (filtered_sales_df[data['sales_sl_code_col']] == sl_code) &
                (filtered_sales_df[data['sales_product_group_col']] == product)
            ]
            
            # Sum all matching sales records for this combination
            if not matching_sales.empty:
                sales_qty_total = matching_sales[data['sales_qty_col']].sum()
                sales_value_total = matching_sales[data['sales_value_col']].sum()
                matches_found += 1
                
                # Debug first few matches
                if matches_found <= 3:
                    print(f"🔍 MATCH {matches_found}: {branch}/{sl_code}/{product} -> Sales Qty: {sales_qty_total}, Sales Value: {sales_value_total}")
            else:
                sales_qty_total = 0
                sales_value_total = 0
            
            # Apply the comparison logic for budget vs billed reports
            final_qty = budget_qty if sales_qty_total > budget_qty else sales_qty_total
            final_value = budget_value if sales_value_total > budget_value else sales_value_total
            
            # Store result
            final_results.append({
                'Branch': branch,
                'SL_Code': sl_code,
                'Product': product,
                'Budget_Qty': budget_qty,
                'Sales_Qty': sales_qty_total,
                'Final_Qty': final_qty,
                'Budget_Value': budget_value,
                'Sales_Value': sales_value_total,
                'Final_Value': final_value
            })
        
        print(f"🔍 Total matches found: {matches_found} out of {len(budget_valid)} budget records")
        
        # Convert to DataFrame for easier manipulation
        results_df = pd.DataFrame(final_results)
        
        if results_df.empty:
            print("❌ No results generated")
            return None
            
        # Debug: Show results summary
        print("🔍 Results summary by branch:")
        branch_summary = results_df.groupby('Branch').agg({
            'Budget_Qty': 'sum',
            'Sales_Qty': 'sum', 
            'Final_Qty': 'sum'
        })
        print(branch_summary.to_string())
        
        # Continue with the rest of the processing...
        # [Rest of the code remains the same as before]
        
        # Get default branches
        if data['selected_branches']:
            default_branches = data['selected_branches']
        else:
            all_sales_branches = filtered_sales_df[data['sales_area_col']].dropna().unique().tolist()
            all_budget_branches = budget_df[data['budget_area_col']].dropna().unique().tolist()
            default_branches = sorted(set(all_sales_branches + all_budget_branches))

        # Build QUANTITY DataFrame
        qty_data = []
        for area in default_branches:
            area_qty_data = results_df[results_df['Branch'] == area]
            
            if not area_qty_data.empty:
                budget_val = round(float(area_qty_data['Budget_Qty'].sum()), 2)
                billed_val = round(float(area_qty_data['Final_Qty'].sum()), 2)
            else:
                budget_val = 0.0
                billed_val = 0.0
            
            percentage = round((billed_val / budget_val) * 100, 2) if budget_val > 0 else 0.0
            
            qty_data.append({
                'Area': area,
                'Budget Qty': budget_val,
                'Billed Qty': billed_val,
                '%': percentage
            })
        
        budget_vs_billed_qty_df = pd.DataFrame(qty_data)
        
        # Build VALUE DataFrame
        value_data = []
        for area in default_branches:
            area_value_data = results_df[results_df['Branch'] == area]
            
            if not area_value_data.empty:
                budget_val = round(float(area_value_data['Budget_Value'].sum()), 2)
                billed_val = round(float(area_value_data['Final_Value'].sum()), 2)
            else:
                budget_val = 0.0
                billed_val = 0.0
            
            percentage = round((billed_val / budget_val) * 100, 2) if budget_val > 0 else 0.0
            
            value_data.append({
                'Area': area,
                'Budget Value': budget_val,
                'Billed Value': billed_val,
                '%': percentage
            })
        
        budget_vs_billed_value_df = pd.DataFrame(value_data)
        
        # Create Overall Sales DataFrames
        overall_sales_data = filtered_sales_df.groupby(data['sales_area_col']).agg({
            data['sales_qty_col']: 'sum',
            data['sales_value_col']: 'sum'
        }).reset_index()
        overall_sales_data.columns = ['Area', 'Overall_Sales_Qty', 'Overall_Sales_Value']
        
        budget_totals = results_df.groupby('Branch').agg({
            'Budget_Qty': 'sum',
            'Budget_Value': 'sum'
        }).reset_index()
        budget_totals.columns = ['Area', 'Budget_Qty', 'Budget_Value']
        
        # Overall Sales Quantity DataFrame
        overall_sales_qty_df = pd.DataFrame({'Area': default_branches})
        overall_sales_qty_df = pd.merge(
            overall_sales_qty_df,
            budget_totals[['Area', 'Budget_Qty']].rename(columns={'Budget_Qty': 'Budget Qty'}),
            on='Area',
            how='left'
        ).fillna({'Budget Qty': 0})
        
        overall_sales_qty_df = pd.merge(
            overall_sales_qty_df,
            overall_sales_data[['Area', 'Overall_Sales_Qty']].rename(columns={'Overall_Sales_Qty': 'Billed Qty'}),
            on='Area',
            how='left'
        ).fillna({'Billed Qty': 0})
        
        # Overall Sales Value DataFrame
        overall_sales_value_df = pd.DataFrame({'Area': default_branches})
        overall_sales_value_df = pd.merge(
            overall_sales_value_df,
            budget_totals[['Area', 'Budget_Value']].rename(columns={'Budget_Value': 'Budget Value'}),
            on='Area',
            how='left'
        ).fillna({'Budget Value': 0})
        
        overall_sales_value_df = pd.merge(
            overall_sales_value_df,
            overall_sales_data[['Area', 'Overall_Sales_Value']].rename(columns={'Overall_Sales_Value': 'Billed Value'}),
            on='Area',
            how='left'
        ).fillna({'Billed Value': 0})
        
        # Add Total Rows
        total_budget_qty = round(budget_vs_billed_qty_df['Budget Qty'].sum(), 2)
        total_billed_qty = round(budget_vs_billed_qty_df['Billed Qty'].sum(), 2)
        total_percentage_qty = round((total_billed_qty / total_budget_qty * 100), 2) if total_budget_qty > 0 else 0.0
        
        total_row_qty = pd.DataFrame({
            'Area': ['TOTAL'],
            'Budget Qty': [total_budget_qty],
            'Billed Qty': [total_billed_qty],
            '%': [total_percentage_qty]
        })
        budget_vs_billed_qty_df = pd.concat([budget_vs_billed_qty_df, total_row_qty], ignore_index=True)
        
        total_budget_value = round(budget_vs_billed_value_df['Budget Value'].sum(), 2)
        total_billed_value = round(budget_vs_billed_value_df['Billed Value'].sum(), 2)
        total_percentage_value = round((total_billed_value / total_budget_value * 100), 2) if total_budget_value > 0 else 0.0
        
        total_row_value = pd.DataFrame({
            'Area': ['TOTAL'],
            'Budget Value': [total_budget_value],
            'Billed Value': [total_billed_value],
            '%': [total_percentage_value]
        })
        budget_vs_billed_value_df = pd.concat([budget_vs_billed_value_df, total_row_value], ignore_index=True)
        
        total_row_overall_qty = pd.DataFrame({
            'Area': ['TOTAL'],
            'Budget Qty': [round(overall_sales_qty_df['Budget Qty'].sum(), 2)],
            'Billed Qty': [round(overall_sales_qty_df['Billed Qty'].sum(), 2)]
        })
        overall_sales_qty_df = pd.concat([overall_sales_qty_df, total_row_overall_qty], ignore_index=True)
        
        total_row_overall_value = pd.DataFrame({
            'Area': ['TOTAL'],
            'Budget Value': [round(overall_sales_value_df['Budget Value'].sum(), 2)],
            'Billed Value': [round(overall_sales_value_df['Billed Value'].sum(), 2)]
        })
        overall_sales_value_df = pd.concat([overall_sales_value_df, total_row_overall_value], ignore_index=True)
        
        # Convert to integers for final display
        budget_vs_billed_value_df['Budget Value'] = budget_vs_billed_value_df['Budget Value'].round(0).astype(int)
        budget_vs_billed_value_df['Billed Value'] = budget_vs_billed_value_df['Billed Value'].round(0).astype(int)
        budget_vs_billed_qty_df['Budget Qty'] = budget_vs_billed_qty_df['Budget Qty'].round(0).astype(int)
        budget_vs_billed_qty_df['Billed Qty'] = budget_vs_billed_qty_df['Billed Qty'].round(0).astype(int)
        overall_sales_qty_df['Budget Qty'] = overall_sales_qty_df['Budget Qty'].round(0).astype(int)
        overall_sales_qty_df['Billed Qty'] = overall_sales_qty_df['Billed Qty'].round(0).astype(int)
        overall_sales_value_df['Budget Value'] = overall_sales_value_df['Budget Value'].round(0).astype(int)
        overall_sales_value_df['Billed Value'] = overall_sales_value_df['Billed Value'].round(0).astype(int)
        
        budget_vs_billed_qty_df['%'] = budget_vs_billed_qty_df['%'].round(0).astype(int)
        budget_vs_billed_value_df['%'] = budget_vs_billed_value_df['%'].round(0).astype(int)
        
        print("✅ Budget vs Billed calculation completed successfully!")
        
        return {
            'budget_vs_billed_qty':{
                'data': budget_vs_billed_qty_df.to_dict(orient='records'),
                'columns': budget_vs_billed_qty_df.columns.tolist(),
            },
            'budget_vs_billed_value': {
                'data': budget_vs_billed_value_df.to_dict(orient='records'),
                'columns': budget_vs_billed_value_df.columns.tolist()
            },
            'overall_sales_qty': {
                'data': overall_sales_qty_df.to_dict(orient='records'),
                'columns': overall_sales_qty_df.columns.tolist()
            },
            'overall_sales_value': {
                'data': overall_sales_value_df.to_dict(orient='records'),
                'columns': overall_sales_value_df.columns.tolist()
            }
        }

    except Exception as e:
        logger.error(f"Error in calculate_budget_vs_billed: {str(e)}")
        print(f"❌ Error calculating budget values: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return None
