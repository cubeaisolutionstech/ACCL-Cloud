# Replace your entire executive_product_growth.py file with this corrected version

import pandas as pd
import numpy as np
import os
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

def standardize_name(name):
    """Standardize company/product group names - EXACT STREAMLIT LOGIC"""
    if pd.isna(name) or not name:
        return ""
    name = str(name).strip().lower()
    name = ''.join(c for c in name if c.isalnum() or c.isspace())
    name = ' '.join(word.capitalize() for word in name.split())
    
    # Handle special cases for 'General'
    general_variants = ['general', 'gen', 'generals', 'general ', 'genral', 'generl']
    if any(variant in name.lower() for variant in general_variants):
        return 'General'
    
    return name

def create_sl_code_mapping(ly_df, cy_df, budget_df, ly_sl_code_col, cy_sl_code_col, budget_sl_code_col, 
                           ly_company_group_col, cy_company_group_col, budget_company_group_col):
    """Create SL Code to Company Group mapping - EXACT STREAMLIT LOGIC"""
    try:
        mappings = []
        
        for df, sl_code_col, company_group_col in [
            (ly_df, ly_sl_code_col, ly_company_group_col),
            (cy_df, cy_sl_code_col, cy_company_group_col),
            (budget_df, budget_sl_code_col, budget_company_group_col)
        ]:
            if sl_code_col and sl_code_col in df.columns and company_group_col in df.columns:
                subset = df[[sl_code_col, company_group_col]].dropna()
                subset = subset[subset[sl_code_col] != ""]
                mappings.append(subset.rename(columns={sl_code_col: 'SL_CODE', company_group_col: 'COMPANY_GROUP'}))
        
        if not mappings:
            return {}
        
        combined = pd.concat(mappings, ignore_index=True)
        combined['COMPANY_GROUP'] = combined['COMPANY_GROUP'].apply(standardize_name)
        
        mapping_df = combined.groupby('SL_CODE')['COMPANY_GROUP'].agg(lambda x: x.mode()[0] if not x.empty else "").reset_index()
        sl_code_map = dict(zip(mapping_df['SL_CODE'], mapping_df['COMPANY_GROUP']))
        
        return sl_code_map
        
    except Exception as e:
        logger.error(f"Error creating SL Code mapping: {e}")
        return {}

def apply_sl_code_mapping(df, sl_code_col, company_group_col, sl_code_map):
    """Apply SL Code mapping - EXACT STREAMLIT LOGIC"""
    if not sl_code_col or sl_code_col not in df.columns or not sl_code_map:
        return df[company_group_col].apply(standardize_name)
    
    try:
        def map_company(row):
            if pd.isna(row[sl_code_col]) or row[sl_code_col] == "":
                return standardize_name(row[company_group_col])
            sl_code = str(row[sl_code_col]).strip()
            return sl_code_map.get(sl_code, standardize_name(row[company_group_col]))
        
        return df.apply(map_company, axis=1)
        
    except Exception as e:
        logger.error(f"Error applying SL Code mapping: {e}")
        return df[company_group_col].apply(standardize_name)

def auto_map_product_growth_columns(ly_df, cy_df, budget_df):
    """Auto-map columns - EXACT STREAMLIT LOGIC"""
    try:
        def find_column(columns, target_names):
            for target in target_names:
                for col in columns:
                    if col.lower() == target.lower():
                        return col
            return None

        ly_mappings = {
            'date': ['Date'],
            'value': ['Value', 'Invoice Value'],
            'quantity': ['Actual Quantity', 'Quantity', 'Qty'],
            'product_group': ['Type (Make)', 'Product Group'],
            'company_group': ['Company Group'],
            'executive': ['Executive Name', 'Executive'],
            'sl_code': ['Customer Code', 'SL Code']
        }

        cy_mappings = {
            'date': ['Date'],
            'value': ['Value', 'Invoice Value'],
            'quantity': ['Actual Quantity', 'Quantity', 'Qty'],
            'product_group': ['Type (Make)', 'Product Group'],
            'company_group': ['Company Group'],
            'executive': ['Executive Name', 'Executive'],
            'sl_code': ['Customer Code', 'SL Code']
        }

        budget_mappings = {
            'company_group': ['Company Group'],
            'product_group': ['Product Group', 'Type (Make)'],
            'quantity': ['Budget Quantity', 'Quantity', 'Qty'],
            'value': ['Budget Value', 'Value'],
            'executive': ['Executive Name', 'Executive'],
            'sl_code': ['SL Code', 'Customer Code']
        }

        ly_mapping = {key: find_column(ly_df.columns.tolist(), targets) for key, targets in ly_mappings.items()}
        cy_mapping = {key: find_column(cy_df.columns.tolist(), targets) for key, targets in cy_mappings.items()}
        budget_mapping = {key: find_column(budget_df.columns.tolist(), targets) for key, targets in budget_mappings.items()}

        return {
            'ly_mapping': ly_mapping,
            'cy_mapping': cy_mapping,
            'budget_mapping': budget_mapping,
            'ly_columns': list(ly_df.columns),
            'cy_columns': list(cy_df.columns),
            'budget_columns': list(budget_df.columns)
        }
        
    except Exception as e:
        return {
            'ly_mapping': {}, 'cy_mapping': {}, 'budget_mapping': {},
            'ly_columns': [], 'cy_columns': [], 'budget_columns': []
        }

def get_product_growth_options(ly_df, cy_df, budget_df, ly_exec_col, cy_exec_col, budget_exec_col,
                              ly_company_group_col, cy_company_group_col, budget_company_group_col,
                              ly_product_group_col, cy_product_group_col, budget_product_group_col,
                              ly_sl_code_col, cy_sl_code_col, budget_sl_code_col):
    """Get options - EXACT STREAMLIT LOGIC"""
    try:
        sl_code_map = create_sl_code_mapping(
            ly_df, cy_df, budget_df, 
            ly_sl_code_col, cy_sl_code_col, budget_sl_code_col,
            ly_company_group_col, cy_company_group_col, budget_company_group_col
        )
        
        ly_df = ly_df.copy()
        cy_df = cy_df.copy()
        budget_df = budget_df.copy()
        
        ly_df[ly_company_group_col] = apply_sl_code_mapping(ly_df, ly_sl_code_col, ly_company_group_col, sl_code_map)
        cy_df[cy_company_group_col] = apply_sl_code_mapping(cy_df, cy_sl_code_col, cy_company_group_col, sl_code_map)
        budget_df[budget_company_group_col] = apply_sl_code_mapping(budget_df, budget_sl_code_col, budget_company_group_col, sl_code_map)
        
        ly_df[ly_product_group_col] = ly_df[ly_product_group_col].apply(standardize_name)
        cy_df[cy_product_group_col] = cy_df[cy_product_group_col].apply(standardize_name)
        budget_df[budget_product_group_col] = budget_df[budget_product_group_col].apply(standardize_name)
        
        all_executives = set()
        for df, exec_col in [(ly_df, ly_exec_col), (cy_df, cy_exec_col), (budget_df, budget_exec_col)]:
            if exec_col and exec_col in df.columns:
                execs = df[exec_col].dropna().astype(str).str.strip().str.upper().unique().tolist()
                all_executives.update(execs)
        
        company_groups = pd.concat([
            ly_df[ly_company_group_col],
            cy_df[cy_company_group_col],
            budget_df[budget_company_group_col]
        ]).dropna().unique().tolist()
        company_groups = sorted([g for g in company_groups if g])
        
        product_groups = pd.concat([
            ly_df[ly_product_group_col],
            cy_df[cy_product_group_col],
            budget_df[budget_product_group_col]
        ]).dropna().unique().tolist()
        product_groups = sorted([g for g in product_groups if g])
        
        return {
            'executives': sorted(list(all_executives)),
            'company_groups': company_groups,
            'product_groups': product_groups
        }
        
    except Exception as e:
        return {'executives': [], 'company_groups': [], 'product_groups': []}

def get_product_growth_months(ly_df, cy_df, ly_date_col, cy_date_col):
    """Get months - EXACT STREAMLIT LOGIC"""
    try:
        ly_df = ly_df.copy()
        cy_df = cy_df.copy()
        
        ly_df[ly_date_col] = pd.to_datetime(ly_df[ly_date_col], dayfirst=True, errors='coerce')
        cy_df[cy_date_col] = pd.to_datetime(cy_df[cy_date_col], dayfirst=True, errors='coerce')
        
        available_ly_months = ly_df[ly_date_col].dt.strftime('%b %y').dropna().unique().tolist()
        available_cy_months = cy_df[cy_date_col].dt.strftime('%b %y').dropna().unique().tolist()
        
        return {
            'ly_months': sorted(available_ly_months),
            'cy_months': sorted(available_cy_months)
        }
        
    except Exception as e:
        return {'ly_months': [], 'cy_months': []}

def calculate_product_growth(ly_df, cy_df, budget_df, ly_month, cy_month, ly_date_col, cy_date_col, 
                            ly_qty_col, cy_qty_col, ly_value_col, cy_value_col, 
                            budget_qty_col, budget_value_col, ly_company_group_col, 
                            cy_company_group_col, budget_company_group_col, 
                            ly_product_group_col, cy_product_group_col, budget_product_group_col,
                            ly_sl_code_col, cy_sl_code_col, budget_sl_code_col,
                            ly_exec_col, cy_exec_col, budget_exec_col, 
                            selected_executives=None, selected_company_groups=None, selected_product_groups=None):
    """
    Calculate product growth for last year, current year, and budget data, aligned with Streamlit logic.
    """
    try:
        logger.info("Starting product growth calculation...")
        
        # Create copies to avoid modifying original DataFrames
        ly_df = ly_df.copy()
        cy_df = cy_df.copy()
        budget_df = budget_df.copy()
        
        # Replace empty strings with NaN and standardize, as in Streamlit
        for df, col in [
            (ly_df, ly_company_group_col), (cy_df, cy_company_group_col), (budget_df, budget_company_group_col),
            (ly_df, ly_product_group_col), (cy_df, cy_product_group_col), (budget_df, budget_product_group_col)
        ]:
            df[col] = df[col].replace("", np.nan).fillna("")
        
        # Validate required columns (relaxed for budget_sl_code_col)
        required_cols = [
            (ly_df, [ly_date_col, ly_qty_col, ly_value_col, ly_company_group_col, ly_product_group_col, ly_exec_col, ly_sl_code_col]),
            (cy_df, [cy_date_col, cy_qty_col, cy_value_col, cy_company_group_col, cy_product_group_col, cy_exec_col, cy_sl_code_col]),
            (budget_df, [budget_qty_col, budget_value_col, budget_company_group_col, budget_product_group_col, budget_exec_col])
        ]
        
        for df, cols in required_cols:
            missing_cols = [col for col in cols if col not in df.columns]
            if missing_cols:
                logger.error(f"Missing columns in DataFrame: {missing_cols}")
                return {"success": False, "error": f"Missing columns in DataFrame: {missing_cols}"}
        
        # Validate optional SL code column for budget
        if budget_sl_code_col and budget_sl_code_col not in budget_df.columns:
            logger.warning(f"Budget SL code column '{budget_sl_code_col}' not found. Proceeding without SL code mapping for budget.")
            budget_sl_code_col = None
        
        # Create SL code mapping (skip if budget_sl_code_col is None)
        if budget_sl_code_col:
            sl_code_map = create_sl_code_mapping(
                ly_df, cy_df, budget_df, 
                ly_sl_code_col, cy_sl_code_col, budget_sl_code_col,
                ly_company_group_col, cy_company_group_col, budget_company_group_col
            )
            ly_df[ly_company_group_col] = apply_sl_code_mapping(ly_df, ly_sl_code_col, ly_company_group_col, sl_code_map)
            cy_df[cy_company_group_col] = apply_sl_code_mapping(cy_df, cy_sl_code_col, cy_company_group_col, sl_code_map)
            budget_df[budget_company_group_col] = apply_sl_code_mapping(budget_df, budget_sl_code_col, budget_company_group_col, sl_code_map)
        else:
            # Standardize company group names without SL code mapping
            ly_df[ly_company_group_col] = ly_df[ly_company_group_col].apply(standardize_name)
            cy_df[cy_company_group_col] = cy_df[cy_company_group_col].apply(standardize_name)
            budget_df[budget_company_group_col] = budget_df[budget_company_group_col].apply(standardize_name)
        
        # Standardize product group names
        ly_df[ly_product_group_col] = ly_df[ly_product_group_col].apply(standardize_name)
        cy_df[cy_product_group_col] = cy_df[cy_product_group_col].apply(standardize_name)
        budget_df[budget_product_group_col] = budget_df[budget_product_group_col].apply(standardize_name)
        
        # Filter by executives
        if selected_executives:
            selected_executives = [str(e).strip().upper() for e in selected_executives]
            if ly_exec_col in ly_df.columns:
                ly_df = ly_df[ly_df[ly_exec_col].astype(str).str.strip().str.upper().isin(selected_executives)]
            if cy_exec_col in cy_df.columns:
                cy_df = cy_df[cy_df[cy_exec_col].astype(str).str.strip().str.upper().isin(selected_executives)]
            if budget_exec_col in budget_df.columns:
                budget_df = budget_df[budget_df[budget_exec_col].astype(str).str.strip().str.upper().isin(selected_executives)]
        
        if ly_df.empty or cy_df.empty or budget_df.empty:
            logger.warning("One or more DataFrames are empty after executive filtering.")
            return {"success": False, "error": "No data remains after executive filtering. Please check executive selections."}
        
        # Convert dates
        ly_df[ly_date_col] = pd.to_datetime(ly_df[ly_date_col], dayfirst=True, errors='coerce', format='mixed')
        cy_df[cy_date_col] = pd.to_datetime(cy_df[cy_date_col], dayfirst=True, errors='coerce', format='mixed')
        
        # Get available months
        available_ly_months = ly_df[ly_date_col].dt.strftime('%b %y').dropna().unique().tolist()
        available_cy_months = cy_df[cy_date_col].dt.strftime('%b %y').dropna().unique().tolist()
        
        if not available_ly_months or not available_cy_months:
            logger.error("No valid dates found in LY or CY data.")
            return {"success": False, "error": "No valid dates found in LY or CY data. Please check date columns."}
        
        # Filter by months
        if ly_month:
            ly_filtered_df = ly_df[ly_df[ly_date_col].dt.strftime('%b %y') == ly_month]
        else:
            latest_ly_month = max(available_ly_months, key=lambda x: pd.to_datetime(f"01 {x}", format="%d %b %y"))
            ly_filtered_df = ly_df[ly_df[ly_date_col].dt.strftime('%b %y') == latest_ly_month]
            ly_month = latest_ly_month
        
        if cy_month:
            cy_filtered_df = cy_df[cy_df[cy_date_col].dt.strftime('%b %y') == cy_month]
        else:
            latest_cy_month = max(available_cy_months, key=lambda x: pd.to_datetime(f"01 {x}", format="%d %b %y"))
            cy_filtered_df = cy_df[cy_df[cy_date_col].dt.strftime('%b %y') == latest_cy_month]
            cy_month = latest_cy_month
        
        if ly_filtered_df.empty or cy_filtered_df.empty:
            logger.error(f"No data for selected months (LY: {ly_month}, CY: {cy_month}).")
            return {"success": False, "error": f"No data for selected months (LY: {ly_month}, CY: {cy_month}). Please check month selections."}
        
        # Get company groups
        company_groups = pd.concat([
            ly_filtered_df[ly_company_group_col], 
            cy_filtered_df[cy_company_group_col], 
            budget_df[budget_company_group_col]
        ]).dropna().unique().tolist()
        
        # Filter by company groups
        if selected_company_groups:
            selected_company_groups = [standardize_name(g) for g in selected_company_groups]
            valid_groups = set(company_groups)
            invalid_groups = [g for g in selected_company_groups if g not in valid_groups]
            if invalid_groups:
                logger.warning(f"The following company groups are not found in the data: {invalid_groups}. Proceeding with valid groups.")
                selected_company_groups = [g for g in selected_company_groups if g in valid_groups]
                if not selected_company_groups:
                    return {"success": False, "error": "No valid company groups selected after validation. Please select valid company groups."}
            
            ly_filtered_df = ly_filtered_df[ly_filtered_df[ly_company_group_col].isin(selected_company_groups)]
            cy_filtered_df = cy_filtered_df[cy_filtered_df[cy_company_group_col].isin(selected_company_groups)]
            budget_df = budget_df[budget_df[budget_company_group_col].isin(selected_company_groups)]
            if ly_filtered_df.empty or cy_filtered_df.empty or budget_df.empty:
                logger.warning(f"No data remains after filtering for company groups: {selected_company_groups}.")
                return {"success": False, "error": f"No data remains after filtering for company groups: {selected_company_groups}. Please check company group selections or data content."}
        
        # Get product groups
        product_groups = pd.concat([
            ly_filtered_df[ly_product_group_col], 
            cy_filtered_df[cy_product_group_col], 
            budget_df[budget_product_group_col]
        ]).dropna().unique().tolist()
        
        # Filter by product groups
        if selected_product_groups:
            selected_product_groups = [standardize_name(g) for g in selected_product_groups]
            valid_product_groups = set(product_groups)
            invalid_product_groups = [g for g in selected_product_groups if g not in valid_product_groups]
            if invalid_product_groups:
                logger.warning(f"The following product groups are not found in the data: {invalid_product_groups}. Proceeding with valid groups.")
                selected_product_groups = [g for g in selected_product_groups if g in valid_product_groups]
                if not selected_product_groups:
                    return {"success": False, "error": "No valid product groups selected after validation. Please select valid product groups."}
            
            ly_filtered_df = ly_filtered_df[ly_filtered_df[ly_product_group_col].isin(selected_product_groups)]
            cy_filtered_df = cy_filtered_df[cy_filtered_df[cy_product_group_col].isin(selected_product_groups)]
            budget_df = budget_df[budget_df[budget_product_group_col].isin(selected_product_groups)]
            if ly_filtered_df.empty or cy_filtered_df.empty or budget_df.empty:
                logger.warning(f"No data remains after filtering for product groups: {selected_product_groups}.")
                return {"success": False, "error": f"No data remains after filtering for product groups: {selected_product_groups}. Please check product group selections or data content."}
        
        # Convert numeric columns
        for df, qty_col, value_col in [
            (ly_filtered_df, ly_qty_col, ly_value_col), 
            (cy_filtered_df, cy_qty_col, cy_value_col), 
            (budget_df, budget_qty_col, budget_value_col)
        ]:
            df[qty_col] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0)
            df[value_col] = pd.to_numeric(df[value_col], errors='coerce').fillna(0)
            if df[qty_col].isna().all() or df[value_col].isna().all():
                logger.warning(f"Invalid numeric data in {qty_col} or {value_col}. All values converted to 0.")
        
        # Get final company groups
        company_groups = selected_company_groups if selected_company_groups else sorted(set(company_groups))
        
        if not company_groups:
            logger.error("No valid company groups found in the data.")
            return {"success": False, "error": "No valid company groups found in the data. Please check company group columns."}
        
        # Helper functions for calculations
        def calc_achievement(row, cy_col, ly_col):
            """Calculate achievement percentage for individual products"""
            if pd.isna(row[ly_col]) or row[ly_col] == 0:
                return 0.00 if row[cy_col] == 0 else 100.00
            return round(((row[cy_col] - row[ly_col]) / row[ly_col]) * 100, 2)

        def calc_total_growth_percentage(total_cy, total_ly):
            """Calculate total growth percentage based on total values"""
            if pd.isna(total_ly) or total_ly == 0:
                return 0.00 if total_cy == 0 else 100.00
            return round(((total_cy - total_ly) / total_ly) * 100, 2)
        
        # Result dictionary
        result = {}
        
        for company in company_groups:
            # Initialize DataFrames
            qty_df = pd.DataFrame(columns=['PRODUCT GROUP', 'LY_QTY', 'BUDGET_QTY', 'CY_QTY', 'ACHIEVEMENT %'])
            value_df = pd.DataFrame(columns=['PRODUCT GROUP', 'LY_VALUE', 'BUDGET_VALUE', 'CY_VALUE', 'ACHIEVEMENT %'])
            
            # Filter data for this company
            ly_company_df = ly_filtered_df[ly_filtered_df[ly_company_group_col] == company]
            cy_company_df = cy_filtered_df[cy_filtered_df[cy_company_group_col] == company]
            budget_company_df = budget_df[budget_df[budget_company_group_col] == company]
            
            if ly_company_df.empty and cy_company_df.empty and budget_company_df.empty:
                logger.warning(f"No data for company group: {company}. Skipping.")
                continue
            
            # Get product groups for this company
            company_product_groups = pd.concat([
                ly_company_df[ly_product_group_col],
                cy_company_df[cy_product_group_col],
                budget_company_df[budget_product_group_col]
            ]).dropna().unique().tolist()
            
            if not company_product_groups:
                logger.warning(f"No product groups for company: {company}. Skipping.")
                continue
            
            if selected_product_groups:
                company_product_groups = [pg for pg in company_product_groups if pg in selected_product_groups]
                if not company_product_groups:
                    logger.warning(f"No valid product groups for company: {company} after filtering. Skipping.")
                    continue
            
            # Aggregate data
            ly_qty = ly_company_df.groupby([ly_company_group_col, ly_product_group_col])[ly_qty_col].sum().reset_index()
            ly_qty = ly_qty.rename(columns={ly_product_group_col: 'PRODUCT GROUP', ly_qty_col: 'LY_QTY'})
            
            cy_qty = cy_company_df.groupby([cy_company_group_col, cy_product_group_col])[cy_qty_col].sum().reset_index()
            cy_qty = cy_qty.rename(columns={cy_product_group_col: 'PRODUCT GROUP', cy_qty_col: 'CY_QTY'})
            
            budget_qty = budget_company_df.groupby([budget_company_group_col, budget_product_group_col])[budget_qty_col].sum().reset_index()
            budget_qty = budget_qty.rename(columns={budget_product_group_col: 'PRODUCT GROUP', budget_qty_col: 'BUDGET_QTY'})
            
            ly_value = ly_company_df.groupby([ly_company_group_col, ly_product_group_col])[ly_value_col].sum().reset_index()
            ly_value = ly_value.rename(columns={ly_product_group_col: 'PRODUCT GROUP', ly_value_col: 'LY_VALUE'})
            
            cy_value = cy_company_df.groupby([cy_company_group_col, cy_product_group_col])[cy_value_col].sum().reset_index()
            cy_value = cy_value.rename(columns={cy_product_group_col: 'PRODUCT GROUP', cy_value_col: 'CY_VALUE'})
            
            budget_value = budget_company_df.groupby([budget_company_group_col, budget_product_group_col])[budget_value_col].sum().reset_index()
            budget_value = budget_value.rename(columns={budget_product_group_col: 'PRODUCT GROUP', budget_value_col: 'BUDGET_VALUE'})
            
            # Create base DataFrames
            product_qty_df = pd.DataFrame({'PRODUCT GROUP': company_product_groups})
            product_value_df = pd.DataFrame({'PRODUCT GROUP': company_product_groups})
            
            # Merge data
            qty_df = product_qty_df.merge(ly_qty[['PRODUCT GROUP', 'LY_QTY']], on='PRODUCT GROUP', how='left')\
                                   .merge(budget_qty[['PRODUCT GROUP', 'BUDGET_QTY']], on='PRODUCT GROUP', how='left')\
                                   .merge(cy_qty[['PRODUCT GROUP', 'CY_QTY']], on='PRODUCT GROUP', how='left').fillna(0)
            
            value_df = product_value_df.merge(ly_value[['PRODUCT GROUP', 'LY_VALUE']], on='PRODUCT GROUP', how='left')\
                                       .merge(budget_value[['PRODUCT GROUP', 'BUDGET_VALUE']], on='PRODUCT GROUP', how='left')\
                                       .merge(cy_value[['PRODUCT GROUP', 'CY_VALUE']], on='PRODUCT GROUP', how='left').fillna(0)
            
            # Calculate achievement percentages for individual products
            qty_df['ACHIEVEMENT %'] = qty_df.apply(lambda row: calc_achievement(row, 'CY_QTY', 'LY_QTY'), axis=1)
            value_df['ACHIEVEMENT %'] = value_df.apply(lambda row: calc_achievement(row, 'CY_VALUE', 'LY_VALUE'), axis=1)
            
            # Round all numeric columns to 2 decimal places
            numeric_cols_qty = ['LY_QTY', 'BUDGET_QTY', 'CY_QTY']
            numeric_cols_value = ['LY_VALUE', 'BUDGET_VALUE', 'CY_VALUE']
            
            for col in numeric_cols_qty:
                qty_df[col] = qty_df[col].round(2)
            
            for col in numeric_cols_value:
                value_df[col] = value_df[col].round(2)
            
            # Reorder columns
            qty_df = qty_df[['PRODUCT GROUP', 'LY_QTY', 'BUDGET_QTY', 'CY_QTY', 'ACHIEVEMENT %']]
            value_df = value_df[['PRODUCT GROUP', 'LY_VALUE', 'BUDGET_VALUE', 'CY_VALUE', 'ACHIEVEMENT %']]
            
            # Calculate totals correctly
            total_ly_qty = qty_df['LY_QTY'].sum()
            total_cy_qty = qty_df['CY_QTY'].sum()
            total_budget_qty = qty_df['BUDGET_QTY'].sum()
            
            total_ly_value = value_df['LY_VALUE'].sum()
            total_cy_value = value_df['CY_VALUE'].sum()
            total_budget_value = value_df['BUDGET_VALUE'].sum()
            
            # Add total rows with correct growth calculation
            qty_totals = pd.DataFrame({
                'PRODUCT GROUP': ['TOTAL'],
                'LY_QTY': [round(total_ly_qty, 2)],
                'BUDGET_QTY': [round(total_budget_qty, 2)],
                'CY_QTY': [round(total_cy_qty, 2)],
                'ACHIEVEMENT %': [calc_total_growth_percentage(total_cy_qty, total_ly_qty)]
            })
            qty_df = pd.concat([qty_df, qty_totals], ignore_index=True)
            
            value_totals = pd.DataFrame({
                'PRODUCT GROUP': ['TOTAL'],
                'LY_VALUE': [round(total_ly_value, 2)],
                'BUDGET_VALUE': [round(total_budget_value, 2)],
                'CY_VALUE': [round(total_cy_value, 2)],
                'ACHIEVEMENT %': [calc_total_growth_percentage(total_cy_value, total_ly_value)]
            })
            value_df = pd.concat([value_df, value_totals], ignore_index=True)
            
            # Store result
            result[company] = {'qty_df': qty_df, 'value_df': value_df}
        
        if not result:
            logger.error("No data available after filtering.")
            return {"success": False, "error": "No data available after filtering. Please review filters and data."}
        
        logger.info("Product growth calculation completed successfully")
        
        # Convert DataFrames to records for JSON serialization
        for company in result:
            result[company]['qty_df'] = result[company]['qty_df'].to_dict('records')
            result[company]['value_df'] = result[company]['value_df'].to_dict('records')
        
        return {
            'success': True,
            'streamlit_result': result,
            'ly_month': ly_month,
            'cy_month': cy_month
        }
        
    except Exception as e:
        logger.error(f"Error in calculate_product_growth: {str(e)}")
        import traceback
        traceback.print_exc()
        return {'success': False, 'error': f"Error calculating product growth: {str(e)}"}

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
