import pandas as pd
from io import BytesIO
import logging

# Set up logger for Flask
logger = logging.getLogger(__name__)

def create_proof_of_calculation_excel(budget_df, sales_df, selected_month, 
                                      budget_exec_col, budget_exec_code_col, budget_area_col, 
                                      budget_sl_code_col, budget_product_group_col, 
                                      budget_qty_col, budget_value_col,
                                      sales_exec_col, sales_date_col, sales_area_col,
                                      sales_sl_code_col, sales_product_group_col,
                                      sales_qty_col, sales_value_col,
                                      selected_executives, selected_branches=None):
    """
    Creates a detailed Excel file showing the proof of calculation for budget vs billed mapping.
    
    Returns:
    - BytesIO buffer containing the Excel file
    """
    try:
        # Filter sales data for the selected month
        sales_df = sales_df.copy()
        budget_df = budget_df.copy()
        
        # Convert sales date and filter by month
        sales_df[sales_date_col] = pd.to_datetime(sales_df[sales_date_col], dayfirst=True, errors='coerce')
        filtered_sales_df = sales_df[sales_df[sales_date_col].dt.strftime('%b %y') == selected_month].copy()
        
        # Apply branch filter if provided
        if selected_branches:
            selected_branches_upper = [str(b).strip().upper() for b in selected_branches]
            filtered_sales_df = filtered_sales_df[
                filtered_sales_df[sales_area_col].astype(str).str.strip().str.upper().isin(selected_branches_upper)
            ]
            budget_df = budget_df[
                budget_df[budget_area_col].astype(str).str.strip().str.upper().isin(selected_branches_upper)
            ]
        
        # Filter by selected executives
        selected_executives_upper = [str(exec).strip().upper() for exec in selected_executives]
        filtered_sales_df = filtered_sales_df[
            filtered_sales_df[sales_exec_col].astype(str).str.strip().str.upper().isin(selected_executives_upper)
        ]
        budget_df = budget_df[
            budget_df[budget_exec_col].astype(str).str.strip().str.upper().isin(selected_executives_upper)
        ]
        
        # Standardize data for mapping
        budget_df['exec_std'] = budget_df[budget_exec_col].astype(str).str.strip().str.upper()
        budget_df['sl_code_std'] = budget_df[budget_sl_code_col].astype(str).str.strip()
        budget_df['product_std'] = budget_df[budget_product_group_col].astype(str).str.strip().str.upper()
        
        filtered_sales_df['exec_std'] = filtered_sales_df[sales_exec_col].astype(str).str.strip().str.upper()
        filtered_sales_df['sl_code_std'] = filtered_sales_df[sales_sl_code_col].astype(str).str.strip()
        filtered_sales_df['product_std'] = filtered_sales_df[sales_product_group_col].astype(str).str.strip().str.upper()
        
        # Convert numeric columns
        budget_df[budget_qty_col] = pd.to_numeric(budget_df[budget_qty_col], errors='coerce').fillna(0)
        budget_df[budget_value_col] = pd.to_numeric(budget_df[budget_value_col], errors='coerce').fillna(0)
        filtered_sales_df[sales_qty_col] = pd.to_numeric(filtered_sales_df[sales_qty_col], errors='coerce').fillna(0)
        filtered_sales_df[sales_value_col] = pd.to_numeric(filtered_sales_df[sales_value_col], errors='coerce').fillna(0)
        
        # Create detailed mapping
        detailed_records = []
        
        # Group budget data by executive, sl_code, and product_group
        budget_grouped = budget_df.groupby(['exec_std', 'sl_code_std', 'product_std']).agg({
            budget_exec_col: 'first',
            budget_exec_code_col: 'first',
            budget_area_col: 'first',
            budget_sl_code_col: 'first',
            budget_product_group_col: 'first',
            budget_qty_col: 'sum',
            budget_value_col: 'sum'
        }).reset_index()
        
        for _, budget_row in budget_grouped.iterrows():
            exec_name = budget_row['exec_std']
            sl_code = budget_row['sl_code_std']
            product = budget_row['product_std']
            budget_qty = budget_row[budget_qty_col]
            budget_value = budget_row[budget_value_col]
            
            # Check if budget conditions are met (qty > 0 AND value > 0)
            budget_conditions_met = budget_qty > 0 and budget_value > 0
            
            if budget_conditions_met:
                # Find matching sales records
                matching_sales = filtered_sales_df[
                    (filtered_sales_df['exec_std'] == exec_name) &
                    (filtered_sales_df['sl_code_std'] == sl_code) &
                    (filtered_sales_df['product_std'] == product) &
                    filtered_sales_df[sales_qty_col].notna() &
                    filtered_sales_df[sales_value_col].notna()
                ]
                
                # Calculate total sales for this combination
                sales_qty_total = matching_sales[sales_qty_col].sum() if not matching_sales.empty else 0
                sales_value_total = matching_sales[sales_value_col].sum() if not matching_sales.empty else 0
                
                # Apply min logic (same as calculate_budget_values)
                final_sales_qty = min(budget_qty, sales_qty_total) if sales_qty_total > 0 else 0
                final_sales_value = min(budget_value, sales_value_total) if sales_value_total > 0 else 0
                
                # Determine match status
                match_status = 'Mapped' if sales_qty_total > 0 or sales_value_total > 0 else 'No Sales Data'
            else:
                # Budget conditions not met, set sales to 0
                final_sales_qty = 0
                final_sales_value = 0
                match_status = 'Budget Invalid'
            
            # Create record for detailed view
            detailed_record = {
                'Executive Name': budget_row[budget_exec_col],
                'Executive Code': budget_row[budget_exec_code_col],
                'Branch': budget_row[budget_area_col],
                'SL Code': budget_row[budget_sl_code_col],
                'Product Group': budget_row[budget_product_group_col],
                'Budget Qty': round(budget_qty, 2),
                'Budget Value': round(budget_value, 2),
                'Sales Qty': round(final_sales_qty, 2),
                'Sales Value': round(final_sales_value, 2),
                'Match Status': match_status
            }
            
            detailed_records.append(detailed_record)
        
        # Create DataFrame
        proof_df = pd.DataFrame(detailed_records)
        
        # Sort by Executive Name, then by SL Code
        proof_df = proof_df.sort_values(['Executive Name', 'SL Code'])
        
        # Create Excel file
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            # Write the detailed data
            proof_df.to_excel(writer, sheet_name='Proof of Calculation', index=False)
            
            # Get the workbook and worksheet to apply formatting
            workbook = writer.book
            worksheet = writer.sheets['Proof of Calculation']
            
            # Apply formatting
            from openpyxl.styles import Font, PatternFill, Alignment
            
            # Header formatting
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
            
            for cell in worksheet[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center')
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        excel_buffer.seek(0)
        return excel_buffer
        
    except Exception as e:
        logger.error(f"Error creating proof of calculation Excel: {e}")
        # Instead of st.error (Streamlit), you can raise an exception or return None
        # For Flask, you might want to handle this differently based on your error handling strategy
        raise Exception(f"Error creating proof of calculation Excel: {e}")
        # return None