# utils/data_processor.py
import pandas as pd
import numpy as np
import re
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

class DataProcessor:
    def __init__(self):
        pass
    
    def handle_duplicate_columns(self, df):
        """Handle duplicate column names by adding suffixes - but be smart about Gr/Ach columns"""
        cols = pd.Series(df.columns)
        
        # Don't treat properly renamed Gr-/Ach- columns as duplicates
        for dup in cols[cols.duplicated()].unique():
            # Skip if this is a properly formatted Gr/Ach column
            if dup.startswith(('Gr-', 'Ach-', 'Act-', 'Budget-', 'LY-')):
                continue
                
            # Only rename actual duplicates
            duplicate_indices = cols[cols == dup].index.tolist()
            
            for i, idx in enumerate(duplicate_indices):
                if i != 0:  # Keep first occurrence as-is, rename subsequent ones
                    cols.iloc[idx] = f"{dup}_{i}"
        
        df.columns = cols
        return df

    def extract_tables(self, df, possible_headers, is_product_analysis=False):
        """Extract tables from Excel sheet based on header patterns"""
        for i in range(len(df)):
            row_text = ' '.join(df.iloc[i].astype(str).str.lower().tolist())
            for header in possible_headers:
                if header.lower() in row_text:
                    # Check for budget/actual style headers
                    potential_header = df.iloc[i]
                    if any(str(col).strip().lower().startswith(('budget-', 'act-', 'ly-', 'gr.', 'ach.')) for col in potential_header[1:]):
                        data_start = i + 2 if i + 2 < len(df) else i + 1
                        if data_start < len(df):
                            first_col = str(df.iloc[data_start, 0]).strip().upper()
                            identifier_cols = ['REGIONS', 'REGION', 'BRANCH', 'ORGANIZATION', 'ORGANIZATION NAME'] if not is_product_analysis else ['PRODUCT', 'PRODUCT GROUP', 'PRODUCT NAME', 'ACETIC ACID', 'AUXILARIES', 'CSF', 'TOTAL']
                            if any(r in first_col for r in identifier_cols) or first_col in ['ACCLLP', 'TOTAL SALES']:
                                return i, data_start
                        
                        data_start = i + 1
                        if data_start < len(df):
                            first_col = str(df.iloc[data_start, 0]).strip().upper()
                            if any(r in first_col for r in identifier_cols) or first_col in ['ACCLLP', 'TOTAL SALES']:
                                return i, data_start
                    
                    # Try alternative approach
                    for j in range(i + 1, min(i + 5, len(df))):
                        if j >= len(df):
                            break
                        row = df.iloc[j]
                        first_col = str(row.iloc[0]).strip().upper()
                        identifier_cols = ['REGIONS', 'REGION', 'BRANCH'] if not is_product_analysis else ['PRODUCT', 'PRODUCT GROUP', 'PRODUCT NAME', 'ACETIC ACID', 'AUXILARIES', 'CSF', 'TOTAL']
                        if any(r in first_col for r in identifier_cols):
                            header_row = j - 1 if j > 0 else j
                            return header_row, j
        
        return None, None

    def rename_columns(self, columns):
        """Standardize column names with consistent formatting"""
        renamed = []
        
        for i, col in enumerate(columns):
            col_clean = str(col).strip()
            
            # Handle identifier columns (first column)
            if col_clean.upper() in ['REGIONS', 'REGION', 'BRANCH', 'ORGANIZATION', 'ORGANIZATION NAME', 
                                   'PRODUCT', 'PRODUCT GROUP', 'PRODUCT NAME', 'ACETIC ACID', 'AUXILARIES', 'CSF']:
                renamed.append(col_clean.upper())
                continue
            
            # Keep simple columns as-is
            if col_clean.upper() in ['MT', 'RS', 'TOTAL']:
                renamed.append(col_clean.upper())
                continue
            
            # Handle Budget columns - multiple patterns
            budget_patterns = [
                r'Budget[- ]*(?:Qty|Value)?[- ]*(\w{3,})[- \']*(\d{2,4})',
                r'(\w{3,})[- \']*(\d{2,4})[- ]*Budget',
                r'Budget[- ]*(\w{3,})[- \']*(\d{2,4})'
            ]
            
            budget_found = False
            for pattern in budget_patterns:
                budget_match = re.match(pattern, col_clean, re.IGNORECASE)
                if budget_match:
                    month, year = budget_match.groups()
                    month = month[:3].capitalize()
                    year = year[-2:] if len(year) > 2 else year
                    renamed.append(f"Budget-{month}-{year}")
                    budget_found = True
                    break
            
            if budget_found:
                continue
            
            # Handle Act/Actual columns - multiple patterns
            act_patterns = [
                r'Act[ual]*[- ]*(\w{3,})[- \']*(\d{2,4})',
                r'(\w{3,})[- \']*(\d{2,4})[- ]*Act[ual]*',
                r'Actual[- ]*(\w{3,})[- \']*(\d{2,4})'
            ]
            
            act_found = False
            for pattern in act_patterns:
                act_match = re.match(pattern, col_clean, re.IGNORECASE)
                if act_match:
                    month, year = act_match.groups()
                    month = month[:3].capitalize()
                    year = year[-2:] if len(year) > 2 else year
                    renamed.append(f"Act-{month}-{year}")
                    act_found = True
                    break
            
            if act_found:
                continue
            
            # Handle YTD columns
            ytd_match = re.match(r'(?:Act[ual]*[- ]*)?YTD[-–\s]*(\d{2})[-–\s]*(\d{2})\s*\((.*?)\)(?:\s*(Act\.|L,Y|LY))?', col_clean, re.IGNORECASE)
            if ytd_match:
                start_year, end_year, period, suffix = ytd_match.groups()
                period = period.replace('June', 'Jun').replace('July', 'Jul').replace('August', 'Aug').replace('September', 'Sep').replace('October', 'Oct').replace('November', 'Nov').replace('December', 'Dec').strip()
                ytd_base = f"YTD-{start_year}-{end_year} ({period})"
                
                if suffix:
                    suffix = suffix.strip().lower()
                    if suffix == 'act.' or 'act' in col_clean.lower():
                        renamed.append(f"Act-{ytd_base}")
                    elif suffix in ['l,y', 'ly']:
                        renamed.append(f"LY-{ytd_base}")
                    else:
                        renamed.append(ytd_base)
                else:
                    # If no suffix but contains 'act' in original column, treat as Act
                    if 'act' in col_clean.lower():
                        renamed.append(f"Act-{ytd_base}")
                    else:
                        renamed.append(ytd_base)
                continue
            
            # Handle Gr/Growth columns - Find nearest preceding Act column
            if col_clean.strip().upper() in ["GR.", "GR", "GROWTH"]:
                # Look backward to find the most recent Act column
                nearest_act_col = None
                
                for j in range(i-1, -1, -1):  # Search backwards from current position
                    if j < len(renamed):
                        if renamed[j].startswith('Act-'):
                            nearest_act_col = renamed[j]
                            break
                
                if nearest_act_col:
                    new_gr_name = nearest_act_col.replace("Act-", "Gr-")
                    renamed.append(new_gr_name)
                else:
                    # Fallback: try to extract month-year from context
                    month_year = self._extract_month_year_from_context(columns, i)
                    if month_year:
                        renamed.append(f"Gr-{month_year}")
                    else:
                        renamed.append("Gr-Unknown")
                continue
            
            elif re.match(r'Gr[owth]*[- ]*(\w{3,})[- \']*(\d{2,4})', col_clean, re.IGNORECASE):
                gr_match = re.match(r'Gr[owth]*[- ]*(\w{3,})[- \']*(\d{2,4})', col_clean, re.IGNORECASE)
                month, year = gr_match.groups()
                month = month[:3].capitalize()
                year = year[-2:] if len(year) > 2 else year
                renamed.append(f"Gr-{month}-{year}")
                continue
            
            # Handle Ach/Achievement columns - Find nearest preceding Act column
            if col_clean.strip().upper() in ["ACH.", "ACH", "ACHIEVEMENT"]:
                # Look backward to find the most recent Act column
                nearest_act_col = None
                
                for j in range(i-1, -1, -1):  # Search backwards from current position
                    if j < len(renamed):
                        if renamed[j].startswith('Act-'):
                            nearest_act_col = renamed[j]
                            break
                
                if nearest_act_col:
                    new_ach_name = nearest_act_col.replace("Act-", "Ach-")
                    renamed.append(new_ach_name)
                else:
                    # Fallback: try to extract month-year from context
                    month_year = self._extract_month_year_from_context(columns, i)
                    if month_year:
                        renamed.append(f"Ach-{month_year}")
                    else:
                        renamed.append("Ach-Unknown")
                continue
            
            elif re.match(r'Ach[ievement]*[- ]*(\w{3,})[- \']*(\d{2,4})', col_clean, re.IGNORECASE):
                ach_match = re.match(r'Ach[ievement]*[- ]*(\w{3,})[- \']*(\d{2,4})', col_clean, re.IGNORECASE)
                month, year = ach_match.groups()
                month = month[:3].capitalize()
                year = year[-2:] if len(year) > 2 else year
                renamed.append(f"Ach-{month}-{year}")
                continue
            
            # Handle LY/Last Year columns
            ly_patterns = [
                r'LY[- ]*(\w{3,})[- \']*(\d{2,4})',
                r'(\w{3,})[- \']*(\d{2,4})[- ]*LY',
                r'Last[- ]*Year[- ]*(\w{3,})[- \']*(\d{2,4})'
            ]
            
            ly_found = False
            for pattern in ly_patterns:
                ly_match = re.match(pattern, col_clean, re.IGNORECASE)
                if ly_match:
                    month, year = ly_match.groups()
                    month = month[:3].capitalize()
                    year = year[-2:] if len(year) > 2 else year
                    renamed.append(f"LY-{month}-{year}")
                    ly_found = True
                    break
            
            if ly_found:
                continue
            
            # Handle month-year patterns without prefix
            month_year_match = re.match(r'(\w{3,})[- \']*(\d{2,4})', col_clean, re.IGNORECASE)
            if month_year_match:
                month, year = month_year_match.groups()
                # Check if it's a valid month
                valid_months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 
                              'jul', 'aug', 'sep', 'oct', 'nov', 'dec',
                              'january', 'february', 'march', 'april', 'may', 'june',
                              'july', 'august', 'september', 'october', 'november', 'december']
                if month.lower() in valid_months:
                    month = month[:3].capitalize()
                    year = year[-2:] if len(year) > 2 else year
                    renamed.append(f"Act-{month}-{year}")  # Default to Act if no prefix
                    continue
            
            # Default case - keep original
            renamed.append(col_clean)
        
        return renamed
    
    def _extract_month_year_from_context(self, columns, current_index):
        """Helper method to extract month-year from surrounding columns"""
        # Look at nearby columns to infer month-year
        for offset in [-1, -2, 1, 2]:  # Check columns before and after
            check_index = current_index + offset
            if 0 <= check_index < len(columns):
                col = str(columns[check_index]).strip()
                
                # Try to extract month-year pattern
                month_year_match = re.match(r'.*?(\w{3,})[- \']*(\d{2,4})', col, re.IGNORECASE)
                if month_year_match:
                    month, year = month_year_match.groups()
                    valid_months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 
                                  'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
                    if month.lower()[:3] in valid_months:
                        month = month[:3].capitalize()
                        year = year[-2:] if len(year) > 2 else year
                        return f"{month}-{year}"
        
        return None

    def clean_and_convert_numeric(self, df):
        """Clean and convert DataFrame columns to appropriate data types"""
        if df is None or df.empty:
            return df
            
        df = df.copy()
        
        # Get the first column (identifier column)
        identifier_col = df.columns[0]
        
        # Process each column
        for col in df.columns:
            if col == identifier_col:
                # Keep identifier column as string
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace({'nan': '', 'NaN': '', 'None': '', 'null': ''})
                continue
                
            # Try to convert to numeric for other columns
            try:
                # First, convert to string and clean
                series_str = df[col].astype(str).str.strip()
                
                # Replace common non-numeric values
                series_str = series_str.replace({
                    'nan': '0', 'NaN': '0', 'None': '0', '': '0', '-': '0', 'null': '0'
                })
                
                # Try to convert to numeric
                df[col] = pd.to_numeric(series_str, errors='coerce')
                
                # Fill any remaining NaN values with 0
                df[col] = df[col].fillna(0.0)
                
                # Ensure it's float64 for consistency
                df[col] = df[col].astype('float64')
                
            except Exception:
                # If conversion fails, keep as string but clean it
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace({'nan': '', 'NaN': '', 'None': '', 'null': ''})
        
        return df