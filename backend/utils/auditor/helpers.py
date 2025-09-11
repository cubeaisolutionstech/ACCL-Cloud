import re
from fuzzywuzzy import process

def extract_tables(df, possible_headers, is_product_analysis=False):
    for i in range(len(df)):
        row_text = ' '.join(df.iloc[i].astype(str).str.lower().tolist())
        for header in possible_headers:
            if header.lower() in row_text:
                for j in range(i + 1, min(i + 5, len(df))):
                    row = df.iloc[j]
                    first_col = str(row.iloc[0]).strip().upper()
                    if any(x in first_col for x in ['REGIONS', 'BRANCH', 'PRODUCT']):
                        return i, j
    raise ValueError(f"Could not locate table header. Tried: {', '.join(possible_headers)}")

def rename_columns(columns):
    renamed = []
    for col in columns:
        col = str(col).strip().upper()
        col = col.replace("ACTUAL", "Act").replace("BUDGET", "Budget").replace("LY", "LY")
        renamed.append(col)
    return renamed
