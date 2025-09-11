import pandas as pd
import io

def get_sheet_names(file):
    excel = pd.ExcelFile(file)
    return excel.sheet_names

def get_sheet_preview(file, sheet, header_row):
    df = pd.read_excel(file, sheet_name=sheet, header=header_row)
    return {
        "columns": list(df.columns),
        "preview": df.fillna("").astype(str).head(10).to_dict(orient="records")
    }
def read_excel_from_binary(binary_data):
    return pd.read_excel(io.BytesIO(binary_data))