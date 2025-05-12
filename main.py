from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import os

app = FastAPI(title="Excel Sheet API", docs_url="/docs")

# CORS setup
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

EXCEL_PATH = "C:\\Users\\KetakiDesale\\Downloads\\Project\\capbudg (3).xlsx"  # Make sure this file exists

def load_excel():
    try:
        if EXCEL_PATH.endswith(".xls"):
            xls = pd.ExcelFile(EXCEL_PATH, engine="xlrd")
        elif EXCEL_PATH.endswith(".xlsx"):
            xls = pd.ExcelFile(EXCEL_PATH, engine="openpyxl")
        else:
            raise HTTPException(status_code=400, detail="Unsupported file format")
        return xls
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error reading Excel: {str(e)}")

@app.get("/list_tables")
def list_tables():
    xls = load_excel()
    return {"tables": xls.sheet_names}

@app.get("/get_table_details")
def get_table_details(table_name: str = Query(...)):
    xls = load_excel()
    if table_name not in xls.sheet_names:
        raise HTTPException(status_code=404, detail="Table not found")
    
    df = xls.parse(sheet_name=table_name)
    if df.empty or df.shape[1] == 0:
        raise HTTPException(status_code=400, detail="Table is empty or malformed")

    row_names = df.iloc[:, 0].dropna().astype(str).tolist()
    return {"table_name": table_name, "row_names": row_names}

@app.get("/row_sum")
def row_sum(table_name: str = Query(...), row_name: str = Query(...)):
    xls = load_excel()
    if table_name not in xls.sheet_names:
        raise HTTPException(status_code=404, detail="Table not found")
    
    df = xls.parse(sheet_name=table_name)
    if df.empty or df.shape[1] < 2:
        raise HTTPException(status_code=400, detail="Table does not have enough columns")

    # Normalize first column to str
    df.iloc[:, 0] = df.iloc[:, 0].astype(str)
    
    # Find the row
    row = df[df.iloc[:, 0] == row_name]
    if row.empty:
        raise HTTPException(status_code=404, detail="Row not found")
    
    numeric_values = row.iloc[0, 1:]
    numeric_sum = numeric_values.apply(pd.to_numeric, errors="coerce").sum(skipna=True)
    
    return {
        "table_name": table_name,
        "row_name": row_name,
        "sum": numeric_sum
    }



























