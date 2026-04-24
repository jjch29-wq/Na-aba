import pandas as pd
import os

db_path = r'c:\Users\-\OneDrive\바탕 화면\PMI Report\home\data\Material_Inventory.xlsx'

print(f"Checking DB: {db_path}")
if not os.path.exists(db_path):
    print("File not found!")
else:
    try:
        xl = pd.ExcelFile(db_path)
        print(f"Sheet names: {xl.sheet_names}")
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            print(f"Sheet '{sheet}': {len(df)} rows")
            if sheet == 'DailyUsage' and len(df) > 0:
                if 'Date' in df.columns:
                    print(f"  - Date range in DailyUsage: {df['Date'].min()} to {df['Date'].max()}")
                else:
                    # Look for Korean '날짜'
                    for col in df.columns:
                        if '날짜' in str(col):
                            print(f"  - Date range in DailyUsage ({col}): {df[col].min()} to {df[col].max()}")
    except Exception as e:
        print(f"Error reading excel: {e}")
