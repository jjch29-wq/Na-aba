import pandas as pd
import os

target_file = r'c:\Users\-\OneDrive\바탕 화면\PMI Report\PMI_Export_20260325_202047.xlsx'

print(f"Inspecting potential backup: {target_file}")
try:
    xl = pd.ExcelFile(target_file)
    print(f"Sheets: {xl.sheet_names}")
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        print(f"Sheet '{sheet}': {len(df)} rows")
        # Check for April 2025
        date_cols = [col for col in df.columns if any(x in str(col).lower() for x in ['date', '날짜', 'entry', 'time'])]
        for col in date_cols:
            dates = pd.to_datetime(df[col], errors='coerce')
            hits = df[(dates >= '2025-04-01') & (dates <= '2025-04-30')]
            if len(hits) > 0:
                print(f"  --> FOUND {len(hits)} records for April 2025 in column '{col}'!")
                # Show sample
                print(hits[['Date', 'Site', 'MaterialID', 'Usage']].head(5))
except Exception as e:
    print(f"Error: {e}")
