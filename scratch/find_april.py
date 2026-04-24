import pandas as pd
import os
import glob

data_dir = r'c:\Users\-\OneDrive\바탕 화면\PMI Report\home\data'
files = glob.glob(os.path.join(data_dir, "*.xls*"))

print(f"Searching in {len(files)} files for April 2025 records...")

for f in files:
    try:
        xl = pd.ExcelFile(f)
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            if df.empty: continue
            
            # Find date-like columns
            date_cols = [col for col in df.columns if any(x in str(col).lower() for x in ['date', '날짜', 'entry', 'time'])]
            for col in date_cols:
                try:
                    dates = pd.to_datetime(df[col], errors='coerce')
                    hits = df[(dates >= '2025-04-01') & (dates <= '2025-04-30')]
                    if len(hits) > 0:
                        print(f"FOUND! File: {os.path.basename(f)}, Sheet: {sheet}, Column: {col} -> {len(hits)} records in April 2025")
                except: pass
    except Exception as e:
        # print(f"Error reading {f}: {e}")
        pass
