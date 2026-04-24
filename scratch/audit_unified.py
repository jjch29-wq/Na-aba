import pandas as pd
import os

unified_file = r'c:\Users\-\OneDrive\바탕 화면\PMI Report\home\data\SIT-PMI-K3-1_Unified_20260309_211135.xlsm'

print(f"Deep Audit of Unified File: {unified_file}")
try:
    xl = pd.ExcelFile(unified_file, engine='openpyxl')
    print(f"Sheets: {xl.sheet_names}")
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        print(f"Sheet '{sheet}': {len(df)} rows")
        
        # Look for 2025-04
        # Since column names might be weird, check all columns for date-like strings
        for col in df.columns:
            try:
                dates = pd.to_datetime(df[col], errors='coerce').dropna()
                hits = df[(dates >= '2025-04-01') & (dates <= '2025-04-30')]
                if len(hits) > 0:
                    print(f"  !!! BINGO !!! Found {len(hits)} records in sheet '{sheet}', column '{col}'")
                    # Save these to a recovery file
                    recovery_path = r'c:\Users\-\OneDrive\바탕 화면\PMI Report\scratch\recovered_april_2025.csv'
                    hits.to_csv(recovery_path, index=False, encoding='utf-8-sig')
                    print(f"  Recovered data saved to {recovery_path}")
            except: pass
except Exception as e:
    print(f"Error reading XLSM: {e}")
