import pandas as pd
import os

db_path = r'c:\Users\-\OneDrive\바탕 화면\PMI Report\home\data\Material_Inventory.xlsx'

print(f"Extraction for DailyUsage in: {db_path}")

try:
    df = pd.read_excel(db_path, sheet_name='DailyUsage')
    # Save to CSV for full visibility in terminal
    df.to_csv(r'c:\Users\-\OneDrive\바탕 화면\PMI Report\scratch\daily_usage_dump.csv', index=False, encoding='utf-8-sig')
    
    # Filter for April 2025 in ANY column as string
    any_2025_04 = df.astype(str).apply(lambda x: x.str.contains('2025', na=False) & (x.str.contains('04', na=False) | x.str.contains('\. 4', na=False)), axis=1)
    hits = df[any_2025_04].copy()
    
    print(f"Total Rows in Sheet: {len(df)}")
    print(f"Total Hits for April 2025 keyword: {len(hits)}")
    if len(hits) > 0:
        print("First few hits:")
        print(hits.head(10))
except Exception as e:
    print(f"Error: {e}")
