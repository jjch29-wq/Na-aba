import pandas as pd
import json
import os
import glob

root_dir = r'c:\Users\-\OneDrive\바탕 화면\PMI Report'
files = glob.glob(os.path.join(root_dir, "**/*.xls*"), recursive=True) + glob.glob(os.path.join(root_dir, "**/*.json"), recursive=True)

print(f"Deep searching in {len(files)} files for '2025' and '04'...")

for f in files:
    try:
        if f.endswith('.json'):
            with open(f, 'r', encoding='utf-8') as jf:
                data = json.load(jf)
                content = str(data)
                if '2025-04' in content or '2025.04' in content or '2025. 4' in content:
                    print(f"POTENTIAL HIT in JSON: {os.path.basename(f)}")
        else:
            xl = pd.ExcelFile(f)
            for sheet in xl.sheet_names:
                df = xl.parse(sheet)
                if df.empty: continue
                # Check for 2025 in any column
                mask = df.astype(str).apply(lambda x: x.str.contains('2025', na=False)).any(axis=1)
                if mask.any():
                    # If 2025 exists, check if April (04) also exists in the same file
                    april_mask = df.astype(str).apply(lambda x: x.str.contains('04', na=False) | x.str.contains('\. 4', na=False)).any(axis=1)
                    if (mask & april_mask).any():
                        print(f"FOUND! File: {os.path.basename(f)}, Sheet: {sheet}")
    except: pass
