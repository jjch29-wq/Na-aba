import pandas as pd
import json
import os
import glob

root_dir = r'c:\Users\-\OneDrive\바탕 화면\PMI Report'
files = glob.glob(os.path.join(root_dir, "**/*.xls*"), recursive=True) + glob.glob(os.path.join(root_dir, "**/*.json"), recursive=True)

print(f"Brute force text search for '4월' in {len(files)} files...")

for f in files:
    try:
        if f.endswith('.json'):
            with open(f, 'r', encoding='utf-8-sig') as jf:
                content = jf.read()
                if '4월' in content:
                    print(f"HIT in JSON: {os.path.basename(f)}")
        else:
            xl = pd.ExcelFile(f)
            for sheet in xl.sheet_names:
                df = xl.parse(sheet)
                # Check all columns for the word '4월' or ' 4.'
                mask = df.astype(str).apply(lambda x: x.str.contains('4월', na=False) | x.str.contains(' 4\.', na=False) | x.str.contains('\.4\.', na=False)).any(axis=1)
                if mask.any():
                    print(f"FOUND '4월' or '.4.'! File: {os.path.basename(f)}, Sheet: {sheet}")
    except: pass
