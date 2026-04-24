import pandas as pd
import os
import glob

root_dir = r'c:\Users\-\OneDrive\바탕 화면\PMI Report'
files = glob.glob(os.path.join(root_dir, "**/*.xls*"), recursive=True)

output_file = r'c:\Users\-\OneDrive\바탕 화면\PMI Report\scratch\date_map.txt'
with open(output_file, 'w', encoding='utf-8') as f_out:
    print(f"Mapping date ranges for {len(files)} files...")
    for f in files:
        try:
            xl = pd.ExcelFile(f)
            for sheet in xl.sheet_names:
                try:
                    df = xl.parse(sheet)
                    date_cols = [col for col in df.columns if any(x in str(col).lower() for x in ['date', '날짜', 'entry', 'time'])]
                    for col in date_cols:
                        dates = pd.to_datetime(df[col], errors='coerce').dropna()
                        if not dates.empty:
                            line = f"FILE: {os.path.basename(f)} | SHEET: {sheet} | COL: {col} | RANGE: {dates.min()} TO {dates.max()}\n"
                            f_out.write(line)
                except: pass
        except: pass
print(f"Done. Results saved to {output_file}")
