import pandas as pd
import os
import glob

root_dir = r'c:\Users\-\OneDrive\바탕 화면\PMI Report'
files = glob.glob(os.path.join(root_dir, "**/*.xls*"), recursive=True)

print(f"Auditing {len(files)} files for row counts...")

results = []
for f in files:
    try:
        xl = pd.ExcelFile(f)
        for sheet in xl.sheet_names:
            try:
                df = xl.parse(sheet)
                results.append({
                    'File': os.path.basename(f),
                    'Sheet': sheet,
                    'Rows': len(df),
                    'Path': f
                })
            except: pass
    except: pass

res_df = pd.DataFrame(results)
# Sort by Rows descending to find the biggest lists
if not res_df.empty:
    res_df = res_df.sort_values(by='Rows', ascending=False)
    print("TOP 20 LARGEST SHEETS FOUND:")
    print(res_df[['File', 'Sheet', 'Rows']].head(20).to_string(index=False))
    print("\nFULL PATHS FOR TOP 5:")
    for i, row in res_df.head(5).iterrows():
        print(f"{row['Rows']} rows: {row['Path']}")
else:
    print("No data found.")
