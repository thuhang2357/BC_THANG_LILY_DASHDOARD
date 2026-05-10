import pandas as pd
import os
import glob
import json

# Get files properly
files = [f for f in os.listdir('.') if 'xlsx' in f and 'THÁNG' in f]
print("Files found:", len(files))

output = []

for f in files:
    output.append(f"=== {f} ===")
    try:
        xls = pd.ExcelFile(f)
        sheet_names = xls.sheet_names
        output.append(f"Sheets: {sheet_names}")

        # Check if 'Báo cáo cửa hàng' exists
        sheet_name = 'Báo cáo cửa hàng'
        if sheet_name in sheet_names:
            df = pd.read_excel(f, sheet_name=sheet_name, header=None)
            output.append(f"Shape: {df.shape}")
            # Find rows with store/channel identifiers
            for i in range(min(80, len(df))):
                row_data = []
                for j in range(min(5, df.shape[1])):
                    val = df.iloc[i, j]
                    if pd.notna(val):
                        row_data.append(str(val)[:30])
                if row_data:
                    output.append(f"Row {i}: {' | '.join(row_data)}")
        output.append("")
    except Exception as e:
        output.append(f"Error: {e}")
        output.append("")

# Write to file
with open('check_output.txt', 'w', encoding='utf-8') as out:
    out.write('\n'.join(output))

print("Output written to check_output.txt")
