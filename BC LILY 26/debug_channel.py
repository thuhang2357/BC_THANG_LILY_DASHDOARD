import pandas as pd
import os

files = [f for f in os.listdir('.') if 'xlsx' in f and 'THÁNG' in f]
output = []

for f in files:
    output.append(f'=== {f} ===')
    try:
        df = pd.read_excel(f, sheet_name='Báo cáo cửa hàng', header=None)

        off_revenue = 0
        tmdt_revenue = 0
        online_revenue = 0

        current_section = None

        for i in range(len(df)):
            row = df.iloc[i]
            first_col = str(row[0]) if pd.notna(row[0]) else ''

            if 'STORE_MB' in first_col:
                current_section = 'off'
                output.append(f"Row {i}: OFF section (MB)")
            elif 'STORE_MN' in first_col:
                current_section = 'off'
                output.append(f"Row {i}: OFF section (MN)")
            elif 'STORE_MT' in first_col:
                current_section = 'off'
                output.append(f"Row {i}: OFF section (MT)")
            elif 'TMĐT' in first_col and len(str(row[1] if len(row) > 1 else '')) > 5:
                current_section = 'tmdt'
                output.append(f"Row {i}: TMĐT section")
            elif any(x in first_col for x in ['VP', 'SEEDING', 'FB', 'IG', 'WEB', 'Đơn sỉ']) or 'SHOPEE' in first_col or 'TIKTOK' in first_col:
                if current_section == 'tmdt':
                    current_section = 'online'
                    output.append(f"Row {i}: ONLINE section - {first_col[:20]}")

            if len(row) > 2 and current_section:
                rev = row[2]
                if pd.notna(rev) and isinstance(rev, (int, float)) and rev > 0:
                    if current_section == 'off':
                        off_revenue += rev
                    elif current_section == 'tmdt':
                        tmdt_revenue += rev
                    elif current_section == 'online':
                        online_revenue += rev

        output.append(f"OFF: {off_revenue:,}")
        output.append(f"TMĐT: {tmdt_revenue:,}")
        output.append(f"ONLINE: {online_revenue:,}")
        output.append("")

    except Exception as e:
        output.append(f"Error: {e}")
        output.append("")

with open('debug_output.txt', 'w', encoding='utf-8') as out:
    out.write('\n'.join(output))

print("Done. Check debug_output.txt")
