import json
import pandas as pd

with open('data/data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Summary
summary_df = pd.DataFrame({
    'Chi tieu': ['Doanh thu', 'Doanh thu thuan', 'Loi nhuan gop', 'Loi nhuan thuan'],
    'Thang 1': [data['summary']['revenue'][0], data['summary']['netRevenue'][0], data['summary']['grossProfit'][0], data['summary']['netProfit'][0]],
    'Thang 2': [data['summary']['revenue'][1], data['summary']['netRevenue'][1], data['summary']['grossProfit'][1], data['summary']['netProfit'][1]],
    'Thang 3': [data['summary']['revenue'][2], data['summary']['netRevenue'][2], data['summary']['grossProfit'][2], data['summary']['netProfit'][2]]
})

# Expenses - 8 categories
exp_keys = [
    'nhan_su', 'thue_khau_hao', 'san_tmdt', 'quang_cao_marketing',
    'van_chuyen', 'dien_nuoc_dich_vu', 'chi_phi_van_hanh', 'khac'
]
exp_titles = [
    'Nhân sự',
    'Thuê nhà & Khấu hao',
    'Sàn TMĐT',
    'Quảng cáo & Marketing',
    'Vận chuyển',
    'Điện nước & Dịch vụ',
    'Chi phí vận hành',
    'Khác'
]
expenses_df = pd.DataFrame({
    'Loai chi phi': exp_titles,
    'Thang 1': [data['expenses'][k][0] for k in exp_keys],
    'Thang 2': [data['expenses'][k][1] for k in exp_keys],
    'Thang 3': [data['expenses'][k][2] for k in exp_keys]
})

# Channels - now with 5 channels
channels_df = pd.DataFrame({
    'Kenh': ['Mien Bac', 'Mien Nam', 'Mien Trung', ' TMDT', 'ONLINE'],
    'Doanh thu T1': [
        data['channels']['north']['netRevenue'][0],
        data['channels']['south']['netRevenue'][0],
        data['channels']['central']['netRevenue'][0],
        data['channels']['tmdt']['netRevenue'][0],
        data['channels']['online']['netRevenue'][0]
    ],
    'Doanh thu T2': [
        data['channels']['north']['netRevenue'][1],
        data['channels']['south']['netRevenue'][1],
        data['channels']['central']['netRevenue'][1],
        data['channels']['tmdt']['netRevenue'][1],
        data['channels']['online']['netRevenue'][1]
    ],
    'Doanh thu T3': [
        data['channels']['north']['netRevenue'][2],
        data['channels']['south']['netRevenue'][2],
        data['channels']['central']['netRevenue'][2],
        data['channels']['tmdt']['netRevenue'][2],
        data['channels']['online']['netRevenue'][2]
    ],
    'LN gop T3': [
        data['channels']['north']['grossProfit'][2],
        data['channels']['south']['grossProfit'][2],
        data['channels']['central']['grossProfit'][2],
        data['channels']['tmdt']['grossProfit'][2],
        data['channels']['online']['grossProfit'][2]
    ],
    'Bien LN T3': [
        data['channels']['north']['grossMargin'][2],
        data['channels']['south']['grossMargin'][2],
        data['channels']['central']['grossMargin'][2],
        data['channels']['tmdt']['grossMargin'][2],
        data['channels']['online']['grossMargin'][2]
    ],
    'Chi phi T3': [
        data['channels']['north']['totalExpense'][2],
        data['channels']['south']['totalExpense'][2],
        data['channels']['central']['totalExpense'][2],
        data['channels']['tmdt']['totalExpense'][2],
        data['channels']['online']['totalExpense'][2]
    ],
    'Ty le CP/DT T3': [
        data['channels']['north']['expenseRatio'][2],
        data['channels']['south']['expenseRatio'][2],
        data['channels']['central']['expenseRatio'][2],
        data['channels']['tmdt']['expenseRatio'][2],
        data['channels']['online']['expenseRatio'][2]
    ]
})

with pd.ExcelWriter('data_check_v2.xlsx', engine='openpyxl') as writer:
    summary_df.to_excel(writer, sheet_name='Tong hop', index=False)
    expenses_df.to_excel(writer, sheet_name='Chi phi', index=False)
    channels_df.to_excel(writer, sheet_name='Kenh', index=False)

print('Da tao data_check_v2.xlsx')
