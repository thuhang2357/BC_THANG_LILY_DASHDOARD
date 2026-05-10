# Dashboard Tài Chính - Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Tạo hệ thống dashboard tài chính tự động đọc từ Excel files và hiển thị Chart.js dashboard.

**Architecture:**
- Python ETL script đọc Excel files và xuất JSON
- HTML Dashboard với Chart.js hiển thị charts
- Filter cho phép chọn 1 hoặc nhiều tháng
- Tab navigation cho các view khác nhau

**Tech Stack:** Python 3, pandas, openpyxl, HTML5, CSS3, Chart.js 4.x (CDN)

---

## File Structure

```
BC LILY 26/
├── THÁNG 01.26 BÁO CÁO THÁNG 01.xlsx      (source data)
├── THÁNG 02.26 BÁO CÁO THÁNG 02 .xlsx    (source data)
├── THÁNG 03.26 BÁO CÁO THÁNG 03.xlsx    (source data)
├── scripts/
│   └── process_data.py                     (ETL script)
├── data/
│   └── data.json                          (central data)
├── dashboard/
│   └── dashboard.html                     (main dashboard)
└── data_nhap_tay.xlsx                     (manual input template)
```

---

## Task 1: Create Folder Structure and Initial Data Schema

**Files:**
- Create: `scripts/.gitkeep`
- Create: `data/.gitkeep`
- Create: `dashboard/.gitkeep`
- Create: `data/data.json` (initial schema)

- [ ] **Step 1: Create folders**

```bash
mkdir -p scripts data dashboard
touch scripts/.gitkeep data/.gitkeep dashboard/.gitkeep
```

- [ ] **Step 2: Create initial data.json schema**

```json
{
  "meta": {
    "company": "Công Ty Cổ Phần ULTD Thịnh Phát",
    "lastUpdated": "2026-05-08",
    "currency": "VND"
  },
  "months": ["Tháng 1", "Tháng 2", "Tháng 3"],
  "monthKeys": ["01", "02", "03"],
  "selectedMonths": [0, 1, 2],
  "summary": {
    "revenue": [0, 0, 0],
    "revenueDiscount": [0, 0, 0],
    "netRevenue": [0, 0, 0],
    "cogs": [0, 0, 0],
    "grossProfit": [0, 0, 0],
    "sellingExpense": [0, 0, 0],
    "managementExpense": [0, 0, 0],
    "netProfit": [0, 0, 0],
    "cash": [0, 0, 0]
  },
  "growth": {
    "revenue": [null, 0, 0],
    "grossProfit": [null, 0, 0],
    "netProfit": [null, 0, 0]
  },
  "expenses": {
    "ads": [0, 0, 0],
    "platform": [0, 0, 0],
    "ship": [0, 0, 0],
    "service": [0, 0, 0],
    "salary": [0, 0, 0],
    "other": [0, 0, 0]
  },
  "channels": {
    "off": {"total": [0, 0, 0]},
    "tmdt": {"total": [0, 0, 0]},
    "online": {"total": [0, 0, 0]}
  }
}
```

- [ ] **Step 3: Commit**

```bash
git add -A && git commit -m "chore: create folder structure and initial data schema"
```

---

## Task 2: Create process_data.py ETL Script

**Files:**
- Create: `scripts/process_data.py`

- [ ] **Step 1: Create ETL script with data extraction**

```python
#!/usr/bin/env python3
"""
process_data.py - ETL script to extract financial data from Excel reports
Usage: python scripts/process_data.py
"""

import pandas as pd
import json
import os
import re
from pathlib import Path

def parse_vietnamese_number(val):
    """Parse Vietnamese number format (1.234.567) to int"""
    if pd.isna(val):
        return 0
    if isinstance(val, (int, float)):
        return int(val) if not pd.isna(val) else 0
    val = str(val).strip()
    if not val:
        return 0
    val = val.replace('.', '').replace(',', '.').replace(' ', '')
    try:
        return int(float(val))
    except:
        return 0

def find_excel_files():
    """Find all monthly report Excel files"""
    pattern = re.compile(r'THÁNG\s+\d+\.26\s+BÁO\s+CÁO\s+THÁNG\s+\d+\.xlsx', re.IGNORECASE)
    files = []
    for f in Path('.').iterdir():
        if f.is_file() and pattern.match(f.name):
            files.append((f.name, f))
    return sorted(files, key=lambda x: x[0])

def extract_summary_data(df):
    """Extract summary metrics from BCKQHĐKD sheet"""
    data = {
        'revenue': 0,
        'revenueDiscount': 0,
        'netRevenue': 0,
        'cogs': 0,
        'grossProfit': 0,
        'sellingExpense': 0,
        'managementExpense': 0,
        'netProfit': 0
    }

    for idx, row in df.iterrows():
        row_values = [str(v).strip() if pd.notna(v) else '' for v in row.values]
        first_col = row_values[0].replace(',', '.').replace(' ', '')

        if first_col.startswith('1.') and 'Doanh thu' in row_values[0] and 'bán hàng' in row_values[0]:
            data['revenue'] = parse_vietnamese_number(row_values[3])
        elif first_col.startswith('2.') and 'giảm trừ' in row_values[0]:
            data['revenueDiscount'] = parse_vietnamese_number(row_values[3])
        elif first_col.startswith('3.') and 'Doanh thu thuần' in row_values[0]:
            data['netRevenue'] = parse_vietnamese_number(row_values[3])
        elif first_col.startswith('4.') and 'Giá vốn' in row_values[0]:
            data['cogs'] = parse_vietnamese_number(row_values[3])
        elif first_col.startswith('5.') and 'Lợi nhuận gộp' in row_values[0]:
            data['grossProfit'] = parse_vietnamese_number(row_values[3])
        elif first_col.startswith('8.') and 'Chi phí bán hàng' in row_values[0]:
            data['sellingExpense'] = parse_vietnamese_number(row_values[3])
        elif first_col.startswith('9.') and 'Chi phí quản lý' in row_values[0]:
            data['managementExpense'] = parse_vietnamese_number(row_values[3])
        elif first_col.startswith('17.') and 'Lợi nhuận sau thuế' in row_values[0]:
            data['netProfit'] = parse_vietnamese_number(row_values[3])

    return data

def extract_expense_data(df):
    """Extract expense breakdown from BC Chi Phí sheet"""
    expenses = {
        'ads': 0,
        'platform': 0,
        'ship': 0,
        'service': 0,
        'salary': 0,
        'other': 0
    }

    for idx, row in df.iterrows():
        if idx < 4:
            continue
        row_values = [str(v).strip() if pd.notna(v) else '' for v in row.values]
        if len(row_values) < 4:
            continue

        name = row_values[1].lower() if len(row_values) > 1 else ''
        amount = parse_vietnamese_number(row_values[3])

        if 'quảng cáo' in name or 'ads' in name or 'gg' in name or 'fb' in name or 'ig' in name or 'tiktok' in name or 'shopee' in name and 'phí' in name:
            expenses['ads'] += amount
        elif 'sàn' in name or ('shopee' in name and 'phí' not in name):
            expenses['platform'] += amount
        elif 'ship' in name or 'vận chuyển' in name or 'ghtk' in name or 'aha' in name:
            expenses['ship'] += amount
        elif 'dịch vụ' in name or 'thuê' in name or 'điện' in name or 'nước' in name:
            expenses['service'] += amount
        elif 'lương' in name or 'bhxh' in name or 'bảo hiểm' in name:
            expenses['salary'] += amount
        elif amount > 0:
            expenses['other'] += amount

    return expenses

def extract_channel_data(df):
    """Extract channel revenue from Báo cáo cửa hàng sheet"""
    channels = {
        'off': 0,
        'tmdt': 0,
        'online': 0
    }
    current_section = None

    for idx, row in df.iterrows():
        if idx < 4:
            continue
        row_values = [str(v).strip() if pd.notna(v) else '' for v in row.values]
        first_col = row_values[0] if len(row_values) > 0 else ''

        if 'STORE_MB' in first_col or 'STORE_MN' in first_col or 'STORE_MT' in first_col:
            current_section = 'off'
        elif 'TMĐT' in first_col or 'BỘ PHẬN THƯƠNG MẠI' in str(row_values[1] if len(row_values) > 1 else ''):
            current_section = 'tmdt'
        elif 'VP' in first_col or 'SEEDING' in first_col or 'FB' in first_col or 'IG' in first_col or 'WEB' in first_col or 'Đơn sỉ' in first_col:
            current_section = 'online'
        elif 'Tổng cộng' in first_col:
            current_section = None

        if current_section and len(row_values) > 2:
            rev = parse_vietnamese_number(row_values[2])
            if rev > 0 and 'STORE' not in first_col:
                channels[current_section] += rev

    return channels

def process_all_files():
    """Main ETL process"""
    files = find_excel_files()
    print(f"Found {len(files)} Excel files")

    result = {
        'meta': {
            'company': 'Công Ty Cổ Phần ULTD Thịnh Phát',
            'lastUpdated': pd.Timestamp.now().strftime('%Y-%m-%d'),
            'currency': 'VND'
        },
        'months': [],
        'monthKeys': [],
        'selectedMonths': [0, 1, 2],
        'summary': {
            'revenue': [],
            'revenueDiscount': [],
            'netRevenue': [],
            'cogs': [],
            'grossProfit': [],
            'sellingExpense': [],
            'managementExpense': [],
            'netProfit': [],
            'cash': []
        },
        'growth': {
            'revenue': [None],
            'grossProfit': [None],
            'netProfit': [None]
        },
        'expenses': {
            'ads': [],
            'platform': [],
            'ship': [],
            'service': [],
            'salary': [],
            'other': []
        },
        'channels': {
            'off': {'total': []},
            'tmdt': {'total': []},
            'online': {'total': []}
        }
    }

    for filename, filepath in files:
        print(f"Processing: {filename}")
        month_key = re.search(r'THÁNG\s+(\d+)', filename, re.IGNORECASE).group(1)
        result['months'].append(f'Tháng {month_key}')
        result['monthKeys'].append(month_key.zfill(2))

        # Read sheets
        xls = pd.ExcelFile(filepath)

        # Summary
        df_summary = pd.read_excel(filepath, sheet_name='BCKQHĐKD', header=None)
        summary = extract_summary_data(df_summary)
        for k, v in summary.items():
            result['summary'][k].append(v)

        # Expenses
        if 'BC Chi Phí' in xls.sheet_names:
            df_exp = pd.read_excel(filepath, sheet_name='BC Chi Phí', header=None)
            expenses = extract_expense_data(df_exp)
            for k, v in expenses.items():
                result['expenses'][k].append(v)

        # Channels
        if 'Báo cáo cửa hàng' in xls.sheet_names:
            df_ch = pd.read_excel(filepath, sheet_name='Báo cáo cửa hàng', header=None)
            channels = extract_channel_data(df_ch)
            result['channels']['off']['total'].append(channels['off'])
            result['channels']['tmdt']['total'].append(channels['tmdt'])
            result['channels']['online']['total'].append(channels['online'])

    # Calculate growth
    for i in range(1, len(result['summary']['revenue'])):
        prev_rev = result['summary']['revenue'][i-1]
        curr_rev = result['summary']['revenue'][i]
        if prev_rev > 0:
            result['growth']['revenue'].append(round((curr_rev - prev_rev) / prev_rev * 100, 1))
        else:
            result['growth']['revenue'].append(None)

        prev_profit = result['summary']['grossProfit'][i-1]
        curr_profit = result['summary']['grossProfit'][i]
        if prev_profit > 0:
            result['growth']['grossProfit'].append(round((curr_profit - prev_profit) / prev_profit * 100, 1))
        else:
            result['growth']['grossProfit'].append(None)

        prev_net = result['summary']['netProfit'][i-1]
        curr_net = result['summary']['netProfit'][i]
        if prev_net != 0:
            result['growth']['netProfit'].append(round((curr_net - prev_net) / abs(prev_net) * 100, 1))
        else:
            result['growth']['netProfit'].append(None)

    # Save to data.json
    with open('data/data.json', 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Data saved to data/data.json")
    print(f"Summary: {len(result['months'])} months processed")
    return result

if __name__ == '__main__':
    process_all_files()
```

- [ ] **Step 2: Run script to verify it works**

```bash
python scripts/process_data.py
```

Expected output: Creates data/data.json with extracted data

- [ ] **Step 3: Commit**

```bash
git add scripts/process_data.py data/data.json
git commit -m "feat: add ETL script for extracting financial data from Excel"
```

---

## Task 3: Create Dashboard HTML with Chart.js

**Files:**
- Create: `dashboard/dashboard.html`

- [ ] **Step 1: Create dashboard HTML with all sections**

```html
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Báo Cáo Tài Chính - ULTD Thịnh Phát</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', system-ui, sans-serif;
            background: #f5f7fa;
            color: #333;
        }

        /* Header */
        .header {
            background: linear-gradient(135deg, #1a5276 0%, #2980b9 100%);
            color: white;
            padding: 20px 30px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .header h1 { font-size: 22px; font-weight: 600; }
        .header-date { font-size: 13px; opacity: 0.9; }

        /* Filter Bar */
        .filter-bar {
            background: white;
            padding: 15px 30px;
            display: flex;
            gap: 20px;
            align-items: center;
            border-bottom: 1px solid #eaecef;
        }
        .filter-label { font-weight: 600; color: #1a5276; }
        .month-filter { display: flex; gap: 10px; }
        .month-checkbox { display: flex; align-items: center; gap: 5px; cursor: pointer; }
        .month-checkbox input { cursor: pointer; }

        /* KPI Cards */
        .kpi-cards {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            padding: 20px 30px;
            margin-top: -20px;
        }
        .card {
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        }
        .card-label { font-size: 12px; color: #666; margin-bottom: 8px; text-transform: uppercase; }
        .card-value { font-size: 26px; font-weight: 700; color: #1a5276; }
        .card-value.positive { color: #27ae60; }
        .card-value.negative { color: #e74c3c; }
        .card-change { font-size: 12px; margin-top: 5px; }
        .card-change.up { color: #27ae60; }
        .card-change.down { color: #e74c3c; }

        /* Tab Navigation */
        .tab-nav {
            background: white;
            padding: 0 30px;
            border-bottom: 1px solid #eaecef;
        }
        .tab-btn {
            padding: 15px 20px;
            border: none;
            background: none;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            color: #666;
            border-bottom: 3px solid transparent;
        }
        .tab-btn.active {
            color: #1a5276;
            border-bottom-color: #1a5276;
        }

        /* Main Content */
        .main-content { padding: 25px 30px; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }

        /* Section */
        .section {
            background: white;
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 25px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        }
        .section-title {
            font-size: 16px;
            font-weight: 600;
            color: #1a5276;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #eaecef;
        }

        /* Charts Grid */
        .charts-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 25px;
        }
        .chart-container { height: 320px; }

        /* Table */
        .metrics-table { width: 100%; border-collapse: collapse; }
        .metrics-table th, .metrics-table td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #eaecef;
        }
        .metrics-table th {
            background: #f8f9fa;
            font-weight: 600;
            color: #1a5276;
            font-size: 13px;
        }
        .metrics-table td { font-size: 14px; }
        .metrics-table tr:hover { background: #f8f9fa; }

        /* Growth Bar */
        .growth-bar { display: flex; align-items: center; margin-bottom: 12px; }
        .growth-label { width: 120px; font-weight: 500; }
        .growth-bar-fill {
            flex: 1;
            height: 24px;
            background: #eaecef;
            border-radius: 12px;
            overflow: hidden;
        }
        .growth-bar-value { width: 80px; text-align: right; padding-right: 10px; }
        .growth-bar-fill-inner {
            height: 100%;
            background: #27ae60;
            border-radius: 12px;
            display: flex;
            align-items: center;
            padding-left: 10px;
            color: white;
            font-size: 12px;
            font-weight: 600;
        }

        /* Footer */
        .footer {
            text-align: center;
            padding: 20px;
            color: #666;
            font-size: 12px;
        }

        @media (max-width: 1024px) {
            .charts-grid { grid-template-columns: 1fr; }
            .kpi-cards { grid-template-columns: repeat(2, 1fr); }
        }
    </style>
</head>
<body>
    <div class="header">
        <div>
            <h1>BÁO CÁO TÀI CHÍNH</h1>
            <div class="header-date" id="company-name">Công Ty Cổ Phần ULTD Thịnh Phát</div>
        </div>
        <div style="text-align: right;">
            <div class="header-date" id="last-updated">Cập nhật: --</div>
        </div>
    </div>

    <div class="filter-bar">
        <span class="filter-label">Lọc tháng:</span>
        <div class="month-filter" id="month-filter"></div>
    </div>

    <div class="kpi-cards">
        <div class="card">
            <div class="card-label">DOANH THU</div>
            <div class="card-value" id="kpi-revenue">--</div>
            <div class="card-change" id="kpi-revenue-change">--</div>
        </div>
        <div class="card">
            <div class="card-label">LỢI NHUẬN</div>
            <div class="card-value" id="kpi-profit">--</div>
            <div class="card-change" id="kpi-profit-change">--</div>
        </div>
        <div class="card">
            <div class="card-label">TỶ SUẤT LN</div>
            <div class="card-value" id="kpi-margin">--</div>
            <div class="card-change" id="kpi-margin-change">--</div>
        </div>
        <div class="card">
            <div class="card-label">ĐÁNH GIÁ</div>
            <div class="card-value" id="kpi-status" style="font-size: 18px;">--</div>
        </div>
    </div>

    <div class="tab-nav">
        <button class="tab-btn active" data-tab="overview">Tổng quan</button>
        <button class="tab-btn" data-tab="expenses">Chi phí</button>
        <button class="tab-btn" data-tab="channels">Kênh</button>
        <button class="tab-btn" data-tab="stores">Cửa hàng</button>
    </div>

    <div class="main-content">
        <!-- Tab: Tổng quan -->
        <div class="tab-content active" id="tab-overview">
            <div class="section">
                <div class="section-title">DOANH THU & LỢI NHUẬN</div>
                <div class="charts-grid">
                    <div class="chart-container"><canvas id="revenueChart"></canvas></div>
                    <div class="chart-container"><canvas id="profitChart"></canvas></div>
                </div>
            </div>
            <div class="section">
                <div class="section-title">TĂNG TRƯỞNG</div>
                <div id="growth-bars"></div>
            </div>
        </div>

        <!-- Tab: Chi phí -->
        <div class="tab-content" id="tab-expenses">
            <div class="section">
                <div class="section-title">CHI PHÍ THEO KHOẢN MỤC</div>
                <div class="charts-grid">
                    <div class="chart-container"><canvas id="expensePieChart"></canvas></div>
                    <div class="chart-container"><canvas id="expenseTrendChart"></canvas></div>
                </div>
            </div>
            <div class="section">
                <div class="section-title">CHI TIẾT CHI PHÍ</div>
                <table class="metrics-table" id="expenseTable">
                    <thead><tr><th>Khoản mục</th><th>Tháng 1</th><th>Tháng 2</th><th>Tháng 3</th><th>Thay đổi</th></tr></thead>
                    <tbody></tbody>
                </table>
            </div>
        </div>

        <!-- Tab: Kênh -->
        <div class="tab-content" id="tab-channels">
            <div class="section">
                <div class="section-title">DOANH THU THEO KÊNH</div>
                <div class="charts-grid">
                    <div class="chart-container"><canvas id="channelBarChart"></canvas></div>
                    <div class="chart-container"><canvas id="channelPieChart"></canvas></div>
                </div>
            </div>
        </div>

        <!-- Tab: Cửa hàng -->
        <div class="tab-content" id="tab-stores">
            <div class="section">
                <div class="section-title">TOP CỬA HÀNG</div>
                <table class="metrics-table" id="storeTable">
                    <thead><tr><th>Mã</th><th>Tên</th><th>Khu vực</th><th>Doanh thu</th></tr></thead>
                    <tbody></tbody>
                </table>
            </div>
        </div>
    </div>

    <div class="footer">
        Dashboard tự động | Dữ liệu từ BC LILY 26
    </div>

    <script>
    // Global state
    let financialData = null;
    let selectedMonths = [0, 1, 2];
    let charts = {};

    // Load data
    async function loadData() {
        try {
            const response = await fetch('../data/data.json');
            financialData = await response.json();
            initDashboard();
        } catch (error) {
            console.error('Error loading data:', error);
            alert('Không thể tải dữ liệu. Vui lòng chạy process_data.py trước.');
        }
    }

    // Initialize dashboard
    function initDashboard() {
        document.getElementById('company-name').textContent = financialData.meta.company;
        document.getElementById('last-updated').textContent = 'Cập nhật: ' + financialData.meta.lastUpdated;

        createMonthFilter();
        updateAll();
        setupTabNavigation();
    }

    // Create month filter checkboxes
    function createMonthFilter() {
        const container = document.getElementById('month-filter');
        container.innerHTML = '';

        financialData.months.forEach((month, idx) => {
            const label = document.createElement('label');
            label.className = 'month-checkbox';
            label.innerHTML = `
                <input type="checkbox" value="${idx}" ${selectedMonths.includes(idx) ? 'checked' : ''}>
                <span>${month}</span>
            `;
            label.querySelector('input').addEventListener('change', (e) => {
                if (e.target.checked) {
                    selectedMonths.push(idx);
                } else {
                    selectedMonths = selectedMonths.filter(i => i !== idx);
                }
                if (selectedMonths.length === 0) {
                    selectedMonths = [idx];
                    e.target.checked = true;
                }
                updateAll();
            });
            container.appendChild(label);
        });
    }

    // Update all charts and KPIs
    function updateAll() {
        updateKPICards();
        updateCharts();
    }

    // Format currency
    function formatBillions(num) {
        if (num === 0) return '0';
        const billions = num / 1e9;
        return billions.toFixed(1) + ' tỷ';
    }
    function formatPercent(num) {
        if (num === null || num === undefined) return '--';
        return (num >= 0 ? '+' : '') + num.toFixed(1) + '%';
    }

    // Update KPI cards
    function updateKPICards() {
        if (selectedMonths.length === 1) {
            const idx = selectedMonths[0];
            document.getElementById('kpi-revenue').textContent = formatBillions(financialData.summary.revenue[idx]);
            document.getElementById('kpi-revenue-change').textContent = formatPercent(financialData.growth.revenue[idx]) || '';
            document.getElementById('kpi-revenue-change').className = 'card-change ' + (financialData.growth.revenue[idx] >= 0 ? 'up' : 'down');

            const profit = financialData.summary.netProfit[idx];
            document.getElementById('kpi-profit').textContent = formatBillions(profit);
            document.getElementById('kpi-profit').className = 'card-value ' + (profit >= 0 ? 'positive' : 'negative');
            document.getElementById('kpi-profit-change').textContent = formatPercent(financialData.growth.netProfit[idx]) || '';
            document.getElementById('kpi-profit-change').className = 'card-change ' + (financialData.growth.netProfit[idx] >= 0 ? 'up' : 'down');

            const margin = financialData.summary.netRevenue[idx] > 0 ? (profit / financialData.summary.netRevenue[idx] * 100) : 0;
            document.getElementById('kpi-margin').textContent = margin.toFixed(1) + '%';

            document.getElementById('kpi-status').textContent = profit >= 0 ? 'TÍCH CỰC' : 'CẦN CẢI THIỆN';
            document.getElementById('kpi-status').className = 'card-value ' + (profit >= 0 ? 'positive' : 'negative');
        } else {
            // Show latest month data
            const idx = financialData.months.length - 1;
            const revenue = financialData.summary.revenue[idx];
            const profit = financialData.summary.netProfit[idx];
            const margin = financialData.summary.netRevenue[idx] > 0 ? (profit / financialData.summary.netRevenue[idx] * 100) : 0;

            document.getElementById('kpi-revenue').textContent = formatBillions(revenue);
            document.getElementById('kpi-revenue-change').textContent = '+' + financialData.growth.revenue[idx] + '%';
            document.getElementById('kpi-revenue-change').className = 'card-change up';

            document.getElementById('kpi-profit').textContent = formatBillions(profit);
            document.getElementById('kpi-profit').className = 'card-value ' + (profit >= 0 ? 'positive' : 'negative');
            document.getElementById('kpi-profit-change').textContent = formatPercent(financialData.growth.netProfit[idx]);
            document.getElementById('kpi-profit-change').className = 'card-change up';

            document.getElementById('kpi-margin').textContent = margin.toFixed(1) + '%';

            document.getElementById('kpi-status').textContent = profit >= 0 ? 'TÍCH CỰC' : 'CẦN CẢI THIỆN';
            document.getElementById('kpi-status').className = 'card-value ' + (profit >= 0 ? 'positive' : 'negative');
        }
    }

    // Update all charts
    function updateCharts() {
        // Get filtered data
        const months = selectedMonths.map(i => financialData.months[i]);
        const revenue = selectedMonths.map(i => financialData.summary.revenue[i] / 1e9);
        const netRevenue = selectedMonths.map(i => financialData.summary.netRevenue[i] / 1e9);
        const grossProfit = selectedMonths.map(i => financialData.summary.grossProfit[i] / 1e9);
        const netProfit = selectedMonths.map(i => financialData.summary.netProfit[i] / 1e9);

        // Revenue Chart
        if (charts.revenue) charts.revenue.destroy();
        charts.revenue = new Chart(document.getElementById('revenueChart'), {
            type: 'bar',
            data: {
                labels: months,
                datasets: [
                    { label: 'Doanh thu', data: revenue, backgroundColor: '#1a5276' },
                    { label: 'DT thuần', data: netRevenue, backgroundColor: '#2980b9' },
                    { label: 'LN gộp', data: grossProfit, backgroundColor: '#27ae60' }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { title: { display: true, text: 'Doanh thu (tỷ VND)' }, legend: { position: 'bottom' } },
                scales: { y: { beginAtZero: true } }
            }
        });

        // Profit Chart
        if (charts.profit) charts.profit.destroy();
        charts.profit = new Chart(document.getElementById('profitChart'), {
            type: 'line',
            data: {
                labels: months,
                datasets: [
                    { label: 'LN gộp', data: grossProfit, borderColor: '#27ae60', backgroundColor: 'rgba(39,174,96,0.1)', fill: true, tension: 0.4 },
                    { label: 'LN ròng', data: netProfit, borderColor: '#e74c3c', backgroundColor: 'rgba(231,76,60,0.1)', fill: true, tension: 0.4 }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { title: { display: true, text: 'Lợi nhuận (tỷ VND)' }, legend: { position: 'bottom' } },
                scales: { y: { beginAtZero: true } }
            }
        });

        // Expense Pie Chart (latest selected month)
        const latestIdx = selectedMonths[selectedMonths.length - 1];
        const exp = financialData.expenses;
        if (charts.expensePie) charts.expensePie.destroy();
        charts.expensePie = new Chart(document.getElementById('expensePieChart'), {
            type: 'doughnut',
            data: {
                labels: ['QC/Ads', 'Phí sàn', 'Ship', 'Dịch vụ', 'Lương', 'Khác'],
                datasets: [{
                    data: [
                        exp.ads[latestIdx] / 1e9,
                        exp.platform[latestIdx] / 1e9,
                        exp.ship[latestIdx] / 1e9,
                        exp.service[latestIdx] / 1e9,
                        exp.salary[latestIdx] / 1e9,
                        exp.other[latestIdx] / 1e9
                    ],
                    backgroundColor: ['#ff6b6b', '#4ecdc4', '#45b7d1', '#96ceb4', '#ffeaa7', '#dfe6e9']
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { title: { display: true, text: 'Chi phí ' + financialData.months[latestIdx] + ' (tỷ VND)' }, legend: { position: 'bottom' } }
            }
        });

        // Expense Trend Chart
        const expMonths = selectedMonths.map(i => financialData.months[i]);
        const expAds = selectedMonths.map(i => exp.ads[i] / 1e9);
        const expPlatform = selectedMonths.map(i => exp.platform[i] / 1e9);
        const expShip = selectedMonths.map(i => exp.ship[i] / 1e9);
        if (charts.expenseTrend) charts.expenseTrend.destroy();
        charts.expenseTrend = new Chart(document.getElementById('expenseTrendChart'), {
            type: 'bar',
            data: {
                labels: expMonths,
                datasets: [
                    { label: 'QC/Ads', data: expAds, backgroundColor: '#ff6b6b' },
                    { label: 'Phí sàn', data: expPlatform, backgroundColor: '#4ecdc4' },
                    { label: 'Ship', data: expShip, backgroundColor: '#45b7d1' }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { title: { display: true, text: 'Chi phí theo tháng (tỷ VND)' }, legend: { position: 'bottom' } },
                scales: { y: { beginAtZero: true, stacked: true } }
            }
        });

        // Channel Bar Chart
        const ch = financialData.channels;
        const offData = selectedMonths.map(i => ch.off.total[i] / 1e9);
        const tmdtData = selectedMonths.map(i => ch.tmdt.total[i] / 1e9);
        const onlineData = selectedMonths.map(i => ch.online.total[i] / 1e9);
        if (charts.channelBar) charts.channelBar.destroy();
        charts.channelBar = new Chart(document.getElementById('channelBarChart'), {
            type: 'bar',
            data: {
                labels: months,
                datasets: [
                    { label: 'OFF', data: offData, backgroundColor: '#11998e' },
                    { label: 'TMĐT', data: tmdtData, backgroundColor: '#f5576c' },
                    { label: 'ONLINE', data: onlineData, backgroundColor: '#4facfe' }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { title: { display: true, text: 'Doanh thu theo kênh (tỷ VND)' }, legend: { position: 'bottom' } }
            }
        });

        // Channel Pie Chart
        if (charts.channelPie) charts.channelPie.destroy();
        charts.channelPie = new Chart(document.getElementById('channelPieChart'), {
            type: 'doughnut',
            data: {
                labels: ['OFF', 'TMĐT', 'ONLINE'],
                datasets: [{
                    data: [
                        ch.off.total[latestIdx] / 1e9,
                        ch.tmdt.total[latestIdx] / 1e9,
                        ch.online.total[latestIdx] / 1e9
                    ],
                    backgroundColor: ['#11998e', '#f5576c', '#4facfe']
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { title: { display: true, text: 'Tỷ lệ kênh ' + financialData.months[latestIdx] }, legend: { position: 'bottom' } }
            }
        });

        // Update Growth Bars
        updateGrowthBars();

        // Update Expense Table
        updateExpenseTable();

        // Update Store Table
        updateStoreTable();
    }

    // Update growth bars
    function updateGrowthBars() {
        const container = document.getElementById('growth-bars');
        const growthData = [
            { label: 'Doanh thu', value: financialData.growth.revenue[financialData.months.length - 1], color: '#27ae60' },
            { label: 'LN gộp', value: financialData.growth.grossProfit[financialData.months.length - 1], color: '#38ef7d' },
            { label: 'LN ròng', value: financialData.growth.netProfit[financialData.months.length - 1], color: '#f5576c' }
        ];

        container.innerHTML = growthData.map(g => `
            <div class="growth-bar">
                <div class="growth-label">${g.label}</div>
                <div class="growth-bar-fill">
                    <div class="growth-bar-fill-inner" style="width: ${Math.min(Math.abs(g.value || 0) * 5, 100)}%; background: ${g.value >= 0 ? g.color : '#e74c3c'}">
                        ${formatPercent(g.value)}
                    </div>
                </div>
            </div>
        `).join('');
    }

    // Update expense table
    function updateExpenseTable() {
        const tbody = document.querySelector('#expenseTable tbody');
        const exp = financialData.expenses;
        const expNames = ['ads', 'platform', 'ship', 'service', 'salary', 'other'];
        const labels = ['QC/Ads', 'Phí sàn', 'Ship', 'Dịch vụ', 'Lương', 'Khác'];

        tbody.innerHTML = labels.map((name, i) => {
            const vals = selectedMonths.map(j => exp[expNames[i]][j] / 1e9);
            const change = exp[expNames[i]][financialData.months.length - 1] - exp[expNames[i]][financialData.months.length - 2];
            return `
                <tr>
                    <td>${name}</td>
                    <td>${vals[0]?.toFixed(1) || '--'}</td>
                    <td>${vals[1]?.toFixed(1) || '--'}</td>
                    <td>${vals[2]?.toFixed(1) || '--'}</td>
                    <td style="color: ${change >= 0 ? '#e74c3c' : '#27ae60'}">${change >= 0 ? '↑' : '↓'} ${Math.abs(change / 1e9).toFixed(1)} tỷ</td>
                </tr>
            `;
        }).join('');
    }

    // Update store table (placeholder - needs store-level data)
    function updateStoreTable() {
        const tbody = document.querySelector('#storeTable tbody');
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#999;">Dữ liệu cửa hàng đang được cập nhật</td></tr>';
    }

    // Setup tab navigation
    function setupTabNavigation() {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                btn.classList.add('active');
                document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
            });
        });
    }

    // Initialize on load
    document.addEventListener('DOMContentLoaded', loadData);
    </script>
</body>
</html>
```

- [ ] **Step 2: Test dashboard opens**

```bash
start dashboard.html
```

Expected: Dashboard loads with charts

- [ ] **Step 3: Commit**

```bash
git add dashboard/dashboard.html
git commit -m "feat: create dashboard HTML with Chart.js"
```

---

## Task 4: Create Manual Input Excel Template

**Files:**
- Create: `data_nhap_tay.xlsx`

- [ ] **Step 1: Create Excel template with openpyxl**

```python
#!/usr/bin/env python3
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Du Lieu Tong Hop'

# Styles
header_fill = PatternFill(start_color='1a5276', end_color='1a5276', fill_type='solid')
header_font = Font(bold=True, color='FFFFFF', size=11)
title_font = Font(bold=True, size=14)
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

# Title
ws['A1'] = 'BÁO CÁO TÀI CHÍNH - DỮ LIỆU TỔNG HỢP'
ws['A1'].font = title_font
ws.merge_cells('A1:F1')

ws['A2'] = 'Đơn vị: VND'
ws['A3'] = ''

# Headers
headers = ['Chỉ tiêu', 'Tháng 1', 'Tháng 2', 'Tháng 3', 'Tăng trưởng', 'Ghi chú']
for col, header in enumerate(headers, 1):
    cell = ws.cell(row=4, column=col, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center')
    cell.border = thin_border

# Data rows
data_rows = [
    ('1. Doanh thu', '', '', '', '', ''),
    ('2. Giảm trừ DT', '', '', '', '', ''),
    ('3. Doanh thu thuần', '', '', '', '', ''),
    ('4. Giá vốn', '', '', '', '', ''),
    ('5. Lợi nhuận gộp', '', '', '', '', ''),
    ('6. Chi phí bán hàng', '', '', '', '', ''),
    ('7. Chi phí quản lý', '', '', '', '', ''),
    ('8. Lợi nhuận thuần', '', '', '', '', ''),
    ('', '', '', '', '', ''),
    ('CHI PHÍ CHI TIẾT', '', '', '', '', ''),
    ('- QC/Ads', '', '', '', '', ''),
    ('- Phí sàn TMĐT', '', '', '', '', ''),
    ('- Phí vận chuyển', '', '', '', '', ''),
    ('- Dịch vụ/Thuê nhà', '', '', '', '', ''),
    ('- Lương/BHXH', '', '', '', '', ''),
    ('- Khác', '', '', '', '', ''),
    ('', '', '', '', '', ''),
    ('KÊNH DOANH THU', '', '', '', '', ''),
    ('- OFF (Cửa hàng)', '', '', '', '', ''),
    ('- TMĐT', '', '', '', '', ''),
    ('- ONLINE', '', '', '', '', ''),
]

for row_idx, row_data in enumerate(data_rows, 5):
    for col_idx, value in enumerate(row_data, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.border = thin_border
        if col_idx > 1 and value:
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal='right')

# Column widths
ws.column_dimensions['A'].width = 25
ws.column_dimensions['B'].width = 15
ws.column_dimensions['C'].width = 15
ws.column_dimensions['D'].width = 15
ws.column_dimensions['E'].width = 12
ws.column_dimensions['F'].width = 20

wb.save('data_nhap_tay.xlsx')
print('Created data_nhap_tay.xlsx')
```

- [ ] **Step 2: Run to create template**

```bash
python create_template.py
```

- [ ] **Step 3: Commit**

```bash
git add data_nhap_tay.xlsx
git commit -m "feat: create manual input Excel template"
```

---

## Task 5: Create README and Final Testing

**Files:**
- Create: `README.md`

- [ ] **Step 1: Create README**

```markdown
# Dashboard Tài Chính - BC LILY 26

Hệ thống báo cáo tự động cho Công Ty Cổ Phần ULTD Thịnh Phát.

## Cấu trúc

```
BC LILY 26/
├── THÁNG *.26 BÁO CÁO THÁNG *.xlsx    # File báo cáo gốc
├── scripts/
│   └── process_data.py                  # ETL script
├── data/
│   └── data.json                        # Dữ liệu JSON
├── dashboard/
│   └── dashboard.html                   # Dashboard HTML
└── data_nhap_tay.xlsx                   # File nhập liệu thủ công
```

## Sử dụng

### 1. Cập nhật dữ liệu tự động

1. Copy file báo cáo Excel tháng mới vào thư mục
2. Chạy: `python scripts/process_data.py`
3. Mở `dashboard/dashboard.html` để xem

### 2. Nhập liệu thủ công

1. Mở `data_nhap_tay.xlsx`
2. Điền số liệu vào cột Tháng 1, 2, 3
3. Chạy script chuyển đổi (nếu cần)

### 3. Xem Dashboard

Mở `dashboard/dashboard.html` trong trình duyệt.

- **Tab Tổng quan**: KPIs, doanh thu, lợi nhuận
- **Tab Chi phí**: Chi tiết chi phí theo khoản mục
- **Tab Kênh**: Performance theo kênh OFF/TMĐT/ONL
- **Tab Cửa hàng**: Top cửa hàng

### 4. Lọc tháng

Sử dụng checkboxes để chọn hiển thị 1 hoặc nhiều tháng.

## Yêu cầu

- Python 3.8+
- pandas
- openpyxl

```bash
pip install pandas openpyxl
```

## Tác giả

Tự động tạo bởi Claude Code
```

- [ ] **Step 2: Final test - run ETL and verify dashboard**

```bash
python scripts/process_data.py
# Verify data.json created
# Open dashboard.html and check charts
```

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: add README and complete project setup"
```

---

## Verification Checklist

After all tasks complete, verify:

- [ ] `python scripts/process_data.py` runs without errors
- [ ] `data/data.json` contains all 3 months of data
- [ ] `dashboard/dashboard.html` opens in browser with charts
- [ ] Month filter checkboxes work
- [ ] All 4 tabs navigate correctly
- [ ] Revenue and Profit charts display data
- [ ] Expense charts display data
- [ ] Channel charts display data
- [ ] `data_nhap_tay.xlsx` template created

---

## Summary

| Task | Files | Status |
|------|-------|--------|
| 1 | Folders + data.json schema | ☐ |
| 2 | scripts/process_data.py | ☐ |
| 3 | dashboard/dashboard.html | ☐ |
| 4 | data_nhap_tay.xlsx | ☐ |
| 5 | README.md + verification | ☐ |
