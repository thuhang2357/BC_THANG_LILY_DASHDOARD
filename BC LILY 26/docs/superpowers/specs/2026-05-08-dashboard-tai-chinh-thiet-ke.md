# SPEC: Hệ Thống Dashboard Tài Chính - BC LILY 26

## 1. Concept & Vision

Hệ thống báo cáo tự động cho Công Ty Cổ Phần ULTD Thịnh Phát, phục vụ CFO và CEO. Dashboard hiển thị tổng quan tài chính với khả năng drill-down theo tháng, kênh, và cửa hàng. Mỗi tháng chỉ cần copy file Excel vào thư mục và chạy script - dashboard tự cập nhật.

**Guiding principle**: "Một dashboard cho CFO thấy chi phí, cho CEO thấy cơ hội tăng trưởng."

---

## 2. Design Language

### Color Palette
- **Primary**: #1a5276 (Navy blue - chuyên nghiệp, tin cậy)
- **Secondary**: #2980b9 (Light blue - tươi hơn cho accent)
- **Success/Up**: #27ae60 (Green)
- **Warning/Down**: #e74c3c (Red)
- **Background**: #f5f7fa (Light gray)
- **Card Background**: #ffffff

### Typography
- Font: Segoe UI, system-ui fallback
- Headings: 18-24px, bold
- Body: 14px regular
- Numbers: 28px bold cho KPI cards

### Chart Colors
- OFF (Offline): #11998e (Teal)
- TMĐT (E-commerce): #f5576c (Coral)
- ONLINE: #4facfe (Sky blue)
- Chi phí ads: #ff6b6b
- Chi phí platform: #4ecdc4
- Chi phí ship: #45b7d1

---

## 3. Data Structure

### 3.1 Core Metrics Schema (data.json)

```json
{
  "meta": {
    "company": "Công Ty Cổ Phần ULTD Thịnh Phát",
    "lastUpdated": "2026-05-08",
    "currency": "VND"
  },
  "months": ["Tháng 1", "Tháng 2", "Tháng 3"],
  "monthKeys": ["01", "02", "03"],

  "summary": {
    "revenue": [21213664770, 22370281393, 25006659252],
    "revenueDiscount": [4765686041, 4471573844, 4290664121],
    "netRevenue": [16447978729, 17898707549, 20715995131],
    "cogs": [5808244302, 6334658083, 6690056300],
    "grossProfit": [10639734427, 11564049466, 14025938831],
    "sellingExpense": [6361170144, 5893273451, 7298613288],
    "managementExpense": [4302304311, 4666746169, 5599533426],
    "netProfit": [-10826144, 1068429882, 1123728119],
    "cash": [12815665162, 14975896465, 19082189578]
  },

  "growth": {
    "revenue": [null, 5.5, 11.8],
    "grossProfit": [null, 8.7, 21.3],
    "netProfit": [null, 101.8, 5.2]
  },

  "expenses": {
    "ads": [1105227512, 967939027, 1617014256],
    "platform": [2605488475, 2254970952, 3206557556],
    "ship": [266980826, 106874420, 515388492],
    "service": [5122302478, 5253218936, 5268257131],
    "salary": [3152975504, 3635098101, 2689803167],
    "other": [19387974227, 19040832880, 19024763572]
  },

  "channels": {
    "off": {
      "total": [0, 15766439000, 16442572570],
      "north": [0, 0, 0],
      "south": [0, 0, 0],
      "central": [0, 0, 0]
    },
    "tmdt": {
      "total": [0, 6292232739, 8011302352],
      "shopee": [0, 0, 0],
      "tiktok": [0, 0, 0]
    },
    "online": {
      "total": [0, 325768322, 490202183],
      "seeding": [0, 0, 0],
      "web": [0, 0, 0]
    }
  },

  "stores": {
    "north": [
      {"code": "115KM", "name": "Kim Mã", "revenue": [0, 488, 520]}
    ],
    "south": [...],
    "central": [...]
  }
}
```

---

## 4. Dashboard Layout

### 4.1 Header
- Company name + "BÁO CÁO TÀI CHÍNH"
- Month/Year selector dropdown
- Last updated date

### 4.2 Filter Bar
- **Month filter**: Checkboxes cho "Hiện tất cả" hoặc chọn từng tháng
- Default: 3 tháng gần nhất
- Khi chọn 1 tháng: Hiển thị chi tiết tháng đó
- Khi chọn nhiều tháng: Hiển thị so sánh

### 4.3 KPI Cards (Row 1)
| Card | Nội dung | Color Logic |
|------|----------|-------------|
| DOANH THU | Giá trị + % growth vs tháng trước | Green if up, Red if down |
| LỢI NHUẬN | Giá trị + % growth | Green if positive, Red if negative |
| TỶ SUẤT LN | Gross margin % | Green if > 50%, Yellow if 30-50%, Red if < 30% |
| ĐÁNH GIÁ | Tích cực / Cần cải thiện | Green/Red badge |

### 4.4 Tab Navigation
1. **Tổng quan**: Summary + Revenue/Profit charts
2. **Chi phí**: Detailed expense breakdown
3. **Kênh**: OFF/TMĐT/ONL performance
4. **Cửa hàng**: Top/Bottom stores table
5. **So sánh**: Month-over-month analysis

---

## 5. Charts Specification

### 5.1 Tab Tổng quan
- **Revenue Bar Chart**: Grouped bar - Revenue, Net Revenue, Gross Profit (tỷ VND)
- **Profit Line Chart**: Line - Gross Profit, Net Profit trend

### 5.2 Tab Chi phí
- **Expense Pie Chart**: Doughnut - ads, platform, ship, service, salary, other
- **Expense Trend Chart**: Stacked bar - 3 tháng chi phí theo category
- **Expense Table**: Chi tiết từng khoản mục với % change

### 5.3 Tab Kênh
- **Channel Bar Chart**: Grouped bar - OFF, TMĐT, ONL theo tháng
- **Channel Pie Chart**: Doughnut - % contribution tháng hiện tại
- **Channel Trend Line**: Line chart - growth trajectory

### 5.4 Tab Cửa hàng
- **Store Table**: Sortable by revenue/profit
  - Columns: Mã cửa hàng, Tên, Khu vực, Doanh thu, Lợi nhuận, Tỷ suất
- **Top 10 Bar Chart**: Horizontal bar - top performers

### 5.5 Tab So sánh
- **Month selector**: Chọn tháng để so sánh
- **Comparison metrics**: Side-by-side KPIs
- **Variance analysis**: +/- vs previous month

---

## 6. Technical Approach

### 6.1 Stack
- **Data Processing**: Python 3 + pandas + openpyxl
- **Dashboard Frontend**: HTML + CSS + Chart.js (CDN)
- **Data Storage**: JSON file (data.json)
- **Export**: Python xlsxwriter for Excel export

### 6.2 File Structure
```
BC LILY 26/
├── THÁNG 01.26 BÁO CÁO THÁNG 01.xlsx
├── THÁNG 02.26 BÁO CÁO THÁNG 02 .xlsx
├── THÁNG 03.26 BÁO CÁO THÁNG 03.xlsx
├── data/
│   ├── data.json              # Central data file
│   └── data_nhap_tay.xlsx    # Manual input template
├── scripts/
│   └── process_data.py       # ETL script
├── dashboard/
│   └── dashboard.html        # Main dashboard
└── export/                   # Generated reports
```

### 6.3 ETL Process (process_data.py)
1. Scan for Excel files matching "THÁNG *.xlsx" pattern
2. For each file:
   - Read BCKQHĐKD sheet → summary metrics
   - Read BC Chi Phí sheet → expense breakdown
   - Read BC LCTT sheet → cash flow
   - Read Báo cáo cửa hàng → channel & store data
3. Aggregate into data.json
4. Calculate growth percentages

---

## 7. User Flows

### 7.1 Monthly Update Flow
1. Nhận file báo cáo từ kế toán
2. Copy vào thư mục data/
3. Run: `python scripts/process_data.py`
4. Open dashboard/dashboard.html
5. (Optional) Export PDF/Excel

### 7.2 Ad-hoc Analysis Flow
1. Open dashboard
2. Use filters to select time period
3. Navigate tabs for different views
4. Drill down to store/channel level
5. Export if needed

---

## 8. MVP Scope

### Included in MVP
- [x] Đọc 3 tháng từ Excel files
- [x] Summary KPIs: Revenue, Profit, Margin
- [x] Revenue & Profit charts
- [x] Expense breakdown chart
- [x] Channel performance chart
- [x] Month filter (all 3 or individual)
- [x] Excel export

### Deferred
- [ ] Store-level drill-down (requires more data parsing)
- [ ] Cash flow detailed chart
- [ ] PDF export
- [ ] Automated email report
- [ ] User authentication

---

## 9. Success Criteria

- CFO có thể xem chi phí chi tiết trong 30 giây
- CEO có thể thấy xu hướng tăng trưởng trong 15 giây
- Update dashboard mới chỉ mất 2 phút (copy file + run script)
- Dashboard load trong < 3 giây với Chart.js
