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
└── data_nhap_tay.xlsx                  # File nhập liệu thủ công
```

## Sử dụng

### 1. Cập nhật dữ liệu tự động

1. Copy file báo cáo Excel tháng mới vào thư mục
2. Chạy: `python scripts/process_data.py`
3. Mở `dashboard/dashboard.html` để xem

### 2. Nhập liệu thủ công

1. Mở `data_nhap_tay.xlsx`
2. Điền số liệu vào cột Tháng 1, 2, 3
3. (Tùy chọn) Viết script để chuyển đổi sang JSON

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

## Dữ liệu

### data.json structure

```json
{
  "meta": { "company": "...", "lastUpdated": "...", "currency": "VND" },
  "months": ["Tháng 1", "Tháng 2", "Tháng 3"],
  "summary": { "revenue": [...], "netRevenue": [...], "grossProfit": [...], "netProfit": [...] },
  "growth": { "revenue": [...], "grossProfit": [...], "netProfit": [...] },
  "expenses": { "ads": [...], "platform": [...], "ship": [...], "service": [...], "salary": [...], "other": [...] },
  "channels": { "off": {"total": [...]}, "tmdt": {"total": [...]}, "online": {"total": [...]} }
}
```

## Tác giả

Tự động tạo bởi Claude Code