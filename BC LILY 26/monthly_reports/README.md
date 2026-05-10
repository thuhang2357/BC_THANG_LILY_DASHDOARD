# Hướng dẫn thêm dữ liệu tháng mới

## Cấu trúc Folder

```
monthly_reports/
├── THÁNG 01.26 BÁO CÁO THÁNG 01.xlsx
├── THÁNG 02.26 BÁO CÁO THÁNG 02 .xlsx
├── THÁNG 03.26 BÁO CÁO THÁNG 03.xlsx
├── THÁNG 04.26 BÁO CÁO THÁNG 04.xlsx
├── THÁNG 05.26 BÁO CÁO THÁNG 05.xlsx
├── THÁNG 06.26 BÁO CÁO THÁNG 06.xlsx
└── ... (các tháng tiếp theo)
```

## Quy tắc đặt tên file

- Format: `THÁNG XX.26 BÁO CÁO THÁNG XX.xlsx`
- XX = số tháng (01, 02, 03, ... 12)
- Lưu ý: có khoảng trắng trước số tháng thứ 2 trong một số file cũ

## Các bước thực hiện

1. **Copy file báo cáo tháng mới** vào folder `monthly_reports/`

2. **Chạy script ETL** để cập nhật data.json:
   ```bash
   python scripts/process_data.py
   ```

3. **Mở dashboard** để xem kết quả:
   - Mở file `dashboard/dashboard.html` trong trình duyệt
   - Hoặc deploy lên web server

## Sheet cần có trong file Excel

File Excel cần có các sheet sau:
- `BCKQHĐKD` - Báo cáo kết quả hoạt động kinh doanh
- `Báo cáo cửa hàng` - Chi tiết theo cửa hàng
- `BC Chi Phí` - Chi tiết chi phí
- `Chi tiết chi phí` - (nếu có)

## Tự động cập nhật

Khi có đủ dữ liệu, dashboard sẽ tự động hiển thị:
- So sánh 2 tháng liên tiếp
- Biểu đồ xu hướng
- Chi tiết theo kênh và cửa hàng