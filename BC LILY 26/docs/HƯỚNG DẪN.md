# HƯỚNG DẪN SỬ DỤNG DASHBOARD TÀI CHÍNH
## Công Ty Cổ Phần ULTD Thịnh Phát

---

## 1. GIỚI THIỆU

Dashboard tài chính giúp theo dõi và phân tích báo cáo tài chính theo tháng, bao gồm:
- Doanh thu, lợi nhuận, chi phí
- Phân tích theo kênh (Miền Bắc, Miền Nam, Miền Trung, TMĐT, Online)
- Xu hướng theo tháng

---

## 2. CẤU TRÚC THƯ MỤC

```
BC LILY 26/
├── dashboard/
│   ├── dashboard.html           ← Mở file này trong trình duyệt
│   ├── revenue_detail.html      ← Chi tiết doanh thu
│   ├── profit_detail.html       ← Chi tiết lợi nhuận
│   ├── expense_detail.html      ← Chi tiết chi phí
│   └── expense_detail_summary.js
├── data/
│   └── data.json               ← Dữ liệu đã xử lý
├── monthly_reports/            ← Copy file Excel vào đây
│   ├── THÁNG 01.26 BÁO CÁO THÁNG 01.xlsx
│   ├── THÁNG 02.26 BÁO CÁO THÁNG 02 .xlsx
│   └── THÁNG 03.26 BÁO CÁO THÁNG 03.xlsx
├── scripts/
│   └── process_data.py         ← Script xử lý dữ liệu
├── generate_expense_detail.py
├── update_all.py               ← Script cập nhật chính
└── docs/
    └── HƯỚNG DẪN.md          ← File này
```

---

## 3. THÊM THÁNG MỚI

### Bước 1: Copy file Excel vào thư mục `monthly_reports/`

- File Excel cần đặt tên đúng định dạng:
  ```
  THÁNG XX.26 BÁO CÁO THÁNG XX.xlsx
  ```
- Ví dụ:
  - `THÁNG 04.26 BÁO CÁO THÁNG 04.xlsx`
  - `THÁNG 05.26 BÁO CÁO THÁNG 05.xlsx`
  - `THÁNG 06.26 BÁO CÁO THÁNG 06.xlsx`

### Bước 2: Chạy script cập nhật

Mở Terminal/Command Prompt và chạy:

```bash
cd "c:\Users\ADMIN\Downloads\BC LILY 26"
python update_all.py
```

Script sẽ tự động:
1. Đọc tất cả file Excel trong `monthly_reports/`
2. Xử lý dữ liệu từ sheet BCKQHĐKD và Báo cáo cửa hàng
3. Tạo file `data/data.json`
4. Cập nhật dashboard

### Bước 3: Mở dashboard kiểm tra

Mở file trong trình duyệt:
```
dashboard/dashboard.html
```

---

## 4. CHIA SẺ DASHBOARD CHO NGƯỜI KHÁC

### Cách 1: Upload lên Web Server (Khuyến nghị)

1. Upload toàn bộ thư mục `dashboard/` lên web hosting
2. Đảm bảo cấu hình CORS cho file JSON (nếu có)
3. Gửi link cho người khác: `https://ten-mien.com/dashboard/dashboard.html`

### Cách 2: Nén và gửi qua Cloud Storage

1. Nén thư mục `dashboard/` thành ZIP
2. Upload lên Google Drive / OneDrive / Dropbox
3. Chia sẻ link tải về
4. Người nhận giải nén và mở `dashboard.html`

### Cách 3: Local Server (cho mạng nội bộ)

1. Cài đặt Python (đã có sẵn)
2. Mở Terminal, chạy:
```bash
cd "c:\Users\ADMIN\Downloads\BC LILY 26"
python -m http.server 8000
```
3. Người khác truy cập: `http://[IP-may-ban]:8000/dashboard.html`

---

## 5. CÁC CHỈ SỐ VÀ GIỚI HẠN

### Lợi nhuận thuần (Tối thiểu - phải trên)
| Kênh | Tối thiểu |
|-------|-----------|
| Miền Bắc | >13% |
| Miền Nam | >9% |
| Miền Trung | >10% |
| TMĐT | >18% |
| Online | >15% |

### Tỷ lệ chiết khấu (Tối đa - phải dưới)
| Kênh | Tối đa |
|-------|--------|
| Miền Bắc/Nam/Trung/Online | <15% |
| TMĐT | <20% |

### Tỷ lệ GV/DT thuần (Tối đa - phải dưới)
| Kênh | Tối đa |
|-------|--------|
| Miền Bắc/Nam/Trung/Online | <36% |
| TMĐT | <45% |

### Tỷ lệ CP/DT thuần (Tối đa - phải dưới)
| Kênh | Tối đa |
|-------|--------|
| Miền Bắc/Nam/Trung/Online | <36% |
| TMĐT | <45% |

---

## 6. XỬ LÝ SỰ CỐ

### Lỗi "Không hiển thị dữ liệu"
- Kiểm tra file `data/data.json` có tồn tại không
- Chạy lại `python update_all.py`

### Lỗi "Font tiếng Việt"
- Đảm bảo trình duyệt hỗ trợ UTF-8 (Chrome, Firefox, Edge đều OK)

### Lỗi "CORS" khi mở trực tiếp
- Nếu mở file HTML trực tiếp bị lỗi, dùng cách Local Server ở trên
- Hoặc upload lên hosting

### Dashboard bị trắng
- Kiểm tra file `dashboard.html` có bị mã độc chèn vào không
- Chạy lại `python update_all.py` để tái tạo

---

## 7. THÔNG TIN LIÊN HỆ

- Công Ty: Cổ Phần ULTD Thịnh Phát
- Dashboard được tạo: 2026-05-09
- Phiên bản: 1.0

---

## 8. GHI CHÚ KỸ THUẬT

### Nguồn dữ liệu
- Sheet `BCKQHĐKD`: Tổng hợp toàn công ty
- Sheet `Báo cáo cửa hàng`: Chi tiết theo kênh
- Sheet `Chi tiết chi phí`: Chi tiết từng khoản mục

### Bugs đã sửa
1. `parse_vietnamese_number` bị lỗi với số thập phân từ Excel
2. Dòng SUM trong Báo cáo cửa hàng không được sử dụng đúng
3. Định dạng số Việt Nam (1.234.567) bị lẫn với decimal

---

*Cập nhật lần cuối: 2026-05-09*
