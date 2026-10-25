# Excel Web Pro - Trình Quản Lý Bảng Tính Chuyên Nghiệp (Excel Clone)

Ứng dụng web clone Excel hoàn chỉnh được xây dựng trên nền tảng **Python Flask** và **JavaScript Native (Vanilla CSS/JS)**, hỗ trợ đọc, chỉnh sửa, tính toán công thức, vẽ biểu đồ và xuất file `.xlsx` chuẩn xác về máy tính.

---

## 🌟 Tính Năng Nổi Bật

1. **Giao Diện Chuẩn Microsoft Excel**:
   - Thanh điều khiển Ribbon đầy đủ: *Tệp (File), Trang chủ (Home), Chèn (Insert), Công thức (Formulas), Dữ liệu (Data), Xem (View)*.
   - Hỗ trợ Dark Mode / Light Mode, thanh trượt phóng to/thu nhỏ (Zoom 50% - 200%).
   - Thanh công thức (Formula Bar) kết hợp hộp tên ô (Name Box), gợi ý hàm tự động (Autocomplete).

2. **Công Thức & Tính Toán Đầy Đủ**:
   - Toán học: `SUM`, `AVERAGE`, `MIN`, `MAX`, `PRODUCT`, `ROUND`, `ABS`, `SQRT`, `POWER`, `MOD`, `INT`...
   - Thống kê & Điều kiện: `COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`, `SUMIF`, `MEDIAN`, `STDEV`...
   - Logic: `IF`, `IFS`, `AND`, `OR`, `NOT`...
   - Văn bản: `CONCAT`, `TEXTJOIN`, `UPPER`, `LOWER`, `PROPER`, `LEN`, `TRIM`, `LEFT`, `RIGHT`, `MID`...
   - Tra cứu & Ngày tháng: `VLOOKUP`, `INDEX`, `MATCH`, `TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`...

3. **Chỉnh Sửa & Tương Tác Ô Siêu Mượt**:
   - Nhập liệu trực tiếp, nhấn **Enter** chuyển ngay xuống ô liền kề.
   - Điều hướng bằng phím mũi tên mượt mà, hỗ trợ Tab, Shift+Tab, F2.
   - Điền tự động thông minh (Smart Drag-to-Fill handle): Tự tăng tiến số và chuỗi (ví dụ `ref-1` ➔ `ref-2`, `biscuit-01` ➔ `biscuit-02`).
   - Lưới ô vô hạn: Tự động mở rộng thêm hàng/cột khi cuộn hoặc điều hướng tới biên.

4. **Định Dạng Ô & Biểu Đồ Động**:
   - Phông chữ, kích thước, in đậm, nghiêng, gạch chân, màu chữ, màu nền, viền ô, căn chỉnh, ngắt dòng (Wrap Text), gộp ô (Merge Cells).
   - Định dạng số: Tiền tệ (₫ VND / $ USD), Phần trăm (%), Tăng/Giảm chữ số thập phân.
   - Tạo biểu đồ động nổi (Floating Charts) từ vùng chọn: Cột, Đường, Tròn, Miền.

5. **Quản Lý Đa Trang Tính (Multi-sheet)**:
   - Thêm sheet mới (`+`), đổi tên sheet (nhấp đúp hoặc chuột phải).
   - Nút xoá trang tính (`x`) thông minh: Xoá tức thì với sheet trống, xác nhận an toàn với sheet có dữ liệu.
   - Hỗ trợ hoàn tác / làm lại đầy đủ (**Ctrl+Z** / **Ctrl+Y**).

6. **Nhập & Xuất File .XLSX Trực Tiếp**:
   - Mở và đọc file Excel (`.xlsx`, `.xls`, `.csv`).
   - Xuất file `.xlsx` chuẩn OpenPyXL trực tiếp về thư mục **Downloads** của máy tính với đầy đủ định dạng, công thức và màu sắc.

---

## 🚀 Cài Đặt & Khởi Chạy

### Yêu Cầu
- Python 3.8 trở lên
- pip

### Cài Đặt Thư Viện
```bash
pip install -r requirements.txt
```

### Khởi Chạy Bản Web (Trình duyệt)
```bash
python app.py
```
Ứng dụng sẽ tự động mở trên trình duyệt tại: `http://127.0.0.1:5000`

### Khởi Chạy Bản Desktop (Cửa sổ độc lập như Excel)
```bash
python desktop_app.py
```

### Đóng Gói Thành File .EXE Độc Lập (Windows)
```bash
python -m PyInstaller --noconsole --onefile --name "ExcelWebPro" --icon "app_icon.ico" --add-data "templates;templates" --add-data "static;static" --add-data "app_icon.ico;." --collect-all webview --collect-all clr_loader --collect-all pythonnet --collect-all openpyxl --collect-all jinja2 --hidden-import "webview.platforms.winforms" --hidden-import "webview.platforms.edgechromium" --hidden-import "clr" desktop_app.py
```
File `ExcelWebPro.exe` sau khi đóng gói sẽ nằm trong thư mục `dist/`.

---

## 📁 Cấu Trúc Dự Án
```
excel-web-editor/
├── app.py                # Backend Flask xử lý đọc/xuất file Excel (OpenPyXL)
├── desktop_app.py        # Trình khởi chạy Desktop Window qua Microsoft Edge WebView2
├── app_icon.ico          # Icon chuẩn màu xanh Excel cho ứng dụng
├── create_icon.py        # Kịch bản tạo icon đa độ phân giải
├── requirements.txt      # Thư viện phụ thuộc (Flask, openpyxl, pywebview, pyinstaller)
├── static/
│   ├── app.js            # Core logic bảng tính, công thức, phím tắt, giao diện
│   ├── styles.css        # Giao diện Fluent UI hiện đại, Dark/Light mode
│   ├── favicon.ico       # Favicon tab trình duyệt
│   └── libs/             # Thư viện SheetJS & Chart.js offline
└── templates/
    └── index.html        # Khung giao diện HTML5 chuẩn SEO
```

---

## 📄 Bản Quyền
Dự án được phát triển phục vụ học tập và làm việc.
