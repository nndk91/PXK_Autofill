# 📦 PXK Manager v5 - Hướng dẫn sử dụng

## ⚡ Cách chạy nhanh

### Bước 1: Cài đặt thư viện (chỉ cần làm 1 lần)
```bash
pip install streamlit pandas openpyxl pdfplumber PyMuPDF
```

### Bước 2: Chạy ứng dụng
```bash
streamlit run app_v5.py
```

Hoặc chạy file batch:
```bash
run_app.bat
```

Sau đó mở trình duyệt vào địa chỉ: http://localhost:8501

---

## 📁 Cấu trúc thư mục

```
AI HỌC/
├── app_v5.py              # Ứng dụng chính (chạy file này)
├── pxk_core_v4.py         # Core xử lý + AI học dữ liệu
├── pdf_extractor.py       # Trích xuất từ PDF
├── excel_writer.py        # Ghi file Excel
├── run_app.bat            # File chạy tự động Windows
├── run_app.ps1            # File chạy PowerShell
├── requirements.txt       # Danh sách thư viện
│
├── 2033-2096/             # Dữ liệu huấn luyện
│   ├── FILE TRONG 2033-2096.xlsx      (form trống)
│   ├── FORM_DA_DIEN_PXK 2033-2096.xlsx (form đã điền - để học)
│   └── pxk/               # Folder chứa PDF
│
├── 2144-2172/             # Dữ liệu huấn luyện
├── 2168-2243/             # Dữ liệu huấn luyện
└── 2273-2316/             # Dữ liệu huấn luyện
```

---

## 🎯 Các chế độ xử lý

### 1️⃣ Chế độ "Chỉ trích xuất PDF" (Bước 1 riêng)
**Dùng khi:** Chỉ cần xem dữ liệu từ file PDF, chưa cần ghép vào form

**Cách dùng:**
1. Chọn chế độ: 📄 Chỉ bước 1: Trích xuất dữ liệu
2. Upload file PDF (có thể chọn nhiều file)
3. Bấm "Bắt đầu xử lý"
4. Tải file Excel kết quả về

**Kết quả:** File Excel chứa dữ liệu đã trích xuất từ tất cả PDF

---

### 2️⃣ Chế độ "Trích xuất + Ghép form" (Cả 2 bước)
**Dùng khi:** Cần ghép số PXK vào form trống

**Cách dùng:**
1. Chọn chế độ: 📋 Cả 2 bước: Trích xuất PDF + Ghép vào Form
2. Upload file PDF (phiếu xuất kho)
3. Upload file Excel form trống
4. Bấm "Bắt đầu xử lý"
5. Tải file form đã điền PXK về

**Kết quả:** File Excel form đã được điền số PXK tự động

---

## 🧠 Tính năng AI học từ dữ liệu (core_v4)

App v5 sử dụng **pxk_core_v4** với khả năng:
- Tự động phát hiện các folder dữ liệu đã điền
- Học từ các mẫu đã điền để chấm điểm các trường hợp mơ hồ
- Sử dụng beam search + scoring để chọn PXK tối ưu
- Cải thiện độ chính xác qua các lần học

---

## 🎨 Màu sắc kết quả

| Màu | Ý nghĩa |
|-----|---------|
| 🟢 Xanh | ✅ Tự động - 1 PXK duy nhất khớp |
| 🟡 Vàng | 🔍 Cần kiểm tra - Có nhiều khả năng |
| 🔴 Đỏ | ❌ Không khớp - Cần điền thủ công |

---

## 📝 Lưu ý quan trọng

### Về file PDF:
- Cần là file Phiếu Xuất Kho định dạng PDF
- Không bị mật khẩu bảo vệ

### Về file Form:
- Định dạng Excel (.xlsx)
- Cột 3: Số hóa đơn (Invoice)
- Cột 4: Mã hàng
- Cột 5: Số lượng GR
- Sheet tên "XUẤT" hoặc "Sheet1"

### Về dữ liệu huấn luyện:
Core_v4 tự động tìm các folder có:
- Thư mục `pxk/` chứa file PDF
- File form trống (tên chứa "không có" / "khong co")
- File form đã điền (tên chứa "có dữ liệu" / "co du lieu")

---

## ❓ Lỗi thường gặp

**Lỗi: ModuleNotFoundError**
```bash
pip install streamlit pandas openpyxl pdfplumber PyMuPDF
```

**Lỗi: Không tìm thấy lệnh streamlit**
```bash
python -m streamlit run app_v5.py
```

**Lỗi: Không đọc được PDF**
- Kiểm tra file PDF không bị mật khẩu bảo vệ
- Thử cài thêm: `pip install pdf2image`

**Lỗi: Không tìm thấy dữ liệu huấn luyện**
- Đảm bảo các folder có định dạng đúng: `XXXX-YYYY`
- Mỗi folder cần có thư mục `pxk/` chứa PDF
- File Excel cần có tên phù hợp (xem phần Lưu ý)
