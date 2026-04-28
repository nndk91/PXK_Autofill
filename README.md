# 🏭 HỆ THỐNG QUẢN LÝ CHI PHÍ SẢN XUẤT

Hệ thống quản lý chi phí sản xuất với bảng biểu Excel, form nhập liệu web và hướng dẫn tạo Power BI dashboard.

---

## 📁 CẤU TRÚC FILE

| File | Mô tả |
|------|-------|
| `QuanLy_ChiPhi_SanXuat.xlsx` | File Excel chính với 3 sheet: Dữ liệu gốc, Nhập liệu, Báo cáo |
| `Data_For_PowerBI.xlsx` | File dữ liệu chuẩn hóa để import vào Power BI |
| `app.py` | Ứng dụng web Streamlit để nhập liệu |
| `HuongDan_PowerBI.md` | Hướng dẫn chi tiết tạo dashboard Power BI |
| `CHIPHI SX T12.2025 - gửi anh Khoa.xlsx` | File dữ liệu gốc |

---

## 🚀 HƯỚNG DẪN SỬ DỤNG

### Cách 1: Sử dụng Excel (Đơn giản)

1. Mở file **`QuanLy_ChiPhi_SanXuat.xlsx`**
2. Vào sheet **"NhapLieuChiPhi"**
3. Nhập dữ liệu vào các dòng trống (đã có 38 dòng trống sẵn)
4. File sẽ tự động tính tổng ở sheet **"BaoCaoTongHop"**

### Cách 2: Sử dụng Web App (Trực quan)

```bash
# 1. Kích hoạt môi trường
source venv/bin/activate

# 2. Chạy ứng dụng
streamlit run app.py
```

Sau đó mở trình duyệt tại: http://localhost:8501

**Tính năng:**
- 📊 Xem tổng quan chi phí
- ➕ Nhập/cập nhật dữ liệu qua form
- 📈 Xem biểu đồ phân tích

### Cách 3: Power BI Dashboard (Chuyên nghiệp)

Xem file **`HuongDan_PowerBI.md`** để tạo dashboard chuyên nghiệp.

Tóm tắt:
1. Mở Power BI Desktop
2. Import file `Data_For_PowerBI.xlsx`
3. Tạo các measure DAX
4. Thiết kế dashboard với charts

---

## 📊 DỮ LIỆU HIỆN TẠI (2025)

| Chỉ Tiêu | Giá Trị |
|----------|---------|
| Tổng Lương CN | 12,443,606,026 VNĐ |
| Tổng Chi Phí Điện | 3,071,157,054 VNĐ |
| Tổng Dầu Dập | 2,888,400,000 VNĐ |
| Tổng Sản Lượng | 1,453,854 SP |
| **TỔNG CHI PHÍ** | **18,403,163,080 VNĐ** |
| Chi Phí/SP | ~12,660 VNĐ |

---

## 🛠️ CÀI ĐẶT MÔI TRƯỜNG

```bash
# Tạo virtual environment
python3 -m venv venv

# Kích hoạt
source venv/bin/activate  # Mac/Linux
# hoặc
venv\Scripts\activate  # Windows

# Cài đặt thư viện
pip install pandas openpyxl xlsxwriter streamlit
```

---

## 📝 GHI CHÚ

- Dữ liệu được lưu tự động vào file Excel
- Có thể nhập dữ liệu cho nhiều năm
- Hỗ trợ cập nhật dữ liệu tháng đã nhập trước đó

---

## 📞 HỖ TRỢ

Nếu cần hỗ trợ thêm, vui lòng liên hệ.
