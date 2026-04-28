# 📊 HƯỚNG DẪN TẠO DASHBOARD POWER BI

## 1. CHUẨN BỊ DỮ LIỆU

### File dữ liệu đã tạo:
- ✅ **Data_For_PowerBI.xlsx** - File Excel chuẩn hóa để import vào Power BI
- ✅ **QuanLy_ChiPhi_SanXuat.xlsx** - File quản lý với nhiều sheet

### Cấu trúc dữ liệu:
| Cột | Mô tả |
|-----|-------|
| Tháng | Kỳ hiệu tháng (T01-T12) |
| Năm | Năm báo cáo |
| Tháng_Số | Số tháng (1-12) |
| Lương_CN_Truc_Tiep | Lương công nhân trực tiếp |
| Chi_Phi_Dien | Chi phí điện |
| Dau_Dap | Chi phí dầu dập |
| Tong_Luong_Hang | Tổng sản lượng |
| Tong_Chi_Phi | Tổng chi phí (tự động tính) |
| Chi_Phi_Don_Vi | Chi phí trên mỗi đơn vị SP |

---

## 2. CÁC BƯỚC TẠO POWER BI

### Bước 1: Import dữ liệu
1. Mở **Power BI Desktop**
2. Click **Home → Get Data → Excel Workbook**
3. Chọn file **Data_For_PowerBI.xlsx**
4. Chọn sheet và click **Load**

### Bước 2: Tạo Measures (DAX)

Mở tab **Modeling → New Measure** và tạo các measure sau:

```dax
// Tổng chi phí lương
Tong_Luong = SUM('Data_For_PowerBI'[Lương_CN_Truc_Tiep])

// Tổng chi phí điện
Tong_Dien = SUM('Data_For_PowerBI'[Chi_Phi_Dien])

// Tổng dầu dập
Tong_Dau = SUM('Data_For_PowerBI'[Dau_Dap])

// Tổng sản lượng
Tong_San_Luong = SUM('Data_For_PowerBI'[Tong_Luong_Hang])

// Tổng tất cả chi phí
Tong_Chi_Phi = [Tong_Luong] + [Tong_Dien] + [Tong_Dau]

// Chi phí trung bình mỗi tháng
TB_Chi_Phi_Thang = AVERAGE('Data_For_PowerBI'[Tong_Chi_Phi])

// Chi phí đơn vị TB
TB_Chi_Phi_Don_Vi = DIVIDE([Tong_Chi_Phi], [Tong_San_Luong], 0)
```

### Bước 3: Thiết kế Dashboard

#### 📌 CARD (Chỉ số tổng quan)
- Tổng chi phí năm: `[Tong_Chi_Phi]`
- Tổng sản lượng: `[Tong_San_Luong]`
- Chi phí TB/tháng: `[TB_Chi_Phi_Thang]`
- Chi phí/SP: `[TB_Chi_Phi_Don_Vi]`

#### 📈 LINE CHART (Xu hướng)
- **Trục X**: Tháng_Số
- **Trục Y**: Lương_CN_Truc_Tiep, Chi_Phi_Dien, Dau_Dap
- **Legend**: Loại chi phí

#### 🥧 PIE CHART (Tỷ lệ chi phí)
- **Values**: Giá trị từng loại chi phí
- **Legend**: Loại chi phí

#### 📊 COLUMN CHART (So sánh)
- **Trục X**: Tháng
- **Trục Y**: Tong_Luong_Hang (Sản lượng)

#### 📉 AREA CHART (Chi phí đơn vị)
- **Trục X**: Tháng
- **Trục Y**: Chi_Phi_Don_Vi

### Bước 4: Thêm Slicer (Bộ lọc)
- Thêm **Slicer** cho cột **Tháng** để lọc theo tháng
- Thêm **Slicer** cho cột **Năm** để lọc theo năm

---

## 3. GỢI Ý LAYOUT DASHBOARD

```
┌─────────────────────────────────────────────────────────────┐
│                    CHI PHÍ SẢN XUẤT 2025                    │
├──────────────┬──────────────┬──────────────┬────────────────┤
│   💰 TỔNG    │   ⚡ TỔNG     │   🛢️ TỔNG    │   📦 TỔNG      │
│   CHI PHÍ    │   ĐIỆN       │   DẦU DẬP    │   SẢN LƯỢNG    │
├──────────────┴──────────────┴──────────────┴────────────────┤
│                                                             │
│   [LINE CHART: Xu hướng chi phí theo tháng]                 │
│                                                             │
├───────────────────────────────┬─────────────────────────────┤
│                               │                             │
│   [PIE CHART: Tỷ lệ chi phí]  │  [BAR CHART: Sản lượng]     │
│                               │                             │
├───────────────────────────────┴─────────────────────────────┤
│                                                             │
│   [TABLE: Chi tiết theo tháng]                              │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## 4. CẬP NHẬT DỮ LIỆU

### Cách 1: Cập nhật thủ công
1. Cập nhật file **Data_For_PowerBI.xlsx**
2. Trong Power BI: **Home → Refresh**

### Cách 2: Tự động (Scheduled Refresh)
1. Publish lên **Power BI Service**
2. Cài đặt **Scheduled Refresh** mỗi ngày/giờ

---

## 5. XUẤT BÁO CÁO

- **File → Export → Power BI template** (.pbit)
- **File → Publish** lên Power BI Service
- **File → Export → PDF** để xuất báo cáo tĩnh

---

## 6. LƯU Ý QUAN TRỌNG

1. **Định dạng số**: Format các cột tiền tệ với dấu phân cách hàng nghìn
2. **Màu sắc**: 
   - Lương CN: Xanh dương (#4472C4)
   - Chi phí điện: Vàng (#FFC000)
   - Dầu dập: Cam (#ED7D31)
3. **Tiêu đề**: Rõ ràng, font size lớn
4. **Tooltip**: Thêm tooltip cho các chart để hiển thị chi tiết khi hover
