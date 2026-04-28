import streamlit as st
import pandas as pd
from datetime import datetime
import os

st.set_page_config(page_title="Quản Lý Chi Phí Sản Xuất", layout="wide")

# CSS tùy chỉnh
st.markdown("""
<style>
    .main-header {
        font-size: 32px;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        padding: 20px;
        background: linear-gradient(90deg, #e3f2fd, #bbdefb);
        border-radius: 10px;
        margin-bottom: 30px;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .form-section {
        background-color: #ffffff;
        padding: 25px;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🏭 QUẢN LÝ CHI PHÍ SẢN XUẤT</div>', unsafe_allow_html=True)

# File Excel
EXCEL_FILE = "QuanLy_ChiPhi_SanXuat.xlsx"

# Hàm đọc dữ liệu
def load_data():
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE, sheet_name='NhapLieuChiPhi')
        df = df.dropna(subset=['Tháng'])
        return df
    return pd.DataFrame()

# Hàm lưu dữ liệu
def save_data(df):
    # Đọc file hiện tại
    with pd.ExcelFile(EXCEL_FILE) as xls:
        df_goc = pd.read_excel(xls, sheet_name='DuLieuGoc', header=None)
        df_baocao = pd.read_excel(xls, sheet_name='BaoCaoTongHop')
    
    # Ghi lại với sheet mới
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        df_goc.to_excel(writer, sheet_name='DuLieuGoc', index=False, header=False)
        df.to_excel(writer, sheet_name='NhapLieuChiPhi', index=False)
        df_baocao.to_excel(writer, sheet_name='BaoCaoTongHop', index=False)

# Tabs
tab1, tab2, tab3 = st.tabs(["📊 Tổng Quan", "➕ Nhập Dữ Liệu", "📈 Phân Tích"])

# ===== TAB 1: Tổng Quan =====
with tab1:
    st.subheader("📊 Tổng Quan Chi Phí Sản Xuất 2025")
    
    df = load_data()
    
    if not df.empty:
        # Tính tổng
        total_luong = df['Lương CN Trực Tiếp'].sum()
        total_dien = df['Chi Phí Điện'].sum()
        total_dau = df['Dầu Dập'].sum()
        total_hang = df['Tổng Lượng Hàng'].sum()
        total_chiphi = total_luong + total_dien + total_dau
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("💰 Tổng Lương CN", f"{total_luong:,.0f} VNĐ")
        with col2:
            st.metric("⚡ Chi Phí Điện", f"{total_dien:,.0f} VNĐ")
        with col3:
            st.metric("🛢️ Dầu Dập", f"{total_dau:,.0f} VNĐ")
        with col4:
            st.metric("📦 Tổng Lượng Hàng", f"{total_hang:,.0f}")
        
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("💵 TỔNG CHI PHÍ", f"{total_chiphi:,.0f} VNĐ")
        with col2:
            if total_hang > 0:
                cost_per_unit = total_chiphi / total_hang
                st.metric("📊 Chi Phí/SP", f"{cost_per_unit:,.0f} VNĐ")
        
        st.markdown("---")
        st.subheader("📋 Dữ Liệu Chi Tiết")
        st.dataframe(df, use_container_width=True, height=400)
    else:
        st.info("Chưa có dữ liệu. Vui lòng nhập dữ liệu ở tab 'Nhập Dữ Liệu'")

# ===== TAB 2: Nhập Dữ Liệu =====
with tab2:
    st.subheader("➕ Nhập Dữ Liệu Mới")
    
    with st.form("nhap_lieu_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            thang = st.selectbox("Tháng", ['T01', 'T02', 'T03', 'T04', 'T05', 'T06', 
                                           'T07', 'T08', 'T09', 'T10', 'T11', 'T12'])
            nam = st.number_input("Năm", min_value=2020, max_value=2030, value=2025)
            luong_cn = st.number_input("Lương CN Trực Tiếp (VNĐ)", min_value=0, value=0, step=1000000)
        
        with col2:
            chi_phi_dien = st.number_input("Chi Phí Điện (VNĐ)", min_value=0, value=0, step=100000)
            dau_dap = st.number_input("Dầu Dập (VNĐ)", min_value=0, value=0, step=100000)
            tong_hang = st.number_input("Tổng Lượng Hàng", min_value=0, value=0, step=1000)
        
        ghi_chu = st.text_area("Ghi Chú", placeholder="Nhập ghi chú nếu có...")
        
        submitted = st.form_submit_button("💾 Lưu Dữ Liệu", use_container_width=True)
        
        if submitted:
            df = load_data()
            
            # Kiểm tra tháng đã tồn tại chưa
            if thang in df['Tháng'].values:
                # Cập nhật
                df.loc[df['Tháng'] == thang, ['Năm', 'Lương CN Trực Tiếp', 'Chi Phí Điện', 
                                               'Dầu Dập', 'Tổng Lượng Hàng', 'Ghi Chú']] = \
                    [nam, luong_cn, chi_phi_dien, dau_dap, tong_hang, ghi_chu]
                st.success(f"✅ Đã cập nhật dữ liệu tháng {thang}!")
            else:
                # Thêm mới
                new_row = pd.DataFrame({
                    'Tháng': [thang],
                    'Năm': [nam],
                    'Lương CN Trực Tiếp': [luong_cn],
                    'Chi Phí Điện': [chi_phi_dien],
                    'Dầu Dập': [dau_dap],
                    'Tổng Lượng Hàng': [tong_hang],
                    'Ghi Chú': [ghi_chu]
                })
                df = pd.concat([df, new_row], ignore_index=True)
                st.success(f"✅ Đã thêm dữ liệu tháng {thang}!")
            
            save_data(df)
            st.balloons()

# ===== TAB 3: Phân Tích =====
with tab3:
    st.subheader("📈 Phân Tích Chi Phí")
    
    df = load_data()
    
    if not df.empty:
        # Sắp xếp theo tháng
        month_order = {f'T{i:02d}': i for i in range(1, 13)}
        df['ThangSo'] = df['Tháng'].map(month_order)
        df = df.sort_values('ThangSo')
        
        # Biểu đồ 1: Xu hướng chi phí
        st.markdown("#### 📉 Xu Hướng Chi Phí Theo Tháng")
        
        chart_data = df.set_index('Tháng')[['Lương CN Trực Tiếp', 'Chi Phí Điện', 'Dầu Dập']]
        st.line_chart(chart_data)
        
        # Biểu đồ 2: Tỷ lệ chi phí
        st.markdown("#### 🥧 Tỷ Lệ Các Khoản Chi Phí")
        
        total_luong = df['Lương CN Trực Tiếp'].sum()
        total_dien = df['Chi Phí Điện'].sum()
        total_dau = df['Dầu Dập'].sum()
        
        pie_data = pd.DataFrame({
            'Loại Chi Phí': ['Lương CN', 'Chi Phí Điện', 'Dầu Dập'],
            'Số Tiền': [total_luong, total_dien, total_dau]
        })
        
        st.bar_chart(pie_data.set_index('Loại Chi Phí'))
        
        # Biểu đồ 3: Sản lượng
        st.markdown("#### 📦 Sản Lượng Theo Tháng")
        st.bar_chart(df.set_index('Tháng')['Tổng Lượng Hàng'])
        
        # Chi phí đơn vị
        st.markdown("#### 💵 Chi Phí/Đơn Vị Sản Phẩm Theo Tháng")
        df['ChiPhiDonVi'] = (df['Lương CN Trực Tiếp'] + df['Chi Phí Điện'] + df['Dầu Dập']) / df['Tổng Lượng Hàng'].replace(0, 1)
        st.line_chart(df.set_index('Tháng')['ChiPhiDonVi'])
        
    else:
        st.info("Chưa có dữ liệu để phân tích.")

# Footer
st.markdown("---")
st.markdown("<p style='text-align: center; color: #666;'>🏭 Hệ Thống Quản Lý Chi Phí Sản Xuất © 2025</p>", unsafe_allow_html=True)
