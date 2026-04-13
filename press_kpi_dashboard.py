#!/usr/bin/env python3
"""
Press KPI Dashboard - Streamlit Version
Run with: streamlit run press_kpi_dashboard.py
"""

import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
from pathlib import Path
import glob

def find_latest_excel_file():
    """Find the most recently modified Excel file in the current directory"""
    excel_patterns = ['*.xlsm', '*.xlsx', '*.xls']
    files = []
    for pattern in excel_patterns:
        files.extend(glob.glob(pattern))
    
    if not files:
        return None
    
    # Get the most recently modified file
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

# Page config
st.set_page_config(
    page_title="Press KPI Dashboard - SMC",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2rem;
        font-weight: bold;
        color: #3b8eea;
        margin-bottom: 1rem;
    }
    .kpi-card {
        background-color: #141720;
        border-radius: 10px;
        padding: 20px;
        border-left: 3px solid #3b8eea;
    }
    .kpi-value {
        font-size: 2rem;
        font-weight: bold;
        color: #e8eaf0;
    }
    .kpi-label {
        font-size: 0.8rem;
        color: #8a90a8;
        text-transform: uppercase;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #1c2030;
        border-radius: 4px 4px 0 0;
        padding: 10px 20px;
        color: #8a90a8;
    }
    .stTabs [aria-selected="true"] {
        background-color: #3b8eea !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# Helper functions
@st.cache_data
def load_data(excel_path, file_mtime=None):
    """Load and process data from Excel file
    
    Args:
        excel_path: Path to the Excel file
        file_mtime: File modification time (used for cache invalidation)
    """
    # file_mtime is used to invalidate cache when file changes
    # Try to detect the correct header row
    df_raw = pd.read_excel(excel_path, sheet_name='Sum', header=None)
    
    # Find the row with 'No' and 'Shift' columns (header row)
    header_row = None
    for i in range(min(15, len(df_raw))):
        row_values = df_raw.iloc[i].astype(str).tolist()
        if 'No' in row_values and 'Shift' in row_values:
            if i + 1 < len(df_raw):
                next_row_first = df_raw.iloc[i + 1, 0]
                if pd.notna(next_row_first) and str(next_row_first).strip().isdigit():
                    header_row = i
                    break
            for j in range(i + 1, min(i + 8, len(df_raw))):
                if j < len(df_raw):
                    check_val = df_raw.iloc[j, 0]
                    if pd.notna(check_val) and str(check_val).strip().isdigit():
                        header_row = i
                        break
    
    if header_row is None:
        header_row = 6  # fallback
    
    # Load data with detected header
    df_sum = df_raw.iloc[header_row + 1:].copy()
    df_sum.columns = df_raw.iloc[header_row]
    df_sum = df_sum.iloc[1:].reset_index(drop=True)
    
    # Clean data types
    df_sum['Date'] = pd.to_datetime(df_sum['Date'], errors='coerce')
    df_sum['Plan'] = pd.to_numeric(df_sum['Plan'], errors='coerce').fillna(0)
    df_sum['Capacity'] = pd.to_numeric(df_sum['Capacity'], errors='coerce').fillna(0)
    df_sum['Loss time'] = pd.to_numeric(df_sum['Loss time'], errors='coerce').fillna(0)
    df_sum['Q.ty defect'] = pd.to_numeric(df_sum['Q.ty defect'], errors='coerce').fillna(0)
    df_sum['Operating time'] = pd.to_numeric(df_sum['Operating time'], errors='coerce').fillna(0)
    df_sum = df_sum.dropna(subset=['Date'])
    
    # Load Loss data
    df_loss = pd.read_excel(excel_path, sheet_name='Data_Loss', header=0)
    df_loss['Date'] = pd.to_datetime(df_loss['Date'], errors='coerce')
    df_loss['Loss time'] = pd.to_numeric(df_loss['Loss time'], errors='coerce').fillna(0)
    df_loss = df_loss.dropna(subset=['Date'])
    
    return df_sum, df_loss

def calculate_uph(capacity, operating_time):
    if operating_time and operating_time > 0:
        return round(capacity / operating_time, 2)
    return 0

def get_iso_week(date):
    """Get ISO week format: Week 1 = Dec 29 - Jan 4"""
    year = date.isocalendar()[0]
    week = date.isocalendar()[1]
    return f"{year}-W{week:02d}"

def ensure_part_code_string(val):
    """Ensure Part Code is string, add ' prefix if it's numeric"""
    if pd.isna(val):
        return ''
    val_str = str(val).strip()
    # Check if value is numeric (all digits)
    if val_str.isdigit():
        return "'" + val_str
    return val_str

def process_data(df_sum, df_loss):
    """Process data for all views"""
    # Add derived columns
    df_sum['YearMonth'] = df_sum['Date'].dt.to_period('M').astype(str)
    df_sum['YearWeek'] = df_sum['Date'].apply(get_iso_week)
    df_sum['Date_str'] = df_sum['Date'].dt.strftime('%Y-%m-%d')
    df_sum['UPH'] = df_sum.apply(lambda x: calculate_uph(x['Capacity'], x['Operating time']), axis=1)
    df_sum['OEE'] = df_sum.apply(lambda x: round(x['Capacity'] / x['Plan'] * 100, 2) if x['Plan'] > 0 else 0, axis=1)
    
    # Ensure Part code is string with ' prefix for numeric values
    df_sum['Part code'] = df_sum['Part code'].apply(ensure_part_code_string)
    
    # Loss data time columns
    df_loss['YearMonth'] = df_loss['Date'].dt.to_period('M').astype(str)
    df_loss['YearWeek'] = df_loss['Date'].apply(get_iso_week)
    df_loss['Date_str'] = df_loss['Date'].dt.strftime('%Y-%m-%d')
    
    # Ensure Part code is string with ' prefix for numeric values
    df_loss['Part code'] = df_loss['Part code'].apply(ensure_part_code_string)
    
    return df_sum, df_loss

def get_time_aggregates(df_sum, period='monthly'):
    """Get time-based aggregates"""
    if period == 'monthly':
        agg = df_sum.groupby('YearMonth').agg({
            'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum', 
            'Loss time': 'sum', 'Operating time': 'sum',
            'Date': ['min', 'max']
        }).reset_index()
        agg.columns = ['Period', 'Plan', 'Capacity', 'Q.ty defect', 'Loss time', 'Operating time', 'Start_Date', 'End_Date']
    elif period == 'weekly':
        agg = df_sum.groupby('YearWeek').agg({
            'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum',
            'Loss time': 'sum', 'Operating time': 'sum',
            'Date': ['min', 'max']
        }).reset_index()
        agg.columns = ['Period', 'Plan', 'Capacity', 'Q.ty defect', 'Loss time', 'Operating time', 'Start_Date', 'End_Date']
    else:  # daily
        agg = df_sum.groupby('Date_str').agg({
            'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum',
            'Loss time': 'sum', 'Operating time': 'sum'
        }).reset_index()
        agg.columns = ['Period', 'Plan', 'Capacity', 'Q.ty defect', 'Loss time', 'Operating time']
        agg['Start_Date'] = agg['Period']
        agg['End_Date'] = agg['Period']
    
    agg['OEE'] = agg.apply(lambda x: round(x['Capacity'] / x['Plan'] * 100, 2) if x['Plan'] > 0 else 0, axis=1)
    agg['PPM'] = agg.apply(lambda x: round(x['Q.ty defect'] / x['Capacity'] * 1e6, 0) if x['Capacity'] > 0 else 0, axis=1)
    agg['UPH'] = agg.apply(lambda x: calculate_uph(x['Capacity'], x['Operating time']), axis=1)
    
    if period != 'daily':
        agg['Start_Date'] = agg['Start_Date'].dt.strftime('%d/%m/%Y')
        agg['End_Date'] = agg['End_Date'].dt.strftime('%d/%m/%Y')
    
    return agg.sort_values('Period')

def get_by_code_data(df_sum, period='all', from_period=None, to_period=None):
    """Get data by part code"""
    if period == 'all':
        code_data = df_sum.groupby(['Part code', 'Part name']).agg({
            'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum',
            'Loss time': 'sum', 'Operating time': 'sum'
        }).reset_index()
    elif period == 'monthly':
        mask = (df_sum['YearMonth'] >= from_period) & (df_sum['YearMonth'] <= to_period)
        code_data = df_sum[mask].groupby(['Part code', 'Part name']).agg({
            'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum',
            'Loss time': 'sum', 'Operating time': 'sum'
        }).reset_index()
    elif period == 'weekly':
        mask = (df_sum['YearWeek'] >= from_period) & (df_sum['YearWeek'] <= to_period)
        code_data = df_sum[mask].groupby(['Part code', 'Part name']).agg({
            'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum',
            'Loss time': 'sum', 'Operating time': 'sum'
        }).reset_index()
    else:  # daily
        mask = (df_sum['Date_str'] >= from_period) & (df_sum['Date_str'] <= to_period)
        code_data = df_sum[mask].groupby(['Part code', 'Part name']).agg({
            'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum',
            'Loss time': 'sum', 'Operating time': 'sum'
        }).reset_index()
    
    # Re-apply string conversion after groupby (groupby may reset type)
    code_data['Part code'] = code_data['Part code'].apply(ensure_part_code_string)
    
    code_data['OEE'] = code_data.apply(lambda x: round(x['Capacity'] / x['Plan'] * 100, 2) if x['Plan'] > 0 else 0, axis=1)
    code_data['PPM'] = code_data.apply(lambda x: round(x['Q.ty defect'] / x['Capacity'] * 1e6, 0) if x['Capacity'] > 0 else 0, axis=1)
    code_data['UPH'] = code_data.apply(lambda x: calculate_uph(x['Capacity'], x['Operating time']), axis=1)
    code_data['NG_Rate'] = code_data.apply(lambda x: round(x['Q.ty defect'] / x['Capacity'] * 100, 2) if x['Capacity'] > 0 else 0, axis=1)
    
    return code_data.sort_values('Capacity', ascending=False)

def get_shift_data(df_sum, period='monthly', from_period=None, to_period=None):
    """Get shift comparison data"""
    if period == 'monthly':
        mask = (df_sum['YearMonth'] >= from_period) & (df_sum['YearMonth'] <= to_period)
        shift_data = df_sum[mask].groupby(['YearMonth', 'Shift']).agg({
            'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum',
            'Loss time': 'sum', 'Operating time': 'sum'
        }).reset_index()
        shift_data.rename(columns={'YearMonth': 'Period'}, inplace=True)
    elif period == 'weekly':
        mask = (df_sum['YearWeek'] >= from_period) & (df_sum['YearWeek'] <= to_period)
        shift_data = df_sum[mask].groupby(['YearWeek', 'Shift']).agg({
            'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum',
            'Loss time': 'sum', 'Operating time': 'sum'
        }).reset_index()
        shift_data.rename(columns={'YearWeek': 'Period'}, inplace=True)
    else:  # daily
        mask = (df_sum['Date_str'] >= from_period) & (df_sum['Date_str'] <= to_period)
        shift_data = df_sum[mask].groupby(['Date_str', 'Shift']).agg({
            'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum',
            'Loss time': 'sum', 'Operating time': 'sum'
        }).reset_index()
        shift_data.rename(columns={'Date_str': 'Period'}, inplace=True)
    
    shift_data['OEE'] = shift_data.apply(lambda x: round(x['Capacity'] / x['Plan'] * 100, 2) if x['Plan'] > 0 else 0, axis=1)
    shift_data['UPH'] = shift_data.apply(lambda x: calculate_uph(x['Capacity'], x['Operating time']), axis=1)
    
    return shift_data

def get_loss_data(df_loss, period='all', from_period=None, to_period=None):
    """Get loss time analysis data with optional period filtering"""
    # Filter by period if specified
    filtered_loss = df_loss.copy()
    if period != 'all' and from_period and to_period:
        if period == 'monthly':
            filtered_loss = df_loss[(df_loss['YearMonth'] >= from_period) & (df_loss['YearMonth'] <= to_period)]
        elif period == 'weekly':
            filtered_loss = df_loss[(df_loss['YearWeek'] >= from_period) & (df_loss['YearWeek'] <= to_period)]
        elif period == 'daily':
            filtered_loss = df_loss[(df_loss['Date_str'] >= from_period) & (df_loss['Date_str'] <= to_period)]
    
    loss_reason = filtered_loss.groupby('Reason')['Loss time'].sum().reset_index().sort_values('Loss time', ascending=False).head(10)
    loss_dept = filtered_loss.groupby('Dept PIC')['Loss time'].sum().reset_index().sort_values('Loss time', ascending=False)
    loss_type = filtered_loss.groupby('Loss type')['Loss time'].sum().reset_index()
    
    # Loss by code với lý do phổ biến nhất
    loss_by_code = filtered_loss.groupby(['Part code', 'Part name'])['Loss time'].sum().reset_index().sort_values('Loss time', ascending=False).head(20)
    # Ensure Part code is string after groupby
    loss_by_code['Part code'] = loss_by_code['Part code'].apply(ensure_part_code_string)
    
    # Thêm lý do phổ biến nhất cho mỗi part code
    top_reasons = filtered_loss.groupby(['Part code', 'Reason'])['Loss time'].sum().reset_index()
    top_reasons = top_reasons.sort_values('Loss time', ascending=False).groupby('Part code').first().reset_index()
    top_reasons = top_reasons[['Part code', 'Reason']].rename(columns={'Reason': 'Top_Reason'})
    # Ensure Part code is string
    top_reasons['Part code'] = top_reasons['Part code'].apply(ensure_part_code_string)
    loss_by_code = loss_by_code.merge(top_reasons, on='Part code', how='left')
    
    # Phân loại theo kế hoạch
    loss_plan_type = filtered_loss.groupby('Loss type')['Loss time'].sum().reset_index().sort_values('Loss time', ascending=False)
    
    # Chi tiết loss với thởi gian
    loss_details = filtered_loss[['Date', 'Part code', 'Part name', 'Loss type', 'Reason', 'Dept PIC', 'Start time', 'End time', 'Loss time']].sort_values('Loss time', ascending=False).head(50)
    
    total_loss = filtered_loss['Loss time'].sum()
    
    return loss_reason, loss_dept, loss_type, loss_by_code, loss_plan_type, loss_details, total_loss

# Main app
def main():
    # Sidebar
    with st.sidebar:
        st.markdown("<div class='main-header'>📊 SMC Press KPI</div>", unsafe_allow_html=True)
        st.markdown("---")
        
        # File upload
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls', 'xlsm'])
        
        if uploaded_file is None:
            # Try to find the latest Excel file
            latest_file = find_latest_excel_file()
            if latest_file:
                uploaded_file = latest_file
                st.success(f"Using latest file: {latest_file}")
            else:
                st.warning("Please upload an Excel file")
                return
        
        st.markdown("---")
        st.markdown("### Navigation")
        page = st.radio("Select Page", [
            "🏠 Overview",
            "📦 By Code", 
            "🕐 By Shift",
            "⚠️ Loss Time"
        ])
        
        st.markdown("---")
        st.markdown("<div style='font-size:0.8rem;color:#8a90a8;'>UPH = Actual / Operating Time</div>", unsafe_allow_html=True)
        
        # Refresh button to clear cache
        if st.button("🔄 Refresh Data", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    
    # Validate and load data
    if isinstance(uploaded_file, str):
        if not uploaded_file.endswith(('.xlsx', '.xls', '.xlsm')):
            st.error("❌ File không hợp lệ! Vui lòng upload file Excel (.xlsx, .xls, .xlsm)")
            return
        # Get file modification time for cache invalidation
        file_mtime = os.path.getmtime(uploaded_file)
        df_sum, df_loss = load_data(uploaded_file, file_mtime)
    else:
        # Check file extension
        if not uploaded_file.name.endswith(('.xlsx', '.xls', '.xlsm')):
            st.error("❌ File không hợp lệ! Vui lòng upload file Excel (.xlsx, .xls, .xlsm)")
            st.stop()
            return
        # Save uploaded file temporarily
        temp_path = f"temp_{uploaded_file.name}"
        with open(temp_path, 'wb') as f:
            f.write(uploaded_file.getvalue())
        try:
            # Get file modification time for cache invalidation
            file_mtime = os.path.getmtime(temp_path)
            df_sum, df_loss = load_data(temp_path, file_mtime)
        except Exception as e:
            st.error(f"❌ Lỗi khi đọc file: {str(e)}")
            st.info("💡 Vui lòng kiểm tra file có đúng định dạng Excel không.")
            os.remove(temp_path)
            return
        os.remove(temp_path)
    
    df_sum, df_loss = process_data(df_sum, df_loss)
    
    # Summary stats
    total_plan = df_sum['Plan'].sum()
    total_capacity = df_sum['Capacity'].sum()
    total_defect = df_sum['Q.ty defect'].sum()
    total_loss = df_loss['Loss time'].sum()
    avg_uph = calculate_uph(total_capacity, df_sum['Operating time'].sum())
    
    # Display page content
    if page == "🏠 Overview":
        show_overview(df_sum, total_plan, total_capacity, total_defect, total_loss, avg_uph)
    elif page == "📦 By Code":
        show_by_code(df_sum, df_loss)
    elif page == "🕐 By Shift":
        show_by_shift(df_sum)
    else:
        show_loss_time(df_sum, df_loss)

def show_overview(df_sum, total_plan, total_capacity, total_defect, total_loss, avg_uph):
    st.markdown("<div class='main-header'>📊 Tổng quan sản xuất</div>", unsafe_allow_html=True)
    
    # KPI Cards
    cols = st.columns(6)
    with cols[0]:
        st.metric("Sản lượng thực tế", f"{total_capacity:,.0f}", "pcs")
    with cols[1]:
        oee = (total_capacity/total_plan*100) if total_plan > 0 else 0
        st.metric("Tỉ lệ hoàn thành", f"{oee:.2f}%", "vs plan")
    with cols[2]:
        st.metric("Kế hoạch", f"{total_plan:,.0f}", "pcs")
    with cols[3]:
        st.metric("Tổng lỗi NG", f"{total_defect:,.0f}", f"{total_defect/total_capacity*1e6:.0f} PPM" if total_capacity > 0 else "0 PPM")
    with cols[4]:
        st.metric("Loss time", f"{total_loss:.1f}", "giờ")
    with cols[5]:
        st.metric("Avg UPH", f"{avg_uph:.2f}", "pcs/hour")
    
    st.markdown("---")
    
    # Time period selection
    tab_monthly, tab_weekly, tab_daily = st.tabs(["📅 Theo tháng", "📆 Theo tuần", "📋 Theo ngày"])
    
    with tab_monthly:
        show_time_view(df_sum, 'monthly')
    with tab_weekly:
        show_time_view(df_sum, 'weekly')
    with tab_daily:
        show_time_view(df_sum, 'daily')

def show_time_view(df_sum, period):
    """Show time-based view with filters"""
    # Get aggregates
    data = get_time_aggregates(df_sum, period)
    
    # Date range filter
    col1, col2 = st.columns(2)
    with col1:
        from_val = st.selectbox("Từ:", data['Period'].tolist(), index=0, key=f"{period}_from")
    with col2:
        to_val = st.selectbox("Đến:", data['Period'].tolist(), index=len(data)-1, key=f"{period}_to")
    
    # Filter data
    from_idx = data[data['Period'] == from_val].index[0]
    to_idx = data[data['Period'] == to_val].index[0]
    filtered = data.iloc[from_idx:to_idx+1]
    
    # Charts - 4 chart trên 1 hàng
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        fig = go.Figure()
        fig.add_trace(go.Bar(name='Plan', x=filtered['Period'], y=filtered['Plan'], marker_color='#3b8eea'))
        fig.add_trace(go.Bar(name='Actual', x=filtered['Period'], y=filtered['Capacity'], marker_color='#2dd4a0'))
        fig.update_layout(
            title="📊 Plan vs Actual", 
            barmode='group',
            plot_bgcolor='#141720',
            paper_bgcolor='#0d0f14',
            font_color='#e8eaf0',
            xaxis_gridcolor='rgba(255,255,255,0.1)',
            yaxis_gridcolor='rgba(255,255,255,0.1)',
            height=250,
            margin=dict(l=10, r=10, t=30, b=10),
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1, font=dict(size=8))
        )
        st.plotly_chart(fig, use_container_width=True, key=f"{period}_time_chart")
    
    with col2:
        # UPH Chart
        fig = go.Figure()
        fig.add_trace(go.Bar(x=filtered['Period'], y=filtered['UPH'], marker_color='#38c4d4', name='UPH'))
        fig.update_layout(
            title="⚡ UPH",
            plot_bgcolor='#141720',
            paper_bgcolor='#0d0f14',
            font_color='#e8eaf0',
            xaxis_gridcolor='rgba(255,255,255,0.1)',
            yaxis_gridcolor='rgba(255,255,255,0.1)',
            height=250,
            margin=dict(l=10, r=10, t=30, b=10),
            showlegend=False
        )
        st.plotly_chart(fig, use_container_width=True, key=f"{period}_uph_chart")
    
    with col3:
        # Line performance
        line_data = df_sum.groupby('Line').agg({'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum'}).reset_index()
        fig = go.Figure()
        fig.add_trace(go.Bar(name='Plan', x=line_data['Line'], y=line_data['Plan'], marker_color='#3b8eea'))
        fig.add_trace(go.Bar(name='Actual', x=line_data['Line'], y=line_data['Capacity'], marker_color='#2dd4a0'))
        fig.add_trace(go.Bar(name='NG', x=line_data['Line'], y=line_data['Q.ty defect'], marker_color='#f05c5c'))
        fig.update_layout(
            title="🏭 Theo dây chuyền",
            barmode='group',
            plot_bgcolor='#141720',
            paper_bgcolor='#0d0f14',
            font_color='#e8eaf0',
            height=250,
            margin=dict(l=10, r=10, t=30, b=10),
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1, font=dict(size=8))
        )
        st.plotly_chart(fig, use_container_width=True, key=f"{period}_line_chart")
    
    with col4:
        # Loss time by period
        fig = go.Figure()
        fig.add_trace(go.Bar(x=filtered['Period'], y=filtered['Loss time'], marker_color='#f05c5c', name='Loss'))
        fig.update_layout(
            title="⏱️ Loss Time",
            plot_bgcolor='#141720',
            paper_bgcolor='#0d0f14',
            font_color='#e8eaf0',
            xaxis_gridcolor='rgba(255,255,255,0.1)',
            yaxis_gridcolor='rgba(255,255,255,0.1)',
            height=250,
            margin=dict(l=10, r=10, t=30, b=10),
            showlegend=False
        )
        st.plotly_chart(fig, use_container_width=True, key=f"{period}_loss_chart")
    
    # Detail table
    st.markdown("### 📋 Chi tiết dữ liệu")
    display_cols = ['Period', 'Start_Date', 'End_Date', 'Plan', 'Capacity', 'OEE', 'Q.ty defect', 'PPM', 'Loss time', 'UPH']
    display_df = filtered[display_cols].copy()
    display_df.columns = ['Thời gian', 'Từ ngày', 'Đến ngày', 'Kế hoạch', 'Thực tế', 'OEE (%)', 'NG (pcs)', 'PPM', 'Loss (h)', 'UPH']
    st.dataframe(display_df, use_container_width=True, hide_index=True)

def show_by_code(df_sum, df_loss=None):
    st.markdown("<div class='main-header'>📦 Phân tích theo Part Code</div>", unsafe_allow_html=True)
    
    # Time period selection
    period = st.radio("Chọn kỳ:", ["Theo tháng", "Theo tuần", "Theo ngày"], horizontal=True)
    period_key = {'Theo tháng': 'monthly', 'Theo tuần': 'weekly', 'Theo ngày': 'daily'}[period]
    
    # Get unique periods
    if period_key == 'monthly':
        periods = sorted(df_sum['YearMonth'].unique())
    elif period_key == 'weekly':
        periods = sorted(df_sum['YearWeek'].unique())
    else:
        periods = sorted(df_sum['Date_str'].unique())
    
    # Date range
    col1, col2 = st.columns(2)
    with col1:
        from_period = st.selectbox("Từ:", periods, index=0, key=f"code_from_{period_key}")
    with col2:
        to_period = st.selectbox("Đến:", periods, index=len(periods)-1, key=f"code_to_{period_key}")
    
    # Get data
    code_data = get_by_code_data(df_sum, period_key, from_period, to_period)
    
    # KPI Cards - Cao nhất & Thấp nhất
    st.markdown("#### 📊 Thống kê cao nhất")
    cols = st.columns(4)
    with cols[0]:
        st.metric("Tổng Part Code", len(code_data))
    if len(code_data) > 0:
        with cols[1]:
            top_vol = code_data.iloc[0]
            st.metric("🔼 Vol cao nhất", top_vol['Part code'], f"{top_vol['Capacity']:,.0f} pcs")
        with cols[2]:
            top_ng = code_data.loc[code_data['Q.ty defect'].idxmax()]
            st.metric("🔼 NG cao nhất", top_ng['Part code'], f"{top_ng['Q.ty defect']:,.0f} pcs")
        with cols[3]:
            uph_data_valid = code_data[code_data['UPH'] > 0]
            if len(uph_data_valid) > 0:
                top_uph = uph_data_valid.loc[uph_data_valid['UPH'].idxmax()]
                st.metric("🔼 UPH cao nhất", top_uph['Part code'], f"{top_uph['UPH']:.2f} pcs/h")
    
    # KPI Cards - NG Rate cao nhất & thấp nhất
    st.markdown("#### 📊 NG Rate (tỷ lệ phần trăm)")
    cols = st.columns(4)
    ng_positive = code_data[code_data['Capacity'] > 0]
    if len(ng_positive) > 0:
        with cols[0]:
            top_ng_rate = ng_positive.loc[ng_positive['NG_Rate'].idxmax()]
            st.metric("🔼 NG Rate cao nhất", top_ng_rate['Part code'], f"{top_ng_rate['NG_Rate']:.2f}%")
        with cols[1]:
            # NG Rate thấp nhất (có sản lượng > 0 và NG > 0)
            ng_with_defects = ng_positive[ng_positive['Q.ty defect'] > 0]
            if len(ng_with_defects) > 0:
                low_ng_rate = ng_with_defects.loc[ng_with_defects['NG_Rate'].idxmin()]
                st.metric("🔽 NG Rate thấp nhất", low_ng_rate['Part code'], f"{low_ng_rate['NG_Rate']:.2f}%")
    
    # KPI Cards - Thấp nhất
    st.markdown("#### 📉 Thống kê thấp nhất (có sản lượng > 0)")
    cols = st.columns(4)
    # Lọc các part có sản lượng > 0
    code_data_with_vol = code_data[code_data['Capacity'] > 0]
    if len(code_data_with_vol) > 0:
        with cols[0]:
            low_vol = code_data_with_vol.iloc[-1]
            st.metric("🔽 Vol thấp nhất", low_vol['Part code'], f"{low_vol['Capacity']:,.0f} pcs")
        with cols[1]:
            # NG thấp nhất (có sản lượng > 0)
            ng_positive = code_data_with_vol[code_data_with_vol['Q.ty defect'] >= 0]
            if len(ng_positive) > 0:
                low_ng = ng_positive.loc[ng_positive['Q.ty defect'].idxmin()]
                st.metric("🔽 NG thấp nhất", low_ng['Part code'], f"{low_ng['Q.ty defect']:,.0f} pcs")
        with cols[2]:
            # UPH thấp nhất (UPH > 0)
            uph_positive = code_data_with_vol[code_data_with_vol['UPH'] > 0]
            if len(uph_positive) > 0:
                low_uph = uph_positive.loc[uph_positive['UPH'].idxmin()]
                st.metric("🔽 UPH thấp nhất", low_uph['Part code'], f"{low_uph['UPH']:.2f} pcs/h")
    
    st.markdown("---")
    
    # ========== CHARTS CAO NHẤT (4 chart trên 1 hàng) ==========
    st.markdown("#### 📈 Charts - Cao nhất")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        # Top 10 Volume - code_data đã sắp xếp theo Capacity giảm dần
        top10_vol = code_data.head(10).sort_values('Capacity', ascending=True).copy()
        top10_vol['Part code'] = top10_vol['Part code'].apply(ensure_part_code_string)
        fig = px.bar(top10_vol, y='Part code', x='Capacity', orientation='h', 
                     title="🔼 Vol cao nhất", color_discrete_sequence=['#2dd4a0'])
        fig.update_layout(plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0', 
                          height=300, margin=dict(l=10, r=10, t=30, b=10), 
                          yaxis=dict(type='category', tickmode='linear'))
        st.plotly_chart(fig, use_container_width=True, key=f"code_vol_top_{period_key}")
    
    with col2:
        # Top 10 UPH - Sắp xếp giảm dần, hiển thị từ trên xuống
        uph_data_top = code_data[code_data['UPH'] > 0].nlargest(10, 'UPH').sort_values('UPH', ascending=True).copy()
        uph_data_top['Part code'] = uph_data_top['Part code'].apply(ensure_part_code_string)
        if len(uph_data_top) > 0:
            fig = px.bar(uph_data_top, y='Part code', x='UPH', orientation='h',
                         title="🔼 UPH cao nhất", color_discrete_sequence=['#3b8eea'])
            fig.update_layout(plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0',
                              height=300, margin=dict(l=10, r=10, t=30, b=10),
                              yaxis=dict(type='category', tickmode='linear'))
            st.plotly_chart(fig, use_container_width=True, key=f"code_uph_top_{period_key}")
    
    with col3:
        # Top 10 NG - Sắp xếp giảm dần, hiển thị từ trên xuống
        ng_data_top = code_data[code_data['Q.ty defect'] > 0].nlargest(10, 'Q.ty defect').sort_values('Q.ty defect', ascending=True).copy()
        ng_data_top['Part code'] = ng_data_top['Part code'].apply(ensure_part_code_string)
        if len(ng_data_top) > 0:
            fig = px.bar(ng_data_top, y='Part code', x='Q.ty defect', orientation='h',
                         title="🔼 NG cao nhất", color_discrete_sequence=['#f05c5c'])
            fig.update_layout(plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0',
                              height=300, margin=dict(l=10, r=10, t=30, b=10),
                              yaxis=dict(type='category', tickmode='linear'))
            st.plotly_chart(fig, use_container_width=True, key=f"code_ng_top_{period_key}")
    
    with col4:
        # Top 10 NG Rate (%) - Sắp xếp giảm dần
        ngrate_data_top = code_data[code_data['Capacity'] > 0].nlargest(10, 'NG_Rate').sort_values('NG_Rate', ascending=True).copy()
        ngrate_data_top['Part code'] = ngrate_data_top['Part code'].apply(ensure_part_code_string)
        if len(ngrate_data_top) > 0:
            fig = px.bar(ngrate_data_top, y='Part code', x='NG_Rate', orientation='h',
                         title="🔼 NG Rate cao nhất (%)", color_discrete_sequence=['#ff6b6b'])
            fig.update_layout(plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0',
                              height=300, margin=dict(l=10, r=10, t=30, b=10),
                              yaxis=dict(type='category', tickmode='linear'))
            st.plotly_chart(fig, use_container_width=True, key=f"code_ngrate_top_{period_key}")
    
    # ========== CHARTS THẤP NHẤT (4 chart trên 1 hàng) ==========
    st.markdown("#### 📉 Charts - Thấp nhất")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        # Low 10 Volume - Sắp xếp theo Capacity tăng dần (thấp nhất ở trên)
        vol_positive = code_data[code_data['Capacity'] > 0].copy()
        if len(vol_positive) > 0:
            low10_vol = vol_positive.nsmallest(10, 'Capacity').sort_values('Capacity', ascending=True).copy()
            low10_vol['Part code'] = low10_vol['Part code'].apply(ensure_part_code_string)
            fig = px.bar(low10_vol, y='Part code', x='Capacity', orientation='h',
                         title="🔽 Vol thấp nhất", color_discrete_sequence=['#f0a832'])
            fig.update_layout(plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0',
                              height=300, margin=dict(l=10, r=10, t=30, b=10),
                              yaxis=dict(type='category', tickmode='linear'))
            st.plotly_chart(fig, use_container_width=True, key=f"code_vol_low_{period_key}")
    
    with col2:
        # Low 10 UPH - Sắp xếp theo UPH tăng dần (thấp nhất ở trên)
        uph_positive = code_data[code_data['UPH'] > 0].copy()
        if len(uph_positive) > 0:
            uph_data_low = uph_positive.nsmallest(10, 'UPH').sort_values('UPH', ascending=True).copy()
            uph_data_low['Part code'] = uph_data_low['Part code'].apply(ensure_part_code_string)
            fig = px.bar(uph_data_low, y='Part code', x='UPH', orientation='h',
                         title="🔽 UPH thấp nhất", color_discrete_sequence=['#38c4d4'])
            fig.update_layout(plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0',
                              height=300, margin=dict(l=10, r=10, t=30, b=10),
                              yaxis=dict(type='category', tickmode='linear'))
            st.plotly_chart(fig, use_container_width=True, key=f"code_uph_low_{period_key}")
    
    with col3:
        # Low 10 NG - Sắp xếp theo NG tăng dần (thấp nhất ở trên)
        ng_positive = code_data[code_data['Q.ty defect'] > 0].copy()
        if len(ng_positive) > 0:
            ng_data_low = ng_positive.nsmallest(10, 'Q.ty defect').sort_values('Q.ty defect', ascending=True).copy()
            ng_data_low['Part code'] = ng_data_low['Part code'].apply(ensure_part_code_string)
            fig = px.bar(ng_data_low, y='Part code', x='Q.ty defect', orientation='h',
                         title="🔽 NG thấp nhất", color_discrete_sequence=['#f0a832'])
            fig.update_layout(plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0',
                              height=300, margin=dict(l=10, r=10, t=30, b=10),
                              yaxis=dict(type='category', tickmode='linear'))
            st.plotly_chart(fig, use_container_width=True, key=f"code_ng_low_{period_key}")
    
    with col4:
        # Low 10 NG Rate (%) - Sắp xếp theo NG Rate tăng dần (có NG > 0)
        ng_with_defects = code_data[(code_data['Capacity'] > 0) & (code_data['Q.ty defect'] > 0)].copy()
        if len(ng_with_defects) > 0:
            ngrate_data_low = ng_with_defects.nsmallest(10, 'NG_Rate').sort_values('NG_Rate', ascending=True).copy()
            ngrate_data_low['Part code'] = ngrate_data_low['Part code'].apply(ensure_part_code_string)
            fig = px.bar(ngrate_data_low, y='Part code', x='NG_Rate', orientation='h',
                         title="🔽 NG Rate thấp nhất (%)", color_discrete_sequence=['#2dd4a0'])
            fig.update_layout(plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0',
                              height=300, margin=dict(l=10, r=10, t=30, b=10),
                              yaxis=dict(type='category', tickmode='linear'))
            st.plotly_chart(fig, use_container_width=True, key=f"code_ngrate_low_{period_key}")
    
    # Search and table
    st.markdown("### 📋 Chi tiết toàn bộ Part Code")
    search = st.text_input("🔍 Tìm kiếm Part Code hoặc Part Name:", "", key=f"code_search_{period_key}")
    
    if search:
        filtered = code_data[
            code_data['Part code'].str.contains(search, case=False, na=False) |
            code_data['Part name'].str.contains(search, case=False, na=False)
        ]
    else:
        filtered = code_data
    
    display_df = filtered[['Part code', 'Part name', 'Plan', 'Capacity', 'OEE', 'Q.ty defect', 'PPM', 'NG_Rate', 'Loss time', 'UPH']].copy()
    display_df.columns = ['Part Code', 'Part Name', 'Plan', 'Actual', 'OEE (%)', 'NG', 'PPM', 'NG Rate (%)', 'Loss (h)', 'UPH']
    display_df['Part Code'] = display_df['Part Code'].apply(ensure_part_code_string)
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    
    # Chi tiết Loss Time với Start/End Time cho Part Code được chọn
    if df_loss is not None and len(filtered) > 0 and search:
        with st.expander("⏱️ Xem chi tiết Loss Time (Start/End Time)", expanded=False):
            selected_codes = filtered['Part code'].tolist()
            loss_details = df_loss[df_loss['Part code'].isin(selected_codes)][
                ['Date', 'Part code', 'Part name', 'Loss type', 'Reason', 'Dept PIC', 'Start time', 'End time', 'Loss time']
            ].sort_values('Loss time', ascending=False)
            
            if len(loss_details) > 0:
                loss_details['Date'] = pd.to_datetime(loss_details['Date']).dt.strftime('%d/%m/%Y')
                loss_details.columns = ['Ngày', 'Part Code', 'Part Name', 'Loại', 'Nguyên nhân', 'Bộ phận', 'Bắt đầu', 'Kết thúc', 'Thởi gian (h)']
                loss_details['Part Code'] = loss_details['Part Code'].apply(ensure_part_code_string)
                st.dataframe(loss_details, use_container_width=True, hide_index=True)
            else:
                st.info("Không có dữ liệu Loss Time cho Part Code này")

def show_by_shift(df_sum):
    st.markdown("<div class='main-header'>🕐 Phân tích theo ca</div>", unsafe_allow_html=True)
    
    # Time period selection
    period = st.radio("Chọn kỳ:", ["Theo tháng", "Theo tuần", "Theo ngày"], horizontal=True)
    period_key = {'Theo tháng': 'monthly', 'Theo tuần': 'weekly', 'Theo ngày': 'daily'}[period]
    
    # Get unique periods
    if period_key == 'monthly':
        periods = sorted(df_sum['YearMonth'].unique())
    elif period_key == 'weekly':
        periods = sorted(df_sum['YearWeek'].unique())
    else:
        periods = sorted(df_sum['Date_str'].unique())
    
    # Date range
    col1, col2 = st.columns(2)
    with col1:
        from_period = st.selectbox("Từ:", periods, index=0)
    with col2:
        to_period = st.selectbox("Đến:", periods, index=len(periods)-1)
    
    # Get shift data
    shift_data = get_shift_data(df_sum, period_key, from_period, to_period)
    
    # Aggregate for display
    day_data = shift_data[shift_data['Shift'] == 'Ngày'].groupby('Shift').agg({
        'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum', 'Loss time': 'sum', 'Operating time': 'sum'
    }).reset_index()
    
    night_data = shift_data[shift_data['Shift'] == 'Đêm'].groupby('Shift').agg({
        'Plan': 'sum', 'Capacity': 'sum', 'Q.ty defect': 'sum', 'Loss time': 'sum', 'Operating time': 'sum'
    }).reset_index()
    
    # Calculate metrics
    if len(day_data) > 0:
        day_data['OEE'] = day_data.apply(lambda x: round(x['Capacity'] / x['Plan'] * 100, 2) if x['Plan'] > 0 else 0, axis=1)
        day_data['UPH'] = day_data.apply(lambda x: calculate_uph(x['Capacity'], x['Operating time']), axis=1)
    if len(night_data) > 0:
        night_data['OEE'] = night_data.apply(lambda x: round(x['Capacity'] / x['Plan'] * 100, 2) if x['Plan'] > 0 else 0, axis=1)
        night_data['UPH'] = night_data.apply(lambda x: calculate_uph(x['Capacity'], x['Operating time']), axis=1)
    
    # Display cards
    cols = st.columns(2)
    with cols[0]:
        st.markdown("#### ☀️ Ca Ngày")
        if len(day_data) > 0:
            d = day_data.iloc[0]
            st.metric("Sản lượng", f"{d['Capacity']:,.0f}")
            st.metric("Đạt KH", f"{d['OEE']:.2f}%")
            st.metric("Lỗi NG", f"{d['Q.ty defect']:,.0f}")
            st.metric("Loss time", f"{d['Loss time']:.1f} h")
    with cols[1]:
        st.markdown("#### 🌙 Ca Đêm")
        if len(night_data) > 0:
            d = night_data.iloc[0]
            st.metric("Sản lượng", f"{d['Capacity']:,.0f}")
            st.metric("Đạt KH", f"{d['OEE']:.2f}%")
            st.metric("Lỗi NG", f"{d['Q.ty defect']:,.0f}")
            st.metric("Loss time", f"{d['Loss time']:.1f} h")
    
    st.markdown("---")
    
    # Charts
    col1, col2 = st.columns(2)
    with col1:
        # Capacity comparison
        cap_data = []
        if len(day_data) > 0:
            cap_data.append({'Ca': 'Ca Ngày', 'Sản lượng': day_data.iloc[0]['Capacity']})
        if len(night_data) > 0:
            cap_data.append({'Ca': 'Ca Đêm', 'Sản lượng': night_data.iloc[0]['Capacity']})
        if cap_data:
            fig = px.bar(cap_data, x='Ca', y='Sản lượng', color='Ca',
                         color_discrete_map={'Ca Ngày': '#f0a832', 'Ca Đêm': '#9b7de8'})
            fig.update_layout(title="Sản lượng theo ca", plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0')
            st.plotly_chart(fig, use_container_width=True, key=f"shift_cap_{period_key}")
    
    with col2:
        # NG comparison
        ng_data = []
        if len(day_data) > 0:
            ng_data.append({'Ca': 'Ca Ngày', 'NG': day_data.iloc[0]['Q.ty defect']})
        if len(night_data) > 0:
            ng_data.append({'Ca': 'Ca Đêm', 'NG': night_data.iloc[0]['Q.ty defect']})
        if ng_data:
            fig = px.bar(ng_data, x='Ca', y='NG', color='Ca',
                         color_discrete_map={'Ca Ngày': '#f0a832', 'Ca Đêm': '#9b7de8'})
            fig.update_layout(title="NG theo ca", plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0')
            st.plotly_chart(fig, use_container_width=True, key=f"shift_ng_{period_key}")
    
    # Pie charts
    col1, col2, col3 = st.columns(3)
    with col1:
        pie_data = []
        if len(day_data) > 0:
            pie_data.append({'Ca': 'Ca Ngày', 'Value': day_data.iloc[0]['Capacity']})
        if len(night_data) > 0:
            pie_data.append({'Ca': 'Ca Đêm', 'Value': night_data.iloc[0]['Capacity']})
        if pie_data:
            fig = px.pie(pie_data, values='Value', names='Ca', title="Sản lượng phân bổ",
                         color_discrete_map={'Ca Ngày': '#f0a832', 'Ca Đêm': '#9b7de8'})
            fig.update_layout(paper_bgcolor='#0d0f14', font_color='#e8eaf0')
            st.plotly_chart(fig, use_container_width=True, key=f"shift_pie_cap_{period_key}")
    
    with col2:
        pie_data = []
        if len(day_data) > 0:
            pie_data.append({'Ca': 'Ca Ngày', 'Value': day_data.iloc[0]['Q.ty defect']})
        if len(night_data) > 0:
            pie_data.append({'Ca': 'Ca Đêm', 'Value': night_data.iloc[0]['Q.ty defect']})
        if pie_data:
            fig = px.pie(pie_data, values='Value', names='Ca', title="NG phân bổ",
                         color_discrete_map={'Ca Ngày': '#f0a832', 'Ca Đêm': '#9b7de8'})
            fig.update_layout(paper_bgcolor='#0d0f14', font_color='#e8eaf0')
            st.plotly_chart(fig, use_container_width=True, key=f"shift_pie_ng_{period_key}")
    
    with col3:
        pie_data = []
        if len(day_data) > 0:
            pie_data.append({'Ca': 'Ca Ngày', 'Value': day_data.iloc[0]['Loss time']})
        if len(night_data) > 0:
            pie_data.append({'Ca': 'Ca Đêm', 'Value': night_data.iloc[0]['Loss time']})
        if pie_data:
            fig = px.pie(pie_data, values='Value', names='Ca', title="Loss time phân bổ",
                         color_discrete_map={'Ca Ngày': '#f0a832', 'Ca Đêm': '#9b7de8'})
            fig.update_layout(paper_bgcolor='#0d0f14', font_color='#e8eaf0')
            st.plotly_chart(fig, use_container_width=True, key=f"shift_pie_loss_{period_key}")
    
    # Comparison table
    st.markdown("### Bảng so sánh chi tiết")
    compare_data = []
    if len(day_data) > 0:
        d = day_data.iloc[0]
        compare_data.append(['☀️ Ca Ngày', d['Plan'], d['Capacity'], d['OEE'], d['Q.ty defect'], d['Loss time'], d['UPH']])
    if len(night_data) > 0:
        d = night_data.iloc[0]
        compare_data.append(['🌙 Ca Đêm', d['Plan'], d['Capacity'], d['OEE'], d['Q.ty defect'], d['Loss time'], d['UPH']])
    
    if compare_data:
        compare_df = pd.DataFrame(compare_data, columns=['Ca', 'Kế hoạch', 'Thực tế', 'OEE (%)', 'NG (pcs)', 'Loss (h)', 'UPH'])
        st.dataframe(compare_df, use_container_width=True, hide_index=True)

def show_loss_time_view(df_loss, period):
    """Show loss time view with period filtering"""
    # Get unique periods
    if period == 'monthly':
        periods = sorted(df_loss['YearMonth'].unique())
    elif period == 'weekly':
        periods = sorted(df_loss['YearWeek'].unique())
    else:
        periods = sorted(df_loss['Date_str'].unique())
    
    if len(periods) == 0:
        st.warning("Không có dữ liệu Loss Time cho kỳ này")
        return
    
    # Date range filter
    col1, col2 = st.columns(2)
    with col1:
        from_period = st.selectbox("Từ:", periods, index=0, key=f"loss_from_{period}")
    with col2:
        to_period = st.selectbox("Đến:", periods, index=len(periods)-1, key=f"loss_to_{period}")
    
    # Get filtered loss data
    loss_reason, loss_dept, loss_type, loss_by_code, loss_plan_type, loss_details, total_loss = \
        get_loss_data(df_loss, period, from_period, to_period)
    
    # KPI Cards - Phân loại theo kế hoạch
    cols = st.columns(5)
    with cols[0]:
        st.metric("📊 Tổng loss time", f"{total_loss:.1f}", "giờ")
    
    # Phân loại Có kế hoạch vs Không kế hoạch
    planned_loss = loss_plan_type[loss_plan_type['Loss type'] == 'Có kế hoạch']['Loss time'].sum() if len(loss_plan_type) > 0 else 0
    unplanned_loss = loss_plan_type[loss_plan_type['Loss type'] == 'Không kế hoạch']['Loss time'].sum() if len(loss_plan_type) > 0 else 0
    
    with cols[1]:
        st.metric("📋 Có kế hoạch", f"{planned_loss:.1f}h", f"{planned_loss/total_loss*100:.1f}%" if total_loss > 0 else "0%")
    with cols[2]:
        st.metric("⚠️ Không kế hoạch", f"{unplanned_loss:.1f}h", f"{unplanned_loss/total_loss*100:.1f}%" if total_loss > 0 else "0%")
    with cols[3]:
        if len(loss_reason) > 0:
            st.metric("🔴 Nguyên nhân #1", loss_reason.iloc[0]['Reason'], f"{loss_reason.iloc[0]['Loss time']:.1f} giờ")
    with cols[4]:
        if len(loss_dept) > 0:
            st.metric("🏭 Bộ phận #1", loss_dept.iloc[0]['Dept PIC'], f"{loss_dept.iloc[0]['Loss time']:.1f} giờ")
    
    st.markdown("---")
    
    # Charts - 4 chart trên 1 hàng
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        fig = px.bar(loss_reason, y='Reason', x='Loss time', orientation='h',
                     title="🔴 Top nguyên nhân", color_discrete_sequence=['#f05c5c'])
        fig.update_layout(plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0',
                          height=280, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig, use_container_width=True, key=f"loss_reason_chart_{period}")
    
    with col2:
        fig = px.pie(loss_plan_type, values='Loss time', names='Loss type', 
                     title="📊 Có/Không kế hoạch",
                     color_discrete_map={'Có kế hoạch': '#3b8eea', 'Không kế hoạch': '#f05c5c'})
        fig.update_layout(paper_bgcolor='#0d0f14', font_color='#e8eaf0',
                          height=280, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig, use_container_width=True, key=f"loss_planned_chart_{period}")
    
    with col3:
        fig = px.bar(loss_dept, y='Dept PIC', x='Loss time', orientation='h',
                     title="🏭 Theo bộ phận", color_discrete_sequence=['#9b7de8'])
        fig.update_layout(plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0',
                          height=280, margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig, use_container_width=True, key=f"loss_dept_chart_{period}")
    
    with col4:
        loss_code_chart = loss_by_code.head(10).copy()
        loss_code_chart['Part code'] = loss_code_chart['Part code'].apply(ensure_part_code_string)
        fig = px.bar(loss_code_chart, y='Part code', x='Loss time', orientation='h',
                     title="📦 Theo Part Code", color_discrete_sequence=['#f0a832'])
        fig.update_layout(plot_bgcolor='#141720', paper_bgcolor='#0d0f14', font_color='#e8eaf0',
                          height=280, margin=dict(l=10, r=10, t=30, b=10), 
                          yaxis=dict(type='category', tickmode='linear'))
        st.plotly_chart(fig, use_container_width=True, key=f"loss_code_chart_{period}")
    
    # Tables
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 🔍 Chi tiết nguyên nhân")
        if len(loss_reason) > 0:
            loss_reason['Tỉ lệ (%)'] = (loss_reason['Loss time'] / loss_reason['Loss time'].sum() * 100).round(1)
            st.dataframe(loss_reason[['Reason', 'Loss time', 'Tỉ lệ (%)']], use_container_width=True, hide_index=True)
    
    with col2:
        st.markdown("### 📦 Chi tiết Part Code bị ảnh hưởng")
        if len(loss_by_code) > 0:
            loss_by_code['Tỉ lệ (%)'] = (loss_by_code['Loss time'] / loss_by_code['Loss time'].sum() * 100).round(1)
            display_loss_code = loss_by_code[['Part code', 'Part name', 'Loss time', 'Tỉ lệ (%)', 'Top_Reason']].copy()
            display_loss_code.columns = ['Part Code', 'Part Name', 'Loss (h)', 'Tỉ lệ (%)', 'Lý do chính']
            display_loss_code['Part Code'] = display_loss_code['Part Code'].apply(ensure_part_code_string)
            st.dataframe(display_loss_code, use_container_width=True, hide_index=True)
    
    # Chi tiết Loss Time với Start/End Time
    st.markdown("---")
    st.markdown("### ⏱️ Chi tiết Loss Time (Top 50 sự cố)")
    
    # Format datetime columns
    if len(loss_details) > 0:
        display_details = loss_details.copy()
        display_details['Date'] = pd.to_datetime(display_details['Date']).dt.strftime('%d/%m/%Y')
        display_details.columns = ['Ngày', 'Part Code', 'Part Name', 'Loại', 'Nguyên nhân', 'Bộ phận', 'Bắt đầu', 'Kết thúc', 'Thởi gian (h)']
        display_details['Part Code'] = display_details['Part Code'].apply(ensure_part_code_string)
        st.dataframe(display_details, use_container_width=True, hide_index=True)
    else:
        st.info("Không có dữ liệu chi tiết")

def show_loss_time(df_sum, df_loss):
    st.markdown("<div class='main-header'>⚠️ Phân tích Loss Time</div>", unsafe_allow_html=True)
    
    # Time period selection tabs
    tab_monthly, tab_weekly, tab_daily = st.tabs(["📅 Theo tháng", "📆 Theo tuần", "📋 Theo ngày"])
    
    with tab_monthly:
        show_loss_time_view(df_loss, 'monthly')
    with tab_weekly:
        show_loss_time_view(df_loss, 'weekly')
    with tab_daily:
        show_loss_time_view(df_loss, 'daily')

if __name__ == "__main__":
    main()
