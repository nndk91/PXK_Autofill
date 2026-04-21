"""
app_v5.py - PXK Manager v5 (Sử dụng core_v4)
===========================================
Cải tiến:
- Cho phép chọn chế độ: Chỉ trích xuất PDF (bước 1) hoặc Trích xuất + Ghép form (cả 2 bước)
- Sử dụng pxk_core_v4 đã được tối ưu
- Giao diện trực quan với Streamlit

Các folder dữ liệu huấn luyện:
- 2033-2096, 2144-2172, 2168-2243, 2273-2316
"""
import io
import os
import tempfile
from collections import defaultdict
from pathlib import Path

import streamlit as st
import pandas as pd

# Import từ pxk_core_v4
from pxk_core_v4 import (
    extract_pdfs_from_files,
    read_form_rows_from_bytes,
    match_pxk_v4,
    build_output_excel,
    load_reference_scorer,
    pxk_sort_key,
)

# Import pdf_extractor để trích xuất riêng
from pdf_extractor import extract_pxk

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title='PXK Manager v5',
    page_icon='📦',
    layout='wide',
    initial_sidebar_state='expanded',
)

st.markdown('''
<style>
.step-header { font-size: 1.1rem; font-weight: 700; color: #1F4E79; margin-bottom: 6px; }
.tag-green  { background: #C6EFCE; color: #276221; padding: 3px 10px; border-radius: 4px; font-size: .9rem; }
.tag-yellow { background: #FFEB9C; color: #7d6608; padding: 3px 10px; border-radius: 4px; font-size: .9rem; }
.tag-red    { background: #FFC7CE; color: #9c1f23; padding: 3px 10px; border-radius: 4px; font-size: .9rem; }
.tag-blue   { background: #BDD7EE; color: #1F4E79; padding: 3px 10px; border-radius: 4px; font-size: .9rem; }
.mode-box { 
    background: #f0f2f6; 
    padding: 15px; 
    border-radius: 8px; 
    border-left: 4px solid #1F4E79;
    margin-bottom: 10px;
}
</style>
''', unsafe_allow_html=True)

st.title('📦 PXK Manager v5 - Trích xuất & Ghép số PXK')
st.caption('🧠 Sử dụng AI học từ dữ liệu đã điền | Tự động hóa thông minh')

# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR - MODE SELECTION & UPLOAD
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header('⚙️ Cấu hình xử lý')
    
    # ── Chế độ xử lý ───────────────────────────────────────────────────────────
    st.markdown('<div class="step-header">Chọn chế độ xử lý</div>', unsafe_allow_html=True)
    
    process_mode = st.radio(
        'Chế độ:',
        options=['full', 'extract_only'],
        format_func=lambda x: {
            'full': '📋 Cả 2 bước: Trích xuất PDF + Ghép vào Form',
            'extract_only': '📄 Chỉ bước 1: Trích xuất dữ liệu PDF',
        }[x],
        help='Chọn chế độ phù hợp với nhu cầu của bạn',
    )
    
    if process_mode == 'full':
        st.markdown('''
        <div class="mode-box">
        <b>Chế độ đầy đủ:</b><br>
        1. Trích xuất dữ liệu từ các file PDF<br>
        2. Ghép số PXK vào form trống
        </div>
        ''', unsafe_allow_html=True)
    else:
        st.markdown('''
        <div class="mode-box">
        <b>Chế độ trích xuất:</b><br>
        Chỉ trích xuất dữ liệu từ PDF<br>
        (Không cần upload form)
        </div>
        ''', unsafe_allow_html=True)
    
    st.divider()
    
    # ── Bước 1: Upload PDF ────────────────────────────────────────────────────
    st.markdown('<div class="step-header">Bước 1 - Upload PDF PXK</div>', unsafe_allow_html=True)
    pdf_files = st.file_uploader(
        'Chọn file PDF Phiếu Xuất Kho',
        type=['pdf'],
        accept_multiple_files=True,
        key='pdf_upload',
        help='Có thể chọn nhiều file (Ctrl+Click)',
    )
    
    if pdf_files:
        st.success(f'✅ {len(pdf_files)} file PDF đã chọn')
    
    # ── Bước 2: Upload Form (chỉ khi mode=full) ───────────────────────────────
    form_file = None
    if process_mode == 'full':
        st.divider()
        st.markdown('<div class="step-header">Bước 2 - Upload Form trống</div>', unsafe_allow_html=True)
        form_file = st.file_uploader(
            'FORM DỮ LIỆU CHƯA NHẬP SỐ PXK.xlsx',
            type=['xlsx'],
            key='form_upload',
        )
        if form_file:
            st.success(f'✅ {form_file.name}')
    
    st.divider()
    
    # ── Nút xử lý ─────────────────────────────────────────────────────────────
    can_run = False
    if process_mode == 'extract_only':
        can_run = bool(pdf_files)
    else:
        can_run = bool(pdf_files and form_file)
    
    run_btn = st.button(
        '🚀 Bắt đầu xử lý',
        type='primary',
        disabled=not can_run,
        use_container_width=True,
    )
    
    st.divider()
    
    # ── Hướng dẫn ─────────────────────────────────────────────────────────────
    st.caption('💡 Màu sắc kết quả:')
    st.markdown('<span class="tag-green">✅ Tự động</span> — 1 PXK duy nhất khớp', unsafe_allow_html=True)
    st.markdown('<span class="tag-yellow">🔍 Cần kiểm tra</span> — Có nhiều khả năng', unsafe_allow_html=True)
    st.markdown('<span class="tag-red">❌ Không khớp</span> — Cần điền thủ công', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN UI - INSTRUCTIONS (when not ready)
# ═══════════════════════════════════════════════════════════════════════════════
if not can_run:
    st.info('👈 Vui lòng upload file ở thanh bên trái để bắt đầu')
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader('📄 Chế độ Trích xuất')
        st.markdown('''
        **Phù hợp khi:**
        - Chỉ cần xem dữ liệu từ PDF
        - Chưa có form để ghép
        - Muốn kiểm tra nội dung PDF trước
        
        **Kết quả:**
        - Bảng dữ liệu trích xuất
        - File Excel tải về
        ''')
    
    with col2:
        st.subheader('📋 Chế độ Đầy đủ')
        st.markdown('''
        **Phù hợp khi:**
        - Cần ghép số PXK vào form
        - Có cả PDF và form trống
        - Muốn xử lý tự động hoàn toàn
        
        **Kết quả:**
        - Form đã điền số PXK
        - Thống kê ghép thành công
        ''')
    
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════════
# PROCESSING
# ═══════════════════════════════════════════════════════════════════════════════
if run_btn or st.session_state.get('processed'):
    if run_btn:
        st.session_state.pop('processed', None)
        st.session_state.pop('results_cache', None)
    
    if 'results_cache' not in st.session_state:
        with st.status('⏳ Đang xử lý...', expanded=True) as status_bar:
            
            # ── Bước 1: Trích xuất PDF ──────────────────────────────────────────
            st.write(f'📄 Đang trích xuất {len(pdf_files)} file PDF...')
            
            file_bytes = [(f.name, f.read()) for f in pdf_files]
            pxk_totals, pxk_items, pxk_dates, pxk_do_no, pdf_errors = extract_pdfs_from_files(file_bytes)
            
            do_covered = sum(1 for p in pxk_totals if p in pxk_do_no)
            st.write(f'✅ {len(pxk_totals)} PXK trích xuất thành công ({do_covered} có D/O No)')
            
            if pdf_errors:
                st.write(f'⚠️ {len(pdf_errors)} file lỗi')
            
            # ── Mode: Extract Only ─────────────────────────────────────────────
            if process_mode == 'extract_only':
                st.write('✅ Hoàn tất trích xuất!')
                
                st.session_state['results_cache'] = {
                    'mode': 'extract_only',
                    'pxk_totals': pxk_totals,
                    'pxk_items': pxk_items,  # Lưu chi tiết từng dòng riêng biệt
                    'pxk_dates': pxk_dates,
                    'pxk_do_no': pxk_do_no,
                    'pdf_errors': pdf_errors,
                }
                st.session_state['processed'] = True
                status_bar.update(label='✅ Trích xuất hoàn tất!', state='complete')
            
            # ── Mode: Full (Extract + Match) ───────────────────────────────────
            else:
                # ── Bước 2: Load dữ liệu huấn luyện ─────────────────────────────
                st.write('📚 Đang nạp dữ liệu huấn luyện từ các folder...')
                scorer = load_reference_scorer('.')
                
                # Hiển thị chi tiết dữ liệu học
                if scorer.folder_count > 0:
                    st.success(
                        f'✅ Đã học từ **{scorer.folder_count}** folder | '
                        f'**{scorer.example_count}** mẫu | '
                        f'**{len(scorer.inv_to_pxk_counter)}** invoice patterns | '
                        f'**{len(scorer.item_to_pxk_counter)}** item patterns'
                    )
                else:
                    st.warning('⚠️ Không tìm thấy folder dữ liệu huấn luyện. Chương trình sẽ chạy không có AI hỗ trợ.')
                
                # ── Bước 3: Đọc form ────────────────────────────────────────────
                st.write('📊 Đang đọc FORM CHƯA NHẬP...')
                form_bytes = form_file.read()
                form_rows = read_form_rows_from_bytes(form_bytes)
                st.write(f'✅ {len(form_rows)} dòng dữ liệu cần ghép')
                
                # ── Bước 4: Ghép PXK ────────────────────────────────────────────
                st.write('🧠 Đang ghép số PXK (sử dụng AI học từ dữ liệu)...')
                result, status_list, note_pxks = match_pxk_v4(
                    form_rows, pxk_totals, pxk_do_no, scorer=scorer
                )
                
                n_auto = sum(1 for s in status_list if s == 'auto')
                n_amb = sum(1 for s in status_list if s == 'ambiguous')
                n_none = sum(1 for s in status_list if s == 'no_match')
                st.write(f'✅ Tự động: **{n_auto}** | Cần KT: **{n_amb}** | Không khớp: **{n_none}**')
                
                # ── Bước 5: Tạo Excel ───────────────────────────────────────────
                st.write('💾 Đang tạo file Excel kết quả...')
                output_bytes = build_output_excel(
                    form_bytes, form_rows, result, status_list, note_pxks, pxk_dates
                )
                
                st.session_state['results_cache'] = {
                    'mode': 'full',
                    'pxk_totals': pxk_totals,
                    'pxk_items': pxk_items,  # Lưu chi tiết từng dòng riêng biệt
                    'pxk_dates': pxk_dates,
                    'pxk_do_no': pxk_do_no,
                    'pdf_errors': pdf_errors,
                    'form_rows': form_rows,
                    'result': result,
                    'status_list': status_list,
                    'note_pxks': note_pxks,
                    'output_bytes': output_bytes,
                    'scorer': scorer,
                }
                st.session_state['processed'] = True
                status_bar.update(label='✅ Xử lý hoàn tất!', state='complete')
    
    # ═══════════════════════════════════════════════════════════════════════════
    # DISPLAY RESULTS
    # ═══════════════════════════════════════════════════════════════════════════
    cache = st.session_state['results_cache']
    mode = cache.get('mode', 'extract_only')
    
    if mode == 'extract_only':
        # ── Hiển thị kết quả trích xuất ───────────────────────────────────────
        st.subheader('📄 Kết quả trích xuất PDF')
        
        pxk_totals = cache['pxk_totals']
        pxk_dates = cache['pxk_dates']
        pxk_do_no = cache['pxk_do_no']
        pdf_errors = cache['pdf_errors']
        
        # KPIs
        c1, c2, c3 = st.columns(3)
        c1.metric('📄 Số PXK', len(pxk_totals))
        c2.metric('📅 Có ngày tháng', len(pxk_dates))
        c3.metric('📝 Có D/O No', len(pxk_do_no))
        
        if pdf_errors:
            with st.expander(f'⚠️ {len(pdf_errors)} file lỗi'):
                st.dataframe(pd.DataFrame(pdf_errors), hide_index=True)
        
        # Bảng chi tiết
        st.divider()
        st.subheader('📋 Chi tiết dữ liệu trích xuất')
        
        # Cho phép chọn chế độ xem
        view_mode = st.radio(
            'Chế độ xem:',
            ['chi_tiet', 'gop'],
            format_func=lambda x: '🔍 Xem từng dòng riêng biệt' if x == 'chi_tiet' else '📊 Xem gộp theo mã hàng',
            horizontal=True
        )
        
        # Lấy pxk_items từ cache
        pxk_items = cache.get('pxk_items', {})
        
        if view_mode == 'chi_tiet' and pxk_items:
            # Hiển thị từng dòng riêng biệt (KHÔNG gộp)
            display_data = []
            for pxk, items in pxk_items.items():
                for item in items:
                    display_data.append({
                        'Số PXK': pxk,
                        'Ngày': pxk_dates.get(pxk, ''),
                        'D/O No': ', '.join(pxk_do_no.get(pxk, [])),
                        'Dòng': item.get('line_no', ''),
                        'Mã hàng': item.get('ma_hang', ''),
                        'Tên hàng': item.get('ten_hang', '')[:50],
                        'Số lượng': item.get('so_luong', 0),
                        'ĐVT': item.get('dvt', ''),
                    })
        else:
            # Hiển thị dạng gộp (mặc định)
            display_data = []
            for pxk, items in pxk_totals.items():
                for ma_hang, sl in items.items():
                    display_data.append({
                        'Số PXK': pxk,
                        'Ngày': pxk_dates.get(pxk, ''),
                        'D/O No': ', '.join(pxk_do_no.get(pxk, [])),
                        'Mã hàng': ma_hang,
                        'Số lượng': sl,
                    })
        
        df_display = pd.DataFrame(display_data)
        
        # Filter
        col1, col2 = st.columns([3, 2])
        with col1:
            search_pxk = st.text_input('🔍 Lọc theo Số PXK', placeholder='VD: 1174')
        with col2:
            search_mh = st.text_input('🔍 Lọc theo Mã hàng', placeholder='VD: DC97-22471T')
        
        df_filtered = df_display.copy()
        if search_pxk:
            df_filtered = df_filtered[df_filtered['Số PXK'].astype(str).str.contains(search_pxk)]
        if search_mh:
            df_filtered = df_filtered[df_filtered['Mã hàng'].str.contains(search_mh, case=False)]
        
        st.dataframe(df_filtered, use_container_width=True, hide_index=True, height=400)
        
        # Thông báo nếu có nhiều dòng cùng mã hàng
        if view_mode == 'chi_tiet' and pxk_items:
            total_lines = sum(len(items) for items in pxk_items.values())
            total_unique = sum(len(set(item['ma_hang'] for item in items)) for items in pxk_items.values())
            if total_lines > total_unique:
                st.info(f'ℹ️ Có {total_lines} dòng chi tiết (bao gồm {total_lines - total_unique} dòng trùng mã hàng)')
        
        # Export
        st.divider()
        st.subheader('⬇️ Tải kết quả')
        
        buffer = io.BytesIO()
        df_display.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)
        
        st.download_button(
            label='📥 Tải dữ liệu trích xuất (.xlsx)',
            data=buffer.getvalue(),
            file_name='du_lieu_pxk_trich_xuat.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type='primary',
        )
    
    else:
        # ── Hiển thị kết quả đầy đủ ───────────────────────────────────────────
        st.subheader('📊 Kết quả ghép số PXK')
        
        form_rows = cache['form_rows']
        result = cache['result']
        status_list = cache['status_list']
        note_pxks = cache['note_pxks']
        pxk_dates = cache['pxk_dates']
        pdf_errors = cache['pdf_errors']
        output_bytes = cache['output_bytes']
        scorer = cache['scorer']
        
        # KPIs
        n_auto = sum(1 for s in status_list if s == 'auto')
        n_amb = sum(1 for s in status_list if s == 'ambiguous')
        n_none = sum(1 for s in status_list if s == 'no_match')
        total = len(form_rows)
        
        # Thông tin chi tiết về AI
        with st.expander('🧠 Chi tiết AI học từ dữ liệu', expanded=False):
            if scorer and scorer.folder_count > 0:
                col1, col2, col3 = st.columns(3)
                col1.metric('📁 Folder đã học', scorer.folder_count)
                col1.metric('📊 Mẫu huấn luyện', scorer.example_count)
                col2.metric('📝 Invoice patterns', len(scorer.inv_to_pxk_counter))
                col2.metric('📦 Item patterns', len(scorer.item_to_pxk_counter))
                col3.metric('🔗 Inv+Item patterns', len(scorer.inv_item_to_pxk_counter))
                col3.metric('📈 Qty patterns', len(scorer.qty_pattern_counter))
                
                st.info(
                    f'🎯 **Tỷ lệ thành công:** {n_auto}/{total} = **{n_auto/total*100:.1f}%** tự động\n\n'
                    f'💡 **Lưu ý:** Các mẫu học giúp AI đưa ra quyết định chính xác hơn '
                    f'khi có nhiều PXK khả dĩ cho cùng một mã hàng.'
                )
            else:
                st.warning('⚠️ Không có dữ liệu huấn luyện. Kết quả có thể không tối ưu.')
        
        # KPIs
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric('📚 Folder học', scorer.folder_count if scorer else 0)
        c2.metric('🧠 Mẫu học', scorer.example_count if scorer else 0)
        c3.metric('✅ Tự động', n_auto)
        c4.metric('🔍 Cần kiểm tra', n_amb)
        c5.metric('❌ Không khớp', n_none)
        
        # Progress bar
        if total > 0:
            pa = n_auto / total * 100
            pb = n_amb / total * 100
            pc = n_none / total * 100
            st.markdown(f'''
            <div style="background: #FFC7CE; border-radius: 6px; height: 20px; width: 100%; overflow: hidden;">
                <div style="background: #C6EFCE; width: {pa:.1f}%; height: 100%; float: left;"></div>
                <div style="background: #FFEB9C; width: {pb:.1f}%; height: 100%; float: left;"></div>
            </div>
            <small>✅ {pa:.1f}% tự động | 🔍 {pb:.1f}% cần KT | ❌ {pc:.1f}% không khớp</small>
            ''', unsafe_allow_html=True)
        
        if pdf_errors:
            with st.expander(f'⚠️ {len(pdf_errors)} file PDF lỗi'):
                st.dataframe(pd.DataFrame(pdf_errors), hide_index=True)
        
        # Thêm phần xem chi tiết dữ liệu PDF đã trích xuất
        st.divider()
        with st.expander('📄 Xem chi tiết dữ liệu PDF đã trích xuất', expanded=False):
            pxk_totals_view = cache['pxk_totals']
            pxk_items_view = cache.get('pxk_items', {})
            pxk_dates_view = cache['pxk_dates']
            pxk_do_no_view = cache['pxk_do_no']
            
            # Cho phép chọn chế độ xem
            view_mode_full = st.radio(
                'Chế độ xem:',
                ['chi_tiet', 'gop'],
                format_func=lambda x: '🔍 Xem từng dòng riêng biệt' if x == 'chi_tiet' else '📊 Xem gộp theo mã hàng',
                horizontal=True,
                key='view_mode_full'
            )
            
            display_data_full = []
            if view_mode_full == 'chi_tiet' and pxk_items_view:
                for pxk, items in pxk_items_view.items():
                    for item in items:
                        display_data_full.append({
                            'Số PXK': pxk,
                            'Ngày': pxk_dates_view.get(pxk, ''),
                            'D/O No': ', '.join(pxk_do_no_view.get(pxk, [])),
                            'Dòng': item.get('line_no', ''),
                            'Mã hàng': item.get('ma_hang', ''),
                            'Tên hàng': item.get('ten_hang', '')[:50],
                            'Số lượng': item.get('so_luong', 0),
                            'ĐVT': item.get('dvt', ''),
                        })
            else:
                for pxk, items in pxk_totals_view.items():
                    for ma_hang, sl in items.items():
                        display_data_full.append({
                            'Số PXK': pxk,
                            'Ngày': pxk_dates_view.get(pxk, ''),
                            'D/O No': ', '.join(pxk_do_no_view.get(pxk, [])),
                            'Mã hàng': ma_hang,
                            'Số lượng': sl,
                        })
            
            df_full_display = pd.DataFrame(display_data_full)
            
            # Filter cho bảng PDF
            col_f1, col_f2 = st.columns([3, 2])
            with col_f1:
                search_pxk_full = st.text_input('🔍 Lọc theo Số PXK', placeholder='VD: 1174', key='search_pxk_full')
            with col_f2:
                search_mh_full = st.text_input('🔍 Lọc theo Mã hàng', placeholder='VD: DC97-22471T', key='search_mh_full')
            
            df_full_filtered = df_full_display.copy()
            if search_pxk_full:
                df_full_filtered = df_full_filtered[df_full_filtered['Số PXK'].astype(str).str.contains(search_pxk_full)]
            if search_mh_full:
                df_full_filtered = df_full_filtered[df_full_filtered['Mã hàng'].str.contains(search_mh_full, case=False)]
            
            st.dataframe(df_full_filtered, use_container_width=True, hide_index=True, height=300)
        
        # Bảng kết quả
        st.divider()
        st.subheader('📋 Chi tiết ghép PXK')
        
        rows_data = []
        for fr in form_rows:
            idx = fr['idx']
            pxk = result[idx]
            state = status_list[idx]
            candidates = sorted([p for p in note_pxks[idx] if p != pxk], key=pxk_sort_key)
            
            rows_data.append({
                'Dòng': fr['row'],
                'Invoice': fr.get('inv') or '',
                'Mã hàng': fr['ma_hang'],
                'Số lượng GR': int(fr['sl']),
                'Số PXK': f'{int(pxk):08d}' if pxk and str(pxk).isdigit() else (pxk or ''),
                'Ngày PXK': pxk_dates.get(pxk, '') if pxk else '',
                'Trạng thái': (
                    '✅ Tự động' if state == 'auto'
                    else '🔍 Cần kiểm tra' if state == 'ambiguous'
                    else '❌ Không khớp'
                ),
                'PXK khả dĩ khác': ' | '.join(
                    f'{int(p):08d}' if str(p).isdigit() else p for p in candidates[:5]
                ),
            })
        
        df = pd.DataFrame(rows_data)
        
        # Filter
        col1, col2 = st.columns([3, 2])
        with col1:
            search = st.text_input('🔍 Lọc theo Mã hàng / Invoice / PXK')
        with col2:
            filter_state = st.selectbox(
                'Lọc trạng thái',
                ['Tất cả', '✅ Tự động', '🔍 Cần kiểm tra', '❌ Không khớp']
            )
        
        df_show = df.copy()
        if search:
            df_show = df_show[
                df_show['Mã hàng'].str.contains(search, case=False, na=False)
                | df_show['Invoice'].astype(str).str.contains(search, case=False, na=False)
                | df_show['Số PXK'].astype(str).str.contains(search, case=False, na=False)
            ]
        if filter_state != 'Tất cả':
            df_show = df_show[df_show['Trạng thái'] == filter_state]
        
        st.dataframe(df_show, use_container_width=True, hide_index=True, height=420)
        st.caption(f'Hiển thị {len(df_show):,} / {len(df):,} dòng')
        
        # Download section - Cả 2 file
        st.divider()
        st.subheader('⬇️ Tải kết quả')
        
        # Tạo file dữ liệu trích xuất PDF để tải xuống
        pxk_totals = cache['pxk_totals']
        pxk_items = cache.get('pxk_items', {})
        pxk_dates = cache['pxk_dates']
        pxk_do_no = cache['pxk_do_no']
        
        # Tạo DataFrame chi tiết từ PDF
        display_data = []
        if pxk_items:
            # Có chi tiết từng dòng riêng biệt
            for pxk, items in pxk_items.items():
                for item in items:
                    display_data.append({
                        'Số PXK': pxk,
                        'Ngày': pxk_dates.get(pxk, ''),
                        'D/O No': ', '.join(pxk_do_no.get(pxk, [])),
                        'Dòng': item.get('line_no', ''),
                        'Mã hàng': item.get('ma_hang', ''),
                        'Tên hàng': item.get('ten_hang', '')[:50],
                        'Số lượng': item.get('so_luong', 0),
                        'ĐVT': item.get('dvt', ''),
                        'Đơn giá': item.get('don_gia', 0),
                        'Thành tiền': item.get('thanh_tien', 0),
                    })
        else:
            # Chỉ có dạng gộp
            for pxk, items in pxk_totals.items():
                for ma_hang, sl in items.items():
                    display_data.append({
                        'Số PXK': pxk,
                        'Ngày': pxk_dates.get(pxk, ''),
                        'D/O No': ', '.join(pxk_do_no.get(pxk, [])),
                        'Mã hàng': ma_hang,
                        'Số lượng': sl,
                    })
        
        df_pdf_data = pd.DataFrame(display_data)
        
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            buffer_pdf = io.BytesIO()
            df_pdf_data.to_excel(buffer_pdf, index=False, engine='openpyxl')
            buffer_pdf.seek(0)
            
            st.download_button(
                label='📄 Tải DỮ LIỆU PDF trích xuất (.xlsx)',
                data=buffer_pdf.getvalue(),
                file_name='du_lieu_pxk_trich_xuat.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='secondary',
                use_container_width=True,
            )
            st.caption('Dữ liệu gốc trích xuất từ các file PDF')
        
        with col_dl2:
            st.download_button(
                label='📋 Tải FORM ĐÃ ĐIỀN PXK (.xlsx)',
                data=output_bytes,
                file_name='FORM_CHUA_NHAP_DA_DIEN_PXK_v5.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
                use_container_width=True,
            )
            st.caption('Form đã được ghép số PXK tự động')
