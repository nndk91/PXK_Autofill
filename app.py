"""
PXK Extractor — Streamlit web app
Upload nhiều file PDF Phiếu Xuất Kho, extract dữ liệu, xuất Excel.
"""

import os
import tempfile

import pandas as pd
import streamlit as st

from excel_writer import results_to_dataframe, results_to_excel
from pdf_extractor import extract_pxk

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PXK Extractor – SMC",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📦 Trích xuất dữ liệu Phiếu Xuất Kho")
st.caption("Upload PDF → Extract tự động → Tải Excel")

# ── Sidebar — upload ────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("1. Chọn file PDF")
    uploaded_files = st.file_uploader(
        "Chọn một hoặc nhiều file PXK",
        type=["pdf"],
        accept_multiple_files=True,
        help="Có thể chọn nhiều file cùng lúc (Ctrl+Click / Cmd+Click)",
    )

    if uploaded_files:
        st.success(f"Đã chọn **{len(uploaded_files)}** file")

    st.divider()
    st.header("2. Xử lý")
    run_btn = st.button("🔄 Bắt đầu Extract", type="primary",
                        disabled=not uploaded_files)

# ── Main area ───────────────────────────────────────────────────────────────────
if not uploaded_files:
    st.info("👈 Chọn file PDF ở thanh bên trái để bắt đầu.")
    st.stop()

# Run extraction when button pressed
if run_btn:
    results = []
    errors  = []

    progress_bar = st.progress(0, text="Đang xử lý…")
    status_text  = st.empty()

    with tempfile.TemporaryDirectory() as tmpdir:
        total = len(uploaded_files)
        for i, uf in enumerate(uploaded_files):
            status_text.text(f"⏳ Đang xử lý ({i+1}/{total}): {uf.name}")

            # Write upload to temp file (pdfplumber needs a path)
            tmp_path = os.path.join(tmpdir, uf.name)
            with open(tmp_path, "wb") as f:
                f.write(uf.getvalue())

            r = extract_pxk(tmp_path)

            if r["error"]:
                errors.append({"file": uf.name, "chi_tiet": r["error"], "loai": "Lỗi đọc PDF"})
            elif not r["items"]:
                errors.append({"file": uf.name, "chi_tiet": "Không tìm thấy dòng hàng hóa", "loai": "Không có dữ liệu"})
            else:
                results.append(r)

            progress_bar.progress((i + 1) / total,
                                  text=f"Đang xử lý… {i+1}/{total}")

    progress_bar.empty()
    status_text.empty()

    # Persist results in session state
    st.session_state["results"] = results
    st.session_state["errors"]  = errors

# Show results if available
results = st.session_state.get("results", [])
errors  = st.session_state.get("errors",  [])

if not results and not errors:
    st.stop()

# ── Error report ────────────────────────────────────────────────────────────────
if errors:
    with st.expander(f"⚠️ {len(errors)} file có vấn đề", expanded=False):
        err_df = pd.DataFrame(errors)
        st.dataframe(err_df, use_container_width=True, hide_index=True)

# ── Success results ─────────────────────────────────────────────────────────────
if results:
    df = results_to_dataframe(results)

    # KPI row
    total_pxk   = df["so_phieu"].nunique()
    total_rows  = len(df)
    total_sl    = df["so_luong"].sum()
    total_tien  = df["thanh_tien"].sum()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("📄 Số PXK", f"{total_pxk}")
    k2.metric("📋 Dòng hàng", f"{total_rows:,}")
    k3.metric("📦 Tổng số lượng", f"{total_sl:,.0f}")
    k4.metric("💰 Tổng thành tiền", f"{total_tien:,.0f} ₫")

    st.divider()

    # Preview table
    st.subheader("Dữ liệu đã extract")

    # Filter controls
    col_search, col_ma = st.columns([3, 2])
    with col_search:
        search_pxk = st.text_input("🔍 Lọc theo Số PXK", placeholder="VD: 1174")
    with col_ma:
        search_ma = st.text_input("🔍 Lọc theo Mã hàng", placeholder="VD: DC97-22471T")

    display_df = df.copy()
    if search_pxk:
        display_df = display_df[display_df["so_phieu"].str.contains(search_pxk, na=False)]
    if search_ma:
        display_df = display_df[display_df["ma_hang"].str.contains(search_ma, case=False, na=False)]

    st.dataframe(
        display_df.rename(columns={
            "so_phieu":   "Số PXK",
            "ngay":       "Ngày",
            "ma_hang":    "Mã hàng",
            "ten_hang":   "Tên hàng",
            "dvt":        "ĐVT",
            "so_luong":   "Số lượng",
            "don_gia":    "Đơn giá",
            "thanh_tien": "Thành tiền",
            "file_name":  "File nguồn",
        }),
        use_container_width=True,
        hide_index=True,
    )

    st.divider()

    # Download
    st.subheader("Xuất Excel")
    excel_bytes = results_to_excel(results)
    st.download_button(
        label="📥 Tải file Excel",
        data=excel_bytes,
        file_name="du_lieu_pxk.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption(f"File Excel chứa **{total_rows}** dòng dữ liệu từ **{total_pxk}** phiếu.")
