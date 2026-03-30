from pathlib import Path
import importlib.util
import sys

import pandas as pd
import streamlit as st


def _load_pxk_core_v4():
    try:
        from pxk_core_v4 import (  # type: ignore
            build_output_excel,
            extract_pdfs_from_files,
            load_reference_scorer,
            match_pxk_v4,
            pxk_sort_key,
            read_form_rows_from_bytes,
        )
        return (
            build_output_excel,
            extract_pdfs_from_files,
            load_reference_scorer,
            match_pxk_v4,
            pxk_sort_key,
            read_form_rows_from_bytes,
        )
    except ModuleNotFoundError:
        core_path = Path(__file__).with_name("pxk_core_v4.py")
        if not core_path.exists():
            raise
        spec = importlib.util.spec_from_file_location("pxk_core_v4", core_path)
        if spec is None or spec.loader is None:
            raise
        module = importlib.util.module_from_spec(spec)
        sys.modules["pxk_core_v4"] = module
        spec.loader.exec_module(module)
        return (
            module.build_output_excel,
            module.extract_pdfs_from_files,
            module.load_reference_scorer,
            module.match_pxk_v4,
            module.pxk_sort_key,
            module.read_form_rows_from_bytes,
        )


(
    build_output_excel,
    extract_pdfs_from_files,
    load_reference_scorer,
    match_pxk_v4,
    pxk_sort_key,
    read_form_rows_from_bytes,
) = _load_pxk_core_v4()


st.set_page_config(
    page_title="PXK Manager – SMC v4",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
.step-header { font-size:1.05rem; font-weight:700; color:#1F4E79; margin-bottom:4px; }
.tag-green  { background:#C6EFCE; color:#276221; padding:2px 8px; border-radius:4px; font-size:.85rem; }
.tag-yellow { background:#FFEB9C; color:#7d6608; padding:2px 8px; border-radius:4px; font-size:.85rem; }
.tag-red    { background:#FFC7CE; color:#9c1f23; padding:2px 8px; border-radius:4px; font-size:.85rem; }
</style>
""",
    unsafe_allow_html=True,
)

st.title("📦 PXK Manager – App v4")
st.caption("App v4 dùng dữ liệu đã điền ở các folder huấn luyện để chấm điểm các ca mơ hồ.")

with st.sidebar:
    st.header("⚙️ Các bước thực hiện")
    st.markdown('<div class="step-header">Bước 1 — Upload PDF</div>', unsafe_allow_html=True)
    pdf_files = st.file_uploader(
        "Chọn file PDF Phiếu Xuất Kho",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_upload",
        help="Có thể chọn nhiều file (Ctrl+Click)",
    )
    if pdf_files:
        st.success(f"✅ {len(pdf_files)} file PDF")

    st.divider()
    st.markdown('<div class="step-header">Bước 2 — Upload FORM CHƯA NHẬP</div>', unsafe_allow_html=True)
    form_file = st.file_uploader(
        "FORM DỮ LIỆU CHƯA NHẬP SỐ PXK.xlsx",
        type=["xlsx"],
        key="form_upload",
    )
    if form_file:
        st.success(f"✅ {form_file.name}")

    st.divider()
    run_btn = st.button(
        "🚀 Bắt đầu xử lý",
        type="primary",
        disabled=not (pdf_files and form_file),
        use_container_width=True,
    )
    st.divider()
    st.caption("💡 Màu sắc kết quả:")
    st.markdown('<span class="tag-green">✅ Tự động</span> — 1 PXK duy nhất khớp', unsafe_allow_html=True)
    st.markdown('<span class="tag-yellow">🔍 Cần kiểm tra</span> — có chấm điểm từ dữ liệu lịch sử', unsafe_allow_html=True)
    st.markdown('<span class="tag-red">❌ Không khớp</span> — cần điền thủ công', unsafe_allow_html=True)


if not (pdf_files and form_file):
    c1, c2 = st.columns(2)
    c1.info("👈 Upload PDF Phiếu Xuất Kho ở thanh bên trái.")
    c2.info("👈 Upload FORM CHƯA NHẬP để App v4 ghép số PXK.")
    st.stop()


if run_btn or st.session_state.get("processed_v4"):
    if run_btn:
        st.session_state.pop("processed_v4", None)
        st.session_state.pop("results_cache_v4", None)

    if "results_cache_v4" not in st.session_state:
        with st.status("⏳ Đang xử lý App v4...", expanded=True) as sb:
            file_bytes = [(f.name, f.read()) for f in pdf_files]
            st.write(f"📄 Đang trích xuất {len(file_bytes)} file PDF...")
            pxk_totals, pxk_dates, pxk_do_no, pdf_errors = extract_pdfs_from_files(file_bytes)
            st.write(
                f"✅ {len(pxk_totals)} PXK"
                + (f" | ⚠️ {len(pdf_errors)} lỗi" if pdf_errors else "")
            )

            st.write("📚 Đang nạp dữ liệu học từ các folder đã điền...")
            scorer = load_reference_scorer(".")
            st.write(
                f"✅ {scorer.folder_count} folder tham chiếu | {scorer.example_count} dòng huấn luyện"
            )

            st.write("📊 Đang đọc FORM CHƯA NHẬP...")
            form_bytes = form_file.read()
            form_rows = read_form_rows_from_bytes(form_bytes)
            st.write(f"✅ {len(form_rows)} dòng dữ liệu")

            st.write("🧠 Đang ghép số PXK bằng App v4...")
            result, status_list, note_pxks = match_pxk_v4(
                form_rows, pxk_totals, pxk_do_no, scorer=scorer
            )
            n_auto = sum(1 for s in status_list if s == "auto")
            n_amb = sum(1 for s in status_list if s == "ambiguous")
            n_none = sum(1 for s in status_list if s == "no_match")
            st.write(f"✅ Auto: **{n_auto}** | Cần KT: **{n_amb}** | Không khớp: **{n_none}**")

            st.write("💾 Đang tạo Excel kết quả...")
            output_bytes = build_output_excel(
                form_bytes, form_rows, result, status_list, note_pxks, pxk_dates
            )

            st.session_state["results_cache_v4"] = {
                "pxk_totals": pxk_totals,
                "pxk_dates": pxk_dates,
                "form_rows": form_rows,
                "result": result,
                "status_list": status_list,
                "note_pxks": note_pxks,
                "output_bytes": output_bytes,
                "pdf_errors": pdf_errors,
                "scorer": scorer,
            }
            st.session_state["processed_v4"] = True
            sb.update(label="✅ Xử lý App v4 hoàn tất", state="complete")

    cache = st.session_state["results_cache_v4"]
    pxk_dates = cache["pxk_dates"]
    form_rows = cache["form_rows"]
    result = cache["result"]
    status_list = cache["status_list"]
    note_pxks = cache["note_pxks"]
    output_bytes = cache["output_bytes"]
    pdf_errors = cache["pdf_errors"]
    scorer = cache["scorer"]

    n_auto = sum(1 for s in status_list if s == "auto")
    n_amb = sum(1 for s in status_list if s == "ambiguous")
    n_none = sum(1 for s in status_list if s == "no_match")
    total = len(form_rows)

    st.subheader("📊 Kết quả tổng quan")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("📚 Folder học", scorer.folder_count)
    c2.metric("🧠 Dòng học", scorer.example_count)
    c3.metric("✅ Tự động", n_auto)
    c4.metric("🔍 Cần kiểm tra", n_amb)
    c5.metric("❌ Không khớp", n_none)

    if pdf_errors:
        with st.expander(f"⚠️ {len(pdf_errors)} file PDF có lỗi"):
            st.dataframe(pd.DataFrame(pdf_errors), hide_index=True, use_container_width=True)

    rows_data = []
    for fr in form_rows:
        idx = fr["idx"]
        pxk = result[idx]
        state = status_list[idx]
        candidates = sorted([p for p in note_pxks[idx] if p != pxk], key=pxk_sort_key)
        rows_data.append(
            {
                "Dòng": fr["row"],
                "Invoice": fr.get("inv") or "",
                "Mã hàng": fr["ma_hang"],
                "Số lượng GR": int(fr["sl"]),
                "Số PXK": f"{int(pxk):08d}" if pxk and str(pxk).isdigit() else (pxk or ""),
                "Ngày PXK": pxk_dates.get(pxk, "") if pxk else "",
                "Trạng thái": (
                    "✅ Tự động"
                    if state == "auto"
                    else "🔍 Cần kiểm tra"
                    if state == "ambiguous"
                    else "❌ Không khớp"
                ),
                "PXK khả dĩ khác": " | ".join(
                    f"{int(p):08d}" if str(p).isdigit() else p for p in candidates[:5]
                ),
            }
        )

    df = pd.DataFrame(rows_data)
    c1, c2 = st.columns([3, 2])
    with c1:
        search = st.text_input("🔍 Lọc theo Mã hàng / Invoice / PXK")
    with c2:
        filter_state = st.selectbox(
            "Lọc trạng thái", ["Tất cả", "✅ Tự động", "🔍 Cần kiểm tra", "❌ Không khớp"]
        )

    df_show = df.copy()
    if search:
        df_show = df_show[
            df_show["Mã hàng"].str.contains(search, case=False, na=False)
            | df_show["Invoice"].astype(str).str.contains(search, case=False, na=False)
            | df_show["Số PXK"].astype(str).str.contains(search, case=False, na=False)
        ]
    if filter_state != "Tất cả":
        df_show = df_show[df_show["Trạng thái"] == filter_state]

    st.dataframe(df_show, use_container_width=True, hide_index=True, height=420)
    st.caption(f"Hiển thị {len(df_show):,} / {len(df):,} dòng")

    st.download_button(
        label="📥 Tải FORM ĐÃ ĐIỀN PXK (.xlsx)",
        data=output_bytes,
        file_name="FORM_CHUA_NHAP_DA_DIEN_PXK_v4.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
