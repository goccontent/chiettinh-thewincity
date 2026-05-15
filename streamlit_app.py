# -*- coding: utf-8 -*-
"""
THE WIN CITY – Công cụ tạo Chiết Tính online (Streamlit).
"""
from __future__ import annotations

import io
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

import openpyxl
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
GIO_HANG_FILE = BASE_DIR / "GIỎ HÀNG LM VAT THE WIN CITY 25.04.2026 (1).xlsx"
CHIETTINH_FILE = BASE_DIR / "CHIETTINH.xlsx"

st.set_page_config(
    page_title="Chiết Tính – THE WIN CITY",
    page_icon="🏙️",
    layout="wide",
)

# ---------- Styling ----------
st.markdown(
    """
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; max-width: 1200px; }
    h1, h2, h3 { font-family: 'Segoe UI', Arial, sans-serif; }
    .hero {
        background: linear-gradient(135deg, #0f2e5c 0%, #1a4d8f 50%, #2563b8 100%);
        padding: 28px 32px; border-radius: 16px; color: white;
        margin-bottom: 24px;
        box-shadow: 0 10px 30px rgba(15,46,92,0.25);
    }
    .hero h1 { color: white; margin: 0; font-size: 2rem; }
    .hero p  { color: #d8e4f5; margin: 6px 0 0 0; font-size: 1rem; }
    .card {
        background: #ffffff; border: 1px solid #e6eaf2;
        border-radius: 14px; padding: 20px 24px;
        box-shadow: 0 4px 14px rgba(15,46,92,0.06);
        margin-bottom: 16px;
    }
    .price-pill {
        display: inline-block; padding: 6px 14px; border-radius: 999px;
        background: #ecf2ff; color: #1a4d8f; font-weight: 600; font-size: 0.95rem;
        margin-right: 8px;
    }
    .price-big {
        font-size: 2rem; font-weight: 700; color: #0f2e5c; margin: 4px 0;
    }
    .muted { color: #6b7280; font-size: 0.9rem; }
    .stDownloadButton button {
        background: linear-gradient(135deg, #1a4d8f, #2563b8) !important;
        color: white !important; border: none !important;
        padding: 12px 24px !important; font-weight: 600 !important;
        border-radius: 10px !important;
    }
    div[data-testid="stMetricValue"] { font-size: 1.4rem; color: #0f2e5c; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- Helpers ----------
def parse_price(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return int(v)
    digits = re.sub(r"[^\d]", "", str(v))
    return int(digits) if digits else None


@st.cache_data(show_spinner=False)
def load_products():
    wb = openpyxl.load_workbook(GIO_HANG_FILE, data_only=True)
    ws = wb["Sheet1"]
    out = {}
    for r in range(3, ws.max_row + 1):
        ma = ws.cell(r, 2).value
        if not ma:
            continue
        ma = str(ma).strip()
        out[ma] = {
            "ma": ma,
            "loai": ws.cell(r, 3).value,
            "so_can": ws.cell(r, 4).value,
            "so_tang": ws.cell(r, 5).value,
            "thap": ws.cell(r, 6).value,
            "dt_thong_thuy": ws.cell(r, 7).value,
            "dt_tim_tuong": ws.cell(r, 8).value,
            "huong": ws.cell(r, 10).value,
            "view": ws.cell(r, 11).value,
            "gia_vn": parse_price(ws.cell(r, 12).value),
            "gia_nn": parse_price(ws.cell(r, 13).value),
            "tinh_trang": ws.cell(r, 14).value,
        }
    return out


def fmt_vnd(n):
    if n is None:
        return "—"
    return f"{n:,.0f}".replace(",", ".") + " ₫"


def compute_breakdown(gia: int):
    """Tái hiện công thức trong CHIETTINH.xlsx để preview."""
    ck_phi_ma_dau = round(gia * 0.02)
    ck_dau_tu = 0
    after_2 = gia - ck_phi_ma_dau - ck_dau_tu
    ck_pttt = round(after_2 * 0.02)
    gia_truoc_vat = gia - ck_phi_ma_dau - ck_dau_tu - ck_pttt
    # Giá đất được trừ: tạm tính theo công thức trong PL1-Chuẩn (sử dụng DT tim tường * 1,000,079)
    # Bỏ qua trong preview - chỉ hiển thị các giá trị chính
    vat = round(gia_truoc_vat * 0.10)
    tong_co_vat = gia_truoc_vat + vat
    phi_bao_tri = round(gia_truoc_vat * 0.02)
    tong_cuoi = tong_co_vat + phi_bao_tri
    return {
        "Giá CĐT (trước VAT)": gia,
        'CK "Phi Mã Đón Đầu" (2%)': -ck_phi_ma_dau,
        "CK theo PTTT Chuẩn (2%)": -ck_pttt,
        "Giá HĐMB trước VAT": gia_truoc_vat,
        "VAT (10%)": vat,
        "Tổng giá HĐMB (gồm VAT)": tong_co_vat,
        "Phí bảo trì (2%)": phi_bao_tri,
        "TỔNG GIÁ TRỊ HỢP ĐỒNG": tong_cuoi,
    }


def try_recalc(path: Path) -> bool:
    """Cố gắng tính lại công thức bằng LibreOffice (có sẵn trên Streamlit Cloud)."""
    for soffice in ["soffice", "libreoffice",
                    r"C:\Program Files\LibreOffice\program\soffice.exe"]:
        try:
            subprocess.run(
                [soffice, "--headless", "--calc",
                 "--convert-to", "xlsx",
                 "--outdir", str(path.parent), str(path)],
                check=True, timeout=90,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            return True
        except (FileNotFoundError, subprocess.SubprocessError):
            continue
    return False


def build_chiettinh(product: dict, is_foreigner: bool, block: str) -> tuple[bytes, bool]:
    gia = product["gia_nn"] if is_foreigner else product["gia_vn"]
    if gia is None:
        raise ValueError("Mã sản phẩm này chưa có giá.")

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        safe = re.sub(r"[^A-Za-z0-9._-]", "_", product["ma"])
        out_path = td / f"CHIETTINH_{safe}.xlsx"
        shutil.copy(CHIETTINH_FILE, out_path)

        wb = openpyxl.load_workbook(out_path)
        pl1a = wb["PL1A"]
        pl1a["C7"] = product["ma"]
        pl1a["D7"] = product["loai"]
        if product["so_tang"] is not None and product["so_can"] is not None:
            pl1a["E7"] = f"CĂN SỐ {product['so_can']} TẦNG {product['so_tang']}"
        pl1a["C11"] = gia
        if block:
            pl1a["C6"] = block
        elif product.get("thap"):
            pl1a["C6"] = str(product["thap"])

        for sn in wb.sheetnames:
            if sn == "PL1A":
                continue
            ws = wb[sn]
            try:
                if ws["A25"].value and "Mã sản phẩm" in str(ws["A25"].value):
                    ws["C25"] = product["ma"]
                    ws["D25"] = product["loai"]
                if ws["A26"].value and "Diện tích" in str(ws["A26"].value):
                    if product["dt_tim_tuong"]:
                        ws["C26"] = product["dt_tim_tuong"]
            except Exception:
                pass

        # Bắt Excel tính lại toàn bộ công thức khi mở file
        from openpyxl.workbook.properties import CalcProperties
        wb.calculation = CalcProperties(calcId=191241, fullCalcOnLoad=True, calcMode="auto")
        for sn in wb.sheetnames:
            ws = wb[sn]
            ws.force_formula_recalculation = True if hasattr(ws, "force_formula_recalculation") else None

        wb.save(out_path)
        ok = try_recalc(out_path)
        data = out_path.read_bytes()
        return data, ok


# ---------- UI ----------
st.markdown(
    """
    <div class="hero">
      <h1>🏙️ Công cụ Chiết Tính – THE WIN CITY</h1>
      <p>Chọn mã sản phẩm → tự động tra giá từ Giỏ hàng → xuất file Chiết Tính kèm tiến độ thanh toán.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

try:
    products = load_products()
except FileNotFoundError as e:
    st.error(f"❌ Thiếu file dữ liệu: {e}")
    st.stop()

codes = sorted(products.keys())

col_l, col_r = st.columns([1, 1.2], gap="large")

with col_l:
    st.markdown("### 1. Chọn căn")
    code = st.selectbox(
        "Mã sản phẩm",
        options=codes,
        index=0,
        help=f"Có {len(codes)} mã căn trong file Giỏ hàng.",
    )
    kind = st.radio(
        "Đối tượng khách hàng",
        options=["Người Việt Nam", "Người Nước ngoài"],
        horizontal=True,
    )
    block = st.text_input("Block (tùy chọn)", placeholder="VD: C1.3 – để trống dùng giá trị mặc định")
    st.caption("Giá và mọi công thức trong file CHIETTINH sẽ tự cập nhật theo lựa chọn.")

product = products[code]
is_nn = kind == "Người Nước ngoài"
gia = product["gia_nn"] if is_nn else product["gia_vn"]

with col_r:
    st.markdown("### 2. Thông tin căn")
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(
        f'<span class="price-pill">{product["loai"] or "—"}</span>'
        f'<span class="price-pill">Tháp {product["thap"]}</span>'
        f'<span class="price-pill">Tầng {product["so_tang"]} – Căn {product["so_can"]}</span>',
        unsafe_allow_html=True,
    )
    st.markdown(f'<div class="price-big">{fmt_vnd(gia)}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="muted">Giá CĐT trước VAT · {kind}</div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("DT thông thủy", f"{product['dt_thong_thuy']} m²" if product['dt_thong_thuy'] else "—")
    c2.metric("DT tim tường", f"{product['dt_tim_tuong']} m²" if product['dt_tim_tuong'] else "—")
    c3.metric("Hướng", product["huong"] or "—")
    st.markdown(f'<div class="muted">View: {product["view"] or "—"} · Tình trạng: {product["tinh_trang"] or "—"}</div>',
                unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("### 3. Bảng tính nhanh (PTTT Chuẩn – tham khảo)")
if gia:
    bd = compute_breakdown(gia)
    cols = st.columns(4)
    items = list(bd.items())
    for i, (k, v) in enumerate(items):
        with cols[i % 4]:
            st.metric(k, fmt_vnd(abs(v)) if v >= 0 else f"−{fmt_vnd(abs(v))}")
    st.caption("Các con số trên là tham khảo cho PTTT Chuẩn. File Excel xuất ra chứa đầy đủ 7 sheet PTTT.")
else:
    st.warning("Mã này chưa có giá trong file Giỏ hàng.")

st.markdown("### 4. Xuất file Chiết Tính")
go = st.button("⚙️ Tạo file Chiết Tính", type="primary", use_container_width=True, disabled=gia is None)

if go and gia:
    with st.spinner("Đang điền dữ liệu vào template và tính công thức..."):
        try:
            data, recalc_ok = build_chiettinh(product, is_nn, block.strip())
        except Exception as e:
            st.error(f"Lỗi: {e}")
            st.stop()

    safe = re.sub(r"[^A-Za-z0-9._-]", "_", product["ma"])
    suffix = "NN" if is_nn else "VN"
    st.success("✅ Đã tạo xong file. Bấm tải xuống bên dưới.")
    if not recalc_ok:
        st.info("ℹ️ Máy chủ chưa có LibreOffice nên công thức chưa được tính sẵn. "
                "Khi mở file trong Excel lần đầu, bấm **Ctrl+Alt+F9** để Excel tự cập nhật giá trị.")
    st.download_button(
        label=f"⬇️ Tải CHIETTINH_{safe}_{suffix}.xlsx",
        data=data,
        file_name=f"CHIETTINH_{safe}_{suffix}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("---")
st.caption("© THE WIN CITY · Dữ liệu lấy từ file Giỏ hàng nội bộ · Bản công cụ chiết tính tự động.")
