"""
app.py
======
Streamlit dashboard for K. Beauty / The Beauty Shop - 3 menus:

1. Welcome        - overview stats once Data Stok is uploaded
2. PO Recommendation - weekly PO draft from Data Stok + Data Penjualan 30 Hari
3. Sales Recap    - dead-stock finder from Data Stok + Data Penjualan 3 Bulan

Upload iPOS 5.0 exports as-is - no manual cleanup needed. No n8n, no paid
service, no external database.

RUN LOCALLY
-----------
    pip install -r requirements.txt
    streamlit run app.py

DEPLOY FREE (so your team can access it from a browser link)
--------------------------------------------------------------
    1. Push this folder to a GitHub repo (app.py, po_logic.py,
       requirements.txt, assets/logo.jpg, .streamlit/config.toml).
    2. Go to https://share.streamlit.io -> "New app" -> connect the repo.
    3. Streamlit Community Cloud builds and hosts it for free.
"""

import io
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from po_logic import (
    COLS,
    FLAG_COLORS,
    load_penjualan,
    load_penjualan_3bulan,
    load_pembelian,
    load_stok,
    dedupe_stok,
    build_dataset,
    build_deadstock,
    build_overstock,
    compute_recommendations,
)

LOGO_PATH = "assets/logo.jpg"
TEAL = "#186156"
LIME = "#ADF901"

MENU_WELCOME = "🏠 Welcome"
MENU_PO = "📊 PO Recommendation"
MENU_DEADSTOCK = "🧟 Deadstock"
MENU_OVERSTOCK = "📈 Overstock"

st.set_page_config(
    page_title="K. Beauty - Dashboard",
    page_icon=LOGO_PATH,
    layout="wide",
)

# ============================================================
# BRAND HEADER + THEME POLISH
# (base colors/backgrounds come from .streamlit/config.toml - this just
# adds a few accents that Streamlit's theme options don't cover)
# ============================================================
st.markdown(
    f"""
    <style>
    div[data-testid="stMetricValue"] {{ color: {TEAL}; }}
    h2, h3 {{ color: {TEAL}; }}
    .stTabs [aria-selected="true"] {{
        border-bottom: 3px solid {LIME} !important;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

logo_col, title_col = st.columns([1, 4])
with logo_col:
    st.image(LOGO_PATH, use_container_width=True)
with title_col:
    st.markdown(
        f"""
        <div style="padding-top: 22px;">
            <span style="font-size:26px; font-weight:700; color:{TEAL};">Inventory & PO Dashboard</span><br>
            <span style="font-size:15px; color:#5b8f83;">K. Beauty &middot; The Beauty Shop</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
st.markdown(
    f'<hr style="border:none; height:4px; background:linear-gradient(90deg,{TEAL},{LIME}); '
    'border-radius:4px; margin-top:6px; margin-bottom:18px;">',
    unsafe_allow_html=True,
)

# ============================================================
# SIDEBAR - menu + file uploaders
# ============================================================

st.sidebar.image(LOGO_PATH, use_container_width=True)
st.sidebar.title("📦 K. Beauty Dashboard")

menu = st.sidebar.radio(
    "Menu", [MENU_WELCOME, MENU_PO, MENU_DEADSTOCK, MENU_OVERSTOCK], label_visibility="collapsed"
)

st.sidebar.divider()
st.sidebar.subheader("📁 Upload Data")

stok_file = st.sidebar.file_uploader(
    "Data Stok (dipakai di semua menu)",
    type=["csv", "xlsx", "xls"], key="stok_file",
)
penjualan_30h_file = st.sidebar.file_uploader(
    "Data Penjualan 30 Hari (untuk PO Recommendation)",
    type=["csv", "xlsx", "xls"], key="p30_file",
)
penjualan_3bulan_file = st.sidebar.file_uploader(
    "Data Penjualan 3 Bulan (untuk Deadstock & Overstock)",
    type=["csv", "xlsx", "xls"], key="p3b_file",
)
pembelian_file = st.sidebar.file_uploader(
    "Data Pembelian (opsional, untuk exclude produk baru dari Deadstock)",
    type=["csv", "xlsx", "xls"], key="pembelian_file",
)

st.sidebar.caption(
    "Upload export asli iPOS 5.0 apa adanya - tidak perlu edit/rapikan file dulu."
)

# ============================================================
# PARSE UPLOADED FILES (once, shared across menus)
# ============================================================

stok_df, stok_error = None, None
if stok_file:
    try:
        stok_df = load_stok(stok_file)
    except Exception as e:
        stok_error = str(e)

penjualan_30h_df, p30_error = None, None
if penjualan_30h_file:
    try:
        penjualan_30h_df = load_penjualan(penjualan_30h_file)
    except Exception as e:
        p30_error = str(e)

penjualan_3bulan_df, p3b_error = None, None
if penjualan_3bulan_file:
    try:
        penjualan_3bulan_df = load_penjualan_3bulan(penjualan_3bulan_file)
    except Exception as e:
        p3b_error = str(e)

pembelian_df, pembelian_error = None, None
if pembelian_file:
    try:
        pembelian_df = load_pembelian(pembelian_file)
    except Exception as e:
        pembelian_error = str(e)


# ============================================================
# SHARED EXCEL HELPERS
# ============================================================

def autosize_and_freeze(ws, dframe):
    ws.freeze_panes = "A2"
    for i, col in enumerate(dframe.columns, start=1):
        width = max(12, min(40, int(dframe[col].astype(str).str.len().max() or 10) + 2))
        ws.column_dimensions[get_column_letter(i)].width = width


def write_sheet(wb, sheet_name, dframe, header_font, header_fill, flag_col=None):
    safe_name = str(sheet_name)[:31] or "Sheet"
    ws = wb.create_sheet(safe_name)
    ws.append(list(dframe.columns))
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    flag_idx = list(dframe.columns).index(flag_col) + 1 if flag_col and flag_col in dframe.columns else None
    for _, row in dframe.iterrows():
        ws.append(list(row))
        if flag_idx:
            color = FLAG_COLORS.get(row[flag_col])
            if color:
                ws.cell(row=ws.max_row, column=flag_idx).fill = PatternFill("solid", fgColor=color)
    autosize_and_freeze(ws, dframe)


# ============================================================
# MENU 1: WELCOME
# ============================================================

def render_welcome():
    st.header("Selamat datang 👋")

    if stok_error:
        st.error(f"Gagal membaca Data Stok: {stok_error}")
        return

    if stok_df is None:
        st.info(
            "Upload minimal **Data Stok** di sidebar untuk mulai. Tambahkan juga "
            "Data Penjualan 30 Hari (untuk menu PO Recommendation) dan/atau Data "
            "Penjualan 3 Bulan (untuk menu Sales Recap) sesuai kebutuhan."
        )
        st.markdown(
            """
**Cara pakai:**
1. Upload **Data Stok** (Daftar Item) di sidebar - dipakai di semua menu.
2. Untuk rekomendasi PO mingguan: tambahkan **Data Penjualan 30 Hari**, lalu
   buka menu **📊 PO Recommendation**.
3. Untuk cari barang yang tidak laku / produk fokus: tambahkan **Data
   Penjualan 3 Bulan** (dan opsional **Data Pembelian** biar produk baru
   tidak ikut ke-flag), lalu buka menu **🧟 Deadstock** atau **📈 Overstock**.

Semua file diupload apa adanya dari export iPOS 5.0 (*Daftar Item* untuk
stok, *Daftar Penjualan per Item per Jenis* untuk penjualan, *Daftar
Pembelian per Item per Jenis* untuk pembelian) - tidak perlu dirapikan dulu.
            """
        )
        return

    missing = []
    if COLS["product_id"] not in stok_df.columns:
        missing.append("Kolom Kode Item / Product ID tidak ditemukan di Data Stok")
    if COLS["stok"] not in stok_df.columns:
        missing.append("Kolom Stok tidak ditemukan di Data Stok")
    if missing:
        st.error("Kolom tidak ditemukan:\n\n" + "\n".join(f"- {m}" for m in missing))
        return

    stok_clean, dup_stok = dedupe_stok(stok_df)
    total_sku = len(stok_clean)
    total_qty_stok = int(stok_clean[COLS["stok"]].sum())
    total_brand = stok_clean[COLS["brand"]].nunique() if COLS["brand"] in stok_clean.columns else 0

    if dup_stok > 0:
        st.caption(f"🧹 {dup_stok} baris duplikat di Data Stok sudah digabung otomatis.")

    col1, col2, col3 = st.columns(3)
    col1.metric("Total SKU (Master Stok)", f"{total_sku:,}")
    col2.metric("Total Qty Stok", f"{total_qty_stok:,}")
    col3.metric("Total Brand", f"{total_brand:,}")

    st.divider()
    ready_po = "✅" if penjualan_30h_df is not None else "⬜"
    ready_recap = "✅" if penjualan_3bulan_df is not None else "⬜"
    st.markdown(
        f"""
- {ready_po} **PO Recommendation** - {"siap dibuka" if penjualan_30h_df is not None else "upload Data Penjualan 30 Hari dulu"}
- {ready_recap} **Deadstock** - {"siap dibuka" if penjualan_3bulan_df is not None else "upload Data Penjualan 3 Bulan dulu"} {"(+ Data Pembelian opsional sudah terpasang)" if penjualan_3bulan_df is not None and pembelian_df is not None else ""}
- {ready_recap} **Overstock** - {"siap dibuka" if penjualan_3bulan_df is not None else "upload Data Penjualan 3 Bulan dulu"}
        """
    )


# ============================================================
# MENU 2: PO RECOMMENDATION
# ============================================================

def render_po_recommendation():
    st.header("📊 PO Recommendation")

    if stok_error:
        st.error(f"Gagal membaca Data Stok: {stok_error}")
        return
    if p30_error:
        st.error(f"Gagal membaca Data Penjualan 30 Hari: {p30_error}")
        return
    if stok_df is None or penjualan_30h_df is None:
        st.info("Upload **Data Stok** dan **Data Penjualan 30 Hari** di sidebar untuk buka menu ini.")
        return

    missing = []
    if COLS["product_id"] not in penjualan_30h_df.columns:
        missing.append("Kolom Kode Item / Product ID tidak ditemukan di file penjualan 30 hari")
    if COLS["qty_terjual_30h"] not in penjualan_30h_df.columns:
        missing.append("Kolom Jumlah / Qty Terjual tidak ditemukan di file penjualan 30 hari")
    if COLS["product_id"] not in stok_df.columns:
        missing.append("Kolom Kode Item / Product ID tidak ditemukan di Data Stok")
    if COLS["stok"] not in stok_df.columns:
        missing.append("Kolom Stok tidak ditemukan di Data Stok")
    if missing:
        st.error("Kolom tidak ditemukan:\n\n" + "\n".join(f"- {m}" for m in missing))
        return

    df, dup_report = build_dataset(penjualan_30h_df, stok_df)
    df = compute_recommendations(df)

    tab_current, tab_brand = st.tabs(["📊 Minggu Ini", "📦 Analitik per Brand"])

    # --------------------------------------------------------
    # TAB 1: MINGGU INI
    # --------------------------------------------------------
    with tab_current:
        total_dupes = dup_report["penjualan_duplicates_merged"] + dup_report["stok_duplicates_merged"]
        if total_dupes > 0:
            st.warning(
                f"🧹 Ditemukan & digabung {dup_report['penjualan_duplicates_merged']} baris duplikat di "
                f"data penjualan dan {dup_report['stok_duplicates_merged']} di data stok "
                f"(Product ID yang sama muncul lebih dari sekali - qty-nya sudah dijumlahkan)."
            )

        only_penjualan = int((df["Sumber Data"] == "Hanya ada di data penjualan (tidak ada di data stok)").sum())
        if only_penjualan > 0:
            st.info(
                f"ℹ️ {only_penjualan} SKU terjual dalam 30 hari terakhir tapi tidak ditemukan di data stok "
                f"(dianggap stok = 0 - cek apakah barang ini masih aktif/terdaftar)."
            )

        counts = df["Flag"].value_counts()
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total SKU", len(df))
        col2.metric("🔴 Urgent", int(counts.get("🔴 URGENT", 0)))
        col3.metric("🟡 Check Sales", int(counts.get("🟡 CHECK SALES", 0)))
        col4.metric("🟢 Normal", int(counts.get("🟢 NORMAL", 0)))

        if df["PO_Value_Rp"].sum() > 0:
            st.metric("Estimasi Nilai PO (harga pokok)", f"Rp {df['PO_Value_Rp'].sum():,.0f}")

        st.bar_chart(counts)

        st.subheader("Detail per SKU")

        filter_cols = st.columns(3)
        flag_filter = filter_cols[0].multiselect(
            "Filter status", options=list(FLAG_COLORS.keys()), default=list(FLAG_COLORS.keys())
        )
        brand_options = ["Semua"] + sorted(df[COLS["brand"]].dropna().unique().tolist()) \
            if COLS["brand"] in df.columns else ["Semua"]
        brand_filter = filter_cols[1].selectbox("Filter brand/jenis", brand_options, key="po_brand_filter")
        only_need_po = filter_cols[2].checkbox("Hanya tampilkan yang perlu PO", value=False)

        view = df[df["Flag"].isin(flag_filter)]
        if brand_filter != "Semua" and COLS["brand"] in df.columns:
            view = view[view[COLS["brand"]] == brand_filter]
        if only_need_po:
            view = view[view["PO_Qty"] > 0]

        display_cols = [c for c in [
            COLS["product_id"], COLS["nama_barang"], COLS.get("brand"),
            COLS["qty_terjual_30h"], COLS["stok"], "PO_Qty", "PO_Value_Rp", "Flag", "Alasan",
        ] if c in view.columns]

        def highlight_flag(row):
            color = FLAG_COLORS.get(row["Flag"])
            return [f"background-color: #{color}" if color else "" for _ in row]

        st.dataframe(
            view[display_cols].style.apply(highlight_flag, axis=1),
            use_container_width=True,
            height=450,
        )

        urgent_items = df[df["Flag"] == "🔴 URGENT"]
        if len(urgent_items) > 0:
            with st.expander(f"⚠️ {len(urgent_items)} item butuh perhatian segera"):
                st.dataframe(urgent_items[display_cols], use_container_width=True)

        # ----------------------------------------------------
        # EXCEL EXPORT (PO Value tidak diikutkan)
        # ----------------------------------------------------
        def build_po_excel_bytes(dframe) -> bytes:
            wb = Workbook()
            wb.remove(wb.active)
            header_font = Font(name="Arial", bold=True, color="FFFFFF")
            header_fill = PatternFill("solid", fgColor="186156")

            ws = wb.create_sheet("Summary", 0)
            ws["A1"] = "PO Recommendation Summary - K. Beauty"
            ws["A1"].font = Font(name="Arial", bold=True, size=14, color="186156")
            ws["A3"], ws["B3"] = "Total SKU", len(dframe)
            ws["A4"], ws["B4"] = "🔴 Urgent", int((dframe["Flag"] == "🔴 URGENT").sum())
            ws["A5"], ws["B5"] = "🟡 Check Sales", int((dframe["Flag"] == "🟡 CHECK SALES").sum())
            ws["A6"], ws["B6"] = "🟢 Normal", int((dframe["Flag"] == "🟢 NORMAL").sum())
            for col, width in zip("AB", (32, 20)):
                ws.column_dimensions[col].width = width

            all_cols = [c for c in [
                COLS["product_id"], COLS["nama_barang"], COLS.get("brand"),
                COLS["qty_terjual_30h"], COLS["stok"], "PO_Qty", "Flag", "Alasan",
            ] if c in dframe.columns]
            write_sheet(wb, "Semua SKU", dframe[all_cols], header_font, header_fill, flag_col="Flag")

            need_po = dframe[dframe["PO_Qty"] > 0]
            if COLS["brand"] in dframe.columns and len(need_po) > 0:
                for brand, group in need_po.groupby(COLS["brand"]):
                    write_sheet(wb, str(brand), group[all_cols], header_font, header_fill, flag_col="Flag")

            buffer = io.BytesIO()
            wb.save(buffer)
            return buffer.getvalue()

        st.divider()
        excel_bytes = build_po_excel_bytes(df)
        st.download_button(
            "⬇️ Download Excel (Summary + Semua SKU + per Brand)",
            data=excel_bytes,
            file_name="PO_Recommendation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="po_download",
        )

    # --------------------------------------------------------
    # TAB 2: ANALITIK PER BRAND
    # --------------------------------------------------------
    with tab_brand:
        st.subheader("📦 Sales & Stock Analytics per Brand - Minggu Ini")

        if COLS["brand"] not in df.columns or df[COLS["brand"]].dropna().empty:
            st.caption("Kolom Brand/Jenis tidak ditemukan di data yang diupload.")
        else:
            brand_agg = df.groupby(COLS["brand"]).agg(
                Total_SKU=(COLS["product_id"], "count"),
                Qty_Terjual=(COLS["qty_terjual_30h"], "sum"),
                Total_Stok=(COLS["stok"], "sum"),
                PO_Qty=("PO_Qty", "sum"),
                Urgent=("Flag", lambda s: (s == "🔴 URGENT").sum()),
                Check_Sales=("Flag", lambda s: (s == "🟡 CHECK SALES").sum()),
            ).reset_index().rename(columns={COLS["brand"]: "Brand"})
            brand_agg = brand_agg.sort_values("Qty_Terjual", ascending=False)

            top_n = st.slider("Tampilkan berapa brand teratas", 5, min(40, len(brand_agg)),
                               value=min(15, len(brand_agg)))
            brand_view = brand_agg.head(top_n)

            chart_cols = st.columns(2)
            with chart_cols[0]:
                st.markdown("**Qty Terjual 30 Hari per Brand**")
                st.bar_chart(brand_view.set_index("Brand")["Qty_Terjual"], color=TEAL)
            with chart_cols[1]:
                st.markdown("**Total Stok per Brand**")
                st.bar_chart(brand_view.set_index("Brand")["Total_Stok"], color=LIME)

            st.markdown("**PO yang direkomendasikan per Brand**")
            st.bar_chart(brand_view.set_index("Brand")["PO_Qty"], color=TEAL)

            st.markdown("**Tabel lengkap per Brand**")
            st.dataframe(
                brand_agg.rename(columns={
                    "Qty_Terjual": "Qty Terjual 30 Hari", "Total_Stok": "Total Stok",
                    "PO_Qty": "PO Qty Direkomendasikan", "Check_Sales": "Check Sales",
                }),
                use_container_width=True, height=400,
            )


# ============================================================
# MENU 3: DEADSTOCK (zero sales in 3 months, excluding new arrivals)
# ============================================================

def render_deadstock():
    st.header("🧟 Deadstock")
    st.caption(
        "Barang dengan stok masih ada TAPI tidak ada penjualan sama sekali "
        "dalam 3 bulan terakhir - kandidat produk fokus / clearance. Produk "
        "baru yang baru direstock (dilihat dari Data Pembelian) tidak ikut "
        "dihitung, karena wajar belum sempat laku."
    )

    if stok_error:
        st.error(f"Gagal membaca Data Stok: {stok_error}")
        return
    if p3b_error:
        st.error(f"Gagal membaca Data Penjualan 3 Bulan: {p3b_error}")
        return
    if pembelian_error:
        st.error(f"Gagal membaca Data Pembelian: {pembelian_error}")
        return
    if stok_df is None or penjualan_3bulan_df is None:
        st.info("Upload **Data Stok** dan **Data Penjualan 3 Bulan** di sidebar untuk buka menu ini.")
        return

    missing = []
    if COLS["product_id"] not in penjualan_3bulan_df.columns:
        missing.append("Kolom Kode Item / Product ID tidak ditemukan di file penjualan 3 bulan")
    if COLS["qty_terjual_3bulan"] not in penjualan_3bulan_df.columns:
        missing.append("Kolom Jumlah / Qty Terjual tidak ditemukan di file penjualan 3 bulan")
    if COLS["product_id"] not in stok_df.columns:
        missing.append("Kolom Kode Item / Product ID tidak ditemukan di Data Stok")
    if COLS["stok"] not in stok_df.columns:
        missing.append("Kolom Stok tidak ditemukan di Data Stok")
    if missing:
        st.error("Kolom tidak ditemukan:\n\n" + "\n".join(f"- {m}" for m in missing))
        return

    if pembelian_df is None:
        st.info(
            "ℹ️ Data Pembelian belum diupload - Deadstock dihitung tanpa exclude produk baru. "
            "Upload di sidebar (opsional) kalau mau produk yang baru direstock tidak ikut ke-flag."
        )

    dead_stock, dup_report = build_deadstock(penjualan_3bulan_df, stok_df, pembelian_df=pembelian_df)

    total_dupes = dup_report["penjualan_3bulan_duplicates_merged"] + dup_report["stok_duplicates_merged"] + dup_report["pembelian_duplicates_merged"]
    if total_dupes > 0:
        st.warning(
            f"🧹 Ditemukan & digabung {dup_report['penjualan_3bulan_duplicates_merged']} baris duplikat "
            f"di data penjualan 3 bulan, {dup_report['stok_duplicates_merged']} di data stok, dan "
            f"{dup_report['pembelian_duplicates_merged']} di data pembelian."
        )
    if dup_report["excluded_new_arrivals"] > 0:
        st.success(
            f"✅ {dup_report['excluded_new_arrivals']} SKU dikecualikan dari Deadstock karena baru "
            f"direstock (ada di Data Pembelian) - dianggap belum sempat laku, bukan barang mati."
        )

    col1, col2 = st.columns(2)
    col1.metric("Total SKU Deadstock", f"{len(dead_stock):,}")
    col2.metric("Total Qty Stok Menumpuk", f"{int(dead_stock[COLS['stok']].sum()):,}" if len(dead_stock) else "0")

    if dead_stock.empty:
        st.success("Tidak ada barang deadstock - semua SKU dengan stok pernah terjual dalam 3 bulan terakhir. 🎉")
        return

    tab_list, tab_brand = st.tabs(["📋 Daftar Deadstock", "📦 Analitik per Brand"])

    # --------------------------------------------------------
    # TAB 1: DAFTAR DEADSTOCK (filter + tabel + download)
    # --------------------------------------------------------
    with tab_list:
        filter_cols = st.columns(2)
        staff_options = ["Semua"] + sorted(dead_stock[COLS["sales_staff"]].dropna().unique().tolist())
        staff_filter = filter_cols[0].selectbox("Filter Sales Staff", staff_options, key="dead_staff_filter")
        brand_options = ["Semua"] + sorted(dead_stock[COLS["brand"]].dropna().unique().tolist()) \
            if COLS["brand"] in dead_stock.columns else ["Semua"]
        brand_filter = filter_cols[1].selectbox("Filter Brand", brand_options, key="dead_brand_filter")

        view = dead_stock
        if staff_filter != "Semua":
            view = view[view[COLS["sales_staff"]] == staff_filter]
        if brand_filter != "Semua" and COLS["brand"] in dead_stock.columns:
            view = view[view[COLS["brand"]] == brand_filter]

        st.caption(f"Menampilkan {len(view)} SKU - sudah otomatis diurutkan per Sales Staff, Brand, lalu Nama Barang.")

        display_cols = [c for c in [
            COLS["product_id"], COLS["nama_barang"], COLS.get("brand"), COLS["sales_staff"],
            COLS["stok"], COLS["qty_terjual_3bulan"],
        ] if c in view.columns]

        st.dataframe(view[display_cols], use_container_width=True, height=500)

        # ------------------------------------------------
        # EXCEL EXPORT - 1 sheet per Sales Staff
        # ------------------------------------------------
        def build_deadstock_excel_bytes(dframe) -> bytes:
            wb = Workbook()
            wb.remove(wb.active)
            header_font = Font(name="Arial", bold=True, color="FFFFFF")
            header_fill = PatternFill("solid", fgColor="186156")

            ws = wb.create_sheet("Summary", 0)
            ws["A1"] = "Deadstock - K. Beauty"
            ws["A1"].font = Font(name="Arial", bold=True, size=14, color="186156")
            ws["A3"], ws["B3"] = "Total SKU Deadstock", len(dframe)
            ws["A4"], ws["B4"] = "Total Qty Stok Menumpuk", int(dframe[COLS["stok"]].sum())
            for col, width in zip("AB", (32, 20)):
                ws.column_dimensions[col].width = width

            export_cols = [c for c in [
                COLS["product_id"], COLS["nama_barang"], COLS.get("brand"),
                COLS["stok"], COLS["qty_terjual_3bulan"],
            ] if c in dframe.columns]
            write_sheet(wb, "Semua Deadstock", dframe[export_cols], header_font, header_fill)

            for staff, group in dframe.groupby(COLS["sales_staff"]):
                write_sheet(wb, str(staff), group[export_cols], header_font, header_fill)

            buffer = io.BytesIO()
            wb.save(buffer)
            return buffer.getvalue()

        st.divider()
        excel_bytes = build_deadstock_excel_bytes(dead_stock)
        st.download_button(
            "⬇️ Download Excel (Semua Deadstock + per Sales Staff)",
            data=excel_bytes,
            file_name="Deadstock.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dead_download",
        )

    # --------------------------------------------------------
    # TAB 2: ANALITIK PER BRAND (Total SKU vs Total Qty per brand)
    # --------------------------------------------------------
    with tab_brand:
        st.subheader("📦 Deadstock per Brand")

        if COLS["brand"] not in dead_stock.columns or dead_stock[COLS["brand"]].dropna().empty:
            st.caption("Kolom Brand/Jenis tidak ditemukan di data yang diupload.")
        else:
            brand_agg = dead_stock.groupby(COLS["brand"]).agg(
                Total_SKU=(COLS["product_id"], "count"),
                Total_Qty=(COLS["stok"], "sum"),
            ).reset_index().rename(columns={COLS["brand"]: "Brand"})
            brand_agg["Rata_rata_Qty_per_SKU"] = (brand_agg["Total_Qty"] / brand_agg["Total_SKU"]).round(1)
            brand_agg = brand_agg.sort_values("Total_SKU", ascending=False)

            st.caption(
                "**Total SKU** = berapa jenis barang mati per brand. **Total Qty** = berapa pcs fisik yang "
                "menumpuk. Brand dengan SKU banyak tapi rata-rata qty/SKU kecil → banyak varian yang perlu "
                "dirasionalisasi. Brand dengan qty besar di sedikit SKU → cek barang spesifiknya untuk clearance."
            )

            sort_by = st.radio(
                "Urutkan berdasarkan", ["Total SKU", "Total Qty"], horizontal=True, key="dead_sort_by"
            )
            sort_col = "Total_SKU" if sort_by == "Total SKU" else "Total_Qty"
            brand_view = brand_agg.sort_values(sort_col, ascending=False)

            top_n = st.slider("Tampilkan berapa brand teratas", 5, min(40, len(brand_view)),
                               value=min(15, len(brand_view)), key="dead_top_n")
            brand_view = brand_view.head(top_n)

            chart_cols = st.columns(2)
            with chart_cols[0]:
                st.markdown("**Total SKU Deadstock per Brand**")
                st.bar_chart(brand_view.set_index("Brand")["Total_SKU"], color=TEAL)
            with chart_cols[1]:
                st.markdown("**Total Qty Menumpuk per Brand**")
                st.bar_chart(brand_view.set_index("Brand")["Total_Qty"], color=LIME)

            st.markdown("**Tabel lengkap per Brand**")
            st.dataframe(
                brand_agg.rename(columns={
                    "Total_SKU": "Total SKU Deadstock",
                    "Total_Qty": "Total Qty Stok Menumpuk",
                    "Rata_rata_Qty_per_SKU": "Rata-rata Qty per SKU",
                }),
                use_container_width=True, height=400,
            )


# ============================================================
# MENU 4: OVERSTOCK (stock > 3-month sales)
# ============================================================

def render_overstock():
    st.header("📈 Overstock")
    st.caption(
        "Barang dengan stok LEBIH BANYAK dari penjualan 3 bulan terakhir - "
        "termasuk yang sempat laku sedikit tapi masih kelebihan stok jauh. "
        "Lebih luas dari Deadstock (yang cuma tangkap barang 0 penjualan)."
    )

    if stok_error:
        st.error(f"Gagal membaca Data Stok: {stok_error}")
        return
    if p3b_error:
        st.error(f"Gagal membaca Data Penjualan 3 Bulan: {p3b_error}")
        return
    if stok_df is None or penjualan_3bulan_df is None:
        st.info("Upload **Data Stok** dan **Data Penjualan 3 Bulan** di sidebar untuk buka menu ini.")
        return

    missing = []
    if COLS["product_id"] not in penjualan_3bulan_df.columns:
        missing.append("Kolom Kode Item / Product ID tidak ditemukan di file penjualan 3 bulan")
    if COLS["qty_terjual_3bulan"] not in penjualan_3bulan_df.columns:
        missing.append("Kolom Jumlah / Qty Terjual tidak ditemukan di file penjualan 3 bulan")
    if COLS["product_id"] not in stok_df.columns:
        missing.append("Kolom Kode Item / Product ID tidak ditemukan di Data Stok")
    if COLS["stok"] not in stok_df.columns:
        missing.append("Kolom Stok tidak ditemukan di Data Stok")
    if missing:
        st.error("Kolom tidak ditemukan:\n\n" + "\n".join(f"- {m}" for m in missing))
        return

    overstock, dup_report = build_overstock(penjualan_3bulan_df, stok_df)

    total_dupes = dup_report["penjualan_3bulan_duplicates_merged"] + dup_report["stok_duplicates_merged"]
    if total_dupes > 0:
        st.warning(
            f"🧹 Ditemukan & digabung {dup_report['penjualan_3bulan_duplicates_merged']} baris duplikat "
            f"di data penjualan 3 bulan dan {dup_report['stok_duplicates_merged']} di data stok."
        )

    col1, col2 = st.columns(2)
    col1.metric("Total SKU Overstock", f"{len(overstock):,}")
    col2.metric("Total Selisih Stok-Penjualan", f"{int(overstock['Selisih Stok-Penjualan'].sum()):,}" if len(overstock) else "0")

    if overstock.empty:
        st.success("Tidak ada barang overstock. 🎉")
        return

    tab_list, tab_brand = st.tabs(["📋 Daftar Overstock", "📦 Analitik per Brand"])

    # --------------------------------------------------------
    # TAB 1: DAFTAR OVERSTOCK (filter + tabel + download)
    # --------------------------------------------------------
    with tab_list:
        filter_cols = st.columns(2)
        staff_options = ["Semua"] + sorted(overstock[COLS["sales_staff"]].dropna().unique().tolist())
        staff_filter = filter_cols[0].selectbox("Filter Sales Staff", staff_options, key="over_staff_filter")
        brand_options = ["Semua"] + sorted(overstock[COLS["brand"]].dropna().unique().tolist()) \
            if COLS["brand"] in overstock.columns else ["Semua"]
        brand_filter = filter_cols[1].selectbox("Filter Brand", brand_options, key="over_brand_filter")

        view = overstock
        if staff_filter != "Semua":
            view = view[view[COLS["sales_staff"]] == staff_filter]
        if brand_filter != "Semua" and COLS["brand"] in overstock.columns:
            view = view[view[COLS["brand"]] == brand_filter]

        st.caption(f"Menampilkan {len(view)} SKU - sudah otomatis diurutkan per Sales Staff, Brand, lalu Nama Barang.")

        display_cols = [c for c in [
            COLS["product_id"], COLS["nama_barang"], COLS.get("brand"), COLS["sales_staff"],
            COLS["stok"], COLS["qty_terjual_3bulan"], "Selisih Stok-Penjualan",
        ] if c in view.columns]

        st.dataframe(
            view[display_cols].sort_values("Selisih Stok-Penjualan", ascending=False)
            if staff_filter != "Semua" or brand_filter != "Semua" else view[display_cols],
            use_container_width=True, height=500,
        )

        # ------------------------------------------------
        # EXCEL EXPORT - 1 sheet per Sales Staff
        # ------------------------------------------------
        def build_overstock_excel_bytes(dframe) -> bytes:
            wb = Workbook()
            wb.remove(wb.active)
            header_font = Font(name="Arial", bold=True, color="FFFFFF")
            header_fill = PatternFill("solid", fgColor="186156")

            ws = wb.create_sheet("Summary", 0)
            ws["A1"] = "Overstock - K. Beauty"
            ws["A1"].font = Font(name="Arial", bold=True, size=14, color="186156")
            ws["A3"], ws["B3"] = "Total SKU Overstock", len(dframe)
            ws["A4"], ws["B4"] = "Total Selisih Stok-Penjualan", int(dframe["Selisih Stok-Penjualan"].sum())
            for col, width in zip("AB", (32, 20)):
                ws.column_dimensions[col].width = width

            export_cols = [c for c in [
                COLS["product_id"], COLS["nama_barang"], COLS.get("brand"),
                COLS["stok"], COLS["qty_terjual_3bulan"], "Selisih Stok-Penjualan",
            ] if c in dframe.columns]
            write_sheet(wb, "Semua Overstock", dframe[export_cols], header_font, header_fill)

            for staff, group in dframe.groupby(COLS["sales_staff"]):
                write_sheet(wb, str(staff), group[export_cols], header_font, header_fill)

            buffer = io.BytesIO()
            wb.save(buffer)
            return buffer.getvalue()

        st.divider()
        excel_bytes = build_overstock_excel_bytes(overstock)
        st.download_button(
            "⬇️ Download Excel (Semua Overstock + per Sales Staff)",
            data=excel_bytes,
            file_name="Overstock.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="over_download",
        )

    # --------------------------------------------------------
    # TAB 2: ANALITIK PER BRAND
    # --------------------------------------------------------
    with tab_brand:
        st.subheader("📦 Overstock per Brand")

        if COLS["brand"] not in overstock.columns or overstock[COLS["brand"]].dropna().empty:
            st.caption("Kolom Brand/Jenis tidak ditemukan di data yang diupload.")
        else:
            brand_agg = overstock.groupby(COLS["brand"]).agg(
                Total_SKU=(COLS["product_id"], "count"),
                Total_Selisih=("Selisih Stok-Penjualan", "sum"),
            ).reset_index().rename(columns={COLS["brand"]: "Brand"})
            brand_agg = brand_agg.sort_values("Total_Selisih", ascending=False)

            top_n = st.slider("Tampilkan berapa brand teratas", 5, min(40, len(brand_agg)),
                               value=min(15, len(brand_agg)), key="over_top_n")
            brand_view = brand_agg.head(top_n)

            chart_cols = st.columns(2)
            with chart_cols[0]:
                st.markdown("**Total SKU Overstock per Brand**")
                st.bar_chart(brand_view.set_index("Brand")["Total_SKU"], color=TEAL)
            with chart_cols[1]:
                st.markdown("**Total Selisih Stok-Penjualan per Brand**")
                st.bar_chart(brand_view.set_index("Brand")["Total_Selisih"], color=LIME)

            st.markdown("**Tabel lengkap per Brand**")
            st.dataframe(
                brand_agg.rename(columns={
                    "Total_SKU": "Total SKU Overstock",
                    "Total_Selisih": "Total Selisih Stok-Penjualan",
                }),
                use_container_width=True, height=400,
            )


# ============================================================
# ROUTE TO SELECTED MENU
# ============================================================

if menu == MENU_WELCOME:
    render_welcome()
elif menu == MENU_PO:
    render_po_recommendation()
elif menu == MENU_DEADSTOCK:
    render_deadstock()
elif menu == MENU_OVERSTOCK:
    render_overstock()
