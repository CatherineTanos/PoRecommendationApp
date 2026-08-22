"""
po_logic.py
===========
Core data logic for the PO recommendation engine.

Reads the 2 raw exports straight from iPOS 5.0 ("Daftar Penjualan per Item
per Jenis" and "Daftar Item" / stock list) - the messy report layout with
merged cells, spacer rows, and "Jenis :" / "Total :" section rows - no manual
cleanup needed before uploading.

RULES (per your latest instructions)
-------------------------------------
1. No fixed SLOW_MOVING_THRESHOLD number. A SKU is flagged slow moving
   whenever Stok == Qty Terjual 30 Hari (exact match, including 0 == 0
   for a dead SKU with no stock and no sales).
2. Only 2 input files: data penjualan (30 hari) & data stok. No master
   barang file.
3. The two files are joined on Product ID ("Kode Item" in your iPOS export).
4. Duplicate rows for the same Product ID (within a file) are detected and
   merged before joining, so nothing gets double counted.
5. PO formula:  PO = kebutuhan 30 hari (penjualan) - stok
   (no "on order" term - simplified, no overstock-ratio override either,
   since with this formula stok > penjualan already always yields PO = 0
   on its own, so a separate overstock rule would be redundant.)
"""

import re
import pandas as pd

# ============================================================
# CONFIG
# ============================================================

# Standardized column names used internally after parsing/merging.
COLS = {
    "product_id": "Product ID",
    "nama_barang": "Nama Barang",
    "brand": "Brand",       # from "Jenis" in the iPOS export
    "harga": "Harga",       # cost price ("Harga Pokok") from the stok file
    "qty_terjual_30h": "Qty Terjual 30 Hari",
    "stok": "Stok",
}

FLAG_COLORS = {
    "🔴 URGENT": "FFC7CE",
    "🟡 CHECK SALES": "FFEB9C",
    "🟢 NORMAL": "C6EFCE",
}

# Days-of-stock-left below which an item still gets flagged urgent even
# though stock isn't literally zero yet.
URGENT_DAYS_LEFT = 3


# ============================================================
# LOW-LEVEL: locate + parse the raw iPOS report layout
# ============================================================

def _norm_id(x):
    """Normalize a Product ID / Kode Item value to a clean string.
    Handles floats like 4545593011270.0 (Excel reading long barcodes as
    numbers) as well as alphanumeric codes like 'LENCIR01'."""
    if pd.isna(x):
        return None
    if isinstance(x, float):
        if x.is_integer():
            return str(int(x))
        return str(x)
    return str(x).strip()


def _norm_header(x):
    if not isinstance(x, str):
        return None
    return re.sub(r"\s+", " ", x).strip().lower()


def _find_header_rows(raw: pd.DataFrame, id_label="kode item"):
    """Return every row index where a cell exactly matches id_label.
    Supports exports that repeat the header per store/section - each
    occurrence starts a new data block."""
    rows = []
    for r in range(len(raw)):
        row = raw.iloc[r]
        for v in row:
            if _norm_header(v) == id_label:
                rows.append(r)
                break
    return rows


def _find_col(raw: pd.DataFrame, header_row: int, labels, contains=False):
    """Find the column index in header_row matching any of `labels`
    (a string or list of candidate strings), case/space-insensitive."""
    if isinstance(labels, str):
        labels = [labels]
    labels = [_norm_header(l) for l in labels]
    row = raw.iloc[header_row]
    for c in range(len(row)):
        v = _norm_header(row.iloc[c])
        if v is None:
            continue
        if contains:
            if any(l in v for l in labels):
                return c
        else:
            if v in labels:
                return c
    return None


def _best_numeric_col(chunk: pd.DataFrame, lo: int, hi: int):
    """Among columns [lo, hi) in `chunk`, return the index with the most
    numeric (non-null) values. Used to locate a data column by its actual
    contents rather than trusting a merged header cell's position, since
    iPOS sometimes merges a header label across a slightly different
    column span than the data cells beneath it."""
    if lo is None or hi is None or hi <= lo:
        return None
    best_col, best_count = None, 0
    for c in range(lo, hi):
        count = pd.to_numeric(chunk.iloc[:, c], errors="coerce").notna().sum()
        if count > best_count:
            best_col, best_count = c, count
    return best_col


def _parse_ipos_report(raw: pd.DataFrame, qty_label, price_labels=None):
    """Generic parser for both the 'Daftar Penjualan' and 'Daftar Item'
    (stock) iPOS reports. Locates every 'Kode Item' header block, pulls out
    Product ID / Nama Barang / Brand / qty (+ optional price), and stacks
    all blocks together (handles multi-store exports too)."""
    header_rows = _find_header_rows(raw)
    if not header_rows:
        raise ValueError(
            "Kolom 'Kode Item' tidak ditemukan. Ini bukan format export iPOS "
            "yang dikenali - pastikan file yang diupload adalah export asli "
            "dari iPOS (belum diedit ulang headernya)."
        )

    blocks = []
    for i, hr in enumerate(header_rows):
        end = header_rows[i + 1] if i + 1 < len(header_rows) else len(raw)

        id_col = _find_col(raw, hr, "kode item")
        name_col = _find_col(raw, hr, "nama item")
        brand_col = _find_col(raw, hr, "jenis")
        satuan_col = _find_col(raw, hr, "satuan")
        qty_col = _find_col(raw, hr, qty_label)
        price_col = _find_col(raw, hr, price_labels, contains=False) if price_labels else None

        chunk = raw.iloc[hr + 1: end].copy()
        chunk = chunk[chunk.iloc[:, id_col].notna()]  # drops spacer/"Total :"/footer rows
        if chunk.empty:
            continue

        # The qty column's header label is sometimes merged one cell off
        # from where the actual numbers sit - re-derive it from the data
        # itself (the numeric column between "Jenis" and "Satuan") whenever
        # that anchor range is available, since it's more reliable.
        if brand_col is not None and satuan_col is not None:
            detected = _best_numeric_col(chunk, brand_col + 1, satuan_col)
            if detected is not None:
                qty_col = detected

        out = pd.DataFrame({
            COLS["product_id"]: chunk.iloc[:, id_col].map(_norm_id),
            COLS["nama_barang"]: chunk.iloc[:, name_col].astype(str).str.strip(),
        })
        if brand_col is not None:
            out[COLS["brand"]] = chunk.iloc[:, brand_col].astype(str).str.strip()
        out[qty_label.title()] = pd.to_numeric(chunk.iloc[:, qty_col], errors="coerce").fillna(0)
        if price_col is not None:
            out[COLS["harga"]] = pd.to_numeric(chunk.iloc[:, price_col], errors="coerce").fillna(0)

        blocks.append(out)

    if not blocks:
        raise ValueError("Tidak ada baris data barang yang ditemukan di file ini.")

    df = pd.concat(blocks, ignore_index=True)
    df = df[df[COLS["product_id"]].notna() & (df[COLS["product_id"]] != "")]
    return df


def _looks_like_raw_ipos(raw: pd.DataFrame) -> bool:
    return len(_find_header_rows(raw)) > 0


# ============================================================
# LOAD - one function per file, with a fallback for already-clean files
# ============================================================

def _read_raw(file) -> pd.DataFrame:
    name = getattr(file, "name", str(file)).lower()
    if hasattr(file, "seek"):
        file.seek(0)
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(file, header=None)
    return pd.read_csv(file, header=None)


def load_penjualan(file) -> pd.DataFrame:
    """Returns columns: Product ID, Nama Barang, Brand (opt), Qty Terjual 30 Hari"""
    raw = _read_raw(file)
    if _looks_like_raw_ipos(raw):
        df = _parse_ipos_report(raw, qty_label="jumlah")
        df = df.rename(columns={"Jumlah": COLS["qty_terjual_30h"]})
        return df
    # Fallback: already a clean table with headers in row 0.
    if hasattr(file, "seek"):
        file.seek(0)
    name = getattr(file, "name", str(file)).lower()
    df = pd.read_excel(file) if name.endswith((".xlsx", ".xls")) else pd.read_csv(file)
    return _coerce_clean_columns(df, qty_col_std=COLS["qty_terjual_30h"],
                                  qty_aliases=["qty terjual 30 hari", "jumlah", "qty terjual", "qty"])


def load_stok(file) -> pd.DataFrame:
    """Returns columns: Product ID, Nama Barang, Brand (opt), Stok, Harga (opt)"""
    raw = _read_raw(file)
    if _looks_like_raw_ipos(raw):
        df = _parse_ipos_report(raw, qty_label="stok", price_labels=["harga pokok", "harga jual"])
        df = df.rename(columns={"Stok": COLS["stok"]})
        return df
    if hasattr(file, "seek"):
        file.seek(0)
    name = getattr(file, "name", str(file)).lower()
    df = pd.read_excel(file) if name.endswith((".xlsx", ".xls")) else pd.read_csv(file)
    return _coerce_clean_columns(df, qty_col_std=COLS["stok"],
                                  qty_aliases=["stok", "stock", "qty stok"])


def _coerce_clean_columns(df: pd.DataFrame, qty_col_std, qty_aliases):
    """Best-effort header matching for already-tabular files (not the raw
    iPOS report), so simple CSV/XLSX exports still work."""
    rename = {}
    for col in df.columns:
        norm = _norm_header(col) or ""
        if norm in ("kode item", "product id", "sku", "barcode"):
            rename[col] = COLS["product_id"]
        elif norm in ("nama item", "nama barang", "nama produk"):
            rename[col] = COLS["nama_barang"]
        elif norm == "jenis" or norm == "brand":
            rename[col] = COLS["brand"]
        elif norm in ("harga pokok", "harga", "cost"):
            rename[col] = COLS["harga"]
        elif norm in qty_aliases:
            rename[col] = qty_col_std
    df = df.rename(columns=rename)
    if COLS["product_id"] in df.columns:
        df[COLS["product_id"]] = df[COLS["product_id"]].map(_norm_id)
    return df


# ============================================================
# DEDUPE
# ============================================================

def _dedupe(df: pd.DataFrame, sum_col) -> tuple[pd.DataFrame, int]:
    """Removes duplicate Product ID rows. Exact duplicate rows are dropped
    outright; if a Product ID still appears more than once (e.g. logged in
    two sections), the quantity column is SUMMED and everything else keeps
    the first non-null value."""
    id_col = COLS["product_id"]
    before = len(df)
    df = df.drop_duplicates()

    keep_cols = [c for c in df.columns if c not in (sum_col, id_col)]
    agg = {sum_col: "sum"} if sum_col in df.columns else {}
    agg.update({c: "first" for c in keep_cols})

    df = df.groupby(id_col, as_index=False).agg(agg)
    removed = before - len(df)
    return df, removed


def dedupe_penjualan(df: pd.DataFrame):
    return _dedupe(df, sum_col=COLS["qty_terjual_30h"])


def dedupe_stok(df: pd.DataFrame):
    return _dedupe(df, sum_col=COLS["stok"])


# ============================================================
# MERGE (join on Product ID)
# ============================================================

def build_dataset(penjualan_df: pd.DataFrame, stok_df: pd.DataFrame):
    penjualan_clean, dup_penjualan = dedupe_penjualan(penjualan_df)
    stok_clean, dup_stok = dedupe_stok(stok_df)

    df = penjualan_clean.merge(
        stok_clean, on=COLS["product_id"], how="outer",
        suffixes=("_penjualan", "_stok"), indicator=True,
    )

    for base_col in [COLS["nama_barang"], COLS["brand"]]:
        left, right = f"{base_col}_penjualan", f"{base_col}_stok"
        if left in df.columns and right in df.columns:
            df[base_col] = df[left].combine_first(df[right])
            df = df.drop(columns=[left, right])

    df[COLS["qty_terjual_30h"]] = df[COLS["qty_terjual_30h"]].fillna(0)
    df[COLS["stok"]] = df[COLS["stok"]].fillna(0)
    if COLS["harga"] not in df.columns:
        df[COLS["harga"]] = 0
    df[COLS["harga"]] = df[COLS["harga"]].fillna(0)

    df["Sumber Data"] = df["_merge"].map({
        "both": "Ada di keduanya",
        "left_only": "Hanya ada di data penjualan (tidak ada di data stok)",
        "right_only": "Hanya ada di data stok (tidak terjual 30 hari terakhir)",
    })
    df = df.drop(columns=["_merge"])

    dup_report = {
        "penjualan_duplicates_merged": dup_penjualan,
        "stok_duplicates_merged": dup_stok,
    }
    return df, dup_report


# ============================================================
# CLASSIFY + PO FORMULA
# ============================================================

def classify_row(row):
    terjual = row[COLS["qty_terjual_30h"]]
    stok = row[COLS["stok"]]

    # Rule 1 (your rule): stock exactly equals last-30-day sales -> slow moving.
    if stok == terjual:
        if terjual == 0:
            alasan = "Tidak ada penjualan 30 hari terakhir & stok 0 - kemungkinan barang mati, cek apakah masih perlu dijual"
        else:
            alasan = (
                f"Stok ({int(stok)}) sama persis dengan penjualan 30 hari terakhir "
                f"({int(terjual)} pcs) - indikasi slow moving, cek dulu sebelum PO"
            )
        return 0, "🟡 CHECK SALES", alasan

    po_raw = terjual - stok

    # Rule: stock already covers 30-day demand -> no PO needed.
    if po_raw <= 0:
        return 0, "🟢 NORMAL", "Stok masih cukup untuk kebutuhan 30 hari ke depan"

    po_qty = po_raw

    # Rule: stock at/below zero while item still sells -> urgent.
    if stok <= 0:
        return po_qty, "🔴 URGENT", "Stok habis, barang ini masih terjual dalam 30 hari terakhir"

    # Rule: stock will run out very soon at current sales pace -> urgent.
    daily_rate = terjual / 30
    days_left = stok / daily_rate if daily_rate else 999
    if days_left < URGENT_DAYS_LEFT:
        return po_qty, "🔴 URGENT", f"Stok tersisa hanya cukup untuk ~{days_left:.0f} hari"

    return po_qty, "🟢 NORMAL", "PO rutin berdasarkan pola penjualan 30 hari"


def compute_recommendations(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    results = df.apply(classify_row, axis=1, result_type="expand")
    df["PO_Qty"], df["Flag"], df["Alasan"] = results[0], results[1], results[2]
    df["PO_Value_Rp"] = df["PO_Qty"] * df[COLS["harga"]]
    return df
