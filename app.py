"""
Aplikasi Klasifikasi Inventory Movement/SKU
-------------------------------------------
Sumber data : Google Sheets (via gspread + Service Account)
Output      : 4 Klasifikasi -> High/Low Margin x Fast/Slow Moving

Alur:
1. Ambil data mentah dari Google Sheets
2. Filter berdasarkan rentang tanggal
3. Filter berdasarkan jenis penjualan (Online / Offline / Marketplace / All Sales)
4. Tampilkan DATA_RAW
5. Pivot per Brand/SKU/Nama Barang/Kategori Barang
6. Klasifikasi (margin threshold fixed 30.000, moving pakai median log(QTY+1))
7. Tampilkan tabel hasil klasifikasi
8. Tampilkan chart klasifikasi (scatter, dengan skor)
"""

import io

import streamlit as st
import pandas as pd
import numpy as np
import gspread
from google.oauth2.service_account import Credentials
import plotly.express as px

# ============================================================
# KONFIGURASI
# ============================================================
SPREADSHEET_ID = "12MKCpCDCQiQVmS81Qgj2VHqMV9nrVOCoTrE8z9V1oLM"
SHEET_NAME = "Data"  # nama tab/worksheet di dalam spreadsheet
SHEET_STOK = "Stok"  # nama tab stok di spreadsheet yang sama

# Nama kolom persis seperti di Google Sheet
COL_TANGGAL = "Tanggal"
COL_SALES = "Nama Tenaga Penjual"
COL_PELANGGAN = "Pelanggan"
COL_BRAND = "Nama Merek Barang Barang & Jasa"
COL_SKU = "Kode #"
COL_NAMA_BARANG = "Nama Barang"
COL_KATEGORI = "Nama Kategori Barang Barang & Jasa"
COL_QTY = "Kuantitas"
COL_HARGA = "@Harga"
COL_TOTAL = "Total Harga"
COL_LABA = "Laba"
COL_GP_ITEM_RAW = "Gross Profit/Item"  # kolom ini ada di data mentah, tapi akan dihitung ulang saat pivot

MARGIN_THRESHOLD = 30000  # fixed, sesuai kode klasifikasi asli

NAMA_ONLINE = ["ARUM", "LUTHFIAH WARDAH"]
NAMA_EXCLUDE_OFFLINE = ["ARUM", "LUTHFIAH WARDAH", "Internal"]
KATA_MARKETPLACE = ["shopee", "tiktok", "lazada", "blibli", "tokopedia"]

QUADRANT_COLORS = {
    "High Margin - Fast Moving": "#2E7D32",
    "High Margin - Slow Moving": "#F9A825",
    "Low Margin - Fast Moving": "#1565C0",
    "Low Margin - Slow Moving": "#C62828",
}

# Kolom yang disembunyikan dari tampilan tabel di UI
# (perhitungan tetap utuh; kolom untuk file Excel diatur terpisah di EXCEL_HIDDEN_COLS)
HIDDEN_UI_COLS = [
    COL_HARGA, COL_TOTAL, COL_LABA, COL_GP_ITEM_RAW,
    "@Harga", "Total Harga", "Laba", "Gross Profit/Item",
    "Skor Kuadran",
]


def hide_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Buang kolom sensitif dari tampilan UI saja."""
    return df.drop(columns=HIDDEN_UI_COLS, errors="ignore")


st.set_page_config(page_title="Klasifikasi Inventory Movement/SKU", layout="wide")


# ============================================================
# KONEKSI GOOGLE SHEETS
# ============================================================
@st.cache_resource(show_spinner=False)
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


@st.cache_data(ttl=300, show_spinner="Mengambil data dari Google Sheets...")
def load_data():
    client = get_gspread_client()
    sh = client.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet(SHEET_NAME)
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    return df

@st.cache_data(ttl=300, show_spinner="Mengambil data stok dari Google Sheets...")
def load_stok() -> pd.DataFrame:
    """
    Ambil sheet 'Stok' berdasarkan posisi kolom:
    A = Kode Barang (SKU), B = Nama Barang, C = Satuan, D = Stok Dapat Dijual.
    Hasil: 1 baris per SKU dengan kolom ['SKU', 'Stok'].
    """
    client = get_gspread_client()
    ws = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_STOK)
    values = ws.get_all_values()

    # baris pertama = header, ambil kolom A-D saja
    rows = [(r + [""] * 4)[:4] for r in values[1:]]
    df = pd.DataFrame(rows, columns=["SKU", "Nama Barang Stok", "Satuan", "Stok"])

    # Normalisasi SKU sama persis dengan di build_pivot (strip + upper)
    df["SKU"] = df["SKU"].astype(str).str.strip().str.upper()
    df = df[~df["SKU"].isin(["", "NAN", "NONE"])]

    df["Stok"] = clean_numeric_column(df["Stok"]).fillna(0)

    # Jika satu SKU muncul lebih dari sekali (mis. beda gudang), stok dijumlahkan
    return df.groupby("SKU", as_index=False)["Stok"].sum()

def clean_numeric_column(series: pd.Series) -> pd.Series:
    """Bersihkan kolom numerik yang mungkin terbawa sebagai string berformat mata uang."""
    if pd.api.types.is_numeric_dtype(series):
        return series
    cleaned = (
        series.astype(str)
        .str.replace(r"[^0-9\-,.]", "", regex=True)
        .str.replace(".", "", regex=False)   # asumsi titik = pemisah ribuan (format ID)
        .str.replace(",", ".", regex=False)  # asumsi koma = pemisah desimal (format ID)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    required_cols = [
        COL_TANGGAL, COL_SALES, COL_PELANGGAN, COL_BRAND, COL_SKU,
        COL_NAMA_BARANG, COL_KATEGORI, COL_QTY, COL_HARGA, COL_TOTAL, COL_LABA,
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Kolom berikut tidak ditemukan di sheet: {missing}")
        st.stop()

    df[COL_TANGGAL] = pd.to_datetime(df[COL_TANGGAL], dayfirst=True, errors="coerce")

    for col in [COL_QTY, COL_HARGA, COL_TOTAL, COL_LABA]:
        df[col] = clean_numeric_column(df[col])

    df[COL_SALES] = df[COL_SALES].astype(str).str.strip()
    df.loc[df[COL_SALES].isin(["", "nan", "None"]), COL_SALES] = ""
    df[COL_PELANGGAN] = df[COL_PELANGGAN].astype(str).str.strip()
    df[COL_BRAND] = df[COL_BRAND].astype(str).str.strip()   # <-- BARU

    df = df.dropna(subset=[COL_TANGGAL])
    return df


# ============================================================
# FILTER JENIS PENJUALAN
# ============================================================
def filter_sales_type(df: pd.DataFrame, jenis: str) -> pd.DataFrame:
    if jenis == "Sales Online":
        return df[df[COL_SALES].isin(NAMA_ONLINE)]

    if jenis == "Sales Offline":
        return df[
            (~df[COL_SALES].isin(NAMA_EXCLUDE_OFFLINE)) & (df[COL_SALES] != "")
        ]

    if jenis == "Marketplace":
        pattern = "|".join(KATA_MARKETPLACE)
        return df[df[COL_PELANGGAN].str.lower().str.contains(pattern, na=False)]

    # ALL SALES
    return df


# ============================================================
# PIVOT (acuan: SKU saja)
# ============================================================
def _pilih_nilai_dominan(series: pd.Series):
    """
    Ambil nilai paling sering muncul dalam satu SKU.
    Jika seri, ambil yang muncul terakhir (data paling baru, karena
    df sudah diurutkan berdasarkan tanggal sebelum groupby).
    """
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    if s.empty:
        return ""
    counts = s.value_counts()
    top = counts[counts == counts.max()].index
    # pilih yang paling akhir muncul di antara kandidat seri
    for val in reversed(s.tolist()):
        if val in top:
            return val
    return counts.index[0]


def build_pivot(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Normalisasi SKU agar variasi spasi/huruf kapital tidak membuat duplikat
    df[COL_SKU] = df[COL_SKU].astype(str).str.strip().str.upper()
    df = df[~df[COL_SKU].isin(["", "NAN", "NONE"])]

    # Urutkan by tanggal supaya "nilai terbaru" bisa dipakai saat tie-break
    df = df.sort_values(COL_TANGGAL)

    grouped = (
        df.groupby(COL_SKU, as_index=False)
        .agg(
            Brand=(COL_BRAND, _pilih_nilai_dominan),
            Nama_Barang=(COL_NAMA_BARANG, _pilih_nilai_dominan),
            Kategori_Barang=(COL_KATEGORI, _pilih_nilai_dominan),
            QTY=(COL_QTY, "sum"),
            Total_Harga=(COL_TOTAL, "sum"),
            Laba=(COL_LABA, "sum"),
        )
    )

    grouped["@Harga"] = np.where(
        grouped["QTY"] != 0, grouped["Total_Harga"] / grouped["QTY"], 0
    )
    grouped["Gross Profit/Item"] = np.where(
        grouped["QTY"] != 0, grouped["Laba"] / grouped["QTY"], 0
    )

    grouped = grouped.rename(
        columns={
            COL_SKU: "SKU",
            "Nama_Barang": "Nama Barang",
            "Kategori_Barang": "Kategori Barang",
            "Total_Harga": "Total Harga",
        }
    )

    grouped = grouped[
        [
            "Brand", "SKU", "Nama Barang", "Kategori Barang",
            "QTY", "@Harga", "Total Harga", "Laba", "Gross Profit/Item",
        ]
    ]
    return grouped


# ============================================================
# KLASIFIKASI INVENTORY MOVEMENT
# (porting dari Apps Script -> Python, moving disederhanakan
#  jadi 2 kelas lewat median split agar hasilnya persis 4 kelompok)
# ============================================================
def classify_quadrant(pivot_df: pd.DataFrame, margin_threshold: float = MARGIN_THRESHOLD):
    df = pivot_df.copy()

    df["Log_QTY"] = np.log(df["QTY"] + 1)
    median_log_qty = df["Log_QTY"].median()

    df["Kategori Moving"] = np.where(
        df["Log_QTY"] >= median_log_qty, "Fast Moving", "Slow Moving"
    )
    df["Kategori Margin"] = np.where(
        df["Gross Profit/Item"] >= margin_threshold, "High Margin", "Low Margin"
    )
    df["Kategori Kuadran"] = df["Kategori Margin"] + " - " + df["Kategori Moving"]

    # ---- Skor Kuadran (0-100), kombinasi posisi margin & moving ----
    min_gp, max_gp = df["Gross Profit/Item"].min(), df["Gross Profit/Item"].max()
    min_log, max_log = df["Log_QTY"].min(), df["Log_QTY"].max()

    df["Skor Margin"] = (
        0.0 if max_gp == min_gp
        else (df["Gross Profit/Item"] - min_gp) / (max_gp - min_gp) * 100
    )
    df["Skor Moving"] = (
        0.0 if max_log == min_log
        else (df["Log_QTY"] - min_log) / (max_log - min_log) * 100
    )
    df["Skor Kuadran"] = ((df["Skor Margin"] + df["Skor Moving"]) / 2).round(1)

    meta = {
        "median_log_qty": median_log_qty,
        "margin_threshold": margin_threshold,
        "jumlah": df["Kategori Kuadran"].value_counts().to_dict(),
    }
    return df, meta


# ============================================================
# UI
# ============================================================
st.title("📊 Klasifikasi Inventory Movement/SKU")
st.caption("High/Low Margin × Fast/Slow Moving — sumber data Google Sheets")

df_raw_all = preprocess(load_data())

if df_raw_all.empty:
    st.warning("Data kosong atau seluruh baris tanggal gagal dibaca.")
    st.stop()

min_date = df_raw_all[COL_TANGGAL].min().date()
max_date = df_raw_all[COL_TANGGAL].max().date()

col1, col2, col3, col4 = st.columns([1, 1, 1, 1.5])

with col1:
    start_date = st.date_input(
        "Tanggal Mulai",
        value=min_date,
        min_value=min_date,
        max_value=max_date,
    )

with col2:
    end_date = st.date_input(
        "Tanggal Akhir",
        value=max_date,
        min_value=min_date,
        max_value=max_date,
    )

with col3:
    jenis_penjualan = st.selectbox(
        "Jenis Penjualan",
        ["ALL SALES", "Sales Online", "Sales Offline", "Marketplace"],
    )

with col4:
    brand_options = sorted(
        b for b in df_raw_all[COL_BRAND].unique()
        if b not in ("", "nan", "None")
    )
    brand_terpilih = st.multiselect(
        "Brand",
        options=brand_options,
        default=[],
        placeholder="Semua brand",
    )

if start_date > end_date:
    st.warning("Tanggal Mulai tidak boleh setelah Tanggal Akhir.")
    st.stop()

mask_date = (df_raw_all[COL_TANGGAL].dt.date >= start_date) & (
    df_raw_all[COL_TANGGAL].dt.date <= end_date
)
df_date_filtered = df_raw_all[mask_date]

df_raw = filter_sales_type(df_date_filtered, jenis_penjualan)
if brand_terpilih:
    df_raw = df_raw[df_raw[COL_BRAND].isin(brand_terpilih)]
brand_label = ", ".join(brand_terpilih) if brand_terpilih else "Semua Brand"

st.divider()
st.subheader("Data Mentah (DATA_RAW)")
st.caption(
    f"{len(df_raw):,} baris — filter: {jenis_penjualan}, {brand_label}, "
    f"{start_date} s/d {end_date}"
)
st.dataframe(hide_cols(df_raw), use_container_width=True, height=300)

if df_raw.empty:
    st.warning("Tidak ada data pada rentang tanggal & filter jenis penjualan ini.")
    st.stop()

st.divider()
st.subheader("Pivot per Produk")
pivot_df = build_pivot(df_raw)
st.dataframe(
    hide_cols(pivot_df),
    use_container_width=True,
    height=300,
)

st.divider()
st.subheader("Hasil Klasifikasi Inventory Movement")
result_df, meta = classify_quadrant(pivot_df)
# Lookup stok berdasarkan SKU (left join: semua SKU hasil klasifikasi tetap tampil)
stok_df = load_stok()
result_df = result_df.merge(stok_df, on="SKU", how="left")
result_df["Stok"] = result_df["Stok"].fillna(0)  # SKU tidak ada di sheet Stok -> 0

info_cols = st.columns(4)
for i, kuadran in enumerate(QUADRANT_COLORS.keys()):
    info_cols[i].metric(kuadran, meta["jumlah"].get(kuadran, 0))

# Kolom lengkap (dipakai untuk perhitungan & sorting)
display_cols = [
    "Brand", "SKU", "Nama Barang", "Kategori Barang", "QTY", "Stok", "@Harga",
    "Total Harga", "Laba", "Gross Profit/Item", "Kategori Kuadran", "Skor Kuadran",
]

# Kolom untuk file Excel hasil download: tanpa @Harga, Total Harga, Laba, Gross Profit/Item
EXCEL_HIDDEN_COLS = ["@Harga", "Total Harga", "Laba", "Gross Profit/Item"]
excel_cols = [c for c in display_cols if c not in EXCEL_HIDDEN_COLS]

# Urutan tetap berdasarkan Skor Kuadran, tapi kolom skor tidak ditampilkan di UI
result_sorted = result_df[display_cols].sort_values("Skor Kuadran", ascending=False)
st.dataframe(
    hide_cols(result_sorted),
    use_container_width=True,
    height=350,
)

excel_buffer = io.BytesIO()
with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    result_df[excel_cols].to_excel(writer, index=False, sheet_name="Hasil Klasifikasi")
excel_buffer.seek(0)

st.download_button(
    "⬇️ Download Hasil Klasifikasi Inventory Movement",
    data=excel_buffer,
    file_name="Hasil Klasifikasi Inventory Movement.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.divider()
st.subheader("Chart Klasifikasi Inventory Movement")

fig = px.scatter(
    result_df,
    x="Log_QTY",
    y="Gross Profit/Item",
    color="Kategori Kuadran",
    size="Skor Kuadran",
    size_max=28,
    color_discrete_map=QUADRANT_COLORS,
    hover_data={
        "Brand": True,
        "SKU": True,
        "Nama Barang": True,
        "QTY": True,
        "Gross Profit/Item": ":,.0f",
        "Skor Kuadran": False,  # disembunyikan dari tooltip
        "Log_QTY": False,
    },
    labels={"Log_QTY": "Log(QTY + 1) — Moving Score", "Gross Profit/Item": "Gross Profit / Item (Rp)"},
)
fig.add_hline(y=meta["margin_threshold"], line_dash="dash", line_color="gray",
              annotation_text=f"Margin threshold ({meta['margin_threshold']:,.0f})")
fig.add_vline(x=meta["median_log_qty"], line_dash="dash", line_color="gray",
              annotation_text="Median Moving")
fig.update_layout(height=550, legend_title_text="Klasifikasi")

st.plotly_chart(fig, use_container_width=True)

# (Keterangan perhitungan sengaja dihapus dari UI)  
