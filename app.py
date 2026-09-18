"""
Aplikasi Klasifikasi Kuadran Sales
-----------------------------------
Sumber data : Google Sheets (via gspread + Service Account)
Output      : 4 Kuadran -> High/Low Margin x Fast/Slow Moving

Alur:
1. Ambil data mentah dari Google Sheets
2. Filter berdasarkan rentang tanggal
3. Filter berdasarkan jenis penjualan (Online / Offline / Marketplace / All Sales)
4. Tampilkan DATA_RAW
5. Pivot per Brand/SKU/Nama Barang/Kategori Barang
6. Klasifikasi kuadran (margin threshold fixed 30.000, moving pakai median log(QTY+1))
7. Tampilkan tabel hasil kuadran
8. Tampilkan chart kuadran (scatter, dengan skor)
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

NAMA_ONLINE = ["Arum", "Luthfiah Wardah"]
NAMA_EXCLUDE_OFFLINE = ["Arum", "Luthfiah Wardah", "Internal"]
KATA_MARKETPLACE = ["shopee", "tiktok", "lazada", "blibli", "tokopedia"]

QUADRANT_COLORS = {
    "High Margin - Fast Moving": "#2E7D32",
    "High Margin - Slow Moving": "#F9A825",
    "Low Margin - Fast Moving": "#1565C0",
    "Low Margin - Slow Moving": "#C62828",
}

st.set_page_config(page_title="Klasifikasi Kuadran Sales", layout="wide")


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
# PIVOT
# ============================================================
def build_pivot(df: pd.DataFrame) -> pd.DataFrame:
    grouped = (
        df.groupby([COL_BRAND, COL_SKU, COL_NAMA_BARANG, COL_KATEGORI], as_index=False)
        .agg(
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
            COL_BRAND: "Brand",
            COL_SKU: "SKU",
            COL_NAMA_BARANG: "Nama Barang",
            COL_KATEGORI: "Kategori Barang",
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
# KLASIFIKASI KUADRAN
# (porting dari Apps Script -> Python, moving disederhanakan
#  jadi 2 kelas lewat median split agar hasilnya persis 4 kuadran)
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
st.title("📊 Klasifikasi Kuadran Sales")
st.caption("High/Low Margin × Fast/Slow Moving — sumber data Google Sheets")

df_raw_all = preprocess(load_data())

if df_raw_all.empty:
    st.warning("Data kosong atau seluruh baris tanggal gagal dibaca.")
    st.stop()

min_date = df_raw_all[COL_TANGGAL].min().date()
max_date = df_raw_all[COL_TANGGAL].max().date()

col1, col2 = st.columns([2, 1])

with col1:
    date_range = st.date_input(
        "1️⃣ Rentang Tanggal",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )

with col2:
    jenis_penjualan = st.selectbox(
        "2️⃣ Jenis Penjualan",
        ["ALL SALES", "Sales Online", "Sales Offline", "Marketplace"],
    )

if len(date_range) != 2:
    st.info("Pilih tanggal awal dan akhir.")
    st.stop()

start_date, end_date = date_range
mask_date = (df_raw_all[COL_TANGGAL].dt.date >= start_date) & (
    df_raw_all[COL_TANGGAL].dt.date <= end_date
)
df_date_filtered = df_raw_all[mask_date]

df_raw = filter_sales_type(df_date_filtered, jenis_penjualan)

st.divider()
st.subheader("3️⃣ Data Mentah (DATA_RAW)")
st.caption(f"{len(df_raw):,} baris — filter: {jenis_penjualan}, {start_date} s/d {end_date}")
st.dataframe(df_raw, use_container_width=True, height=300)

if df_raw.empty:
    st.warning("Tidak ada data pada rentang tanggal & filter jenis penjualan ini.")
    st.stop()

st.divider()
st.subheader("4️⃣ Pivot per Produk")
pivot_df = build_pivot(df_raw)
st.dataframe(
    pivot_df.style.format(
        {"@Harga": "{:,.0f}", "Total Harga": "{:,.0f}", "Laba": "{:,.0f}", "Gross Profit/Item": "{:,.0f}"}
    ),
    use_container_width=True,
    height=300,
)

st.divider()
st.subheader("5️⃣ - 6️⃣ Hasil Klasifikasi Kuadran")
result_df, meta = classify_quadrant(pivot_df)

info_cols = st.columns(4)
for i, kuadran in enumerate(QUADRANT_COLORS.keys()):
    info_cols[i].metric(kuadran, meta["jumlah"].get(kuadran, 0))

display_cols = [
    "Brand", "SKU", "Nama Barang", "Kategori Barang", "QTY", "@Harga",
    "Total Harga", "Laba", "Gross Profit/Item", "Kategori Kuadran", "Skor Kuadran",
]
st.dataframe(
    result_df[display_cols]
    .sort_values("Skor Kuadran", ascending=False)
    .style.format(
        {"@Harga": "{:,.0f}", "Total Harga": "{:,.0f}", "Laba": "{:,.0f}", "Gross Profit/Item": "{:,.0f}"}
    ),
    use_container_width=True,
    height=350,
)

excel_buffer = io.BytesIO()
with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    result_df[display_cols].to_excel(writer, index=False, sheet_name="Hasil Kuadran")
excel_buffer.seek(0)

st.download_button(
    "⬇️ Download Hasil Kuadran (Excel)",
    data=excel_buffer,
    file_name="hasil_kuadran.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.divider()
st.subheader("7️⃣ Chart Kuadran")

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
        "Skor Kuadran": True,
        "Log_QTY": False,
    },
    labels={"Log_QTY": "Log(QTY + 1) — Moving Score", "Gross Profit/Item": "Gross Profit / Item (Rp)"},
)
fig.add_hline(y=meta["margin_threshold"], line_dash="dash", line_color="gray",
              annotation_text=f"Margin threshold ({meta['margin_threshold']:,.0f})")
fig.add_vline(x=meta["median_log_qty"], line_dash="dash", line_color="gray",
              annotation_text="Median Moving")
fig.update_layout(height=550, legend_title_text="Kuadran")

st.plotly_chart(fig, use_container_width=True)

with st.expander("ℹ️ Keterangan perhitungan"):
    st.markdown(
        f"""
        - **Margin**: `Gross Profit/Item = Laba (total) / QTY (total)` per produk hasil pivot.
          Threshold fixed **Rp {meta['margin_threshold']:,.0f}** → ≥ threshold = **High Margin**.
        - **Moving**: `Log(QTY + 1)`, dibandingkan terhadap **median** dari data yang sedang tampil
          (median = {meta['median_log_qty']:.3f}) → ≥ median = **Fast Moving**.
        - **Skor Kuadran** (0–100): rata-rata dari skor margin dan skor moving yang sudah
          dinormalisasi (min-max) terhadap data yang sedang tampil. Dipakai sebagai ukuran
          bubble pada chart — semakin besar & semakin ke kanan-atas, semakin unggul produk itu.
        """
    )