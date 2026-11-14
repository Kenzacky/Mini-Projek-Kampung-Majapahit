"""
app.py
======
Sistem Informasi & Interactive Data Analytics Dashboard
Desa Wisata Kampung Majapahit Bejijong
Mitra: LP3SDM Sanggar Bhagaskara | Proyek MSIB, D4 Manajemen Informatika UNESA

Cara menjalankan:
    pip install -r requirements.txt
    python generate_data.py          # membuat sample_desa_wisata.xlsx (sekali saja)
    streamlit run app.py

Akun demo:
    admin / bejijong123   (Admin Desa)
    unesa / msib2026      (Dosen / Penguji)

Arsitektur singkat (agar mudah dibaca dosen penguji):
    1. KONFIGURASI   : konstanta, tema warna, kredensial demo
    2. TEMA (CSS)    : tampilan Majapahit Heritage ala aplikasi SaaS
    3. DATA          : baca, bersihkan, fallback, penyimpanan input petugas
    4. AUTENTIKASI   : login berbasis st.session_state
    5. ANALITIK      : fungsi hitung bersama (per bulan, per saluran, rekomendasi, narasi)
    6. HALAMAN       : 7 menu (lihat dict MENU di bawah)
    7. MAIN          : router yang memilih halaman sesuai status login dan menu,
                       plus tombol pindah menu di bagian bawah setiap halaman

Tujuh menu sidebar:
    🏠 Beranda               informasi proyek, mitra, panduan
    📊 Analytics Dashboard   KPI + grafik Plotly + tabel & ekspor
    🔬 Eksplorasi Data (EDA) kualitas data, sebaran, pola waktu, segmentasi (alur CRISP-DM)
    💡 Rekomendasi Strategis Recommendation Engine berbasis aturan + simulasi alokasi anggaran
    📝 Laporan Ringkas       ringkasan eksekutif otomatis (data storytelling) + unduh Excel/TXT
    ➕ Input Data Pengunjung formulir petugas desa
    ⚙️ Pengaturan & Logout   info sesi dan keluar
"""

from __future__ import annotations

import hmac
import io
import time
from datetime import date, datetime
from html import escape
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================================
# 1. KONFIGURASI
# ============================================================================
st.set_page_config(
    page_title="Sistem Informasi & Analytics Desa Wisata Kampung Majapahit Bejijong",
    page_icon="🏛️",
    layout="wide",
)

# Palet "Majapahit Heritage"
TERAKOTA = "#8B3A2B"   # warna utama: bata Majapahit
EMAS = "#D4AF37"       # aksen: emas kerajaan
KAYU = "#6E473B"       # aksen: cokelat kayu hangat
KREM = "#FAF8F5"       # latar: krem terang
WARNA_GRAFIK = [TERAKOTA, EMAS, KAYU, "#B9805F", "#3E7C74", "#A89F91"]

BASE_DIR = Path(__file__).resolve().parent
FILE_SAMPEL = BASE_DIR / "sample_desa_wisata.xlsx"
FILE_INPUT = BASE_DIR / "data_input_petugas.csv"   # tempat menyimpan data dari form petugas

KOLOM_WAJIB = [
    "Tanggal", "Kategori_Pengunjung", "Saluran_Pemasaran", "Asal_Kota",
    "Jumlah_Pengunjung", "Harga_Tiket_Per_Orang", "Total_Pendapatan", "Tingkat_Kepuasan",
]
PILIHAN_KATEGORI = ["Wisatawan Lokal", "Wisatawan Mancanegara", "Rombongan Edukasi/Sekolah", "Peneliti/Akademisi"]
PILIHAN_SALURAN = ["Instagram", "TikTok", "Website Resmi", "Walk-in / Langsung", "Brosur / Event", "Travel Agent"]
PILIHAN_KOTA = ["Surabaya", "Mojokerto", "Malang", "Sidoarjo", "Jakarta", "Yogyakarta", "Mancanegara"]

# Kredensial DEMO untuk keperluan proyek. Pada sistem produksi, jangan menaruh kata sandi
# di dalam kode: simpan hash (bcrypt/argon2) di database atau st.secrets.
PENGGUNA = {
    "admin": {"password": "bejijong123", "nama": "Admin Desa", "peran": "Admin Desa"},
    "unesa": {"password": "msib2026", "nama": "Dosen / Penguji", "peran": "Dosen / Penguji UNESA"},
}
BATAS_GAGAL_LOGIN = 5      # setelah 5 kali salah, login dikunci sementara
DURASI_KUNCI_DETIK = 30

# Urutan dict ini = urutan menu di sidebar. Kuncinya dipakai oleh tombol pindah menu.
MENU = {
    "beranda": "🏠 Beranda",
    "dashboard": "📊 Analytics Dashboard",
    "eda": "🔬 Eksplorasi Data (EDA)",
    "rekomendasi": "💡 Rekomendasi Strategis",
    "laporan": "📝 Laporan Ringkas",
    "input": "➕ Input Data Pengunjung",
    "pengaturan": "⚙️ Pengaturan & Logout",
}

# Nama hari dan bulan Indonesia ditulis manual agar tidak bergantung pada locale komputer
NAMA_HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
NAMA_BULAN = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
NAMA_BULAN_PENUH = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
                    "September", "Oktober", "November", "Desember"]
SKALA_PANAS = [[0.0, "#FBF3E8"], [0.5, EMAS], [1.0, TERAKOTA]]   # skala warna heatmap

# ============================================================================
# 2. TEMA (CSS KUSTOM)
# ============================================================================
CSS_UMUM = f"""
<style>
  .stApp {{ background-color: {KREM}; }}
  #MainMenu, footer {{ visibility: hidden; }}
  [data-testid="stAppDeployButton"], .stDeployButton {{ display: none; }}
  .block-container {{ padding-top: 1.6rem; max-width: 1280px; }}
  h1, h2, h3 {{ color: {TERAKOTA}; letter-spacing: -0.01em; }}

  /* Banner header halaman */
  .banner {{
      background: linear-gradient(110deg, {TERAKOTA} 0%, #A04A37 55%, {KAYU} 100%);
      border-bottom: 5px solid {EMAS}; border-radius: 14px 14px 0 0;
      padding: 26px 32px 22px 32px; color: #FFF6EA; margin-bottom: 0;
  }}
  .banner small {{ color: {EMAS}; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; }}
  .banner h1 {{ color: #FFFFFF; font-size: 2rem; margin: 6px 0 4px 0; }}
  .banner p {{ color: #F1DED0; margin: 0; max-width: 70ch; }}
  .bata {{ height: 9px; margin-bottom: 22px; border-radius: 0 0 14px 14px;
      background-color: {KAYU};
      background-image: linear-gradient(90deg, transparent calc(100% - 2px), rgba(0,0,0,.35) 0),
                        linear-gradient(0deg, rgba(0,0,0,.35) 2px, transparent 0);
      background-size: 26px 100%, 100% 9px; }}

  /* Kartu KPI (st.metric) */
  [data-testid="stMetric"] {{
      background: #FFFFFF; border: 1px solid #E8DFD6; border-left: 6px solid {TERAKOTA};
      border-radius: 12px; padding: 14px 18px; min-height: 118px;
      box-shadow: 0 6px 16px -12px rgba(110,71,59,.5);
  }}
  /* Keterangan di bawah KPI bukan perubahan naik/turun, jadi panahnya disembunyikan */
  [data-testid="stMetricDeltaIcon-Up"], [data-testid="stMetricDeltaIcon-Down"] {{ display: none; }}
  [data-testid="stMetricLabel"] p {{ color: {KAYU}; font-weight: 600; }}
  [data-testid="stMetricValue"] {{ color: #2B1F1A; font-size: 1.65rem; }}
  [data-testid="stMetricValue"] > div {{ white-space: normal; overflow: visible; text-overflow: clip; line-height: 1.2; }}

  /* Kartu informasi umum */
  .kartu {{ background: #FFFFFF; border: 1px solid #E8DFD6; border-radius: 12px; padding: 18px 20px; height: 100%; }}
  .kartu h4 {{ color: {TERAKOTA}; margin: 0 0 6px 0; }}
  .kartu p, .kartu li {{ color: #4A3A33; font-size: .95rem; }}
  .langkah {{ display: flex; gap: 12px; align-items: flex-start; }}
  .langkah b.no {{ background: {EMAS}; color: #3A2A10; border-radius: 50%; min-width: 30px; height: 30px;
      display: grid; place-items: center; }}

  /* Sidebar */
  [data-testid="stSidebar"] {{ border-right: 1px solid #E8DFD6; }}
  .akun {{ background: #FFFFFF; border: 1px solid #E8DFD6; border-top: 4px solid {EMAS};
      border-radius: 12px; padding: 12px 14px; margin-bottom: 10px; }}
  .akun b {{ color: {TERAKOTA}; }} .akun span {{ color: #6B5A52; font-size: .85rem; }}

  /* Tombol */
  .stButton > button, [data-testid="stFormSubmitButton"] > button, .stDownloadButton > button {{
      border-radius: 10px; font-weight: 600; }}
  button[kind="primary"]:hover, button[kind="primaryFormSubmit"]:hover {{ background-color: {KAYU}; border-color: {KAYU}; }}
  button[data-baseweb="tab"] {{ font-weight: 600; }}

  /* Kartu rekomendasi */
  .rekom {{ background: #FFFFFF; border: 1px solid #E8DFD6; border-left: 6px solid {KAYU}; border-radius: 12px;
      padding: 16px 18px; margin-bottom: 14px; }}
  .rekom.tinggi {{ border-left-color: {TERAKOTA}; }}
  .rekom.sedang {{ border-left-color: {EMAS}; }}
  .rekom.info {{ border-left-color: #3E7C74; }}
  .rekom h4 {{ margin: 0 0 8px 0; color: #2B1F1A; font-size: 1.02rem; }}
  .rekom p {{ margin: 8px 0 0 0; color: #4A3A33; font-size: .93rem; line-height: 1.55; }}
  .rekom p.aksi {{ background: #FAF3EA; border-radius: 8px; padding: 8px 10px; }}
  .chip {{ display: inline-block; font-size: .72rem; font-weight: 700; padding: 2px 10px; border-radius: 99px; letter-spacing: .03em; }}
  .chip.tinggi {{ background: {TERAKOTA}; color: #FFFFFF; }}
  .chip.sedang {{ background: {EMAS}; color: #3A2A10; }}
  .chip.info {{ background: #3E7C74; color: #FFFFFF; }}

  /* Alur CRISP-DM */
  .crisp {{ display: flex; flex-wrap: wrap; gap: 8px; margin: 0 0 14px 0; }}
  .crisp span {{ border: 1px solid #E8DFD6; background: #FFFFFF; color: #6B5A52; border-radius: 99px;
      padding: 4px 12px; font-size: .8rem; }}
  .crisp span.aktif {{ background: {TERAKOTA}; border-color: {TERAKOTA}; color: #FFFFFF; font-weight: 600; }}

  /* Narasi laporan */
  .narasi {{ background: #FFFFFF; border: 1px solid #E8DFD6; border-left: 6px solid {EMAS}; border-radius: 12px;
      padding: 14px 20px; margin-bottom: 12px; }}
  .narasi h4 {{ color: {TERAKOTA}; margin: 0 0 6px 0; }}
  .narasi p {{ margin: 0; color: #3A2D27; line-height: 1.65; }}

  /* Garis pemisah di atas tombol pindah menu */
  .navbawah {{ margin-top: 30px; border-top: 2px solid #E8DFD6; }}
</style>
"""

# CSS tambahan khusus halaman login: sembunyikan sidebar dan jadikan form sebagai kartu login
CSS_LOGIN = f"""
<style>
  [data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], [data-testid="collapsedControl"],
  [data-testid="stHeader"] {{ display: none; }}
  .stApp {{ background: radial-gradient(1200px 500px at 50% -10%, #F1DFD2 0%, {KREM} 60%); }}
  .block-container {{ padding-top: 7vh; }}
  [data-testid="stForm"] {{
      background: #FFFFFF; border: 1px solid #E8DFD6; border-top: 6px solid {TERAKOTA};
      border-radius: 16px; padding: 28px 30px 24px 30px; box-shadow: 0 24px 48px -28px rgba(110,71,59,.55);
  }}
  .login-kepala {{ text-align: center; margin-bottom: 6px; }}
  .login-kepala .judul {{ margin: 8px 0 2px 0; font-size: 1.5rem; font-weight: 700; color: {TERAKOTA}; }}
  .login-kepala p {{ color: #6B5A52; margin: 0; font-size: .92rem; }}
  .login-kaki {{ text-align: center; color: #8B7A70; font-size: .8rem; margin-top: 14px; }}
</style>
"""

def html_rapat(teks: str) -> str:
    """
    Meratakan blok HTML menjadi satu baris tanpa indentasi dan tanpa baris kosong.
    Penting: Markdown menganggap teks berindentasi 4 spasi setelah baris kosong sebagai
    BLOK KODE, sehingga tag HTML tampil mentah di layar. Dengan meratakan HTML,
    hal itu tidak akan terjadi.
    """
    return " ".join(baris.strip() for baris in teks.strip().splitlines())


# Logo candi bentar (gerbang terbelah khas Majapahit) dengan matahari emas di tengahnya
LOGO_SVG = html_rapat(f"""
<svg viewBox="0 0 28 24" width="84" height="72" xmlns="http://www.w3.org/2000/svg" aria-label="Candi Bentar">
  <circle cx="14" cy="11" r="3.4" fill="{EMAS}"/>
  <path d="M1 23V18H4V14H7V10H9.5V3H12V23Z M27 23V18H24V14H21V10H18.5V3H16V23Z" fill="{TERAKOTA}"/>
</svg>
""")

# ============================================================================
# 3. FUNGSI DATA
# ============================================================================
_VERSI_BARU = tuple(int(x) for x in st.__version__.split(".")[:2]) >= (1, 50)


def tampil_grafik(fig) -> None:
    """Menampilkan grafik Plotly selebar kolom (kompatibel versi Streamlit lama dan baru)."""
    fig.update_layout(
        template="plotly_white", font=dict(family="Segoe UI, Arial, sans-serif"),
        margin=dict(l=10, r=10, t=50, b=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    )
    if _VERSI_BARU:
        st.plotly_chart(fig, width="stretch")
    else:
        st.plotly_chart(fig, use_container_width=True)


def tampil_tabel(df: pd.DataFrame, **kwargs) -> None:
    """Menampilkan tabel selebar kolom (kompatibel versi Streamlit lama dan baru)."""
    if _VERSI_BARU:
        st.dataframe(df, width="stretch", **kwargs)
    else:
        st.dataframe(df, use_container_width=True, **kwargs)


def lebar_penuh() -> dict:
    """Argumen 'tombol selebar kolom' yang sesuai versi Streamlit terpasang (dipakai sebagai **kwargs)."""
    return {"width": "stretch"} if _VERSI_BARU else {"use_container_width": True}


def rupiah(angka: float) -> str:
    """Format 'Rp 25.000.000' dengan pemisah ribuan titik."""
    return "Rp " + f"{angka:,.0f}".replace(",", ".")


def angka_id(angka: float) -> str:
    """Format 1250 menjadi '1.250'."""
    return f"{angka:,.0f}".replace(",", ".")


def rapikan_angka(tabel: pd.DataFrame) -> pd.DataFrame:
    """
    Salinan tabel khusus untuk TAMPILAN: bilangan bulat diberi pemisah ribuan titik (3.410),
    desimal dibulatkan (kepuasan 2 angka, lainnya 1 angka). Data asli tetap numerik untuk unduhan.
    """
    tampil = tabel.copy()
    for kolom in tampil.columns:
        if pd.api.types.is_integer_dtype(tampil[kolom]):
            tampil[kolom] = tampil[kolom].map(angka_id)
        elif pd.api.types.is_float_dtype(tampil[kolom]):
            desimal = 2 if "Kepuasan" in str(kolom) else 1
            tampil[kolom] = tampil[kolom].map(lambda x, d=desimal: "-" if pd.isna(x) else f"{x:.{d}f}")
    return tampil


def persen(nilai: float) -> str:
    """Format 0.4312 menjadi '43%'."""
    return f"{nilai:.0%}"


def tgl_id(ts) -> str:
    """Format tanggal gaya Indonesia: '12 Nov 2026'."""
    return f"{ts.day} {NAMA_BULAN[ts.month - 1]} {ts.year}"


def label_bulan(ts, penuh: bool = False) -> str:
    """Format bulan: 'Nov 2026' atau, bila penuh=True, 'November 2026'."""
    nama = (NAMA_BULAN_PENUH if penuh else NAMA_BULAN)[ts.month - 1]
    return f"{nama} {ts.year}"


def data_cadangan(n: int = 300, seed: int = 7) -> pd.DataFrame:
    """DataFrame sintetis otomatis bila tidak ada unggahan dan file sampel tidak ditemukan."""
    rng = np.random.default_rng(seed)
    tanggal = pd.Timestamp.today().normalize() - pd.to_timedelta(rng.integers(0, 365, n), unit="D")
    kat = rng.choice(PILIHAN_KATEGORI, n, p=[0.5, 0.12, 0.28, 0.10])
    df = pd.DataFrame({
        "Tanggal": tanggal,
        "Kategori_Pengunjung": kat,
        "Saluran_Pemasaran": rng.choice(PILIHAN_SALURAN, n),
        "Asal_Kota": np.where(kat == "Wisatawan Mancanegara", "Mancanegara", rng.choice(PILIHAN_KOTA[:-1], n)),
        "Jumlah_Pengunjung": rng.integers(2, 51, n),
        "Harga_Tiket_Per_Orang": rng.integers(10, 26, n) * 1000,
        "Tingkat_Kepuasan": np.round(np.clip(rng.normal(4.3, 0.6, n), 1.0, 5.0), 1),
    })
    df["Total_Pendapatan"] = df["Jumlah_Pengunjung"] * df["Harga_Tiket_Per_Orang"]
    return df[KOLOM_WAJIB].sort_values("Tanggal").reset_index(drop=True)


def bersihkan(df: pd.DataFrame) -> pd.DataFrame:
    """Merapikan nama kolom dan tipe data. Melempar ValueError bila kolom wajib tidak lengkap."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if "Total_Pendapatan" not in df.columns and {"Jumlah_Pengunjung", "Harga_Tiket_Per_Orang"} <= set(df.columns):
        df["Total_Pendapatan"] = (pd.to_numeric(df["Jumlah_Pengunjung"], errors="coerce")
                                  * pd.to_numeric(df["Harga_Tiket_Per_Orang"], errors="coerce"))
    hilang = [c for c in KOLOM_WAJIB if c not in df.columns]
    if hilang:
        raise ValueError("Kolom wajib tidak ditemukan: " + ", ".join(hilang))
    df = df[KOLOM_WAJIB].copy()
    df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
    for kolom in ["Jumlah_Pengunjung", "Harga_Tiket_Per_Orang", "Total_Pendapatan", "Tingkat_Kepuasan"]:
        df[kolom] = pd.to_numeric(df[kolom], errors="coerce")
    for kolom in ["Kategori_Pengunjung", "Saluran_Pemasaran", "Asal_Kota"]:
        df[kolom] = df[kolom].astype(str).str.strip()
    df = df.dropna(subset=["Tanggal", "Jumlah_Pengunjung"]).reset_index(drop=True)
    if df.empty:
        raise ValueError("File tidak memiliki baris data yang valid.")
    return df


def baca_unggahan(berkas) -> pd.DataFrame:
    """Membaca file unggahan (.xlsx, .xls, atau .csv) menjadi DataFrame mentah."""
    if berkas.name.lower().endswith(".csv"):
        return pd.read_csv(berkas, sep=None, engine="python")   # sep=None: tebak koma atau titik koma
    return pd.read_excel(berkas)


@st.cache_data(show_spinner=False)
def muat_sampel(mtime: float) -> pd.DataFrame:
    """Membaca file sampel. Argumen `mtime` membuat cache diperbarui saat file berubah."""
    return bersihkan(pd.read_excel(FILE_SAMPEL))


def baca_input_petugas() -> pd.DataFrame:
    """
    Membaca semua data yang diinput petugas lewat form.
    Sumber: file CSV, ditambah cadangan di session_state bila file tidak bisa ditulis
    (misalnya server hosting yang read-only).
    """
    bagian = []
    if FILE_INPUT.exists():
        try:
            bagian.append(pd.read_csv(FILE_INPUT))
        except (OSError, pd.errors.ParserError, pd.errors.EmptyDataError):
            pass
    cache = st.session_state.get("input_cache", [])
    if cache:
        bagian.append(pd.DataFrame(cache))
    if not bagian:
        return pd.DataFrame(columns=KOLOM_WAJIB)
    try:
        return bersihkan(pd.concat(bagian, ignore_index=True))
    except ValueError:
        return pd.DataFrame(columns=KOLOM_WAJIB)


def simpan_input(baris: dict) -> tuple[bool, str]:
    """
    Menambahkan satu baris data ke file CSV. Bila gagal menulis file,
    baris disimpan di session_state agar tetap tampil pada sesi ini.
    Mengembalikan (berhasil_ke_file, pesan).
    """
    try:
        pd.DataFrame([baris]).to_csv(FILE_INPUT, mode="a", header=not FILE_INPUT.exists(), index=False, encoding="utf-8")
        return True, f"Tersimpan di {FILE_INPUT.name}"
    except OSError:
        st.session_state.setdefault("input_cache", []).append(baris)
        return False, "Server tidak mengizinkan menulis file. Data hanya tersimpan selama sesi ini."


def data_dasar() -> tuple[pd.DataFrame, str, str | None]:
    """
    Menentukan data dasar dengan urutan prioritas:
    1) hasil unggahan pengguna (tersimpan di session_state),
    2) file sampel sample_desa_wisata.xlsx,
    3) data cadangan sintetis.
    Mengembalikan (DataFrame, keterangan sumber, peringatan atau None).
    """
    unggahan = st.session_state.get("unggahan")
    if unggahan is not None:
        return unggahan["df"], f"Unggahan: {unggahan['nama']}", None
    peringatan = st.session_state.pop("peringatan_unggah", None)
    if FILE_SAMPEL.exists():
        try:
            return muat_sampel(FILE_SAMPEL.stat().st_mtime), "Data contoh: sample_desa_wisata.xlsx", peringatan
        except Exception as galat:  # noqa: BLE001 - kegagalan dibaca pengguna, bukan crash
            peringatan = f"File contoh rusak ({galat})."
    return data_cadangan(), "Data cadangan otomatis (simulasi)", peringatan


def gabung_input_petugas(data: pd.DataFrame, sertakan: bool) -> tuple[pd.DataFrame, int]:
    """Menambahkan data input petugas ke `data`. Mengembalikan (DataFrame, jumlah baris tambahan)."""
    if not sertakan:
        return data, 0
    tambahan = baca_input_petugas()
    if tambahan.empty:
        return data, 0
    return pd.concat([data, tambahan], ignore_index=True), len(tambahan)


def data_analisis() -> tuple[pd.DataFrame, str]:
    """
    Data untuk menu EDA, Rekomendasi, dan Laporan: data dasar + input petugas.
    Mengikuti pilihan kotak centang 'Sertakan data input petugas' di Analytics Dashboard
    (disimpan di session_state["pref_petugas"] agar tidak hilang saat pindah menu).
    """
    data, sumber, peringatan = data_dasar()
    if peringatan:
        st.warning(peringatan)
    gabungan, tambahan = gabung_input_petugas(data, st.session_state.get("pref_petugas", True))
    if tambahan:
        sumber += f" + {angka_id(tambahan)} entri petugas"
    return gabungan, sumber


PILIHAN_PERIODE = {"Seluruh data": None, "6 bulan terakhir": 6, "3 bulan terakhir": 3}


def pilih_periode(df: pd.DataFrame, kunci: str) -> tuple[pd.DataFrame, str]:
    """
    Pemilih periode analisis di dalam halaman (bukan di sidebar).
    Jendela waktu dihitung mundur dari tanggal data terakhir, bukan dari hari ini,
    sehingga tetap masuk akal untuk data unggahan yang lama.
    """
    pilihan = st.radio("Periode analisis", list(PILIHAN_PERIODE), horizontal=True, key=f"periode_{kunci}")
    n_bulan = PILIHAN_PERIODE[pilihan]
    if n_bulan is None:
        return df, pilihan
    batas = df["Tanggal"].max() - pd.DateOffset(months=n_bulan)
    return df[df["Tanggal"] > batas].copy(), pilihan


# ============================================================================
# 4. AUTENTIKASI
# ============================================================================
def sudah_login() -> bool:
    return bool(st.session_state.get("logged_in", False))


def cek_kredensial(username: str, password: str) -> dict | None:
    """Memeriksa username dan password. compare_digest mencegah serangan timing."""
    akun = PENGGUNA.get(username.strip().lower())
    if akun is None:
        return None
    benar = hmac.compare_digest(password.encode("utf-8"), akun["password"].encode("utf-8"))
    return akun if benar else None


def logout() -> None:
    """Menghapus seluruh state sesi lalu kembali ke halaman login."""
    st.session_state.clear()
    st.session_state["logged_in"] = False
    st.rerun()


def halaman_login() -> None:
    """Halaman login terpisah: kartu di tengah layar dengan logo candi."""
    st.markdown(CSS_LOGIN, unsafe_allow_html=True)
    _, tengah, _ = st.columns([1, 1.15, 1])
    with tengah:
        with st.form("form_login", clear_on_submit=False):
            st.markdown(
                html_rapat(f"""<div class="login-kepala">{LOGO_SVG}
                    <div class="judul">Kampung Majapahit Bejijong</div>
                    <p>Sistem Informasi &amp; Analytics Desa Wisata</p></div>"""),
                unsafe_allow_html=True,
            )
            username = st.text_input("Username", placeholder="Masukkan username")
            password = st.text_input("Password", type="password", placeholder="Masukkan password")
            masuk = st.form_submit_button("Masuk", type="primary", **lebar_penuh())

        # Proteksi sederhana terhadap tebak-tebakan password
        sisa_kunci = st.session_state.get("kunci_sampai", 0) - time.time()
        if sisa_kunci > 0:
            st.error(f"Terlalu banyak percobaan gagal. Coba lagi dalam {int(sisa_kunci) + 1} detik.")
        elif masuk:
            akun = cek_kredensial(username, password)
            if akun:
                st.session_state.update(
                    logged_in=True, username=username.strip().lower(), nama=akun["nama"], peran=akun["peran"],
                    login_waktu=datetime.now(), gagal=0, menu=MENU["beranda"],
                )
                st.rerun()   # langsung alihkan ke aplikasi utama
            else:
                gagal = st.session_state.get("gagal", 0) + 1
                st.session_state["gagal"] = gagal
                if gagal >= BATAS_GAGAL_LOGIN:
                    st.session_state["kunci_sampai"] = time.time() + DURASI_KUNCI_DETIK
                    st.session_state["gagal"] = 0
                st.error("Username atau password salah. Periksa kembali lalu coba lagi.")
        st.markdown('<div class="login-kaki">Proyek MSIB UNESA · LP3SDM Sanggar Bhagaskara</div>', unsafe_allow_html=True)


# ============================================================================
# 5. ANALITIK (fungsi hitung bersama, tidak menampilkan apa pun ke layar)
# ============================================================================
# Bagian ini murni menghitung dari DataFrame, sehingga mudah diuji dan dipakai ulang
# oleh menu EDA, Rekomendasi Strategis, dan Laporan Ringkas.

def ringkas_per(df: pd.DataFrame, kolom: str) -> pd.DataFrame:
    """
    Ringkasan per kelompok (kategori, saluran, atau kota): jumlah kunjungan, pengunjung,
    pendapatan, kepuasan rata-rata, ukuran rombongan, pendapatan per pengunjung, dan porsi
    terhadap total. Hasil diurutkan dari pengunjung terbanyak.
    """
    g = (df.groupby(kolom)
           .agg(Kunjungan=("Jumlah_Pengunjung", "size"), Pengunjung=("Jumlah_Pengunjung", "sum"),
                Pendapatan=("Total_Pendapatan", "sum"), Kepuasan=("Tingkat_Kepuasan", "mean"))
           .reset_index())
    g["Rata_per_Kunjungan"] = g["Pengunjung"] / g["Kunjungan"].clip(lower=1)
    g["Pendapatan_per_Pengunjung"] = g["Pendapatan"] / g["Pengunjung"].clip(lower=1)
    g["Porsi_Pengunjung"] = g["Pengunjung"] / max(g["Pengunjung"].sum(), 1)
    g["Porsi_Pendapatan"] = g["Pendapatan"] / max(g["Pendapatan"].sum(), 1)
    return g.sort_values("Pengunjung", ascending=False).reset_index(drop=True)


def tabel_ringkas(df: pd.DataFrame, kolom: str, judul: str) -> pd.DataFrame:
    """Versi ramah baca dari ringkas_per(): nama kolom Indonesia, porsi dalam persen. Dipakai di layar dan Excel."""
    g = ringkas_per(df, kolom)
    return pd.DataFrame({
        judul: g[kolom],
        "Jumlah Kunjungan": g["Kunjungan"],
        "Total Pengunjung": g["Pengunjung"].astype(int),
        "Total Pendapatan (Rp)": g["Pendapatan"].round(0).astype(int),
        "Rata-rata Kepuasan": g["Kepuasan"].round(2),
        "Rata-rata per Kunjungan": g["Rata_per_Kunjungan"].round(1),
        "Pendapatan per Pengunjung (Rp)": g["Pendapatan_per_Pengunjung"].round(0).astype(int),
        "Porsi Pengunjung (%)": (g["Porsi_Pengunjung"] * 100).round(1),
        "Porsi Pendapatan (%)": (g["Porsi_Pendapatan"] * 100).round(1),
    })


def ringkas_bulanan(df: pd.DataFrame) -> pd.DataFrame:
    """
    Total per bulan (bulan tanpa data diisi 0). Kolom `Lengkap` bernilai True bila data menutup
    seluruh bulan tersebut. Bulan di awal atau akhir rentang data biasanya parsial, sehingga
    tidak adil dibandingkan dengan bulan penuh saat mencari bulan puncak.
    """
    bulan = df["Tanggal"].dt.to_period("M").dt.start_time
    r = df.groupby(bulan).agg(Kunjungan=("Jumlah_Pengunjung", "size"), Pengunjung=("Jumlah_Pengunjung", "sum"),
                              Pendapatan=("Total_Pendapatan", "sum"))
    semua_bulan = pd.date_range(r.index.min(), r.index.max(), freq="MS")
    r = r.reindex(semua_bulan, fill_value=0).rename_axis("Bulan").reset_index()
    awal, akhir = df["Tanggal"].min().normalize(), df["Tanggal"].max().normalize()
    r["Lengkap"] = (r["Bulan"] >= awal) & ((r["Bulan"] + pd.offsets.MonthEnd(0)) <= akhir)
    return r


def pilih_bulan_pembanding(bulanan: pd.DataFrame) -> pd.DataFrame:
    """Bulan yang layak dibandingkan: bulan lengkap bila ada minimal dua, selain itu semua bulan."""
    return bulanan[bulanan["Lengkap"]] if int(bulanan["Lengkap"].sum()) >= 2 else bulanan


def rata_per_hari_minggu(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rata-rata pengunjung per hari kalender untuk tiap hari dalam seminggu. Total dibagi jumlah
    hari tersebut dalam rentang data (bukan jumlah baris), jadi hari yang kebetulan muncul
    lebih sering dalam rentang tidak terlihat lebih ramai.
    """
    hari = pd.date_range(df["Tanggal"].min().normalize(), df["Tanggal"].max().normalize(), freq="D")
    jumlah_hari = pd.Series(hari.dayofweek).value_counts().reindex(range(7), fill_value=0)
    total = df.groupby(df["Tanggal"].dt.dayofweek)["Jumlah_Pengunjung"].sum().reindex(range(7), fill_value=0)
    rata = (total / jumlah_hari.replace(0, np.nan)).fillna(0)
    return pd.DataFrame({"Hari": NAMA_HARI, "Total": total.to_numpy(),
                         "Hari_Kalender": jumlah_hari.to_numpy(), "Rata_per_Hari": rata.to_numpy()})


def periksa_kualitas(df: pd.DataFrame) -> pd.DataFrame:
    """Memeriksa data terhadap aturan bisnis. Mengembalikan tabel: pemeriksaan, aturan, jumlah baris bermasalah, status."""
    hari_ini = pd.Timestamp.today().normalize()
    puas = df["Tingkat_Kepuasan"]
    jumlah, harga = df["Jumlah_Pengunjung"], df["Harga_Tiket_Per_Orang"]
    mancanegara_kat = df["Kategori_Pengunjung"] == "Wisatawan Mancanegara"
    mancanegara_kota = df["Asal_Kota"] == "Mancanegara"
    di_luar_daftar = (~df["Kategori_Pengunjung"].isin(PILIHAN_KATEGORI)
                      | ~df["Saluran_Pemasaran"].isin(PILIHAN_SALURAN)
                      | ~df["Asal_Kota"].isin(PILIHAN_KOTA))
    aturan = [
        ("Jumlah pengunjung", "2 sampai 50 orang per kunjungan", ((jumlah < 2) | (jumlah > 50)).sum()),
        ("Harga tiket", "Rp 10.000 sampai Rp 25.000 per orang", ((harga < 10_000) | (harga > 25_000)).sum()),
        ("Total pendapatan", "sama dengan jumlah pengunjung x harga tiket",
         (~np.isclose(df["Total_Pendapatan"], jumlah * harga)).sum()),
        ("Tingkat kepuasan", "1,0 sampai 5,0 (bila diisi)", ((puas < 1) | (puas > 5)).sum()),
        ("Konsistensi asal", "kategori Mancanegara hanya dari kota Mancanegara", (mancanegara_kat != mancanegara_kota).sum()),
        ("Tanggal", "tidak melewati hari ini", (df["Tanggal"] > hari_ini).sum()),
        ("Nilai kategori", "sesuai daftar baku kategori, saluran, dan kota", di_luar_daftar.sum()),
    ]
    hasil = pd.DataFrame(aturan, columns=["Pemeriksaan", "Aturan", "Baris bermasalah"])
    hasil["Baris bermasalah"] = hasil["Baris bermasalah"].astype(int)
    hasil["Status"] = np.where(hasil["Baris bermasalah"] == 0, "✅ Lolos", "⚠️ Perlu dicek")
    return hasil


def hitung_alokasi(df: pd.DataFrame, anggaran: float, bobot_pendapatan: float = 0.6) -> pd.DataFrame:
    """
    Usulan pembagian anggaran promosi ke tiap saluran pemasaran. Aturannya sengaja sederhana
    agar bisa dijelaskan kepada pengelola desa:

        skor = bobot x porsi pendapatan + (1 - bobot) x porsi pengunjung
        skor disesuaikan sedikit dengan kepuasan saluran (faktor 0,9 sampai 1,1)
        porsi anggaran = skor saluran / jumlah skor semua saluran

    Status: 'Prioritas' bila porsinya >= 1,25 x porsi rata-rata, 'Pertahankan' bila >= 0,75 x,
    selain itu 'Uji coba kecil'.
    """
    rs = ringkas_per(df, "Saluran_Pemasaran")
    n = len(rs)
    skor = bobot_pendapatan * rs["Porsi_Pendapatan"] + (1 - bobot_pendapatan) * rs["Porsi_Pengunjung"]
    rata_puas = rs["Kepuasan"].mean()
    if pd.notna(rata_puas) and rata_puas > 0:
        faktor = (rs["Kepuasan"] / rata_puas).clip(0.9, 1.1).fillna(1.0)
        skor = skor * faktor
    porsi = skor / skor.sum() if skor.sum() > 0 else pd.Series(1.0 / n, index=rs.index)

    alokasi = np.floor(porsi.to_numpy() * anggaran)
    alokasi[int(np.argmax(porsi.to_numpy()))] += anggaran - alokasi.sum()   # sisa pembulatan ke saluran teratas

    rata_porsi = 1.0 / n
    status = np.where(porsi >= 1.25 * rata_porsi, "Prioritas",
                      np.where(porsi >= 0.75 * rata_porsi, "Pertahankan", "Uji coba kecil"))
    hasil = pd.DataFrame({
        "Saluran_Pemasaran": rs["Saluran_Pemasaran"], "Pengunjung": rs["Pengunjung"].astype(int),
        "Pendapatan": rs["Pendapatan"].round(0).astype(int), "Kepuasan": rs["Kepuasan"].round(2),
        "Porsi_Usulan": porsi.to_numpy(), "Alokasi_Rp": alokasi.astype(int), "Status": status,
    })
    return hasil.sort_values("Porsi_Usulan", ascending=False).reset_index(drop=True)


URUTAN_PRIORITAS = {"Tinggi": 0, "Sedang": 1, "Info": 2}


def buat_rekomendasi(df: pd.DataFrame) -> list[dict]:
    """
    Recommendation Engine berbasis aturan. Setiap aturan membaca satu pola dari data
    (musim, hari, saluran, segmen, kepuasan, kota) lalu menghasilkan temuan dan tindakan.
    Hasilnya berupa daftar dict dengan kunci: Ikon, Judul, Prioritas, Temuan, Aksi.
    """
    hasil: list[dict] = []

    def tambah(ikon: str, judul: str, prioritas: str, temuan: str, aksi: str) -> None:
        hasil.append({"Ikon": ikon, "Judul": judul, "Prioritas": prioritas, "Temuan": temuan, "Aksi": aksi})

    # --- Aturan 1: periode puncak kunjungan ---
    kandidat = pilih_bulan_pembanding(ringkas_bulanan(df))
    if len(kandidat) >= 2:
        urut = kandidat.sort_values("Pengunjung", ascending=False)
        puncak = urut.iloc[0]
        rata_bulan = kandidat["Pengunjung"].mean()
        lebih = puncak["Pengunjung"] / rata_bulan - 1 if rata_bulan > 0 else 0.0
        tiga_teramai = ", ".join(NAMA_BULAN_PENUH[b.month - 1] for b in urut["Bulan"].head(3))
        nama_puncak = NAMA_BULAN_PENUH[puncak["Bulan"].month - 1]
        tambah("📅", "Siapkan kampanye sebelum musim puncak", "Tinggi",
               f"Bulan paling ramai: {label_bulan(puncak['Bulan'], True)} dengan {angka_id(puncak['Pengunjung'])} "
               f"pengunjung, {persen(lebih)} di atas rata-rata bulanan. Tiga bulan teramai: {tiga_teramai}.",
               f"Mulai promosi 4 sampai 6 minggu sebelum {nama_puncak}, dan pastikan pemandu, homestay, serta stok "
               f"suvenir kuningan siap menampung lonjakan kunjungan.")
    else:
        tambah("📅", "Pola musim belum bisa dibaca", "Info",
               "Data baru mencakup kurang dari dua bulan, sehingga bulan puncak belum dapat ditentukan.",
               "Lanjutkan pencatatan harian minimal 6 bulan agar periode puncak dapat diprediksi.")

    # --- Aturan 2: hari tersibuk dan tersepi (efisiensi kapasitas) ---
    if (df["Tanggal"].max() - df["Tanggal"].min()).days >= 13:
        hr = rata_per_hari_minggu(df)
        ramai, sepi = hr.loc[hr["Rata_per_Hari"].idxmax()], hr.loc[hr["Rata_per_Hari"].idxmin()]
        if ramai["Rata_per_Hari"] > sepi["Rata_per_Hari"]:
            tambah("🗓️", "Ratakan beban kunjungan dengan paket hari sepi", "Sedang",
                   f"{ramai['Hari']} paling ramai (rata-rata {ramai['Rata_per_Hari']:.1f} pengunjung per hari), "
                   f"sedangkan {sepi['Hari']} paling sepi ({sepi['Rata_per_Hari']:.1f}).",
                   f"Tawarkan paket khusus hari {sepi['Hari']} (diskon rombongan sekolah, paket edukasi, atau homestay "
                   f"hari kerja) agar pemandu dan homestay terpakai merata sepanjang minggu.")

    # --- Aturan 3 dan 4: saluran pemasaran ---
    rs = ringkas_per(df, "Saluran_Pemasaran")
    if len(rs) >= 2:
        alok = hitung_alokasi(df, 1_000_000)
        atas, bawah = alok.iloc[0], alok.iloc[-1]
        porsi = rs.set_index("Saluran_Pemasaran")
        tambah("💰", "Prioritaskan anggaran pada saluran paling produktif", "Tinggi",
               f"{atas['Saluran_Pemasaran']} menyumbang {persen(porsi.loc[atas['Saluran_Pemasaran'], 'Porsi_Pengunjung'])} "
               f"pengunjung dan {persen(porsi.loc[atas['Saluran_Pemasaran'], 'Porsi_Pendapatan'])} pendapatan. "
               f"{bawah['Saluran_Pemasaran']} paling kecil kontribusinya "
               f"({persen(porsi.loc[bawah['Saluran_Pemasaran'], 'Porsi_Pengunjung'])} pengunjung).",
               f"Alokasikan sekitar {persen(atas['Porsi_Usulan'])} anggaran promosi ke {atas['Saluran_Pemasaran']}. "
               f"Evaluasi {bawah['Saluran_Pemasaran']} lewat uji coba kecil sebelum menambah biaya. "
               f"Hitung rincinya di tab Simulasi Alokasi Anggaran.")
        teratas = rs.iloc[0]
        if teratas["Porsi_Pengunjung"] > 0.40:
            tambah("⚠️", "Kurangi ketergantungan pada satu saluran", "Tinggi",
                   f"{teratas['Saluran_Pemasaran']} menghasilkan {persen(teratas['Porsi_Pengunjung'])} dari seluruh pengunjung. "
                   f"Bila saluran ini terganggu, kunjungan bisa turun tajam.",
                   f"Kembangkan saluran lain (konten Instagram dan TikTok, kerja sama langsung dengan sekolah dan kampus) "
                   f"sambil menjaga hubungan baik dengan {teratas['Saluran_Pemasaran']}.")

    # --- Aturan 5: segmen pengunjung ---
    sg = ringkas_per(df, "Kategori_Pengunjung")
    if len(sg) >= 2:
        terbesar = sg.iloc[0]
        bernilai = sg.sort_values("Pendapatan_per_Pengunjung", ascending=False).iloc[0]
        kat_bernilai = bernilai["Kategori_Pengunjung"]
        saluran_bernilai = (df[df["Kategori_Pengunjung"] == kat_bernilai]
                            .groupby("Saluran_Pemasaran")["Jumlah_Pengunjung"].sum().idxmax())
        if terbesar["Kategori_Pengunjung"] == kat_bernilai:
            temuan = (f"{kat_bernilai} paling banyak ({persen(terbesar['Porsi_Pengunjung'])} pengunjung) sekaligus paling "
                      f"bernilai ({rupiah(bernilai['Pendapatan_per_Pengunjung'])} per pengunjung).")
            aksi = (f"Jadikan {kat_bernilai} sasaran utama: rancang paket khusus dan perkuat promosi lewat {saluran_bernilai}.")
        else:
            temuan = (f"{terbesar['Kategori_Pengunjung']} paling banyak ({persen(terbesar['Porsi_Pengunjung'])} pengunjung). "
                      f"{kat_bernilai} memberi pendapatan per pengunjung tertinggi ({rupiah(bernilai['Pendapatan_per_Pengunjung'])}) "
                      f"dan paling sering datang lewat {saluran_bernilai}.")
            aksi = (f"Rancang paket khusus {kat_bernilai} dan promosikan lewat {saluran_bernilai}, "
                    f"sambil menjaga layanan {terbesar['Kategori_Pengunjung']} sebagai basis volume.")
        tambah("🎯", "Garap segmen bernilai tinggi", "Sedang", temuan, aksi)

    # --- Aturan 6: kepuasan ---
    sg_puas = sg.dropna(subset=["Kepuasan"])
    if not sg_puas.empty:
        rendah = sg_puas.sort_values("Kepuasan").iloc[0]
        rata_puas = df["Tingkat_Kepuasan"].mean()
        prioritas = "Tinggi" if rendah["Kepuasan"] < 4.0 else "Sedang" if rendah["Kepuasan"] < 4.3 else "Info"
        temuan = f"Kepuasan rata-rata {rata_puas:.1f} dari 5."
        if len(sg_puas) >= 2:
            temuan += f" Segmen terendah: {rendah['Kategori_Pengunjung']} ({rendah['Kepuasan']:.1f})."
        aksi = ("Kumpulkan masukan langsung dari segmen ini lewat survei singkat di akhir kunjungan, lalu perbaiki "
                "titik yang paling sering dikeluhkan." if prioritas != "Info" else
                "Pertahankan kualitas layanan dan jadikan ulasan positif sebagai bahan promosi di media sosial.")
        tambah("⭐", "Jaga dan tingkatkan kepuasan", prioritas, temuan, aksi)
    else:
        tambah("⭐", "Data kepuasan belum tersedia", "Info",
               "Belum ada baris yang memiliki rating kepuasan.",
               "Minta petugas mengisi rating pada form Input Data agar kualitas layanan dapat dipantau.")

    # --- Aturan 7: asal kota domestik ---
    kt = ringkas_per(df[df["Asal_Kota"] != "Mancanegara"], "Asal_Kota")
    if len(kt) >= 2:
        besar, kecil = kt.iloc[0], kt.iloc[-1]
        tambah("🗺️", "Perluas jangkauan ke kota yang masih kecil kontribusinya", "Sedang",
               f"Kota asal domestik terbesar: {besar['Asal_Kota']} ({persen(besar['Porsi_Pengunjung'])}). "
               f"Terkecil: {kecil['Asal_Kota']} ({persen(kecil['Porsi_Pengunjung'])}).",
               f"Pertahankan pasar {besar['Asal_Kota']}. Uji kampanye digital bertarget di {kecil['Asal_Kota']} dan "
               f"kota sekitarnya, lalu bandingkan hasilnya bulan berikutnya.")

    hasil.sort(key=lambda r: URUTAN_PRIORITAS[r["Prioritas"]])   # sort stabil: urutan aturan tetap terjaga
    return hasil


def buat_narasi(df: pd.DataFrame, periode: str) -> list[dict]:
    """
    Ringkasan eksekutif otomatis (data storytelling): tiap bagian berupa dict {Judul, Teks}.
    Kalimat dirangkai dari angka-angka di data, sehingga laporan selalu ikut berubah
    bila periode atau data berubah.
    """
    bagian: list[dict] = []
    total_p = int(df["Jumlah_Pengunjung"].sum())
    total_r = float(df["Total_Pendapatan"].sum())
    n = len(df)
    rata_puas = df["Tingkat_Kepuasan"].mean()

    teks = (f"Pada periode {periode.lower()} ({tgl_id(df['Tanggal'].min())} sampai {tgl_id(df['Tanggal'].max())}), "
            f"Kampung Majapahit Bejijong mencatat {angka_id(total_p)} pengunjung dari {angka_id(n)} kunjungan, "
            f"dengan pendapatan tiket {rupiah(total_r)}. Satu kunjungan rata-rata membawa {total_p / max(n, 1):.1f} orang, "
            f"dan setiap pengunjung menyumbang {rupiah(total_r / max(total_p, 1))}.")
    if pd.notna(rata_puas):
        teks += f" Tingkat kepuasan rata-rata {rata_puas:.1f} dari 5."
    bagian.append({"Judul": "Gambaran umum", "Teks": teks})

    bulanan = ringkas_bulanan(df)
    kandidat = pilih_bulan_pembanding(bulanan)
    if len(kandidat) >= 2:
        puncak = kandidat.loc[kandidat["Pengunjung"].idxmax()]
        sepi = kandidat.loc[kandidat["Pengunjung"].idxmin()]
        teks = (f"Bulan teramai adalah {label_bulan(puncak['Bulan'], True)} ({angka_id(puncak['Pengunjung'])} pengunjung), "
                f"sedangkan bulan tersepi adalah {label_bulan(sepi['Bulan'], True)} ({angka_id(sepi['Pengunjung'])}).")
        lengkap = bulanan[bulanan["Lengkap"]]
        if len(lengkap) >= 2:
            terakhir, sebelumnya = lengkap.iloc[-1], lengkap.iloc[-2]
            if sebelumnya["Pengunjung"] > 0:
                selisih = terakhir["Pengunjung"] / sebelumnya["Pengunjung"] - 1
                arah = "naik" if selisih >= 0 else "turun"
                teks += (f" Dibanding {label_bulan(sebelumnya['Bulan'], True)}, kunjungan {label_bulan(terakhir['Bulan'], True)} "
                         f"{arah} {persen(abs(selisih))}.")
    else:
        teks = "Data belum mencakup cukup bulan untuk membandingkan tren dari bulan ke bulan."
    bagian.append({"Judul": "Tren kunjungan", "Teks": teks})

    sg = ringkas_per(df, "Kategori_Pengunjung")
    terbesar = sg.iloc[0]
    teks = (f"Segmen terbesar adalah {terbesar['Kategori_Pengunjung']} dengan {persen(terbesar['Porsi_Pengunjung'])} "
            f"pengunjung dan {persen(terbesar['Porsi_Pendapatan'])} pendapatan.")
    if len(sg) >= 2:
        bernilai = sg.sort_values("Pendapatan_per_Pengunjung", ascending=False).iloc[0]
        if bernilai["Kategori_Pengunjung"] != terbesar["Kategori_Pengunjung"]:
            teks += (f" Pendapatan per pengunjung tertinggi justru dari {bernilai['Kategori_Pengunjung']} "
                     f"({rupiah(bernilai['Pendapatan_per_Pengunjung'])}).")
    bagian.append({"Judul": "Segmen pengunjung", "Teks": teks})

    rs = ringkas_per(df, "Saluran_Pemasaran")
    teratas = rs.iloc[0]
    teks = (f"{teratas['Saluran_Pemasaran']} adalah saluran paling efektif: {persen(teratas['Porsi_Pengunjung'])} "
            f"pengunjung dan {persen(teratas['Porsi_Pendapatan'])} pendapatan.")
    if len(rs) >= 2:
        terkecil = rs.iloc[-1]
        teks += (f" Saluran dengan kontribusi terkecil adalah {terkecil['Saluran_Pemasaran']} "
                 f"({persen(terkecil['Porsi_Pengunjung'])} pengunjung).")
    bagian.append({"Judul": "Efektivitas pemasaran", "Teks": teks})

    sg_puas = sg.dropna(subset=["Kepuasan"])
    if not sg_puas.empty and pd.notna(rata_puas):
        tinggi = sg_puas.sort_values("Kepuasan", ascending=False).iloc[0]
        teks = f"Kepuasan rata-rata {rata_puas:.1f} dari 5."
        if len(sg_puas) >= 2:
            rendah = sg_puas.sort_values("Kepuasan").iloc[0]
            teks += (f" Segmen paling puas: {tinggi['Kategori_Pengunjung']} ({tinggi['Kepuasan']:.1f}). "
                     f"Segmen yang perlu diperhatikan: {rendah['Kategori_Pengunjung']} ({rendah['Kepuasan']:.1f}).")
        bagian.append({"Judul": "Kepuasan pengunjung", "Teks": teks})

    rekom = buat_rekomendasi(df)[:3]
    daftar = "\n".join(f"{i}. {r['Judul']}: {r['Aksi']}" for i, r in enumerate(rekom, start=1))   # satu baris per langkah
    bagian.append({"Judul": "Langkah yang disarankan", "Teks": daftar})
    return bagian


# ============================================================================
# 6. KOMPONEN & HALAMAN
# ============================================================================
def banner(label: str, judul: str, deskripsi: str) -> None:
    """Banner header bernuansa terakota dan emas di bagian atas setiap halaman."""
    st.markdown(
        f'<div class="banner"><small>{label}</small><h1>{judul}</h1><p>{deskripsi}</p></div><div class="bata"></div>',
        unsafe_allow_html=True,
    )


def sidebar_navigasi() -> str:
    """Sidebar setelah login: info akun aktif dan menu navigasi (st.sidebar.radio)."""
    with st.sidebar:
        st.markdown(f"<h3 style='margin:0;color:{TERAKOTA}'>🏛️ Kampung Majapahit</h3>", unsafe_allow_html=True)
        st.caption("Bejijong, Trowulan, Mojokerto")
        st.markdown(
            html_rapat(f"""<div class="akun"><b>{st.session_state['nama']}</b><br>
            <span>{st.session_state['peran']} · @{st.session_state['username']}</span></div>"""),
            unsafe_allow_html=True,
        )
        pilihan = st.radio("Menu", list(MENU.values()), key="menu", label_visibility="collapsed")
        st.markdown("---")
    return pilihan


# ---------- Menu 1: Beranda ----------
def halaman_beranda() -> None:
    banner("Selamat datang", f"Halo, {st.session_state['nama']}",
           "Sistem informasi dan dashboard analitik untuk memantau kunjungan, pemasaran, dan pendapatan "
           "Desa Wisata Kampung Majapahit Bejijong.")

    # Ringkasan cepat dari data yang tersedia
    df, sumber, _ = data_dasar()
    gabungan = pd.concat([df, baca_input_petugas()], ignore_index=True)
    k1, k2, k3 = st.columns(3)
    k1.metric("Pengunjung tercatat", f"{angka_id(gabungan['Jumlah_Pengunjung'].sum())} Orang")
    k2.metric("Total pendapatan tiket", rupiah(gabungan["Total_Pendapatan"].sum()))
    k3.metric("Periode data", f"{gabungan['Tanggal'].min():%d %b %Y} - {gabungan['Tanggal'].max():%d %b %Y}")
    st.caption(f"Sumber ringkasan: {sumber} dan data input petugas.")
    st.write("")

    kiri, kanan = st.columns(2)
    with kiri:
        st.markdown(html_rapat(
            """<div class="kartu"><h4>Tentang proyek</h4>
            <p>Aplikasi ini adalah luaran proyek akhir Studi Independen Bersertifikat (MSIB) mahasiswa D4 Manajemen
            Informatika, Universitas Negeri Surabaya, bersama PT Mitra Talenta Group (Celerates). Tujuannya
            mengubah data operasional desa wisata menjadi informasi yang mudah dibaca: tren kunjungan, asal wisatawan,
            efektivitas saluran promosi, dan pendapatan tiket. Hasilnya membantu pengelola menyusun strategi promosi
            berbasis data.</p></div>"""), unsafe_allow_html=True)
    with kanan:
        st.markdown(html_rapat(
            """<div class="kartu"><h4>Mitra desa wisata</h4>
            <p><b>LP3SDM Sanggar Bhagaskara</b><br>Kampung Majapahit Bejijong<br>
            Jl. Kebudayaan 3, Bejijong, Trowulan, Mojokerto 61362.</p>
            <p>Desa wisata warisan Majapahit dengan homestay berarsitektur bata merah, kerajinan patung kuningan,
            dan atraksi seni budaya. Pengelola membutuhkan pengolahan data dan optimalisasi promosi, dan aplikasi ini
            dirancang untuk kebutuhan tersebut.</p></div>"""), unsafe_allow_html=True)

    st.write("")
    st.subheader("Panduan penggunaan")
    langkah = [
        ("Pantau di Analytics Dashboard", "Lihat KPI, tren, dan efektivitas saluran. Unggah Excel atau CSV di sidebar bila punya data sendiri."),
        ("Pahami data lewat EDA", "Cek kualitas data, sebaran, pola waktu, dan segmen pengunjung sebelum mengambil keputusan."),
        ("Ambil keputusan", "Rekomendasi Strategis memberi saran dan simulasi anggaran. Laporan Ringkas merangkumnya untuk diunduh."),
        ("Catat kunjungan harian", "Petugas mengisi form di menu Input Data. Total pendapatan dihitung otomatis dan langsung masuk analisis."),
    ]
    kolom = st.columns(4)
    for i, (judul, isi) in enumerate(langkah):
        kolom[i].markdown(
            html_rapat(f"""<div class="kartu" style="min-height:215px"><div class="langkah"><b class="no">{i + 1}</b>
            <div><b>{judul}</b><p style="margin:4px 0 0 0">{isi}</p></div></div></div>"""),
            unsafe_allow_html=True,
        )


# ---------- Menu 2: Analytics Dashboard (filter sidebar dinamis) ----------
def sidebar_filter(data: pd.DataFrame, sumber: str):
    """Filter dinamis di sidebar. Mengembalikan DataFrame yang sudah difilter."""
    sb = st.sidebar
    sb.subheader("Sumber data")
    versi = st.session_state.get("versi_uploader", 0)   # naikkan versi untuk mengosongkan widget uploader
    berkas = sb.file_uploader("Unggah data kunjungan", type=["xlsx", "xls", "csv"], key=f"uploader_{versi}",
                              help="Kolom wajib: " + ", ".join(KOLOM_WAJIB))
    if berkas is not None:
        tanda_berkas = (berkas.name, berkas.size)
        if st.session_state.get("tanda_berkas") != tanda_berkas:   # proses hanya saat file baru
            st.session_state["tanda_berkas"] = tanda_berkas
            try:
                st.session_state["unggahan"] = {"nama": berkas.name, "df": bersihkan(baca_unggahan(berkas))}
            except Exception as galat:  # noqa: BLE001
                st.session_state.pop("unggahan", None)
                st.session_state["peringatan_unggah"] = f"File '{berkas.name}' tidak bisa dipakai ({galat}). Menampilkan data contoh."
            st.rerun()
    if "unggahan" in st.session_state:
        if sb.button("Hapus unggahan, pakai data contoh", **lebar_penuh()):
            st.session_state.pop("unggahan", None)
            st.session_state.pop("tanda_berkas", None)
            st.session_state["versi_uploader"] = versi + 1
            st.rerun()
    sb.caption(f"Sumber aktif: {sumber}")

    sertakan = sb.checkbox("Sertakan data input petugas", value=st.session_state.get("pref_petugas", True),
                           help="Gabungkan data dari form Input Data Pengunjung ke dashboard dan menu analisis lain.")
    st.session_state["pref_petugas"] = sertakan   # dibaca data_analisis() agar pilihan ini ikut ke menu lain
    df = data
    if sertakan:
        tambahan = baca_input_petugas()
        if not tambahan.empty:
            df = pd.concat([data, tambahan], ignore_index=True)
            sb.caption(f"+ {angka_id(len(tambahan))} entri dari petugas")

    sb.markdown("---")
    sb.subheader("Filter data")
    tanda = f"{sumber}-{len(df)}"   # key widget ikut berubah saat data berubah, sehingga filter kembali ke bawaan
    tgl_min, tgl_maks = df["Tanggal"].min().date(), df["Tanggal"].max().date()
    rentang = sb.date_input("Rentang tanggal", value=(tgl_min, tgl_maks), min_value=tgl_min, max_value=tgl_maks, key=f"tgl_{tanda}")
    if isinstance(rentang, (tuple, list)):
        mulai, selesai = (rentang[0], rentang[-1]) if len(rentang) else (tgl_min, tgl_maks)
    else:
        mulai = selesai = rentang

    def multi(label: str, kolom: str, kunci: str) -> list[str]:
        pilihan = sorted(df[kolom].dropna().unique().tolist())
        return sb.multiselect(label, pilihan, default=pilihan, key=f"{kunci}_{tanda}")

    kat = multi("Kategori pengunjung", "Kategori_Pengunjung", "kat")
    sal = multi("Saluran pemasaran", "Saluran_Pemasaran", "sal")
    kota = multi("Asal kota", "Asal_Kota", "kota")

    # Semua filter diterapkan sekaligus. Tiap perubahan widget menjalankan ulang skrip,
    # sehingga KPI dan grafik diperbarui secara real-time.
    mask = ((df["Tanggal"].dt.date >= mulai) & (df["Tanggal"].dt.date <= selesai)
            & df["Kategori_Pengunjung"].isin(kat) & df["Saluran_Pemasaran"].isin(sal) & df["Asal_Kota"].isin(kota))
    hasil = df.loc[mask].copy()
    sb.caption(f"{angka_id(len(hasil))} dari {angka_id(len(df))} baris tampil")
    return hasil


def halaman_dashboard() -> None:
    banner("Analytics Dashboard", "Dashboard Analytics Desa Wisata",
           "Pantau kunjungan, efektivitas pemasaran, dan pendapatan tiket secara real-time.")
    data, sumber, peringatan = data_dasar()
    if peringatan:
        st.sidebar.warning(peringatan)
    df = sidebar_filter(data, sumber)

    if df.empty:
        st.warning("Tidak ada data untuk kombinasi filter ini. Ubah rentang tanggal atau pilihan filter di sidebar.")
        return

    # ---- Baris KPI ----
    total_pengunjung = int(df["Jumlah_Pengunjung"].sum())
    total_pendapatan = float(df["Total_Pendapatan"].sum())
    per_saluran = df.groupby("Saluran_Pemasaran")["Jumlah_Pengunjung"].sum().sort_values(ascending=False)
    saluran_top, pengunjung_top = per_saluran.index[0], int(per_saluran.iloc[0])
    rata_puas = df["Tingkat_Kepuasan"].mean()   # entri petugas tanpa rating diabaikan otomatis (NaN)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Pengunjung", f"{angka_id(total_pengunjung)} Orang")
    k2.metric("Total Pendapatan Tiket", rupiah(total_pendapatan))
    k3.metric("Saluran Terefektif", saluran_top, f"{angka_id(pengunjung_top)} pengunjung ({pengunjung_top / total_pengunjung:.0%})",
              delta_color="off")
    k4.metric("Rata-rata Kepuasan", f"⭐ {rata_puas:.1f} / 5.0" if pd.notna(rata_puas) else "Belum ada data")
    st.write("")

    tab1, tab2, tab3 = st.tabs(["📈 Tren & Demografi", "🎯 Efektivitas Pemasaran", "📑 Tabel Data & Ekspor"])

    # ---- Tab 1: tren dan demografi ----
    with tab1:
        kiri, kanan = st.columns([3, 2])
        with kiri:
            gran = st.radio("Tampilan tren", ["Harian", "Mingguan", "Bulanan"], horizontal=True, index=2)
            if gran == "Harian":
                periode = df["Tanggal"].dt.normalize()
            elif gran == "Mingguan":
                periode = df["Tanggal"].dt.to_period("W").dt.start_time
            else:
                periode = df["Tanggal"].dt.to_period("M").dt.start_time
            tren = df.groupby(periode)["Jumlah_Pengunjung"].sum().reset_index()
            tren.columns = ["Periode", "Pengunjung"]
            fig = px.line(tren, x="Periode", y="Pengunjung", markers=True, title=f"Tren jumlah pengunjung ({gran.lower()})",
                          color_discrete_sequence=[TERAKOTA])
            if gran == "Harian" and len(tren) >= 7:
                tren["Rata-rata 7 titik"] = tren["Pengunjung"].rolling(7, min_periods=1).mean()
                fig.add_scatter(x=tren["Periode"], y=tren["Rata-rata 7 titik"], mode="lines",
                                name="Rata-rata 7 titik", line=dict(color=EMAS, width=3))
            fig.update_traces(hovertemplate="%{x|%d %b %Y}<br>%{y:,.0f} orang")
            fig.update_layout(yaxis_title="Jumlah pengunjung", xaxis_title=None, showlegend=False)
            tampil_grafik(fig)
        with kanan:
            kat = df.groupby("Kategori_Pengunjung")["Jumlah_Pengunjung"].sum().reset_index()
            fig = px.pie(kat, names="Kategori_Pengunjung", values="Jumlah_Pengunjung", hole=0.55,
                         title="Proporsi kategori pengunjung", color_discrete_sequence=WARNA_GRAFIK)
            fig.update_traces(textinfo="percent", hovertemplate="%{label}<br>%{value:,.0f} orang (%{percent})")
            fig.update_layout(legend=dict(orientation="h", y=-0.15))
            tampil_grafik(fig)

    # ---- Tab 2: efektivitas pemasaran ----
    with tab2:
        ringkas = (df.groupby("Saluran_Pemasaran")
                   .agg(Pengunjung=("Jumlah_Pengunjung", "sum"), Pendapatan=("Total_Pendapatan", "sum"))
                   .reset_index())
        kiri, kanan = st.columns(2)
        with kiri:
            urut = ringkas.sort_values("Pengunjung", ascending=False)
            fig = px.bar(urut, x="Saluran_Pemasaran", y="Pengunjung", text="Pengunjung",
                         title="Jumlah pengunjung per saluran pemasaran", color_discrete_sequence=[TERAKOTA])
            fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
            fig.update_layout(xaxis_title=None, yaxis_title="Jumlah pengunjung")
            fig.update_xaxes(categoryorder="array", categoryarray=urut["Saluran_Pemasaran"])
            tampil_grafik(fig)
        with kanan:
            urut = ringkas.sort_values("Pendapatan", ascending=False)
            fig = px.bar(urut, x="Saluran_Pemasaran", y="Pendapatan", title="Total pendapatan per saluran pemasaran",
                         color_discrete_sequence=[EMAS])
            fig.update_traces(hovertemplate="%{x}<br>Rp %{y:,.0f}")
            fig.update_layout(xaxis_title=None, yaxis_title="Pendapatan (Rp)")
            fig.update_xaxes(categoryorder="array", categoryarray=urut["Saluran_Pemasaran"])
            tampil_grafik(fig)

    # ---- Tab 3: tabel dan ekspor ----
    with tab3:
        cari = st.text_input("Cari di semua kolom teks", placeholder="Contoh: Instagram, Surabaya, Edukasi")
        tabel = df.copy()
        if cari.strip():
            cocok = np.zeros(len(tabel), dtype=bool)
            for kolom in ["Kategori_Pengunjung", "Saluran_Pemasaran", "Asal_Kota"]:
                cocok |= tabel[kolom].str.contains(cari.strip(), case=False, na=False, regex=False).to_numpy()
            tabel = tabel[cocok]
        st.caption(f"{angka_id(len(tabel))} baris. Klik judul kolom untuk mengurutkan.")
        tampil = tabel.copy()
        tampil["Tanggal"] = tampil["Tanggal"].dt.date
        tampil_tabel(tampil, hide_index=True, height=440, column_config={
            "Tanggal": st.column_config.DateColumn(format="YYYY-MM-DD"),
            "Jumlah_Pengunjung": st.column_config.NumberColumn(format="%d"),
            "Harga_Tiket_Per_Orang": st.column_config.NumberColumn(format="Rp %d"),
            "Total_Pendapatan": st.column_config.NumberColumn(format="Rp %d"),
            "Tingkat_Kepuasan": st.column_config.NumberColumn(format="%.1f"),
        })
        ekspor = tabel.copy()
        ekspor["Tanggal"] = ekspor["Tanggal"].dt.strftime("%Y-%m-%d")
        unduh_csv = ekspor.to_csv(index=False).encode("utf-8-sig")   # utf-8-sig agar terbaca benar di Excel
        penyangga = io.BytesIO()
        with pd.ExcelWriter(penyangga, engine="openpyxl") as penulis:
            ekspor.to_excel(penulis, index=False, sheet_name="Laporan_Kunjungan")
        d1, d2, _ = st.columns([1, 1, 3])
        d1.download_button("⬇️ Unduh Excel", penyangga.getvalue(), "laporan_kunjungan.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2.download_button("⬇️ Unduh CSV", unduh_csv, "laporan_kunjungan.csv", mime="text/csv")


# ---------- Komponen bersama: alur CRISP-DM dan tombol pindah menu ----------
FASE_CRISP = ["Business Understanding", "Data Understanding", "Data Preparation", "Modeling", "Evaluation", "Deployment"]


def tampil_crisp(aktif: list[int]) -> None:
    """Menampilkan enam fase CRISP-DM dan menyorot fase yang dikerjakan di halaman ini (indeks mulai dari 0)."""
    chip = "".join(
        f'<span class="{"aktif" if i in aktif else ""}">{i + 1}. {nama}</span>' for i, nama in enumerate(FASE_CRISP)
    )
    st.markdown(f'<div class="crisp">{chip}</div>', unsafe_allow_html=True)


def pindah_menu(tujuan: str) -> None:
    """
    Callback tombol pindah menu. Callback berjalan SEBELUM skrip dirender ulang, sehingga boleh
    mengubah nilai radio sidebar. Bila diubah setelah radio dibuat, Streamlit akan memunculkan error.
    """
    st.session_state["menu"] = tujuan


def navigasi_antar_menu(aktif: str) -> None:
    """Baris tombol di bagian bawah setiap halaman untuk berpindah ke menu lain tanpa membuka sidebar."""
    daftar = list(MENU.values())
    posisi = daftar.index(aktif)
    berikutnya = daftar[posisi + 1] if posisi + 1 < len(daftar) else None
    lainnya = [m for m in daftar if m != aktif]

    st.markdown('<div class="navbawah"></div>', unsafe_allow_html=True)
    st.markdown("##### Pindah ke menu lain")
    for awal in range(0, len(lainnya), 3):          # tiga tombol per baris
        kolom = st.columns(3)
        for kol, tujuan in zip(kolom, lainnya[awal:awal + 3]):
            kol.button(tujuan, key=f"nav_{aktif}_{tujuan}", on_click=pindah_menu, args=(tujuan,),
                       type="primary" if tujuan == berikutnya else "secondary", **lebar_penuh())
    if berikutnya:
        st.caption(f"Menu berikutnya yang disarankan ditandai warna terakota: {berikutnya}")


# ---------- Menu 3: Eksplorasi Data (EDA) ----------
def halaman_eda() -> None:
    banner("Exploratory Data Analysis", "Eksplorasi Data (EDA)",
           "Pahami kualitas, sebaran, pola waktu, dan segmen pengunjung sebelum mengambil keputusan.")
    tampil_crisp([1, 2])
    data, sumber = data_analisis()
    df, periode = pilih_periode(data, "eda")
    if df.empty:
        st.warning("Tidak ada data pada periode ini.")
        return
    st.caption(f"{sumber} · {periode} · {angka_id(len(df))} baris · {tgl_id(df['Tanggal'].min())} sampai {tgl_id(df['Tanggal'].max())}")
    rentang_hari = (df["Tanggal"].max() - df["Tanggal"].min()).days + 1

    tab_kualitas, tab_sebaran, tab_waktu, tab_segmen = st.tabs(
        ["🧹 Kualitas Data", "📊 Sebaran", "📅 Pola Waktu", "🧩 Segmentasi"])

    # ---- Tab 1: kualitas data ----
    with tab_kualitas:
        st.caption("Fase CRISP-DM: Data Understanding dan Data Preparation. Data yang salah akan menghasilkan rekomendasi yang salah.")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Jumlah baris", angka_id(len(df)))
        m2.metric("Rentang data", f"{angka_id(rentang_hari)} hari")
        m3.metric("Sel kosong", angka_id(int(df.isna().sum().sum())))
        m4.metric("Baris kembar", angka_id(int(df.duplicated().sum())))
        st.write("")

        cek = periksa_kualitas(df)
        bermasalah = int((cek["Baris bermasalah"] > 0).sum())
        if bermasalah == 0:
            st.success("Semua pemeriksaan aturan data lolos.")
        else:
            st.warning(f"{bermasalah} dari {len(cek)} pemeriksaan menemukan baris yang perlu dicek. "
                       "Entri petugas di luar rentang tipikal (misalnya rombongan lebih dari 50 orang) ikut terhitung di sini.")
        tampil_tabel(cek, hide_index=True)

        kiri, kanan = st.columns(2)
        with kiri:
            st.markdown("**Nilai kosong per kolom**")
            kosong = df.isna().sum().reset_index()
            kosong.columns = ["Kolom", "Kosong"]
            kosong["Persen"] = (kosong["Kosong"] / len(df) * 100).round(1)
            tampil_tabel(kosong, hide_index=True)
            if int(df["Tingkat_Kepuasan"].isna().sum()):
                st.caption("Kepuasan kosong wajar bila petugas tidak mengisi rating pada form.")
        with kanan:
            st.markdown("**Statistik deskriptif**")
            angka = ["Jumlah_Pengunjung", "Harga_Tiket_Per_Orang", "Total_Pendapatan", "Tingkat_Kepuasan"]
            stat = df[angka].agg(["mean", "median", "min", "max", "std"]).T.round(1).reset_index()
            stat.columns = ["Kolom", "Rata-rata", "Median", "Minimum", "Maksimum", "Simpangan baku"]
            tampil_tabel(stat, hide_index=True)

    # ---- Tab 2: sebaran ----
    with tab_sebaran:
        st.caption("Fase CRISP-DM: Data Understanding. Lihat bentuk sebaran sebelum percaya pada angka rata-rata.")
        a, b = st.columns(2)
        with a:
            fig = px.histogram(df, x="Jumlah_Pengunjung", nbins=12, title="Ukuran rombongan per kunjungan",
                               color_discrete_sequence=[TERAKOTA])
            fig.update_layout(bargap=0.06, xaxis_title="Orang per kunjungan", yaxis_title="Jumlah kunjungan")
            tampil_grafik(fig)
        with b:
            fig = px.histogram(df, x="Harga_Tiket_Per_Orang", nbins=8, title="Harga tiket per orang",
                               color_discrete_sequence=[EMAS])
            fig.update_layout(bargap=0.06, xaxis_title="Rupiah per orang", yaxis_title="Jumlah kunjungan")
            tampil_grafik(fig)
        c, d = st.columns(2)
        with c:
            fig = px.histogram(df.dropna(subset=["Tingkat_Kepuasan"]), x="Tingkat_Kepuasan", nbins=10,
                               title="Tingkat kepuasan", color_discrete_sequence=[KAYU])
            fig.update_layout(bargap=0.06, xaxis_title="Rating (1 sampai 5)", yaxis_title="Jumlah kunjungan")
            tampil_grafik(fig)
        with d:
            fig = px.box(df, x="Kategori_Pengunjung", y="Jumlah_Pengunjung", title="Ukuran rombongan per kategori",
                         color_discrete_sequence=[TERAKOTA])
            fig.update_layout(xaxis_title=None, yaxis_title="Orang per kunjungan")
            tampil_grafik(fig)

        miring = df["Jumlah_Pengunjung"].skew()
        median, rata = df["Jumlah_Pengunjung"].median(), df["Jumlah_Pengunjung"].mean()
        if pd.notna(miring) and miring > 0.5:
            st.info(f"Ukuran rombongan condong ke kanan (rata-rata {rata:.1f}, median {median:.0f}): kebanyakan kunjungan "
                    "kecil, tetapi ada rombongan besar yang menarik rata-rata naik. Rombongan besar di sini umumnya ciri segmen "
                    "edukasi, bukan kesalahan data.")
        else:
            st.info(f"Ukuran rombongan cukup merata (rata-rata {rata:.1f}, median {median:.0f}).")

    # ---- Tab 3: pola waktu ----
    with tab_waktu:
        st.caption("Fase CRISP-DM: Data Understanding. Pola waktu menentukan kapan promosi dan kapasitas harus disiapkan.")
        bulanan = ringkas_bulanan(df)
        pembanding = pilih_bulan_pembanding(bulanan)
        indeks_puncak = pembanding["Pengunjung"].idxmax()
        bulanan["Status"] = np.where(bulanan["Lengkap"], "Bulan lengkap", "Bulan parsial")
        if len(pembanding) >= 2:
            bulanan.loc[indeks_puncak, "Status"] = "Bulan puncak"
        bulanan["Label"] = bulanan["Bulan"].map(label_bulan)
        fig = px.bar(bulanan, x="Label", y="Pengunjung", color="Status", text="Pengunjung",
                     title="Pengunjung per bulan",
                     color_discrete_map={"Bulan puncak": EMAS, "Bulan lengkap": TERAKOTA, "Bulan parsial": "#D9B8A8"})
        fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
        fig.update_layout(xaxis_title=None, yaxis_title="Jumlah pengunjung", legend_title_text=None)
        fig.update_xaxes(categoryorder="array", categoryarray=bulanan["Label"])
        tampil_grafik(fig)
        st.caption("Bulan parsial (awal atau akhir rentang data tidak penuh) berwarna pucat dan tidak dipakai menentukan bulan puncak.")

        if rentang_hari >= 14:
            kiri, kanan = st.columns([2, 3])
            hr = rata_per_hari_minggu(df)
            hr["Jenis"] = np.where(hr["Hari"].isin(["Sabtu", "Minggu"]), "Akhir pekan", "Hari kerja")
            with kiri:
                fig = px.bar(hr, x="Hari", y="Rata_per_Hari", color="Jenis", title="Rata-rata pengunjung per hari",
                             color_discrete_map={"Akhir pekan": EMAS, "Hari kerja": TERAKOTA})
                fig.update_layout(xaxis_title=None, yaxis_title="Pengunjung per hari", legend_title_text=None)
                fig.update_xaxes(categoryorder="array", categoryarray=NAMA_HARI)
                tampil_grafik(fig)
            with kanan:
                tmp = df.assign(Hari=df["Tanggal"].dt.dayofweek.map(lambda i: NAMA_HARI[i]),
                                Bulan=df["Tanggal"].dt.to_period("M").dt.start_time)
                piv = (tmp.pivot_table(index="Hari", columns="Bulan", values="Jumlah_Pengunjung", aggfunc="sum", fill_value=0)
                       .reindex(NAMA_HARI).fillna(0))
                fig = px.imshow(piv.to_numpy(), x=[label_bulan(c) for c in piv.columns], y=NAMA_HARI, text_auto=True,
                                aspect="auto", color_continuous_scale=SKALA_PANAS, title="Peta panas: hari x bulan")
                fig.update_layout(coloraxis_colorbar=dict(title="Pengunjung"))
                tampil_grafik(fig)
        else:
            st.info("Pola per hari membutuhkan data minimal dua minggu.")

    # ---- Tab 4: segmentasi ----
    with tab_segmen:
        st.caption("Fase CRISP-DM: Data Understanding. Segmen yang berbeda butuh pendekatan promosi yang berbeda.")
        sg = tabel_ringkas(df, "Kategori_Pengunjung", "Kategori Pengunjung")
        tampil_tabel(rapikan_angka(sg), hide_index=True)

        porsi = sg.melt(id_vars="Kategori Pengunjung", value_vars=["Porsi Pengunjung (%)", "Porsi Pendapatan (%)"],
                        var_name="Ukuran", value_name="Persen")
        fig = px.bar(porsi, x="Kategori Pengunjung", y="Persen", color="Ukuran", barmode="group",
                     title="Porsi pengunjung dibanding porsi pendapatan per segmen",
                     color_discrete_map={"Porsi Pengunjung (%)": TERAKOTA, "Porsi Pendapatan (%)": EMAS})
        fig.update_layout(xaxis_title=None, yaxis_title="Persen dari total", legend_title_text=None)
        tampil_grafik(fig)

        kiri, kanan = st.columns(2)
        with kiri:
            pivot_kota = df.pivot_table(index="Kategori_Pengunjung", columns="Asal_Kota", values="Jumlah_Pengunjung",
                                        aggfunc="sum", fill_value=0)
            fig = px.imshow(pivot_kota.to_numpy(), x=list(pivot_kota.columns), y=list(pivot_kota.index), text_auto=True,
                            aspect="auto", color_continuous_scale=SKALA_PANAS, title="Pengunjung: segmen x kota asal")
            fig.update_layout(coloraxis_showscale=False)
            tampil_grafik(fig)
        with kanan:
            pivot_saluran = df.pivot_table(index="Kategori_Pengunjung", columns="Saluran_Pemasaran",
                                           values="Jumlah_Pengunjung", aggfunc="sum", fill_value=0)
            persen_baris = pivot_saluran.div(pivot_saluran.sum(axis=1).replace(0, np.nan), axis=0).fillna(0)
            fig = px.imshow(persen_baris.to_numpy(), x=list(persen_baris.columns), y=list(persen_baris.index),
                            text_auto=".0%", aspect="auto", color_continuous_scale=SKALA_PANAS,
                            title="Saluran asal tiap segmen (% dalam segmen)")
            fig.update_layout(coloraxis_showscale=False)
            tampil_grafik(fig)

        baris_info = []
        for kategori, baris in persen_baris.iterrows():
            baris_info.append(f"- **{kategori}** paling banyak datang lewat **{baris.idxmax()}** ({persen(baris.max())} dari segmen).")
        st.markdown("**Saluran utama tiap segmen**\n\n" + "\n".join(baris_info))


# ---------- Menu 4: Rekomendasi Strategis ----------
def kartu_rekomendasi(r: dict) -> str:
    """HTML satu kartu rekomendasi. Teks di-escape karena sebagian berasal dari file unggahan pengguna."""
    prio = r["Prioritas"].lower()
    ikon, judul = r["Ikon"], escape(r["Judul"])
    temuan, aksi = escape(r["Temuan"]), escape(r["Aksi"])
    return html_rapat(f"""<div class="rekom {prio}">
        <h4>{ikon} {judul}</h4>
        <span class="chip {prio}">Prioritas {r['Prioritas']}</span>
        <p><b>Temuan:</b> {temuan}</p>
        <p class="aksi"><b>Saran tindakan:</b> {aksi}</p></div>""")


def halaman_rekomendasi() -> None:
    banner("Recommendation Engine", "Rekomendasi Strategis",
           "Saran promosi, alokasi anggaran, dan pengelolaan kapasitas yang dihitung otomatis dari data kunjungan.")
    tampil_crisp([3, 4])
    data, sumber = data_analisis()
    df, periode = pilih_periode(data, "rekomendasi")
    st.caption(f"{sumber} · {periode} · {angka_id(len(df))} baris")

    tab_utama, tab_simulasi, tab_cara = st.tabs(["🎯 Rekomendasi Utama", "💰 Simulasi Alokasi Anggaran", "🧭 Cara Kerja & Batasan"])

    # ---- Tab 1: daftar rekomendasi ----
    with tab_utama:
        rekom = buat_rekomendasi(df)
        hitung = {p: sum(1 for r in rekom if r["Prioritas"] == p) for p in URUTAN_PRIORITAS}
        m1, m2, m3 = st.columns(3)
        m1.metric("Prioritas tinggi", f"{hitung['Tinggi']} saran")
        m2.metric("Prioritas sedang", f"{hitung['Sedang']} saran")
        m3.metric("Informasi", f"{hitung['Info']} catatan")
        st.write("")
        kolom = st.columns(2)
        for i, r in enumerate(rekom):
            kolom[i % 2].markdown(kartu_rekomendasi(r), unsafe_allow_html=True)
        unduh = pd.DataFrame(rekom)[["Prioritas", "Judul", "Temuan", "Aksi"]].to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Unduh rekomendasi (CSV)", unduh, "rekomendasi_strategis.csv", mime="text/csv")

    # ---- Tab 2: simulasi alokasi anggaran ----
    with tab_simulasi:
        st.caption("Fase CRISP-DM: Modeling. Geser bobot dan ubah anggaran, hasilnya langsung dihitung ulang.")
        c1, c2 = st.columns(2)
        anggaran = c1.number_input("Total anggaran promosi periode berikutnya (Rp)", min_value=0, max_value=1_000_000_000,
                                   value=5_000_000, step=500_000)
        bobot = c2.slider("Bobot pendapatan terhadap jumlah pengunjung (%)", 0, 100, 60, 10,
                          help="100 berarti alokasi hanya mengikuti pendapatan. 0 berarti hanya mengikuti jumlah pengunjung.")
        alok = hitung_alokasi(df, anggaran, bobot / 100)
        atas = alok.iloc[0]
        k1, k2 = st.columns(2)
        k1.metric("Saluran prioritas", atas["Saluran_Pemasaran"], f"{persen(atas['Porsi_Usulan'])} dari anggaran", delta_color="off")
        k2.metric("Alokasi saluran prioritas", rupiah(atas["Alokasi_Rp"]))
        st.write("")

        fig = px.bar(alok, x="Saluran_Pemasaran", y="Alokasi_Rp", color="Status", text="Alokasi_Rp",
                     title="Usulan alokasi anggaran per saluran",
                     color_discrete_map={"Prioritas": TERAKOTA, "Pertahankan": EMAS, "Uji coba kecil": "#A89F91"})
        fig.update_traces(texttemplate="Rp %{text:,.0f}", textposition="outside", cliponaxis=False)
        fig.update_layout(xaxis_title=None, yaxis_title="Alokasi (Rp)", legend_title_text=None)
        fig.update_xaxes(categoryorder="array", categoryarray=alok["Saluran_Pemasaran"])
        tampil_grafik(fig)

        tabel = pd.DataFrame({
            "Saluran": alok["Saluran_Pemasaran"], "Status": alok["Status"],
            "Porsi usulan (%)": (alok["Porsi_Usulan"] * 100).round(1), "Alokasi (Rp)": alok["Alokasi_Rp"],
            "Pengunjung": alok["Pengunjung"], "Pendapatan (Rp)": alok["Pendapatan"], "Kepuasan": alok["Kepuasan"],
        })
        tampil_tabel(rapikan_angka(tabel), hide_index=True)
        st.download_button("⬇️ Unduh simulasi (CSV)", tabel.to_csv(index=False).encode("utf-8-sig"),
                           "simulasi_alokasi_anggaran.csv", mime="text/csv")
        st.caption("Rumus: skor = bobot x porsi pendapatan + (1 - bobot) x porsi pengunjung, disesuaikan sedikit dengan "
                   "kepuasan saluran. Porsi anggaran = skor saluran dibagi jumlah skor seluruh saluran. Ini alat bantu "
                   "diskusi, bukan angka mutlak.")

    # ---- Tab 3: cara kerja dan batasan ----
    with tab_cara:
        st.markdown("**Cara kerja mesin rekomendasi**")
        st.markdown(
            "- Mesin membaca tujuh pola dari data: bulan puncak, hari tersibuk dan tersepi, saluran paling produktif, "
            "ketergantungan pada satu saluran, segmen paling bernilai, tingkat kepuasan, dan asal kota.\n"
            "- Tiap pola diubah menjadi temuan (angka dari data) dan saran tindakan. Aturannya berbasis ambang batas "
            "yang bisa diperiksa, sehingga setiap saran dapat ditelusuri ke angka sumbernya.\n"
            "- Prioritas tinggi dipakai untuk pola yang berdampak besar pada pendapatan atau risiko, misalnya bulan puncak "
            "atau satu saluran yang menguasai lebih dari 40% pengunjung.")
        st.markdown("**Batasan yang perlu diketahui**")
        st.markdown(
            "- Data contoh adalah simulasi. Rekomendasi baru bermakna setelah memakai data kunjungan nyata dari LP3SDM Sanggar Bhagaskara.\n"
            "- Belum ada data okupansi homestay dan paket wisata. Saran kapasitas memakai beban kunjungan harian sebagai pendekatan.\n"
            "- Pola yang terlihat adalah korelasi, bukan bukti sebab-akibat. Uji saran dalam skala kecil sebelum menambah anggaran besar.\n"
            "- Makin pendek periode analisis, makin mudah hasil berubah oleh satu rombongan besar.")
        fase = pd.DataFrame({
            "Fase CRISP-DM": FASE_CRISP,
            "Wujud di aplikasi ini": [
                "Proposal MSIB dan kebutuhan mitra (halaman Beranda)",
                "Menu Eksplorasi Data: tab Kualitas Data dan Sebaran",
                "Fungsi pembersihan data dan validasi form Input Data",
                "Mesin rekomendasi berbasis aturan dan simulasi alokasi anggaran",
                "Pembacaan hasil bersama mitra, lalu uji dengan data nyata (tahap lanjutan)",
                "Aplikasi Streamlit untuk pengguna nonteknis dan Laporan Ringkas yang bisa diunduh",
            ],
        })
        tampil_tabel(fase, hide_index=True)


# ---------- Menu 5: Laporan Ringkas ----------
def tulis_sheet(penulis, tabel: pd.DataFrame, nama: str, lebar_maks: int = 45, bungkus: bool = False) -> None:
    """Menulis satu sheet Excel dengan header terakota, lebar kolom otomatis, dan baris pertama dibekukan."""
    tabel.to_excel(penulis, sheet_name=nama, index=False)
    ws = penulis.sheets[nama]
    for sel in ws[1]:
        sel.font = Font(bold=True, color="FFFFFF")
        sel.fill = PatternFill("solid", fgColor="8B3A2B")
        sel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, kolom in enumerate(tabel.columns, start=1):
        terpanjang = max([len(str(kolom))] + [len(str(v)) for v in tabel[kolom].head(300)])
        ws.column_dimensions[get_column_letter(i)].width = min(terpanjang + 3, lebar_maks)
    if bungkus:
        for baris in ws.iter_rows(min_row=2):
            for sel in baris:
                sel.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"


def buat_excel_laporan(df: pd.DataFrame, narasi: list[dict], rekom: list[dict], periode: str, sumber: str) -> bytes:
    """Menyusun laporan Excel multi-sheet: ringkasan, tabel pendukung, rekomendasi, dan data mentah."""
    bulanan = ringkas_bulanan(df)
    per_bulan = pd.DataFrame({
        "Bulan": bulanan["Bulan"].map(lambda b: label_bulan(b, True)), "Jumlah Kunjungan": bulanan["Kunjungan"],
        "Total Pengunjung": bulanan["Pengunjung"].astype(int), "Total Pendapatan (Rp)": bulanan["Pendapatan"].round(0).astype(int),
        "Keterangan": np.where(bulanan["Lengkap"], "Bulan lengkap", "Bulan parsial"),
    })
    ringkasan = pd.DataFrame(
        [{"Bagian": "Periode", "Isi": f"{periode}: {tgl_id(df['Tanggal'].min())} sampai {tgl_id(df['Tanggal'].max())}"},
         {"Bagian": "Sumber data", "Isi": sumber},
         {"Bagian": "Dibuat oleh", "Isi": f"{st.session_state.get('nama', '-')} pada {datetime.now():%d-%m-%Y %H:%M}"}]
        + [{"Bagian": b["Judul"], "Isi": b["Teks"]} for b in narasi])
    data_mentah = df.copy()
    data_mentah["Tanggal"] = data_mentah["Tanggal"].dt.strftime("%Y-%m-%d")

    penyangga = io.BytesIO()
    with pd.ExcelWriter(penyangga, engine="openpyxl") as penulis:
        tulis_sheet(penulis, ringkasan, "Ringkasan", lebar_maks=110, bungkus=True)
        tulis_sheet(penulis, per_bulan, "Per_Bulan")
        tulis_sheet(penulis, tabel_ringkas(df, "Saluran_Pemasaran", "Saluran Pemasaran"), "Per_Saluran")
        tulis_sheet(penulis, tabel_ringkas(df, "Kategori_Pengunjung", "Kategori Pengunjung"), "Per_Kategori")
        tulis_sheet(penulis, tabel_ringkas(df, "Asal_Kota", "Asal Kota"), "Per_Kota")
        tulis_sheet(penulis, pd.DataFrame(rekom)[["Prioritas", "Judul", "Temuan", "Aksi"]], "Rekomendasi", lebar_maks=70, bungkus=True)
        tulis_sheet(penulis, data_mentah, "Data")
    return penyangga.getvalue()


def buat_teks_laporan(narasi: list[dict], rekom: list[dict], periode: str, sumber: str) -> str:
    """Versi teks polos dari laporan, cocok untuk ditempel ke surel atau pesan WhatsApp."""
    baris = ["LAPORAN RINGKAS KINERJA DESA WISATA", "Kampung Majapahit Bejijong · LP3SDM Sanggar Bhagaskara", "",
             f"Periode : {periode}", f"Sumber  : {sumber}", f"Dibuat  : {datetime.now():%d-%m-%Y %H:%M}", ""]
    for b in narasi:
        baris += [b["Judul"].upper(), b["Teks"], ""]
    baris.append("REKOMENDASI")
    for i, r in enumerate(rekom, start=1):
        baris += [f"{i}. [{r['Prioritas']}] {r['Judul']}", f"   Temuan: {r['Temuan']}", f"   Tindakan: {r['Aksi']}"]
    return "\n".join(baris) + "\n"


def halaman_laporan() -> None:
    banner("Data storytelling", "Laporan Ringkas",
           "Ringkasan eksekutif yang disusun otomatis dari data, siap dibaca pimpinan dan diunduh.")
    tampil_crisp([4, 5])
    data, sumber = data_analisis()
    df, periode = pilih_periode(data, "laporan")
    st.caption(f"{sumber} · {periode} · {angka_id(len(df))} baris")

    total_p, total_r = int(df["Jumlah_Pengunjung"].sum()), float(df["Total_Pendapatan"].sum())
    rata_puas = df["Tingkat_Kepuasan"].mean()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Pengunjung", f"{angka_id(total_p)} Orang")
    k2.metric("Total Pendapatan Tiket", rupiah(total_r))
    k3.metric("Pendapatan per Pengunjung", rupiah(total_r / max(total_p, 1)))
    k4.metric("Rata-rata Kepuasan", f"⭐ {rata_puas:.1f} / 5.0" if pd.notna(rata_puas) else "Belum ada data")
    st.write("")

    st.subheader("Ringkasan eksekutif")
    narasi = buat_narasi(df, periode)
    for bagian in narasi:
        st.markdown(html_rapat(f'<div class="narasi"><h4>{escape(bagian["Judul"])}</h4><p>{escape(bagian["Teks"]).replace(chr(10), "<br>")}</p></div>'),
                    unsafe_allow_html=True)

    st.subheader("Tabel pendukung")
    t_bulan, t_saluran, t_kategori, t_kota = st.tabs(["Per bulan", "Per saluran", "Per kategori", "Per kota"])
    with t_bulan:
        bulanan = ringkas_bulanan(df)
        tampil_tabel(rapikan_angka(pd.DataFrame({
            "Bulan": bulanan["Bulan"].map(lambda b: label_bulan(b, True)), "Jumlah Kunjungan": bulanan["Kunjungan"],
            "Total Pengunjung": bulanan["Pengunjung"].astype(int),
            "Total Pendapatan (Rp)": bulanan["Pendapatan"].round(0).astype(int),
            "Keterangan": np.where(bulanan["Lengkap"], "Bulan lengkap", "Bulan parsial"),
        })), hide_index=True)
    with t_saluran:
        tampil_tabel(rapikan_angka(tabel_ringkas(df, "Saluran_Pemasaran", "Saluran Pemasaran")), hide_index=True)
    with t_kategori:
        tampil_tabel(rapikan_angka(tabel_ringkas(df, "Kategori_Pengunjung", "Kategori Pengunjung")), hide_index=True)
    with t_kota:
        tampil_tabel(rapikan_angka(tabel_ringkas(df, "Asal_Kota", "Asal Kota")), hide_index=True)

    st.subheader("Unduh laporan")
    rekom = buat_rekomendasi(df)
    d1, d2, _ = st.columns([1, 1, 2])
    d1.download_button("⬇️ Unduh Excel (lengkap)", buat_excel_laporan(df, narasi, rekom, periode, sumber),
                       "laporan_ringkas_bejijong.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    d2.download_button("⬇️ Unduh Teks (.txt)", buat_teks_laporan(narasi, rekom, periode, sumber).encode("utf-8"),
                       "laporan_ringkas_bejijong.txt", mime="text/plain")
    st.caption("Excel berisi sheet Ringkasan, Per_Bulan, Per_Saluran, Per_Kategori, Per_Kota, Rekomendasi, dan Data.")


# ---------- Menu 6: Input Data Pengunjung ----------
def halaman_input() -> None:
    banner("Formulir petugas", "Input Data Pengunjung",
           "Catat kunjungan harian langsung dari sini, tanpa perlu membuka Excel. Data otomatis masuk ke dashboard.")

    with st.form("form_input", clear_on_submit=True):
        st.markdown("##### Data kunjungan")
        c1, c2 = st.columns(2)
        tanggal = c1.date_input("Tanggal kunjungan", value=date.today(), max_value=date.today())
        kategori = c2.selectbox("Kategori pengunjung", PILIHAN_KATEGORI)
        c3, c4 = st.columns(2)
        saluran = c3.selectbox("Saluran pemasaran", PILIHAN_SALURAN, help="Dari mana pengunjung mengetahui desa wisata ini?")
        kota = c4.selectbox("Asal kota", PILIHAN_KOTA)
        c5, c6 = st.columns(2)
        jumlah = c5.number_input("Jumlah pengunjung (orang)", min_value=1, max_value=500, value=2, step=1)
        harga = c6.number_input("Harga tiket per orang (Rp)", min_value=0, max_value=500_000, value=15_000, step=1_000)
        nilai = st.selectbox("Tingkat kepuasan (opsional)", ["Belum dinilai", 5, 4, 3, 2, 1],
                             help="Isi bila pengunjung memberi penilaian 1 sampai 5.")
        st.caption("Total pendapatan dihitung otomatis: jumlah pengunjung × harga tiket per orang.")
        kirim = st.form_submit_button("💾 Simpan data", type="primary")

    if kirim:
        # Validasi konsistensi data sebelum disimpan
        galat = []
        if kategori == "Wisatawan Mancanegara" and kota != "Mancanegara":
            galat.append("Kategori Wisatawan Mancanegara harus berasal dari kota 'Mancanegara'.")
        if kota == "Mancanegara" and kategori != "Wisatawan Mancanegara":
            galat.append("Asal kota 'Mancanegara' hanya cocok untuk kategori Wisatawan Mancanegara.")
        if galat:
            for pesan in galat:
                st.error(pesan)
        else:
            total = int(jumlah) * int(harga)   # perhitungan otomatis Total_Pendapatan
            baris = {
                "Tanggal": tanggal.strftime("%Y-%m-%d"), "Kategori_Pengunjung": kategori, "Saluran_Pemasaran": saluran,
                "Asal_Kota": kota, "Jumlah_Pengunjung": int(jumlah), "Harga_Tiket_Per_Orang": int(harga),
                "Total_Pendapatan": total, "Tingkat_Kepuasan": None if nilai == "Belum dinilai" else float(nilai),
                "Dicatat_Oleh": st.session_state["username"], "Waktu_Input": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            ke_file, info = simpan_input(baris)
            st.success(f"Data berhasil disimpan. {angka_id(jumlah)} pengunjung × {rupiah(harga)} = **{rupiah(total)}**.")
            (st.caption if ke_file else st.warning)(info)

    st.write("")
    st.subheader("Entri terbaru dari petugas")
    riwayat = baca_input_petugas()
    if riwayat.empty:
        st.info("Belum ada data yang diinput. Isi form di atas untuk mencatat kunjungan pertama.")
    else:
        terbaru = riwayat.iloc[::-1].head(10).copy()
        terbaru["Tanggal"] = terbaru["Tanggal"].dt.date
        st.caption(f"{angka_id(len(riwayat))} entri tersimpan. Menampilkan 10 terakhir.")
        tampil_tabel(terbaru, hide_index=True, column_config={
            "Jumlah_Pengunjung": st.column_config.NumberColumn(format="%d"),
            "Harga_Tiket_Per_Orang": st.column_config.NumberColumn(format="Rp %d"),
            "Total_Pendapatan": st.column_config.NumberColumn(format="Rp %d"),
            "Tingkat_Kepuasan": st.column_config.NumberColumn(format="%.1f"),
        })


# ---------- Menu 7: Pengaturan & Logout ----------
def halaman_pengaturan() -> None:
    banner("Akun", "Pengaturan & Logout", "Informasi sesi aktif Anda dan pengelolaan keluar dari aplikasi.")
    mulai: datetime = st.session_state["login_waktu"]
    lama = datetime.now() - mulai
    menit = int(lama.total_seconds() // 60)
    _, sumber, _ = data_dasar()

    st.subheader("Status sesi")
    info = pd.DataFrame({
        "Keterangan": ["Nama pengguna", "Username", "Peran", "Masuk sejak", "Lama sesi", "Sumber data aktif", "Entri input petugas"],
        "Nilai": [st.session_state["nama"], st.session_state["username"], st.session_state["peran"],
                  mulai.strftime("%d %b %Y, %H:%M:%S"), f"{menit} menit", sumber, f"{angka_id(len(baca_input_petugas()))} entri"],
    })
    tampil_tabel(info, hide_index=True)
    st.success("Sesi aktif. Anda masuk sebagai " + st.session_state["peran"] + ".")

    st.write("")
    st.subheader("Keluar")
    st.caption("Keluar menghapus seluruh data sesi, termasuk file yang diunggah pada dashboard.")
    if st.button("🚪 Log Out (Keluar)", type="primary"):
        logout()


# ============================================================================
# 7. MAIN: ROUTER HALAMAN
# ============================================================================
def main() -> None:
    st.session_state.setdefault("logged_in", False)   # status login awal
    st.markdown(CSS_UMUM, unsafe_allow_html=True)

    if not sudah_login():
        halaman_login()
        return

    menu = sidebar_navigasi()
    peta_halaman = {
        MENU["beranda"]: halaman_beranda,
        MENU["dashboard"]: halaman_dashboard,
        MENU["eda"]: halaman_eda,
        MENU["rekomendasi"]: halaman_rekomendasi,
        MENU["laporan"]: halaman_laporan,
        MENU["input"]: halaman_input,
        MENU["pengaturan"]: halaman_pengaturan,
    }
    peta_halaman[menu]()
    navigasi_antar_menu(menu)   # tombol pindah menu di bagian bawah SETIAP halaman


main()
