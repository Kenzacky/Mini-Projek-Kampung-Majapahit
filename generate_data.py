"""
generate_data.py
================
Membuat dataset DUMMY kunjungan wisata Kampung Majapahit Bejijong
dan menyimpannya sebagai `sample_desa_wisata.xlsx`.

Cara menjalankan:
    pip install pandas numpy openpyxl
    python generate_data.py

Catatan:
- Seluruh data adalah SIMULASI untuk keperluan pengembangan dashboard.
- Ganti dengan data riil dari LP3SDM Sanggar Bhagaskara setelah tersedia.
  Selama nama kolomnya sama, `app.py` bisa langsung membacanya.
"""

from datetime import date, timedelta

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# 1. KONFIGURASI
# ----------------------------------------------------------------------------
SEED = 2026                      # seed tetap agar hasilnya bisa diulang
JUMLAH_BARIS = 320               # minimal 250 baris sesuai kebutuhan
RENTANG_BULAN = 12               # data 12 bulan terakhir (boleh 6-12)
NAMA_FILE = "sample_desa_wisata.xlsx"

KATEGORI = [
    "Wisatawan Lokal",
    "Wisatawan Mancanegara",
    "Rombongan Edukasi/Sekolah",
    "Peneliti/Akademisi",
]
SALURAN = [
    "Instagram",
    "TikTok",
    "Website Resmi",
    "Walk-in / Langsung",
    "Brosur / Event",
    "Travel Agent",
]
KOTA = [
    "Surabaya",
    "Mojokerto",
    "Malang",
    "Sidoarjo",
    "Jakarta",
    "Yogyakarta",
    "Mancanegara",
]

# Peluang tiap kategori pengunjung
BOBOT_KATEGORI = [0.52, 0.12, 0.26, 0.10]

# Peluang saluran pemasaran untuk tiap kategori (urutan = SALURAN)
BOBOT_SALURAN = {
    "Wisatawan Lokal":           [0.28, 0.22, 0.08, 0.22, 0.12, 0.08],
    "Wisatawan Mancanegara":     [0.18, 0.04, 0.26, 0.04, 0.06, 0.42],
    "Rombongan Edukasi/Sekolah": [0.04, 0.02, 0.10, 0.12, 0.22, 0.50],
    "Peneliti/Akademisi":        [0.10, 0.02, 0.50, 0.12, 0.16, 0.10],
}

# Peluang kota asal untuk kategori domestik (urutan = KOTA tanpa "Mancanegara")
KOTA_DOMESTIK = KOTA[:-1]
BOBOT_KOTA = [0.24, 0.26, 0.16, 0.12, 0.12, 0.10]

# Rentang jumlah rombongan (min, maks) per kategori. Semuanya tetap dalam 2-50.
RENTANG_JUMLAH = {
    "Wisatawan Lokal":           (2, 12),
    "Wisatawan Mancanegara":     (2, 10),
    "Rombongan Edukasi/Sekolah": (20, 50),
    "Peneliti/Akademisi":        (2, 15),
}

# Rentang harga tiket per orang (min, maks) per kategori. Semuanya dalam Rp 10.000-25.000.
RENTANG_HARGA = {
    "Wisatawan Lokal":           (10_000, 15_000),
    "Wisatawan Mancanegara":     (20_000, 25_000),
    "Rombongan Edukasi/Sekolah": (10_000, 12_000),
    "Peneliti/Akademisi":        (12_000, 18_000),
}

# Rata-rata kepuasan dasar per kategori (sebelum diberi noise)
KEPUASAN_DASAR = {
    "Wisatawan Lokal": 4.3,
    "Wisatawan Mancanegara": 4.5,
    "Rombongan Edukasi/Sekolah": 4.1,
    "Peneliti/Akademisi": 4.4,
}


# ----------------------------------------------------------------------------
# 2. FUNGSI PEMBANTU
# ----------------------------------------------------------------------------
def buat_tanggal(rng: np.random.Generator, n: int) -> pd.Series:
    """
    Membuat n tanggal acak dalam RENTANG_BULAN terakhir.
    Akhir pekan dan bulan libur sekolah (Jun-Jul, Des) dibuat lebih ramai
    supaya pola musimnya terlihat di grafik tren.
    """
    akhir = date.today()
    awal = akhir - timedelta(days=30 * RENTANG_BULAN)
    semua_hari = pd.date_range(awal, akhir, freq="D")

    bobot = np.ones(len(semua_hari))
    # Akhir pekan (Sabtu=5, Minggu=6) lebih ramai
    bobot[semua_hari.dayofweek >= 5] *= 2.2
    # Hari Jumat sedikit lebih ramai
    bobot[semua_hari.dayofweek == 4] *= 1.3
    # Musim libur sekolah dan akhir tahun
    bobot[np.isin(semua_hari.month, [6, 7])] *= 1.5
    bobot[semua_hari.month == 12] *= 1.6
    bobot = bobot / bobot.sum()

    pilihan = rng.choice(len(semua_hari), size=n, p=bobot)
    return pd.Series(semua_hari[pilihan]).dt.normalize()


def buat_dataframe(n: int = JUMLAH_BARIS, seed: int = SEED) -> pd.DataFrame:
    """Membangun DataFrame dummy dengan struktur kolom sesuai kebutuhan dashboard."""
    rng = np.random.default_rng(seed)

    tanggal = buat_tanggal(rng, n)
    kategori = rng.choice(KATEGORI, size=n, p=BOBOT_KATEGORI)

    saluran, kota, jumlah, harga, kepuasan = [], [], [], [], []
    for kat in kategori:
        # Saluran pemasaran bergantung pada kategori pengunjung
        saluran.append(rng.choice(SALURAN, p=BOBOT_SALURAN[kat]))

        # Wisatawan mancanegara selalu berasal dari "Mancanegara"
        if kat == "Wisatawan Mancanegara":
            kota.append("Mancanegara")
        else:
            kota.append(rng.choice(KOTA_DOMESTIK, p=BOBOT_KOTA))

        lo, hi = RENTANG_JUMLAH[kat]
        jumlah.append(int(rng.integers(lo, hi + 1)))

        # Harga dibulatkan ke kelipatan Rp 1.000
        hlo, hhi = RENTANG_HARGA[kat]
        harga.append(int(rng.integers(hlo // 1000, hhi // 1000 + 1) * 1000))

        # Kepuasan condong tinggi dengan sedikit variasi, dibatasi 1.0-5.0
        nilai = rng.normal(KEPUASAN_DASAR[kat], 0.6)
        kepuasan.append(round(float(np.clip(nilai, 1.0, 5.0)), 1))

    df = pd.DataFrame(
        {
            "Tanggal": tanggal,
            "Kategori_Pengunjung": kategori,
            "Saluran_Pemasaran": saluran,
            "Asal_Kota": kota,
            "Jumlah_Pengunjung": jumlah,
            "Harga_Tiket_Per_Orang": harga,
        }
    )
    # Total pendapatan = jumlah pengunjung x harga tiket per orang
    df["Total_Pendapatan"] = df["Jumlah_Pengunjung"] * df["Harga_Tiket_Per_Orang"]
    df["Tingkat_Kepuasan"] = kepuasan

    return df.sort_values("Tanggal").reset_index(drop=True)


def simpan_excel(df: pd.DataFrame, nama_file: str = NAMA_FILE) -> None:
    """Menyimpan DataFrame ke Excel dengan header berwarna dan format kolom rapi."""
    with pd.ExcelWriter(nama_file, engine="openpyxl", datetime_format="YYYY-MM-DD") as penulis:
        df.to_excel(penulis, sheet_name="Data_Kunjungan", index=False)
        ws = penulis.sheets["Data_Kunjungan"]

        # Gaya header: latar bata merah, teks putih tebal
        for sel in ws[1]:
            sel.font = Font(bold=True, color="FFFFFF")
            sel.fill = PatternFill("solid", fgColor="8B3A2B")
            sel.alignment = Alignment(horizontal="center", vertical="center")

        # Lebar kolom otomatis mengikuti isi terpanjang
        for i, kolom in enumerate(df.columns, start=1):
            terpanjang = max(len(str(kolom)), df[kolom].astype(str).str.len().max())
            ws.column_dimensions[get_column_letter(i)].width = min(terpanjang + 4, 34)

        # Format angka: pemisah ribuan untuk harga dan pendapatan, satu desimal untuk rating
        for baris in ws.iter_rows(min_row=2):
            baris[4].number_format = "#,##0"       # Jumlah_Pengunjung
            baris[5].number_format = "#,##0"       # Harga_Tiket_Per_Orang
            baris[6].number_format = "#,##0"       # Total_Pendapatan
            baris[7].number_format = "0.0"         # Tingkat_Kepuasan

        ws.freeze_panes = "A2"                     # header tetap terlihat saat scroll
        ws.auto_filter.ref = ws.dimensions         # tombol filter di setiap kolom


# ----------------------------------------------------------------------------
# 3. PROGRAM UTAMA
# ----------------------------------------------------------------------------
if __name__ == "__main__":
    data = buat_dataframe()
    simpan_excel(data)
    print(f"Berhasil membuat '{NAMA_FILE}' dengan {len(data)} baris.")
    print(f"Rentang tanggal : {data['Tanggal'].min():%Y-%m-%d} s.d. {data['Tanggal'].max():%Y-%m-%d}")
    print(f"Total pengunjung: {data['Jumlah_Pengunjung'].sum():,}")
    print(f"Total pendapatan: Rp {data['Total_Pendapatan'].sum():,}")
