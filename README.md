# Sistem Informasi & Interactive Data Analytics Dashboard
Desa Wisata Kampung Majapahit Bejijong · Mitra: LP3SDM Sanggar Bhagaskara · Proyek MSIB D4 Manajemen Informatika UNESA

Aplikasi web berbasis Streamlit dengan 7 menu: Beranda, Analytics Dashboard, Eksplorasi Data (EDA),
Rekomendasi Strategis, Laporan Ringkas, Input Data Pengunjung, dan Pengaturan & Logout.

## 1. Isi folder

```
streamlit_v2/
├── app.py                  aplikasi utama (login, 7 menu, analitik)
├── generate_data.py        pembuat data contoh (sample_desa_wisata.xlsx)
├── sample_desa_wisata.xlsx data contoh 320 baris (hasil generate_data.py)
├── requirements.txt        daftar library
├── README.md               panduan ini
└── .streamlit/
    └── config.toml         tema warna Majapahit (mengunci tampilan terang)
```

Folder `.streamlit` harus berada di samping `app.py`, dengan nama persis seperti itu (diawali titik).

## 2. Menjalankan di komputer (VS Code)

Prasyarat: Python 3.9 atau lebih baru dan VS Code. Cek versi Python dengan `python --version`.

1. Ekstrak zip, lalu di VS Code pilih **File > Open Folder** dan buka folder `streamlit_v2`.
2. Buka terminal VS Code (**Terminal > New Terminal**).
3. Buat lingkungan virtual agar library tidak bercampur dengan proyek lain:
   - Windows: `python -m venv .venv` lalu `.venv\Scripts\activate`
   - macOS/Linux: `python3 -m venv .venv` lalu `source .venv/bin/activate`
   - Bila VS Code menanyakan interpreter, pilih yang berada di `.venv`.
4. Pasang library: `pip install -r requirements.txt`
5. (Opsional) Buat ulang data contoh: `python generate_data.py`
6. Jalankan aplikasi: `streamlit run app.py`

Browser akan terbuka di `http://localhost:8501`. Hentikan aplikasi dengan `Ctrl + C` di terminal.

## 3. Akun demo

| Username | Password      | Peran           |
|----------|---------------|-----------------|
| admin    | bejijong123   | Admin Desa      |
| unesa    | msib2026      | Dosen / Penguji |

Ini akun contoh yang tertulis di kode (`PENGGUNA` di `app.py`). Untuk pemakaian nyata, jangan simpan
kata sandi di kode. Pindahkan ke `st.secrets` atau basis data, dan simpan dalam bentuk hash.

## 4. Alur demo untuk presentasi (sekitar 7 menit)

1. **Login** dengan akun `unesa`. Tunjukkan juga pesan error saat password salah.
2. **Beranda**: jelaskan proyek, mitra, dan panduan penggunaan.
3. **Analytics Dashboard**: tunjukkan 4 KPI, ubah filter sidebar (rentang tanggal, kategori, saluran, kota), buka tiga tab, lalu unduh Excel dan CSV.
4. **Eksplorasi Data (EDA)**: tab Kualitas Data (aturan data lolos), Pola Waktu (bulan puncak berwarna emas), Segmentasi.
5. **Rekomendasi Strategis**: bacakan satu saran prioritas tinggi, lalu mainkan simulasi anggaran (ubah anggaran dan bobot).
6. **Laporan Ringkas**: tunjukkan narasi otomatis dan unduh laporan Excel.
7. **Input Data Pengunjung**: isi satu form, lalu buka Laporan Ringkas lagi untuk menunjukkan angkanya ikut berubah.
8. **Pengaturan & Logout**: akhiri sesi.

Setiap halaman punya tombol "Pindah ke menu lain" di bagian bawah, jadi menu bisa dipindah tanpa sidebar.

## 5. Memakai data nyata desa

1. Siapkan file Excel (.xlsx) atau CSV dengan nama kolom persis berikut (urutan bebas):
   `Tanggal`, `Kategori_Pengunjung`, `Saluran_Pemasaran`, `Asal_Kota`, `Jumlah_Pengunjung`,
   `Harga_Tiket_Per_Orang`, `Total_Pendapatan`, `Tingkat_Kepuasan`.
   Bila `Total_Pendapatan` tidak ada, aplikasi menghitungnya dari jumlah x harga.
2. Cara cepat: unggah file lewat sidebar **Analytics Dashboard**. Data langsung dipakai semua menu.
3. Cara tetap: ganti file `sample_desa_wisata.xlsx` dengan file data nyata (nama dan kolom sama).
4. Buka menu **Eksplorasi Data > Kualitas Data** untuk melihat baris yang melanggar aturan, misalnya salah ketik nama saluran.
5. Bila nama kategori, saluran, atau kota di data desa berbeda dari daftar baku, ubah daftar `PILIHAN_KATEGORI`, `PILIHAN_SALURAN`, dan `PILIHAN_KOTA` di bagian atas `app.py`.

Data yang diinput petugas lewat form tersimpan di `data_input_petugas.csv` (dibuat otomatis di samping `app.py`).

## 6. Menayangkan online (Streamlit Community Cloud, gratis)

1. Buat repositori GitHub, lalu unggah `app.py`, `requirements.txt`, `sample_desa_wisata.xlsx`, dan folder `.streamlit`.
2. Buka `https://share.streamlit.io`, masuk dengan akun GitHub, pilih **Create app**.
3. Pilih repositori dan cabang, isi **Main file path** dengan `app.py`, lalu **Deploy**.
4. Setelah beberapa menit aplikasi mendapat alamat web yang bisa dibagikan.

Catatan penting sebelum dipakai petugas desa:
- Penyimpanan file di layanan gratis bersifat sementara. `data_input_petugas.csv` bisa hilang saat aplikasi restart. Untuk pemakaian nyata, simpan data ke basis data atau Google Sheets.
- Ganti kredensial demo dulu (lihat bagian 3). Repositori publik membuat kata sandi di kode terbaca siapa saja.

## 7. Pemecahan masalah

| Gejala | Penyelesaian |
|--------|--------------|
| `streamlit` is not recognized | Jalankan `python -m streamlit run app.py`, atau aktifkan lagi `.venv` |
| `ModuleNotFoundError` | Jalankan `pip install -r requirements.txt` di terminal yang `.venv`-nya aktif |
| Warna tema tidak muncul | Pastikan `.streamlit/config.toml` berada di samping `app.py`, lalu jalankan ulang |
| Port 8501 dipakai | `streamlit run app.py --server.port 8502` |
| File unggahan ditolak | Cek nama kolom wajib (bagian 5). Pesan galat tampil di sidebar |
| Data contoh hilang | Jalankan `python generate_data.py` |

## 8. Kaitan dengan CRISP-DM

| Fase | Wujud di aplikasi |
|------|-------------------|
| Business Understanding | Beranda dan proposal MSIB |
| Data Understanding | Menu Eksplorasi Data: Kualitas Data, Sebaran |
| Data Preparation | Fungsi `bersihkan()` dan validasi form Input Data |
| Modeling | Mesin rekomendasi berbasis aturan dan simulasi anggaran |
| Evaluation | Pembacaan hasil bersama mitra dengan data nyata (tahap lanjutan) |
| Deployment | Aplikasi Streamlit dan Laporan Ringkas yang bisa diunduh |

## 9. Batasan

- Data contoh adalah simulasi. Rekomendasi baru bermakna dengan data nyata mitra.
- Belum ada data okupansi homestay dan paket wisata, jadi saran kapasitas memakai beban kunjungan harian.
- Pola yang terlihat adalah korelasi, bukan bukti sebab-akibat.
