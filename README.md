# Analisis Terpadu Posko Web App

Web app lokal untuk menjalankan analisis terpadu posko dari satu file Excel yang sudah berisi kolom data gangguan dan mapping posko:

- `NO_LAPORAN`
- `IDPEL`
- `LAT_PLGN`
- `LONG_PLGN`
- `TGL_LAPOR`
- `BULAN`
- `JAM_LAPOR`
- `RPT`
- `NAMA_REGU`
- `NAMA_POSKO`

Output setiap run disimpan di `webapp/output/web_runs/<timestamp>/`.

## Install

```powershell
cd "D:\OneDrive - PLN\1 KULIAH\SKRIPSI\PENGOLAHAN DATA\webapp"
pip install -r requirements.txt
```

## Jalankan

```powershell
streamlit run app.py
```

Atau dari folder project utama:

```powershell
streamlit run webapp/app.py
```

## Pilih Sheet

- Aplikasi membaca daftar sheet dari file upload.
- Jika ada sheet `data_gangguan_valid_14052026`, sheet itu dipilih sebagai default.

Template Excel dan contoh data 100 baris bisa diunduh dari sidebar aplikasi.

## Parameter Posko

User perlu mengisi jumlah posko existing. Range silhouette otomatis dihitung:

- `K_MIN = max(2, jumlah posko existing - 10)`
- `K_MAX = jumlah posko existing`

## Output

Setelah proses selesai, dashboard menampilkan ringkasan, response time, beban kerja, posko, detail perbandingan, outlier, data tidak terpetakan, map OpenStreetMap, grafik, dan visualisasi KMeans.

File Excel lengkap dan semua gambar juga tersedia sebagai download.
Map juga tersedia sebagai file HTML.
