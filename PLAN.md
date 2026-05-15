# Streamlit Web App Analisis Terpadu Posko

## Summary
Bangun web app lokal berbasis Streamlit untuk menggantikan alur manual `analisis_terpadu_posko.py`: user upload 1 file Excel yang sudah berisi data gangguan dan mapping posko, mengisi jumlah posko existing, app menjalankan analisis, lalu menampilkan dashboard lengkap serta menyediakan download Excel/PNG hasil analisis.

## Key Changes
- Refactor `analisis_terpadu_posko.py` agar punya fungsi reusable seperti `run_analysis(gangguan_file, gangguan_sheet, output_dir, jumlah_posko_existing, progress_callback)`.
- Pertahankan mode CLI lama dengan `if __name__ == "__main__": main()`, jadi script tetap bisa dijalankan langsung.
- Buat `app.py` Streamlit dengan:
  - upload satu file Excel,
  - dropdown sheet yang dibaca dari workbook upload, dengan default `data_gangguan_valid_14052026` jika tersedia,
  - input jumlah posko existing,
  - tombol download template Excel dan contoh data 100 baris,
  - progress/status proses dari 12 tahap analisis,
  - dashboard tab untuk ringkasan, grafik, map OpenStreetMap, posko existing/usulan, beban kerja, response time, detail data, outlier, dan data tidak terpetakan,
  - download output Excel lengkap dan semua gambar hasil analisis.
- Gunakan output per-run di folder terpisah, misalnya `output/web_runs/<timestamp>/`, supaya hasil user tidak menimpa file lama.
- Tambahkan `requirements.txt` berisi minimal:
  - `streamlit`
  - `pandas`
  - `numpy`
  - `matplotlib`
  - `scikit-learn`
  - `openpyxl`

## Input/Template
- Template upload berisi header minimal:
  `NO_LAPORAN`, `IDPEL`, `LAT_PLGN`, `LONG_PLGN`, `TGL_LAPOR`, `BULAN`, `JAM_LAPOR`, `RPT`, `NAMA_REGU`, `NAMA_POSKO`.
- Contoh data 100 baris tersedia sebagai `contoh_data_100_baris.xlsx`.
- Range silhouette dihitung dari input jumlah posko existing:
  - `K_MIN = max(2, jumlah_posko_existing - 10)`
  - `K_MAX = jumlah_posko_existing`
- Validasi wajib sebelum proses:
  - file punya `NO_LAPORAN`, `IDPEL`, `LAT_PLGN`, `LONG_PLGN`, `TGL_LAPOR`, `BULAN`, `JAM_LAPOR`, `RPT`, `NAMA_REGU`, `NAMA_POSKO`,
  - sheet dipilih dari daftar sheet workbook upload,
  - koordinat bisa dikonversi ke angka.

## Test Plan
- Jalankan `python analisis_terpadu_posko.py` untuk memastikan CLI lama tetap bekerja.
- Jalankan `streamlit run app.py`, upload `contoh_data_100_baris.xlsx`, isi jumlah posko existing, lalu pastikan:
  - progress berjalan sampai selesai,
  - output Excel bisa di-download,
  - grafik silhouette, grafik beban, dan visualisasi KMeans tampil,
  - tabel dashboard tampil dari semua sheet output utama,
  - hasil disimpan di folder run baru tanpa menimpa `hasil_analisis_terpadu_posko.xlsx`.
- Uji error handling:
  - sheet salah,
  - kolom wajib hilang,
  - file bukan Excel,
  - data koordinat kosong/tidak valid.

## Assumptions
- App hanya ditargetkan untuk pemakaian lokal via `streamlit run app.py`.
- User upload 1 workbook yang sudah dianggap benar dan berisi kolom mapping posko.
- Dashboard lengkap berarti semua output utama ditampilkan di UI, termasuk map OpenStreetMap, tetapi Excel tetap menjadi artifact lengkap untuk arsip.
- Streamlit belum terpasang di environment saat ini, jadi perlu install dari `requirements.txt`.
