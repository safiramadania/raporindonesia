# Ruang Sekolah — prototipe lokal

Prototipe portal diagnostik untuk membantu sekolah membaca estimasi LIT dan NUM peserta, menyimpan ringkasan analisis secara lokal, lalu mencatat rencana tindak lanjut.

## Menjalankan di Windows

1. Pastikan Python 3.13 tersedia.
2. Buka PowerShell pada folder ini. Jika paket belum tersedia, buat environment dengan `py -3.13 -m venv .venv`, lalu jalankan `.venv\Scripts\python.exe -m pip install -r requirements.txt`.
3. Jalankan `Jalankan_App.bat`.
4. Buka `http://127.0.0.1:5055` di browser.

## Menyiapkan deploy online

Folder ini berisi konfigurasi Render (`render.yaml`) untuk menjalankan Flask dengan Gunicorn dalam mode publik. Jika dibuat sebagai repo Git terpisah, unggah hanya isi folder aplikasi ini—jangan unggah notebook atau dataset proyek yang berada di folder induk. Pengunjung dapat membuka aplikasi melalui tautan tanpa login.

Prediksi diproses tanpa menyimpan baris peserta di server. Riwayat dan rencana disimpan pada browser masing-masing, sehingga tidak dibagikan ke pengunjung lain. Gunakan untuk demo dan refleksi; aplikasi ini belum memiliki akun sekolah atau sinkronisasi antarperangkat.

## Alur aplikasi

- Isi satu peserta melalui formulir, atau unduh template Excel/CSV dari halaman **Analisis peserta**.
- Isi nomor induk peserta, nama sekolah, dan informasi yang diketahui. Nama kabupaten/kota boleh ditulis sebagai keterangan; versi model saat ini belum menggunakannya untuk menghitung estimasi.
- Kolom yang belum diketahui boleh dikosongkan. Satu baris mewakili satu peserta.
- Unggah CSV atau Excel. Hasil individu diproses di perangkat ini dan tidak ditulis ke database.
- Unduh hasil estimasi, lalu simpan ringkasan saja bila ingin masuk ke riwayat lokal.
- Catat rencana tindak lanjut di halaman **Rencana tindak lanjut**.

Database lokal `rapor_local.sqlite3` hanya menyimpan ringkasan proses dan catatan rencana. Tidak ada autentikasi, akses daring, atau sinkronisasi cloud pada prototipe ini.

## Model dan batasan

Bundle yang tersedia adalah FT-Transformer + PLE-64 dengan tiga seed, dua output (LIT dan NUM), dan 72 prediktor. Skor keluaran dibatasi pada rentang 0–100. Metrik yang tercatat pada bundle: MAE LIT 10.33, RMSE LIT 13.49, MAE NUM 8.62, RMSE NUM 11.14. Ini adalah hasil evaluasi pada eksperimen sebelumnya dan bukan jaminan akurasi untuk semua sekolah.

Bundle ini belum setara dengan varian PLE terbaru yang diuji dengan tambahan kelompok fitur. Jangan menyebut aplikasi ini sebagai implementasi final penelitian sebelum artefak model terbaru, ketersediaan input sekolah, responden, kode kategori, dan alur perlindungan data divalidasi.

Hasil adalah estimasi penelitian, bukan skor resmi AN, bukan diagnosis peserta, dan bukan dasar pemeringkatan atau keputusan berisiko tinggi. Untuk uji coba, gunakan data sintetis atau data yang sudah diizinkan dan diminimalkan.
