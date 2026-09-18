# KPJ PDF Auto Renamer

Aplikasi Streamlit untuk mencocokkan nomor KPJ 11 digit di PDF dengan Excel, lalu langsung menulis PDF ke nama akhir dalam ZIP hasil. Isi PDF tidak diubah.

## Jalankan lokal

Gunakan Python 3.11 atau 3.12. Dari folder project:

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
python -m pip install -r requirements.txt
streamlit run app.py
```

## Deploy ke Streamlit Community Cloud

1. Upload isi folder project ke repository GitHub (termasuk `.streamlit/config.toml`). Jangan upload data pegawai.
2. Masuk ke Streamlit Community Cloud, pilih **Create app**, repository dan branch tersebut.
3. Isi main file path dengan `app.py`, pilih Python 3.12 di pengaturan lanjutan, lalu deploy. Tidak memerlukan secrets atau API key.

Referensi: [deployment resmi](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy) dan [dependensi](https://docs.streamlit.io/deploy/concepts/dependencies).

## Penggunaan

Upload ZIP dan Excel `.xlsx`/`.xls`, pilih sheet, periksa preview dan pilihan kolom, lalu klik **Proses File**. Header berada di baris pertama. Header dikenali tanpa membedakan kapitalisasi, spasi atau tanda baca; pilihan dapat diubah. `.xls` dibaca menggunakan xlrd, `.xlsx` menggunakan openpyxl melalui pandas.

Contoh: `NO KPJ` = `26133068531`, `NAMA BARU` = `PEGAWAI_PEG26072942_KARTU BENEFIT_KARTU BPJS TK.pdf`.

KPJ disimpan sebagai string. Spasi, tanda kutip di luar angka, dan akhiran `.0` dibersihkan. Leading zero pada **nilai teks Excel** dipertahankan. Simpan kolom KPJ sebagai Text: nol yang sudah hilang karena Excel menyimpan nilai numerik tidak bisa dipulihkan dengan pasti. Format tampilan `00000000000` saja bukan nilai teks; nilai kurang dari 11 digit ditolak dan dicatat, tidak ditebak.

## Aturan matching dan duplikat

- PyMuPDF membaca seluruh halaman. Jika gagal atau tidak ada kandidat, pdfminer.six menjadi fallback. Regex `(?<!\d)[0-9]{11}(?!\d)` menolak potongan angka yang lebih panjang.
- Satu kandidat yang cocok di Excel digunakan. Bila beberapa kandidat cocok, aplikasi mempertahankan nama asli di `Tidak_Ditemukan/` dan menandai `DUPLIKAT` agar tidak memilih pegawai secara sembarang. Angka 11 digit tanpa label tetap dapat menjadi kandidat, jadi periksa laporan.
- KPJ Excel berulang dengan nama sama memakai satu mapping dan dicatat. KPJ sama dengan nama berbeda tidak di-rename sampai Excel diperbaiki.
- Dua PDF dengan KPJ sama tetap disimpan. Benturan nama, termasuk perbedaan kapitalisasi saja, diberi `_2`, `_3`, dst. Nilai `Duplikat` adalah jumlah PDF terdampak; duplikat Excel juga dirinci pada sheet validasi.
- Nama dibersihkan dari path, karakter terlarang, nama perangkat Windows dan panjang berlebihan. Nama asli hanya disesuaikan jika tidak aman atau bertabrakan.

ZIP hasil berisi `Hasil_Rename/`, `Tidak_Ditemukan/`, `Gagal_Membaca_KPJ/`, dan `laporan_mapping.xlsx`. Laporan juga tersedia sebagai download terpisah, dengan sheet hasil, validasi Excel, dan ringkasan. PDF rusak/password tetap disertakan di folder gagal dengan status `ERROR`. ZIP rusak/terenkripsi yang tidak dapat dibaca membatalkan batch; hasil parsial tidak diterbitkan. File non-PDF dalam ZIP diabaikan.

## Privasi dan batas pemrosesan

Input, hasil, dan laporan diproses di memori, tanpa menyimpan upload permanen, tanpa cache lintas pengguna, tanpa mengekstrak path ZIP ke disk. Hasil berada dalam session_state agar kedua download dapat digunakan. Tombol **Bersihkan sesi** menghapus referensi input dan hasil dari sesi aplikasi. Dokumen dikirim ke server tempat aplikasi dijalankan; atur akses deployment sesuai kebutuhan organisasi.

Batas: upload 100 MB/file, maksimal 2.000 PDF, 30 MB/PDF, total PDF 250 MB setelah dekompresi. Pemrosesan serta laporan membutuhkan memori tambahan; pecah batch jika resource server terbatas. Tidak ada hasil jika seluruh mapping Excel tidak valid. KPJ ambigu bisa menyebabkan jumlah kategori ringkasan tidak menjumlah tepat ke total; semua detail ada dalam tabel.

## OCR opsional

OCR nonaktif secara default dan bukan dependensi wajib. Hook `ocr_pdf()` tersedia di `utils.py`. Untuk mengaktifkan, install `pytesseract` dan `Pillow`, install Tesseract OS (Community Cloud: tambahkan `packages.txt` berisi `tesseract-ocr`), lalu panggil `read_pdf_kpjs(payload, enable_ocr=True)` pada `process_zip`. Verifikasi hasil OCR karena digit dapat salah dikenali. PDF scan tanpa OCR masuk folder gagal.

## Pengujian

```bash
python -m unittest discover -s tests -v
```

Pengujian membuat PDF/Excel sintetis di memori; mencakup matching, fallback, leading zero, duplikat, konflik, sanitasi, byte PDF tidak berubah, ZIP rusak, dan laporan Excel. Tidak membutuhkan data pegawai.

Status verifikasi saat penyerahan: 8 pengujian pemrosesan lulus dengan Python 3.12, PyMuPDF 1.24.14, pandas 3.0.1, openpyxl 3.1.5, dan pdfminer.six 20251230 yang tersedia di lingkungan pengujian. Pemeriksaan sintaks juga lulus. Pengujian UI `tests/test_app.py` disediakan tetapi belum dapat dijalankan karena keterbatasan disk/dependensi lingkungan. Instalasi bersih seluruh requirements, pembacaan fixture `.xls`, tampilan browser, dan deployment Cloud belum diverifikasi langsung. Requirements menggunakan rentang versi untuk resolusi dependensi Cloud, dengan pandas 2.x untuk kompatibilitas Streamlit lintas versi.
