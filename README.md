# Aplikasi Web Kuesioner Analisis Kebutuhan Pembelajaran - UNJ

Aplikasi web **Streamlit** untuk pengisian kuesioner penelitian disertasi **Analisis Kebutuhan Pembelajaran** oleh **Ruslina Irianty** (Program Doktor S3 Penelitian dan Evaluasi Pendidikan, Universitas Negeri Jakarta).

---

## 🎯 Fitur Utama:
1. **2 Jenis Kuesioner**: Kuesioner Dosen dan Kuesioner Mahasiswa dalam 1 aplikasi.
2. **Informed Consent**: Form persetujuan wajib diisi sebelum melanjutkan.
3. **Pengantar Penelitian**: Menampilkan sambutan resmi dan jaminan kerahasiaan data.
4. **Biodata Terpisah**: Form identitas berbeda untuk Dosen dan Mahasiswa.
5. **Navigasi Bagian**: Pertanyaan utama dan panduan per bagian dengan progress bar.
6. **Auto-Sync Google Sheets**: Hasil pengisian tersimpan otomatis via Google Apps Script.
7. **Simpan Excel Lokal**: Data tersimpan di folder `responses/` secara otomatis.

---

## 📁 Struktur File:
- `app_survey.py`: Kode utama aplikasi Streamlit.
- `questions_analisis_kebutuhan.json`: Bank pertanyaan kuesioner Dosen & Mahasiswa.
- `Google_Apps_Script.gs`: Kode Apps Script untuk dipasang di Google Sheets.
- `requirements.txt`: Dependensi Python.
- `.streamlit/secrets.toml`: Konfigurasi URL Google Apps Script (jangan di-upload ke GitHub).

---

## 🚀 Cara Menjalankan Aplikasi:

### 1. Di Komputer Lokal:
```bash
pip install -r requirements.txt
streamlit run app_survey.py
```

### 2. Deploy ke Streamlit Cloud:
1. Upload 3 file ke GitHub: `app_survey.py`, `questions_analisis_kebutuhan.json`, `requirements.txt`
2. Buka [share.streamlit.io](https://share.streamlit.io) dan hubungkan dengan repository.
3. Di dasbor Streamlit Cloud, buka **App Settings > Secrets**, lalu salin konfigurasi berikut:

```toml
[connections.gsheets]
spreadsheet = "https://script.google.com/macros/s/AKfycbyT_n9dXE9zWLXIMUln7pPikWYjX4wB2T-IDmTT_5ftNPhx8QjoXQ9KRy1fU0WYYsqx/exec"
```
