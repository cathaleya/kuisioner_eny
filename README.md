# Aplikasi Web Kuesioner Evaluasi LDP (Need Assessment) - UNJ

Aplikasi web terpadu **Streamlit** untuk pengisian 5 instrumen kuesioner kualitatif penelitian disertasi pengembangan **Model Evaluasi Leadership Development Program (LDP)** oleh Eny Cahyaningsih (Penelitian dan Evaluasi Pendidikan, Universitas Negeri Jakarta).

---

## 🎯 Fitur Utama:
1. **Unified Application**: 5 instrumen kuesioner disajikan dalam 1 aplikasi web tanpa perlu memecah menjadi aplikasi terpisah.
2. **Pengantar Penelitian**: Menampilkan sambutan resmi, tujuan penelitian, dan jaminan kerahasiaan data sesuai dokumen `Pengantar Kuesioner Penelitian.docx`.
3. **Pilihan Kuesioner Interaktif**:
   - 1. Kuesioner Peserta LDP
   - 2. Kuesioner Atasan / Mentor
   - 3. Kuesioner Pengelola LDP / Fasilitator
   - 4. Kuesioner Manajemen SDM / Talent Management
   - 5. Kuesioner Pimpinan Unit Kerja
4. **Navigasi Modul & Indikator**: Menampilkan Pertanyaan Utama dan Pertanyaan Mendalam secara sistematis dengan *progress bar*.
5. **Auto-Sync Google Sheets**: Hasil pengisian tersimpan otomatis ke Google Sheets target via Google Apps Script:
   `https://docs.google.com/spreadsheets/d/1171QQzfhf--vDczUZm-Ty71L38g6RjGHZoAoU_EsvLQ/edit?usp=sharing`
6. **Animasi & Ucapan Terima Kasih**: Efek `st.balloons()` dan banner ucapan terima kasih pada halaman perayaan akhir.

---

## 📁 Struktur File:
- `app_survey.py`: Kode utama aplikasi Streamlit.
- `questions.json`: Bank pertanyaan lengkap hasil konversi 5 dokumen kuesioner & pengantar penelitian.
- `Google_Apps_Script.gs`: Kode Apps Script untuk dipasang di Google Sheets target.
- `requirements.txt`: Dependensi Python (`streamlit`, `pandas`, `requests`).
- `.streamlit/secrets.toml`: Konfigurasi URL Google Apps Script Web App.

---

## 🚀 Cara Menjalankan Aplikasi:

### 1. Di Komputer Lokal:
```bash
pip install -r requirements.txt
streamlit run app_survey.py
```

### 2. Deploy ke Streamlit Cloud:
1. Push/unggah folder ini ke repository GitHub.
2. Buka [share.streamlit.io](https://share.streamlit.io) dan hubungkan dengan repository.
3. Di **Advanced Settings > Secrets**, tambahkan konfigurasi Apps Script:
```toml
[connections.gsheets]
spreadsheet = "https://script.google.com/macros/s/YOUR_APPS_SCRIPT_DEPLOYMENT_ID/exec"
```
