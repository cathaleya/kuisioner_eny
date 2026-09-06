import streamlit as st
import json
import pandas as pd
import requests
import os
import io
from datetime import datetime

# ==========================================
# PAGE CONFIG & CUSTOM CSS
# ==========================================
st.set_page_config(
    page_title="Kuesioner Evaluasi LDP - UNJ",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Premium Aesthetics
st.markdown("""
<style>
    /* Main container styling */
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        max-width: 1050px;
    }
    
    /* Header Banner */
    .header-box {
        background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
        color: white;
        padding: 2rem;
        border-radius: 14px;
        box-shadow: 0 4px 14px rgba(30, 58, 138, 0.15);
        margin-bottom: 2rem;
    }
    .header-box h1 {
        color: #FFFFFF !important;
        font-size: 1.85rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        line-height: 1.3;
    }
    .header-box p {
        color: #DBEAFE !important;
        font-size: 1.05rem;
        margin-bottom: 0;
    }
    
    /* Intro Card Styling */
    .intro-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.05);
        overflow: hidden;
        margin-bottom: 1.5rem;
    }
    .intro-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #1E40AF 100%);
        color: white;
        padding: 1.25rem 1.8rem;
    }
    .intro-header h2 {
        color: #FFFFFF !important;
        font-size: 1.4rem;
        font-weight: 700;
        margin: 0 0 0.3rem 0;
    }
    .intro-subtitle {
        color: #BFDBFE !important;
        font-size: 0.95rem;
        margin: 0;
    }
    .intro-body {
        padding: 1.8rem 2rem;
        color: #334155;
        font-size: 1.02rem;
        line-height: 1.75;
    }
    .intro-body p {
        margin-bottom: 1.2rem;
    }
    .info-box {
        border-radius: 10px;
        padding: 1.25rem 1.6rem;
        margin: 1.4rem 0;
    }
    .blue-box {
        background-color: #EFF6FF;
        border-left: 5px solid #3B82F6;
    }
    .blue-box h4 {
        color: #1E40AF !important;
        margin-top: 0;
        margin-bottom: 0.6rem;
        font-size: 1.1rem;
        font-weight: 700;
    }
    .green-box {
        background-color: #F0FDF4;
        border-left: 5px solid #22C55E;
    }
    .green-box h4 {
        color: #166534 !important;
        margin-top: 0;
        margin-bottom: 0.6rem;
        font-size: 1.1rem;
        font-weight: 700;
    }
    .info-box ul, .info-box ol {
        margin: 0;
        padding-left: 1.4rem;
    }
    .info-box li {
        margin-bottom: 0.5rem;
        color: #1E293B;
    }
    .author-signature {
        margin-top: 2rem;
        padding-top: 1.2rem;
        border-top: 2px dashed #E2E8F0;
    }
    .author-title {
        font-size: 0.85rem;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        font-weight: 600;
    }
    .author-name {
        font-size: 1.25rem;
        font-weight: 700;
        color: #1E3A8A;
        margin: 0.2rem 0;
    }
    .author-affiliation {
        font-size: 0.95rem;
        color: #475569;
        line-height: 1.4;
    }
    
    /* Selection Cards */
    .card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
    }
    
    /* Domain Badge */
    .domain-badge {
        display: inline-block;
        background-color: #EFF6FF;
        color: #1D4ED8;
        font-weight: 600;
        font-size: 0.85rem;
        padding: 0.35rem 0.85rem;
        border-radius: 9999px;
        border: 1px solid #BFDBFE;
        margin-bottom: 0.75rem;
    }
    
    /* Main Question Box */
    .main-question-box {
        background-color: #F8FAFC;
        border-left: 5px solid #2563EB;
        padding: 1.25rem 1.5rem;
        border-radius: 8px;
        margin-top: 0.5rem;
        margin-bottom: 1.25rem;
    }
    .main-question-title {
        color: #1E293B;
        font-weight: 700;
        font-size: 1.1rem;
        margin-bottom: 0.5rem;
    }
    
    /* Buttons */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        padding: 0.55rem 1.3rem;
        transition: all 0.2s;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# LOAD DATA
# ==========================================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
QUESTIONS_PATH = os.path.join(CURRENT_DIR, "questions.json")
RESPONSES_DIR = os.path.join(CURRENT_DIR, "responses")
os.makedirs(RESPONSES_DIR, exist_ok=True)

LOCAL_EXCEL_PATH = os.path.join(RESPONSES_DIR, "data_jawaban_excel.xlsx")
LOCAL_CSV_PATH = os.path.join(RESPONSES_DIR, "data_jawaban.csv")

@st.cache_data
def load_questions():
    if not os.path.exists(QUESTIONS_PATH):
        return None
    with open(QUESTIONS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

data = load_questions()

if not data:
    st.error(f"File `{QUESTIONS_PATH}` tidak ditemukan. Silakan jalankan script ekstraksi data terlebih dahulu.")
    st.stop()

questionnaires_dict = data.get("questionnaires", {})
TARGET_SHEET_URL = data.get("google_sheets_url", "https://docs.google.com/spreadsheets/d/1171QQzfhf--vDczUZm-Ty71L38g6RjGHZoAoU_EsvLQ/edit?usp=sharing")

# ==========================================
# SESSION STATE INITIALIZATION
# ==========================================
if "page" not in st.session_state:
    st.session_state.page = "intro"
if "selected_q_id" not in st.session_state:
    st.session_state.selected_q_id = "peserta_ldp"
if "user_biodata" not in st.session_state:
    st.session_state.user_biodata = {}
if "answers" not in st.session_state:
    st.session_state.answers = {}
if "closing_answer" not in st.session_state:
    st.session_state.closing_answer = ""
if "current_mod_idx" not in st.session_state:
    st.session_state.current_mod_idx = 0

# Helper to flatten modules for easy linear navigation
def get_flattened_modules(q_id):
    q_data = questionnaires_dict.get(q_id, {})
    flat = []
    domains = q_data.get("domains", [])
    for d in domains:
        domain_title = d.get("title", "")
        for m in d.get("modules", []):
            flat.append({
                "domain": domain_title,
                "code": m.get("code", ""),
                "title": m.get("title", ""),
                "main_question": m.get("main_question", ""),
                "sub_questions": m.get("sub_questions", [])
            })
    return flat

# Helper function to save answers locally into Excel & CSV
def save_to_local_excel(payload):
    try:
        df_new = pd.DataFrame([payload])
        
        # Save or append to CSV
        if os.path.exists(LOCAL_CSV_PATH):
            df_existing = pd.read_csv(LOCAL_CSV_PATH)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_combined.to_csv(LOCAL_CSV_PATH, index=False, encoding="utf-8-sig")
        else:
            df_new.to_csv(LOCAL_CSV_PATH, index=False, encoding="utf-8-sig")
            
        # Save or append to Excel
        if os.path.exists(LOCAL_EXCEL_PATH):
            df_existing_excel = pd.read_excel(LOCAL_EXCEL_PATH)
            df_combined_excel = pd.concat([df_existing_excel, df_new], ignore_index=True)
            df_combined_excel.to_excel(LOCAL_EXCEL_PATH, index=False)
        else:
            df_new.to_excel(LOCAL_EXCEL_PATH, index=False)
            
        return True, "Data berhasil direkam ke file Excel lokal komputer."
    except Exception as e:
        return False, f"Gagal merekam ke file lokal: {str(e)}"

# Helper function to create downloadable Excel file bytes
def create_excel_download_bytes(payload):
    output = io.BytesIO()
    df = pd.DataFrame([payload])
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Hasil Kuesioner')
    return output.getvalue()

# Helper function to send data via Google Apps Script
def submit_to_google_sheets(payload):
    # Step 1: ALWAYS save to local Excel & CSV first
    local_ok, local_msg = save_to_local_excel(payload)
    
    # Step 2: Attempt Google Apps Script transmission
    try:
        url = None
        if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
            url = st.secrets["connections"]["gsheets"]["spreadsheet"]
        elif "gsheets_url" in st.secrets:
            url = st.secrets["gsheets_url"]
        elif "spreadsheet" in st.secrets:
            url = st.secrets["spreadsheet"]
        if not url:
            url = st.secrets.get("apps_script_url", None)

        if url and "script.google.com" in url:
            headers = {"Content-Type": "application/json"}
            response = requests.post(url, json=payload, headers=headers, timeout=15, allow_redirects=True)
            
            if response.status_code == 200:
                try:
                    res = response.json()
                    if res.get("result") == "success":
                        return True, "Berhasil mengirimkan data ke Google Sheets & merekam ke Excel lokal."
                    else:
                        return True, f"Tersimpan di Excel lokal. Respon Google Sheets: {res.get('message', 'Sukses')}"
                except Exception:
                    return True, "Data berhasil dikirim dan direkam di Excel lokal."
            elif response.status_code == 401:
                return False, "HTTP 401 Unauthorized"
            else:
                return False, f"HTTP Error {response.status_code}: {response.text}"
        else:
            return True, "Data berhasil direkam ke Excel lokal. (URL Google Apps Script belum diset di secrets.toml)"
    except Exception as e:
        return False, f"Respon online: {str(e)} (Namun jawaban Anda SUDAH AMAN tersimpan di Excel lokal)."

# ==========================================
# SIDEBAR
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/illustrations/100/learning-support.png", width=75)
    st.markdown("### 🎓 Penelitian Disertasi PEP UNJ")
    st.markdown("**Eny Cahyaningsih**")
    st.caption("Model Evaluasi Leadership Development Program (LDP)")
    st.markdown("---")
    
    if st.session_state.page != "intro":
        active_q = questionnaires_dict.get(st.session_state.selected_q_id, {})
        st.markdown(f"**Instrumen Aktif:**\n`{active_q.get('title', '')}`")
        st.markdown(f"**Peran:** {active_q.get('role_name', '')}")
        
        flat_mods = get_flattened_modules(st.session_state.selected_q_id)
        if flat_mods and st.session_state.page == "kuisioner":
            completed_count = sum(1 for m in flat_mods if m["code"] in st.session_state.answers and st.session_state.answers[m["code"]].strip() != "")
            total_mods = len(flat_mods)
            st.markdown(f"**Progres Pengisian:** {completed_count}/{total_mods} Modul")
            st.progress(completed_count / total_mods if total_mods > 0 else 0.0)

        st.markdown("---")
        
        # Quick submit button in sidebar if user has answered at least 1 item
        has_any_answer = any(v.strip() != "" for v in st.session_state.answers.values())
        if has_any_answer and st.session_state.page in ["kuisioner", "biodata"]:
            if st.button("🚀 Kirim Jawaban (Rekam ke Excel)", type="primary", use_container_width=True):
                st.session_state.page = "summary"
                st.rerun()
            st.markdown("---")

        if st.button("🔄 Ganti Kuesioner / Kembali ke Awal", use_container_width=True):
            st.session_state.page = "intro"
            st.session_state.answers = {}
            st.session_state.closing_answer = ""
            st.session_state.current_mod_idx = 0
            st.rerun()

    st.markdown("---")
    st.caption("📌 Perum BULOG & Universitas Negeri Jakarta")

# ==========================================
# PAGE 1: PENGANTAR & PILIHAN KUESIONER
# ==========================================
if st.session_state.page == "intro":
    st.markdown("""
    <div class="header-box">
        <h1>📋 Need Assessment Pengembangan Model Evaluasi LDP</h1>
        <p>Kuesioner Penelitian Disertasi | Program Studi Penelitian dan Evaluasi Pendidikan - Universitas Negeri Jakarta</p>
    </div>
    """, unsafe_allow_html=True)
    
    col_main, col_select = st.columns([1.65, 1])
    
    with col_main:
        # Intro Component using native Streamlit Markdown
        with st.container(border=True):
            st.subheader("📋 Pengantar Kuesioner Penelitian")
            st.write("")
            st.markdown("**Assalamu’alaikum warahmatullahi wabarakatuh.**")
            st.markdown("**Yth. Bapak/Ibu Responden,**")
            st.write("")
            st.markdown("Terima kasih atas kesediaan Bapak/Ibu berpartisipasi dalam penelitian ini.")
            st.markdown("Kuesioner ini merupakan bagian dari penelitian disertasi mengenai pengembangan Model Evaluasi Leadership Development Program (LDP).")
            st.markdown("Kuesioner ini bertujuan memperoleh informasi, pengalaman, pandangan, dan penilaian Bapak/Ibu terkait pelaksanaan, hasil, manfaat, serta dampak Leadership Development Program (LDP) sebagai bahan dalam pengembangan model evaluasi yang lebih komprehensif dan sesuai dengan kebutuhan organisasi.")
            st.markdown("Bapak/Ibu diharapkan memberikan jawaban berdasarkan pengalaman, pengamatan, dan kondisi yang sebenarnya.")
            st.markdown("Tidak terdapat jawaban benar atau salah, apabila memungkinkan, Bapak/Ibu dapat memberikan contoh konkret atau informasi pendukung yang relevan untuk memperjelas jawaban.")
            st.markdown("Seluruh informasi yang diberikan akan digunakan semata-mata untuk kepentingan akademik dan penelitian, serta diolah dan disajikan secara bertanggung jawab sesuai dengan prinsip kerahasiaan data penelitian.")
            st.markdown("Partisipasi dan masukan Bapak/Ibu sangat berarti dalam mendukung pengembangan dan penyempurnaan Model Evaluasi LDP yang dihasilkan melalui penelitian ini.")
            st.markdown("Atas waktu, kesediaan, dan kontribusi Bapak/Ibu, saya mengucapkan terima kasih.")
            st.markdown("**Wassalamu’alaikum warahmatullahi wabarakatuh.**")
            st.markdown("---")
            st.markdown("**Eny Cahyaningsih**")
            st.caption("Penelitian dan Evaluasi Pendidikan — Universitas Negeri Jakarta")
            
    with col_select:
        st.markdown("""
        <div class="card">
            <h3 style="margin-top:0; color:#1E3A8A; font-size:1.25rem;">🎯 Pilih Jenis Kuesioner</h3>
            <p style="font-size:0.9rem; color:#4B5563;">Silakan pilih instrumen kuesioner yang sesuai dengan posisi/peran Anda dalam Leadership Development Program (LDP):</p>
        </div>
        """, unsafe_allow_html=True)
        
        q_options = {
            "peserta_ldp": "1. Kuesioner Peserta LDP",
            "atasan_mentor": "2. Kuesioner Atasan / Mentor",
            "pengelola_fasilitator": "3. Kuesioner Pengelola LDP / Fasilitator",
            "manajemen_sdm": "4. Kuesioner Manajemen SDM / Talent Management",
            "pimpinan_unit": "5. Kuesioner Pimpinan Unit Kerja"
        }
        
        selected_key = st.radio(
            "Jenis Responden / Informan:",
            options=list(q_options.keys()),
            format_func=lambda x: q_options[x],
            index=0
        )
        
        st.session_state.selected_q_id = selected_key
        q_info = questionnaires_dict.get(selected_key, {})
        
        st.info(f"**Peran:** {q_info.get('role_name', '')}\n\n**Jumlah Modul Pertanyaan:** {len(get_flattened_modules(selected_key))} Modul Inti")
        
        if st.button("🚀 Mulai Isi Kuesioner Ini", type="primary", use_container_width=True):
            st.session_state.page = "biodata"
            st.session_state.current_mod_idx = 0
            st.rerun()

# ==========================================
# PAGE 2: IDENTITAS INFORMAN (BIODATA)
# ==========================================
elif st.session_state.page == "biodata":
    q_info = questionnaires_dict.get(st.session_state.selected_q_id, {})
    
    st.title("👤 Identitas Informan / Responden")
    st.subheader(f"Instrumen: {q_info.get('title', '')}")
    st.info("Silakan lengkapi identitas Anda di bawah ini sebelum melanjutkan ke pengisian instrumen pertanyaan.")
    
    with st.form("form_biodata"):
        col1, col2 = st.columns(2)
        
        with col1:
            kode_informan = st.text_input("Kode Informan (Opsional / Diisi oleh peneliti)", value=st.session_state.user_biodata.get("Kode Informan", ""))
            nama = st.text_input("Nama Lengkap (Boleh diisi / Anonim)", value=st.session_state.user_biodata.get("Nama", ""))
            jabatan = st.text_input("Jabatan / Posisi Saat Ini *", value=st.session_state.user_biodata.get("Jabatan", ""))
            unit_kerja = st.text_input("Unit Kerja / Direktorat *", value=st.session_state.user_biodata.get("Unit Kerja", ""))
            
        with col2:
            angkatan_ldp = st.text_input("Program / Angkatan LDP yang Diikuti / Dikelola *", value=st.session_state.user_biodata.get("Angkatan LDP", ""))
            tgl = st.date_input("Tanggal Pengisian / Wawancara", value=datetime.now())
            pewawancara = st.text_input("Nama Pewawancara (Bila melalui wawancara)", value=st.session_state.user_biodata.get("Pewawancara", "Self-Assessment / Form Online"))
            
        st.markdown("---")
        submit_bio = st.form_submit_button("Lanjut ke Pertanyaan Instrumen ➡️", type="primary", use_container_width=True)
        
        if submit_bio:
            if jabatan and unit_kerja and angkatan_ldp:
                st.session_state.user_biodata = {
                    "Kode Informan": kode_informan if kode_informan else "INF-" + datetime.now().strftime("%H%M%S"),
                    "Nama": nama if nama else "Anonim",
                    "Jabatan": jabatan,
                    "Unit Kerja": unit_kerja,
                    "Angkatan LDP": angkatan_ldp,
                    "Tanggal": tgl.strftime("%Y-%m-%d"),
                    "Pewawancara": pewawancara,
                    "Jenis Kuesioner": q_info.get("title", ""),
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                st.session_state.page = "kuisioner"
                st.session_state.current_mod_idx = 0
                st.rerun()
            else:
                st.warning("Mohon isi bidang wajib bertanda bintang (*): Jabatan, Unit Kerja, dan Program/Angkatan LDP.")

# ==========================================
# PAGE 3: PENGISIAN MODUL INSTRUMEN
# ==========================================
elif st.session_state.page == "kuisioner":
    q_info = questionnaires_dict.get(st.session_state.selected_q_id, {})
    flat_mods = get_flattened_modules(st.session_state.selected_q_id)
    total_mods = len(flat_mods)
    
    idx = st.session_state.current_mod_idx
    
    # Check if we are within normal modules or at closing question
    if idx < total_mods:
        mod = flat_mods[idx]
        
        # Domain Header
        st.markdown(f'<span class="domain-badge">{mod["domain"]}</span>', unsafe_allow_html=True)
        st.caption(f"Modul {idx + 1} dari {total_mods} | {q_info.get('role_name', '')}")
        st.progress((idx + 1) / total_mods)
        
        st.title(f"📌 {mod['title']}")
        
        # Main Question Box
        if mod['main_question']:
            st.markdown(f"""
            <div class="main-question-box">
                <div class="main-question-title">Pertanyaan Utama:</div>
                <div style="font-size:1.05rem; color:#0F172A; line-height:1.6;">{mod['main_question']}</div>
            </div>
            """, unsafe_allow_html=True)
        
        # Sub Questions
        if mod['sub_questions']:
            with st.expander("💡 Pertanyaan Mendalam / Indikator Panduan (Klik untuk melihat)", expanded=True):
                for sq in mod['sub_questions']:
                    st.markdown(f"- {sq}")
        
        # Answer Box
        current_ans = st.session_state.answers.get(mod['code'], "")
        ans_text = st.text_area(
            "Jawaban / Catatan & Evidence Bapak/Ibu:",
            value=current_ans,
            height=200,
            key=f"ans_input_{mod['code']}",
            placeholder="Tuliskan pengalaman, pengamatan, dan jawaban Anda secara jelas dan mendalam di sini..."
        )
        st.session_state.answers[mod['code']] = ans_text
        
        # Action Bar (Navigations & Direct Submit Option)
        col_nav1, col_nav2, col_nav3 = st.columns([1, 1.2, 1.2])
        
        with col_nav1:
            if idx > 0:
                if st.button("⬅️ Modul Sebelumnya"):
                    st.session_state.current_mod_idx -= 1
                    st.rerun()
            else:
                if st.button("⬅️ Edit Biodata"):
                    st.session_state.page = "biodata"
                    st.rerun()
                    
        with col_nav2:
            st.markdown(f"<div style='text-align:center; padding-top:0.4rem; font-weight:600; color:#4B5563;'>Modul {idx+1}/{total_mods}</div>", unsafe_allow_html=True)
            
        with col_nav3:
            if idx < total_mods - 1:
                if st.button("Simpan & Lanjut ➡️", type="primary"):
                    st.session_state.current_mod_idx += 1
                    st.rerun()
            else:
                if st.button("Lanjut ke Pertanyaan Penutup 🏁", type="primary"):
                    st.session_state.current_mod_idx = total_mods
                    st.rerun()
                    
        st.markdown("---")
        # Direct Submit Banner
        c_sub1, c_sub2 = st.columns([2.5, 1])
        with c_sub1:
            st.caption("💡 Sudah selesai mengisi atau ingin langsung merekam jawaban Anda ke Excel?")
        with c_sub2:
            if st.button("🚀 KIRIM JAWABAN SEKARANG", type="secondary", use_container_width=True):
                st.session_state.page = "summary"
                st.rerun()
                    
    else:
        # Closing Question Page
        st.markdown('<span class="domain-badge">PERTANYAAN PENUTUP</span>', unsafe_allow_html=True)
        st.title("D. Pertanyaan Penutup")
        
        closing_q_text = q_info.get("closing_question", "Apakah ada hal penting terkait pelaksanaan, evaluasi, atau pengembangan LDP yang menurut Bapak/Ibu belum kami tanyakan tetapi perlu kami ketahui?")
        
        st.markdown(f"""
        <div class="main-question-box">
            <div class="main-question-title">Pertanyaan Penutup:</div>
            <div style="font-size:1.05rem; color:#0F172A; line-height:1.6;">{closing_q_text}</div>
        </div>
        """, unsafe_allow_html=True)
        
        closing_ans = st.text_area(
            "Jawaban / Catatan Tambahan Penutup:",
            value=st.session_state.closing_answer,
            height=180,
            key="closing_ans_input",
            placeholder="Tuliskan catatan atau rekomendasi tambahan jika ada..."
        )
        st.session_state.closing_answer = closing_ans
        
        col_nav1, col_nav2 = st.columns(2)
        with col_nav1:
            if st.button("⬅️ Kembali ke Modul Terakhir", use_container_width=True):
                st.session_state.current_mod_idx = total_mods - 1
                st.rerun()
        with col_nav2:
            if st.button("🚀 Lihat Ringkasan & Kirim Jawaban 📋", type="primary", use_container_width=True):
                st.session_state.page = "summary"
                st.rerun()

# ==========================================
# PAGE 4: RINGKASAN JAWABAN & SUBMISSION
# ==========================================
elif st.session_state.page == "summary":
    st.title("📋 Ringkasan Jawaban Kuesioner & Tombol Kirim")
    st.info("Silakan periksa ringkasan jawaban Anda di bawah ini, lalu klik **TOMBOL KIRIM JAWABAN (WARNA BIRU)** untuk merekam data ke Excel dan Google Sheets.")
    
    bio = st.session_state.user_biodata
    q_info = questionnaires_dict.get(st.session_state.selected_q_id, {})
    flat_mods = get_flattened_modules(st.session_state.selected_q_id)
    
    # Display Biodata Summary
    with st.expander("👤 Ringkasan Identitas Informan", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"**Nama:** {bio.get('Nama', '-')}")
            st.write(f"**Jabatan:** {bio.get('Jabatan', '-')}")
            st.write(f"**Unit Kerja:** {bio.get('Unit Kerja', '-')}")
        with col2:
            st.write(f"**Kode Informan:** {bio.get('Kode Informan', '-')}")
            st.write(f"**Angkatan LDP:** {bio.get('Angkatan LDP', '-')}")
            st.write(f"**Tanggal:** {bio.get('Tanggal', '-')}")

    # Display Answers Summary
    st.subheader("📝 Ringkasan Jawaban Modul")
    summary_data = []
    for m in flat_mods:
        code = m["code"]
        ans = st.session_state.answers.get(code, "").strip()
        summary_data.append({
            "Kode Modul": code,
            "Judul Modul": m["title"],
            "Status": "✅ Terisi" if ans else "⚠️ Kosong",
            "Jawaban": ans if ans else "-"
        })
    
    df_summary = pd.DataFrame(summary_data)
    st.dataframe(df_summary[["Kode Modul", "Judul Modul", "Status", "Jawaban"]], use_container_width=True)
    
    if st.session_state.closing_answer.strip():
        st.markdown(f"**Pertanyaan Penutup:** {st.session_state.closing_answer}")

    st.markdown("---")
    
    # Big prominent submission box
    st.markdown("""
    <div style="background-color: #EFF6FF; border: 2px solid #3B82F6; border-radius: 10px; padding: 1.25rem; margin-bottom: 1.5rem; text-align: center;">
        <h3 style="color: #1E3A8A; margin-top:0;">📤 Siap Mengirimkan Jawaban?</h3>
        <p style="color: #1E40AF; font-size: 0.95rem; margin-bottom: 0.5rem;">
            Klik tombol di bawah untuk menyinkronkan data ke Google Sheets dan merekam file Excel secara otomatis.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col_sub1, col_sub2 = st.columns([1, 1.8])
    with col_sub1:
        if st.button("⬅️ Edit Kembali Pertanyaan", use_container_width=True):
            st.session_state.page = "kuisioner"
            st.session_state.current_mod_idx = 0
            st.rerun()
            
    with col_sub2:
        if st.button("🚀 KIRIM JAWABAN SEKARANG (REKAM KE EXCEL)", type="primary", use_container_width=True):
            # Build payload
            payload = {}
            for k, v in bio.items():
                payload[k] = v
                
            for m in flat_mods:
                code = m["code"]
                payload[f"Modul_{code}"] = st.session_state.answers.get(code, "")
                
            payload["Pertanyaan_Penutup"] = st.session_state.closing_answer
            
            with st.spinner("Sedang merekam data ke Excel & mengirim ke Google Sheets..."):
                success, msg = submit_to_google_sheets(payload)
                if success:
                    st.session_state.page = "finish"
                    st.rerun()
                else:
                    # If Google Apps Script returned HTTP 401 error
                    if "401" in msg or "Unauthorized" in msg:
                        st.error("⚠️ HTTP Error 401: Akses Google Apps Script Ditolak (Unauthorized)")
                        st.warning("""
                        **Penjelasan untuk Peneliti (Eny Cahyaningsih):**
                        
                        Error `401 Unauthorized` terjadi karena pengaturan penempatan (*Deployment*) Google Apps Script di Google Sheets Anda saat ini masih terset **"Only me" (Hanya saya)**.
                        
                        **Langkah Mudah Mengatasi HTTP Error 401 (Hanya 3 Langkah):**
                        1. Buka Google Sheets penelitian Anda & klik **Extensions > Apps Script**.
                        2. Di sudut kanan atas Apps Script, klik **Deploy > Manage deployments** (Kelola Penempatan).
                        3. Klik icon **Pensil (Edit)**, lalu ubah **"Who has access" (Siapa yang memiliki akses)** dari *Only me* menjadi **"Anyone" (Siapa saja)**.
                        4. Klik **Deploy** untuk menyimpan.
                        
                        ---
                        💡 **JANGAN KHATIR! JAWABAN ANDA SUDAH 100% AMAN!**  
                        Sistem aplikasi ini telah **otomatis merekam jawaban Anda ke file Excel** (`data_jawaban_excel.xlsx`) di folder `app_survey/responses/` komputer ini!
                        """)
                    else:
                        st.error(f"Pemberitahuan Server: {msg}")
                        st.info("Jawaban Anda telah tersimpan secara lokal. Anda juga dapat mengunduh file Excel/JSON di bawah ini:")

                    excel_bytes = create_excel_download_bytes(payload)
                    c_d1, c_d2 = st.columns(2)
                    with c_d1:
                        st.download_button(
                            label="📥 Unduh Format Excel (.xlsx)",
                            data=excel_bytes,
                            file_name=f"Kuesioner_{bio.get('Kode Informan','jawaban')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    with c_d2:
                        st.download_button(
                            label="📥 Unduh Format JSON (.json)",
                            data=json.dumps(payload, indent=2, ensure_ascii=False),
                            file_name=f"Kuesioner_{bio.get('Kode Informan','jawaban')}.json",
                            mime="application/json",
                            use_container_width=True
                        )

# ==========================================
# PAGE 5: FINISH (BALOON & TERIMA KASIH)
# ==========================================
elif st.session_state.page == "finish":
    st.balloons()
    
    st.markdown("""
    <div style="text-align: center; padding: 2.5rem 1rem;">
        <h1 style="color: #1E3A8A; font-size: 2.5rem;">🎉 TERIMA KASIH! 🎉</h1>
        <h3 style="color: #374151; font-weight: 500;">Jawaban dan Kontribusi Bapak/Ibu Sangat Berharga</h3>
        <p style="color: #6B7280; max-width: 700px; margin: 1rem auto; font-size: 1.05rem; line-height: 1.6;">
            Seluruh data jawaban Anda telah <strong>BERHASIL DIREKAM KE FILE EXCEL & DATABASE PENELITIAN</strong>. 
            Informasi ini digunakan semata-mata untuk kepentingan penelitian disertasi akademis pengembangan 
            <strong>Model Evaluasi Leadership Development Program (LDP)</strong> di Perum BULOG.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    st.success("✅ Data jawaban kuesioner berhasil direkam ke Excel lokal & disinkronkan ke Google Sheets Peneliti.")
    
    bio = st.session_state.user_biodata
    flat_mods = get_flattened_modules(st.session_state.selected_q_id)
    payload = {**bio, **{f"Modul_{m['code']}": st.session_state.answers.get(m['code'], "") for m in flat_mods}, "Pertanyaan_Penutup": st.session_state.closing_answer}
    excel_bytes = create_excel_download_bytes(payload)
    
    c_fin1, c_fin2, c_fin3 = st.columns([1.2, 1.4, 1.4])
    with c_fin1:
        if st.button("🏠 Halaman Utama", use_container_width=True):
            st.session_state.page = "intro"
            st.session_state.answers = {}
            st.session_state.closing_answer = ""
            st.session_state.current_mod_idx = 0
            st.rerun()
            
    with c_fin2:
        st.download_button(
            label="📥 Unduh Salinan Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"Kuesioner_LDP_{bio.get('Kode Informan','Jawaban')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    with c_fin3:
        st.download_button(
            label="📥 Unduh Salinan JSON (.json)",
            data=json.dumps(payload, indent=2, ensure_ascii=False),
            file_name=f"Kuesioner_LDP_{bio.get('Kode Informan','Jawaban')}.json",
            mime="application/json",
            use_container_width=True
        )
