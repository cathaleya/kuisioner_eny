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
    page_title="Kuesioner Analisis Kebutuhan Pembelajaran",
    page_icon="📋",
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
        background: linear-gradient(135deg, #065F46 0%, #059669 100%);
        color: white;
        padding: 2rem;
        border-radius: 14px;
        box-shadow: 0 4px 14px rgba(6, 95, 70, 0.18);
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
        color: #D1FAE5 !important;
        font-size: 1.05rem;
        margin-bottom: 0;
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
        background-color: #ECFDF5;
        color: #065F46;
        font-weight: 600;
        font-size: 0.85rem;
        padding: 0.35rem 0.85rem;
        border-radius: 9999px;
        border: 1px solid #A7F3D0;
        margin-bottom: 0.75rem;
    }
    
    /* Main Question Box */
    .main-question-box {
        background-color: #F8FAFC;
        border-left: 5px solid #059669;
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
    
    /* Consent Box */
    .consent-box {
        background: linear-gradient(135deg, #FFFBEB 0%, #FEF3C7 100%);
        border: 2px solid #F59E0B;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
    }
    .consent-box h3 {
        color: #92400E;
        margin-top: 0;
    }

    /* Buttons */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        padding: 0.55rem 1.3rem;
        transition: all 0.2s;
    }

    /* Submit Callout Card */
    .submit-banner {
        background: linear-gradient(135deg, #ECFDF5 0%, #D1FAE5 100%);
        border: 2px solid #059669;
        border-radius: 12px;
        padding: 1.25rem;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
        text-align: center;
    }

    /* Mobile Responsive Optimizations for Smartphones (HP & Tablet) */
    @media (max-width: 768px) {
        .main .block-container {
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
            padding-top: 1rem !important;
        }
        .header-box {
            padding: 1.2rem !important;
            border-radius: 10px !important;
        }
        .header-box h1 {
            font-size: 1.35rem !important;
            line-height: 1.35 !important;
        }
        .header-box p {
            font-size: 0.9rem !important;
        }
        .main-question-box {
            padding: 1rem !important;
            border-left-width: 4px !important;
        }
        .main-question-title {
            font-size: 1rem !important;
        }
        .stButton>button {
            width: 100% !important;
            padding: 0.65rem 1rem !important;
            font-size: 0.95rem !important;
            margin-bottom: 0.4rem !important;
        }
        .stTextArea textarea {
            font-size: 0.95rem !important;
        }
        .domain-badge {
            font-size: 0.78rem !important;
            padding: 0.3rem 0.65rem !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# LOAD DATA
# ==========================================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
QUESTIONS_PATH = os.path.join(CURRENT_DIR, "questions_analisis_kebutuhan.json")
RESPONSES_DIR = os.path.join(CURRENT_DIR, "responses")
os.makedirs(RESPONSES_DIR, exist_ok=True)

LOCAL_EXCEL_PATH = os.path.join(RESPONSES_DIR, "data_analisis_kebutuhan_excel.xlsx")
LOCAL_CSV_PATH = os.path.join(RESPONSES_DIR, "data_analisis_kebutuhan.csv")

@st.cache_data
def load_questions():
    if not os.path.exists(QUESTIONS_PATH):
        return None
    with open(QUESTIONS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

data = load_questions()

if not data:
    st.error(f"File `{QUESTIONS_PATH}` tidak ditemukan.")
    st.stop()

questionnaires_dict = data.get("questionnaires", {})
meta = data.get("meta", {})
consent_text = data.get("consent_text", "")

# ==========================================
# SESSION STATE INITIALIZATION
# ==========================================
if "page" not in st.session_state:
    st.session_state.page = "intro"
if "selected_q_id" not in st.session_state:
    st.session_state.selected_q_id = "dosen"
if "user_biodata" not in st.session_state:
    st.session_state.user_biodata = {}
if "answers" not in st.session_state:
    st.session_state.answers = {}
if "closing_answer" not in st.session_state:
    st.session_state.closing_answer = ""
if "current_mod_idx" not in st.session_state:
    st.session_state.current_mod_idx = 0
if "last_submit_msg" not in st.session_state:
    st.session_state.last_submit_msg = ""
if "last_payload" not in st.session_state:
    st.session_state.last_payload = {}
if "consent_given" not in st.session_state:
    st.session_state.consent_given = False

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
        try:
            if os.path.exists(LOCAL_CSV_PATH):
                df_existing = pd.read_csv(LOCAL_CSV_PATH)
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                df_combined.to_csv(LOCAL_CSV_PATH, index=False, encoding="utf-8-sig")
            else:
                df_new.to_csv(LOCAL_CSV_PATH, index=False, encoding="utf-8-sig")
        except Exception:
            pass
            
        # Save or append to Excel
        try:
            if os.path.exists(LOCAL_EXCEL_PATH):
                df_existing_excel = pd.read_excel(LOCAL_EXCEL_PATH)
                df_combined_excel = pd.concat([df_existing_excel, df_new], ignore_index=True)
                df_combined_excel.to_excel(LOCAL_EXCEL_PATH, index=False)
            else:
                df_new.to_excel(LOCAL_EXCEL_PATH, index=False)
        except Exception:
            pass
            
        return True, "Data berhasil direkam."
    except Exception as e:
        return False, f"Catatan penyimpanan lokal: {str(e)}"

# Helper function to create downloadable Excel file bytes (with CSV fallback)
def create_excel_download_bytes(payload):
    try:
        output = io.BytesIO()
        df = pd.DataFrame([payload])
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Hasil Kuesioner')
        return output.getvalue()
    except Exception:
        try:
            df = pd.DataFrame([payload])
            return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        except Exception:
            return b""

# Helper function to send data via Google Apps Script & Local Excel
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
                return False, "HTTP 401 Unauthorized (Google Apps Script butuh akses 'Anyone')"
            else:
                return False, f"HTTP Error {response.status_code}: {response.text}"
        else:
            return True, "Data berhasil direkam ke Excel lokal komputer."
    except Exception as e:
        return True, f"Jawaban Anda SUDAH AMAN tersimpan di Excel lokal. (Koneksi online: {str(e)})"

# Global helper to perform direct submission from ANY page
def process_direct_submission():
    payload = {}
    bio = st.session_state.user_biodata
    for k, v in bio.items():
        payload[k] = v
        
    flat_mods = get_flattened_modules(st.session_state.selected_q_id)
    for m in flat_mods:
        code = m["code"]
        payload[f"Modul_{code}"] = st.session_state.answers.get(code, "")
        
    payload["Pertanyaan_Penutup"] = st.session_state.closing_answer
    
    with st.spinner("Sedang merekam jawaban Anda ke Excel & Google Sheets..."):
        success, msg = submit_to_google_sheets(payload)
        st.session_state.last_submit_msg = msg
        st.session_state.last_payload = payload
        st.session_state.page = "finish"
        st.rerun()

# ==========================================
# SIDEBAR
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/illustrations/100/student-male.png", width=75)
    st.markdown("### 🎓 Penelitian Disertasi PEP UNJ")
    st.markdown(f"**{meta.get('peneliti', 'Eny Cahyaningsih')}**")
    st.caption("Analisis Kebutuhan Pembelajaran")
    st.markdown("---")
    
    if st.session_state.page != "intro":
        active_q = questionnaires_dict.get(st.session_state.selected_q_id, {})
        st.markdown(f"**Instrumen Aktif:**\n`{active_q.get('title', '')}`")
        st.markdown(f"**Peran:** {active_q.get('role_name', '')}")
        
        flat_mods = get_flattened_modules(st.session_state.selected_q_id)
        if flat_mods and st.session_state.page == "kuisioner":
            completed_count = sum(1 for m in flat_mods if m["code"] in st.session_state.answers and st.session_state.answers[m["code"]].strip() != "")
            total_mods = len(flat_mods)
            st.markdown(f"**Progres Pengisian:** {completed_count}/{total_mods} Bagian")
            st.progress(completed_count / total_mods if total_mods > 0 else 0.0)

        st.markdown("---")
        
        # PROMINENT SIDEBAR SUBMIT BUTTON
        if st.session_state.page in ["kuisioner", "biodata", "summary"]:
            if st.button("🚀 KIRIM JAWABAN SEKARANG", type="primary", use_container_width=True):
                process_direct_submission()
            st.markdown("---")

        if st.button("🔄 Ganti Kuesioner / Kembali ke Awal", use_container_width=True):
            st.session_state.page = "intro"
            st.session_state.answers = {}
            st.session_state.closing_answer = ""
            st.session_state.current_mod_idx = 0
            st.session_state.consent_given = False
            st.rerun()

    st.markdown("---")
    st.caption(f"📌 {meta.get('institusi', 'Universitas Negeri Jakarta')}")

# ==========================================
# PAGE 1: PENGANTAR & PILIHAN KUESIONER
# ==========================================
if st.session_state.page == "intro":
    st.markdown("""
    <div class="header-box">
        <h1>📋 Kuesioner Analisis Kebutuhan Pembelajaran</h1>
        <p>Penelitian Disertasi | Program Studi Penelitian dan Evaluasi Pendidikan — Universitas Negeri Jakarta</p>
    </div>
    """, unsafe_allow_html=True)
    
    col_main, col_select = st.columns([1.65, 1])
    
    with col_main:
        active_q_id = st.session_state.selected_q_id
        active_q_data = questionnaires_dict.get(active_q_id, {})
        intro_key = active_q_data.get("intro_key", "intro_dosen")
        intro_text = data.get(intro_key, "")
        
        with st.container(border=True):
            st.subheader("📋 Pengantar Kuesioner Penelitian")
            st.write("")
            for line in intro_text.split("\n"):
                if line.strip():
                    if line.startswith("Eny Cahyaningsih") or line.startswith("Penelitian"):
                        st.markdown(f"**{line.strip()}**")
                    else:
                        st.markdown(line.strip())
                else:
                    st.write("")
            
    with col_select:
        st.markdown("""
        <div class="card">
            <h3 style="margin-top:0; color:#065F46; font-size:1.25rem;">🎯 Pilih Jenis Kuesioner</h3>
            <p style="font-size:0.9rem; color:#4B5563;">Silakan pilih instrumen kuesioner yang sesuai dengan peran Anda:</p>
        </div>
        """, unsafe_allow_html=True)
        
        q_options = {
            "dosen": "1. Kuesioner untuk Dosen",
            "mahasiswa": "2. Kuesioner untuk Mahasiswa",
        }
        
        selected_key = st.radio(
            "Jenis Responden:",
            options=list(q_options.keys()),
            format_func=lambda x: q_options[x],
            index=list(q_options.keys()).index(st.session_state.selected_q_id) if st.session_state.selected_q_id in q_options else 0
        )
        
        st.session_state.selected_q_id = selected_key
        q_info = questionnaires_dict.get(selected_key, {})
        
        st.info(f"**Peran:** {q_info.get('role_name', '')}\n\n**Jumlah Bagian:** {len(get_flattened_modules(selected_key))} Bagian Pertanyaan")
        
        if st.button("📄 Lanjut ke Persetujuan & Identitas", type="primary", use_container_width=True):
            st.session_state.page = "consent"
            st.session_state.current_mod_idx = 0
            st.session_state.consent_given = False
            st.rerun()

# ==========================================
# PAGE 1.5: FORM CONSENT (PERNYATAAN PERSETUJUAN)
# ==========================================
elif st.session_state.page == "consent":
    q_info = questionnaires_dict.get(st.session_state.selected_q_id, {})
    
    st.markdown("""
    <div class="header-box">
        <h1>📜 Pernyataan Persetujuan (Informed Consent)</h1>
        <p>Silakan baca dengan seksama sebelum melanjutkan pengisian kuesioner</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="consent-box">
        <h3>⚠️ Pernyataan Persetujuan Penelitian</h3>
        <p style="font-size: 0.95rem; color: #78350F; line-height: 1.7;">
            {consent_text.replace(chr(10), '<br>')}
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    with st.container(border=True):
        st.subheader("🔒 Informasi Perlindungan Data")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            **Data yang dikumpulkan:**
            - Identitas (opsional/anonim)
            - Jawaban kuesioner
            - Tanggal pengisian
            """)
        with col2:
            st.markdown("""
            **Kerahasiaan Anda:**
            - Data disimpan secara aman
            - Tidak disebarluaskan ke publik
            - Hanya untuk kepentingan penelitian
            """)
    
    st.markdown("---")
    
    consent_check = st.checkbox(
        "✅ Saya telah membaca, memahami, dan **MENYETUJUI** untuk berpartisipasi dalam penelitian ini secara sukarela.",
        value=st.session_state.consent_given,
        key="consent_checkbox"
    )
    
    col_c1, col_c2, col_c3 = st.columns([1, 1, 1.5])
    with col_c1:
        if st.button("⬅️ Kembali ke Awal", use_container_width=True):
            st.session_state.page = "intro"
            st.rerun()
    with col_c3:
        if st.button("Lanjut ke Identitas ➡️", type="primary", use_container_width=True, disabled=not consent_check):
            if consent_check:
                st.session_state.consent_given = True
                st.session_state.page = "biodata"
                st.rerun()
            else:
                st.warning("Mohon centang pernyataan persetujuan terlebih dahulu untuk melanjutkan.")

# ==========================================
# PAGE 2: IDENTITAS INFORMAN (BIODATA)
# ==========================================
elif st.session_state.page == "biodata":
    q_info = questionnaires_dict.get(st.session_state.selected_q_id, {})
    
    st.title("👤 Identitas Responden")
    st.subheader(f"Instrumen: {q_info.get('title', '')}")
    st.info("Silakan lengkapi identitas Anda di bawah ini. Kolom bertanda (*) wajib diisi. Identitas dapat diisi dengan anonim jika diinginkan.")
    
    # --- BIODATA FORM FOR DOSEN ---
    if st.session_state.selected_q_id == "dosen":
        with st.form("form_biodata_dosen"):
            col1, col2 = st.columns(2)
            
            with col1:
                kode_informan = st.text_input("Kode Responden (Diisi oleh peneliti / opsional)", value=st.session_state.user_biodata.get("Kode Responden", ""))
                nama = st.text_input("Nama Lengkap (Boleh anonim)", value=st.session_state.user_biodata.get("Nama", ""))
                institusi_asal = st.text_input("Nama Perguruan Tinggi / Institusi *", value=st.session_state.user_biodata.get("Institusi", ""))
                prodi = st.text_input("Program Studi *", value=st.session_state.user_biodata.get("Program Studi", ""))
                
            with col2:
                mata_kuliah = st.text_input("Mata Kuliah yang Diampu *", value=st.session_state.user_biodata.get("Mata Kuliah Diampu", ""))
                jenjang_mengajar = st.selectbox("Jenjang Mengajar *", ["D3", "S1", "S2", "S3", "D3 & S1", "S1 & S2"], index=0)
                lama_mengajar = st.text_input("Lama Mengajar (tahun) *", value=st.session_state.user_biodata.get("Lama Mengajar", ""))
                tgl = st.date_input("Tanggal Pengisian", value=datetime.now())
                
            st.markdown("---")
            submit_bio = st.form_submit_button("Lanjut ke Kuesioner ➡️", type="primary", use_container_width=True)
            
            if submit_bio:
                if institusi_asal and prodi and mata_kuliah and lama_mengajar:
                    st.session_state.user_biodata = {
                        "Kode Responden": kode_informan if kode_informan else "DSN-" + datetime.now().strftime("%H%M%S"),
                        "Nama": nama if nama else "Anonim",
                        "Institusi": institusi_asal,
                        "Program Studi": prodi,
                        "Mata Kuliah Diampu": mata_kuliah,
                        "Jenjang Mengajar": jenjang_mengajar,
                        "Lama Mengajar": lama_mengajar,
                        "Tanggal": tgl.strftime("%Y-%m-%d"),
                        "Jenis Kuesioner": q_info.get("title", ""),
                        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Consent": "Setuju"
                    }
                    st.session_state.page = "kuisioner"
                    st.session_state.current_mod_idx = 0
                    st.rerun()
                else:
                    st.warning("Mohon isi semua kolom wajib bertanda (*): Institusi, Program Studi, Mata Kuliah, dan Lama Mengajar.")
    
    # --- BIODATA FORM FOR MAHASISWA ---
    else:
        with st.form("form_biodata_mahasiswa"):
            col1, col2 = st.columns(2)
            
            with col1:
                kode_informan = st.text_input("Kode Responden (Diisi oleh peneliti / opsional)", value=st.session_state.user_biodata.get("Kode Responden", ""))
                nama = st.text_input("Nama Lengkap (Boleh anonim)", value=st.session_state.user_biodata.get("Nama", ""))
                nim = st.text_input("NIM (Nomor Induk Mahasiswa — Boleh anonim)", value=st.session_state.user_biodata.get("NIM", ""))
                institusi_asal = st.text_input("Nama Perguruan Tinggi / Institusi *", value=st.session_state.user_biodata.get("Institusi", ""))
                prodi = st.text_input("Program Studi *", value=st.session_state.user_biodata.get("Program Studi", ""))
                
            with col2:
                semester = st.selectbox("Semester Saat Ini *", ["1", "2", "3", "4", "5", "6", "7", "8", "> 8"], index=0)
                angkatan = st.text_input("Angkatan / Tahun Masuk *", value=st.session_state.user_biodata.get("Angkatan", ""))
                tgl = st.date_input("Tanggal Pengisian", value=datetime.now())
                
            st.markdown("---")
            submit_bio = st.form_submit_button("Lanjut ke Kuesioner ➡️", type="primary", use_container_width=True)
            
            if submit_bio:
                if institusi_asal and prodi and angkatan:
                    st.session_state.user_biodata = {
                        "Kode Responden": kode_informan if kode_informan else "MHS-" + datetime.now().strftime("%H%M%S"),
                        "Nama": nama if nama else "Anonim",
                        "NIM": nim if nim else "Anonim",
                        "Institusi": institusi_asal,
                        "Program Studi": prodi,
                        "Semester": semester,
                        "Angkatan": angkatan,
                        "Tanggal": tgl.strftime("%Y-%m-%d"),
                        "Jenis Kuesioner": q_info.get("title", ""),
                        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Consent": "Setuju"
                    }
                    st.session_state.page = "kuisioner"
                    st.session_state.current_mod_idx = 0
                    st.rerun()
                else:
                    st.warning("Mohon isi semua kolom wajib bertanda (*): Institusi, Program Studi, dan Angkatan.")

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
        st.caption(f"Bagian {idx + 1} dari {total_mods} | {q_info.get('role_name', '')}")
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
            with st.expander("💡 Pertanyaan Panduan / Indikator (Klik untuk melihat)", expanded=True):
                for sq in mod['sub_questions']:
                    st.markdown(f"- {sq}")
        
        # Answer Box
        current_ans = st.session_state.answers.get(mod['code'], "")
        ans_text = st.text_area(
            "Jawaban / Catatan Anda:",
            value=current_ans,
            height=200,
            key=f"ans_input_{mod['code']}",
            placeholder="Tuliskan pengalaman, pengamatan, dan jawaban Anda secara jelas di sini..."
        )
        st.session_state.answers[mod['code']] = ans_text
        
        # Action Bar (Navigations & Direct Submit Option)
        if idx < total_mods - 1:
            col_nav1, col_nav2, col_nav3 = st.columns([1, 1.2, 1.5])
            with col_nav1:
                if idx > 0:
                    if st.button("⬅️ Bagian Sebelumnya", use_container_width=True):
                        st.session_state.current_mod_idx -= 1
                        st.rerun()
                else:
                    if st.button("⬅️ Edit Identitas", use_container_width=True):
                        st.session_state.page = "biodata"
                        st.rerun()
                        
            with col_nav2:
                st.markdown(f"<div style='text-align:center; padding-top:0.4rem; font-weight:600; color:#4B5563;'>Bagian {idx+1}/{total_mods}</div>", unsafe_allow_html=True)
                
            with col_nav3:
                if st.button("Simpan & Lanjut ➡️", type="primary", use_container_width=True):
                    st.session_state.current_mod_idx += 1
                    st.rerun()
        else:
            # LAST MODULE - PRIMARY BUTTON IS KIRIM JAWABAN SEKARANG!
            col_nav1, col_nav2, col_nav3 = st.columns([1, 1.2, 1.8])
            with col_nav1:
                if st.button("⬅️ Bagian Sebelumnya", use_container_width=True):
                    st.session_state.current_mod_idx -= 1
                    st.rerun()
            with col_nav2:
                if st.button("📝 Ke Pertanyaan Penutup", use_container_width=True):
                    st.session_state.current_mod_idx = total_mods
                    st.rerun()
            with col_nav3:
                if st.button("🚀 KIRIM JAWABAN SEKARANG", type="primary", use_container_width=True, key="btn_last_mod_submit_primary"):
                    process_direct_submission()
                    
        # PROMINENT SUBMIT CARD ON EVERY QUESTION PAGE
        st.markdown("""
        <div class="submit-banner">
            <h4 style="color:#065F46; margin:0 0 0.5rem 0;">📤 Selesai Mengisi? Kirim Jawaban Sekarang</h4>
            <p style="color:#065F46; font-size:0.9rem; margin-bottom:0.8rem;">
                Klik tombol hijau di atas atau di bawah untuk langsung merekam seluruh jawaban Anda ke file Excel.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("🚀 KIRIM JAWABAN SEKARANG (REKAM KE EXCEL)", type="primary", use_container_width=True, key=f"btn_submit_mod_{idx}"):
            process_direct_submission()
                    
    else:
        # Closing Question Page
        st.markdown('<span class="domain-badge">PERTANYAAN PENUTUP</span>', unsafe_allow_html=True)
        st.title("E. Pertanyaan Penutup")
        
        closing_q_text = q_info.get("closing_question", "Apakah ada hal penting yang belum kami tanyakan dan ingin Anda sampaikan kepada peneliti?")
        
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
        
        # CLOSING PAGE NAVIGATION - PRIMARY BUTTON IS KIRIM JAWABAN SEKARANG!
        col_nav1, col_nav2, col_nav3 = st.columns([1, 1, 1.8])
        with col_nav1:
            if st.button("⬅️ Bagian Terakhir", use_container_width=True):
                st.session_state.current_mod_idx = total_mods - 1
                st.rerun()
        with col_nav2:
            if st.button("📋 Lihat Ringkasan", use_container_width=True):
                st.session_state.page = "summary"
                st.rerun()
        with col_nav3:
            if st.button("🚀 KIRIM JAWABAN SEKARANG", type="primary", use_container_width=True, key="btn_closing_submit_primary"):
                process_direct_submission()

        # PROMINENT SUBMIT CARD ON CLOSING PAGE
        st.markdown("""
        <div class="submit-banner">
            <h3 style="color:#065F46; margin:0 0 0.5rem 0;">🚀 SELESAI PENGISIAN? KIRIM JAWABAN SEKARANG</h3>
            <p style="color:#065F46; font-size:0.95rem; margin-bottom:0.8rem;">
                Klik tombol di bawah ini untuk menyimpan seluruh data jawaban Anda ke Excel.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("🚀 KIRIM JAWABAN SEKARANG (REKAM KE EXCEL)", type="primary", use_container_width=True, key="btn_submit_closing"):
            process_direct_submission()

# ==========================================
# PAGE 4: RINGKASAN JAWABAN & SUBMISSION
# ==========================================
elif st.session_state.page == "summary":
    st.title("📋 Ringkasan Jawaban Kuesioner")
    st.info("Periksa ringkasan jawaban Anda di bawah ini, lalu klik **TOMBOL KIRIM JAWABAN (WARNA HIJAU)** untuk merekam data ke Excel.")
    
    bio = st.session_state.user_biodata
    q_info = questionnaires_dict.get(st.session_state.selected_q_id, {})
    flat_mods = get_flattened_modules(st.session_state.selected_q_id)
    
    # Display Biodata Summary
    with st.expander("👤 Ringkasan Identitas Responden", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"**Nama:** {bio.get('Nama', '-')}")
            if bio.get('NIM'):
                st.write(f"**NIM:** {bio.get('NIM', '-')}")
            st.write(f"**Institusi:** {bio.get('Institusi', '-')}")
            st.write(f"**Program Studi:** {bio.get('Program Studi', '-')}")
        with col2:
            st.write(f"**Kode Responden:** {bio.get('Kode Responden', '-')}")
            if bio.get('Mata Kuliah Diampu'):
                st.write(f"**Mata Kuliah:** {bio.get('Mata Kuliah Diampu', '-')}")
            if bio.get('Semester'):
                st.write(f"**Semester:** {bio.get('Semester', '-')}")
            st.write(f"**Tanggal:** {bio.get('Tanggal', '-')}")

    # Display Answers Summary
    st.subheader("📝 Ringkasan Jawaban")
    summary_data = []
    for m in flat_mods:
        code = m["code"]
        ans = st.session_state.answers.get(code, "").strip()
        summary_data.append({
            "Kode": code,
            "Judul Bagian": m["title"],
            "Status": "✅ Terisi" if ans else "⚠️ Kosong",
            "Jawaban": ans if ans else "-"
        })
    
    df_summary = pd.DataFrame(summary_data)
    st.dataframe(df_summary[["Kode", "Judul Bagian", "Status", "Jawaban"]], use_container_width=True)
    
    if st.session_state.closing_answer.strip():
        st.markdown(f"**Catatan Penutup:** {st.session_state.closing_answer}")

    st.markdown("---")
    
    # Big prominent submission box
    st.markdown("""
    <div class="submit-banner">
        <h3 style="color: #065F46; margin-top:0;">📤 Siap Mengirimkan Jawaban?</h3>
        <p style="color: #065F46; font-size: 0.95rem; margin-bottom: 0.5rem;">
            Klik tombol di bawah untuk merekam seluruh jawaban ke file Excel secara otomatis.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col_sub1, col_sub2 = st.columns([1, 1.8])
    with col_sub1:
        if st.button("⬅️ Edit Kembali Jawaban", use_container_width=True):
            st.session_state.page = "kuisioner"
            st.session_state.current_mod_idx = 0
            st.rerun()
            
    with col_sub2:
        if st.button("🚀 KIRIM JAWABAN SEKARANG (REKAM KE EXCEL)", type="primary", use_container_width=True, key="btn_submit_summary"):
            process_direct_submission()

# ==========================================
# PAGE 5: FINISH (BALOON & TERIMA KASIH)
# ==========================================
elif st.session_state.page == "finish":
    st.balloons()
    
    st.markdown("""
    <div style="text-align: center; padding: 2.5rem 1rem;">
        <h1 style="color: #065F46; font-size: 2.5rem;">🎉 TERIMA KASIH! 🎉</h1>
        <h3 style="color: #374151; font-weight: 500;">Jawaban dan Kontribusi Anda Sangat Berharga</h3>
        <p style="color: #6B7280; max-width: 700px; margin: 1rem auto; font-size: 1.05rem; line-height: 1.6;">
            Seluruh data jawaban Anda telah <strong>BERHASIL DIREKAM KE FILE EXCEL</strong>. 
            Informasi ini digunakan semata-mata untuk kepentingan penelitian disertasi akademis 
            <strong>Analisis Kebutuhan Pembelajaran</strong>.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    msg = st.session_state.last_submit_msg
    if "HTTP 401" in msg:
        st.warning("✅ Data jawaban Anda SUDAH 100% AMAN DIREKAM ke file Excel lokal (`data_analisis_kebutuhan_excel.xlsx`) di komputer ini.")
        st.info("💡 (Catatan Google Sheets online: Memerlukan perizinan 'Anyone' pada Apps Script).")
    else:
        st.success(f"✅ {msg}")
    
    payload = st.session_state.last_payload
    if payload:
        excel_bytes = create_excel_download_bytes(payload)
    else:
        excel_bytes = b""
        
    bio = st.session_state.user_biodata
    
    c_fin1, c_fin2, c_fin3 = st.columns([1.2, 1.4, 1.4])
    with c_fin1:
        if st.button("🏠 Halaman Utama", use_container_width=True):
            st.session_state.page = "intro"
            st.session_state.answers = {}
            st.session_state.closing_answer = ""
            st.session_state.current_mod_idx = 0
            st.session_state.consent_given = False
            st.rerun()
            
    with c_fin2:
        if excel_bytes:
            st.download_button(
                label="📥 Unduh Salinan Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"Kuesioner_AK_{bio.get('Kode Responden','Jawaban')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
    with c_fin3:
        if payload:
            st.download_button(
                label="📥 Unduh Salinan JSON (.json)",
                data=json.dumps(payload, indent=2, ensure_ascii=False),
                file_name=f"Kuesioner_AK_{bio.get('Kode Responden','Jawaban')}.json",
                mime="application/json",
                use_container_width=True
            )
