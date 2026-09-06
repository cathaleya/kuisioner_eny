import streamlit as st
import json
import pandas as pd
import requests
import os
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
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1000px;
    }
    
    /* Header Banner */
    .header-box {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        color: white;
        padding: 2rem;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        margin-bottom: 2rem;
    }
    .header-box h1 {
        color: #FFFFFF !important;
        font-size: 1.8rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    .header-box p {
        color: #E0E7FF !important;
        font-size: 1rem;
        margin-bottom: 0;
    }
    
    /* Cards */
    .card {
        background-color: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 2px 6px rgba(0,0,0,0.04);
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
        padding: 1.25rem;
        border-radius: 6px;
        margin-top: 0.5rem;
        margin-bottom: 1.25rem;
    }
    .main-question-title {
        color: #1E293B;
        font-weight: 700;
        font-size: 1.1rem;
        margin-bottom: 0.5rem;
    }
    
    /* Sub questions styling */
    .sub-questions-box {
        background-color: #FEFCE8;
        border: 1px solid #FEF08A;
        border-radius: 8px;
        padding: 1rem 1.25rem;
        margin-bottom: 1.25rem;
    }
    .sub-questions-title {
        color: #854D0E;
        font-weight: 600;
        font-size: 0.95rem;
        margin-bottom: 0.5rem;
    }
    
    /* Buttons */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        padding: 0.5rem 1.5rem;
        transition: all 0.2s;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# LOAD DATA
# ==========================================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
QUESTIONS_PATH = os.path.join(CURRENT_DIR, "questions.json")

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

intro_text = data.get("intro", "")
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

# Helper function to send data via Google Apps Script
def submit_to_google_sheets(payload):
    try:
        url = None
        if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
            url = st.secrets["connections"]["gsheets"]["spreadsheet"]
        elif "gsheets_url" in st.secrets:
            url = st.secrets["gsheets_url"]
        elif "spreadsheet" in st.secrets:
            url = st.secrets["spreadsheet"]
        
        # Fallback default Apps Script URL if provided in secrets
        if not url:
            # Check if default script URL is saved in session or hardcoded
            url = st.secrets.get("apps_script_url", None)

        if url and "script.google.com" in url:
            response = requests.post(url, json=payload, timeout=15)
            if response.status_code == 200:
                try:
                    res = response.json()
                    if res.get("result") == "success":
                        return True, "Berhasil mengirimkan jawaban ke Google Sheets."
                    else:
                        return False, f"Apps Script Response Error: {res.get('message', 'Unknown error')}"
                except Exception:
                    return True, "Data berhasil dikirim."
            else:
                return False, f"HTTP Error {response.status_code}: {response.text}"
        else:
            return False, "URL Apps Script belum dikonfigurasi di secrets.toml."
    except Exception as e:
        return False, f"Gagal menghubungi server: {str(e)}"

# ==========================================
# SIDEBAR
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/illustrations/100/learning-support.png", width=70)
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
        if st.button("🔄 Ganti Kuesioner / Kembali ke Awal"):
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
    
    col_main, col_select = st.columns([1.6, 1])
    
    with col_main:
        st.subheader("📜 Pengantar Kuesioner Penelitian")
        with st.container():
            st.markdown(intro_text)
            
    with col_select:
        st.markdown("""
        <div class="card">
            <h3 style="margin-top:0; color:#1E3A8A;">🎯 Pilih Jenis Kuesioner</h3>
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
        
        col_nav1, col_nav2, col_nav3 = st.columns([1, 1, 1])
        
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
            st.markdown(f"<div style='text-align:center; padding-top:0.5rem; font-weight:600; color:#6B7280;'>Modul {idx+1}/{total_mods}</div>", unsafe_allow_html=True)
            
        with col_nav3:
            if idx < total_mods - 1:
                if st.button("Simpan & Lanjut ➡️", type="primary"):
                    st.session_state.current_mod_idx += 1
                    st.rerun()
            else:
                if st.button("Lanjut ke Pertanyaan Penutup 🏁", type="primary"):
                    st.session_state.current_mod_idx = total_mods
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
            if st.button("⬅️ Kembali ke Modul Terakhir"):
                st.session_state.current_mod_idx = total_mods - 1
                st.rerun()
        with col_nav2:
            if st.button("Lihat Ringkasan Jawaban 📋", type="primary"):
                st.session_state.page = "summary"
                st.rerun()

# ==========================================
# PAGE 4: RINGKASAN JAWABAN & SUBMISSION
# ==========================================
elif st.session_state.page == "summary":
    st.title("📋 Ringkasan Jawaban Kuesioner")
    st.info("Silakan periksa kembali jawaban Anda sebelum mengirimkan ke database penelitian.")
    
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
    
    col_sub1, col_sub2 = st.columns(2)
    with col_sub1:
        if st.button("⬅️ Kembali Edit Pertanyaan"):
            st.session_state.page = "kuisioner"
            st.session_state.current_mod_idx = 0
            st.rerun()
            
    with col_sub2:
        if st.button("🚀 Kirim Jawaban Sekarang", type="primary", use_container_width=True):
            # Prepare payload for Apps Script / Google Sheets
            payload = {}
            for k, v in bio.items():
                payload[k] = v
                
            for m in flat_mods:
                code = m["code"]
                payload[f"Modul_{code}"] = st.session_state.answers.get(code, "")
                
            payload["Pertanyaan_Penutup"] = st.session_state.closing_answer
            
            with st.spinner("Sedang menyimpan data ke Google Sheets..."):
                success, msg = submit_to_google_sheets(payload)
                if success:
                    st.session_state.page = "finish"
                    st.rerun()
                else:
                    st.error(f"Gagal mengirimkan data: {msg}")
                    st.info("Anda dapat mengunduh salinan data jawaban di bawah ini dan mengirungkannya ke peneliti:")
                    st.download_button(
                        label="📥 Unduh File Jawaban (JSON)",
                        data=json.dumps(payload, indent=2, ensure_ascii=False),
                        file_name=f"Kuesioner_{bio.get('Kode Informan','data')}.json",
                        mime="application/json"
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
            Seluruh data yang Bapak/Ibu sampaikan telah berhasil disimpan secara aman. 
            Informasi ini akan digunakan semata-mata untuk kepentingan penelitian akademis pengembangan 
            <strong>Model Evaluasi Leadership Development Program (LDP)</strong> di Perum BULOG.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    st.success("✅ Data jawaban kuesioner berhasil dikirimkan ke Google Sheets Peneliti.")
    
    col_fin1, col_fin2 = st.columns(2)
    with col_fin1:
        if st.button("🏠 Kembali ke Halaman Utama / Mulai Baru", use_container_width=True):
            st.session_state.page = "intro"
            st.session_state.answers = {}
            st.session_state.closing_answer = ""
            st.session_state.current_mod_idx = 0
            st.rerun()
            
    with col_fin2:
        bio = st.session_state.user_biodata
        payload = {**bio, **{f"Modul_{k}": v for k, v in st.session_state.answers.items()}, "Pertanyaan_Penutup": st.session_state.closing_answer}
        st.download_button(
            label="📥 Unduh Salinan Jawaban Anda (JSON)",
            data=json.dumps(payload, indent=2, ensure_ascii=False),
            file_name=f"Kuesioner_LDP_{bio.get('Kode Informan','Jawaban')}.json",
            mime="application/json",
            use_container_width=True
        )
