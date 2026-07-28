# ============================================================
# RPM CERDAS AI v2.5 ONLINE & COMPREHENSIVE GENERATION
# BAGIAN 1: CONFIG, LLM INTEGRATION, & CORE LOGIC
# ============================================================

import io
import json
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Menggunakan pustaka resmi Google GenAI terbaru
try:
    from google import genai
    from google.genai import types
except ImportError:
    st.error("Pustaka 'google-genai' belum terinstal. Silakan instal terlebih dahulu menggunakan perintah: pip install google-genai")

# Konfigurasi Halaman Utama Streamlit
st.set_page_config(
    page_title="RPM CERDAS AI ONLINE",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded"
)

if "hasil" not in st.session_state:
    st.session_state.hasil = ""
if "components" not in st.session_state:
    st.session_state.components = None

# Gaya Tampilan Antarmuka (CSS Custom)
st.markdown("""
<style>
.main-title { font-size: 34px; font-weight: bold; color: #1E3A8A; margin-bottom: 5px; }
.sub-title { font-size: 15px; color: #4B5563; margin-bottom: 20px; }
textarea { font-size: 14px !important; }
.stButton>button { width: 100%; border-radius: 8px; }
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📘 RPM CERDAS AI v2.5 (Online)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Penyusun Rencana Pembelajaran Mendalam Kurikulum Merdeka Terintegrasi Google Gemini AI</div>', unsafe_allow_html=True)
st.divider()

# Manajemen API Key di Sidebar untuk Akses Publik Gratis
st.sidebar.title("🔐 Konfigurasi Server AI")
st.sidebar.markdown(
    "Aplikasi ini menggunakan **Google Gemini API**. "
    "Setiap guru dapat memasukkan API Key gratis mereka sendiri untuk mulai menggunakannya secara mandiri."
)

api_key_input = st.sidebar.text_input(
    "Masukkan Gemini API Key Anda:",
    type="password",
    help="Dapatkan API Key gratis di Google AI Studio (https://google.com)"
)

st.sidebar.divider()
if api_key_input:
    st.sidebar.success("Status: API Key Terpasang & Siap Digunakan")
else:
    st.sidebar.warning("Status: Menunggu API Key untuk Mengaktifkan Mesin AI")

def call_gemini_ai(topik, subtopik, tujuan, cp, api_key):
    """
    Menghubungi Google Gemini API untuk menghasilkan komponen Kurikulum Merdeka 
    secara mendalam, orisinal, dan berbasis HOTS dalam format data JSON terstruktur.
    """
    if not api_key:
        return None
        
    try:
        # Inisialisasi klien dengan SDK google-genai terbaru
        client = genai.Client(api_key=api_key)
        
        prompt_sistem = (
            "Anda adalah pakar kurikulum dan pengembang perangkat pembelajaran Kurikulum Merdeka "
            "Kementerian Pendidikan Indonesia. Tugas Anda adalah menganalisis Topik, Subtopik, "
            "Capaian Pembelajaran (CP), dan Tujuan Pembelajaran yang dikirimkan oleh pengguna, "
            "lalu menghasilkan rancangan Rencana Pembelajaran Mendalam (RPM) yang komprehensif, "
            "kontekstual, berorientasi HOTS (High Order Thinking Skills), serta terperinci.\n\n"
            "Wajib merespons dalam struktur JSON murni dengan key yang tepat seperti contoh di bawah ini tanpa komentar di luar JSON:\n"
            "{\n"
            "  \"karakteristik\": \"deskripsi mendalam karakteristik siswa adaptif terhadap topik\",\n"
            "  \"dimensi\": [\"Dimensi 1 + penjelasan kontekstual\", \"Dimensi 2 + penjelasan\"],\n"
            "  \"pendahuluan\": [\"langkah 1\", \"langkah 2\", \"langkah 3\", \"langkah 4\"],\n"
            "  \"inti\": [\"langkah 1 berbasis PBL/HOTS\", \"langkah 2\", \"langkah 3\", \"langkah 4\", \"langkah 5\"],\n"
            "  \"penutup\": [\"langkah 1\", \"langkah 2\", \"langkah 3\", \"langkah 4\"],\n"
            "  \"diagnostik\": \"detail teknis instrumen diagnostik awal materi ini\",\n"
            "  \"formatif\": \"detail rubrik/aktivitas formatif proses diskusi\",\n"
            "  \"sumatif\": \"detail tugas akhir atau tes esai sumatif\",\n"
            "  \"lkpd\": [\"soal analisis konsep HOTS\", \"studi kasus nyata\", \"soal problem solving kreatif\"],\n"
            "  \"rubrik\": [\"rubrik penilaian sikap beserta indikatornya\", \"rubrik pengetahuan\", \"rubrik presentasi\"]\n"
            "}"
        )
        
        prompt_pengguna = f"""
        Analisis data pembelajaran berikut:
        - Topik Utama: {topik}
        - Sub Topik: {subtopik}
        - Capaian Pembelajaran (CP): {cp}
        - Tujuan Pembelajaran: {tujuan}
        
        Tolong hasilkan seluruh isian Rencana Pembelajaran Mendalam sesuai format JSON tersebut secara kreatif, lengkap, dan tidak normatif/potong-potong.
        """
        
        response = client.models.generate_content(
            model='gemini-1.5-flash',
            contents=prompt_pengguna,
            config=types.GenerateContentConfig(
                system_instruction=prompt_sistem,
                response_mime_type="application/json",
                temperature=0.3
            )
        )
        
        # Parse output teks JSON ke objek Python dictionary
        return json.loads(response.text)
        
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data ke server AI: {str(e)}")
        return None

def generate_text_preview(sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, tujuan, data_ai):
    """Membuat format teks mentah untuk pratinjau langsung di layar aplikasi."""
    hasil = f"""============================================================
RENCANA PEMBELAJARAN MENDALAM (RPM) - FORMULASI AI
============================================================

A. IDENTITAS PEMBELAJARAN
Nama Sekolah      : {sekolah}
Nama Guru         : {guru}
Mata Pelajaran    : {mapel}
Kelas / Fase      : {kelas} / {fase}
Semester          : {semester}
Tahun Pelajaran   : {tahun}
Alokasi Waktu     : {alokasi}

============================================================
B. IDENTIFIKASI PEMBELAJARAN
Topik Utama       : {topik}
Sub Topik         : {subtopik}
Capaian (CP)      : {cp}

============================================================
C. KARAKTERISTIK PESERTA DIDIK
{data_ai.get('karakteristik', '')}

============================================================
D. DIMENSI PROFIL PELAJAR PANCASILA
"""
    for d in data_ai.get('dimensi', []):
        hasil += f"• {d}\n"

    hasil += f"""
============================================================
E. TUJUAN PEMBELAJARAN
{tujuan}

============================================================
F. LANGKAH-LANGKAH PEMBELAJARAN (MODEL AKTIF/HOTS)
1. KEGIATAN PENDAHULUAN (15 Menit)
"""
    for p in data_ai.get('pendahuluan', []):
        hasil += f"• {p}\n"

    hasil += "\n2. KEGIATAN INTI (50 Menit)\n"
    for i in data_ai.get('inti', []):
        hasil += f"• {i}\n"

    hasil += "\n3. KEGIATAN PENUTUP (15 Menit)\n"
    for pen in data_ai.get('penutup', []):
        hasil += f"• {pen}\n"

    hasil += f"""
============================================================
G. ASESMEN TRI-PARTIT
• Diagnostik : {data_ai.get('diagnostik', '')}
• Formatif   : {data_ai.get('formatif', '')}
• Sumatif    : {data_ai.get('sumatif', '')}

============================================================
H. LEMBAR KERJA PESERTA DIDIK (LKPD)
Mata Pelajaran : {mapel}
Topik          : {topik} ({subtopik})

SOAL / AKTIVITAS DISKUSI:
"""
    for idx, task in enumerate(data_ai.get('lkpd', []), 1):
        hasil += f"{idx}. {task}\n"

    hasil += "\n============================================================\nI. RUBRIK PENILAIAN\n"
    for r in data_ai.get('rubrik', []):
        hasil += f"• {r}\n"

    return hasil
# ============================================================
# RPM CERDAS AI v2.5 ONLINE & COMPREHENSIVE GENERATION
# BAGIAN 2: WORD EXPORTER (TERMASUK TABEL TANDA TANGAN RAPI)
# ============================================================

def set_cell_background(cell, fill_color):
    """Memberi warna latar belakang pada sel tabel Word."""
    tc_pr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tc_pr.append(shd)

def add_header_row(table, headers, fill_color="1565C0"):
    """Membuat baris header tabel dengan latar berwarna biru."""
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], fill_color)
        p = hdr_cells[i].paragraphs[0]
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

def export_word(sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, tujuan, data_ai, nama_kepsek, nip_kepsek, nip_guru, tempat_tanggal):
    """Menyusun seluruh komponen hasil analisis AI ke dalam dokumen Microsoft Word berbasis tabel."""
    doc = Document()
    
    # Konfigurasi Margin Halaman Dokumen
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # JUDUL UTAMA DOKUMEN
    title = doc.add_heading(level=0)
    run_title = title.add_run("RENCANA PEMBELAJARAN MENDALAM (RPM)")
    run_title.font.size = Pt(18)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(21, 101, 192)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if not karakteristik.strip():
        karakter = f"Peserta didik kelas {kelas} memiliki profil belajar variatif. Pembelajaran diakomodasi melalui diferensiasi proses berbasis pemecahan masalah konkret terhadap materi {topik}."
    else:
        karakter = karakteristik

    # --- 1. TABEL IDENTITAS ---
    doc.add_heading("I. IDENTITAS & IDENTIFIKASI PEMBELAJARAN", level=2)
    table_id = doc.add_table(rows=10, cols=2)
    table_id.style = "Table Grid"
    
    data_id = [
        ("Nama Sekolah", sekolah),
        ("Nama Guru", guru),
        ("Mata Pelajaran", mapel),
        ("Kelas / Fase / Sem", f"{kelas} / Fase {fase} / Semester {semester}"),
        ("Tahun Pelajaran", tahun),
        ("Alokasi Waktu", alokasi),
        ("Topik Utama", topik),
        ("Sub Topik", subtopik),
        ("Capaian Pembelajaran", cp),
        ("Karakteristik Siswa (AI)", data_ai.get('karakteristik', ''))
    ]
    
    for idx, (label, val) in enumerate(data_id):
        row = table_id.rows[idx]
        row.cells[0].text = label
        row.cells[1].text = val
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        set_cell_background(row.cells[0], "F2F2F2")

    doc.add_paragraph()

    # --- 2. TABEL PROFIL LULUSAN & TUJUAN ---
    doc.add_heading("II. ORIENTASI PROFIL & TUJUAN PEMBELAJARAN", level=2)
    table_goal = doc.add_table(rows=2, cols=2)
    table_goal.style = "Table Grid"
    add_header_row(table_goal, ["Dimensi Profil Pelajar Pancasila", "Tujuan Pembelajaran Sesuai Target"])
    
    row_g = table_goal.rows[1]
    p_dim = row_g.cells[0].paragraphs[0]
    for d in data_ai.get('dimensi', []):
        p_dim.add_run(f"• {d}\n")
    row_g.cells[1].text = tujuan

    doc.add_paragraph()

    # --- 3. TABEL LANGKAH PEMBELAJARAN ---
    doc.add_heading("III. STRUKTUR LANGKAH PEMBELAJARAN MENDALAM", level=2)
    table_steps = doc.add_table(rows=4, cols=3)
    table_steps.style = "Table Grid"
    add_header_row(table_steps, ["Tahap Kegiatan", "Durasi", "Detail Aktivitas Berbasis AI"])
    
    # Kegiatan Pendahuluan
    r_pend = table_steps.rows[1]
    r_pend.cells[0].text = "Kegiatan Pendahuluan"
    r_pend.cells[1].text = "15 Menit"
    p_pend = r_pend.cells[2].paragraphs[0]
    for p in data_ai.get('pendahuluan', []):
        p_pend.add_run(f"• {p}\n")
        
    # Kegiatan Inti
    r_inti = table_steps.rows[2]
    r_inti.cells[0].text = "Kegiatan Inti (Model HOTS/PBL)"
    r_inti.cells[1].text = "50 Menit"
    p_inti = r_inti.cells[2].paragraphs[0]
    for i in data_ai.get('inti', []):
        p_inti.add_run(f"• {i}\n")
        
    # Kegiatan Penutup
    r_pen = table_steps.rows[3]
    r_pen.cells[0].text = "Kegiatan Penutup & Refleksi"
    r_pen.cells[1].text = "15 Menit"
    p_pen = r_pen.cells[2].paragraphs[0]
    for pen in data_ai.get('penutup', []):
        p_pen.add_run(f"• {pen}\n")

    doc.add_paragraph()

    # --- 4. TABEL EVALUASI, LKPD, & RUBRIK ---
    doc.add_heading("IV. EVALUASI, LKPD, & RUBRIK PENILAIAN OBJEKTIF", level=2)
    table_eval = doc.add_table(rows=4, cols=2)
    table_eval.style = "Table Grid"
    add_header_row(table_eval, ["Komponen Evaluasi", "Rancangan Dokumen Pokok"])
    
    # Asesmen
    r_as = table_eval.rows[1]
    r_as.cells[0].text = "Asesmen Tripartit"
    r_as.cells[0].paragraphs[0].runs[0].font.bold = True
    r_as.cells[1].text = f"1. Diagnostik: {data_ai.get('diagnostik', '')}\n\n2. Formatif: {data_ai.get('formatif', '')}\n\n3. Sumatif: {data_ai.get('sumatif', '')}"
    
    # LKPD
    r_lk = table_eval.rows[2]
    r_lk.cells[0].text = "Lembar Kerja Siswa (LKPD)"
    r_lk.cells[0].paragraphs[0].runs[0].font.bold = True
    p_lk = r_lk.cells[1].paragraphs[0]
    for idx, task in enumerate(data_ai.get('lkpd', []), 1):
        p_lk.add_run(f"{idx}. {task}\n\n")
        
    # Rubrik
    r_rb = table_eval.rows[3]
    r_rb.cells[0].text = "Rubrik Penilaian Kinerja"
    r_rb.cells[0].paragraphs[0].runs[0].font.bold = True
    p_rb = r_rb.cells[1].paragraphs[0]
    for r in data_ai.get('rubrik', []):
        p_rb.add_run(f"• {r}\n\n")

    doc.add_paragraph()
    doc.add_paragraph()

    # --- 5. BAGIAN TANDA TANGAN (MENGGUNAKAN TABEL TANPA GARIS BORDER) ---
    table_ttd = doc.add_table(rows=3, cols=2)
    table_ttd.style = "Table Grid" 
    
    # Menghilangkan border agar tampak bersih seperti ketikan manual biasa
    for row in table_ttd.rows:
        for cell in row.cells:
            tcPr = cell._element.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), 'none')
                tcBorders.append(border)
            tcPr.append(tcBorders)

    # Baris 1: Tempat, Tanggal Dokumen
    cell_tgl = table_ttd.rows[0].cells[1]
    p_tgl = cell_tgl.paragraphs[0]
    p_tgl.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_tgl.add_run(f"{tempat_tanggal}\n")

    # Baris 2: Jabatan
    cell_jabatan_kepsek = table_ttd.rows[1].cells[0]
    p_jkep = cell_jabatan_kepsek.paragraphs[0]
    p_jkep.add_run("Mengetahui,\nKepala Sekolah\n\n\n\n")
    
    cell_jabatan_guru = table_ttd.rows[1].cells[1]
    p_jguru = cell_jabatan_guru.paragraphs[0]
    p_jguru.add_run("Guru Mata Pelajaran\n\n\n\n")

    # Baris 3: Nama Lengkap & NIP
    cell_nama_kepsek = table_ttd.rows[2].cells[0]
    p_nkep = cell_nama_kepsek.paragraphs[0]
    run_nkep = p_nkep.add_run(nama_kepsek)
    run_nkep.font.underline = True
    run_nkep.font.bold = True
    p_nkep.add_run(f"\nNIP. {nip_kepsek if nip_kepsek else '-'}")

    cell_nama_guru = table_ttd.rows[2].cells[1]
    p_nguru = cell_nama_guru.paragraphs[0]
    run_nguru = p_nguru.add_run(guru)
    run_nguru.font.underline = True
    run_nguru.font.bold = True
    p_nguru.add_run(f"\nNIP. {nip_guru if nip_guru else '-'}")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
# ============================================================
# RPM CERDAS AI v2.5 ONLINE & COMPREHENSIVE GENERATION
# BAGIAN 3: FORM INPUT PARAMETER DAN INTERFACES + FITUR ADMINISTRASI
# ============================================================

st.header("⚙️ Form Parameter & Input Rencana Pembelajaran")
st.write("Silakan lengkapi identitas pokok kurikulum, topik utama, serta target capaian pembelajaran Anda:")

# Pembagian Tab agar Input Administrasi Tanda Tangan Terpisah Rapi
tab_identitas, tab_ttd_admin = st.tabs(["📝 Identitas Pembelajaran", "✍️ Pengesahan & Tanda Tangan"])

with tab_identitas:
    col1, col2, col3 = st.columns(3)
    with col1:
        sekolah = st.text_input("Nama Sekolah", value="SMA Negeri 1 Pembelajaran")
        guru = st.text_input("Nama Guru & Gelar *", value="Ahmad Sucipto, S.Pd.")
        tahun = st.text_input("Tahun Pelajaran", value="2026/2027")
    with col2:
        mapel = st.text_input("Mata Pelajaran", value="Informatika / Sains Terpadu")
        kelas = st.text_input("Kelas / Rombel", value="X-A")
        fase = st.text_input("Fase Kurikulum", value="E")
    with col3:
        semester = st.selectbox("Semester", ["1 (Ganjil)", "2 (Genap)"])
        alokasi = st.text_input("Alokasi Waktu Efektif", value="2 JP x 45 Menit")

    st.divider()

    col_topik, col_tujuan = st.columns(2)
    with col_topik:
        topik = st.text_input("Topik Utama (Materi Pokok) *", value="Kecerdasan Buatan (AI)")
        subtopik = st.text_input("Sub Topik Bahasan *", value="Penerapan LLM dalam Pendidikan")
        cp = st.text_area("Capaian Pembelajaran (CP) Resmi *", value="Peserta didik mampu memahami perkembangan teknologi terkini, menganalisis dampak pemanfaatan tools AI secara bijak, dan merancang solusi penyelesaian masalah sehari-hari menggunakan konsep komputasi modern.")

    with col_tujuan:
        tujuan = st.text_area(
            "Tujuan Pembelajaran Target Kurikulum * (Gunakan penomoran)",
            value="1. Menjelaskan konsep dasar Large Language Model (LLM) dengan bahasanya sendiri secara runtut.\n2. Menganalisis keuntungan dan batasan etis penggunaan AI dalam menyusun materi belajar.\n3. Berkolaborasi dalam kelompok kecil untuk merumuskan petunjuk penggunaan AI yang aman di sekolah."
        )

with tab_ttd_admin:
    st.write("Sesuaikan data penandatanganan dokumen perangkat pembelajaran untuk lembar pengesahan:")
    col_admin1, col_admin2 = st.columns(2)
    
    with col_admin1:
        tempat_tanggal = st.text_input("Tempat, Tanggal Dokumen", value="Jakarta, 17 Juli 2026")
        nama_kepsek = st.text_input("Nama Kepala Sekolah & Gelar", value="Dr. H. Supriadi, M.Pd.")
        nip_kepsek = st.text_input("NIP Kepala Sekolah (Kosongkan jika tidak ada)", value="19750824 200003 1 002")
        
    with col_admin2:
        st.write("") # Spacer kilat
        st.write("") 
        nip_guru = st.text_input("NIP Guru Mata Pelajaran (Kosongkan jika tidak ada)", value="19891210 201504 2 003")

st.divider()

if st.button("🚀 Hubungkan ke Server AI & Formulasikan Perangkat Pembelajaran", type="primary"):
    if not api_key_input:
        st.error("Proses Dihentikan! Silakan masukkan Google Gemini API Key Anda di panel sebelah kiri terlebih dahulu untuk dapat menggunakan fasilitas pencarian kecerdasan AI.")
    elif not topik.strip() or not subtopik.strip() or not tujuan.strip() or not cp.strip() or not guru.strip():
        st.error("Gagal! Kolom bertanda (*) seperti Nama Guru, Topik Utama, Sub Topik, CP, dan Tujuan Pembelajaran wajib diisi.")
    else:
        with st.spinner("Mesin kecerdasan buatan sedang melakukan penalaran kurikulum mendalam dan menyelaraskan instrumen evaluasi..."):
            components = call_gemini_ai(topik, subtopik, tujuan, cp, api_key_input)
            
            if components:
                st.session_state.components = components
                st.session_state.hasil = generate_text_preview(
                    sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, tujuan, components
                )
                st.success("Berhasil! Seluruh komponen pembelajaran prasyarat telah dirumuskan secara otentik.")

if st.session_state.hasil and st.session_state.components:
    st.header("📄 Pratinjau Dokumen Hasil Formulasi AI")
    st.text_area("Preview Output Analisis Teks", value=st.session_state.hasil, height=400)
    
    # Proses pembuatan file Word dengan parameter tambahan Tanda Tangan
    word_file = export_word(
        sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, tujuan, 
        st.session_state.components, nama_kepsek, nip_kepsek, nip_guru, tempat_tanggal
    )
    
    st.download_button(
        label="📥 Unduh Dokumen RPM Hasil AI (Format Tabel Word .docx)",
        data=word_file,
        file_name=f"RPM_Mendalam_AI_{topik.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
