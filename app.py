# ============================================================
# RPM CERDAS AI v2.6 PRODUCTION READY (ONLINE PRO)
# BAGIAN 1: KONFIGURASI HALAMAN & ENGINES GOOGLE GEMINI AI
# ============================================================

import io
import json
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Menggunakan pustaka google-generativeai baku yang sangat stabil untuk server publik
import google.generativeai as genai

# Konfigurasi Halaman Utama Streamlit agar responsif dan muat banyak info
st.set_page_config(
    page_title="RPM CERDAS AI - Komunitas Guru",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# State Management agar halaman tidak reset saat guru mengunduh file
if "hasil" not in st.session_state:
    st.session_state.hasil = ""
if "components" not in st.session_state:
    st.session_state.components = None

# Tampilan UI Khas Kurikulum Merdeka (Biru Pendidikan)
st.markdown("""
<style>
.main-title { font-size: 34px; font-weight: bold; color: #1E3A8A; margin-bottom: 2px; }
.sub-title { font-size: 15px; color: #4B5563; margin-bottom: 15px; }
textarea { font-size: 14px !important; }
.stButton>button { width: 100%; border-radius: 8px; font-weight: bold; }
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📘 RPM CERDAS AI v2.6 (Edisi Berbagi Guru)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Aplikasi Penyusun Rencana Pembelajaran Mendalam Kurikulum Merdeka Terintegrasi Kecerdasan AI</div>', unsafe_allow_html=True)
st.divider()

# Sidebar untuk Panduan Guru Lain
st.sidebar.title("🔐 Akses Server AI Guru")
st.sidebar.markdown(
    "Aplikasi ini didesain gratis untuk dibagikan. "
    "Agar server tidak kelebihan beban (*overload*), bapak/ibu guru disarankan menggunakan "
    "**Gemini API Key** pribadi."
)

api_key_input = st.sidebar.text_input(
    "Masukkan Gemini API Key Anda:",
    type="password",
    help="Dapatkan kunci API 100% gratis di Google AI Studio menggunakan akun Gmail Anda."
)

st.sidebar.markdown("[👉 Ambil API Key Gratis Di Sini](https://google.com)")
st.sidebar.divider()

if api_key_input:
    st.sidebar.success("✅ Mesin AI Siap Bekerja")
else:
    st.sidebar.warning("⚠️ Menunggu Pengisian API Key")

def call_gemini_ai(topik, subtopik, tujuan, cp, api_key):
    """Menghubungi Google Gemini API untuk melakukan penalaran kurikulum murni."""
    if not api_key:
        return None
        
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt_sistem = (
            "Anda adalah konsultan kurikulum senior Kemendikbudristek Indonesia. "
            "Tugas Anda menganalisis Topik, Subtopik, CP, dan Tujuan Pembelajaran, "
            "lalu menyusun Rencana Pembelajaran Mendalam (RPM) komprehensif, logis, dan saling sinergi.\n\n"
            "Wajib memberikan jawaban dalam bentuk format JSON murni tanpa hiasan teks lain dengan struktur berikut:\n"
            "{\n"
            "  \"karakteristik\": \"deskripsi analisis gaya belajar dan kesiapan siswa terkait materi ini\",\n"
            "  \"dimensi\": [\"Dimensi Pancasila 1 (Alasan kontekstual)\", \"Dimensi Pancasila 2 (Alasan kontekstual)\"],\n"
            "  \"pendahuluan\": [\"Langkah orientasi kelas\", \"Aktivitas apersepsi materi\", \"Penyampaian Pertanyaan Pemantik HOTS\", \"Penyampaian tujuan\"],\n"
            "  \"inti\": [\"Orientasi masalah nyata terkait topik\", \"Pengorganisasian kelompok heterogen\", \"Penyelidikan mandiri/kelompok\", \"Presentasi hasil karya/solusi LKPD\", \"Evaluasi proses dan penguatan konsep oleh guru\"],\n"
            "  \"penutup\": [\"Refleksi mendalam siswa\", \"Kesimpulan bersama guru\", \"Aktivitas tindak lanjut\", \"Doa penutup\"],\n"
            "  \"diagnostik\": \"Bentuk kuis awal/pertanyaan pemantik lisan pembuka materi\",\n"
            "  \"formatif\": \"Teknik observasi jalannya diskusi kelompok dan kualitas pengerjaan LKPD\",\n"
            "  \"sumatif\": \"Teknik penilaian produk solusi nyata atau soal tes esai HOTS akhir bab\",\n"
            "  \"lkpd\": [\"Pertanyaan HOTS Tingkat Analisis Konsep\", \"Studi Kasus Konkrit Lingkungan Sekitar\", \"Tantangan Pemecahan Masalah & Rekomendasi Solusi Kelompok\"],\n"
            "  \"rubrik\": [\"Rubrik penilaian sikap ilmiah\", \"Rubrik kedalaman analisis jawaban LKPD\", \"Rubrik performa presentasi publik\"]\n"
            "}"
        )
        
        prompt_pengguna = f"Topik: {topik}\nSubtopik: {subtopik}\nCP: {cp}\nTujuan Pembelajaran:\n{tujuan}"
        
        response = model.generate_content(
            contents=prompt_pengguna,
            generation_config={"response_mime_type": "application/json", "temperature": 0.3}
        )
        
        return json.loads(response.text)
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data ke server AI: {str(e)}")
        return None

def generate_text_preview(sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, tujuan, data_ai):
    """Menghasilkan pratinjau teks di layar aplikasi web."""
    hasil = f"""============================================================
RENCANA PEMBELAJARAN MENDALAM (RPM) - FORMULASI AI
============================================================
A. IDENTITAS PEMBELAJARAN
Sekolah/Guru : {sekolah} / {guru}
Mapel / Kelas : {mapel} / {kelas} ({fase}) - Sem {semester} - TP {tahun}
Alokasi Waktu : {alokasi}

B. IDENTIFIKASI PEMBELAJARAN
Topik / Sub   : {topik} / {subtopik}
Capaian (CP)  : {cp}

C. KARAKTERISTIK SISWA
{data_ai.get('karakteristik', '')}

D. PROFIL PELAJAR PANCASILA
"""
    for d in data_ai.get('dimensi', []):
        hasil += f"• {d}\n"
    hasil += f"\nE. TUJUAN PEMBELAJARAN\n{tujuan}\n\nF. STRUKTUR KEGIATAN"
    hasil += "\n[Pendahuluan]\n" + "\n".join([f"- {x}" for x in data_ai.get('pendahuluan', [])])
    hasil += "\n\n[Inti]\n" + "\n".join([f"- {x}" for x in data_ai.get('inti', [])])
    hasil += "\n\n[Penutup]\n" + "\n".join([f"- {x}" for x in data_ai.get('penutup', [])])
    return hasil
# ============================================================
# RPM CERDAS AI v2.6 PRODUCTION READY (ONLINE PRO)
# BAGIAN 2: ENGINE PEMBUATAN TABEL DAN STRUKTUR MICROSOFT WORD
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
    """Fungsi krusial ekspor dokumen berbentuk tabel tanpa bug tuple."""
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # JUDUL DOKUMEN
    title = doc.add_heading(level=0)
    run_title = title.add_run("RENCANA PEMBELAJARAN MENDALAM (RPM)")
    run_title.font.size = Pt(18)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(21, 101, 192)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # I. TABEL IDENTITAS
    doc.add_heading("I. IDENTITAS & IDENTIFIKASI PEMBELAJARAN", level=2)
    table_id = doc.add_table(rows=10, cols=2)
    table_id.style = "Table Grid"
    
    data_id = [
        ("Nama Sekolah", sekolah), ("Nama Guru", guru), ("Mata Pelajaran", mapel),
        ("Kelas / Fase / Sem", f"{kelas} / Fase {fase} / Semester {semester}"),
        ("Tahun Pelajaran", tahun), ("Alokasi Waktu", alokasi),
        ("Topik Utama", topik), ("Sub Topik", subtopik),
        ("Capaian Pembelajaran", cp), ("Karakteristik Siswa (AI)", data_ai.get('karakteristik', ''))
    ]
    
    for idx, (label, val) in enumerate(data_id):
        row = table_id.rows[idx]
        row.cells[0].text = label
        row.cells[1].text = val
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        set_cell_background(row.cells[0], "F2F2F2")

    doc.add_paragraph()

    # II. TABEL PROFIL & TUJUAN
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

    # III. TABEL LANGKAH PEMBELAJARAN
    doc.add_heading("III. STRUKTUR LANGKAH PEMBELAJARAN MENDALAM", level=2)
    table_steps = doc.add_table(rows=4, cols=3)
    table_steps.style = "Table Grid"
    add_header_row(table_steps, ["Tahap Kegiatan", "Durasi", "Detail Aktivitas Berbasis AI"])
    
    tahapan = [
        ("Kegiatan Pendahuluan", "15 Menit", data_ai.get('pendahuluan', [])),
        ("Kegiatan Inti (PBL/HOTS)", "50 Menit", data_ai.get('inti', [])),
        ("Kegiatan Penutup & Refleksi", "15 Menit", data_ai.get('penutup', []))
    ]
    
    for idx, (t_name, dur, acts) in enumerate(tahapan, 1):
        row = table_steps.rows[idx]
        row.cells[0].text = t_name
        row.cells[1].text = dur
        p_act = row.cells[2].paragraphs[0]
        for a in acts:
            p_act.add_run(f"• {a}\n")

    doc.add_paragraph()

    # IV. TABEL EVALUASI, LKPD, & RUBRIK
    doc.add_heading("IV. EVALUASI, LKPD, & RUBRIK PENILAIAN OBJEKTIF", level=2)
    table_eval = doc.add_table(rows=4, cols=2)
    table_eval.style = "Table Grid"
    add_header_row(table_eval, ["Komponen Evaluasi", "Rancangan Dokumen Pokok"])
    
    # Isi Asesmen
    table_eval.rows[1].cells[0].text = "Asesmen Tripartit"
    table_eval.rows[1].cells[0].paragraphs[0].runs[0].font.bold = True
    table_eval.rows[1].cells[1].text = f"1. Diagnostik: {data_ai.get('diagnostik', '')}\n\n2. Formatif: {data_ai.get('formatif', '')}\n\n3. Sumatif: {data_ai.get('sumatif', '')}"
    
    # Isi LKPD
    table_eval.rows[2].cells[0].text = "Lembar Kerja Siswa (LKPD)"
    table_eval.rows[2].cells[0].paragraphs[0].runs[0].font.bold = True
    p_lk = table_eval.rows[2].cells[1].paragraphs[0]
    for i, t in enumerate(data_ai.get('lkpd', []), 1):
        p_lk.add_run(f"{i}. {t}\n\n")
        
    # Isi Rubrik
    table_eval.rows[3].cells[0].text = "Rubrik Penilaian Kinerja"
    table_eval.rows[3].cells[0].paragraphs[0].runs[0].font.bold = True
    p_rb = table_eval.rows[3].cells[1].paragraphs[0]
    for r in data_ai.get('rubrik', []):
        p_rb.add_run(f"• {r}\n\n")

    doc.add_paragraph()
    doc.add_paragraph()

    # V. BAGIAN TANDA TANGAN (FORMULASI INDEKS SEL PRESTISIUS)
    table_ttd = doc.add_table(rows=3, cols=2)
    
    for row in table_ttd.rows:
        for cell in row.cells:
            tcPr = cell._element.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), 'none')
                tcBorders.append(border)
            tcPr.append(tcBorders)

    # Isi Data Tanda Tangan
    table_ttd.rows[0].cells[1].text = f"{tempat_tanggal}\n"
    table_ttd.rows[1].cells[0].text = "Mengetahui,\nKepala Sekolah\n\n\n\n"
    table_ttd.rows[1].cells[1].text = "Guru Mata Pelajaran\n\n\n\n"

    # Format Nama Kepala Sekolah
    p_nkep = table_ttd.rows[2].cells[0].paragraphs[0]
    run_nkep = p_nkep.add_run(nama_kepsek)
    run_nkep.font.underline = True
    run_nkep.font.bold = True
    p_nkep.add_run(f"\nNIP. {nip_kepsek if nip_kepsek else '-'}")

    # Format Nama Guru
    p_nguru = table_ttd.rows[2].cells[1].paragraphs[0]
    run_nguru = p_nguru.add_run(guru)
    run_nguru.font.underline = True
    run_nguru.font.bold = True
    p_nguru.add_run(f"\nNIP. {nip_guru if nip_guru else '-'}")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
# ============================================================
# RPM CERDAS AI v2.6 PRODUCTION READY (ONLINE PRO)
# BAGIAN 3: SELEKTOR ANTARMUKA, FORM INPUT, & PROSES DOWNLOAD
# ============================================================

# Pembagian data input guru menggunakan Tab modern
tab_identitas, tab_ttd_admin = st.tabs(["📝 Identitas Pembelajaran", "✍️ Administrasi Tanda Tangan"])

with tab_identitas:
    col1, col2, col3 = st.columns(3)
    with col1:
        sekolah = st.text_input("Nama Sekolah", value="SMA Negeri 1 Guru Berbagi")
        guru = st.text_input("Nama Guru & Gelar *", value="Ahmad Sucipto, S.Pd.")
        tahun = st.text_input("Tahun Pelajaran", value="2026/2027")
    with col2:
        mapel = st.text_input("Mata Pelajaran", value="Informatika / Sains Terpadu")
        kelas = st.text_input("Kelas / Rombel", value="X-A")
        fase = st.text_input("Fase", value="E")
    with col3:
        semester = st.selectbox("Semester", ["1 (Ganjil)", "2 (Genap)"])
        alokasi = st.text_input("Alokasi Waktu", value="2 JP x 45 Menit")

    st.divider()

    col_topik, col_tujuan = st.columns(2)
    with col_topik:
        topik = st.text_input("Topik Utama (Materi Pokok) *", value="Kecerdasan Buatan (AI)")
        subtopik = st.text_input("Sub Topik Bahasan *", value="Penerapan LLM dalam Pendidikan")
        cp = st.text_area("Capaian Pembelajaran (CP) Resmi *", value="Peserta didik mampu memahami perkembangan teknologi terkini, menganalisis dampak pemanfaatan tools AI secara bijak, dan merancang solusi penyelesaian masalah sehari-hari menggunakan konsep komputasi modern.")

    with col_tujuan:
        tujuan = st.text_area(
            "Tujuan Pembelajaran Target Kurikulum * (Tulis Berurutan)",
            value="1. Menjelaskan konsep dasar Large Language Model (LLM) dengan bahasanya sendiri.\n2. Menganalisis batasan etis penggunaan AI dalam menyusun materi belajar.\n3. Merumuskan petunjuk penggunaan AI yang aman di sekolah."
        )

with tab_ttd_admin:
    st.write("Isi kelengkapan dokumen pengesahan dinas berikut:")
    col_admin1, col_admin2 = st.columns(2)
    
    with col_admin1:
        tempat_tanggal = st.text_input("Tempat, Tanggal Dokumen", value="Jakarta, 28 Juli 2026")
        nama_kepsek = st.text_input("Nama Kepala Sekolah", value="Dr. H. Supriadi, M.Pd.")
        nip_kepsek = st.text_input("NIP Kepala Sekolah", value="19750824 200003 1 002")
        
    with col_admin2:
        st.write("") # Spacer vertikal layout
        st.write("")
        nip_guru = st.text_input("NIP Guru Mata Pelajaran", value="19891210 201504 2 003")

st.divider()

# Logika Tombol Formulasi Utama
if st.button("🚀 Hubungkan ke Server AI & Susun Modul Pembelajaran", type="primary"):
    if not api_key_input:
        st.error("Proses Ditolak! Masukkan Gemini API Key di menu sidebar sebelah kiri terlebih dahulu untuk dapat menggunakan fasilitas publik ini.")
    elif not topik.strip() or not subtopik.strip() or not tujuan.strip() or not cp.strip() or not guru.strip():
        st.error("Gagal! Mohon lengkapi semua kolom input yang bertanda bintang (*).")
    else:
        with st.spinner("Server AI sedang melakukan penelusuran pedagogi dan sinkronisasi instrumen..."):
            components = call_gemini_ai(topik, subtopik, tujuan, cp, api_key_input)
            
            if components:
                st.session_state.components = components
                st.session_state.hasil = generate_text_preview(
                    sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, tujuan, components
                )
                st.success("Selesai! Komponen Rencana Pembelajaran Berhasil dirancang.")

# Menampilkan hasil dokumen serta tombol download jika proses generate berhasil
if st.session_state.hasil and st.session_state.components:
    st.header("📄 Pratinjau Dokumen Hasil Formulasi AI")
    st.text_area("Preview Struktur Teks", value=st.session_state.hasil, height=350)
    
    # Generator file Word yang aman dari error tuple karena indeks sel [0] & [1] sudah dikunci
    word_file = export_word(
        sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, tujuan, 
        st.session_state.components, nama_kepsek, nip_kepsek, nip_guru, tempat_tanggal
    )
    
    st.download_button(
        label="📥 Unduh Dokumen RPM Resmi Berformat Tabel (.docx)",
        data=word_file,
        file_name=f"RPM_Merdeka_AI_{topik.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
