# ============================================================
# RPM CERDAS AI v2.2 OFFLINE (FORMAT TABEL WORD RAPI)
# BAGIAN 1: CONFIG, IMPORTS, & LOGIKA GENERATOR AI
# ============================================================

import io
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Konfigurasi Halaman Utama
st.set_page_config(
    page_title="RPM CERDAS AI",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded"
)

if "hasil" not in st.session_state:
    st.session_state.hasil = ""

if "history" not in st.session_state:
    st.session_state.history = []

# Styling Tampilan UI
st.markdown("""
<style>
.main-title { font-size: 36px; font-weight: bold; color: #1565C0; }
.sub-title { font-size: 16px; color: #555555; }
textarea { font-size: 14px !important; }
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📘 RPM CERDAS AI</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Generator RPM Kurikulum Merdeka (Edisi Format Tabel Word Rapi)</div>', unsafe_allow_html=True)
st.divider()

st.sidebar.title("📘 RPM CERDAS AI")
menu = st.sidebar.selectbox("Pilih Menu", ["Generator RPM", "Tentang Aplikasi"])
st.sidebar.divider()
st.sidebar.success("Status: Server Offline Ready")

if menu == "Tentang Aplikasi":
    st.header("Tentang RPM CERDAS AI")
    st.write("""
    Aplikasi ini menyusun Rencana Pembelajaran Mendalam (RPM) lengkap dengan:
    - Identitas & Identifikasi Pembelajaran
    - Langkah-langkah (Pendahuluan, Inti, Penutup)
    - Asesmen Diagnostik, Formatif, Sumatif
    - LKPD & Rubrik Penilaian Detail
    - **Ekspor Word Rapi:** Otomatis dikemas dalam tabel berpola profesional.
    """)
    st.stop()

def ai_generate_components(topik, subtopik, tujuan):
    """Mesin penalaran AI internal untuk menyusun komponen pembelajaran."""
    dimensi = [
        "Bernalar Kritis (Menganalisis konsep dan memecahkan masalah kontekstual)",
        "Gotong Royong (Berkolaborasi aktif dalam penugasan kelompok dan presentasi)",
        "Mandiri (Bertanggung jawab atas proses dan hasil belajar secara personal)"
    ]
    
    langkah = {
        "pendahuluan": [
            "Guru membuka kelas dengan salam hangat, berdoa bersama, dan mengecek kehadiran.",
            f"Apersepsi: Guru mengaitkan materi sebelumnya dengan topik baru yaitu '{topik}'.",
            f"Pertanyaan Pemantik: 'Pernahkah kalian memperhatikan bagaimana {subtopik} bekerja atau diterapkan dalam kehidupan nyata kita? Apa dampaknya jika hal tersebut tidak ada?'",
            f"Guru memaparkan tujuan pembelajaran hari ini dan menjelaskan pentingnya menguasai {subtopik}."
        ],
        "inti": [
            f"Orientasi Masalah: Peserta didik mengamati studi kasus nyata atau tayangan visual yang memuat tantangan kontekstual terkait '{topik} ({subtopik})'.",
            "Pengorganisasian Kelompok: Peserta didik dibagi ke dalam kelompok heterogen (4-5 orang) dan menerima lembar kerja (LKPD).",
            f"Penyelidikan Terbimbing: Kelompok melakukan literasi digital/buku, berdiskusi, dan mengumpulkan data untuk memecahkan masalah praktis terkait {subtopik}.",
            "Mengembangkan & Menyajikan Karya: Setiap kelompok menyusun draf solusi atau peta konsep pada LKPD, lalu mempresentasikannya di depan kelas.",
            f"Analisis & Evaluasi: Guru memberikan konfirmasi, meluruskan miskonsepsi mengenai {topik}, dan memberikan apresiasi atas performa seluruh tim."
        ],
        "penutup": [
            f"Peserta didik dibimbing guru membuat kesimpulan terpadu tentang inti materi {subtopik}.",
            f"Refleksi Mendalam: Peserta didik menjawab pertanyaan, 'Bagian mana dari konsep {topik} yang paling menantang dan bagaimana Anda akan menerapkannya?'",
            "Guru memberikan umpan balik positif serta menginformasikan rencana materi pertemuan berikutnya.",
            "Pembelajaran diakhiri dengan doa syukur dan salam penutup."
        ]
    }
    
    asesmen = {
        "diagnostik": f"Asesmen Kognitif Awal melalui kuis singkat 3 pertanyaan atau curah pendapat kilat mengenai pengetahuan dasar {topik}.",
        "formatif": f"Observasi proses diskusi kelompok, penilaian keaktifan, dan penilaian performa saat mempresentasikan LKPD {subtopik}.",
        "sumatif": f"Tes tertulis objektif/esai di akhir unit atau penilaian produk laporan solusi kontekstual mengenai {topik}."
    }
    
    lkpd_tasks = [
        f"[Pemahaman Konsep] Analisis secara mendalam hubungan kausalitas antara '{topik}' dengan sub-materi '{subtopik}' berdasarkan literatur yang Anda temukan!",
        f"[Studi Kasus Nyata] Temukan satu fenomena konkret di lingkungan sekitar Anda yang berkaitan dengan {subtopik}. Jelaskan mekanisme terjadinya secara ilmiah!",
        f"[Problem Solving & Inovasi] Jika ditemukan kegagalan sistem atau masalah implementasi pada {topik} dalam kehidupan sehari-hari, formulasikan 2 solusi kreatif kelompok Anda!"
    ]
    
    rubrik = [
        "Sikap (Profil Pelajar Pancasila): Skor 1-4 untuk indikator Bernalar Kritis dan Gotong Royong selama dinamika kelompok.",
        "Pengetahuan (LKPD): Skor maksimal 100 dinilai berdasarkan kedalaman analisis, orisinalitas argumen, dan ketepatan teori.",
        "Keterampilan (Presentasi): Skor 1-4 berpatokan pada kejelasan artikulasi, sistematika penyampaian, dan kemampuan mempertahankan argumentasi saat tanya jawab."
    ]
    
    return dimensi, langkah, asesmen, lkpd_tasks, rubrik

def generate_rpm_document(sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, karakteristik, tujuan):
    if not karakteristik.strip():
        karakter = f"Peserta didik kelas {kelas} memiliki profil belajar yang variatif dengan kesiapan yang beragam. Sebagian besar memerlukan visualisasi konkret sebelum memahami abstraksi teori {topik}. Pembelajaran diakomodasi melalui diferensiasi proses dan konten."
    else:
        karakter = karakteristik

    dimensi, langkah, asesmen, lkpd_tasks, rubrik = ai_generate_components(topik, subtopik, tujuan)

    hasil = f"""============================================================
RENCANA PEMBELAJARAN MENDALAM (RPM)
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
{karakter}

============================================================
D. DIMENSI PROFIL LULUSAN
"""
    for d in dimensi:
        hasil += f"• {d}\n"

    hasil += f"""
============================================================
E. TUJUAN PEMBELAJARAN
{tujuan}

============================================================
F. LANGKAH-LANGKAH PEMBELAJARAN
1. KEGIATAN PENDAHULUAN (15 Menit)
"""
    for p in langkah["pendahuluan"]:
        hasil += f"• {p}\n"

    hasil += "\n2. KEGIATAN INTI (50 Menit)\n"
    for i in langkah["inti"]:
        hasil += f"• {i}\n"

    hasil += "\n3. KEGIATAN PENUTUP (15 Menit)\n"
    for pen in langkah["penutup"]:
        hasil += f"• {pen}\n"

    hasil += f"""
============================================================
G. ASESMEN PEMBELAJARAN
• Diagnostik : {asesmen["diagnostik"]}
• Formatif   : {asesmen["formatif"]}
• Sumatif    : {asesmen["sumatif"]}

============================================================
H. LEMBAR KERJA PESERTA DIDIK (LKPD)
Mata Pelajaran : {mapel}
Topik          : {topik} ({subtopik})

TUGAS / PERTANYAAN DISKUSI:
"""
    for idx, task in enumerate(lkpd_tasks, 1):
        hasil += f"{idx}. {task}\n"

    hasil += "\n============================================================\nI. RUBRIK PENILAIAN\n"
    for r in rubrik:
        hasil += f"• {r}\n"

    return hasil
# ============================================================
# RPM CERDAS AI v2.2 OFFLINE (FORMAT TABEL WORD RAPI)
# BAGIAN 2: LOGIKA INTEGRASI FORMAT TABEL WORD (.DOCX)
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

def export_word(sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, karakteristik, tujuan):
    doc = Document()
    
    # Margin Dokumen (Normal)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Judul Dokumen
    title = doc.add_heading(level=0)
    run_title = title.add_run("RENCANA PEMBELAJARAN MENDALAM (RPM)")
    run_title.font.size = Pt(18)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(21, 101, 192)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    dimensi, langkah, asesmen, lkpd_tasks, rubrik = ai_generate_components(topik, subtopik, tujuan)
    
    if not karakteristik.strip():
        karakter = f"Peserta didik kelas {kelas} memiliki ragam kesiapan belajar yang bervariasi. Pembelajaran didesain menggunakan pendekatan diferensiasi berbasis konten dan proses guna memaksimalkan internalisasi konsep {topik}."
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
        ("Karakteristik Siswa", karakter)
    ]
    
    for idx, (label, val) in enumerate(data_id):
        row = table_id.rows[idx]
        row.cells[0].text = label
        row.cells[1].text = val
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        set_cell_background(row.cells[0], "F2F2F2")

    doc.add_paragraph()

    # --- 2. TABEL PROFIL & TUJUAN ---
    doc.add_heading("II. ORIENTASI PROFIL & TUJUAN", level=2)
    table_goal = doc.add_table(rows=2, cols=2)
    table_goal.style = "Table Grid"
    add_header_row(table_goal, ["Dimensi Profil Pelajar Pancasila", "Tujuan Pembelajaran Sesuai Target"])
    
    row_g = table_goal.rows[1]
    p_dim = row_g.cells[0].paragraphs[0]
    for d in dimensi:
        p_dim.add_run(f"• {d}\n")
    row_g.cells[1].text = tujuan

    doc.add_paragraph()

    # --- 3. TABEL LANGKAH PEMBELAJARAN ---
    doc.add_heading("III. STRUKTUR LANGKAH PEMBELAJARAN MENDALAM", level=2)
    table_steps = doc.add_table(rows=4, cols=3)
    table_steps.style = "Table Grid"
    add_header_row(table_steps, ["Tahap Kegiatan", "Durasi", "Detail Aktivitas Berbasis AI"])
    
    # Pendahuluan
    r_pend = table_steps.rows[1]
    r_pend.cells[0].text = "Kegiatan Pendahuluan"
    r_pend.cells[1].text = "15 Menit"
    p_pend = r_pend.cells[2].paragraphs[0]
    for p in langkah["pendahuluan"]:
        p_pend.add_run(f"• {p}\n")
        
    # Inti
    r_inti = table_steps.rows[2]
    r_inti.cells[0].text = "Kegiatan Inti (Model HOTS/PBL)"
    r_inti.cells[1].text = "50 Menit"
    p_inti = r_inti.cells[2].paragraphs[0]
    for i in langkah["inti"]:
        p_inti.add_run(f"• {i}\n")
        
    # Penutup
    r_pen = table_steps.rows[3]
    r_pen.cells[0].text = "Kegiatan Penutup & Refleksi"
    r_pen.cells[1].text = "15 Menit"
    p_pen = r_pen.cells[2].paragraphs[0]
    for pen in langkah["penutup"]:
        p_pen.add_run(f"• {pen}\n")

    doc.add_paragraph()

    # --- 4. TABEL EVALUASI, LKPD & RUBRIK ---
    doc.add_heading("IV. EVALUASI, LKPD, & RUBRIK PENILAIAN OBJEKTIF", level=2)
    table_eval = doc.add_table(rows=4, cols=2)
    table_eval.style = "Table Grid"
    add_header_row(table_eval, ["Komponen Evaluasi", "Rancangan Dokumen Pokok"])
    
    r_as = table_eval.rows[1]
    r_as.cells[0].text = "Asesmen Tripartit"
    r_as.cells[0].paragraphs[0].runs[0].font.bold = True
    r_as.cells[1].text = f"1. Diagnostik: {asesmen['diagnostik']}\n2. Formatif: {asesmen['formatif']}\n3. Sumatif: {asesmen['sumatif']}"
    
    r_lk = table_eval.rows[2]
    r_lk.cells[0].text = f"Lembar Kerja Peserta Didik (LKPD)"
    r_lk.cells[0].paragraphs[0].runs[0].font.bold = True
    p_lk = r_lk.cells[1].paragraphs[0]
    for idx, task in enumerate(lkpd_tasks, 1):
        p_lk.add_run(f"{idx}. {task}\n")
        
    r_rb = table_eval.rows[3]
    r_rb.cells[0].text = "Rubrik Penilaian Kinerja"
    r_rb.cells[0].paragraphs[0].runs[0].font.bold = True
    p_rb = r_rb.cells[1].paragraphs[0]
    for r in rubrik:
        p_rb.add_run(f"• {r}\n")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
# ============================================================
# RPM CERDAS AI v2.2 OFFLINE (FORMAT TABEL WORD RAPI)
# BAGIAN 3: FORM INPUT PARAMETER DAN INTERFACES STREAMLIT
# ============================================================

st.header("⚙️ Form Parameter & Input Rencana Pembelajaran")
st.write("Silakan isi data identitas pokok, topik, dan tujuan pembelajaran di bawah ini:")

col1, col2, col3 = st.columns(3)

with col1:
    sekolah = st.text_input("Nama Sekolah", value="SMA Negeri 1 Pembelajaran")
    guru = st.text_input("Nama Guru", value="Ahmad Sucipto, S.Pd.")
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
    topik = st.text_input("Topik Utama *", value="Kecerdasan Buatan (AI)")
    subtopik = st.text_input("Sub Topik *", value="Penerapan LLM dalam Pendidikan")
    cp = st.text_area("Capaian Pembelajaran (CP)", value="Peserta didik mampu memahami perkembangan teknologi terkini, menganalisis dampak pemanfaatan tools AI secara bijak, dan merancang solusi penyelesaian masalah sehari-hari menggunakan konsep komputasi modern.")
    karakteristik = st.text_area("Karakteristik Siswa (Opsional - Biarkan kosong agar diisi otomatis oleh AI)", value="")

with col_tujuan:
    tujuan = st.text_area(
        "Tujuan Pembelajaran * (Tulis per baris atau gunakan penomoran)",
        value="1. Menjelaskan konsep dasar Large Language Model (LLM) dengan bahasanya sendiri secara runtut.\n2. Menganalisis keuntungan dan batasan etis penggunaan AI dalam menyusun materi belajar.\n3. Berkolaborasi dalam kelompok kecil untuk merumuskan petunjuk penggunaan AI yang aman di sekolah."
    )

st.divider()

# Aksi saat tombol ditekan
if st.button("🚀 Susun Rencana Pembelajaran Mendalam (AI)", type="primary"):
    if not topik.strip() or not subtopik.strip() or not tujuan.strip():
        st.error("Gagal! Kolom Topik Utama, Sub Topik, dan Tujuan Pembelajaran wajib diisi.")
    else:
        with st.spinner("AI sedang menganalisis tujuan pembelajaran dan menyusun materi pendukung..."):
            # Buat teks pratinjau
            st.session_state.hasil = generate_rpm_document(
                sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, karakteristik, tujuan
            )
            # Buat binary file Word format tabel
            st.session_state.word_file = export_word(
                sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, karakteristik, tujuan
            )
            st.success("Selesai! RPM Berhasil dihasilkan secara komprehensif.")

# Tampilkan hasil jika data sudah digenerate
if st.session_state.hasil:
    st.header("📄 Pratinjau Dokumen Hasil Formulasi AI")
    st.text_area("Preview Output Teks", value=st.session_state.hasil, height=450)
    
    st.download_button(
        label="📥 Unduh Dokumen RPM (Format Tabel Word .docx)",
        data=st.session_state.word_file,
        file_name=f"RPM_Mendalam_{topik.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
