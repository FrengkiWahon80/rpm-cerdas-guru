# ============================================================
# RPM CERDAS AI v2.2 OFFLINE (FORMAT TABEL WORD RAPI)
# Aplikasi Penyusun Rencana Pembelajaran Mendalam
# ============================================================

import io
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ============================================================
# KONFIGURASI HALAMAN STREAMLIT
# ============================================================

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

# CSS Custom
st.markdown("""
<style>
.main-title { font-size: 36px; font-weight: bold; color: #1565C0; }
.sub-title { font-size: 16px; color: #555555; }
textarea { font-size: 14px !important; }
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-title">📘 RPM CERDAS AI</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Generator RPM Kurikulum Merdeka (Edisi Format Tabel Word Rapi)</div>', unsafe_allow_html=True)
st.divider()

# Sidebar
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

# ============================================================
# FUNGSI GENERATOR TEKS DOKUMEN RPM
# ============================================================

def generate_rpm_document(sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, karakteristik):
    if not karakteristik.strip():
        karakter = f"Peserta didik kelas {kelas} memiliki beragam gaya belajar (visual, auditori, kinestetik) serta tingkat kesiapan belajar yang bervariasi. Pembelajaran dirancang menggunakan diferensiasi proses dan konten."
    else:
        karakter = karakteristik

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
• Beriman, Bertakwa kepada Tuhan YME, dan Berakhlak Mulia
• Berkebinekaan Global
• Bergotong Royong
• Mandiri
• Bernalar Kritis
• Kreatif

============================================================
E. TUJUAN PEMBELAJARAN
1. Melalui pengamatan dan diskusi, peserta didik mampu menjelaskan konsep dasar {topik} ({subtopik}) dengan tepat.
2. Melalui eksplorasi masalah, peserta didik mampu menganalisis keterkaitan {topik} dengan kehidupan sehari-hari.
3. Melalui penugasan kelompok pada LKPD, peserta didik mampu menyusun dan mempresentasikan solusi secara kolaboratif.
4. Menunjukkan sikap bernalar kritis, gotong royong, dan tanggung jawab selama pembelajaran.

============================================================
F. LANGKAH-LANGKAH PEMBELAJARAN
1. KEGIATAN PENDAHULUAN (15 Menit)
• Salam, berdoa, dan presensi.
• Apersepsi dan apersepsi materi {topik}.
• Pertanyaan pemantik: "Bagaimana penerapan {topik} dalam kehidupan kita sehari-hari?"
• Menyampaikan tujuan pembelajaran dan alur kegiatan.

2. KEGIATAN INTI (50 Menit)
• Orientasi Masalah: Pengamatan tayangan/studi kasus kontekstual terkait {topik}.
• Pengorganisasian Kelompok: Pembagian kelompok heterogen dan pembagian LKPD.
• Penyelidikan Mandiri/Kelompok: Eksplorasi materi, diskusi, dan pengumpulan data.
• Penyajian Karya: Presentasi hasil diskusi kelompok di depan kelas.
• Analisis & Evaluasi: Tanya jawab, penguatan konsep oleh guru, dan klarifikasi miskonsepsi.

3. KEGIATAN PENUTUP (15 Menit)
• Rangkuman dan kesimpulan bersama peserta didik.
• Refleksi pembelajaran dan umpan balik.
• Penugasan mandiri/informasi materi berikutnya, penutup doa dan salam.

============================================================
G. ASESMEN PEMBELAJARAN
• Diagnostik : Tanya jawab / Pertanyaan Pemantik awal.
• Formatif   : Observasi diskusi kelompok dan unjuk kerja LKPD.
• Sumatif    : Tes tertulis / penilaian produk akhir.

============================================================
H. LEMBAR KERJA PESERTA DIDIK (LKPD)
Mata Pelajaran : {mapel}
Topik          : {topik} ({subtopik})

TUGAS / PERTANYAAN DISKUSI:
1. [Pemahaman Konsep] Jelaskan definisi serta prinsip dasar dari {topik}!
2. [Analisis Kasus] Analisis 2 contoh penerapan {subtopik} di lingkungan sekitar!
3. [Penyelesaian Masalah] Rekomendasikan solusi terhadap kendala penerapan {topik}!

============================================================
I. RUBRIK PENILAIAN
1. Penilaian Sikap (Bernalar Kritis, Gotong Royong, Tanggung Jawab)
2. Penilaian Pengetahuan (Kelengkapan & Ketepatan Jawaban LKPD)
3. Penilaian Keterampilan (Penguasaan Materi & Penyampaian Presentasi)
"""
    return hasil

# ============================================================
# FUNGSI EXPORT WORD LENGKAP DALAM FORMAT TABEL RAPI
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

def export_word(sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, karakteristik):
    doc = Document()
    
    # Setting Margin Dokumen (Normal)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # JUDUL UTAMA
    title = doc.add_heading(level=0)
    run_title = title.add_run("RENCANA PEMBELAJARAN MENDALAM (RPM)")
    run_title.font.size = Pt(18)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(21, 101, 192)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- 1. TABEL IDENTITAS & IDENTIFIKASI ---
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
        ("Capaian Pembelajaran (CP)", cp),
        ("Karakteristik Siswa", karakteristik if karakteristik else "Siswa memiliki gaya belajar bervariasi; diterapkan pembelajaran berpusat pada siswa.")
    ]

    for i, (k, v) in enumerate(data_id):
        cell_k = table_id.cell(i, 0)
        cell_v = table_id.cell(i, 1)
        cell_k.text = k
        cell_v.text = v
        set_cell_background(cell_k, "F2F4F7")
        cell_k.paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph()

    # --- 2. TABEL LANGKAH-LANGKAH PEMBELAJARAN ---
    doc.add_heading("II. LANGKAH-LANGKAH PEMBELAJARAN", level=2)
    table_step = doc.add_table(rows=4, cols=3)
    table_step.style = "Table Grid"
    add_header_row(table_step, ["Kegiatan", "Deskripsi Aktivitas", "Waktu"])

    steps = [
        ("Pendahuluan", 
         "1. Salam pembuka, doa bersama, dan mengecek presensi.\n"
         "2. Apersepsi dan apersepsi materi sebelumnya terkait topik.\n"
         "3. Penyampaian pertanyaan pemantik & tujuan pembelajaran.", 
         "15 Menit"),
        ("Kegiatan Inti", 
         f"1. Orientasi Masalah: Mengamati studi kasus kontekstual materi {topik}.\n"
         "2. Pengorganisasian: Pembentukan kelompok & pembagian LKPD.\n"
         "3. Penyelidikan: Diskusi kelompok dan eksplorasi data.\n"
         "4. Penyajian Karya: Presentasi hasil diskusi kelompok.\n"
         "5. Evaluasi: Penguatan konsep oleh guru & penyamaan persepsi.", 
         "50 Menit"),
        ("Penutup", 
         "1. Guru dan siswa merangkum kesimpulan pembelajaran.\n"
         "2. Refleksi pembelajaran dan umpan balik.\n"
         "3. Penugasan mandiri / info materi berikutnya, doa, dan salam.", 
         "15 Menit")
    ]

    for idx, (keg, desk, wkt) in enumerate(steps, start=1):
        table_step.cell(idx, 0).text = keg
        table_step.cell(idx, 1).text = desk
        table_step.cell(idx, 2).text = wkt
        table_step.cell(idx, 0).paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph()

    # --- 3. TABEL ASESMEN PEMBELAJARAN ---
    doc.add_heading("III. ASESMEN PEMBELAJARAN", level=2)
    table_as = doc.add_table(rows=4, cols=3)
    table_as.style = "Table Grid"
    add_header_row(table_as, ["Jenis Asesmen", "Teknik Penilaian", "Instrumen Penilaian"])

    asesmen_data = [
        ("Asesmen Diagnostik", "Tanya Jawab / Wawancara Singkat", "Pertanyaan Pemantik Kesiapan Belajar"),
        ("Asesmen Formatif", "Observasi & Unjuk Kerja", "Lembar Observasi Diskusi & Rubrik LKPD"),
        ("Asesmen Sumatif", "Tes Tertulis / Penilaian Produk", "Soal Evaluasi / Rubrik Laporan")
    ]

    for idx, (jns, tek, ins) in enumerate(asesmen_data, start=1):
        table_as.cell(idx, 0).text = jns
        table_as.cell(idx, 1).text = tek
        table_as.cell(idx, 2).text = ins
        table_as.cell(idx, 0).paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph()

    # --- 4. LEMBAR KERJA PESERTA DIDIK (LKPD) ---
    doc.add_page_break()
    doc.add_heading("IV. LEMBAR KERJA PESERTA DIDIK (LKPD)", level=2)
    
    table_lkpd = doc.add_table(rows=4, cols=2)
    table_lkpd.style = "Table Grid"
    
    table_lkpd.cell(0, 0).text = "Topik / Sub Topik"
    table_lkpd.cell(0, 1).text = f"{topik} / {subtopik}"
    table_lkpd.cell(1, 0).text = "Petunjuk Kerja"
    table_lkpd.cell(1, 1).text = "1. Bacalah materi pendukung dengan cermat.\n2. Diskusikan pertanyaan berikut dalam kelompok.\n3. Tuliskan jawaban pada kolom yang tersedia."
    
    table_lkpd.cell(2, 0).text = "Pertanyaan Diskusi"
    table_lkpd.cell(2, 1).text = (
        f"1. Jelaskan definisi dan konsep dasar dari {topik}!\n"
        f"2. Analisis 2 contoh penerapan {subtopik} di lingkungan sekitarmu!\n"
        f"3. Rekomendasikan solusi terhadap kendala penerapan {topik}!"
    )
    
    table_lkpd.cell(3, 0).text = "Kolom Jawaban & Kesimpulan"
    table_lkpd.cell(3, 1).text = "\n\n\n\n\n"

    for row in table_lkpd.rows:
        set_cell_background(row.cells[0], "F2F4F7")
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph()

    # --- 5. RUBRIK PENILAIAN ---
    doc.add_heading("V. RUBRIK PENILAIAN", level=2)
    
    table_rubrik = doc.add_table(rows=4, cols=5)
    table_rubrik.style = "Table Grid"
    add_header_row(table_rubrik, ["Aspek", "Skor 4 (Sangat Baik)", "Skor 3 (Baik)", "Skor 2 (Cukup)", "Skor 1 (Kurang)"])

    rubrik_data = [
        ("Pengetahuan", "Jawaban sangat tepat & analisis mendalam", "Jawaban tepat & penjelasan jelas", "Jawaban cukup tepat & kurang rinci", "Jawaban kurang tepat"),
        ("Keterampilan", "Menguasai materi & penyampaian sangat lancar", "Menguasai materi & bahasa jelas", "Membaca teks saat presentasi", "Pasif & tidak menguasai materi"),
        ("Sikap (P3)", "Sangat aktif, kritis, & bertanggung jawab", "Aktif bekerjasama dalam tim", "Kurang berkontribusi dalam kelompok", "Pasif & tidak mengumpulkan tugas")
    ]

    for idx, (asp, s4, s3, s2, s1) in enumerate(rubrik_data, start=1):
        table_rubrik.cell(idx, 0).text = asp
        table_rubrik.cell(idx, 1).text = s4
        table_rubrik.cell(idx, 2).text = s3
        table_rubrik.cell(idx, 3).text = s2
        table_rubrik.cell(idx, 4).text = s1
        table_rubrik.cell(idx, 0).paragraphs[0].runs[0].font.bold = True

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ============================================================
# FORM INPUT DATA
# ============================================================

st.header("📋 Identitas Pembelajaran")
col1, col2 = st.columns(2)

with col1:
    sekolah = st.text_input("Nama Sekolah", "SMP Negeri 1 Cerdas")
    guru = st.text_input("Nama Guru", "Nama Pengajar, S.Pd.")
    mapel = st.selectbox("Mata Pelajaran", [
        "Pendidikan Agama", "PPKn", "Bahasa Indonesia", "Matematika",
        "IPA", "IPS", "Bahasa Inggris", "Seni Budaya", "PJOK", "Prakarya", "Informatika"
    ])
    kelas = st.selectbox("Kelas", ["VII", "VIII", "IX"])

with col2:
    semester = st.selectbox("Semester", ["Ganjil", "Genap"])
    fase = st.selectbox("Fase", ["D", "E"])
    tahun = st.text_input("Tahun Pelajaran", "2026/2027")
    alokasi = st.text_input("Alokasi Waktu", "2 x 40 Menit")

st.divider()

st.header("📖 Identifikasi Pembelajaran")
topik = st.text_input("Topik Utama", "Ekosistem dan Interaksi Makhluk Hidup")
subtopik = st.text_input("Sub Topik", "Rantai Makanan dan Jaring-Jaring Makanan")
cp = st.text_area("Capaian Pembelajaran (CP)", "Peserta didik mampu mengidentifikasi interaksi antar makhluk hidup dan lingkungannya serta merancang upaya pelestarian lingkungan.", height=100)
karakteristik = st.text_area("Karakteristik Peserta Didik (Opsional)", height=80)

# Status Kelengkapan Data
status_count = sum([bool(sekolah.strip()), bool(guru.strip()), bool(topik.strip()), bool(cp.strip())])
persen = int((status_count / 4) * 100)
st.progress(persen / 100)
st.caption(f"Kelengkapan Data Utama: {persen}%")

# ============================================================
# TOMBOL GENERATE
# ============================================================

st.divider()

if st.button("🚀 Generate RPM Lengkap & Format Tabel Word", use_container_width=True, type="primary"):
    if not sekolah.strip():
        st.warning("Silakan isi Nama Sekolah.")
    elif not guru.strip():
        st.warning("Silakan isi Nama Guru.")
    elif not topik.strip():
        st.warning("Silakan isi Topik Utama.")
    elif not cp.strip():
        st.warning("Silakan isi Capaian Pembelajaran.")
    else:
        st.session_state.hasil = generate_rpm_document(
            sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, karakteristik
        )
        st.session_state.history.append({
            "Sekolah": sekolah, "Guru": guru, "Mapel": mapel, "Kelas": kelas, "Topik": topik
        })
        st.success("✅ RPM Lengkap Berhasil Di-generate!")

# ============================================================
# PREVIEW HASIL & DOWNLOAD
# ============================================================

if st.session_state.hasil:
    st.divider()
    st.header("📄 Hasil Dokumen RPM")
    
    st.session_state.hasil = st.text_area(
        "Tinjauan Dokumen (Dapat Disesuaikan):",
        value=st.session_state.hasil,
        height=350
    )

    # Metric Summary
    c1, c2, c3 = st.columns(3)
    c1.metric("Mata Pelajaran", mapel)
    c2.metric("Kelas", kelas)
    c3.metric("Semester", semester)

    # Download Word dengan Tabel Rapi
    file_word = export_word(sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, karakteristik)
    st.download_button(
        label="📥 Download Microsoft Word Rapi (.docx)",
        data=file_word,
        file_name=f"RPM_Tabel_{mapel}_{kelas}_{topik}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

# ============================================================
# FOOTER
# ============================================================

st.divider()
st.markdown("""
<div style="text-align:center;color:gray;font-size:13px">
<b>RPM CERDAS AI v2.2 OFFLINE</b><br>
Penyusun RPM Kurikulum Merdeka Terformat Rapi<br>
© 2026 RPM CERDAS AI
</div>
""", unsafe_allow_html=True)
