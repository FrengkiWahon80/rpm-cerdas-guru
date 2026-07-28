# ============================================================
# RPM CERDAS AI v2.2 OFFLINE (FORMAT TABEL WORD RAPI)
# BAGIAN 1: CONFIG, IMPORTS, & LOGIKA TEKS PREVIEW ASLI
# ============================================================

import io
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Konfigurasi Halaman Utama Streamlit
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

# Fungsi Pembuat Preview Teks Asli Anda
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
# RPM CERDAS AI v2.2 OFFLINE (FORMAT TABEL WORD RAPI)
# BAGIAN 2: ENGINE EKSPOR TABEL WORD & PENAMBAHAN FITUR TTD
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

def export_word(sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, karakteristik, nama_kepsek, nip_kepsek, nip_guru, tempat_tanggal):
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

    if not karakteristik.strip():
        karakter = f"Peserta didik kelas {kelas} memiliki beragam gaya belajar (visual, auditori, kinestetik) serta tingkat kesiapan belajar yang bervariasi. Pembelajaran dirancang menggunakan diferensiasi proses dan konten."
    else:
        karakter = karakteristik

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
        ("Capaian (CP)", cp),
        ("Karakteristik Siswa", karakter)
    ]
    
    for idx, (label, val) in enumerate(data_id):
        row = table_id.rows[idx]
        row.cells[0].text = label
        row.cells[1].text = val
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        set_cell_background(row.cells[0], "F2F2F2")

    doc.add_paragraph()

    # --- 2. TABEL PROFIL LULUSAN & TUJUAN ---
    doc.add_heading("II. ORIENTASI PROFIL & TUJUAN", level=2)
    table_goal = doc.add_table(rows=2, cols=2)
    table_goal.style = "Table Grid"
    add_header_row(table_goal, ["Dimensi Profil Lulusan", "Tujuan Pembelajaran"])
    
    row_g = table_goal.rows[1]
    p_dim = row_g.cells[0].paragraphs[0]
    p_dim.add_run("• Beriman, Bertakwa kepada Tuhan YME, dan Berakhlak Mulia\n")
    p_dim.add_run("• Berkebinekaan Global\n")
    p_dim.add_run("• Bergotong Royong\n")
    p_dim.add_run("• Mandiri\n")
    p_dim.add_run("• Bernalar Kritis\n")
    p_dim.add_run("• Kreatif")
    
    p_tujuan = row_g.cells[1].paragraphs[0]
    p_tujuan.add_run(f"1. Melalui pengamatan dan diskusi, peserta didik mampu menjelaskan konsep dasar {topik} ({subtopik}) dengan tepat.\n")
    p_tujuan.add_run(f"2. Melalui eksplorasi masalah, peserta didik mampu menganalisis keterkaitan {topik} dengan kehidupan sehari-hari.\n")
    p_tujuan.add_run(f"3. Melalui penugasan kelompok pada LKPD, peserta didik mampu menyusun dan mempresentasikan solusi secara kolaboratif.\n")
    p_tujuan.add_run("4. Menunjukkan sikap bernalar kritis, gotong royong, dan tanggung jawab selama pembelajaran.")

    doc.add_paragraph()

    # --- 3. TABEL LANGKAH PEMBELAJARAN ---
    doc.add_heading("III. LANGKAH-LANGKAH PEMBELAJARAN", level=2)
    table_steps = doc.add_table(rows=4, cols=3)
    table_steps.style = "Table Grid"
    add_header_row(table_steps, ["Tahap Kegiatan", "Durasi", "Detail Aktivitas"])
    
    # Pendahuluan
    r_pend = table_steps.rows[1]
    r_pend.cells[0].text = "Kegiatan Pendahuluan"
    r_pend.cells[1].text = "15 Menit"
    p_pend = r_pend.cells[2].paragraphs[0]
    p_pend.add_run("• Salam, berdoa, dan presensi.\n")
    p_pend.add_run(f"• Apersepsi dan apersepsi materi {topik}.\n")
    p_pend.add_run(f"• Pertanyaan pemantik: \"Bagaimana penerapan {topik} dalam kehidupan kita sehari-hari?\"\n")
    p_pend.add_run("• Menyampaikan tujuan pembelajaran dan alur kegiatan.")
        
    # Inti
    r_inti = table_steps.rows[2]
    r_inti.cells[0].text = "Kegiatan Inti"
    r_inti.cells[1].text = "50 Menit"
    p_inti = r_inti.cells[2].paragraphs[0]
    p_inti.add_run(f"• Orientasi Masalah: Pengamatan tayangan/studi kasus kontekstual terkait {topik}.\n")
    p_inti.add_run("• Pengorganisasian Kelompok: Pembagian kelompok heterogen dan pembagian LKPD.\n")
    p_inti.add_run(f"• Penyelidikan Mandiri/Kelompok: Eksplorasi materi, diskusi, dan pengumpulan data.\n")
    p_inti.add_run("• Penyajian Karya: Presentasi hasil diskusi kelompok di depan kelas.\n")
    p_inti.add_run("• Analisis & Evaluasi: Tanya jawab, penguatan konsep oleh guru, dan klarifikasi miskonsepsi.")
        
    # Penutup
    r_pen = table_steps.rows[3]
    r_pen.cells[0].text = "Kegiatan Penutup"
    r_pen.cells[1].text = "15 Menit"
    p_pen = r_pen.cells[2].paragraphs[0]
    p_pen.add_run("• Rangkuman dan kesimpulan bersama peserta didik.\n")
    p_pen.add_run("• Refleksi pembelajaran dan umpan balik.\n")
    p_pen.add_run("• Penugasan mandiri/informasi materi berikutnya, penutup doa dan salam.")

    doc.add_paragraph()

    # --- 4. TABEL ASESMEN, LKPD & RUBRIK ---
    doc.add_heading("IV. ASESMEN, LKPD & RUBRIK PENILAIAN", level=2)
    table_eval = doc.add_table(rows=4, cols=2)
    table_eval.style = "Table Grid"
    add_header_row(table_eval, ["Komponen", "Rancangan Dokumen"])
    
    # Baris Asesmen
    r_as = table_eval.rows[1]
    r_as.cells[0].text = "Asesmen Pembelajaran"
    r_as.cells[0].paragraphs[0].runs[0].font.bold = True
    r_as.cells[1].text = "• Diagnostik : Tanya jawab / Pertanyaan Pemantik awal.\n• Formatif   : Observasi diskusi kelompok dan unjuk kerja LKPD.\n• Sumatif    : Tes tertulis / penilaian produk akhir."
    
    # Baris LKPD
    r_lk = table_eval.rows[2]
    r_lk.cells[0].text = "Lembar Kerja Peserta Didik (LKPD)"
    r_lk.cells[0].paragraphs[0].runs[0].font.bold = True
    r_lk.cells[1].text = f"1. [Pemahaman Konsep] Jelaskan definisi serta prinsip dasar dari {topik}!\n2. [Analisis Kasus] Analisis 2 contoh penerapan {subtopik} di lingkungan sekitar!\n3. [Penyelesaian Masalah] Rekomendasikan solusi terhadap kendala penerapan {topik}!"
        
    # Baris Rubrik
    r_rb = table_eval.rows[3]
    r_rb.cells[0].text = "Rubrik Penilaian"
    r_rb.cells[0].paragraphs[0].runs[0].font.bold = True
    r_rb.cells[1].text = "1. Penilaian Sikap (Bernalar Kritis, Gotong Royong, Tanggung Jawab)\n2. Penilaian Pengetahuan (Kelengkapan & Ketepatan Jawaban LKPD)\n3. Penilaian Keterampilan (Penguasaan Materi & Penyampaian Presentasi)"

    doc.add_paragraph()
    doc.add_paragraph()

    # --- 5. LAMPIRAN TANDA TANGAN (KODE AMAN PRESTISIUS TANPA ERROR TUPLE) ---
    table_ttd = doc.add_table(rows=3, cols=2)
    
    for row in table_ttd.rows:
        for cell in row.cells:
            tcPr = cell._element.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            for b_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                b = OxmlElement(f'w:{b_name}')
                b.set(qn('w:val'), 'none')
                tcBorders.append(b)
            tcPr.append(tcBorders)

    table_ttd.rows[0].cells[1].text = f"{tempat_tanggal}\n"
    table_ttd.rows[1].cells[0].text = "Mengetahui,\nKepala Sekolah\n\n\n\n"
    table_ttd.rows[1].cells[1].text = "Guru Mata Pelajaran\n\n\n\n"

    p_nkep = table_ttd.rows[2].cells[0].paragraphs[0]
    run_nkep = p_nkep.add_run(nama_kepsek)
    run_nkep.font.underline = True
    run_nkep.font.bold = True
    p_nkep.add_run(f"\nNIP. {nip_kepsek if nip_kepsek else '-'}")

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
# RPM CERDAS AI v2.2 OFFLINE (FORMAT TABEL WORD RAPI)
# BAGIAN 3: USER INTERFACE (TABS) & STATE MANAJEMEN DOWNLOAD
# ============================================================

st.header("⚙️ Form Parameter & Input Rencana Pembelajaran")
st.write("Silakan isi data identitas pokok, topik, dan kelengkapan penandatanganan dokumen di bawah ini:")

# Pembagian Form Input Menggunakan Sistem Tab Sederhana Berdampingan
tab_utama, tab_tanda_tangan = st.tabs(["📝 Identitas Pembelajaran", "✍️ Lembar Pengesahan TTD"])

with tab_utama:
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

    topik = st.text_input("Topik Utama *", value="Kecerdasan Buatan (AI)")
    subtopik = st.text_input("Sub Topik *", value="Penerapan LLM dalam Pendidikan")
    cp = st.text_area("Capaian Pembelajaran (CP)", value="Peserta didik mampu memahami perkembangan teknologi terkini, menganalisis dampak pemanfaatan tools AI secara bijak, dan merancang solusi penyelesaian masalah sehari-hari menggunakan konsep komputasi modern.")
    karakteristik = st.text_area("Karakteristik Siswa (Opsional - Biarkan kosong agar diisi otomatis oleh template)", value="")

with tab_tanda_tangan:
    st.write("Sesuaikan teks administratif penandatanganan lembar terbawah perangkat pembelajaran:")
    col_adm1, col_adm2 = st.columns(2)
    
    with col_adm1:
        tempat_tanggal = st.text_input("Tempat, Tanggal Dokumen", value="Jakarta, 17 Juli 2026")
        nama_kepsek = st.text_input("Nama Kepala Sekolah", value="Dr. H. Supriadi, M.Pd.")
        nip_kepsek = st.text_input("NIP Kepala Sekolah", value="19750824 200003 1 002")
        
    with col_adm2:
        st.write("") 
        st.write("") 
        nip_guru = st.text_input("NIP Guru Mata Pelajaran", value="19891210 201504 2 003")

st.divider()

# Tombol Eksekusi Pembuatan Dokumen
if st.button("🚀 Susun Rencana Pembelajaran Mendalam (AI)", type="primary"):
    if not topik.strip() or not subtopik.strip():
        st.error("Gagal! Kolom Topik Utama dan Sub Topik wajib diisi.")
    else:
        with st.spinner("Template sedang menyelaraskan seluruh parameter identitas..."):
            # Pembuatan Teks Preview
            st.session_state.hasil = generate_rpm_document(
                sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, karakteristik
            )
            # Pembuatan File Binary .docx via fungsi ekspor tabel dengan Tanda Tangan
            st.session_state.word_file = export_word(
                sekolah, guru, mapel, kelas, semester, fase, tahun, alokasi, topik, subtopik, cp, karakteristik, nama_kepsek, nip_kepsek, nip_guru, tempat_tanggal
            )
            st.success("Selesai! RPM Berhasil dihasilkan.")

# Menampilkan Hasil Preview dan Tombol Unduh Dokumen
if st.session_state.hasil:
    st.header("📄 Pratinjau Dokumen Hasil")
    st.text_area("Preview Output Teks", value=st.session_state.hasil, height=450)
    
    st.download_button(
        label="📥 Unduh Dokumen RPM (Format Tabel Word .docx)",
        data=st.session_state.word_file,
        file_name=f"RPM_Mendalam_{topik.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
