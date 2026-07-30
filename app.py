import io
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- CONFIG HALAMAN ---
st.set_page_config(page_title="RPM CERDAS AI - SMPN 7 MARET", page_icon="🧠", layout="wide")

# --- KONSTANTA ---
DIMENSI_LULUSAN = [
    "1. Beriman, Bertakwa kepada Tuhan YME, dan Berakhlak Mulia",
    "2. Berkebinekaan Global", "3. Mandiri", "4. Bergotong Royong",
    "5. Bernalar Kritis", "6. Kreatif", "7. Literasi", "8. Numerasi"
]

# --- LOGIKA AI: PENYEMPURNAAN KONTEN ---
def generate_detailed_steps(topik, subtopik):
    return {
        "pendahuluan": (
            f"1. Guru menyapa siswa dan melakukan doa bersama.\n"
            f"2. Apersepsi: Guru mengaitkan {topik} dengan pengalaman nyata siswa.\n"
            f"3. Guru menyampaikan tujuan pembelajaran dan manfaat mempelajari {subtopik}.\n"
            f"4. Diagnostik singkat: Menanyakan satu pertanyaan pemantik untuk cek kesiapan."
        ),
        "surface": (
            f"• Literasi Dasar: Siswa membaca/menonton materi tentang {topik}.\n"
            f"• Penemuan Konsep: Siswa mengidentifikasi kosakata kunci dan struktur dasar {subtopik}.\n"
            f"• Modeling: Guru memberikan contoh konkret dan menjelaskan 'apa' dan 'bagaimana' konsep tersebut."
        ),
        "deep": (
            f"• Investigasi: Siswa bekerja kelompok menganalisis hubungan antar unsur dalam {subtopik}.\n"
            f"• Diskusi Kritis: Debat atau diskusi mendalam mengenai 'mengapa' dan 'bagaimana jika' terkait {topik}.\n"
            f"• Feedback: Guru memberikan penguatan pada miskonsepsi yang muncul saat diskusi."
        ),
        "transfer": (
            f"• Proyek Kreatif: Siswa membuat karya original (tulisan/produk) berbasis {topik}.\n"
            f"• Kontekstualisasi: Siswa menerapkan solusi dari {subtopik} untuk masalah di lingkungan sekolah/rumah.\n"
            f"• Presentasi: Siswa menyajikan hasil karyanya untuk mendapatkan umpan balik dari rekan sejawat."
        ),
        "penutup": (
            f"1. Refleksi: Siswa menuliskan 'Satu hal baru yang saya pelajari hari ini'.\n"
            f"2. Simpulan: Guru dan siswa merangkum poin-poin utama pembelajaran.\n"
            f"3. Evaluasi: Penjelasan singkat mengenai tugas mandiri atau materi pertemuan berikutnya.\n"
            f"4. Doa dan salam penutup."
        ),
        "rubrik": (
            "Kriteria Penilaian (Skala 1-4):\n"
            "1. Pemahaman Konsep: (1: Kurang, 2: Cukup, 3: Baik, 4: Sangat Baik)\n"
            "2. Analisis Kritis: Kemampuan menghubungkan ide-ide dalam diskusi.\n"
            "3. Kreativitas/Transfer: Orisinalitas dalam produk akhir.\n"
            "4. Kolaborasi: Keaktifan dalam kerja kelompok."
        )
    }

# --- FUNGSI EKSPOR WORD ---
def set_cell_background(cell, fill_color):
    tc_pr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tc_pr.append(shd)

def add_header_row(table, headers):
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        set_cell_background(hdr_cells[i], "1F4E78")
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

def export_word_rpm(data):
    doc = Document()
    
    # Judul
    t = doc.add_heading(level=0)
    run = t.add_run("RENCANA PEMBELAJARAN MENDALAM (RPM)")
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(31, 78, 120)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # I. Identitas
    doc.add_heading("I. IDENTITAS & KARAKTERISTIK", level=2)
    tbl_id = doc.add_table(rows=8, cols=2)
    tbl_id.style = "Table Grid"
    rows_id = [
        ("Nama Sekolah", data['sekolah']), ("Nama Guru", data['guru']),
        ("Mata Pelajaran", data['mapel']), ("Kelas / Fase", f"{data['kelas']} / {data['fase']}"),
        ("Topik Utama", data['topik']), ("Alokasi Waktu", f"{data['total_menit']} Menit"),
        ("Capaian Pembelajaran", data['cp']), ("Karakteristik Siswa", data['karakteristik'])
    ]
    for i, (k, v) in enumerate(rows_id):
        tbl_id.rows[i].cells[0].text = k
        tbl_id.rows[i].cells[1].text = str(v)
        set_cell_background(tbl_id.rows[i].cells[0], "F2F2F2")

    # II. Alur Pembelajaran (Langkah Rinci)
    doc.add_heading("II. LANGKAH-LANGKAH PEMBELAJARAN (METODE SDT)", level=2)
    
    # Pendahuluan
    doc.add_heading("A. PENDAHULUAN", level=3)
    doc.add_paragraph(data['pendahuluan'])

    # Inti (Tabel)
    doc.add_heading("B. KEGIATAN INTI (Surface - Deep - Transfer)", level=3)
    tbl_step = doc.add_table(rows=4, cols=3)
    tbl_step.style = "Table Grid"
    add_header_row(tbl_step, ["Tahapan", "Aktivitas Strategis (Guru & Siswa)", "Durasi"])
    
    steps = [
        ("SURFACE (Pemerolehan)", data['surf'], f"{data['m_surf']} Menit"),
        ("DEEP (Pengolahan)", data['deep'], f"{data['m_deep']} Menit"),
        ("TRANSFER (Penerapan)", data['tran'], f"{data['m_tran']} Menit")
    ]
    for i, (name, content, time) in enumerate(steps):
        row = tbl_step.rows[i+1]
        row.cells[0].text = name
        row.cells[1].text = content
        row.cells[2].text = time

    # Penutup
    doc.add_heading("C. PENUTUP", level=3)
    doc.add_paragraph(data['penutup'])

    # III. Asesmen & Rubrik
    doc.add_heading("III. ASESMEN & RUBRIK PENILAIAN", level=2)
    tbl_as = doc.add_table(rows=4, cols=2)
    tbl_as.style = "Table Grid"
    as_rows = [
        ("Asesmen Formatif", data['as_f']),
        ("Asesmen Sumatif", data['as_s']),
        ("Rubrik Penilaian", data['as_r']),
        ("Pertanyaan Pemantik AI", data['ai_q'])
    ]
    for i, (k, v) in enumerate(as_rows):
        tbl_as.rows[i].cells[0].text = k
        tbl_as.rows[i].cells[1].text = v
        set_cell_background(tbl_as.rows[i].cells[0], "F2F2F2")

    # Tanda Tangan
    doc.add_paragraph("\n")
    ttd_table = doc.add_table(rows=3, cols=2)
    ttd_table.rows[1].cells[0].text = "Mengetahui,\nKepala Sekolah"
    ttd_table.rows[1].cells[1].text = f"{data['tgl_doc']}\nGuru Mata Pelajaran"
    ttd_table.rows[2].cells[0].paragraphs[0].add_run(f"\n\n{data['kepsek']}").font.bold = True
    ttd_table.rows[2].cells[1].paragraphs[0].add_run(f"\n\n{data['guru']}").font.bold = True

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- INTERFACE UTAMA ---
st.title("🧠 RPM CERDAS AI v3.1")
st.caption("Penyempurnaan Langkah Pembelajaran & Rubrik Penilaian")

with st.sidebar:
    st.header("📋 Administrasi")
    sekolah = st.text_input("Sekolah", "SMPN Tujuh Maret Hadakewa")
    guru = st.text_input("Guru", "Daniel F. L. Wahon, S.S")
    kepsek = st.text_input("Kepala Sekolah", "Nama Kepala Sekolah")
    tgl_doc = st.text_input("Tempat, Tanggal", "Hadakewa, 17 Juli 2024")
    total_menit = st.number_input("Total Durasi (Menit)", value=80)

# Kalkulasi Waktu Otomatis
m_surf = int(total_menit * 0.25)
m_deep = int(total_menit * 0.55)
m_tran = total_menit - (m_surf + m_deep)

col1, col2 = st.columns(2)
with col1:
    mapel = st.text_input("Mata Pelajaran", "Bahasa Inggris")
    topik = st.text_input("Topik Utama", "Narrative Text")
with col2:
    kelas = st.text_input("Kelas / Fase", "VII / D")
    subtopik = st.text_input("Sub-Topik", "Legenda Rakyat NTT")

# Generate Konten Default dari AI
ai_content = generate_detailed_steps(topik, subtopik)

st.divider()
st.subheader("📝 Detail Rencana Pembelajaran")

tab1, tab2, tab3 = st.tabs(["Langkah Pembelajaran", "Asesmen & Rubrik", "CP & Karakteristik"])

with tab1:
    pend = st.text_area("1. Pendahuluan (Opening)", ai_content['pendahuluan'], height=120)
    st.markdown("**2. Kegiatan Inti (SDT Framework)**")
    c1, c2, c3 = st.columns(3)
    surf = c1.text_area(f"Surface ({m_surf}m)", ai_content['surface'], height=200)
    deep = c2.text_area(f"Deep ({m_deep}m)", ai_content['deep'], height=200)
    tran = c3.text_area(f"Transfer ({m_tran}m)", ai_content['transfer'], height=200)
    penu = st.text_area("3. Penutup (Closing)", ai_content['penutup'], height=120)

with tab2:
    as_f = st.text_area("Asesmen Formatif (Proses)", "Observasi diskusi kelompok dan ceklis pemahaman konsep.")
    as_s = st.text_area("Asesmen Sumatif (Hasil)", "Produk narasi kreatif atau tes uraian analisis teks.")
    as_r = st.text_area("Rubrik Penilaian Terperinci", ai_content['rubrik'], height=150)
    ai_q = st.text_input("Pertanyaan Pemantik AI", f"Mengapa sebuah legenda tetap diceritakan meskipun sudah berumur ratusan tahun?")

with tab3:
    cp = st.text_area("Capaian Pembelajaran (CP)", "Peserta didik mampu memahami, mengolah, dan menyajikan teks naratif secara kritis dan kreatif.")
    karakter = st.text_input("Karakteristik Siswa", "Heterogen dengan minat literasi visual dan auditori.")

if st.button("📝 Susun & Unduh Dokumen RPM", type="primary"):
    payload = {
        "sekolah": sekolah, "guru": guru, "mapel": mapel, "kelas": kelas, "fase": "D",
        "topik": topik, "subtopik": subtopik, "total_menit": total_menit, "cp": cp,
        "karakteristik": karakter, "pendahuluan": pend, "surf": surf, "deep": deep, 
        "tran": tran, "penutup": penu, "m_surf": m_surf, "m_deep": m_deep, "m_tran": m_tran,
        "as_f": as_f, "as_s": as_s, "as_r": as_r, "ai_q": ai_q,
        "kepsek": kepsek, "tgl_doc": tgl_doc
    }
    final_file = export_word_rpm(payload)
    st.download_button(label="📥 Unduh RPM (.docx)", data=final_file, file_name=f"RPM_{topik}.docx")
