# ============================================================
# NAMA APLIKASI: RPM CERDAS AI v3.0 (EDISI KHUSUS PAK DANIEL)
# SEKOLAH: SMPN TUJUH MARET HADAKEWA
# ============================================================

import io
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- CONFIG HALAMAN ---
st.set_page_config(
    page_title="RPM CERDAS AI - SMPN 7 MARET HADAKEWA",
    page_icon="🧠",
    layout="wide"
)

# --- KONSTANTA: 8 DIMENSI KOMPETENSI LULUSAN (UPDATE RPM TERBARU) ---
DIMENSI_LULUSAN = [
    "1. Beriman, Bertakwa kepada Tuhan YME, dan Berakhlak Mulia",
    "2. Berkebinekaan Global (Nasionalisme & Inklusivitas)",
    "3. Mandiri (Self-Regulated Learning)",
    "4. Bergotong Royong (Kolaborasi & Empati)",
    "5. Bernalar Kritis (Analisis & Evaluasi)",
    "6. Kreatif (Inovasi & Orisinalitas)",
    "7. Literasi (Kemampuan Memahami Konteks Teks Kompleks)",
    "8. Numerasi (Kemampuan Logika Matematika dalam Realitas)"
]

# --- LOGIKA AI: DEEP CONTENT ENHANCER ---
def ai_brain_enhancer(topik, subtopik):
    suggestions = {
        "essential_question": f"Bagaimana prinsip {topik} dapat memprediksi atau menyelesaikan tantangan di masa depan terkait {subtopik}?",
        "surface_tips": f"Fokus pada penguasaan konsep dasar {topik} dan pengenalan istilah kunci.",
        "deep_tips": f"Bimbing siswa menghubungkan {topik} dengan kasus nyata di lingkungan Hadakewa.",
        "transfer_tips": f"Tantang siswa menciptakan solusi unik menggunakan pemahaman {subtopik} mereka."
    }
    return suggestions

# --- FUNGSI EKSPOR WORD (TABEL PROFESIONAL) ---
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
    
    # Judul Dokumen
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
        ("Nama Sekolah", data['sekolah']),
        ("Nama Guru", data['guru']),
        ("Mata Pelajaran", data['mapel']),
        ("Kelas / Fase / Semester", f"{data['kelas']} / {data['fase']} / {data['semester']}"),
        ("Topik Utama / Sub-Topik", f"{data['topik']} / {data['subtopik']}"),
        ("Alokasi Waktu", data['alokasi']),
        ("Capaian Pembelajaran", data['cp']),
        ("Karakteristik Siswa", data['karakteristik'])
    ]
    for i, (k, v) in enumerate(rows_id):
        tbl_id.rows[i].cells[0].text = k
        tbl_id.rows[i].cells[1].text = v
        set_cell_background(tbl_id.rows[i].cells[0], "F2F2F2")
        tbl_id.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True

    # II. Dimensi & Tujuan
    doc.add_heading("II. KOMPETENSI & TUJUAN (DEEP LEARNING)", level=2)
    tbl_dim = doc.add_table(rows=2, cols=2)
    tbl_dim.style = "Table Grid"
    add_header_row(tbl_dim, ["8 Dimensi Kompetensi Lulusan", "Tujuan Pembelajaran"])
    
    cell_dim = tbl_dim.rows[1].cells[0]
    for d in DIMENSI_LULUSAN:
        p = cell_dim.add_paragraph(d, style='List Bullet')
        p.paragraph_format.space_after = Pt(0)
    
    cell_tujuan = tbl_dim.rows[1].cells[1]
    for t_item in data['tujuan_list']:
        if t_item.strip():
            cell_tujuan.add_paragraph(t_item, style='List Number')

    # III. Langkah SDT
    doc.add_heading("III. ALUR PEMBELAJARAN MENDALAM (S-D-T)", level=2)
    tbl_step = doc.add_table(rows=4, cols=3)
    tbl_step.style = "Table Grid"
    add_header_row(tbl_step, ["Tahapan", "Aktivitas Strategis (Penerapan AI/Mendalam)", "Alokasi"])
    
    steps = [
        ("SURFACE (Pemerolehan)", data['surf'], "20%"),
        ("DEEP (Pengolahan)", data['deep'], "60%"),
        ("TRANSFER (Penerapan)", data['tran'], "20%")
    ]
    for i, (name, content, time) in enumerate(steps):
        row = tbl_step.rows[i+1]
        row.cells[0].text = name
        row.cells[1].text = content
        row.cells[2].text = time
        row.cells[0].paragraphs[0].runs[0].font.bold = True

    # IV. Asesmen & Tanda Tangan
    doc.add_heading("IV. EVALUASI & PENGESAHAN", level=2)
    tbl_as = doc.add_table(rows=3, cols=2)
    tbl_as.style = "Table Grid"
    tbl_as.rows[0].cells[0].text = "Asesmen Diagnostik/Formatif"
    tbl_as.rows[0].cells[1].text = data['asesmen_f']
    tbl_as.rows[1].cells[0].text = "Asesmen Sumatif"
    tbl_as.rows[1].cells[1].text = data['asesmen_s']
    tbl_as.rows[2].cells[0].text = "Pertanyaan Pemantik AI"
    tbl_as.rows[2].cells[1].text = data['ai_q']

    doc.add_paragraph("\n\n")
    ttd_table = doc.add_table(rows=3, cols=2)
    ttd_table.rows[0].cells[1].text = f"{data['tgl_doc']}"
    ttd_table.rows[1].cells[0].text = "Mengetahui,\nKepala Sekolah"
    ttd_table.rows[1].cells[1].text = "Guru Mata Pelajaran"
    
    p1 = ttd_table.rows[2].cells[0].paragraphs[0]
    r1 = p1.add_run(f"\n\n{data['kepsek']}\nNIP. {data['nip_kepsek']}")
    r1.font.bold = True
    
    p2 = ttd_table.rows[2].cells[1].paragraphs[0]
    r2 = p2.add_run(f"\n\n{data['guru']}\nNIP. {data['nip_guru']}")
    r2.font.bold = True

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- INTERFACE UTAMA (BROWSER) ---
st.markdown("<h1 style='text-align: center; color: #1F4E78;'>🧠 RPM CERDAS AI</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-weight: bold;'>SMPN TUJUH MARET HADAKEWA</p>", unsafe_allow_html=True)
st.divider()

# Sidebar - Isian Otomatis Pak Daniel
with st.sidebar:
    st.header("📋 Administrasi")
    sekolah = st.text_input("Nama Sekolah", "SMPN Tujuh Maret Hadakewa")
    guru = st.text_input("Nama Guru", "Daniel F. L. Wahon, S.S") # Nama Default Anda
    nip_guru = st.text_input("NIP Guru", "Masukkan NIP Di Sini")
    st.divider()
    kepsek = st.text_input("Kepala Sekolah", "Nama Kepala Sekolah")
    nip_kepsek = st.text_input("NIP Kepsek", "NIP Kepala Sekolah")
    tgl_doc = st.text_input("Tempat, Tanggal", "Hadakewa, 17 Juli 2024")

# Input Konten Pembelajaran
col1, col2 = st.columns(2)
with col1:
    mapel = st.text_input("Mata Pelajaran", "Bahasa Inggris / Informatika")
    kelas = st.text_input("Kelas / Fase", "VII / D")
    semester = st.selectbox("Semester", ["1 (Ganjil)", "2 (Genap)"])
    alokasi = st.text_input("Alokasi Waktu", "2 JP x 40 Menit")
with col2:
    topik = st.text_input("Topik Utama", "Narrative Text")
    subtopik = st.text_input("Sub-Topik", "Legenda Lokal Nusa Tenggara Timur")
    cp = st.text_area("Capaian Pembelajaran (CP)", "Peserta didik mampu memahami konteks literasi dan menghasilkan teks kreatif secara mandiri.")

# Rekomendasi AI
ai_res = ai_brain_enhancer(topik, subtopik)

st.subheader("🚀 Strategi Pembelajaran Mendalam (Deep Learning)")
tab1, tab2, tab3 = st.tabs(["Alur SDT", "Tujuan & Dimensi", "Asesmen & AI Thinking"])

with tab1:
    c1, c2, c3 = st.columns(3)
    surf = c1.text_area("Surface (Pemerolehan Fakta)", f"• {ai_res['surface_tips']}\n• Brainstorming kosa kata baru.\n• Menyimak materi awal.")
    deep = c2.text_area("Deep (Pengolahan Makna)", f"• {ai_res['deep_tips']}\n• Diskusi kelompok menganalisis nilai moral.\n• Membandingkan berbagai versi teks.")
    tran = c3.text_area("Transfer (Penerapan Solusi)", f"• {ai_res['transfer_tips']}\n• Menulis ulang cerita dalam konteks modern.")

with tab2:
    st.info("Sesuai Standar Lulusan terbaru, dokumen mencakup Literasi & Numerasi.")
    tujuan = st.text_area("Tujuan Pembelajaran", 
                         f"Siswa mampu mengidentifikasi struktur {topik}.\nSiswa dapat mengevaluasi pesan moral dalam {subtopik}.\nSiswa mahir mempresentasikan karya secara kolaboratif.")

with tab3:
    as_f = st.text_input("Asesmen Formatif", "Diskusi & Umpan Balik Teman Sejawat")
    as_s = st.text_input("Asesmen Sumatif", "Proyek Menulis Kreatif")
    st.success(f"**AI Thinking (Pertanyaan Pemantik):** {ai_res['essential_question']}")

# Tombol Download
if st.button("📝 Susun & Unduh RPM Sekarang", type="primary"):
    payload = {
        "sekolah": sekolah, "guru": guru, "nip_guru": nip_guru, "mapel": mapel,
        "kelas": kelas, "fase": "D", "semester": semester, "topik": topik,
        "subtopik": subtopik, "alokasi": alokasi, "cp": cp,
        "karakteristik": "Beragam Profil Literasi (Diferensiasi Proses)", "tujuan_list": tujuan.split('\n'),
        "surf": surf, "deep": deep, "tran": tran, "asesmen_f": as_f, "asesmen_s": as_s,
        "kepsek": kepsek, "nip_kepsek": nip_kepsek, "tgl_doc": tgl_doc, "ai_q": ai_res['essential_question']
    }
    
    final_docx = export_word_rpm(payload)
    st.download_button(
        label="📥 Klik Untuk Unduh RPM (.docx)", 
        data=final_docx, 
        file_name=f"RPM_Mendalam_SMPN7_{topik.replace(' ', '_')}.docx"
    )
