# ============================================================
# NAMA APLIKASI: RPM CERDAS AI v3.1 (EDISI PENYEMPURNAAN)
# PENYUSUN: Daniel F. L. Wahon, S.S
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
st.set_page_config(page_title="RPM CERDAS AI - SMPN 7 MARET", page_icon="🧠", layout="wide")

# --- KONSTANTA: 8 DIMENSI KOMPETENSI LULUSAN ---
DIMENSI_LULUSAN = [
    "1. Beriman, Bertakwa kepada Tuhan YME, dan Berakhlak Mulia",
    "2. Berkebinekaan Global", "3. Mandiri", "4. Bergotong Royong",
    "5. Bernalar Kritis", "6. Kreatif", "7. Literasi", "8. Numerasi"
]

# --- LOGIKA AI: ASESMEN & RUBRIK ENHANCER ---
def ai_assessment_logic(topik):
    return {
        "diagnostik": f"Kuis cepat/tanya jawab lisan mengenai pengetahuan awal siswa tentang {topik}.",
        "formatif": f"Observasi diskusi kelompok, ceklis keaktifan, dan penilaian antar teman saat proses pengolahan makna.",
        "sumatif": f"Penilaian produk akhir (proyek/laporan) atau tes tertulis uraian yang menguji daya nalar.",
        "rubrik": f"Kriteria: 1. Kedalaman analisis (40%), 2. Orisinalitas ide (30%), 3. Kualitas presentasi/produk (30%).",
        "pemantik": f"Jika konsep {topik} ini tidak pernah ditemukan, bagaimana cara manusia memecahkan masalah tersebut secara manual?"
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
        ("Mata Pelajaran", data['mapel']), ("Kelas / Fase / Semester", f"{data['kelas']} / {data['fase']} / {data['semester']}"),
        ("Topik / Sub-Topik", f"{data['topik']} / {data['subtopik']}"), ("Alokasi Waktu", f"{data['total_menit']} Menit"),
        ("Capaian Pembelajaran", data['cp']), ("Karakteristik Siswa", data['karakteristik'])
    ]
    for i, (k, v) in enumerate(rows_id):
        tbl_id.rows[i].cells[0].text = k
        tbl_id.rows[i].cells[1].text = str(v)
        set_cell_background(tbl_id.rows[i].cells[0], "F2F2F2")
        tbl_id.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True

    # II. 8 Dimensi & Tujuan
    doc.add_heading("II. KOMPETENSI & TUJUAN", level=2)
    tbl_dim = doc.add_table(rows=2, cols=2)
    tbl_dim.style = "Table Grid"
    add_header_row(tbl_dim, ["8 Dimensi Kompetensi Lulusan", "Tujuan Pembelajaran"])
    cell_dim = tbl_dim.rows[1].cells[0]
    for d in DIMENSI_LULUSAN:
        p = cell_dim.add_paragraph(d, style='List Bullet')
        p.paragraph_format.space_after = Pt(0)
    cell_tujuan = tbl_dim.rows[1].cells[1]
    for t_item in data['tujuan_list']:
        if t_item.strip(): cell_tujuan.add_paragraph(t_item, style='List Number')

    # III. Langkah SDT (Waktu dalam MENIT)
    doc.add_heading("III. ALUR PEMBELAJARAN (SURFACE - DEEP - TRANSFER)", level=2)
    tbl_step = doc.add_table(rows=4, cols=3)
    tbl_step.style = "Table Grid"
    add_header_row(tbl_step, ["Tahapan", "Aktivitas Strategis", "Durasi"])
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
        row.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # IV. Asesmen Detail & Rubrik
    doc.add_heading("IV. RANCANGAN ASESMEN & RUBRIK", level=2)
    tbl_as = doc.add_table(rows=5, cols=2)
    tbl_as.style = "Table Grid"
    as_rows = [
        ("Asesmen Diagnostik", data['as_d']),
        ("Asesmen Formatif (Proses)", data['as_f']),
        ("Asesmen Sumatif (Hasil)", data['as_s']),
        ("Kriteria/Rubrik Penilaian", data['as_r']),
        ("Pertanyaan Pemantik AI", data['ai_q'])
    ]
    for i, (k, v) in enumerate(as_rows):
        tbl_as.rows[i].cells[0].text = k
        tbl_as.rows[i].cells[1].text = v
        set_cell_background(tbl_as.rows[i].cells[0], "F2F2F2")

    # Tanda Tangan
    doc.add_paragraph("\n")
    ttd_table = doc.add_table(rows=3, cols=2)
    ttd_table.rows[0].cells[1].text = f"{data['tgl_doc']}"
    ttd_table.rows[1].cells[0].text = "Mengetahui,\nKepala Sekolah"
    ttd_table.rows[1].cells[1].text = "Guru Mata Pelajaran"
    ttd_table.rows[2].cells[0].paragraphs[0].add_run(f"\n\n{data['kepsek']}\nNIP. {data['nip_kepsek']}").font.bold = True
    ttd_table.rows[2].cells[1].paragraphs[0].add_run(f"\n\n{data['guru']}\nNIP. {data['nip_guru']}").font.bold = True

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- INTERFACE UTAMA ---
st.markdown("<h2 style='text-align: center; color: #1F4E78;'>🧠 RPM CERDAS AI v3.1</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'><b>Daniel F. L. Wahon, S.S</b> | SMPN TUJUH MARET HADAKEWA</p>", unsafe_allow_html=True)
st.divider()

with st.sidebar:
    st.header("📋 Administrasi")
    sekolah = st.text_input("Sekolah", "SMPN Tujuh Maret Hadakewa")
    guru = st.text_input("Guru", "Daniel F. L. Wahon, S.S")
    nip_guru = st.text_input("NIP Guru", "19...")
    kepsek = st.text_input("Kepala Sekolah", "Nama Kepala Sekolah")
    nip_kepsek = st.text_input("NIP Kepsek", "19...")
    tgl_doc = st.text_input("Tempat, Tanggal", "Hadakewa, 17 Juli 2024")
    st.divider()
    total_menit = st.number_input("Total Durasi Pembelajaran (Menit)", min_value=1, value=80)

# Kalkulasi Menit Otomatis (20% - 60% - 20%)
m_surf = int(total_menit * 0.2)
m_deep = int(total_menit * 0.6)
m_tran = total_menit - (m_surf + m_deep)

col1, col2 = st.columns(2)
with col1:
    mapel = st.text_input("Mata Pelajaran", "Bahasa Inggris")
    kelas = st.text_input("Kelas / Fase", "VII / D")
    topik = st.text_input("Topik Utama", "Narrative Text")
with col2:
    semester = st.selectbox("Semester", ["1 (Ganjil)", "2 (Genap)"])
    subtopik = st.text_input("Sub-Topik", "Legenda Rakyat NTT")
    cp = st.text_area("Capaian Pembelajaran (CP)", "Peserta didik memahami konteks literasi dan memproduksi teks kreatif.")

ai_as = ai_assessment_logic(topik)

st.subheader("🚀 Strategi & Asesmen Mendalam")
tab1, tab2, tab3 = st.tabs(["Langkah Pembelajaran (Menit)", "Asesmen Lengkap", "Tujuan & Rubrik"])

with tab1:
    st.caption(f"Distribusi Waktu Otomatis: Surface {m_surf}m, Deep {m_deep}m, Transfer {m_tran}m")
    c1, c2, c3 = st.columns(3)
    surf = c1.text_area("Surface Activity", f"• Penjelasan konsep dasar {topik}.\n• Menyimak kosakata baru.", height=150)
    deep = c2.text_area("Deep Activity", f"• Analisis unsur intrinsik {subtopik}.\n• Diskusi kelompok makna tersirat.", height=150)
    tran = c3.text_area("Transfer Activity", f"• Menulis ulang/menceritakan kembali cerita dalam konteks saat ini.", height=150)

with tab2:
    as_d = st.text_area("Asesmen Diagnostik", ai_as['diagnostik'])
    as_f = st.text_area("Asesmen Formatif (Proses)", ai_as['formatif'])
    as_s = st.text_area("Asesmen Sumatif (Hasil Akhir)", ai_as['sumatif'])

with tab3:
    tujuan = st.text_area("Tujuan Pembelajaran", f"Siswa dapat menganalisis {topik}.\nSiswa terampil menyusun {subtopik}.")
    as_r = st.text_area("Rubrik Penilaian (Kriteria)", ai_as['rubrik'])
    st.info(f"**AI Prompt (Pertanyaan Pemantik):** {ai_as['pemantik']}")

if st.button("📝 Susun & Unduh Dokumen RPM", type="primary"):
    payload = {
        "sekolah": sekolah, "guru": guru, "nip_guru": nip_guru, "mapel": mapel,
        "kelas": kelas, "fase": "D", "semester": semester, "topik": topik,
        "subtopik": subtopik, "total_menit": total_menit, "cp": cp,
        "karakteristik": "Diferensiasi Minat", "tujuan_list": tujuan.split('\n'),
        "surf": surf, "deep": deep, "tran": tran,
        "m_surf": m_surf, "m_deep": m_deep, "m_tran": m_tran,
        "as_d": as_d, "as_f": as_f, "as_s": as_s, "as_r": as_r,
        "kepsek": kepsek, "nip_kepsek": nip_kepsek, "tgl_doc": tgl_doc, "ai_q": ai_as['pemantik']
    }
    final_file = export_word_rpm(payload)
    st.download_button(label="📥 Unduh RPM (.docx)", data=final_file, file_name=f"RPM_Mendalam_{topik}.docx")
