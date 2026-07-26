import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

def panggil_ai_guru(topik, cp, komponen, instruksi):
    try:
        import requests
        url = "https://vercel.app"
        p = {"topik": topik, "cp": cp, "komponen": komponen, "instruksi": instruksi}
        res = requests.get(url, params=p, timeout=25).json()
        hasil = res.get('text', "")
        if len(hasil) > 15: return hasil
        return "⏳ Server memproses draf padat. Sila klik kembali tombol AI."
    except Exception:
        return "⚠️ Koneksi sibuk. Silakan klik kembali tombol AI untuk memicu respon."

def buat_dokumen_rpm(d):
    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(1)
    doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(11)
    
    t = doc.add_paragraph()
    tr = t.add_run("RENCANA PEMBELAJARAN MENDALAM (RPM)")
    tr.bold = True; tr.font.size = Pt(14); t.alignment = 1
    doc.add_paragraph()
    
    doc.add_heading("I. IDENTITAS DAN VALIDASI", level=2)
    ti = doc.add_table(rows=7, cols=2); ti.style = 'Table Grid'
    lbls = [
        ("Nama Sekolah", d.get('sekolah', '')), ("Nama Guru", d.get('guru', '')),
        ("Mata Pelajaran", d.get('mapel', '')), ("Kelas / Semester", d.get('kelas_semester', '')),
        ("Alokasi Waktu", d.get('alokasi_waktu', '')), ("Topik Utama", d.get('topik', '')),
        ("Capaian Pembelajaran (CP)", d.get('cp', ''))
    ]
    for i, (l, v) in enumerate(lbls):
        ti.rows[i].cells[0].paragraphs[0].text = str(l)
        ti.rows[i].cells[1].paragraphs[0].text = str(v)
        ti.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True
    doc.add_paragraph()
    
    doc.add_heading("II. KOMKONEN INTI RPM MENDALAM", level=2)
    t_inti = doc.add_table(rows=9, cols=2); t_inti.style = 'Table Grid'
    
    # Pengisian Header Utama Berwarna Abu-abu
    t_inti.rows[0].cells[0].paragraphs[0].text = 'Komponen RPM'
    t_inti.rows[0].cells[1].paragraphs[0].text = 'Deskripsi / Detail Rencana Kerja (Hasil AI & Guru)'
    t_inti.rows[0].cells[0].paragraphs[0].runs[0].font.bold = True
    t_inti.rows[0].cells[1].paragraphs[0].runs[0].font.bold = True
    t_inti.rows[0].cells[0]._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="E6E6E6"/>'.format(nsdecls('w'))))
    t_inti.rows[0].cells[1]._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="E6E6E6"/>'.format(nsdecls('w'))))
    
    k_data = [
        ("1. Dimensi Profil Lulusan", d.get('dimensi_profil', '')),
        ("2. Tujuan Pembelajaran", d.get('tujuan_pembelajaran', '')),
        ("3. Praktik Pedagogis", d.get('praktik_pedagogis', '')),
        ("4. Lingkungan Pembelajaran", d.get('lingkungan_belajar', '')),
        ("5. Kemitraan Pembelajaran", d.get('kemitraan_belajar', '')),
        ("6. Pemanfaatan Digital", d.get('pemanfaatan_digital', '')),
        ("7. Langkah Pembelajaran Rinci", d.get('langkah_pembelajaran', '')),
        ("8. Asesmen & Lembar Kerja", d.get('asesmen_total', ''))
    ]
    for i, (k, isi) in enumerate(k_data):
        t_inti.rows[i+1].cells[0].paragraphs[0].text = str(k)
        t_inti.rows[i+1].cells[1].paragraphs[0].text = str(isi)
        t_inti.rows[i+1].cells[0].paragraphs[0].runs[0].font.bold = True
    doc.add_paragraph(); doc.add_paragraph()
    
    doc.add_heading("III. PENGESAHAN", level=2)
    ttd = doc.add_table(rows=1, cols=2)
    for cell in ttd.rows[0].cells:
        cell._tc.get_or_add_tcPr().append(parse_xml(r'<w:tcBorders {}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w'))))
    ttd.rows[0].cells[0].paragraphs[0].text = f"Mengetahui,\nKepala Sekolah {d.get('sekolah', '')}\n\n\n\n\n( _______________________ )"
    ttd.rows[0].cells[1].paragraphs[0].text = f"Guru Mata Pelajaran,\n\n\n\n\n\n( {d.get('guru', '')} )"
    
    stream = io.BytesIO(); doc.save(stream); stream.seek(0)
    return stream

st.set_page_config(page_title="Aplikasi Pembuat RPM Cerdas", layout="wide")
st.title("🤖 Aplikasi Pembuat Rencana Pembelajaran Mendalam (RPM) Berbasis AI")

if "tujuan_ai" not in st.session_state: st.session_state.tujuan_ai = ""
if "langkah_ai" not in st.session_state: st.session_state.langkah_ai = ""
if "asesmen_ai" not in st.session_state: st.session_state.asesmen_ai = ""

col1, col2 = st.columns(2)
with col1:
    st.subheader("I. Identitas Dasar")
    sekolah = st.text_input("Nama Sekolah", "SMA Negeri 1 Pembelajaran")
    guru = st.text_input("Nama Guru", "Nama Guru, S.Pd.")
    mapel = st.text_input("Mata Pelajaran", "Agama Katolik / Budi Pekerti")
    kelas_semester = st.text_input("Kelas / Semester", "XI / Ganjil")
    alokasi_waktu = st.text_input("Alokasi Waktu", "2 x 45 Menit")
    topik = st.text_input("Topik Pembelajaran", "Kebebasan dan Tanggapan Iman")
    cp = st.text_area("Capaian Pembelajaran (CP)", "Murid mampu menganalisis, mengevaluasi, dan mewujudkan imannya secara nyata...")
    
    # Komponen Pilihan Multi-Select 7 Dimensi Lulusan Sesuai Permintaan
    st.subheader("⚙️ Pemilihan Dimensi Lulusan")
    opsi_dimensi = [
        "Keimanan dan Ketaqwaan terhadap Tuhan YME",
        "Kewargaan", "Penalaran Kritis", "Kreativitas", 
        "Kolaborasi", "Kesehatan", "Komunikasi"
    ]
    dimensi_pilihan = st.multiselect("Pilih Dimensi Lulusan yang Dikembangkan:", opsi_dimensi, default=["Keimanan dan Ketaqwaan terhadap Tuhan YME", "Penalaran Kritis"])

with col2:
    st.subheader("II. Tombol Generator Cerdas AI")
    st.info("💡 Klik tombol di bawah ini satu per satu untuk menyusun rincian komponen ajar.")
    
    if st.button("✨ 2. Rumuskan Tujuan Pembelajaran Mendalam (AI)"):
        with st.spinner("AI memproses..."):
            ins = "Buat Tujuan Pembelajaran mendalam yang wajib memenuhi 3 kriteria: 1) Berkesadaran tinggi (kesadaran diri/iman), 2) Bermakna nyata bagi kehidupan sehari-hari murid, 3) Menggembirakan bagi atmosfer belajar."
            st.session_state.tujuan_ai = panggil_ai_guru(topik, cp, "Tujuan Pembelajaran", ins)
            st.rerun()
            
    if st.button("🔥 7. Kembangkan Kegiatan Rinci (Pendanduluan, Isi, Penutup) (AI)"):
        with st.spinner("AI memproses..."):
            ins_l = "Susun skenario pembelajaran berbasis masalah (PBL) sangat detail per menit dengan uraian wajib mencakup 3 tahap mutlak: 1) Pendahuluan (Apersepsi mendalam & pemantik), 2) Isi/Inti (Penyelidikan masalah kelompok & pemanfaatan teknologi digital), dan 3) Penutup (Refleksi emosional bermakna)."
            st.session_state.langkah_ai = panggil_ai_guru(topik, cp, "Langkah Pembelajaran", ins_l)
            st.rerun()
            
    if st.button("📊 8. Buat Paket Asesmen, LKPD & Rubrik Skor 1-4 (AI)"):
        with st.spinner("AI memproses..."):
            ins_a = "Buat paket instrumen evaluasi lengkap: 1) Metode Evaluasi Formatif & Sumatif. 2) Lembar Kerja Peserta Didik (LKPD/LKM) siap pakai berisi studi kasus logika tinggi dan pertanyaan refleksi murid. 3) Kriteria Penilaian / Rubrik skor 1, 2, 3, sampai 4 lengkap beserta indikator ketercapaiannya."
            st.session_state.asesmen_ai = panggil_ai_guru(topik, cp, "Asesmen & LKPD", ins_a)
            st.rerun()

st.markdown("---")
st.subheader("III. Peninjauan & Penyempurnaan Teks (Dapat Diedit Manual)")

# Konversi daftar pilihan dimensi menjadi untaian teks rapi untuk tabel Word
teks_dimensi = ", ".join(dimensi_pilihan)

dimensi_profil = st.text_area("1. Dimensi Profil Lulusan (Pilihan Guru)", teks_dimensi, height=80)
tujuan_pembelajaran = st.text_area("2. Tujuan Pembelajaran (Hasil AI)", st.session_state.tujuan_ai if st.session_state.tujuan_ai else "Klik tombol AI di atas", height=100)
praktik_pedagogis = st.text_area("3. Praktik Pedagogis", "Menggunakan pendekatan Problem-Based Learning (PBL) berbasis penyelidikan kasus nyata secara berkelompok.")
lingkungan_belajar = st.text_area("4. Lingkungan Pembelajaran", "Fisik: Susunan meja berkelompok. Budaya: Saling menghargai argumen, ramah kesalahan, refleksi terbuka.")
kemitraan_belajar = st.text_area("5. Kemitraan Pembelajaran", "Kolaborasi aktif antar peserta didik, guru sebagai fasilitator, dan pemanfaatan gawai cerdas.")
pemanfaatan_digital = st.text_area("6. Pemanfaatan Digital", "Platform kolaborasi online untuk pengerjaan tugas kelompok secara real-time.")
langkah_pembelajaran = st.text_area("7. Langkah Pembelajaran Rinci (Pendahuluan, Isi, Penutup)", st.session_state.langkah_ai if st.session_state.langkah_ai else "Klik tombol AI di atas", height=150)
asesmen_total = st.text_area("8. Asesmen Pembelajaran (Evaluasi, LKPD, & Rubrik 1-4)", st.session_state.asesmen_ai if st.session_state.asesmen_ai else "Klik tombol AI di atas", height=150)

rpm_data = {
    'sekolah': sekolah, 'guru': guru, 'mapel': mapel, 'kelas_semester': kelas_semester, 'alokasi_waktu': alokasi_waktu,
    'topik': topik, 'cp': cp, 'dimensi_profil': dimensi_profil, 'tujuan_pembelajaran': tujuan_pembelajaran,
    'praktik_pedagogis': praktik_pedagogis, 'lingkungan_belajar': lingkungan_belajar, 'kemitraan_belajar': kemitraan_belajar,
    'pemanfaatan_digital': pemanfaatan_digital, 'langkah_pembelajaran': langkah_pembelajaran, 'asesmen_total': asesmen_total
}

st.markdown("---")
st.subheader("IV. Finalisasi Dokumen RPP")
st.write("Klik tombol di bawah ini untuk langsung menyimpan hasil kerja ke berkas Word:")
try:
    file_word_ready = buat_dokumen_rpm(rpm_data)
    st.download_button(
        label="📥 Simpan Dokumen Word (.docx)", 
        data=file_word_ready, 
        file_name=f"RPM_Cerdas_{topik.replace(' ', '_')}.docx", 
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
except Exception as e:
    st.error(f"⚠️ Gagal menyiapkan tombol simpan berkas. (Detail: {e})")
