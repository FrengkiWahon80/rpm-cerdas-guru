import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

# =====================================================================
# FUNGSI AI RESMI GOOGLE SDK - DIJAMIN LOLOS BLOKIR GOOGLE
# =====================================================================
def panggil_ai_guru(topik, cp, komponen_rpp, instruksi_khusus):
    # Mengambil kunci API secara aman dari Secrets Streamlit Cloud
    api_key_ai = st.secrets.get("GEMINI_API_KEY", "")
    
    if not api_key_ai:
        return "⚠️ Kunci API kosong. Mohon isi 'GEMINI_API_KEY' di menu Secrets Streamlit Cloud Anda."
        
    try:
        # Menggunakan pustaka resmi google-genai SDK standar industri
        from google import genai
        
        # Inisialisasi klien resmi dengan kunci API murni
        client = genai.Client(api_key=str(api_key_ai).strip())
        
        prompt = f"""
        Anda adalah pakar kurikulum pendidikan modern abad 21 dan perancang Rencana Pembelajaran Mendalam (RPM).
        Tugas Anda adalah mengembangkan komponen '{komponen_rpp}' secara sangat rinci, mendalam, aplikatif, 
        dan berbobot sebagai referensi utama guru di kelas.
        
        Informasi Dasar Kelas:
        - Topik Pembelajaran: {topik}
        - Capaian Pembelajaran (CP): {cp}
        
        Instruksi Khusus untuk Komponen Ini:
        {instruksi_khusus}
        
        Berikan jawaban dalam Bahasa Indonesia yang padat, aplikatif, tuntas, tanpa basa-basi kalimat pembuka.
        """
        
        # Memanggil model resmi Gemini 2.5 Flash melalui jalur SDK
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
        )
        return response.text
    except Exception as e:
        return f"⚠️ Gagal memuat AI otomatis. Anda dapat mengetik manual. (Detail Eror SDK: {str(e)})"

# =====================================================================
# FUNGSI UTAMA: MENYUSUN DATA MENJADI TABEL WORD YANG RAPI
# =====================================================================
def buat_dokumen_rpm(data):
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
        ("Nama Sekolah", data.get('sekolah', '')), ("Nama Guru", data.get('guru', '')),
        ("Mata Pelajaran", data.get('mapel', '')), ("Kelas / Semester", data.get('kelas_semester', '')),
        ("Alokasi Waktu", data.get('alokasi_waktu', '')), ("Topik Utama", data.get('topik', '')),
        ("Capaian Pembelajaran (CP)", data.get('cp', ''))
    ]
    for i, (l, v) in enumerate(lbls):
        ti.rows[i].cells[0].text = str(l)
        ti.rows[i].cells[1].text = str(v)
        ti.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True
    doc.add_paragraph()
    
    doc.add_heading("II. KOMKONEN INTI RPM MENDALAM", level=2)
    t_inti = doc.add_table(rows=9, cols=2); t_inti.style = 'Table Grid'
    
    # Pengisian header tabel inti secara aman per sel kolom 0 dan 1
    t_inti.rows[0].cells[0].text = 'Komponen RPM'
    t_inti.rows[0].cells[1].text = 'Deskripsi / Detail Rencana Kerja (Hasil AI & Guru)'
    t_inti.rows[0].cells[0].paragraphs[0].runs[0].font.bold = True
    t_inti.rows[0].cells[1].paragraphs[0].runs[0].font.bold = True
    t_inti.rows[0].cells[0]._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="E6E6E6"/>'.format(nsdecls('w'))))
    t_inti.rows[0].cells[1]._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="E6E6E6"/>'.format(nsdecls('w'))))
    
    k_data = [
        ("1. Dimensi Profil Lulusan", data.get('dimensi_profil', '')),
        ("2. Tujuan Pembelajaran", data.get('tujuan_pembelajaran', '')),
        ("3. Praktik Pedagogis", data.get('praktik_pedagogis', '')),
        ("4. Lingkungan Pembelajaran", data.get('lingkungan_belajar', '')),
        ("5. Kemitraan Pembelajaran", data.get('kemitraan_belajar', '')),
        ("6. Pemanfaatan Digital", data.get('pemanfaatan_digital', '')),
        ("7. Langkah Pembelajaran Rinci", data.get('langkah_pembelajaran', '')),
        ("8. Asesmen & Lembar Kerja", data.get('asesmen_total', ''))
    ]
    for i, (k, isi) in enumerate(k_data):
        t_inti.rows[i+1].cells[0].text = str(k)
        t_inti.rows[i+1].cells[1].text = str(isi)
        t_inti.rows[i+1].cells[0].paragraphs[0].runs[0].font.bold = True
    doc.add_paragraph(); doc.add_paragraph()
    
    doc.add_heading("III. PENGESAHAN", level=2)
    ttd = doc.add_table(rows=1, cols=2)
    for cell in ttd.rows[0].cells:
        cell._tc.get_or_add_tcPr().append(parse_xml(r'<w:tcBorders {}><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w'))))
    ttd.rows[0].cells[0].paragraphs[0].text = f"Mengetahui,\nKepala Sekolah {data.get('sekolah', '')}\n\n\n\n\n( _______________________ )"
    ttd.rows[0].cells[1].paragraphs[0].text = f"Guru Mata Pelajaran,\n\n\n\n\n\n( {data.get('guru', '')} )"
    
    stream = io.BytesIO(); doc.save(stream); stream.seek(0)
    return stream

# =====================================================================
# ANTARMUKA WEB STREAMLIT
# =====================================================================
st.set_page_config(page_title="Aplikasi Pembuat RPM Cerdas", layout="wide")
st.title("🤖 Aplikasi Pembuat Rencana Pembelajaran Mendalam (RPM) Berbasis AI")

if "profil_ai" not in st.session_state: st.session_state.profil_ai = ""
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
    cp = st.text_area("Capaian Pembelajaran (CP)", "Murid mampu menganalisis, mengevaluasi, dan mewujudkan imannya secara nyata dalam konteks kebebasan...")

with col2:
    st.subheader("II. Tombol Generator Cerdas AI")
    st.info("💡 Klik tombol di bawah ini satu per satu untuk mengisi draf RPM secara otomatis.")
    if st.button("✨ 1 & 2. Rumuskan Profil Lulusan & Tujuan (AI)"):
        with st.spinner("AI memproses..."):
            st.session_state.profil_ai = panggil_ai_guru(topik, cp, "Dimensi Profil Lulusan", "Rincikan Keterampilan abad 21.")
            st.session_state.tujuan_ai = panggil_ai_guru(topik, cp, "Tujuan Pembelajaran", "Rumuskan Tujuan Pembelajaran yang Berkesadaran, Bermakna, dan Menggembirakan.")
            st.rerun()
    if st.button("🔥 7. Kembangkan Kegiatan Pembelajaran Rinci (AI)"):
        with st.spinner("AI memproses..."):
            st.session_state.langkah_ai = panggil_ai_guru(topik, cp, "Langkah Pembelajaran", "Buat tahapan proses PBL rinci per menit: Pembukaan, Inti, Penutup.")
            st.rerun()
    if st.button("📊 8. Buat Instrumen Asesmen & LKM Lengkap (AI)"):
        with st.spinner("AI memproses..."):
            st.session_state.asesmen_ai = panggil_ai_guru(topik, cp, "Asesmen & LKM", "Buat evaluasi Formatif Sumatif, Lembar Kerja Murid (LKM), dan Rubrik skor 1-4.")
            st.rerun()

st.markdown("---")
st.subheader("III. Peninjauan & Penyempurnaan Teks (Dapat Diedit Manual)")
dimensi_profil = st.text_area("1. Dimensi Profil Lulusan", st.session_state.profil_ai if st.session_state.profil_ai else "Klik tombol AI di atas", height=100)
tujuan_pembelajaran = st.text_area("2. Tujuan Pembelajaran", st.session_state.tujuan_ai if st.session_state.tujuan_ai else "Klik tombol AI di atas", height=100)
praktik_pedagogis = st.text_area("3. Praktik Pedagogis", "Menggunakan pendekatan Problem-Based Learning (PBL) berbasis penyelidikan kasus nyata secara berkelompok.")
lingkungan_belajar = st.text_area("4. Lingkungan Pembelajaran", "Fisik: Susunan meja berkelompok. Budaya: Saling menghargai argumen, ramah kesalahan, refleksi terbuka.")
kemitraan_belajar = st.text_area("5. Kemitraan Pembelajaran", "Kolaborasi aktif antar peserta didik, guru sebagai fasilitator, dan pemanfaatan gawai cerdas.")
pemanfaatan_digital = st.text_area("6. Pemanfaatan Digital", "Platform kolaborasi online untuk pengerjaan tugas kelompok secara real-time.")
langkah_pembelajaran = st.text_area("7. Langkah Pembelajaran Rinci", st.session_state.langkah_ai if st.session_state.langkah_ai else "Klik tombol AI di atas", height=150)
asesmen_total = st.text_area("8. Asesmen Pembelajaran & LKM", st.session_state.asesmen_ai if st.session_state.asesmen_ai else "Klik tombol AI di atas", height=150)

rpm_data = {'sekolah': sekolah, 'guru': guru, 'mapel': mapel, 'kelas_semester': kelas_semester, 'alokasi_waktu': alokasi_waktu, 'topik': topik, 'cp': cp, 'dimensi_profil': dimensi_profil, 'tujuan_pembelajaran': tujuan_pembelajaran, 'praktik_pedagogis': praktik_pedagogis, 'lingkungan_belajar': lingkungan_belajar, 'kemitraan_belajar': kemitraan_belajar, 'pemanfaatan_digital': pemanfaatan_digital, 'langkah_pembelajaran': langkah_pembelajaran, 'asesmen_total': asesmen_total }st.markdown("---")
st.subheader("IV. Finalisasi Dokumen RPP")
try: ile_word_ready = buat_dokumen_rpm(rpm_data)
st.download_button(label="📥 Unduh Dokumen RPM (.docx)", data=file_word_ready, file_name=f"RPM_Cerdas_{topik.replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
except Exception as e:
st.error(f"⚠️ Gagal menyiapkan tombol unduh. (Detail: {e})")
