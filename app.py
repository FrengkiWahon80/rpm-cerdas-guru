# ============================================================
# RPM CERDAS AI
# Generator RPP Kurikulum Merdeka
# Versi Sederhana
# Bahasa Indonesia
# ============================================================

import streamlit as st
from docx import Document
import io

# ------------------------------------------------------------
# KONFIGURASI HALAMAN
# ------------------------------------------------------------

st.set_page_config(
    page_title="RPM CERDAS AI",
    page_icon="📘",
    layout="wide"
)

# ------------------------------------------------------------
# CSS SEDERHANA
# ------------------------------------------------------------

st.markdown("""
<style>

.main-title{
font-size:34px;
font-weight:bold;
color:#1565C0;
}

.sub-title{
font-size:18px;
color:#555555;
}

.block-container{
padding-top:1rem;
}

</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# JUDUL
# ------------------------------------------------------------

st.markdown(
'<div class="main-title">📘 RPM CERDAS AI</div>',
unsafe_allow_html=True
)

st.markdown(
'<div class="sub-title">Generator Rencana Pelaksanaan Pembelajaran (RPP) Kurikulum Merdeka</div>',
unsafe_allow_html=True
)

st.divider()

# ============================================================
# IDENTITAS
# ============================================================

st.header("📋 Identitas Pembelajaran")

col1,col2=st.columns(2)

with col1:

    sekolah=st.text_input(
        "Nama Sekolah",
        "SMP Negeri 1 Indonesia"
    )

    guru=st.text_input(
        "Nama Guru",
        ""
    )

    mapel=st.selectbox(

        "Mata Pelajaran",

        [

        "Pendidikan Agama",

        "PPKn",

        "Bahasa Indonesia",

        "Matematika",

        "IPA",

        "IPS",

        "Bahasa Inggris",

        "Seni Budaya",

        "PJOK",

        "Informatika",

        "Prakarya"

        ]

    )

    kelas=st.selectbox(

        "Kelas",

        [

        "VII",

        "VIII",

        "IX"

        ]

    )

with col2:

    semester=st.selectbox(

        "Semester",

        [

        "Ganjil",

        "Genap"

        ]

    )

    tahun=st.text_input(

        "Tahun Pelajaran",

        "2026/2027"

    )

    alokasi=st.text_input(

        "Alokasi Waktu",

        "2 x 40 Menit"

    )

    model=st.selectbox(

        "Model Pembelajaran",

        [

        "Problem Based Learning",

        "Project Based Learning",

        "Discovery Learning",

        "Inquiry Learning",

        "Cooperative Learning"

        ]

    )

# ============================================================
# TOPIK DAN CAPAIAN
# ============================================================

st.divider()

st.header("📖 Materi Pembelajaran")

topik=st.text_input(

    "Topik / Materi",

    ""

)

cp=st.text_area(

    "Capaian Pembelajaran",

    height=150

)

tujuan_manual=st.text_area(

    "Tujuan Pembelajaran (Opsional)",

    help="Kosongkan jika ingin dibuat otomatis oleh AI.",

    height=120

)

# ============================================================
# TOMBOL GENERATOR
# ============================================================

st.divider()

generate=st.button(

    "🚀 GENERATE RPP",

    use_container_width=True

)

# ============================================================
# FUNGSI AI LOKAL
# (Bagian 2 akan melanjutkan dari sini)
# ============================================================# ============================================================
# GENERATOR AI LOKAL
# ============================================================

def buat_dimensi(topik):

    return f"""
• Beriman, bertakwa kepada Tuhan Yang Maha Esa, dan berakhlak mulia melalui penghayatan nilai-nilai dalam materi "{topik}".

• Bernalar kritis dengan menganalisis fakta, konsep, dan permasalahan yang berkaitan dengan materi.

• Kreatif dalam menyampaikan ide dan solusi terhadap permasalahan yang dipelajari.

• Bergotong royong melalui diskusi dan kerja kelompok.
"""


def buat_tujuan(topik, mapel, kelas, tujuan_manual):

    if tujuan_manual.strip() != "":
        return tujuan_manual

    return f"""
Setelah mengikuti pembelajaran, peserta didik kelas {kelas} pada mata pelajaran {mapel} diharapkan mampu:

1. Memahami konsep {topik} secara benar.

2. Menganalisis keterkaitan materi dengan kehidupan sehari-hari.

3. Menunjukkan sikap bertanggung jawab, jujur, dan mampu bekerja sama.

4. Menyajikan hasil diskusi secara percaya diri.

5. Merefleksikan manfaat materi dalam kehidupan nyata.

Pembelajaran dirancang agar berkesadaran, bermakna, dan menggembirakan.
"""


def buat_pedagogis(model):

    return f"""
Pembelajaran menggunakan model **{model}** yang menempatkan peserta didik sebagai subjek belajar.

Guru berperan sebagai fasilitator yang membimbing peserta didik menemukan konsep melalui pengamatan, diskusi, eksplorasi, pemecahan masalah, refleksi, dan presentasi hasil belajar.
"""


def buat_lingkungan():

    return """
Lingkungan Fisik

• Ruang kelas yang bersih, nyaman, dan tertata.

• Tempat duduk berkelompok.

• LCD/Proyektor.

• Buku dan media pembelajaran.

Lingkungan Sosial

• Suasana saling menghargai.

• Aman untuk bertanya.

• Kolaboratif.

• Menghargai perbedaan pendapat.
"""


def buat_kemitraan():

    return """
Guru memfasilitasi pembelajaran sebagai pendamping.

Peserta didik bekerja sama dalam kelompok.

Teknologi digital dimanfaatkan sebagai sumber belajar.

Orang tua dapat memberikan dukungan belajar di rumah.
"""


def buat_digital(mapel):

    return f"""
Pemanfaatan teknologi digital dilakukan melalui:

• Google Classroom

• Google Docs

• Canva

• YouTube Edu

• Quizizz

Pemanfaatan aplikasi disesuaikan dengan karakteristik mata pelajaran {mapel}.
"""

# ============================================================
# GENERATE
# ============================================================

if generate:

    dimensi = buat_dimensi(topik)

    tujuan = buat_tujuan(
        topik,
        mapel,
        kelas,
        tujuan_manual
    )

    pedagogis = buat_pedagogis(model)

    lingkungan = buat_lingkungan()

    kemitraan = buat_kemitraan()

    digital = buat_digital(mapel)

    st.session_state["dimensi"] = dimensi
    st.session_state["tujuan"] = tujuan
    st.session_state["pedagogis"] = pedagogis
    st.session_state["lingkungan"] = lingkungan
    st.session_state["kemitraan"] = kemitraan
    st.session_state["digital"] = digital

# ============================================================
# TAMPILKAN HASIL BAGIAN AWAL RPP
# ============================================================

if "tujuan" in st.session_state:

    st.divider()

    st.header("📄 HASIL GENERATOR RPP")

    st.subheader("1. Dimensi Profil Lulusan")

    st.text_area(
        "",
        st.session_state["dimensi"],
        height=170
    )

    st.subheader("2. Tujuan Pembelajaran")

    st.text_area(
        "",
        st.session_state["tujuan"],
        height=220
    )

    st.subheader("3. Praktik Pedagogis")

    st.text_area(
        "",
        st.session_state["pedagogis"],
        height=160
    )

    st.subheader("4. Lingkungan Pembelajaran")

    st.text_area(
        "",
        st.session_state["lingkungan"],
        height=180
    )

    st.subheader("5. Kemitraan Pembelajaran")

    st.text_area(
        "",
        st.session_state["kemitraan"],
        height=150
    )

    st.subheader("6. Pemanfaatan Digital")

    st.text_area(
        "",
        st.session_state["digital"],
        height=170
    )
