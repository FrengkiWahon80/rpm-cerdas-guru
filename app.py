# ============================================================
# RPM CERDAS AI v1.0
# Generator RPP Pembelajaran Mendalam
# SMP Semua Mata Pelajaran
# Kurikulum Merdeka
# ============================================================

import streamlit as st
from docx import Document
import io

# ------------------------------------------------------------
# KONFIGURASI
# ------------------------------------------------------------

st.set_page_config(
    page_title="RPM CERDAS AI",
    page_icon="📘",
    layout="wide"
)

# ------------------------------------------------------------
# CSS
# ------------------------------------------------------------

st.markdown("""
<style>

.main-title{
font-size:36px;
font-weight:bold;
color:#1565C0;
}

.sub-title{
font-size:18px;
color:#555;
}

.box{
background:#F7F9FC;
padding:20px;
border-radius:10px;
border:1px solid #DDDDDD;
margin-bottom:20px;
}

textarea{
font-size:16px !important;
}

</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# HEADER
# ------------------------------------------------------------

st.markdown(
"""
<div class="main-title">
📘 RPM CERDAS AI
</div>
""",
unsafe_allow_html=True
)

st.markdown(
"""
<div class="sub-title">
Generator Rencana Pembelajaran Mendalam (PM)
Kurikulum Merdeka
</div>
""",
unsafe_allow_html=True
)

st.divider()

# ============================================================
# IDENTITAS
# ============================================================

st.header("📋 IDENTITAS")

c1,c2=st.columns(2)

with c1:

    sekolah=st.text_input(
        "Nama Sekolah",
        "SMP Negeri"
    )

    guru=st.text_input(
        "Nama Guru"
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

with c2:

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
# MATERI PEMBELAJARAN
# ============================================================

st.divider()

st.header("📖 IDENTIFIKASI PEMBELAJARAN")

topik=st.text_input(

    "Topik"

)

subtopik=st.text_input(

    "Sub Topik"

)

cp=st.text_area(

    "Capaian Pembelajaran",

    height=130

)

karakteristik=st.text_area(

    "Karakteristik Peserta Didik (Opsional)",

    height=90,

    help="Kosongkan jika ingin dibuat otomatis."

)

tujuan_manual=st.text_area(

    "Tujuan Pembelajaran (Opsional)",

    height=120,

    help="Kosongkan jika ingin dibuat otomatis."

)

st.info(
"""
💡 Tips

Semakin lengkap Capaian Pembelajaran (CP) yang dimasukkan,
semakin baik hasil RPP yang akan dihasilkan.
"""
)

# ============================================================
# GENERATOR
# ============================================================

st.divider()

st.header("🤖 GENERATOR AI")

st.write(
"""
Klik tombol di bawah ini untuk membuat RPP Pembelajaran
Mendalam secara otomatis.
"""
)

generate=st.button(

    "🚀 GENERATE RPP",

    use_container_width=True,

    type="primary"

)

# ============================================================
# TEMPAT PENYIMPANAN HASIL
# ============================================================

if "hasil" not in st.session_state:

    st.session_state.hasil={}

# ============================================================
# BAGIAN 2 AKAN DIMULAI DARI SINI
# ============================================================
