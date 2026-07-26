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
# ============================================================
# BAGIAN 2
# MESIN GENERATOR RPM CERDAS AI
# ============================================================

# ------------------------------------------------------------
# KARAKTERISTIK PESERTA DIDIK
# ------------------------------------------------------------

def generate_karakteristik(mapel, kelas):

    return f"""
Peserta didik kelas {kelas} memiliki kemampuan, minat, dan gaya belajar yang beragam.
Pembelajaran dirancang menggunakan pendekatan diferensiasi sehingga seluruh peserta didik
memperoleh kesempatan belajar yang aktif, bermakna, dan menyenangkan pada mata pelajaran {mapel}.
"""


# ------------------------------------------------------------
# DIMENSI PROFIL LULUSAN
# ------------------------------------------------------------

def generate_dimensi(cp, topik):

    teks = f"""
Pembelajaran pada topik **{topik}** diarahkan untuk mengembangkan Dimensi Profil Lulusan berikut:

• Beriman dan berakhlak mulia
• Bernalar kritis
• Kreatif
• Gotong royong
• Mandiri

Pengembangan dimensi tersebut dilakukan melalui aktivitas penyelidikan,
diskusi, presentasi, refleksi, dan pemecahan masalah yang berkaitan dengan
Capaian Pembelajaran.
"""

    return teks


# ------------------------------------------------------------
# LINTAS DISIPLIN
# ------------------------------------------------------------

def generate_lintas(mapel):

    data = {

        "IPA":"Matematika, Informatika, Bahasa Indonesia",

        "IPS":"Bahasa Indonesia, Informatika",

        "Matematika":"IPA, Informatika",

        "Bahasa Indonesia":"IPS, Seni Budaya",

        "Bahasa Inggris":"Bahasa Indonesia, Informatika",

        "PPKn":"IPS, Bahasa Indonesia",

        "Pendidikan Agama":"PPKn, Bahasa Indonesia",

        "PJOK":"IPA",

        "Seni Budaya":"Bahasa Indonesia",

        "Prakarya":"IPA, Informatika",

        "Informatika":"Matematika"

    }

    return data.get(mapel,"Lintas disiplin disesuaikan dengan kebutuhan pembelajaran.")


# ------------------------------------------------------------
# TUJUAN PEMBELAJARAN
# ------------------------------------------------------------

def generate_tujuan(topik, mapel, kelas, tujuan_manual):

    if tujuan_manual.strip()!="":

        return tujuan_manual

    return f"""
Setelah mengikuti pembelajaran, peserta didik kelas {kelas} mampu:

1. Memahami konsep {topik} pada mata pelajaran {mapel}.

2. Menganalisis informasi secara kritis berdasarkan berbagai sumber belajar.

3. Menyampaikan hasil diskusi secara percaya diri.

4. Bekerja sama dalam menyelesaikan masalah nyata.

5. Merefleksikan manfaat materi dalam kehidupan sehari-hari.

Pembelajaran dilaksanakan secara:

✔ Berkesadaran

✔ Bermakna

✔ Menggembirakan
"""


# ------------------------------------------------------------
# PRAKTIK PEDAGOGIS
# ------------------------------------------------------------

def generate_pedagogis(model):

    return f"""
Guru menerapkan model **{model}** dengan pembelajaran berpusat pada peserta didik.

Peserta didik belajar melalui eksplorasi, diskusi kelompok,
pemecahan masalah, refleksi, dan presentasi hasil belajar.

Guru berperan sebagai fasilitator, motivator, dan pemberi umpan balik.
"""


# ------------------------------------------------------------
# LINGKUNGAN PEMBELAJARAN
# ------------------------------------------------------------

def generate_lingkungan():

    return """
Lingkungan Fisik

• Ruang kelas nyaman

• Tempat duduk berkelompok

• LCD Proyektor

• Internet

• Buku Guru dan Buku Peserta Didik

Lingkungan Budaya

• Aman

• Inklusif

• Kolaboratif

• Menghargai pendapat

• Menyenangkan
"""


# ------------------------------------------------------------
# KEMITRAAN PEMBELAJARAN
# ------------------------------------------------------------

def generate_kemitraan():

    return """
Kolaborasi dilakukan antara:

• Guru sebagai fasilitator pembelajaran.

• Peserta didik sebagai pembelajar aktif.

• Teman sebaya melalui kerja kelompok.

• Teknologi digital sebagai sumber belajar.

• Orang tua sebagai pendamping belajar di rumah.
"""


# ------------------------------------------------------------
# PEMANFAATAN DIGITAL
# ------------------------------------------------------------

def generate_digital(mapel):

    return f"""
Teknologi digital dimanfaatkan melalui:

• Google Classroom

• Canva

• Google Docs

• YouTube Edu

• Quizizz

Penggunaan media disesuaikan dengan karakteristik mata pelajaran {mapel}.
"""


# ============================================================
# PROSES GENERATE
# ============================================================

if generate:

    karakter = karakteristik

    if karakter.strip()=="":
        karakter = generate_karakteristik(mapel,kelas)

    hasil = {

        "karakteristik":karakter,

        "dimensi":generate_dimensi(cp,topik),

        "lintas":generate_lintas(mapel),

        "tujuan":generate_tujuan(
            topik,
            mapel,
            kelas,
            tujuan_manual
        ),

        "pedagogis":generate_pedagogis(model),

        "lingkungan":generate_lingkungan(),

        "kemitraan":generate_kemitraan(),

        "digital":generate_digital(mapel)

    }

    st.session_state.hasil = hasil
