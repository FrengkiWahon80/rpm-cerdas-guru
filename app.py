# ==========================================================
# RPM CERDAS AI v1.0
# Generator Rencana Pembelajaran Mendalam
# Jenjang : SMP
# Kurikulum : Merdeka
# Bahasa : Indonesia
# ==========================================================

import streamlit as st
from datetime import datetime

# ----------------------------------------------------------
# KONFIGURASI HALAMAN
# ----------------------------------------------------------

st.set_page_config(
    page_title="RPM CERDAS AI",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------------------------------------
# CSS
# ----------------------------------------------------------

st.markdown("""
<style>

.block-container{
    padding-top:1rem;
}

.main-title{
    font-size:34px;
    font-weight:bold;
    color:#1565C0;
}

.sub-title{
    font-size:18px;
    color:#555;
}

.card{
    background:#F8F9FA;
    border-radius:12px;
    padding:20px;
    border:1px solid #DDDDDD;
}

.footer{
    text-align:center;
    color:gray;
    margin-top:50px;
}

</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------
# SESSION STATE
# ----------------------------------------------------------

default_data = {

    "sekolah":"",
    "guru":"",
    "nip":"",
    "mapel":"",
    "kelas":"",
    "fase":"",
    "semester":"",
    "tahun":"",
    "materi":"",
    "topik":"",
    "cp":""

}

for key,value in default_data.items():

    if key not in st.session_state:

        st.session_state[key]=value

# ----------------------------------------------------------
# SIDEBAR
# ----------------------------------------------------------

with st.sidebar:

    st.image(
        "https://img.icons8.com/color/96/books.png",
        width=80
    )

    st.title("RPM CERDAS AI")

    st.caption("Versi Lokal 1.0")

    st.divider()

    menu = st.radio(

        "Menu",

        [

            "🏠 Dashboard",

            "📋 Identitas",

            "🎯 Tujuan",

            "📚 Langkah",

            "📝 Asesmen",

            "📄 LKPD",

            "📊 Rubrik",

            "⬇ Export Word"

        ]

    )

    st.divider()

    st.success("Jenjang : SMP")

    st.info("Kurikulum Merdeka")

    st.caption("Versi Offline")

# ----------------------------------------------------------
# DASHBOARD
# ----------------------------------------------------------

if menu=="🏠 Dashboard":

    st.markdown(
        '<div class="main-title">📘 RPM CERDAS AI</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="sub-title">Generator Rencana Pembelajaran Mendalam</div>',
        unsafe_allow_html=True
    )

    st.divider()

    col1,col2,col3,col4 = st.columns(4)

    with col1:
        st.metric("Jenjang","SMP")

    with col2:
        st.metric("Kurikulum","Merdeka")

    with col3:
        st.metric("Versi","1.0")

    with col4:
        st.metric(
            "Tahun",
            datetime.now().year
        )

    st.write("")

    st.markdown("""
### Selamat Datang

RPM CERDAS AI membantu guru menyusun perangkat pembelajaran secara lebih cepat.

Aplikasi ini akan menghasilkan:

- Tujuan Pembelajaran
- Langkah Pembelajaran
- Pembelajaran Mendalam
- Asesmen Diagnostik
- Asesmen Formatif
- Asesmen Sumatif
- LKPD
- Rubrik Penilaian
- Pengayaan
- Remedial
- Refleksi
- Dokumen Word siap cetak
""")

    st.info(
        "Silakan pilih menu **Identitas** pada sidebar untuk mulai membuat perangkat pembelajaran."
    )

# ----------------------------------------------------------
# HALAMAN BELUM DIBUAT
# ----------------------------------------------------------

else:

    st.warning(
        "Halaman ini akan dibuat pada file berikutnya."
    )

# ----------------------------------------------------------
# FOOTER
# ----------------------------------------------------------

st.divider()

st.markdown(
"""
<div class="footer">

RPM CERDAS AI v1.0<br>

Dikembangkan menggunakan Python + Streamlit

</div>
""",
unsafe_allow_html=True
)
# ==========================================================
# HALAMAN IDENTITAS
# RPM CERDAS AI
# ==========================================================

import streamlit as st

st.set_page_config(page_title="Identitas", page_icon="📋", layout="wide")

st.title("📋 Identitas Pembelajaran")

st.markdown("Lengkapi identitas perangkat pembelajaran terlebih dahulu.")

st.divider()

# ==========================================================
# IDENTITAS SEKOLAH
# ==========================================================

st.subheader("🏫 Identitas Sekolah")

col1, col2 = st.columns(2)

with col1:

    sekolah = st.text_input(
        "Nama Sekolah",
        value=st.session_state.get("sekolah", "")
    )

    guru = st.text_input(
        "Nama Guru",
        value=st.session_state.get("guru", "")
    )

    nip = st.text_input(
        "NIP",
        value=st.session_state.get("nip", "")
    )

    mapel = st.selectbox(
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

with col2:

    kelas = st.selectbox(
        "Kelas",
        ["VII", "VIII", "IX"]
    )

    fase = st.selectbox(
        "Fase",
        ["D"]
    )

    semester = st.selectbox(
        "Semester",
        ["Ganjil", "Genap"]
    )

    tahun = st.text_input(
        "Tahun Pelajaran",
        value="2026/2027"
    )

# ==========================================================
# PEMBELAJARAN
# ==========================================================

st.divider()

st.subheader("📖 Informasi Pembelajaran")

materi = st.text_input(
    "Materi Pokok"
)

topik = st.text_input(
    "Topik Pembelajaran"
)

alokasi = st.text_input(
    "Alokasi Waktu",
    value="2 x 40 Menit"
)

cp = st.text_area(
    "Capaian Pembelajaran (CP)",
    height=150
)

# ==========================================================
# MODEL PEMBELAJARAN
# ==========================================================

st.divider()

st.subheader("🎯 Strategi Pembelajaran")

col3, col4 = st.columns(2)

with col3:

    model = st.selectbox(
        "Model Pembelajaran",
        [
            "Problem Based Learning",
            "Project Based Learning",
            "Discovery Learning",
            "Inquiry Learning",
            "Cooperative Learning",
            "Problem Solving",
            "Contextual Teaching Learning",
            "Role Playing"
        ]
    )

    pendekatan = st.selectbox(
        "Pendekatan",
        [
            "Deep Learning",
            "Saintifik",
            "Teaching at The Right Level",
            "Konstruktivistik"
        ]
    )

with col4:

    metode = st.multiselect(
        "Metode Pembelajaran",
        [
            "Diskusi",
            "Ceramah",
            "Presentasi",
            "Tanya Jawab",
            "Penugasan",
            "Observasi",
            "Eksperimen",
            "Studi Kasus"
        ],
        default=["Diskusi", "Presentasi"]
    )

    media = st.text_input(
        "Media Pembelajaran",
        value="LCD, Laptop, LKPD"
    )

# ==========================================================
# SUMBER BELAJAR
# ==========================================================

st.divider()

st.subheader("📚 Sumber Belajar")

sumber = st.text_area(
    "Sumber Belajar",
    value="""1. Buku Guru
2. Buku Siswa
3. Internet
4. Lingkungan Sekitar"""
)

# ==========================================================
# PROFIL LULUSAN
# ==========================================================

st.divider()

st.subheader("🌟 Profil Lulusan")

profil = st.multiselect(
    "Pilih Dimensi Profil Lulusan",
    [
        "Keimanan dan Ketakwaan terhadap Tuhan YME",
        "Kewargaan",
        "Penalaran Kritis",
        "Kreativitas",
        "Kolaborasi",
        "Komunikasi",
        "Kesehatan"
    ],
    default=["Penalaran Kritis"]
)

# ==========================================================
# UPLOAD
# ==========================================================

st.divider()

st.subheader("🖼 Upload Dokumen")

logo = st.file_uploader(
    "Logo Sekolah",
    type=["png", "jpg", "jpeg"]
)

ttd_guru = st.file_uploader(
    "Tanda Tangan Guru",
    type=["png", "jpg"]
)

ttd_kepsek = st.file_uploader(
    "Tanda Tangan Kepala Sekolah",
    type=["png", "jpg"]
)

# ==========================================================
# SIMPAN SESSION
# ==========================================================

if st.button("💾 Simpan Identitas", use_container_width=True):

    st.session_state["sekolah"] = sekolah
    st.session_state["guru"] = guru
    st.session_state["nip"] = nip
    st.session_state["mapel"] = mapel
    st.session_state["kelas"] = kelas
    st.session_state["fase"] = fase
    st.session_state["semester"] = semester
    st.session_state["tahun"] = tahun
    st.session_state["materi"] = materi
    st.session_state["topik"] = topik
    st.session_state["alokasi"] = alokasi
    st.session_state["cp"] = cp
    st.session_state["model"] = model
    st.session_state["pendekatan"] = pendekatan
    st.session_state["metode"] = metode
    st.session_state["media"] = media
    st.session_state["sumber"] = sumber
    st.session_state["profil"] = profil

    st.success("✅ Identitas berhasil disimpan.")
"""
=========================================================
GENERATOR RPM CERDAS AI
Versi Lokal
=========================================================
"""

# =========================================================
# TUJUAN PEMBELAJARAN
# =========================================================

def generate_tujuan(mapel, topik, kelas):

    tujuan = f"""
Peserta didik kelas {kelas} mampu memahami konsep {topik}
pada mata pelajaran {mapel},
menunjukkan sikap kritis,
mampu bekerja sama,
mengkomunikasikan hasil belajar,
serta menerapkan pengetahuan dalam kehidupan sehari-hari.
"""

    return tujuan.strip()


# =========================================================
# PEMAHAMAN BERMAKNA
# =========================================================

def generate_pemahaman(topik):

    return f"""
Peserta didik memahami bahwa {topik}
berkaitan erat dengan kehidupan sehari-hari,
sehingga dapat diterapkan secara bertanggung jawab.
""".strip()


# =========================================================
# PERTANYAAN PEMANTIK
# =========================================================

def generate_pemantik(topik):

    return f"""
1. Apa yang kamu ketahui tentang {topik}?

2. Mengapa {topik} penting dipelajari?

3. Bagaimana penerapan {topik} dalam kehidupan sehari-hari?
""".strip()


# =========================================================
# PROFIL LULUSAN
# =========================================================

def generate_profil(profil):

    hasil = []

    for item in profil:

        hasil.append(
            f"• Peserta didik menunjukkan dimensi {item} selama proses pembelajaran."
        )

    return "\n".join(hasil)


# =========================================================
# PEMBELAJARAN MENDALAM
# =========================================================

def generate_deep_learning(topik):

    return f"""
Pembelajaran dirancang agar peserta didik tidak hanya mengetahui konsep
{topik}, tetapi mampu menganalisis,
mengevaluasi,
mengaitkan dengan pengalaman nyata,
serta menghasilkan solusi terhadap berbagai persoalan kontekstual.
""".strip()


# =========================================================
# SARANA
# =========================================================

def generate_sarana():

    return """
• Laptop

• LCD

• LKPD

• Buku Guru

• Buku Siswa

• Internet

• Lingkungan sekitar
""".strip()


# =========================================================
# KARAKTERISTIK
# =========================================================

def generate_karakteristik():

    return """
Peserta didik memiliki kemampuan yang beragam.

Guru menerapkan pembelajaran berdiferensiasi
sesuai kebutuhan belajar peserta didik.
""".strip()
import streamlit as st

from modules.generator import (
    generate_tujuan,
    generate_pemahaman,
    generate_pemantik,
    generate_profil,
    generate_deep_learning,
    generate_sarana,
    generate_karakteristik
)

st.title("🎯 Generator Tujuan Pembelajaran")

st.write(
    "Klik tombol di bawah untuk membuat rancangan awal secara otomatis."
)

if st.button("🚀 Generate Tujuan", use_container_width=True):

    tujuan = generate_tujuan(
        st.session_state.get("mapel", ""),
        st.session_state.get("topik", ""),
        st.session_state.get("kelas", "")
    )

    pemahaman = generate_pemahaman(
        st.session_state.get("topik", "")
    )

    pemantik = generate_pemantik(
        st.session_state.get("topik", "")
    )

    profil = generate_profil(
        st.session_state.get("profil", [])
    )

    deep = generate_deep_learning(
        st.session_state.get("topik", "")
    )

    sarana = generate_sarana()

    karakteristik = generate_karakteristik()

    st.session_state["tujuan"] = tujuan
    st.session_state["pemahaman"] = pemahaman
    st.session_state["pemantik"] = pemantik
    st.session_state["profil_hasil"] = profil
    st.session_state["deep"] = deep
    st.session_state["sarana"] = sarana
    st.session_state["karakteristik"] = karakteristik

st.subheader("1. Tujuan Pembelajaran")
st.text_area("", st.session_state.get("tujuan", ""), height=120)

st.subheader("2. Pemahaman Bermakna")
st.text_area("", st.session_state.get("pemahaman", ""), height=100)

st.subheader("3. Pertanyaan Pemantik")
st.text_area("", st.session_state.get("pemantik", ""), height=120)

st.subheader("4. Profil Lulusan")
st.text_area("", st.session_state.get("profil_hasil", ""), height=120)

st.subheader("5. Pembelajaran Mendalam")
st.text_area("", st.session_state.get("deep", ""), height=140)

st.subheader("6. Sarana dan Prasarana")
st.text_area("", st.session_state.get("sarana", ""), height=120)

st.subheader("7. Karakteristik Peserta Didik")
st.text_area("", st.session_state.get("karakteristik", ""), height=120)
