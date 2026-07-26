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
