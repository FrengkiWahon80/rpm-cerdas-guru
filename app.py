import streamlit as st

st.set_page_config(
    page_title="RPM CERDAS AI",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ===========================
# SESSION STATE
# ===========================

default = {
    "guru":"",
    "sekolah":"",
    "mapel":"",
    "kelas":"",
    "fase":"",
    "semester":"",
    "tahun":"",
    "topik":"",
    "cp":""
}

for k,v in default.items():
    if k not in st.session_state:
        st.session_state[k]=v

# ===========================
# SIDEBAR
# ===========================

with st.sidebar:

    st.title("📘 RPM CERDAS AI")

    st.success("Versi 1.0")

    st.markdown("---")

    st.write(
    """
    Generator

    ✅ RPM

    ✅ Modul Ajar

    ✅ LKPD

    ✅ Asesmen

    ✅ Word Export
    """
    )

st.title("🤖 RPM CERDAS AI")

st.subheader("Generator Rencana Pembelajaran Mendalam")

st.info("Silakan pilih menu di sidebar.")

st.markdown("---")

st.write("""
Selamat datang di RPM CERDAS AI.

Aplikasi ini dibuat khusus untuk Guru SMP Kurikulum Merdeka.

Versi Lokal (Offline)
""")
