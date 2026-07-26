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
# SARANA DAN PRASARANA
# =========================================================

def generate_sarana():

    return """
• Laptop

• LCD Proyektor

• Buku Guru

• Buku Siswa

• LKPD

• Internet

• Lingkungan sekitar sebagai sumber belajar
""".strip()


# =========================================================
# KARAKTERISTIK PESERTA DIDIK
# =========================================================

def generate_karakteristik():

    return """
Peserta didik memiliki kemampuan, minat, dan gaya belajar yang beragam.

Guru menerapkan pembelajaran berdiferensiasi agar seluruh peserta didik
dapat berkembang sesuai dengan kebutuhannya.
""".strip()
