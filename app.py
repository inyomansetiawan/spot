import streamlit as st
import pandas as pd
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
import io
import os
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# ID folder Google Drive tujuan (GANTI DENGAN ID FOLDER ANDA)
FOLDER_ID = "1oE3xhsmyW_zeMRyP9inST20fiob6rylt"

# Fungsi autentikasi Google Drive
def authenticate_drive():
    gauth = GoogleAuth()
    
    # Cek apakah kredensial sudah ada
    if os.path.exists("credentials.json"):
        gauth.LoadCredentialsFile("credentials.json")
        if gauth.credentials is None:
            gauth.CommandLineAuth()  # Mode CLI tanpa localhost
        elif gauth.access_token_expired:
            gauth.Refresh()
        else:
            gauth.Authorize()
    else:
        gauth.CommandLineAuth()  # Mode CLI tanpa localhost
        gauth.SaveCredentialsFile("credentials.json")  # Simpan token baru
    
    return GoogleDrive(gauth)



def export_pdf(data, filename):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=50, bottomMargin=30)
    elements = []
    
    styles = getSampleStyleSheet()

    # Mengatur style untuk title dan subtitle
    title_style = styles["Title"]
    title_style.alignment = TA_CENTER

    subtitle_style = styles["Heading2"]
    subtitle_style.alignment = TA_CENTER

    # Membuat custom style untuk jawaban
    answer_style1 = ParagraphStyle("answer_style1", parent=styles["Normal"], alignment=TA_CENTER)
    answer_style2 = ParagraphStyle("answer_style2", parent=styles["Normal"], alignment=TA_JUSTIFY)

    # Membuat Title dan Subtitle
    title = Paragraph("SPOT", title_style)
    subtitle = Paragraph("Summary of Progress & Objectives Tracker", subtitle_style)

    # Menambah elemen ke dalam dokumen
    elements.append(title)
    elements.append(Spacer(1, 2))  # Spacer untuk memberi jarak
    elements.append(subtitle)
    elements.append(Spacer(1, 16))  # Spacer tambahan sebelum konten berikutnya

    for idx, (key, value) in enumerate(data.items()):
        # Heading (Pertanyaan)
        question = Paragraph(f"<b>{key}</b>", styles["Heading2"])
        elements.append(question)
        elements.append(Spacer(1, 6))

        # Pilih gaya berdasarkan indeks
        answer_style = answer_style1 if idx <= 3 else answer_style2

        # Jawaban dalam paragraf
        answer = Paragraph(str(value), answer_style)
        elements.append(answer)
        elements.append(Spacer(1, 12))  # Beri jarak antar pertanyaan

    doc.build(elements)

    buffer.seek(0)

    with open(filename, "wb") as f:
        f.write(buffer.getbuffer())

    return filename, buffer


# Simpan data sementara di session state
if "data" not in st.session_state:
    st.session_state.data = {}


# Menggunakan Markdown dengan HTML untuk center alignment
st.markdown(
    """
    <h1 style='text-align: center;'>SPOT</h1>
    <h3 style='text-align: center;'>Summary of Progress & Objectives Tracker</h3>
    <br>
    """,
    unsafe_allow_html=True
)


# Form Input
with st.form("data_form"):
    # Identitas Tim
    nama_tim = st.text_input("Nama Tim")
    ketua = st.text_input("Nama Ketua")
    coach = st.text_input("Nama Coach")
    jumlah_anggota = st.number_input("Jumlah Anggota", min_value=1)

    # Objective/Goal Tahunan
    objective = st.text_area("Objective/Goal Tahunan")

    # Progress Bulanan vs Target Triwulanan
    st.subheader("Progress Bulanan vs Target Triwulanan")
    progress_bulanan = st.text_area("Progress Bulanan dengan Indikator Pencapaian")
    target_triwulanan = st.text_area("Target Triwulanan dengan Indikator Pencapaian")

    # Hasil Retrospective
    st.subheader("Hasil Retrospektif")
    what_went_well = st.text_area("What went Well?")
    what_can_be_improved = st.text_area("What can be Improved?")
    action_points = st.text_area("Action Points")

    submitted = st.form_submit_button("Simpan Data")

    if submitted:
        st.session_state.data = {
            "Nama Tim": nama_tim,
            "Ketua": ketua,
            "Coach": coach,
            "Jumlah Anggota": jumlah_anggota,
            "Objective/Goal Tahunan": objective,
            "Progress Bulanan": progress_bulanan,
            "Target Triwulanan": target_triwulanan,
            "What went well": what_went_well,
            "What can be improved": what_can_be_improved,
            "Action Points": action_points
        }
        st.success("Data berhasil disimpan!")

# Tampilkan Data
if st.session_state.data:
    st.write("### Data Tersimpan")
    st.json(st.session_state.data)

# Tombol ekspor ke PDF & Upload ke Google Drive
if st.button("Ekspor & Unggah ke Google Drive"):
    if not nama_tim:
        st.warning("Harap isi Nama Tim terlebih dahulu.")
    else:
        # Format nama file: nama_tim_bulan.pdf
        bulan = datetime.datetime.now().strftime("%B")  # Nama bulan dalam bahasa Inggris
        filename = f"{nama_tim}_{bulan}.pdf"

        # Buat PDF
        pdf_path, pdf_buffer = export_pdf(st.session_state.data, filename)

        # Upload ke Google Drive
        drive = authenticate_drive()
        file_drive = drive.CreateFile({"title": filename, "parents": [{"id": FOLDER_ID}]})
        file_drive.SetContentFile(pdf_path)
        file_drive.Upload()

        # Dapatkan link file di Google Drive
        file_drive.InsertPermission({'type': 'anyone', 'value': 'anyone', 'role': 'reader'})
        gdrive_link = f"https://drive.google.com/file/d/{file_drive['id']}/view"

        st.success(f"PDF berhasil diunggah ke Google Drive!")
        st.markdown(f"[Lihat File di Google Drive]({gdrive_link})")

        # Tombol download PDF
        st.download_button("Download PDF", pdf_buffer, file_name=filename, mime="application/pdf")

        # Hapus file lokal setelah berhasil diunggah
        os.remove(pdf_path)