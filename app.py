import streamlit as st
import pandas as pd
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
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
    gauth.LoadCredentialsFile("credentials.json")  # Gunakan token yang sudah ada
    if gauth.credentials is None:
        gauth.LocalWebserverAuth()  # Login jika token tidak ditemukan
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.Authorize()

    return GoogleDrive(gauth)


def export_pdf(data, filename):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=50, bottomMargin=30)
    elements = []
    
    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    title_style.alignment = TA_CENTER
    subtitle_style = styles["Heading2"]
    subtitle_style.alignment = TA_CENTER
    answer_style = styles["Normal"]
    answer_style.alignment = TA_JUSTIFY

    elements.append(Paragraph("SPOT", title_style))
    elements.append(Spacer(1, 2))
    elements.append(Paragraph("Summary of Progress & Objectives Tracker", subtitle_style))
    elements.append(Spacer(1, 16))

    for key, value in data.items():
        elements.append(Paragraph(f"<b>{key}</b>", styles["Heading2"]))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(str(value), answer_style))
        elements.append(Spacer(1, 12))

    doc.build(elements)
    buffer.seek(0)

    with open(filename, "wb") as f:
        f.write(buffer.getbuffer())

    return filename, buffer


if "data" not in st.session_state:
    st.session_state.data = {}

st.markdown(
    """
    <h1 style='text-align: center;'>SPOT</h1>
    <h3 style='text-align: center;'>Summary of Progress & Objectives Tracker</h3>
    <br>
    """,
    unsafe_allow_html=True
)

with st.form("data_form"):
    nama_tim = st.text_input("Nama Tim")
    ketua = st.text_input("Nama Ketua")
    coach = st.text_input("Nama Coach")
    jumlah_anggota = st.number_input("Jumlah Anggota", min_value=1)
    objective = st.text_area("Objective/Goal Tahunan")
    progress_bulanan = st.text_area("Progress Bulanan dengan Indikator Pencapaian")
    target_triwulanan = st.text_area("Target Triwulanan dengan Indikator Pencapaian")
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

if st.session_state.data:
    st.write("### Data Tersimpan")
    st.json(st.session_state.data)

if st.button("Ekspor & Upload ke Google Drive"):
    if not nama_tim:
        st.warning("Harap isi Nama Tim terlebih dahulu.")
    else:
        bulan = datetime.datetime.now().strftime("%B")
        filename = f"{nama_tim}_{bulan}.pdf"
        pdf_path, pdf_buffer = export_pdf(st.session_state.data, filename)

        drive = authenticate_drive()
        file_drive = drive.CreateFile({"title": filename, "parents": [{"id": FOLDER_ID}]})
        file_drive.SetContentFile(pdf_path)
        file_drive.Upload()

        file_drive.InsertPermission({'type': 'anyone', 'value': 'anyone', 'role': 'reader'})
        gdrive_link = f"https://drive.google.com/file/d/{file_drive['id']}/view"

        st.success("PDF berhasil diunggah ke Google Drive!")
        st.markdown(f'<a href="{gdrive_link}" target="_blank">Lihat File di Google Drive</a>', unsafe_allow_html=True)
        st.download_button("Download PDF", pdf_buffer, file_name=filename, mime="application/pdf")
        os.remove(pdf_path)
