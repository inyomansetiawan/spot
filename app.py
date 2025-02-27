import streamlit as st
import pandas as pd
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
import io
import os
import requests

# ID folder Google Drive tujuan (GANTI DENGAN ID FOLDER ANDA)
FOLDER_ID = "1oE3xhsmyW_zeMRyP9inST20fiob6rylt"
API_KEY = "AIzaSyBPVpnju6gMRPXk8qUZJJHPW9c6ua6DDvE"  # Ganti dengan API Key Anda
UPLOAD_URL = f"https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&key={API_KEY}"


def export_pdf(data, filename):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=50, bottomMargin=30)
    elements = []
    
    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    title_style.alignment = TA_CENTER
    subtitle_style = styles["Heading2"]
    subtitle_style.alignment = TA_CENTER

    answer_style1 = ParagraphStyle("answer_style1", parent=styles["Normal"], alignment=TA_CENTER)
    answer_style2 = ParagraphStyle("answer_style2", parent=styles["Normal"], alignment=TA_JUSTIFY)

    title = Paragraph("SPOT", title_style)
    subtitle = Paragraph("Summary of Progress & Objectives Tracker", subtitle_style)
    elements.append(title)
    elements.append(Spacer(1, 2))
    elements.append(subtitle)
    elements.append(Spacer(1, 16))

    for idx, (key, value) in enumerate(data.items()):
        question = Paragraph(f"<b>{key}</b>", styles["Heading2"])
        elements.append(question)
        elements.append(Spacer(1, 6))
        answer_style = answer_style1 if idx <= 3 else answer_style2
        answer = Paragraph(str(value), answer_style)
        elements.append(answer)
        elements.append(Spacer(1, 12))

    doc.build(elements)
    buffer.seek(0)
    return buffer


def upload_to_drive(file_buffer, filename):
    headers = {"Authorization": f"Bearer {API_KEY}"}
    metadata = {
        "name": filename,
        "parents": [FOLDER_ID],
    }
    files = {
        "metadata": (None, str(metadata), "application/json"),
        "file": (filename, file_buffer, "application/pdf"),
    }
    response = requests.post(UPLOAD_URL, headers=headers, files=files)
    if response.status_code == 200:
        file_id = response.json().get("id")
        return f"https://drive.google.com/file/d/{file_id}/view"
    else:
        return None

# Streamlit UI
st.title("SPOT - Summary of Progress & Objectives Tracker")

with st.form("data_form"):
    nama_tim = st.text_input("Nama Tim")
    ketua = st.text_input("Nama Ketua")
    coach = st.text_input("Nama Coach")
    jumlah_anggota = st.number_input("Jumlah Anggota", min_value=1)
    objective = st.text_area("Objective/Goal Tahunan")
    progress_bulanan = st.text_area("Progress Bulanan")
    target_triwulanan = st.text_area("Target Triwulanan")
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

if st.session_state.get("data"):
    st.write("### Data Tersimpan")
    st.json(st.session_state.data)

if st.button("Ekspor & Unggah ke Google Drive"):
    if not nama_tim:
        st.warning("Harap isi Nama Tim terlebih dahulu.")
    else:
        bulan = datetime.datetime.now().strftime("%B")
        filename = f"{nama_tim}_{bulan}.pdf"
        pdf_buffer = export_pdf(st.session_state.data, filename)
        gdrive_link = upload_to_drive(pdf_buffer, filename)
        
        if gdrive_link:
            st.success("PDF berhasil diunggah ke Google Drive!")
            st.markdown(f"[Lihat File di Google Drive]({gdrive_link})")
            st.download_button("Download PDF", pdf_buffer, file_name=filename, mime="application/pdf")
        else:
            st.error("Gagal mengunggah ke Google Drive. Periksa API Key Anda.")
