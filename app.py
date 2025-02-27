import streamlit as st
import os
import subprocess
from PIL import Image
from pdf2docx import Converter
import pandas as pd

try:
    from docx2pdf import convert
except ImportError:
    st.error("docx2pdf module not installed. Run: `pip install docx2pdf`")
    st.stop()

supported_formats = {
    ".pdf": ".docx",
    ".docx": ".pdf",
    ".xlsx": ".csv",
    ".csv": ".xlsx",
}

icon_map = {
    ".pdf": "icons/pdf_icon.png",
    ".docx": "icons/docx_icon.png",
    ".xlsx": "icons/xlsx_icon.png",
    ".csv": "icons/csv_icon.png"
}
default_icon = "icons/default_icon.png"

# Function to get file icon as PIL Image
def get_file_icon(extension):
    icon_path = icon_map.get(extension, default_icon)
    return Image.open(icon_path) if os.path.exists(icon_path) else Image.open(default_icon)

# Function to open files (only works locally)
def open_file(file_path):
    if file_path and os.path.exists(file_path):
        try:
            subprocess.run(["open" if os.name == "posix" else "start", file_path], shell=True)
        except Exception as e:
            st.error(f"Failed to open file: {e}")

# File conversion functions
def convert_pdf_to_docx(pdf_path, docx_path):
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

def convert_docx_to_pdf(docx_path):
    try:
        convert(docx_path)
    except Exception as e:
        st.error(f"Failed to convert DOCX to PDF: {e}")

def convert_excel_to_csv(excel_path, csv_path):
    df = pd.read_excel(excel_path)
    df.to_csv(csv_path, index=False)

def convert_csv_to_excel(csv_path, excel_path):
    df = pd.read_csv(csv_path)
    df.to_excel(excel_path, index=False)

# Streamlit UI
st.title("File Converter Web App")

uploaded_file = st.file_uploader("Upload a file", type=["pdf", "docx", "xlsx", "csv"])

if uploaded_file:
    file_ext = os.path.splitext(uploaded_file.name)[1]

    # Show Uploaded File Icon & Name
    col1, col2 = st.columns([1, 4])
    with col1:
        st.image(get_file_icon(file_ext), width=50)
    with col2:
        st.write(f"**Selected File:** {uploaded_file.name}")

    if file_ext not in supported_formats:
        st.error("Unsupported file format")
    else:
        output_format = supported_formats[file_ext]
        converted_file_name = os.path.splitext(uploaded_file.name)[0] + output_format
        convert_button = st.button(f"Convert to {output_format.upper()}")

        if convert_button:
            with st.spinner("Converting..."):
                try:
                    temp_input_path = f"temp{file_ext}"
                    temp_output_path = f"temp{output_format}"

                    with open(temp_input_path, "wb") as f:
                        f.write(uploaded_file.read())

                    if file_ext == ".pdf":
                        convert_pdf_to_docx(temp_input_path, temp_output_path)
                    elif file_ext == ".docx":
                        convert_docx_to_pdf(temp_input_path)
                    elif file_ext == ".xlsx":
                        convert_excel_to_csv(temp_input_path, temp_output_path)
                    elif file_ext == ".csv":
                        convert_csv_to_excel(temp_input_path, temp_output_path)

                    st.success("Conversion successful!")

                    # Show Converted File Icon & Name
                    col3, col4 = st.columns([1, 4])
                    with col3:
                        st.image(get_file_icon(output_format), width=50)
                    with col4:
                        st.write(f"**Converted File:** {converted_file_name}")

                    # Add Download Button
                    st.download_button(
                        label="Download Converted File",
                        data=open(temp_output_path, "rb"),
                        file_name=converted_file_name
                    )

                except Exception as e:
                    st.error(f"Conversion failed: {e}")
