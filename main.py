import streamlit as st
import tempfile
import os
from pdf2docx import Converter
import fitz  # PyMuPDF

st.set_page_config(
    page_title="PDF to Word Converter",
    page_icon="📄",
    layout="centered"
)

st.title("📄 PDF to Word Converter")
st.write("Upload a PDF file and convert it into an editable Microsoft Word document.")

uploaded_file = st.file_uploader(
    "Choose a PDF file",
    type=["pdf"]
)

def validate_pdf(path):
    try:
        doc = fitz.open(path)
        doc.close()
        return True
    except Exception:
        return False


if uploaded_file is not None:

    with tempfile.TemporaryDirectory() as tmpdir:

        pdf_path = os.path.join(tmpdir, uploaded_file.name)

        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        if not validate_pdf(pdf_path):
            st.error("The uploaded file is not a valid PDF.")
            st.stop()

        output_filename = os.path.splitext(uploaded_file.name)[0] + ".docx"
        output_path = os.path.join(tmpdir, output_filename)

        if st.button("Convert to Word"):

            with st.spinner("Converting... Please wait."):

                try:
                    cv = Converter(pdf_path)
                    cv.convert(output_path)
                    cv.close()

                    with open(output_path, "rb") as docx_file:
                        st.success("Conversion completed successfully!")

                        st.download_button(
                            label="⬇ Download Word File",
                            data=docx_file,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )

                except Exception as e:
                    st.error(f"Conversion failed.\n\n{e}")
