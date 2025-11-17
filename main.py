import streamlit as st
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import pandas as pd
import tempfile
import os
os.environ["PATH"] += os.pathsep + "/usr/bin"

# ------------------------------
# Clean extracted text
# ------------------------------
def clean_text(text):
    text = text.replace("\n\n", "\n")
    text = " ".join(text.split())
    return text

# ------------------------------
# Extract text from image
# ------------------------------
def extract_from_image(image_file):
    img = Image.open(image_file)
    text = pytesseract.image_to_string(img)
    text = clean_text(text)
    return text

# ------------------------------
# Extract text from PDF
# ------------------------------
def extract_from_pdf(pdf_file):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(pdf_file.read())
        temp_pdf_path = temp_pdf.name

    pages = convert_from_path(temp_pdf_path, dpi=300)   # improved accuracy
    full_text = ""

    for page in pages:
        text = pytesseract.image_to_string(page)
        text = clean_text(text)
        full_text += text + "\n"

    return full_text, pages[0]   # return text + first page for preview

# ------------------------------
# Streamlit Dashboard
# ------------------------------
st.title("📄 OCR Dashboard — Convert Image/PDF to Excel")

uploaded_files = st.file_uploader(
    "Upload one or multiple scanned PDFs / Images",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True
)

final_data = []

if uploaded_files:

    if st.button("Extract & Generate Excel"):
        for uploaded_file in uploaded_files:

            with st.spinner(f"Extracting: {uploaded_file.name} ... ⏳"):

                # IMAGE PREVIEW
                if uploaded_file.type in ["image/jpeg", "image/png", "image/jpg"]:
                    img = Image.open(uploaded_file)
                    st.image(img, caption=f"Preview: {uploaded_file.name}", width=300)

                    extracted_text = extract_from_image(uploaded_file)

                # PDF PREVIEW (first page only)
                elif uploaded_file.type == "application/pdf":
                    extracted_text, first_page_img = extract_from_pdf(uploaded_file)
                    st.image(first_page_img, caption=f"Preview (Page 1): {uploaded_file.name}", width=300)

                # Save result
                final_data.append([uploaded_file.name, extracted_text])

                # Show extracted text
                st.subheader(f"📌 Extracted Text — {uploaded_file.name}")
                st.text_area("", extracted_text, height=200)

        # Save Excel
        df = pd.DataFrame(final_data, columns=["File Name", "Extracted Text"])
        df.to_excel("output.xlsx", index=False)

        with open("output.xlsx", "rb") as f:
            st.download_button(
                label="📥 Download Excel File",
                data=f,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.success("All files processed and Excel generated!")
