import streamlit as st
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import io
import fitz  # PyMuPDF for PDF splitting
from docx import Document  # for converting Word to PDF
import pandas as pd  # for converting Excel to PDF
from pptx import Presentation  # for converting PowerPoint to PDF
from PIL import Image  # for converting images to PDF
from template import new

# Function to merge PDFs
def merge_pdfs(pdf_files):
    merger = PdfMerger()
    for pdf_file in pdf_files:
        merger.append(pdf_file)
    merger.write("merged_output.pdf")
    merger.close()

# Function to compress PDF
def compress_pdf(input_pdf):
    output_pdf = io.BytesIO()
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    
    for page in reader.pages:
        page.compress_content_streams()
        writer.add_page(page)
    
    writer.write(output_pdf)
    output_pdf.seek(0)
    return output_pdf


# Function to split PDF
def split_pdf(pdf_file):
    doc = fitz.open(pdf_file)
    for i in range(len(doc)):
        output = f"{i}_output.pdf"
        page = doc.load_page(i)
        page.extract_text()
        page.write(output)
    doc.close()

# Function to convert Word to PDF
def convert_to_pdf_word(word_file):
    doc = Document(word_file)
    pdf = io.BytesIO()
    doc.save(pdf)
    return pdf

# Function to convert Excel to PDF
def convert_to_pdf_excel(excel_file):
    df = pd.read_excel(excel_file)
    pdf = io.BytesIO()
    df.to_excel(pdf, index=False)
    return pdf

# Function to convert PowerPoint to PDF
def convert_to_pdf_powerpoint(ppt_file):
    prs = Presentation(ppt_file)
    pdf = io.BytesIO()
    prs.save(pdf)
    return pdf

# Function to convert image to PDF
def convert_to_pdf_image(image_file):
    img = Image.open(image_file)
    pdf = io.BytesIO()
    img.save(pdf, format="PDF")
    return pdf

# Main function for merge PDF page
def merge_page():
    st.title("Merge PDF")
    st.write("Upload the PDF files you want to merge:")
    uploaded_files = st.file_uploader("Choose PDF files", accept_multiple_files=True, type="pdf")

    if st.button("Merge PDFs"):
        if uploaded_files:
            merge_pdfs(uploaded_files)
            st.success("PDFs merged successfully!")
            st.download_button(
                label="Download Merged PDF",
                data=open("merged_output.pdf", "rb").read(),
                file_name="merged_output.pdf",
                mime="application/pdf",
            )
        else:
            st.warning("Please upload at least one PDF file.")

# Main function for compress PDF page
def compress_page():
    st.title("Compress PDF")
    pdf_file = st.file_uploader("Upload the PDF file you want to compress:", type="pdf")

    if st.button("Compress PDF") and pdf_file:
        with st.spinner("Compressing PDF..."):
            compressed_pdf = compress_pdf(pdf_file)
        st.success("PDF compressed successfully!")
        st.download_button(
            label="Download Compressed PDF",
            data=compressed_pdf,
            file_name="compressed_output.pdf",
            mime="application/pdf",
        )

# Main function for split PDF page
def split_page():
    st.title("Split PDF")
    pdf_file = st.file_uploader("Upload the PDF file you want to split:", type="pdf")

    if st.button("Split PDF") and pdf_file:
        split_pdf(pdf_file)
        st.success("PDF split successfully!")
        st.write("The PDF has been split into individual pages. You can download each page separately.")

# Main function for convert from PDF page
def convert_from_pdf_page():
    st.title("Convert From PDF")
    st.write("Select the format you want to convert the PDF to:")

    selected_format = st.selectbox("Select Format", ["Word", "Excel", "PowerPoint", "Image"])

    pdf_file = st.file_uploader("Upload the PDF file you want to convert:", type="pdf")

    if st.button(f"Convert to {selected_format}") and pdf_file:
        if selected_format == "Word":
            converted_pdf = convert_to_pdf_word(pdf_file)
            st.success("PDF converted to Word successfully!")
        elif selected_format == "Excel":
            converted_pdf = convert_to_pdf_excel(pdf_file)
            st.success("PDF converted to Excel successfully!")
        elif selected_format == "PowerPoint":
            converted_pdf = convert_to_pdf_powerpoint(pdf_file)
            st.success("PDF converted to PowerPoint successfully!")
        elif selected_format == "Image":
            converted_pdf = convert_to_pdf_image(pdf_file)
            st.success("PDF converted to Image successfully!")

        st.download_button(
            label=f"Download Converted {selected_format}",
            data=converted_pdf,
            file_name=f"converted_output.{selected_format.lower()}",
            mime=f"application/{selected_format.lower()}",
        )

# Main function for convert to PDF page
def convert_to_pdf_page():
    st.title("Convert To PDF")
    st.write("Select the format of the file you want to convert to PDF:")

    selected_format = st.selectbox("Select Format", ["Word", "Excel", "PowerPoint", "Image"])

    file = st.file_uploader(f"Upload the {selected_format} file you want to convert to PDF:", type=["docx", "xlsx", "pptx", "jpg", "jpeg", "png"])

    if st.button(f"Convert to PDF") and file:
        if selected_format == "Word":
            converted_pdf = io.BytesIO()
            converted_pdf.write(convert_to_pdf_word(file).getvalue())
            st.success("Word file converted to PDF successfully!")
        elif selected_format == "Excel":
            converted_pdf = io.BytesIO()
            converted_pdf.write(convert_to_pdf_excel(file).getvalue())
            st.success("Excel file converted to PDF successfully!")
        elif selected_format == "PowerPoint":
            converted_pdf = io.BytesIO()
            converted_pdf.write(convert_to_pdf_powerpoint(file).getvalue())
            st.success("PowerPoint file converted to PDF successfully!")
        elif selected_format == "Image":
            converted_pdf = io.BytesIO()
            converted_pdf.write(convert_to_pdf_image(file).getvalue())
            st.success("Image file converted to PDF successfully!")

        st.download_button(
            label=f"Download Converted PDF",
            data=converted_pdf.getvalue(),
            file_name=f"converted_output.pdf",
            mime="application/pdf",
        )

# Main function to switch between pages
def main():
    st.sidebar.title("PDF Toolkit")
    app_mode = st.sidebar.selectbox(
        "Choose an option",
        ["Merge PDF", "Compress PDF", "Split PDF", "Convert From PDF", "Convert To PDF"]
    )

    if app_mode == "Merge PDF":
        merge_page()
    elif app_mode == "Compress PDF":
        compress_page()
    elif app_mode == "Split PDF":
        split_page()
    elif app_mode == "Convert From PDF":
        convert_from_pdf_page()
    elif app_mode == "Convert To PDF":
        convert_to_pdf_page()

if __name__ == "__main__":
    main()
