import os
import streamlit as st
from docx import Document
import pythoncom
from win32com import client  # pip install pywin32


def replace_placeholders(template_path, replacements):
    """
    Replace placeholders in the Word template with user inputs while preserving formatting.
    """
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        for placeholder, value in replacements.items():
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, value)

    return doc


def save_word_as_pdf(word_path, pdf_path):
    """
    Convert a Word document to a PDF using Microsoft Word for perfect formatting.
    """
    pythoncom.CoInitialize()  # Initialize COM for multithreading (required by Streamlit)
    try:
        # Open Word application
        word = client.Dispatch("Word.Application")
        word.Visible = False  # Run Word in the background

        # Open the Word document
        doc = word.Documents.Open(word_path)

        # Save as PDF
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF format
        doc.Close()
        word.Quit()
    except Exception as e:
        raise RuntimeError(f"Error during Word to PDF conversion: {e}")


def main():
    st.title("Automated NDA PDF Generator")
    st.write("Fill in the details below to generate a customized NDA PDF.")

    # User inputs
    name = st.text_input("Enter Full Name")
    email = st.text_input("Enter Email")
    phone = st.text_input("Enter Phone Number")
    address = st.text_area("Enter Address (multi-line allowed)")

    # Paths for template and output files
    template_path = os.path.abspath("NDA Template - INDIA 3.docx")  # Predefined common template
    output_docx = os.path.abspath("updated_template.docx")
    output_pdf = os.path.abspath("generated_nda.pdf")

    if st.button("Generate PDF"):
        if not os.path.exists(template_path):
            st.error("Template file not found. Please ensure 'template.docx' is in the working directory.")
            return

        if not (name and email and phone and address):
            st.error("Please fill in all the fields!")
            return

        try:
            # Define the placeholder replacements
            replacements = {
                "Client Name": name,
                "Client Email": email,
                "Client Phone": phone,
                "Client Address": address,
            }

            # Replace placeholders in the template
            updated_doc = replace_placeholders(template_path, replacements)
            updated_doc.save(output_docx)  # Save updated Word document

            # Convert the updated Word document to PDF
            save_word_as_pdf(output_docx, output_pdf)

            # Provide the download link for the PDF
            with open(output_pdf, "rb") as pdf_file:
                st.download_button(
                    label="Download NDA PDF",
                    data=pdf_file,
                    file_name="generated_nda.pdf",
                    mime="application/pdf",
                )

        except RuntimeError as re:
            st.error(f"Word-to-PDF Conversion Error: {re}")
        except Exception as e:
            st.error(f"An unexpected error occurred: {e}")


if __name__ == "__main__":
    main()
