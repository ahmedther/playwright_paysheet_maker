from PyPDF2 import PdfReader, PdfWriter
from env.env import PDF_PASSWORD


class PyPDF2_Handler:
    def __init__(self):
        pass

    def set_pdf_password(self, pdf_path):
        with open(pdf_path, "rb") as file:
            reader = PdfReader(file)
            writer = PdfWriter()

            # Copy all pages to the writer
            for page_num in range(len(reader.pages)):
                writer.add_page(reader.pages[page_num])

            # Add a password to the PDF
            writer.encrypt(PDF_PASSWORD)

            # Save the new PDF with the password
            with open(pdf_path, "wb") as output_file:
                writer.write(output_file)
