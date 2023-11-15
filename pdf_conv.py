import base64
import io
from docx import Document
import fitz  # PyMuPDF
from pdf2docx import Converter


def base64_to_pdf(base64_string, output_pdf_path):
    # Convert base64 to bytes
    pdf_bytes = base64.b64decode(base64_string)

    # Create a BytesIO object
    pdf_stream = io.BytesIO(pdf_bytes)

    # Open the PDF stream with PyMuPDF
    pdf_document = fitz.open(pdf_stream)

    # Create a new PDF document
    new_pdf_document = fitz.open()

    # Iterate through each page of the original PDF
    for page_number in range(pdf_document.page_count):
        # Get the page
        page = pdf_document[page_number]

        # Create a new page in the new PDF document
        new_page = new_pdf_document.new_page(width=page.rect.width, height=page.rect.height)

        # Copy the content from the original page to the new page
        new_page.insert_page(page_number, page)

    # Save the new PDF document
    new_pdf_document.save(output_pdf_path)

    # Close the PDF documents
    pdf_document.close()
    new_pdf_document.close()

def convert_pdf_to_docx(pdf_path, docx_path):
    # Create a PDF to DOCX converter object
    cv = Converter(pdf_path)

    # Convert the PDF to DOCX
    cv.convert(docx_path, start=0, end=None)

    # Close the converter
    cv.close()

def docx_to_base64(docx_path):
    # Open the DOCX file
    doc = Document(docx_path)

    # Create a BytesIO object to store the binary data
    output_stream = io.BytesIO()

    # Save the DOCX document to the BytesIO object
    doc.save(output_stream)

    # Get the binary data from the BytesIO object
    binary_data = output_stream.getvalue()

    # Encode the binary data as base64
    base64_string = base64.b64encode(binary_data).decode('utf-8')

    return base64_string

if __name__ == "__main__":
    # Replace "base64_string" with your actual base64-encoded PDF string
    base64_string = "YOUR_BASE64_STRING_HERE"

    # Specify the output path for the new PDF file
    output_pdf_path = "output.pdf"

    # Convert base64 to PDF
    base64_to_pdf(base64_string, output_pdf_path)

    input_pdf_path = output_pdf_path

    output_docx_path = "output.docx"

    # Convert the PDF to DOCX
    convert_pdf_to_docx(input_pdf_path, output_docx_path)

    input_docx_path = output_docx_path

    # Convert the DOCX to base64
    base64_string = docx_to_base64(input_docx_path)
