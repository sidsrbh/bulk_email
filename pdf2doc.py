import fitz  # PyMuPDF
from docx import Document

def pdf_to_docx(pdf_path, docx_path):
    document = Document()
    pdf_document = fitz.open(pdf_path)
    
    for page in pdf_document:
        text = page.get_text()
        document.add_paragraph(text)

    document.save(docx_path)
    pdf_document.close()

# Example usage
pdf_to_docx('input.pdf', 'output.docx')
