from docx import Document
import os
from datetime import datetime


def generate_word_doc(data):
    doc = Document()
    doc.add_heading(data.get("title", "Untitled Document"), level=0)

    for section in data.get("sections", []):
        doc.add_heading(section.get("heading", ""), level=1)
        doc.add_paragraph(section.get("content", ""))

    filename = f"word_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    filepath = os.path.join("outputs", filename)
    os.makedirs("outputs", exist_ok=True)
    doc.save(filepath)
    return filepath
