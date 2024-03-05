import streamlit as st
from docx import Document
from docx.shared import Pt
import io

def adauga_text_la_inceput(document, text="A"):
    # Crează un nou paragraf la începutul documentului
    paragraph = document.paragraphs[0]._element  # Accesează elementul XML al primului paragraf
    p = document.add_paragraph()  # Creează un nou paragraf
    p.add_run(text)  # Adaugă textul specificat în noul paragraf
    p_font = p.runs[0].font
    p_font.size = Pt(1)  # Setează dimensiunea fontului la 1 pentru a fi cât mai mic vizual
    p_element = p._element
    paragraph.addprevious(p_element)  # Adaugă noul paragraf înaintea primului paragraf existent

def main():
    st.header("Adaugă Text la Începutul Documentului")

    uploaded_file = st.file_uploader("Încărcați documentul DOCX aici:", type="docx")
    if uploaded_file is not None:
        document = Document(uploaded_file)
        adauga_text_la_inceput(document)
        
        # Salvăm documentul modificat într-un buffer
        buffer = io.BytesIO()
        document.save(buffer)
        buffer.seek(0)
        
        st.download_button(label="Descarcă Documentul Modificat",
                           data=buffer,
                           file_name="document_modificat.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == "__main__":
    main()

