import streamlit as st
from docx import Document
import io

def rescrie_si_copiaza_documentul(document):
    # Crează un nou document
    new_doc = Document()
    
    # Copiază fiecare paragraf din documentul original în documentul nou
    for para in document.paragraphs:
        new_doc.add_paragraph(para.text)
    
    # Pentru a copia și conținutul tabelelor, este necesară o logică suplimentară
    # Acest exemplu nu include copierea tabelelor, stilurilor sau obiectelor încorporate
    # pentru a menține codul simplu și concentrat pe text
    
    return new_doc

def main():
    st.header("Rescriere Document DOCX și Descărcare")

    uploaded_file = st.file_uploader("Încărcați documentul DOCX aici:", type="docx")
    if uploaded_file is not None:
        original_doc = Document(uploaded_file)
        new_doc = rescrie_si_copiaza_documentul(original_doc)
        
        # Salvăm documentul nou creat într-un buffer pentru descărcare
        buffer = io.BytesIO()
        new_doc.save(buffer)
        buffer.seek(0)
        
        st.download_button(label="Descarcă Documentul Rescris",
                           data=buffer,
                           file_name="document_rescris.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == "__main__":
    main()

