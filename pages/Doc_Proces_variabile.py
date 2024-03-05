import streamlit as st
from docx import Document
import io

def adauga_paragraf_la_inceput(document):
    # Adaugă un nou paragraf la începutul documentului
    p = document.paragraphs[0].insert_paragraph_before("Text adăugat la începutul documentului")
    p.add_run("Acesta este un text clar și vizibil adăugat.")

def main():
    st.header("Adaugă Text la Începutul Documentului și Descarcă")

    uploaded_file = st.file_uploader("Încărcați documentul DOCX aici:", type="docx")
    if uploaded_file is not None:
        document = Document(uploaded_file)
        adauga_paragraf_la_inceput(document)
        
        # Salvăm documentul modificat într-un buffer pentru descărcare
        buffer = io.BytesIO()
        document.save(buffer)
        buffer.seek(0)
        
        st.download_button(label="Descarcă Documentul Modificat",
                           data=buffer,
                           file_name="document_modificat.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == "__main__":
    main()

