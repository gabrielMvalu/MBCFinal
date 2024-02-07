# pages/Completare_Doc.py
import streamlit as st
import pandas as pd
import re
from docx import Document
from constatator import extrage_informatii_firma, extrage_asociati_admini, extrage_situatie_angajati, extrage_coduri_caen
from datesolicitate import extrage_date_solicitate
from bilantsianaliza import extrage_date_bilant, extrage_date_contpp, extrage_indicatori_financiari
from serviciisiutilaje import extrage_pozitii, coreleaza_date


st.set_page_config(layout="wide")

st.header(':blue[Procesul de înlocuire a Placeholder-urilor]', divider='rainbow')

caen_nr_extras = None
document_succes = False   # variabilă pentru a ține evidența succesului procesării primului document

col1, col2 = st.columns(2)

with col1:
    uploaded_doc1 = st.file_uploader("Încărcați fișierul Date Solicitate", type=["xlsx"], key="dateSolicitate")
    st.success("Adaugati DATE SOLICITATE")
    if uploaded_template is not None:
        
        datesolicitate_doc = pd.read_excel(uploaded_doc1)
        date_din_xlsx_date_solicitate = extrage_date_solicitate(datesolicitate_doc)
        
        caen_extras = date_din_xlsx_date_solicitate.get('Cod CAEN', 'Cod CAEN necunoscut')
        firma = date_din_xlsx_date_solicitate.get('Denumirea firmei SRL', 'Firmă necunoscută')
        
        match = re.search(r'CAEN (\d+)', caen_extras)
        # Verificăm si extragem numărul CAEN
        if match:
            caen_nr_extras = match.group(1)  
        else:
            caen_nr_extras = None 
        
        st.success(f"Vom începe prelucrarea firmei: {firma} cu prelucrarea pe codul CAEN: {caen_nr_extras} - {caen_extras}")

        document_succes = True  # Setăm variabila pe True pentru a indica că primul document a fost procesat cu succes

# Utilizarea celei de-a doua coloane pentru încărcarea celui de-al doilea document, dacă primul a fost procesat cu succes
with col2:
    if document_succes:
        uploaded_file2 = st.file_uploader("Încărcați al doilea document", type=["docx"], key="RaportInterogare")

