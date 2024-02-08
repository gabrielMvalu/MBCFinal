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

# Inițializare variabile în st.session_state dacă nu există deja
if 'caen_nr_extras_foi' not in st.session_state:
    st.session_state.caen_nr_extras_foi = None
if 'judet_foi' not in st.session_state:
    st.session_state.judet_foi = None
if 'noua_veche_foi' not in st.session_state:
    st.session_state.noua_veche_foi = None

document_succes = False  
document2_succes = False  

datesolicitate_doc = None
date_din_xlsx_date_solicitate = None 

col1, col2, col3 = st.columns(3)

with col1:
    uploaded_doc1 = st.file_uploader("Încărcați fișierul Date Solicitate", type=["xlsx"], key="dateSolicitate")
    st.success("Date Solicitate")
    
    if uploaded_doc1 is not None:
        datesolicitate_doc = pd.read_excel(uploaded_doc1)
        date_din_xlsx_date_solicitate = extrage_date_solicitate(datesolicitate_doc)
        
        caen_extras = date_din_xlsx_date_solicitate.get('Cod CAEN', 'Cod CAEN necunoscut')
        st.session_state.judet_foi = date_din_xlsx_date_solicitate.get('Județ', 'Judet necunoscut')
        st.session_state.noua_veche_foi = date_din_xlsx_date_solicitate.get('Activitate', 'Activitate necunoscuta')
        
        firma = date_din_xlsx_date_solicitate.get('Denumirea firmei SRL', 'Firmă necunoscută')
        
        match = re.search(r'CAEN (\d+)', caen_extras)
        if match:
            st.session_state.caen_nr_extras_foi = match.group(1)  
        else:
            st.session_state.caen_nr_extras_foi = None 
        
        st.success(f"Vom începe prelucrarea firmei: {firma} cu prelucrarea pe codul CAEN: {st.session_state.caen_nr_extras_foi} - {caen_extras} ")

        document_succes = True  

with col2:
    if document_succes:
        uploaded_doc2 = st.file_uploader("Încărcați al doilea document", type=["docx"], key="RaportInterogare")
        st.info("Raport interogare")

        if uploaded_doc2 is not None:
            template_doc = Document(uploaded_doc2)
            st.toast('Incepem procesarea Planului de afaceri', icon='⭐') 
            ion = date_din_xlsx_date_solicitate.get('Cod CAEN', 'Cod CAEN necunoscut')
            st.info(f"Vom începe prelucrarea firmei: {ion} cu prelucrarea pe codul CAEN: {st.session_state.caen_nr_extras_foi} ")
            document2_succes = True

with col3:
    if document2_succes:
        uploaded_doc3 = st.file_uploader("Încărcați al 3 lea document", type=["xlsx"], key="AnalizaMacheta")
        st.success("Incepem prelucrarea analizei")
        
        if uploaded_doc3 is not None:
            template_doc3 = pdf.read_excel(uploaded_doc3)
            st.success(f"Vom începe prelucrarea analizei financiare CAEN: {st.session_state.caen_nr_extras_foi} JUDET: {st.session_state.judet_foi} NOUA SAU VECHE: {st.session_state.noua_veche_foi} ")


            
