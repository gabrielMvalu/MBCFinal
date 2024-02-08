# pages/Completare_Doc.py
import streamlit as st
import pandas as pd
import re
from docx import Document
from constatator import extrage_informatii_firma, extrage_asociati_admini, extrage_situatie_angajati, extrage_coduri_caen
from datesolicitate import extrage_date_solicitate, extrage_date_suplimentare
from bilantsianaliza import extrage_date_bilant, extrage_date_contpp, extrage_indicatori_financiari
from serviciisiutilaje import extrage_pozitii, coreleaza_date

st.set_page_config(layout="wide")

st.header(':blue[Procesul de înlocuire a Placeholder-urilor]', divider='rainbow')


document_succes = False  
document2_succes = False
document3_succes = False

col1, col2 = st.columns(2)

with col1:
    uploaded_doc1 = st.file_uploader("Încărcați fișierul Date Solicitate", type=["xlsx"], key="dateSolicitate")
    if uploaded_doc1 is not None:
        
        df = pd.read_excel(uploaded_doc1)
        solicitate_data = extrage_date_solicitate(df)
        st.success(f"Vom prelucra")
        document_succes = True
        st.json({"Date Solicitate": solicitate_data})



with col2:
    if document_succes:
        uploaded_doc2 = st.file_uploader("Încărcați al doilea document", type=["docx"], key="RaportInterogare")
        if uploaded_doc2 is not None:
            template_doc = Document(uploaded_doc2)
            st.info(f"Vom începe prelucrar")
            document2_succes = True        
    else:
        st.warning("Vă rugăm să încărcați și să procesați mai întâi documentul din prima coloană.")
    
col3, col4 = st.columns(2)

with col3:
    if document2_succes:
        uploaded_doc3 = st.file_uploader("Încărcați al 3-lea document", type=["xlsx"], key="AnalizaMacheta")
        if uploaded_doc3 is not None:
            st.success(f"Vom începe")
            document3_succes = True
    else:
        st.warning("Vă rugăm să încărcați și să procesați documentele din primele două coloane.")


with col4:
    if document3_succes:
        uploaded_doc4 = st.file_uploader("Încărcați al 4-lea document", type=["docx"], key="MachetaPA")
        if uploaded_doc4 is not None:
            
            st.info(f"Vom începe ")
            
    else:
        st.warning("Vă rugăm să încărcați și să procesați documentele din primele două coloane.")
            
