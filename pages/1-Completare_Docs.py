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

nr_CAEN = None
document_succes = False  
document2_succes = False
document3_succes = False

col1, col2 = st.columns(2)

with col1:
    uploaded_doc1 = st.file_uploader("Încărcați fișierul Date Solicitate", type=["xlsx"], key="dateSolicitate")
    if uploaded_doc1 is not None:
        df = pd.read_excel(uploaded_doc1)
        solicitate_data = extrage_date_solicitate(df)

        firma = solicitate_data.get('Denumirea firmei SRL', 'N/A')
        nr_CAEN = solicitate_data.get('Doar nr CAEN','N/A')
        st.session_state.codCAEN = nr_CAEN
        st.success(f"Primul pas, pentru: {firma}, completat.")
        document_succes = True
        #st.json({"Date extrase": solicitate_data})


with col2:
    if document_succes:
        uploaded_doc2 = st.file_uploader(f"Încărcați Raportul Interogare al {firma}", type=["docx"], key="RaportInterogare")
        if uploaded_doc2 is not None:
            constatator_doc = Document(uploaded_doc2)
    
            informatii_firma = extrage_informatii_firma(constatator_doc)
            asociati_info, administratori_info = extrage_asociati_admini(constatator_doc)
            situatie_angajati = extrage_situatie_angajati(constatator_doc)
            full_text_constatator = "\n".join([p.text for p in constatator_doc.paragraphs])
            coduri_caen = extrage_coduri_caen(full_text_constatator)
            
            def curata_duplicate_coduri_caen(coduri_caen):
                coduri_unice = {}
                for cod, descriere in coduri_caen:
                    coduri_unice[cod] = descriere
                return list(coduri_unice.items())
    
            coduri_caen_curatate = curata_duplicate_coduri_caen(coduri_caen)
            
            adrese_secundare_text = '\n'.join(informatii_firma.get('Adresa sediul secundar', [])) if informatii_firma.get('Adresa sediul secundar', []) else "N/A"
            asociati_text = '\n'.join(asociati_info) if asociati_info else "N/A"
            administratori_text = administratori_info if administratori_info else "N/A"
            coduri_caen_text = '\n'.join([f"{cod} - {descriere}" for cod, descriere in coduri_caen_curatate]) if coduri_caen_curatate else "N/A"    
           
            st.info(f"Prelucrarea 'Rapor Interogare' al {firma}, este completa.")
            document2_succes = True        
            
             # Afișarea datelor în format JSON
            #st.json({
            #    "Date Generale": informatii_firma,
            #    "Informații Detaliate": {"Asociați": asociati_info, "Administratori": administratori_info},
            #    "Situație Angajati": situatie_angajati,
            #    "Coduri CAEN": coduri_caen
            #}) 
    else:
        st.warning("Prima dată, încărcați și procesați, 'date solicitate.xlsx'.")



col3, col4 = st.columns(2)

with col3:
    if document2_succes:
        uploaded_doc3 = st.file_uploader("Încărcați al 3-lea document", type=["xlsx"], key="AnalizaMacheta")
        if uploaded_doc3 is not None:
            df_bilant = pd.read_excel(uploaded_doc3, sheet_name='1-Bilant')
            df_contpp = pd.read_excel(uploaded_doc3, sheet_name='2-ContPP')
            df_analiza_fin = pd.read_excel(uploaded_doc3, sheet_name='1D-Analiza_fin_indicatori')    
            df_financiar = pd.read_excel(uploaded_doc3, sheet_name='P. FINANCIAR')
            date_financiare = extrage_pozitii(df_financiar)
            if nr_CAEN != 'N/A' and date_financiare:
                rezultate_corelate, rezultate_corelate1, rezultate_corelate2 = coreleaza_date(date_financiare)
                rezultate_text = '\n'.join([rezultat for _, _, rezultat in rezultate_corelate])
                cheltuieli_text = '\n'.join([rezultat for _, _, rezultat in rezultate_corelate1])
                cantitati_corelate = [pd.to_numeric(item[1], errors='coerce') for item in rezultate_corelate]
                cantitati_corelate = [0 if pd.isna(x) else x for x in cantitati_corelate]
                numar_total_utilaje = sum(cantitati_corelate)
                rezultate_corelate, rezultate_corelate1, rezultate_corelate2 = coreleaza_date(date_financiare)
                rezultate2_text = '\n'.join([f"{descriere}" for nume, _, descriere in rezultate_corelate2])


            
            st.success(f"Vom incepe prelucrarea datelor din Analiza Financiara")
            document3_succes = True
    else:
        st.warning("Vă rugăm să încărcați și să procesați 'Date Solicitate', apoi 'Raport Interogare'.")


with col4:
    if document3_succes:
        uploaded_template = st.file_uploader("Încărcați al 4-lea document", type=["docx"], key="MachetaPA")
        if uploaded_template is not None:
            
            st.info(f"Vom începe ")
            
    else:
        st.warning("Vă rugăm să încărcați și să procesați documentele din primele două coloane.")
            
