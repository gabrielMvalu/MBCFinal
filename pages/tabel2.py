import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from io import BytesIO


st.title(':blue[Transformare Date Excel]')

uploaded_file = st.file_uploader("Alegeți fișierul Excel:", type='xlsx')
uploaded_word_file = st.file_uploader("Încarcă documentul Word", type=['docx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name='P. FINANCIAR')


    # Lista cu valorile pe care dorim să le excludem din coloana B
    valori_de_exclus1 = [
        "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati",
        "Rampa mobila",
        "Toaleta ecologica",
        "Total active corporale",
        "Total active necorporale",
        "Publicitate",
        "Consultanta management",
        "Consultanta achizitii",
        "Consultanta scriere",
        "Cursuri instruire personal",
    ]

    valori_de_exclus2 = [
        "Total active corporale",
        "Total active necorporale",
        "Publicitate",
        "Consultanta management",
        "Consultanta achizitii",
        "Consultanta scriere",
    ]    
    
    # Filtrăm DataFrame-ul pentru a exclude rândurile cu valorile specificate în lista 'valori_de_exclus'
    df_filtrat_pt_subtotal1 = df[df.iloc[:, 1].notna() & ~df.iloc[:, 1].isin(valori_de_exclus1) & (df.iloc[:, 1] != 0) & (df.iloc[:, 1] != '-')]
    df_filtrat_pt_subtotal1 = df[df.iloc[:, 1].notna() & ~df.iloc[:, 1].isin(valori_de_exclus2) & (df.iloc[:, 1] != 0) & (df.iloc[:, 1] != '-')]

    subtotal_1 = 0
    subtotal_2 = 0

    for i, row in enumerate(df_filtrat_pt_subtotal1.itertuples(), 1):
    subtotal_1 += row[4]

    for i, row in enumerate(df_filtrat_pt_subtotal2.itertuples(), 1):
    subtotal_1 += row[4]

    

    stop_row = None

    for index, row in df.iterrows():
        if row[1] == 'Total proiect':
            stop_row = index
            break  
   
    if stop_row is not None:
      
        valoare_total_proiect = df.iloc[stop_row, 4]
     
    else:
        pass

st.write(f"Valoare totala proiect: {valoare_total_proiect}, Subtotal_1: {subtotal_1}, Subtotal_2 {subtotal_2}")
