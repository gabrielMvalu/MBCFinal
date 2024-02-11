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

    stop_row = None

    for index, row in df.iterrows():
        if row[1] == 'Total proiect':
            stop_row = index
            break  
   
    if stop_row is not None:
      
        valoare_total_proiect = df.iloc[stop_row, 5]
     
    else:
        pass

st.write(f"Valoare totala proiect: {valoare_total_proiect}")
