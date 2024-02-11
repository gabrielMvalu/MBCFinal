import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from io import BytesIO


st.title(':blue[Transformare Date Excel]')

stop_text = None


uploaded_file = st.file_uploader("Alegeți fișierul Excel:", type='xlsx')
uploaded_word_file = st.file_uploader("Încarcă documentul Word", type=['docx'])



if uploaded_file is not None:
    
    stop_text = 'Total proiect'
    
    df = pd.read_excel(uploaded_file, sheet_name='P. FINANCIAR')
    stop_index = df[df.iloc[:, 1] == stop_text].index.min()
    df_filtrat = df.iloc[3:stop_index] if pd.notna(stop_index) else df.iloc[3:]
    df_filtrat = df_filtrat[df_filtrat.iloc[:, 1].notna() & (df_filtrat.iloc[:, 1] != 0) & (df_filtrat.iloc[:, 1] != '-')]


    # Lista cu valorile pe care dorim să le excludem din coloana B
    valori_de_exclus = [
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

    # Filtrăm DataFrame-ul pentru a exclude rândurile cu valorile specificate în lista 'valori_de_exclus'
    df_filtrat_pt_subtotal1 = df_filtrat[~df_filtrat.iloc[:, 1].isin(valori_de_exclus)]
    
    st.dataframe(df_filtrat_pt_subtotal1)
    
    subtotal_1 = df_filtrat_pt_subtotal1.iloc[:, 3].sum()
    
    # Afișăm subtotal_1
    st.write(f"Subtotal 1: {subtotal_1}")


stop_row = None
subtotal_2 = 0  # Inițializează subtotal_2 cu 0

elemente_specifice = [
    "Cursuri instruire personal",
    "Toaleta ecologica",
    "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati",
    "Rampa mobila"
]


for index, row in df.iterrows():
    # Verifică dacă ai ajuns la 'Total proiect'
    if row[1] == 'Total proiect':
        stop_row = index
        break  # Ieși din buclă când găsești 'Total proiect'
    
    # Verifică dacă rândul curent conține unul dintre elementele specifice
    if row[1] in elemente_specifice:
        # Adună valoarea din coloana dorită la subtotal_2
        subtotal_2 += row[4]  # Presupunând că valorile sunt în coloana cu indexul 4

# Verifică dacă ai găsit 'Total proiect' și calculează valoarea totală a proiectului
if stop_row is not None:
    valoare_total_proiect = df.iloc[stop_row, 4]
else:
    pass  # Poți gestiona cazul în care 'Total proiect' nu este găsit

# Afișează subtotal_2 și valoare_total_proiect

st.write(f"Total: {valoare_total_proiect}")
st.write(f"Subtotal 2: {subtotal_2}")
st.write(f"Subtotal 1: {subtotal_1}")
