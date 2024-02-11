import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from io import BytesIO


st.title(':blue[Transformare Date Excel]')



def transforma_date_tabel2(df):
    stop_text = 'Total proiect'
    
    # Filtrează și preprocesează df
    stop_index = df[df.iloc[:, 1] == stop_text].index.min()
    df_filtrat = df.iloc[3:stop_index] if pd.notna(stop_index) else df.iloc[3:]
    df_filtrat = df_filtrat[df_filtrat.iloc[:, 1].notna() & (df_filtrat.iloc[:, 1] != 0) & (df_filtrat.iloc[:, 1] != '-')]

    valori_de_eliminat = [
        "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati",
        "Rampa mobila", "Total active corporale", "Total active necorporale", 
        "Publicitate", "Consultanta management", "Consultanta achizitii", "Consultanta scriere"
    ]
    
    df_filtrat = df_filtrat[~df_filtrat.iloc[:, 1].isin(valori_de_eliminat)]
   
    nr_crt = []
    denumire = []
    um = []
    cantitate = []
    pret_unitar = []
    valoare_totala = []
    
    for i, row in df_filtrat.iterrows():
        nr_crt.append(i)
        denumire.append(row[1])  
        um.append(row[2]) 
        cantitate.append(row[3])  
        pret_unitar.append(row[4])  
        valoare_totala.append(row[5])  
    
    tabel_2 = pd.DataFrame({
        "Nr. crt.": nr_crt,
        "Denumire": denumire,
        "UM": um,
        "Cantitate": cantitate,
        "Preț unitar (fără TVA)": pret_unitar,
        "Valoare Totală (fără TVA)": valoare_totala
    })
    
    return tabel_2

tabel_2 = transforma_date_tabel2(df)

st.dataframe(tabel_2)


stop_text = None

uploaded_file = st.file_uploader("Alegeți fișierul Excel:", type='xlsx')
uploaded_word_file = st.file_uploader("Încarcă documentul Word", type=['docx'])

if uploaded_file is not None:
    stop_text = 'Total proiect'
    df = pd.read_excel(uploaded_file, sheet_name='P. FINANCIAR')
    stop_index = df[df.iloc[:, 1] == stop_text].index.min()
    df_filtrat = df.iloc[3:stop_index] if pd.notna(stop_index) else df.iloc[3:]
    df_filtrat = df_filtrat[df_filtrat.iloc[:, 1].notna() & (df_filtrat.iloc[:, 1] != 0) & (df_filtrat.iloc[:, 1] != '-')]

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
    df_filtrat_pt_subtotal1 = df_filtrat[~df_filtrat.iloc[:, 1].isin(valori_de_exclus)]
    st.dataframe(df_filtrat_pt_subtotal1)
    subtotal_1 = df_filtrat_pt_subtotal1.iloc[:, 3].sum()

stop_row = None
subtotal_2 = 0  
elemente_specifice = [
    "Cursuri instruire personal",
    "Toaleta ecologica",
    "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati",
    "Rampa mobila"
]
for index, row in df.iterrows():
    if row[1] == 'Total proiect':
        stop_row = index
        break 
    if row[1] in elemente_specifice:
        subtotal_2 += row[4]  
if stop_row is not None:
    valoare_total_proiect = df.iloc[stop_row, 4]
else:
    pass 

st.write(f"Total: {valoare_total_proiect}")
st.write(f"Subtotal 2: {subtotal_2:.2f}")
st.write(f"Subtotal 1: {subtotal_1}")
