import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io


def transforma_date(df, start_row, stop_text):
    stop_indexes = df.index[df.iloc[:, 1].eq(stop_text)].tolist()
    if not stop_indexes:
        raise ValueError(f"'{stop_text}' nu a fost găsit în DataFrame.")
    stop_index = stop_indexes[0]
    df = df.iloc[start_row:stop_index]
    return transforma(df)

def transforma_date_alt(df, start_text, stop_text):
    start_indexes = df.index[df.iloc[:, 1].eq(start_text)].tolist()
    if not start_indexes:
        raise ValueError(f"'{start_text}' nu a fost găsit în DataFrame.")
    start_index = start_indexes[0] + 1
    
    stop_indexes = df.index[df.iloc[:, 1].eq(stop_text)].tolist()
    if not stop_indexes:
        raise ValueError(f"'{stop_text}' nu a fost găsit în DataFrame.")
    stop_index = stop_indexes[0]

    df = df.iloc[start_index:stop_index]
    return transforma(df)

def transforma(df):
    df = df[df.iloc[:, 1].notna() & (df.iloc[:, 1] != 0) & (df.iloc[:, 1] != '-')]
    nr_crt, um_list, cantitate_list, pret_unitar_list, valoare_totala_list, linie_bugetara_list, eligibil_neeligibil = [], [], [], [], [], [], []
    counter = 1

    for _, row in df.iterrows():
        nr_crt.append(counter)
        um_list.append("buc")
        cantitate = int(row.iloc[11]) if pd.notna(row.iloc[11]) else None
        cantitate_list.append(cantitate)
        pret_unitar_list.append(row.iloc[3])
        valoare_totala = row.iloc[3] * cantitate if cantitate is not None else None
        valoare_totala_list.append(valoare_totala)
        linie_bugetara_list.append(row.iloc[14])
        counter += 1

        val_6 = pd.to_numeric(row.iloc[6], errors='coerce')
        val_4 = pd.to_numeric(row.iloc[4], errors='coerce')
        eligibil_neeligibil.append(determina_eligibilitate(val_6, val_4))

    return pd.DataFrame({
        "Nr. crt.": nr_crt,
        "Denumirea lucrărilor / bunurilor/ serviciilor": df.iloc[:, 1],
        "UM": um_list,
        "Cantitate": cantitate_list,
        "Preţ unitar (fără TVA)": pret_unitar_list,
        "Valoare Totală (fără TVA)": valoare_totala_list,
        "Linie bugetară": linie_bugetara_list,
        "Eligibil/ neeligibil": eligibil_neeligibil,
    })

def determina_eligibilitate(val_6, val_4):
    if pd.isna(val_6) or pd.isna(val_4):
        return "Data Missing"
    elif val_6 == 0 and val_4 != 0:
        return f"0 // {round(val_4, 2)}"
    elif val_6 == 0 and val_4 == 0:
        return "0 // 0"
    elif val_6 < val_4:
        return f"{round(val_6, 2)} // {round(val_4 - val_6, 2)}"
    else:
        return f"{round(val_6, 2)} // {round(val_6 - val_4, 2)}"

st.title('Transformare Date Excel')

uploaded_file = st.file_uploader("Alegeți fișierul Excel:", type='xlsx')
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    stop_text1 = 'Total active corporale'
    start_text2 = 'Publicitate'
    stop_text2 = 'Total active necorporale'

    df1_transformed = transforma_date(df, 4, stop_text1)
    st.write("Tabel 1:", df1_transformed)

    df2_transformed = transforma_date_alt(df, start_text2, stop_text2)
    st.write("Tabel 2:", df2_transformed)
