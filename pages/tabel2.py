import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from io import BytesIO

stop_text = 'Total proiect'

def transforma_date_tabel2(df):
    stop_index = df[df.iloc[:, 1] == stop_text].index.min()
    df_filtrat = df.iloc[3:stop_index] if pd.notna(stop_index) else df.iloc[3:]
    df_filtrat = df_filtrat[df_filtrat.iloc[:, 1].notna() & (df_filtrat.iloc[:, 1] != 0) & (df_filtrat.iloc[:, 1] != '-')]

    stop_in = df.index[df.iloc[:, 1].eq("Total proiect")].tolist()
    if stop_in:
        val_total_proiect = df.iloc[stop_in[0], 4]  # Asumând că valoarea totală a proiectului se află în coloana 5 (index 4)
    else:
        val_total_proiect = None  # Sau setează la o valoare implicită dacă 'Total proiect' nu este găsit

    valori_de_eliminat = ["Total active corporale", "Total active necorporale", "Publicitate", "Consultanta management", "Consultanta achizitii", "Consultanta scriere"]
    
    df_filtrat = df_filtrat[~df_filtrat.iloc[:, 1].isin(valori_de_eliminat)]

    # Reordonare specifică bazată pe index
    cursuri_index = df_filtrat.index[df_filtrat.iloc[:, 1] == "Cursuri instruire personal"].tolist()
    toaleta_index = df_filtrat.index[df_filtrat.iloc[:, 1] == "Toaleta ecologica"].tolist()
    rampa_iddex = df_filtrat.index[df_filtrat.iloc[:, 1] == "Rampa mobila"].tolist()
    servicii_index = df_filtrat.index[df_filtrat.iloc[:, 1] == "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati"].tolist()
    
    if cursuri_index and toaleta_index:
        toaleta_row = df_filtrat.loc[toaleta_index[0]]
        df_filtrat = df_filtrat.drop(toaleta_index)
        df_filtrat = pd.concat([df_filtrat.iloc[:cursuri_index[0]], toaleta_row.to_frame().T, df_filtrat.iloc[cursuri_index[0]:]], ignore_index=True)

    # Inițializarea variabilelor pentru datele finale
    nr_crt = []
    denumire = []
    um = []
    cantitate = []
    pret_unitar = []
    valoare_totala = []

    # Variabile pentru subtotaluri
    subtotal_1 = 0
    subtotal_2 = 0

    for i, row in df_filtrat.iterrows():
        item = row[1]  # Presupunând că 'Denumire' este în a doua coloană

        if item not in ["Cursuri instruire personal", "Toaleta ecologica", "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati", "Rampa mobila"]:
            subtotal_1 += row[5]  # Suma valorilor pentru 'Valoare Totală'

        if item in  ["Cursuri instruire personal", "Toaleta ecologica", "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati", "Rampa mobila"]:
            subtotal_2 += row[5]

        if item == "Cursuri instruire personal":
            nr_crt.append("Subtotal 1")
            denumire.append("Total valoare cheltuieli cu investiția care contribuie substanțial la obiectivele de mediu")
            um.append(None)
            cantitate.append(None)
            pret_unitar.append(None)
            valoare_totala.append(subtotal_1)

        nr_crt.append(i + 1)  # Ajustează numărul de ordine dacă este necesar
        denumire.append(item)
        um.append(row[2])  # Presupunând că 'UM' este în a treia coloană
        cantitate.append(row[3])  # Presupunând că 'Cantitate' este în a patra coloană
        pret_unitar.append(row[4])  # Presupunând că 'Preţ unitar' este în a cincea coloană
        valoare_totala.append(row[5])  # Presupunând că 'Valoare Totală' este în a șasea coloană

    # Adăugarea celorlalte intrări specifice după procesarea tuturor elementelor
    additional_entries = [
        {"Nr. crt.": "Subtotal 2", "Denumire": "Total valoare cheltuieli cu investiția care contribuie substanțial la egalitatea de șanse, de tratament și accesibilitatea pentru persoanele cu dizabilități", "UM": None, "Cantitate": None, "Preţ unitar (fără TVA)": None, "Valoare Totală (fără TVA)": subtotal_2},
        {"Nr. crt.": None, "Denumire": "Valoare totala eligibila proiect", "UM": None, "Cantitate": None, "Preţ unitar (fără TVA)": None, "Valoare Totală (fără TVA)": val_total_proiect},
        {"Nr. crt.": "Pondere", "Denumire": "Total valoare cheltuieli cu investiția care contribuie substanțial la obiectivele de mediu / Valoare totala eligibila proiect", "UM": None, "Cantitate": None, "Preţ unitar (fără TVA)": None, "Valoare Totală (fără TVA)": 100 * subtotal_1 / val_total_proiect if val_total_proiect else None},
        {"Nr. crt.": "Pondere", "Denumire": "Total valoare cheltuieli cu investiția care contribuie substanțial la egalitatea de șanse, de tratament și accesibilitatea pentru persoanele cu dizabilități / Valoare totala eligibila proiect", "UM": None, "Cantitate": None, "Preţ unitar (fără TVA)": None, "Valoare Totală (fără TVA)": 100 * subtotal_2 / val_total_proiect if val_total_proiect else None}
    ]

    for entry in additional_entries:
        nr_crt.append(entry["Nr. crt."])
        denumire.append(entry["Denumire"])
        um.append(entry["UM"])
        cantitate.append(entry["Cantitate"])
        pret_unitar.append(entry["Preţ unitar (fără TVA)"])
        valoare_totala.append(entry["Valoare Totală (fără TVA)"])

    # Crearea DataFrame-ului final
    tabel_2 = pd.DataFrame({
        "Nr. crt.": nr_crt,
        "Denumire": denumire,
        "UM": um,
        "Cantitate": cantitate,
        "Preţ unitar (fără TVA)": pret_unitar,
        "Valoare Totală (fără TVA)": valoare_totala
    })

    return tabel_2



 

st.title('Transformare Date Excel')

uploaded_file = st.file_uploader("Alegeți fișierul Excel:", type='xlsx')
uploaded_word_file = st.file_uploader("Încarcă documentul Word", type=['docx'])


if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name='P. FINANCIAR')
    tabel_2 = transforma_date_tabel2(df)
    st.dataframe(tabel_2)
    

    placeholder_found = False

    doc = Document(uploaded_word_file) if uploaded_word_file is not None else None
    
    if doc:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if "#tabel3" in cell.text:
                        cell.text = ""  # Șterge placeholder
                        data_frame = tabel_2  # Alege DataFrame-ul corespunzător
                        for i, data_row in data_frame.iterrows():
                            new_row = table.add_row()
                            for j, value in enumerate(data_row):
                                new_row.cells[j].text = str(value)
                        break  # Ieșire din bucla celulelor după popularea tabelului
                break  # Ieșire din bucla rândurilor după găsirea și popularea tabelului
    
        word_modified_bytes = BytesIO()
        doc.save(word_modified_bytes)
        word_modified_bytes.seek(0)
    
        st.download_button(label="Descarcă documentul Word modificat", data=word_modified_bytes, file_name="Document_modificat.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
