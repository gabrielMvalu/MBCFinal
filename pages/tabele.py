import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io


def transforma_date(df, start_row, stop_text):
    stop_indexes = df.index[df.iloc[:, 1].str.contains(stop_text, na=False)].tolist()
    if not stop_indexes:
        raise ValueError(f"'{stop_text}' nu a fost găsit în DataFrame.")
    stop_index = stop_indexes[0]
    df = df.iloc[start_row:stop_index]
    return transforma(df)

def transforma(df):
    df = df[df.iloc[:, 1].notna() & (df.iloc[:, 1] != '0') & (df.iloc[:, 1] != '-')]
    nr_crt, um_list, cantitate_list, pret_unitar_list, valoare_totala_list, linie_bugetara_list, eligibil_neeligibil, = [], [], [], [], [], [], [],
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
        return f"Eligibil: 0 \nNeeligibil: {round(val_4, 2)}"
    elif val_6 == 0 and val_4 == 0:
        return "Eligibil: 0 \nNeeligibil: 0"
    elif val_6 < val_4:
        return f"Eligibil: {round(val_6, 2)} \nNeeligibil: {round(val_4 - val_6, 2)}"
    else:
        return f"Eligibil: {round(val_6, 2)} \nNeeligibil: {round(val_6 - val_4, 2)}"



st.title('Transformare Date Excel')

uploaded_file = st.file_uploader("Alegeți fișierul Excel:", type='xlsx')
uploaded_word_file = st.file_uploader("Încarcă documentul Word", type=['docx'])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name='P. FINANCIAR')
    # Definirea textelor de oprire
    stop_text1 = 'Total active corporale'
    stop_text2 = 'Total active necorporale'
    # Transformarea datelor
    df1_transformed = transforma_date(df, 3, stop_text1)
    st.write("Tabel 1:", df1_transformed)

if uploaded_word_file is not None and df1_transformed is not None:
    word_bytes = io.BytesIO(uploaded_word_file.getvalue())
    doc = Document(word_bytes)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "#tabel1" in cell.text:
                    # Ștergerea textului placeholder și începerea populării tabelului
                    cell_index = row.cells.index(cell)
                    for i, data_row in df1_transformed.iterrows():
                        # Adăugarea unui rând nou sau utilizarea rândului curent dacă este primul rând de date
                        if i > 0 or cell_index > 0:  # Dacă nu este prima celulă sau primul rând de date
                            new_row = table.add_row()
                        else:
                            new_row = row
                        for j in range(len(data_row)):
                            new_row.cells[j].text = str(df1_transformed.iloc[i, j])
                    break  # Ieșire din bucla celulelor după popularea tabelului
            if "#tabel1" in cell.text:
                break  # Ieșire din bucla rândurilor după găsirea și popularea tabelului

    # Salvarea și descărcarea documentului Word modificat
    word_modified_bytes = io.BytesIO()
    doc.save(word_modified_bytes)
    word_modified_bytes.seek(0)
    st.download_button(label="Descarcă documentul Word modificat", data=word_modified_bytes, file_name="Document_modificat.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
