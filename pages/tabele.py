import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io


def transforma_date(df):
    stop_index = df.index[df.iloc[:, 1].eq(stop_text)].tolist()
    df = df.iloc[3:stop_index[0]] if stop_index else df.iloc[3:]
    df = df[df.iloc[:, 1].notna() & (df.iloc[:, 1] != 0) & (df.iloc[:, 1] != '-')]

    nr_crt, um_list, cantitate_list, pret_unitar_list, valoare_totala_list, linie_bugetara_list, eligibil_neeligibil = [], [], [], [], [], [], []
    counter = 1

    for index, row in df.iterrows():
        item = row.iloc[1].strip().lower()
        if item in ["total active corporale", "total active necorporale"]:
            nr_crt.append(None)
            um_list.append(None)
            cantitate_list.append(None)
            pret_unitar_list.append(None)
            valoare_totala_list.append(None)
            linie_bugetara_list.append(None)
        else:
            nr_crt.append(counter)
            um_list.append("buc")
            cantitate = int(row.iloc[11]) if pd.notna(row.iloc[11]) else None
            cantitate_list.append(cantitate)
            pret_unitar_list.append(row.iloc[3])
            valoare_totala = row.iloc[3] * cantitate if cantitate is not None else None
            valoare_totala_list.append(valoare_totala)
            linie_bugetara_list.append(row.iloc[14])
            counter += 1

    for index, row in df.iterrows():
        val_6 = pd.to_numeric(row.iloc[6], errors='coerce')
        val_4 = pd.to_numeric(row.iloc[4], errors='coerce')
        if pd.isna(val_6) or pd.isna(val_4):
            eligibil_neeligibil.append("Data Missing")
        elif val_6 == 0 and val_4 != 0:
            eligibil_neeligibil.append(f"0 // {round(val_4, 2)}")
        elif val_6 == 0 and val_4 == 0:
            eligibil_neeligibil.append("0 // 0")
        elif val_6 < val_4:
            eligibil_neeligibil.append(f"{round(val_6, 2)} // {round(val_4 - val_6, 2)}")
        else:
            eligibil_neeligibil.append(f"{round(val_6, 2)} // {round(val_6 - val_4, 2)}")

      
    
    df_nou = pd.DataFrame({
        "Nr. crt.": [str(int(x)) if pd.notna(x) else "" for x in nr_crt],
        "Denumirea lucrărilor / bunurilor/ serviciilor": df.iloc[:, 1],
        "UM": um_list,
        "Cantitate": [str(int(x)) if pd.notna(x) else "" for x in cantitate_list],
        "Preţ unitar (fără TVA)": pret_unitar_list,
        "Valoare Totală (fără TVA)": valoare_totala_list,
        "Linie bugetară": linie_bugetara_list,
        "Eligibil/ neeligibil": eligibil_neeligibil,
        "Contribuie la criteriile de evaluare a,b,c,d": df.iloc[:, 15]
    })

    
    return df_nou 



def populate_table_from_placeholder(doc, placeholder, df):
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                if placeholder in cell.text:
                    # Începem să populăm datele în tabel de la acest rând și celulă
                    cell.text = ""  # Înlăturăm placeholder-ul
                    current_row_index = i
                    current_cell_index = j
                    # Iterăm prin DataFrame și adăugăm datele în tabel
                    for df_index, data_row in df.iterrows():
                        # Asigură-te că avem suficiente rânduri în tabel
                        if current_row_index == len(table.rows):
                            table.add_row()
                        # Obținem rândul curent în tabel
                        row = table.rows[current_row_index]
                        # Populăm datele începând de la celula curentă
                        for df_cell_index in range(len(data_row)):
                            table_index = current_cell_index + df_cell_index
                            # Asigură-te că nu depășim numărul de celule din rând
                            if table_index < len(row.cells):
                                row.cells[table_index].text = str(data_row[df_cell_index]) if pd.notna(data_row[df_cell_index]) else ""
                        current_row_index += 1
                    # După ce am terminat de populat secțiunea, ieșim din buclă
                    return



# Titlul aplicației
st.title('Încărcare și prelucrare fișier Excel și Word')

# Widget pentru încărcarea fișierului Excel
uploaded_excel_file = st.file_uploader("Alege un fișier Excel (.xlsx)", type=['xlsx'], key="excel")

# Widget pentru încărcarea fișierului Word
uploaded_word_file = st.file_uploader("Încarcă documentul Word", type=['docx'], key="word")

# Variabila care reprezintă textul de stop
stop_text = 'Total proiect'

# Procesarea fișierului Excel
if uploaded_excel_file is not None:
    df = pd.read_excel(uploaded_excel_file, sheet_name='P. FINANCIAR')
    df_nou = transforma_date(df)

# Procesarea și modificarea fișierului Word
if uploaded_word_file is not None and df_nou is not None:
    # Încărcarea documentului Word din buffer
    word_bytes = io.BytesIO(uploaded_word_file.getvalue())
    doc = Document(word_bytes)

    # Populăm tabelul începând de la placeholder-ul #qq
    populate_table_from_placeholder(doc, "#qq", df_nou)

    # Populăm tabelul începând de la placeholder-ul #pp
    populate_table_from_placeholder(doc, "#pp", df_nou)


    # Salvarea documentului modificat într-un buffer
    word_modified_bytes = io.BytesIO()
    doc.save(word_modified_bytes)
    word_modified_bytes.seek(0)

    # Oferirea documentului modificat pentru descărcare
    st.download_button(label="Descarcă documentul Word modificat",
                       data=word_modified_bytes,
                       file_name="Document_modificat.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
