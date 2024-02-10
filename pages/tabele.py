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


# Funcție pentru setarea bordurilor unei celule
def set_cell_border(cell):
    border_xml = """
    <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tcBorders>
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(border_xml)
    tcPr.append(tcBorders)



def populate_table(doc, start_placeholder, end_placeholder, df):
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            for cell in row.cells:
                if start_placeholder in cell.text:
                    # Am găsit începutul secțiunii unde trebuie să adăugăm date
                    cell.text = cell.text.replace(start_placeholder, "")
                    # Populăm datele în tabel începând cu rândul următor
                    current_row = i + 1
                    for _, data_row in df.iterrows():
                        # Dacă rândul curent este mai mare decât numărul de rânduri din tabel, adăugăm un rând nou
                        if current_row >= len(table.rows):
                            table.add_row()
                        for j, value in enumerate(data_row):
                            table.cell(current_row, j).text = str(value) if pd.notna(value) else ""
                        current_row += 1
                elif end_placeholder in cell.text:
                    # Am găsit sfârșitul secțiunii și ne oprim din adăugat date
                    return
                    
def populate_rows(table, start_index, df):
    for i, data in df.iterrows():
        # Adăugăm un rând nou dacă am depășit numărul de rânduri din tabel
        if i + start_index == len(table.rows):
            table.add_row()
        row = table.rows[i + start_index]
        for j, value in enumerate(data):
            cell = row.cells[j]
            cell.text = str(value) if pd.notna(value) else ""



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

    # Populăm secțiunea de active corporale
    populate_table(doc, "#qq", "#pp", df_nou[df_nou['Denumirea lucrărilor / bunurilor/ serviciilor'].str.contains("Total active corporale", case=False)])
    
    # Populăm secțiunea de active necorporale
    populate_table(doc, "#pp", None, df_nou[df_nou['Denumirea lucrărilor / bunurilor/ serviciilor'].str.contains("Total active necorporale", case=False)])


    # Salvarea documentului modificat într-un buffer
    word_modified_bytes = io.BytesIO()
    doc.save(word_modified_bytes)
    word_modified_bytes.seek(0)

    # Oferirea documentului modificat pentru descărcare
    st.download_button(label="Descarcă documentul Word modificat",
                       data=word_modified_bytes,
                       file_name="Document_modificat.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
