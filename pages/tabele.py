import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import io

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

# Funcție pentru adăugarea unui tabel cu borduri într-un document Word
def add_df_with_borders_to_doc(doc, df):
    table = doc.add_table(rows=1, cols=len(df.columns))  # Creăm tabelul
    
    # Setăm bordurile pentru header
    for cell in table.rows[0].cells:
        set_cell_border(cell)
    
    # Adăugăm anteturile coloanelor
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(df.columns):
        hdr_cells[i].text = str(column)
    
    # Adăugăm rândurile din DataFrame
    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            row_cells[i].text = str(value) if pd.notna(value) else ""
            set_cell_border(row_cells[i])  # Setăm bordurile pentru fiecare celulă din tabel

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
    
    #functia pt rearanjarea tabelului
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
            "Nr. crt.": nr_crt,
            "Denumirea lucrărilor / bunurilor/ serviciilor": df.iloc[:, 1],
            "UM": um_list,
            "Cantitate": cantitate_list,
            "Preţ unitar (fără TVA)": pret_unitar_list,
            "Valoare Totală (fără TVA)": valoare_totala_list,
            "Linie bugetară": linie_bugetara_list,
            "Eligibil/ neeligibil": eligibil_neeligibil,
            "Contribuie la criteriile de evaluare a,b,c,d": df.iloc[:, 15]
        })
    
        return df_nou

    st.write('Datele filtrate (limitate la 16 coloane):')
    st.dataframe(df_nou)


# Procesarea și modificarea fișierului Word
if uploaded_word_file is not None and df_filtered is not None:
    # Încărcarea documentului Word din buffer
    word_bytes = io.BytesIO(uploaded_word_file.getvalue())
    doc = Document(word_bytes)

    # Căutarea placeholder-ului și înlocuirea cu tabelul
    for paragraph in doc.paragraphs:
        if '#Tabel1' in paragraph.text:
            add_df_with_borders_to_doc(doc, df_filtered)
            p = paragraph._element
            p.getparent().remove(p)
            p._p = p._element = None
            break

    # Salvarea documentului modificat într-un buffer
    word_modified_bytes = io.BytesIO()
    doc.save(word_modified_bytes)
    word_modified_bytes.seek(0)

    # Oferirea documentului modificat pentru descărcare
    st.download_button(label="Descarcă documentul Word modificat",
                       data=word_modified_bytes,
                       file_name="Document_modificat.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
