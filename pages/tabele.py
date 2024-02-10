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
    stop_row = df[df.iloc[:, 1] == stop_text].index.min()

    # Verificăm dacă stop_text a fost găsit și limităm DataFrame-ul la primele 15 coloane
    if pd.notna(stop_row):
        df_filtered = df.iloc[4:stop_row + 1, :16]  # Selectăm rândurile și primele 15 coloane
    else:
        st.write('Textul de stop nu a fost găsit. Se afișează toate datele începând cu rândul 5, limitat la 15 coloane.')
        df_filtered = df.iloc[4:, :16]  # Selectăm toate datele începând cu rândul 5, limitat la 15 coloane

    st.write('Datele filtrate (limitate la 16 coloane):')
    st.dataframe(df_filtered)


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

