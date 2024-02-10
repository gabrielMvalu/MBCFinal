import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Titlul aplicației
st.title('Încărcare și prelucrare fișier Excel')

# Widget pentru încărcarea fișierului
uploaded_file = st.file_uploader("Alege un fișier Excel (.xlsx)", type=['xlsx'])

# Variabila care reprezintă textul de stop
stop_text = 'Total proiect'

# Verifică dacă a fost încărcat un fișier
if uploaded_file is not None:
    # Citirea datelor din fișierul Excel încărcat
    df = pd.read_excel(uploaded_file, sheet_name='P. FINANCIAR')

    # Găsirea rândului care conține stop_text în coloana B (folosind indexul 1 pentru a accesa coloana B)
    stop_row = df[df.iloc[:, 1] == stop_text].index.min()

    # Verificăm dacă stop_text a fost găsit
    if pd.notna(stop_row):
        # Selectarea datelor de la rândul 5 până la rândul stop_index
        df_filtered = df.iloc[4:stop_row + 1]  # Indexarea începe de la 0, de aceea folosim 4 pentru rândul 5
    else:
        st.write('Textul de stop nu a fost găsit. Se afișează toate datele începând cu rândul 5.')
        df_filtered = df.iloc[4:]  # Selectăm toate datele începând cu rândul 5

    # Afișarea datelor filtrate în Streamlit
    st.write('Datele filtrate:')
    st.dataframe(df_filtered)
else:
    st.write('Așteptând încărcarea unui fișier Excel...')



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

# Încărcăm documentul Word
doc = Document("x.docx")

# Căutăm placeholder-ul și înlocuim cu tabelul
for paragraph in doc.paragraphs:
    if '#TABEL' in paragraph.text:
        # Adăugăm tabelul cu borduri în document
        add_df_with_borders_to_doc(doc, df_filtered)  # Presupunem că df_filtered este DataFrame-ul tău
        
        # Ștergem paragraful cu placeholder
        p = paragraph._element
        p.getparent().remove(p)
        p._p = p._element = None  # Înlăturăm referințele pentru a evita scurgerile de memorie
        break  # Opriți bucla după primul placeholder găsit și înlocuit

# Salvăm documentul modificat
doc.save("xcomplet.docx")
