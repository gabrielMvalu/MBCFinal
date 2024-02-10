import streamlit as st
import pandas as pd
from docx import Document

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



def replace_placeholder_with_table(doc_path, placeholder, df):
    # Deschiderea documentului Word existent
    doc = Document(doc_path)

    # Căutarea fiecărui paragraf pentru placeholder
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            # Crearea unui tabel în locul placeholder-ului
            table = doc.add_table(rows=df.shape[0]+1, cols=df.shape[1])
            table.style = 'Table Grid'  # Adăugarea unui stil de tabel (opțional)

            # Completarea anteturilor de coloane
            for j in range(df.shape[1]):
                table.cell(0, j).text = df.columns[j]

            # Completarea celulelor tabelului cu date
            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    table.cell(i+1, j).text = str(df.iloc[i, j])

            # Ștergerea paragrafului cu placeholder
            p = paragraph._element
            p.getparent().remove(p)
            p._p = p._element = None

            break  # Oprește bucla după ce găsește și înlocuiește primul placeholder

    # Salvarea documentului modificat
    doc.save('document_modificat.docx')

# Presupunând că ai un file_uploader pentru încărcarea documentului Word
uploaded_word_file = st.file_uploader("Încarcă documentul Word care conține placeholder-ul", type=['docx'])

if uploaded_word_file is not None:
    # Salvarea temporară a fișierului încărcat pentru a-l putea deschide cu Document
    with open("temp_document.docx", "wb") as f:
        f.write(uploaded_word_file.getbuffer())
    
    # Apelarea funcției pentru a înlocui placeholder-ul cu tabelul
    replace_placeholder_with_table("temp_document.docx", "#Tabel1", df_filtered)
    st.success("Placeholder-ul a fost înlocuit cu tabelul în documentul Word.")



