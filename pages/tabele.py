import streamlit as st
import pandas as pd

# Titlul aplicației
st.title('Încărcare și prelucrare fișier Excel')

# Widget pentru încărcarea fișierului
uploaded_file = st.file_uploader("Alege un fișier Excel (.xlsx)", type=['xlsx'])

# Variabila care reprezintă textul de stop
stop_text = 'Total_proiect'

# Verifică dacă a fost încărcat un fișier
if uploaded_file is not None:
    # Citirea datelor din fișierul Excel încărcat
    df = pd.read_excel(uploaded_file, sheet_name='P. FINANCIAR')

    # Găsirea rândului care conține stop_text în coloana B (folosind indexul 1 pentru a accesa coloana B)
    stop_index = df[df.iloc[:, 2] == stop_text].index.min()

    # Verificăm dacă stop_text a fost găsit
    if pd.notna(stop_index):
        # Selectarea datelor de la rândul 5 până la rândul stop_index
        df_filtered = df.iloc[4:stop_index]  # Indexarea începe de la 0, de aceea folosim 4 pentru rândul 5
    else:
        st.write('Textul de stop nu a fost găsit. Se afișează toate datele începând cu rândul 5.')
        df_filtered = df.iloc[4:]  # Selectăm toate datele începând cu rândul 5

    # Afișarea datelor filtrate în Streamlit
    st.write('Datele filtrate:')
    st.dataframe(df_filtered)
else:
    st.write('Așteptând încărcarea unui fișier Excel...')

