import streamlit as st
import pandas as pd

# Titlul aplicației
st.header(':blue[Prelucrare Tabelara]', divider='rainbow')

# Widget pentru încărcarea fișierului
uploaded_file = st.file_uploader("Incarca Anexa 3 Macheta Financiara (.xlsx)", type=['xlsx'])

# Verifică dacă a fost încărcat un fișier
if uploaded_file is not None:
  
    # Citirea datelor din fișierul Excel încărcat
    df = pd.read_excel(uploaded_file, sheet_name='P. FINANCIAR')

    # Afișarea datelor în Streamlit
    st.write('Datele încărcate:')
    st.dataframe(df)
  
else:
    st.write('Așteptând încărcarea unui fișier Excel...')
