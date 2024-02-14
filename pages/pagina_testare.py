import streamlit as st
import pandas as pd

stop_text = 'Total proiect'
stop_text2 = 'Total active corporale'
stop_text3 = 'Total active necorporale'

def extrage_cheltuieli_buget(df):
    stop_row = df[df.iloc[:, 1] == stop_text].index.min()
    df_filtrat = df.iloc[3:stop_row, [1, 11]] if pd.notna(stop_row) else df.iloc[3:, [1, 11]] 
    conditii_de_excludere = (df_filtrat.iloc[:, 0] == stop_text) | (df_filtrat.iloc[:, 0] == stop_text2) | (df_filtrat.iloc[:, 0] == stop_text3) | (df_filtrat.iloc[:, 0] == '-') | (df_filtrat.iloc[:, 0] == 0) | df_filtrat.iloc[:, 0].isna()
    df_filtrat = df_filtrat[~conditii_de_excludere]
    return df_filtrat

uploaded_file = st.file_uploader("Alegeți fișierul Excel:", type='xlsx')

if uploaded_file is not None:
    try:
        df_financiar = pd.read_excel(uploaded_file, sheet_name='P. FINANCIAR')
        df_rezultate = extrage_cheltuieli_buget(df_financiar)

        if not df_rezultate.empty:
            # Afișarea dataframe-ului rezultat
            st.write(df_rezultate)

            # Construirea stringului cheltuieli_buget
            cheltuieli_buget = '\n'.join([f"{row[0]} - {row[1]} buc. \n" for _, row in df_rezultate.iterrows()])

            # Afișarea cheltuielilor buget sub forma specificată
            st.write(cheltuieli_buget)
        else:
            st.error("Nu s-au găsit date valide în foaia 'P. FINANCIAR'.")

    except ValueError as e:
        st.error(f'Eroare: {e}')
else:
    st.error("Vă rugăm să încărcați un fișier.")