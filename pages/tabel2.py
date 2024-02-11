import streamlit as st
import pandas as pd

# Titlul aplicației
st.title(':blue_heart: Transformare Date Excel')

# Funcție pentru transformarea datelor din tabel
def transforma_date_tabel2(df, valori_de_eliminat, elemente_specifice):
    stop_text = 'Total proiect'
    
    # Filtrează DataFrame-ul până la 'Total proiect'
    stop_index = df[df.iloc[:, 1] == stop_text].index.min()
    df_filtrat = df.iloc[3:stop_index] if pd.notna(stop_index) else df.iloc[3:]
    df_filtrat = df_filtrat[df_filtrat.iloc[:, 1].notna() & (df_filtrat.iloc[:, 1] != 0) & (df_filtrat.iloc[:, 1] != '-')]
    df_filtrat = df_filtrat[~df_filtrat.iloc[:, 1].isin(valori_de_eliminat)]
    
    # Calculează subtotaluri și valoarea totală a proiectului
    subtotal_1 = df_filtrat.iloc[:, 3].sum()  # Presupunând că subtotalul se calculează din coloana 4
    subtotal_2 = df_filtrat[df_filtrat.iloc[:, 1].isin(elemente_specifice)].iloc[:, 3].sum()  # Suma pentru elemente specifice
    val_total_proiect = subtotal_1 + subtotal_2  # Valoarea totală a proiectului

    # Reorganizează elementele specifice în DataFrame, dacă este necesar
    # Logica pentru reordonarea elementelor specifice poate fi adăugată aici

    # Construiește noul DataFrame pentru afișare
    tabel_2 = pd.DataFrame({
        "Nr. crt.": range(1, len(df_filtrat) + 1),
        "Denumire": df_filtrat.iloc[:, 1],
        "UM": df_filtrat.iloc[:, 2],
        "Cantitate": df_filtrat.iloc[:, 3],
        "Preţ unitar (fără TVA)": df_filtrat.iloc[:, 4],
        "Valoare Totală (fără TVA)": df_filtrat.iloc[:, 5]
    })

    # Adaugă subtotaluri și valoarea totală la sfârșitul DataFrame-ului
    tabel_2 = tabel_2.append([
        {"Denumire": "Subtotal 1", "Valoare Totală (fără TVA)": subtotal_1},
        {"Denumire": "Subtotal 2", "Valoare Totală (fără TVA)": subtotal_2},
        {"Denumire": "Valoare totală proiect", "Valoare Totală (fără TVA)": val_total_proiect}
    ], ignore_index=True)
    
    return tabel_2

# Încărcarea fișierului Excel
uploaded_file = st.file_uploader("Alegeți fișierul Excel:", type='xlsx')

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name='P. FINANCIAR')

    valori_de_eliminat = [
        "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati",
        "Rampa mobila", "Toaleta ecologica", "Total active corporale",
        "Total active necorporale", "Publicitate", "Consultanta management",
        "Consultanta achizitii", "Consultanta scriere", "Cursuri instruire personal",
    ]

    elemente_specifice = [
        "Cursuri instruire personal",
        "Toaleta ecologica",
        "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati",
        "Rampa mobila"
    ]

    tabel_2 = transforma_date_tabel2(df, valori_de_eliminat, elemente_specifice)
    st.dataframe(tabel_2)

