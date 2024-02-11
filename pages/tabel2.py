import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from io import BytesIO



stop_text = 'Total proiect'

def transforma_date_tabel2(df):
            stop_index = df[df.iloc[:, 1] == stop_text].index.min()
            df_filtrat = df.iloc[3:stop_index] if pd.notna(stop_index) else df.iloc[3:]
            df_filtrat = df_filtrat[df_filtrat.iloc[:, 1].notna() & (df_filtrat.iloc[:, 1] != 0) & (df_filtrat.iloc[:, 1] != '-')]
            
            stop_in = df.index[df.iloc[:, 1].eq("Total proiect")].tolist()

            if stop_in:
                val_total_proiect = df.iloc[stop_in[0], 4]
            else:
                val_total_proiect = None  
                
            valori_de_eliminat = [
                "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati",
                "Rampa mobila", "Total active corporale", "Total active necorporale", 
                "Publicitate", "Consultanta management", "Consultanta achizitii", "Consultanta scriere"
            ]
            df_filtrat = df_filtrat[~df_filtrat.iloc[:, 1].isin(valori_de_eliminat)]


             # Identificarea indexurilor pentru fiecare element specific
            cursuri_index = df_filtrat.index[df_filtrat.iloc[:, 1] == "Cursuri instruire personal"].tolist()
            toaleta_index = df_filtrat.index[df_filtrat.iloc[:, 1] == "Toaleta ecologica"].tolist()
            rampa_index = df_filtrat.index[df_filtrat.iloc[:, 1] == "Rampa mobila"].tolist()
            servicii_index = df_filtrat.index[df_filtrat.iloc[:, 1] == "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati"].tolist()
            
            # Adăugarea "Toaleta ecologica" după "Cursuri instruire personal"
            if cursuri_index and toaleta_index:
                toaleta_row = df_filtrat.loc[toaleta_index[0]]
                df_filtrat = df_filtrat.drop(toaleta_index)
                df_filtrat = pd.concat([df_filtrat.iloc[:cursuri_index[0]], toaleta_row.to_frame().T, df_filtrat.iloc[cursuri_index[0]:]], ignore_index=True)
            
            # Adăugarea "Rampa mobila" după "Toaleta ecologica" sau "Cursuri instruire personal" dacă "Toaleta ecologica" nu este prezentă
            if cursuri_index and rampa_index:
                rampa_row = df_filtrat.loc[rampa_index[0]]
                df_filtrat = df_filtrat.drop(rampa_index)
                df_filtrat = pd.concat([df_filtrat.iloc[:cursuri_index[0]+1], rampa_row.to_frame().T, df_filtrat.iloc[cursuri_index[0]+1:]], ignore_index=True)
            
            # Adăugarea "Servicii de adaptare a utilajelor..." după "Rampa mobila" sau ultimul element adăugat anterior
            if cursuri_index and servicii_index:
                servicii_row = df_filtrat.loc[servicii_index[0]]
                df_filtrat = df_filtrat.drop(servicii_index)
                df_filtrat = pd.concat([df_filtrat.iloc[:cursuri_index[0]+2], servicii_row.to_frame().T, df_filtrat.iloc[cursuri_index[0]+2:]], ignore_index=True)


            nr_crt_counter = 1
            nr_crt = []
            denumire = []
            um = []
            cantitate = []
            pret_unitar = []
            valoare_totala = []
        
                # Add "Subtotal 1" before "Cursuri instruire personal"
                if item == "Cursuri instruire personal":
                    nr_crt.append("Subtotal 1")
                    denumire.append("Total valoare cheltuieli cu investiția care contribuie substanțial la obiectivele de mediu")
                    um.append(None)
                    cantitate.append(None)
                    pret_unitar.append(None)
                    valoare_totala.append(subtotal_1)
        
                # Add items to lists
                nr_crt.append(nr_crt_counter)
                denumire.append(item)
                um.append("buc")
                cantitate.append(df_filtrat.iloc[i-1, 11])  # Adjust the index as necessary
                pret_unitar.append(df_filtrat.iloc[i-1, 3])
                valoare_totala.append(df_filtrat.iloc[i-1, 3] * df_filtrat.iloc[i-1, 11])
                nr_crt_counter += 1
        
            # Add other specific entries after processing all items
            nr_crt.extend(["Subtotal 2", None, "Pondere", "Pondere"])
            denumire.extend([
                "Total valoare cheltuieli cu investiția care contribuie substanțial la egalitatea de șanse, de tratament și accesibilitatea pentru persoanele cu dizabilități",
                "Valoare totala eligibila proiect",
                "Total valoare cheltuieli cu investiția care contribuie substanțial la obiectivele de mediu / Valoare totala eligibila proiect",
                "Total valoare cheltuieli cu investiția care contribuie substanțial la egalitatea de șanse, de tratament și accesibilitatea pentru persoanele cu dizabilități / Valoare totala eligibila proiect"
            ])
            um.extend([None, None, None, None])
            cantitate.extend([None, None, None, None])
            pret_unitar.extend([None, None, None, None])
            valoare_totala.extend([subtotal_2, val_total_proiect, 100*subtotal_1/val_total_proiect, 100*subtotal_2/val_total_proiect])
        
            # Create the final DataFrame
            tabel_2 = pd.DataFrame({
                "Nr. crt.": nr_crt,
                "Denumire": denumire,
                "UM": um,
                "Cantitate": cantitate,
                "Preţ unitar (fără TVA)": pret_unitar,
                "Valoare Totală (fără TVA)": valoare_totala
            })
        
            return tabel_2
    


st.title(':blue[Transformare Date Excel]')


uploaded_file = st.file_uploader("Alegeți fișierul Excel:", type='xlsx')
uploaded_word_file = st.file_uploader("Încarcă documentul Word", type=['docx'])

if uploaded_file is not None:
    stop_text = 'Total proiect'
    df = pd.read_excel(uploaded_file, sheet_name='P. FINANCIAR')
    stop_index = df[df.iloc[:, 1] == stop_text].index.min()
    df_filtrat = df.iloc[3:stop_index] if pd.notna(stop_index) else df.iloc[3:]
    df_filtrat = df_filtrat[df_filtrat.iloc[:, 1].notna() & (df_filtrat.iloc[:, 1] != 0) & (df_filtrat.iloc[:, 1] != '-')]

    valori_de_exclus = [
        "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati",
        "Rampa mobila",
        "Toaleta ecologica",
        "Total active corporale",
        "Total active necorporale",
        "Publicitate",
        "Consultanta management",
        "Consultanta achizitii",
        "Consultanta scriere",
        "Cursuri instruire personal",
    ]
    df_filtrat_pt_subtotal1 = df_filtrat[~df_filtrat.iloc[:, 1].isin(valori_de_exclus)]
    st.dataframe(df_filtrat_pt_subtotal1)
    subtotal_1 = df_filtrat_pt_subtotal1.iloc[:, 3].sum()

stop_row = None
subtotal_2 = 0  
elemente_specifice = [
    "Cursuri instruire personal",
    "Toaleta ecologica",
    "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati",
    "Rampa mobila"
]
for index, row in df.iterrows():
    if row[1] == 'Total proiect':
        stop_row = index
        break 
    if row[1] in elemente_specifice:
        subtotal_2 += row[4]  
if stop_row is not None:
    valoare_total_proiect = df.iloc[stop_row, 4]
else:
    pass 

st.write(f"Total: {valoare_total_proiect}")
st.write(f"Subtotal 2: {subtotal_2:.2f}")
st.write(f"Subtotal 1: {subtotal_1}")
