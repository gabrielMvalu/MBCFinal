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

            valori_de_exclus1 = [
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
            
            valori_de_exclus2 = [
                    "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati",
                    "Rampa mobila",
                    "Toaleta ecologica",
                    "Cursuri instruire personal",
                ]    
                        

            # Filtrăm DataFrame-ul pentru a exclude rândurile cu valorile specificate în lista 'valori_de_exclus'
            df_filtrat_pt_subtotal1 = df_filtrat[~df_filtrat.iloc[:, 1].isin(valori_de_exclus1)]
            df_filtrat_pt_subtotal2 = df_filtrat[df_filtrat.iloc[:, 1].isin(valori_de_exclus2)]    

            subtotal_1 = 0
            subtotal_2 = 0
                        
            for i, row in enumerate(df_filtrat_pt_subtotal1.itertuples(), 1):
                subtotal_1 += row[4]
            
            for i, row in enumerate(df_filtrat_pt_subtotal2.itertuples(), 1):
                subtotal_2 += row[4]


            st.write(f"{subtotal_1}")
            st.write(f"{subtotal_2: .2f}")   
         

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
            
            # Adăugarea "Toaleta ecologica" după "Cursuri instruire personal"
            if cursuri_index and toaleta_index:
                toaleta_row = df_filtrat.loc[toaleta_index[0]]
                df_filtrat = df_filtrat.drop(toaleta_index)
                df_filtrat = pd.concat([df_filtrat.iloc[:cursuri_index[0]], toaleta_row.to_frame().T, df_filtrat.iloc[cursuri_index[0]:]], ignore_index=True)
            
           

            # Initialize 'Nr. crt.' counter and lists for all columns
            nr_crt_counter = 1
            nr_crt = []
            denumire = []
            um = []
            cantitate = []
            pret_unitar = []
            valoare_totala = []
        
        
            # Bucla de procesare a elementelor
            for i, row in enumerate(df_filtrat.itertuples(), 1):
                item = row[2]  # Assuming 'Denumire' is the second column
        
                if item == "Cursuri instruire personal":
                    nr_crt.append("Subtotal 1")
                    denumire.append("Total valoare cheltuieli cu investiția care contribuie substanțial la obiectivele de mediu")
                    um.append(" ")
                    cantitate.append(" ")
                    pret_unitar.append(" ")
                    valoare_totala.append(subtotal_1)
        
                nr_crt.append(nr_crt_counter)
                denumire.append(item)
                um.append("buc")
                cantitate.append(df_filtrat.iloc[i-1, 11])  # Adjust the index as necessary
                pret_unitar.append(df_filtrat.iloc[i-1, 3])
                valoare_totala.append(df_filtrat.iloc[i-1, 3] * df_filtrat.iloc[i-1, 11])
                nr_crt_counter += 1
        
            nr_crt.extend([nr_crt_counter, nr_crt_counter + 1, "Subtotal 2", " ", "Pondere", "Pondere"])
            denumire.extend([
                "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati",
                "Rampa mobila",
                "Total valoare cheltuieli cu investiția care contribuie substanțial la egalitatea de șanse, de tratament și accesibilitatea pentru persoanele cu dizabilități",
                "Valoare totala eligibila proiect",
                "Total valoare cheltuieli cu investiția care contribuie substanțial la obiectivele de mediu / Valoare totala eligibila proiect",
                "Total valoare cheltuieli cu investiția care contribuie substanțial la egalitatea de șanse, de tratament și accesibilitatea pentru persoanele cu dizabilități / Valoare totala eligibila proiect"
            ])
            um.extend(["buc", "buc", " ", " "," ", " "])
            cantitate.extend([" ", " ", " ", " "," ", " "])
            pret_unitar.extend([" ", " ", " ", " "," ", " "])
            valoare_totala.extend([" val pt serv", " val pt rampa", f"{subtotal_2: .2f}", val_total_proiect, f"{100*subtotal_1/val_total_proiect: .2f}%", f"{100*subtotal_2/val_total_proiect: .2f}%"])
        
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


 

st.title('Transformare Date Excel')

uploaded_file = st.file_uploader("Alegeți fișierul Excel:", type='xlsx')
uploaded_word_file = st.file_uploader("Încarcă documentul Word", type=['docx'])


if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name='P. FINANCIAR')
    tabel_2 = transforma_date_tabel2(df)
    st.dataframe(tabel_2)

    placeholder_found = False

    doc = Document(uploaded_word_file) if uploaded_word_file is not None else None
    
    if doc:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if "#tabel3" in cell.text:
                        cell.text = ""  # Șterge placeholder
                        data_frame = tabel_2  # Alege DataFrame-ul corespunzător
                        for i, data_row in data_frame.iterrows():
                            new_row = table.add_row()
                            for j, value in enumerate(data_row):
                                new_row.cells[j].text = str(value)
                        break  # Ieșire din bucla celulelor după popularea tabelului
                break  # Ieșire din bucla rândurilor după găsirea și popularea tabelului
    
        word_modified_bytes = BytesIO()
        doc.save(word_modified_bytes)
        word_modified_bytes.seek(0)
    
        st.download_button(label="Descarcă documentul Word modificat", data=word_modified_bytes, file_name="Document_modificat.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

