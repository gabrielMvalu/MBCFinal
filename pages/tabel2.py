import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from io import BytesIO

stop_text = 'Total proiect'

def transforma_date_tabel2(df):
    stop_text = 'Total proiect'
    stop_index = df[df.iloc[:, 1] == stop_text].index.min()
    df_filtrat = df.iloc[3:stop_index] if pd.notna(stop_index) else df.iloc[3:]
    df_filtrat = df_filtrat[df_filtrat.iloc[:, 1].notna() & (df_filtrat.iloc[:, 1] != 0) & (df_filtrat.iloc[:, 1] != '-')]
    valori_de_eliminat = ["Total active corporale", "Total active necorporale", "Publicitate", "Consultanta management", "Consultanta achizitii", "Consultanta scriere"]
    df_filtrat = df_filtrat[~df_filtrat.iloc[:, 1].isin(valori_de_eliminat)]
    valori_speciale = ["Cursuri instruire personal", "Toaleta ecologica", "Rampa mobila", "Servicii de adaptare a utilajelor pentru operarea acestora de persoanele cu dizabilitati"]
    df_filtrat2 = df_filtrat[df_filtrat.iloc[:, 1].isin(valori_speciale)]
    df_filtrat3 = df_filtrat[~df_filtrat.iloc[:, 1].isin(valori_speciale)]
    subtotal_1 = df_filtrat2.iloc[:, 5].sum()
    subtotal_2 = df_filtrat3.iloc[:, 5].sum()
    val_total_proiect = df_filtrat.iloc[:, 5].sum()  # Assuming total project value is the sum of all values in column 5

    df_final = pd.DataFrame({
        "Denumire": df_filtrat.iloc[:, 1],
        "UM": df_filtrat.iloc[:, 2],
        "Cantitate": df_filtrat.iloc[:, 3],
        "Preţ unitar (fără TVA)": df_filtrat.iloc[:, 4],
        "Valoare Totală (fără TVA)": df_filtrat.iloc[:, 5]
    }).reset_index(drop=True)

    df_final["Nr. crt."] = range(1, len(df_final) + 1)

    # Append rows without "Nr. crt." for subtotal and total
    df_final = df_final.append({
        "Denumire": "Subtotal 1",
        "Valoare Totală (fără TVA)": subtotal_1
    }, ignore_index=True)

    df_final = df_final.append(df_filtrat2.drop(columns=["Nr. crt."]), ignore_index=True)

    df_final = df_final.append({
        "Denumire": "Subtotal 2",
        "Valoare Totală (fără TVA)": subtotal_2
    }, ignore_index=True)

    df_final = df_final.append({
        "Denumire": "Valoare totală proiect",
        "Valoare Totală (fără TVA)": val_total_proiect
    }, ignore_index=True)

    # Correct "Nr. crt." for rows where applicable
    df_final["Nr. crt."] = df_final.index + 1
    df_final.loc[df_final['Denumire'].isin(['Subtotal 1', 'Subtotal 2', 'Valoare totală proiect']), 'Nr. crt.'] = ''

    return df_final


 

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
