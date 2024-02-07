#serviciisiutilaje.py
import streamlit as st
import pandas as pd
import re

def extrage_pozitii(df_financiar):
    values_col2_and_col12 = []
    for index, row in df_financiar.iterrows():
        value_col2 = row.iloc[1]
        value_col12 = row.iloc[11]
        if value_col2 == "Total proiect":
            break
        if value_col2 and value_col12 and value_col12 != 0:
            values_col2_and_col12.append((value_col2, value_col12))
    return values_col2_and_col12

def coreleaza_date(date_financiar):
    df_amortizare = pd.read_excel('./assets/utilaje.xlsx', sheet_name='amortizare')
    df_utilaje = pd.read_excel('./assets/utilaje1.xlsx', sheet_name='utilajeservicii')
    df_descriere = pd.read_excel(f'./assets/{st.session_state.codCAEN}.xlsx', sheet_name='utilajedescriere')
    
    amortizare_data = {}
    servicii_data = {}
    descriere_data = {}

    for index, row in df_amortizare.iterrows():
        if row.iloc[1]:
            nume = ' '.join(re.sub(r'\d+$', '', str(row.iloc[1])).strip().split()).lower()
            cod = ' '.join(str(row.iloc[2]).strip().split()) if row.iloc[2] else ''
            descriere = ' '.join(str(row.iloc[3]).strip().split()) if row.iloc[3] else ''
            amortizare_data[nume] = (cod, descriere)

    for index, row in df_utilaje.iterrows():
        if row.iloc[1]:
            nume = ' '.join(re.sub(r'\d+$', '', str(row.iloc[1])).strip().split()).lower()
            servicii_data[nume] = nume

    for index, row in df_descriere.iterrows():
        if row.iloc[1]:
            nume = ' '.join(re.sub(r'\d+$', '', str(row.iloc[1])).strip().split()).lower()
            descriere = ' '.join(str(row.iloc[2]).strip().split()) if row.iloc[2] else ''
            descriere_data[nume] = descriere

    rezultate_corelate = []
    rezultate_corelate1 = []
    rezultate_corelate2 = []

    for nume, cantitate in date_financiar:
        nume_curat = ' '.join(re.sub(r'\d+$', '', str(nume)).strip().split()).lower()
        if nume_curat in amortizare_data:
            cod, descriere = amortizare_data[nume_curat]
            rezultat = f"{nume}, {cantitate} buc., ce aparține clasei {cod} {descriere}, conform HG 2139/2004"
            rezultate_corelate.append((nume, cantitate, rezultat))
        if nume_curat in servicii_data:
            rezultat1 = f"{nume}, {cantitate} buc"
            rezultate_corelate1.append((nume, cantitate, rezultat1))
        if nume_curat in descriere_data:
            descriere = descriere_data[nume_curat]
            rezultat2 = f"{descriere}"
            rezultate_corelate2.append((nume, cantitate, rezultat2))

    return rezultate_corelate, rezultate_corelate1, rezultate_corelate2
