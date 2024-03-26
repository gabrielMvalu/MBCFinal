#bilantsianalizanou.py
import pandas as pd
import streamlit as st

def extrage_date_bilant(df):
    cpa20 = f"{df.iloc[99, 1]:.2f}"
    cpa21 = f"{df.iloc[99, 2]:.2f}"
    cpa22 = f"{df.iloc[99, 3]:.2f}"
    
    data = {
        "Capitalul propriu al actionarilor 2020": cpa20, 
        "Capitalul propriu al actionarilor 2021": cpa21,
        "Capitalul propriu al actionarilor 2022": cpa22
    }
    return data

def extrage_date_contpp(df1):
    ca20 = f"{df1.iloc[4, 1]:.2f}"
    ca21 = f"{df1.iloc[4, 2]:.2f}"
    ca22 = f"{df1.iloc[4, 3]:.2f}"
    vt20 = f"{df1.iloc[55, 1]:.2f}"
    vt21 = f"{df1.iloc[55, 2]:.2f}"
    vt22 = f"{df1.iloc[55, 3]:.2f}"

    pfexpl = f"{df1.iloc[33, 3]:.2f}"

    if df1.iloc[4, 6] != 0:
        pc = f"{(df1.iloc[4, 6] / df1.iloc[4, 3] * 100) - 100:.2f}%"
    else:
        pc = "N/A" 
    
    if df1.iloc[4, 1] > df1.iloc[4, 2] and df1.iloc[4, 1] > df1.iloc[4, 3]:
        camax = 2020
    elif df1.iloc[4, 2] > df1.iloc[4, 1] and df1.iloc[4, 2] > df1.iloc[4, 3]:
        camax = 2021
    else:
        camax = 2022
    
    re20 = f"{df1.iloc[66, 1]:.2f}" if df1.iloc[66, 1] > 0 else f"{df1.iloc[67, 1]:.2f}"
    re21 = f"{df1.iloc[66, 2]:.2f}" if df1.iloc[66, 2] > 0 else f"{df1.iloc[67, 2]:.2f}"
    re22 = f"{df1.iloc[66, 3]:.2f}" if df1.iloc[66, 3] > 0 else f"{df1.iloc[67, 3]:.2f}"
    
    data = {
        "Cifra de afaceri 2020": ca20, 
        "Cifra de afaceri 2021": ca21,
        "Cifra de afaceri 2022": ca22,
        "Venituri totale 2020": vt20, 
        "Venituri totale 2021": vt21,
        "Venituri totale 2022": vt22,
        "Rezultat al exercitiului 2020": re20, 
        "Rezultat al exercitiului 2021": re21,
        "Rezultat al exercitiului 2022": re22,  
        "Anul cu cea mai mare cifra de afaceri": camax, 
        "Procent crestere": pc,
        
        "Profit exploatare": pfexpl,
    }
    return data

def extrage_indicatori_financiari(df2):
    rs20 = f"{df2.iloc[89, 1]:.2f}"  
    rs21 = f"{df2.iloc[89, 2]:.2f}"
    rs22 = f"{df2.iloc[89, 3]:.2f}"
    
    gdi20 = f"{df2.iloc[94, 1]:.0%}"
    gdi21 = f"{df2.iloc[94, 2]:.0%}"
    gdi22 = f"{df2.iloc[94, 3]:.0%}"
    
    roa20 = f"{df2.iloc[43, 1]:.0%}" 
    roa21 = f"{df2.iloc[43, 2]:.0%}" 
    roa22 = f"{df2.iloc[43, 3]:.0%}"

    
    roe20 = f"{df2.iloc[47, 1]:.0%}"
    roe21 = f"{df2.iloc[47, 2]:.0%}"
    roe22 = f"{df2.iloc[47, 3]:.0%}"


    
    #adaugare dupa schimbari din 16.feb 
    rpe22 = f"{df2.iloc[32, 3]:.2%}"

    data = {
        "Rata solvabilitatii generale 2020": rs20, 
        "Rata solvabilitatii generale 2021": rs21,
        "Rata solvabilitatii generale 2022": rs22,
        "Gradul de indatorare pe termen scurt 2020": gdi20, 
        "Gradul de indatorare pe termen scurt 2021": gdi21,
        "Gradul de indatorare pe termen scurt 2022": gdi22,
        "Rentabilitatea activelor (ROA) 2020": roa20, 
        "Rentabilitatea activelor (ROA) 2021": roa21,
        "Rentabilitatea activelor (ROA) 2022": roa22,
        "Rentabilitatea capitalului propriu (ROE) 2020": roe20, 
        "Rentabilitatea capitalului propriu (ROE) 2021": roe21,
        "Rentabilitatea capitalului propriu (ROE) 2022": roe22,
        #adaugare  schimbari din 16.feb 
        "Rata profitului din exploatare RPE22": rpe22
    }

    return data


    """
    Extrage valoarea bazată pe un termen de căutare dintr-o anumită coloană.
    
    :param df: DataFrame-ul din care se extrag datele.
    :param coloana_cautare: Indexul coloanei în care se caută termenul.
    :param termen_cautare: Termenul care se caută în coloană.
    :param coloana_valoare: Indexul coloanei de unde se extrage valoarea.
    :param mesaj_eroare: Mesajul afișat dacă termenul de căutare nu este găsit.
    :return: Valoarea extrasă sau None dacă termenul de căutare nu este găsit.
    """
def extrage_valoare_din_df(df, coloana_cautare, termen_cautare, coloana_valoare, mesaj_eroare):
    index = df.index[df.iloc[:, coloana_cautare].eq(termen_cautare)].tolist()
    
    if index:
        valoare = df.iloc[index[0], coloana_valoare]
        return valoare
    else:
        st.error(mesaj_eroare)
        return None
