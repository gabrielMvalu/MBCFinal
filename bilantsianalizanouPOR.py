#bilantsianalizanou.py
import pandas as pd
import streamlit as st

def extrage_date_bilant(df):
    cpa21 = f"{df.iloc[97, 2]:.2f}"
    cpa22 = f"{df.iloc[97, 3]:.2f}"
    cpa23 = f"{df.iloc[97, 4]:.2f}"
    
    data = {
        "Capitalul propriu al actionarilor 2020": cpa21, 
        "Capitalul propriu al actionarilor 2021": cpa22,
        "Capitalul propriu al actionarilor 2022": cpa23
    }
    return data

def extrage_date_contpp(df1):
    ca21 = f"{df1.iloc[4, 2]:.2f}"
    ca22 = f"{df1.iloc[4, 3]:.2f}"
    ca23 = f"{df1.iloc[4, 4]:.2f}"
    vt21 = f"{df1.iloc[55, 2]:.2f}"
    vt22 = f"{df1.iloc[55, 3]:.2f}"
    vt23 = f"{df1.iloc[55, 4]:.2f}"

    pfexpl = f"{df1.iloc[33, 4]:.2f}"

    if df1.iloc[4, 6] != 0:
        pc = f"{(df1.iloc[4, 6] / df1.iloc[4, 4] * 100) - 100:.2f}%"
    else:
        pc = "N/A" 
    
    if df1.iloc[4, 2] > df1.iloc[4, 3] and df1.iloc[4, 2] > df1.iloc[4, 4]:
        camax = 2021
    elif df1.iloc[4, 3] > df1.iloc[4, 2] and df1.iloc[4, 3] > df1.iloc[4, 2]:
        camax = 2022
    else:
        camax = 2023
    
    re21 = f"{df1.iloc[31, 2]:.2f}" if df1.iloc[31, 2] > 0 else f"{df1.iloc[32, 2]:.2f}"
    re22 = f"{df1.iloc[31, 3]:.2f}" if df1.iloc[31, 3] > 0 else f"{df1.iloc[32, 3]:.2f}"
    re23 = f"{df1.iloc[31, 4]:.2f}" if df1.iloc[31, 4] > 0 else f"{df1.iloc[32, 4]:.2f}"
    
    data = {
        "Cifra de afaceri 2020": ca21, 
        "Cifra de afaceri 2021": ca22,
        "Cifra de afaceri 2022": ca23,
        "Venituri totale 2020": vt21, 
        "Venituri totale 2021": vt22,
        "Venituri totale 2022": vt23,
        "Rezultat al exercitiului 2020": re21, 
        "Rezultat al exercitiului 2021": re22,
        "Rezultat al exercitiului 2022": re23,  
        "Anul cu cea mai mare cifra de afaceri": camax, 
        "Procent crestere": pc,
        
        "Profit exploatare": pfexpl,
    }
    return data

def extrage_indicatori_financiari(df2):
    rs21 = f"{df2.iloc[89, 2]:.2f}"  
    rs22 = f"{df2.iloc[89, 3]:.2f}"
    rs23 = f"{df2.iloc[89, 4]:.2f}"
    
    gdi21 = f"{df2.iloc[94, 2]:.0%}"
    gdi22 = f"{df2.iloc[94, 3]:.0%}"
    gdi23 = f"{df2.iloc[94, 4]:.0%}"

    roa21 = f"{df2.iloc[58, 2]:.0%}" if isinstance(df2.iloc[58, 2], (int, float)) else "nu se calculeaza"
    roa22 = f"{df2.iloc[58, 3]:.0%}" if isinstance(df2.iloc[58, 3], (int, float)) else "nu se calculeaza"
    roa23 = f"{df2.iloc[58, 4]:.0%}" if isinstance(df2.iloc[58, 4], (int, float)) else "nu se calculeaza"

    roe21 = f"{df2.iloc[62, 2]:.0%}" if isinstance(df2.iloc[62, 2], (int, float)) else "nu se calculeaza"
    roe22 = f"{df2.iloc[62, 3]:.0%}" if isinstance(df2.iloc[62, 3], (int, float)) else "nu se calculeaza"
    roe23 = f"{df2.iloc[62, 4]:.0%}" if isinstance(df2.iloc[62, 4], (int, float)) else "nu se calculeaza"

    
    #adaugare dupa schimbari din 16.feb 
    rpe23 = f"{df2.iloc[32, 4]:.2%}"

    data = {
        "Rata solvabilitatii generale 2020": rs21, 
        "Rata solvabilitatii generale 2021": rs22,
        "Rata solvabilitatii generale 2022": rs23,
        "Gradul de indatorare pe termen scurt 2020": gdi21, 
        "Gradul de indatorare pe termen scurt 2021": gdi22,
        "Gradul de indatorare pe termen scurt 2022": gdi23,
        "Rentabilitatea activelor (ROA) 2020": roa21, 
        "Rentabilitatea activelor (ROA) 2021": roa22,
        "Rentabilitatea activelor (ROA) 2022": roa23,
        "Rentabilitatea capitalului propriu (ROE) 2020": roe21, 
        "Rentabilitatea capitalului propriu (ROE) 2021": roe22,
        "Rentabilitatea capitalului propriu (ROE) 2022": roe23,
        #adaugare  schimbari din 16.feb 
        "Rata profitului din exploatare RPE22": rpe23
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
