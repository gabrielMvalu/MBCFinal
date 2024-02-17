#bilantsianaliza.py
import pandas as pd

def extrage_date_bilant(df):
    cpa20 = f"{df.iloc[76, 1]:.2f}"
    cpa21 = f"{df.iloc[76, 2]:.2f}"
    cpa22 = f"{df.iloc[76, 3]:.2f}"
    
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
    }
    return data

def extrage_indicatori_financiari(df2):
    rs20 = f"{df2.iloc[89, 1]:.2f}"  
    rs21 = f"{df2.iloc[89, 2]:.2f}"
    rs22 = f"{df2.iloc[89, 3]:.2f}"
    gdi20 = f"{df2.iloc[95, 1]:.0%}"
    gdi21 = f"{df2.iloc[95, 2]:.0%}"
    gdi22 = f"{df2.iloc[95, 3]:.0%}"
    roa20 = f"{df2.iloc[43, 1]:.0%}"
    roa21 = f"{df2.iloc[43, 2]:.0%}"
    roa22 = f"{df2.iloc[43, 3]:.0%}"
    roe20 = f"{df2.iloc[47, 1]:.0%}"
    roe21 = f"{df2.iloc[47, 2]:.0%}"
    roe22 = f"{df2.iloc[47, 3]:.2%}"
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
    
   # df = pd.read_excel(uploaded_file, sheet_name='1-Bilant')
   # df1 = pd.read_excel(uploaded_file, sheet_name='2-ContPP')
   # df2 = pd.read_excel(uploaded_file, sheet_name='1D-Analiza_fin_indicatori')    
