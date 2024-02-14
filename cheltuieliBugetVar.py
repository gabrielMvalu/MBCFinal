import pandas as pd

c_stop_text = 'Total proiect'
c_stop_text2 = 'Total active corporale'
c_stop_text3 = 'Total active necorporale'

def extrage_cheltuieli_buget(df):
    c_stop_row = df[df.iloc[:, 1] == c_stop_text].index.min()
    df_filtrat = df.iloc[3:c_stop_row, [1, 11]] if pd.notna(c_stop_row) else df.iloc[3:, [1, 11]] 
    c_conditii_de_excludere = (df_filtrat.iloc[:, 0] == c_stop_text) | (df_filtrat.iloc[:, 0] == c_stop_text2) | (df_filtrat.iloc[:, 0] == c_stop_text3) | (df_filtrat.iloc[:, 0] == '-') | (df_filtrat.iloc[:, 0] == 0) | df_filtrat.iloc[:, 0].isna()
    df_filtrat = df_filtrat[~c_conditii_de_excludere]
    return df_filtrat
