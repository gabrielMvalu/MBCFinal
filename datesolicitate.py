import pandas as pd


def extrage_date_suplimentare(judet, caen, tip_activitate):
    data_foi = {}
    with pd.ExcelFile('./variabile/machetaVariabile.xlsx') as xls:
        df_judet = pd.read_excel(xls, sheet_name=judet)
        df_caen = pd.read_excel(xls, sheet_name=caen)
        df_activitate = pd.read_excel(xls, sheet_name=activitate)

    contributia_proiectului_la_TJ = df_judet.iloc[0, 1]
    strategii_materiale = df_judet.iloc[1, 1]
    strategii_reciclare = df_judet.iloc[2, 1]
    
    activitate = df_caen.iloc[0, 1] 
    descriere_utilaj_reciclare = df_caen.iloc[1, 1]
    lucrari_inovatie = df_caen.iloc[2, 1]
    lucrari_caen = df_caen.iloc[3, 1]
    aDNSH = df_caen.iloc[4, 1]
    cDNSH = df_caen.iloc[5, 1]
    dDNSH = df_caen.iloc[6, 1]
    materiale_locale = df_caen.iloc[7, 1]
    pregatireaTeren = df_caen.iloc[8, 1]
    reciclareaMaterialelor = df_caen.iloc[9, 1]
    clientiFirma = df_caen.iloc[10, 1]
    descriere_serviciu = df_caen.iloc[11, 1]
    piata_tinta = df_caen.iloc[12, 1]

    crestere_creare = df_tip_activitate.iloc[0, 1]
    creareActivVizata = df_tip_activitate.iloc[1, 1]
    dezavantajeConcurentiale = df_tip_activitate.iloc[2, 1]
    
    data_foi = {
        #Variabile din foile ce pot fi nume judet
        "Contribuția proiectului la tranziția justă": contributia_proiectului_la_TJ,
        "Strategii materiale": strategii_materiale,
        "Strategii materiale reciclate": strategii_reciclare,

        #Variabilele din foile ce pot fi nr cod caen
        "Activitate specifică": activitate,
        "Inovații în lucrări": lucrari_inovatie,
        "Lucrări conform codurilor CAEN": lucrari_caen,
        "Detalii DNSH - A": aDNSH,
        "Detalii DNSH - C": cDNSH,
        "Detalii DNSH - D": dDNSH,
        "Utilizarea materialelor locale": materiale_locale,
        "Pregătirea terenului pentru lucrări": pregatireaTeren,
        "Procesul de reciclare a materialelor": reciclareaMaterialelor,
        "Clienți principali ai firmei": clientiFirma,
        "Descriere serviciului": descriere_serviciu,
        "Piata tinta": piata_tinta,

        #Variabile din foi ce pot fi veche sau nou
        "Creșterea sau crearea de noi surse de venit": crestere_creare,
        "Crearea de activități în domeniul vizat": creareActivVizata,
        "Identificarea dezavantajelor concurențiale": dezavantajeConcurentiale, 
    }

    return data_foi


def extrage_date_solicitate(df):
    firma = df.iloc[2, 2]
    categ_intreprindere = df.iloc[3, 2]
    firme_legate = df.iloc[4, 2]
    tip_investitie = df.iloc[5, 2]
    activitate = df.iloc[6, 2]
    caen = df.iloc[7, 2]
    nr_caen = df.iloc[43, 2]
    nr_locuri_munca_noi = df.iloc[8, 2]
    judet = df.iloc[9, 2]
    utilaj_dizabilitati = df.iloc[10, 2]
    utilaj_cu_tocator = df.iloc[11, 2]
    adresa_loc_implementare = df.iloc[12, 2]
    nr_clasare_notificare = df.iloc[13, 2]
    clienti_actuali = df.iloc[14, 2]
    furnizori = df.iloc[15, 2]
    tip_activitate = df.iloc[16, 2]
    iso = df.iloc[17, 2]
    activitate_curenta = df.iloc[18, 2]
    dotari_activitate_curenta = df.iloc[19, 2]
    info_ctr_implementare = df.iloc[20, 2]
    zonele_vizate_prioritar = df.iloc[21, 2]
    utilaj_ghidare = df.iloc[22, 2]
    legaturi = df.iloc[23, 2]
    rude = df.iloc[24, 2]
    concluzie_CA = df.iloc[36, 2]
    caracteristici_tehnice = df.iloc[37, 2]
    flux_tehnologic = df.iloc[38, 2]
    utilajeDNSH = df.iloc[39, 2]
    utilaj_ghidare_descriere = df.iloc[40, 2]
    dacaTipInvest = df.iloc[41, 2]
    nrlocmunca30 = df.iloc[49, 2]
    nrlocmunca20 = df.iloc[50, 2]
    iso14001 = df.iloc[53, 2]

    
    data = {
        #Variabile extrase din date solicitate completate de consultanti
        "Denumirea firmei SRL": firma, 
        "Categorie întreprindere": categ_intreprindere, 
        "Firme legate": firme_legate,  
        "Tipul investiției": tip_investitie,  
        "Activitate": activitate,
        "Cod CAEN": caen,
        "Doar nr CAEN": nr_caen,
        "Număr locuri de muncă noi": nr_locuri_munca_noi,
        "Județ": judet,
        "Utilaj pentru persoane cu dizabilități": utilaj_dizabilitati,
        "Utilaj cu tocător": utilaj_cu_tocator,
        "Adresa locației de implementare": adresa_loc_implementare,
        "Număr clasare notificare": nr_clasare_notificare,
        "Clienți actuali": clienti_actuali,
        "Furnizori": furnizori,
        "Tip activitate": tip_activitate,
        "Certificări ISO": iso,
        "Activitate curentă": activitate_curenta,
        "Dotări pentru activitatea curentă": dotari_activitate_curenta,
        "Informații despre contractul de implementare": info_ctr_implementare,
        "Zonele vizate prioritare": zonele_vizate_prioritar,
        "Utilaj de ghidare": utilaj_ghidare,
        "Legaturi": legaturi,
        "Rude": rude,
        "Concluzie cifra de afaceri": concluzie_CA,
        "Caracteristici tehnice utilaje": caracteristici_tehnice,
        "Fluxul tehnologic": flux_tehnologic,
        "DNSH pentru utilaje": utilajeDNSH,
        "Descrierea utilaj ghidare": utilaj_ghidare_descriere,
        "Tipul investitiei": dacaTipInvest,
        "Procent 30% din total locuri munca nou create": nrlocmunca30,
        "Procent 30% din total locuri munca nou create": nrlocmunca20,
        "Daca are sau nu iso14001": iso14001,
    }
    
    # Apelarea funcției pentru a extrage datele suplimentare
    date_suplimentare = extrage_date_suplimentare(judet, nr_caen, tip_activitate)
    # Actualizarea dicționarului `data` cu datele suplimentare extrase
    data.update(date_suplimentare)

    return data
