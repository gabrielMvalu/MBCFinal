import pandas as pd

def extrage_date_solicitate(df):
    firma = df.iloc[2, 2]
    categ_intreprindere = df.iloc[3, 2]
    firme_legate = df.iloc[4, 2]
    tip_investitie = df.iloc[5, 2]
    activitate = df.iloc[6, 2]
    caen = df.iloc[7, 2]
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
    contributia_proiectului_la_TJ = df.iloc[42, 2]
    strategii_materiale = df.iloc[43, 2]
    strategii_reciclate = df.iloc[44, 2]
    activitate = df.iloc[45, 2]
    utilaj_reciclare_descriere = df.iloc[46, 2]
    lucrari_inovatie = df.iloc[47, 2]
    lucrari_caen = df.iloc[48, 2]
    aDNSH = df.iloc[49, 2]
    cDNSH = df.iloc[50, 2]
    dDNSH = df.iloc[50, 2]
    materiale_locale = df.iloc[52, 2]
    pregatireaTeren = df.iloc[53, 2]
    reciclareaMaterialelor = df.iloc[54, 2]
    clientiFirma = df.iloc[55, 2]
    dacaTipInvest = df.iloc[56, 2]
    crearea = df.iloc[59, 2]
    diversificarea = df.iloc[62, 2]
    extinderea = df.iloc[63, 2]
    crestCreare = df.iloc[64, 2]
    creareActivVizata = df.iloc[65, 2]
    dezavantajeConcurentiale = df.iloc[66, 2]
    locuriMunca30 = df.iloc[67, 2]
    locuriMunca20 = df.iloc[68, 2]
    zoneVizatePrioritar = df.iloc[69, 2]
    iSO14001 = df.iloc[71, 2]
    
    data = {
        "Denumirea firmei SRL": firma, 
        "Categorie întreprindere": categ_intreprindere, 
        "Firme legate": firme_legate,  
        "Tipul investiției": tip_investitie,  
        "Activitate": activitate,
        "Cod CAEN": caen,
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
        "Legături": legaturi,
        "Rude în cadrul firmei": rude,
        "Concluzie_CA": concluzie_CA, 
        "Caracteristici tehnice relevante": caracteristici_tehnice,
        "Flux tehnologic": flux_tehnologic,
        "Utilaje DNSH": utilajeDNSH,
        "Descriere utilaj ghidare": utilaj_ghidare_descriere,
        "Descriere utilaj reciclare": utilaj_reciclare_descriere,
        "Contribuția proiectului la tranziția justă": contributia_proiectului_la_TJ,
        "Strategii materiale": strategii_materiale,
        "Strategii materiale reciclate": strategii_reciclate,
        "Activitate specifică": activitate,
        "Descriere utilaj de reciclare": utilaj_reciclare_descriere,
        "Inovații în lucrări": lucrari_inovatie,
        "Lucrări conform codurilor CAEN": lucrari_caen,
        "Detalii DNSH - A": aDNSH,
        "Detalii DNSH - C": cDNSH,
        "Detalii DNSH - D": dDNSH,
        "Utilizarea materialelor locale": materiale_locale,
        "Pregătirea terenului pentru lucrări": pregatireaTeren,
        "Procesul de reciclare a materialelor": reciclareaMaterialelor,
        "Clienți principali ai firmei": clientiFirma,
        "Tipul investiției planificate": dacaTipInvest,
        "Crearea de noi oportunități": crearea,
        "Diversificarea activităților firmei": diversificarea,
        "Extinderea capacității firmei": extinderea,
        "Creșterea și crearea de noi activități": crestCreare,
        "Crearea de activități în domeniul vizat": creareActivVizata,
        "Identificarea dezavantajelor concurențiale": dezavantajeConcurentiale,
        "Locuri Noi Create 30%": locuriMunca30 ,
        "Locuri Noi Create 20%": locuriMunca20,
        "Zone vizate Prioritar": zoneVizatePrioritar,
        "Certificat ISO 14001 ": iSO14001,                     

    }

    return data
