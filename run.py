# -*- coding: utf-8 -*-
"""
Created on Tue Jan 21 07:37:49 2025

@author: olmel6349
"""
import matplotlib.pyplot as plt
import pandas as pd

#Oppgave a
def oppgave_a():
    #leser excel filen ved bruk av pandas
    data = pd.read_excel('support_uke_24.xlsx')

    #panda series til lister
    u_dag = data['Ukedag'].values #liste med ukedag
    kl_slett = data['Klokkeslett'].values #liste med klokkeslett
    varighet = data['Varighet'].values #liste med varighet
    score = data['Tilfredshet'].values #liste med score
    return u_dag, kl_slett, varighet, score

#Oppgave b
def oppgave_b(u_dag):
    #liste fra mandag til fredag (1-5), som holder på antall samtaler
    antall_per_uke_dag = [0, 0, 0, 0, 0] 

    #looper gjennom antall ukedager og legger til verdier i listen basert på hvilken dag samtalen fant sted
    for x in u_dag:
        if(x == "Mandag"):
            antall_per_uke_dag[0] += 1
        elif(x == "Tirsdag"):
            antall_per_uke_dag[1] += 1
        elif(x == "Onsdag"):
            antall_per_uke_dag[2] += 1
        elif(x == "Torsdag"):
            antall_per_uke_dag[3] += 1
        elif(x == "Fredag"):
            antall_per_uke_dag[4] += 1

    #bruker matplotlib.pyplot til å vise stolpediagram
    plt.bar(["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag"], antall_per_uke_dag)

#Oppgave c
def oppgave_c(varighet):
    print()
    print("################### SAMTALELENGDE ###################")

    """
    Samtalelengde er lagret som string i xlsx. Konverter til float så man kan sorterte 
    med bruk av sort funksjonen på lista og finne minste tid med [0] og lengeste med [-1]
    """

    #lager en liste som holder på samtaletid i desimalform
    liste_samtale_tid_som_desimaler = []

    #looper gjennom varighet listen og populerer liste_samtale_tid_som_desimaler med float format
    for v in varighet:
        arr = v.split(":")
        desimal_string = f"{arr[1]}.{arr[2]}"
        liste_samtale_tid_som_desimaler.append(float(desimal_string))

    liste_samtale_tid_som_desimaler.sort() #sorterer listen fra lav til høy

    minste_samtale = liste_samtale_tid_som_desimaler[0] #henter ut korteste (minste) samtale

    minste_samtale_liste = str(minste_samtale).split(".") #Lager en liste av korteste (minste) samtale, så det er enklere å skrive ut på en pen måte

    print(f"Minste samtale var på {minste_samtale_liste[0]} minutter og {minste_samtale_liste[1]} sekunder")

    lengste_samtale = liste_samtale_tid_som_desimaler[-1] #henter ut lengste samtale

    lengste_samtale_liste = str(lengste_samtale).split(".") #lager en liste av lengste samtale, så det er enklere å skrive ut på en pen måte

    print(f"Lengste samtale var på {lengste_samtale_liste[0]} minutter og {lengste_samtale_liste[1]} sekunder")

    return liste_samtale_tid_som_desimaler

#Oppgave d
def oppgave_d(liste_samtale_tid_som_desimaler):
    print()
    print("################### Gjennomsnitt ###################")

    #variabel som holder total varighet på samtalene
    total_tidsbruk = 0

    #lopper gjennom liste_samtale_tid_som_desimaler og legger tid til total_tidsbruk
    for tid in liste_samtale_tid_som_desimaler:
        total_tidsbruk += tid

    #finner gjennomsnittet ved å ta total_tidsbruk og dele på antall samtaler i liste_samtale_tid_som_desimaler
    gjennomsnitt = total_tidsbruk / len(liste_samtale_tid_som_desimaler)

    gjennomsnitt_formatert = "{:.2f}".format(gjennomsnitt) #formaterer for finere ukskrift

    gjennomsnitt_liste = gjennomsnitt_formatert.split(".") #splitter for å dele opp i minutter og sekunder

    print(f"Gjennomsnitt tidsbruk på samtalene er {gjennomsnitt_liste[0]} minutter og {gjennomsnitt_liste[1]} sekunder")

#Oppgave e
def oppgave_e(kl_slett):
    #liste som holder på klokkeslett som desimaler
    kL_slett_desimaler = []

    #gjør om klokkeslett til float så de er enkelere å jobbe med
    for k in kl_slett:
        arr = k.split(":")
        desimal_string = f"{arr[0]}.{arr[1]}"
        kL_slett_desimaler.append(float(desimal_string))

    #lager fire variabler som holder på intervalene
    aatte_til_ti = ti_til_tolv = tolv_til_to = to_til_fire = 0

    #looper igjennom kL_slett_desimaler listen og tilordner variablene som holder på intervalene
    for i in kL_slett_desimaler:
        if(i >= 8.0 and i <= 09.59):
            aatte_til_ti += 1
        elif(i >= 10.00 and i <= 11.59):
            ti_til_tolv += 1
        elif(i >= 12.00 and i <= 13.59):
            tolv_til_to += 1
        elif(i >= 14.00 and i <= 16.00):
            to_til_fire +=1

    #lager merkelapper til kakediagrammet
    kl_slett_labels = [f"8-10 ({aatte_til_ti})", f"10-12 ({ti_til_tolv})", f"12-14 ({tolv_til_to})", f"14-16 ({to_til_fire})"]

    #Bruker matplotlib.pyplot til å vise kakediagram
    plt.pie([aatte_til_ti, ti_til_tolv, tolv_til_to, to_til_fire], labels=kl_slett_labels)

#Oppgave f
def oppgave_f(score):
    #variabler som holder på antallet av de forskjellige meningene
    positive_kunder = negative_kunder = passive_kunder = antall_kunder = 0

    for s in score: #looper igjennom tilfredshet listen
        if((s != s) == False): #fjerner de som er NaN
            antall_kunder += 1 #teller opp antall kunder som har svart på undersøkelsen
            if(s >= 9):
                positive_kunder += 1
            elif(s >= 7 and s <= 8):
                passive_kunder += 1
            elif(s <= 6):
                negative_kunder += 1

    positive_kunder_prosentandel = positive_kunder / antall_kunder #prosentandel positive kunder
    negative_kunder_prosentandel = negative_kunder / antall_kunder #prosentandel negative kunder
    passive_kunder_prosentandel = passive_kunder / antall_kunder #prosentandel passive kunder

    #regner ut nps basert på formelen i oppgaven
    nps = (positive_kunder_prosentandel - negative_kunder_prosentandel) * 100

    #formaterer kundene sine meninger, så det blir fin utskrift
    nps_formatert = "{:.2f}".format(nps)
    negative_kunder_formatert = "{:.2f}".format(negative_kunder_prosentandel * 100)
    passive_kunder_formatert = "{:.2f}".format(passive_kunder_prosentandel * 100)
    positive_kunder_formaatert = "{:.2f}".format(positive_kunder_prosentandel * 100)

    print()
    print("################### NET PROMOTER SCORE ###################")
    print(f"NPS: {nps_formatert}%")
    print(f"Detractors: {negative_kunder_formatert}%")
    print(f"Passives: {passive_kunder_formatert}%")
    print(f"Promoters: {positive_kunder_formaatert}%")
    print(f"Total: {antall_kunder}")

#Start funksjonen kjører alle funksjonene fra oppgave_a til oppgave_f
def start():
    u_dag, kl_slett, varighet, score = oppgave_a()
    oppgave_b(u_dag)
    liste_samtale_tid_som_desimaler = oppgave_c(varighet)
    oppgave_d(liste_samtale_tid_som_desimaler)
    oppgave_e(kl_slett)
    oppgave_f(score)

#Kjører start funksjonen  
start()