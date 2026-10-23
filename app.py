import json
import random
import streamlit as st

st.set_page_config(page_title="Convocazioni Agape", page_icon="⚽", layout="centered")

# Carica la rosa dal file esterno "giocatori.json"
@st.cache_data
def carica_rosa():
    try:
        with open("giocatori.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        st.error("Errore: Il file 'giocatori.json' non è stato trovato!")
        return []

ROSA = carica_rosa()
MIA_SQUADRA = "AGAPE UMBERTIDE PIERANTONIO 🔴🔵"

def genera_convocazione(rosa):
    avanzati = [g for g in rosa if g["Livello"] == "Avanzato"]
    intermedi = [g for g in rosa if g["Livello"] == "Intermedio"]
    base = [g for g in rosa if g["Livello"] == "Base"]
    
    random.shuffle(avanzati)
    random.shuffle(intermedi)
    random.shuffle(base)
    
    # Squadra Rossa (12): 4-4-4
    squadra_rossa = avanzati[:4] + intermedi[:4] + base[:4]
    # Squadra Blu (11): 4-4-3
    squadra_blu = avanzati[4:] + intermedi[4:] + base[4:]
    
    return squadra_rossa, squadra_blu

def formatta_giocatori(giocatori):
    return "\n".join([f"- {g['Cognome']} {g['Nome']}" for g in giocatori])

st.title("⚽ Convocazioni AGAPE")
st.write("Compila i campi e premi il pulsante per generare i messaggi per WhatsApp.")

data_gara = st.text_input("📅 Data Gara", "SABATO 28 OTTOBRE")
ora_inizio = st.text_input("Rb Ora Inizio Gare", "15:00")

st.markdown("---")
st.subheader("🔴 GRUPPO ROSSO")
avversario_rosso = st.text_input("Avversario Rosso", "Castel del Piano")
in_casa_rosso = st.checkbox("Gioca in Casa (Rosso)", value=True)
stadio_rosso = st.text_input("Stadio / Campo Rosso", "Stadio Comunale Umbertide")
ritrovo_rosso = st.text_input("Ritrovo Rosso", "Campo Sportivo Umbertide")
spogliatoio_rosso = st.text_input("Spogliatoio Rosso", "Spogliatoio 1")
ora_conv_rossa = st.text_input("Ora Convocazione Rossa", "14:00")

st.markdown("---")
st.subheader("🔵 GRUPPO BLU")
avversario_blu = st.text_input("Avversario Blu", "Cannara")
in_casa_blu = st.checkbox("Gioca in Casa (Blu)", value=False)
stadio_blu = st.text_input("Stadio / Campo Blu", "Stadio Spoletini Cannara")
ritrovo_blu = st.text_input("Ritrovo Blu", "Piazzale Parcheggio Umbertide")
spogliatoio_blu = st.text_input("Spogliatoio Blu", "Spogliatoio Ospiti")
ora_conv_blu = st.text_input("Ora Convocazione Blu", "13:30")

st.markdown("---")

if st.button("⚡ GENERATE MATCH DAY E CONVOCAZIONI", type="primary", use_container_width=True):
    if not ROSA:
        st.error("Nessun giocatore caricato.")
    else:
        rossa, blu = genera_convocazione(ROSA)
        
        casa_r = MIA_SQUADRA if in_casa_rosso else avversario_rosso
        trasf_r = avversario_rosso if in_casa_rosso else MIA_SQUADRA
        
        casa_b = MIA_SQUADRA if in_casa_blu else avversario_blu
        trasf_b = avversario_blu if in_casa_blu else MIA_SQUADRA

        matchday_txt = f"""🔴GRUPPO ROSSO🔴
📅Match Day – {data_gara}
⚽ {casa_r} vs {trasf_r}
Rb Inizio gara: {ora_inizio}
📍 Stadio/Campo - {stadio_rosso}


🔵GRUPPO BLU🔵
📅 Match Day – 
{data_gara}
⚽ {casa_b} vs {trasf_b}
Rb Inizio gara: ore {ora_inizio}
📍 Stadio/Campo - {stadio_blu}"""

        convocazione_txt = f"""📌 La convocazione seguirà nella giornata di giovedì:

📣 CONVOCAZIONE GARA
GRUPPO ROSSO🔴
📅 Data: {data_gara}
⏰ Convocazione: ore {ora_conv_rossa}
📍 Ritrovo: {ritrovo_rosso}
🚿 Spogliatoio : {spogliatoio_rosso}

🕑 Inizio gara: ore {ora_inizio}
🏟️ Campo: {stadio_rosso}

⚔️ Gara: {casa_r} vs {trasf_r}

👥 Convocati:
{formatta_giocatori(rossa)}

📲 *RISPONDERE AL SONDAGGIO ENTRO VENERDÌ ALLE ORE 13:00*


📣 CONVOCAZIONE GARA
GRUPPO BLU 🔵
📅 Data: {data_gara}
⏰ Convocazione: {ora_conv_blu}
📍 Ritrovo: {ritrovo_blu}
🚿 Spogliatoio : {spogliatoio_blu}

🕑 Inizio gara: ore {ora_inizio}
🏟️ Campo: {stadio_blu}

⚔️ Gara: {casa_b} vs {trasf_b}

👥 Convocati:
{formatta_giocatori(blu)}

📲 *RISPONDERE AL SONDAGGIO ENTRO VENERDÌ ALLE ORE 13:00*"""

        st.subheader("1. Testo per Match Day")
        st.code(matchday_txt, language=None)
        
        st.subheader("2. Testo Convocazioni (Giovedì)")
        st.code(convocazione_txt, language=None)
