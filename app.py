import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Analisi Dati Excel", layout="wide")

st.title("📊 Analisi Excel - Estrazione Punt. B1d & %OS")

# --- 1. SELETTORE DECIMALI ---
st.markdown("### Impostazioni")
decimali = st.number_input("Scegli il numero di decimali per i dati estratti:", min_value=0, max_value=6, value=2)

st.markdown("---")

# Intestazione esatta da cercare
HEADER_ATTESO = [
    "Viste", "Giorno", "Settimana", "Mese", "Anno",
    "Media Treni Circolati", "Punt. Reale", "Punt. B1d",
    "Punt. SB", "Punt. RFI", "Punt. IF",
    "Circolati", "In Fascia", "Fuori Fascia",
    "Fuori Fascia Per Esterne", "Fuori Fascia per RFI",
    "Fuori Fascia per Propria IF", "Fuori Fascia per Altre IF",
    "%OS (soglia propria)", "T Rit", "T Eff"
]

# --- FUNZIONE ESTRAZIONE COLONNA A ---
def estrai_metadato(df, chiave):
    chiave_lower = chiave.lower()
    for i in range(len(df)):
        valore = df.iloc[i, 0]
        if pd.notna(valore):
            testo = str(valore).strip()
            testo_lower = testo.lower()
            
            if testo_lower.startswith(chiave_lower):
                # Se c'è il simbolo "=" prende tutto quello che c'è dopo
                if "=" in testo:
                    return testo.split("=", 1)[1].strip()
                # Altrimenti prende il testo successivo alla parola chiave
                else:
                    return testo[len(chiave):].strip()
    return None

# --- 2. CARICAMENTO FILE ---
uploaded_files = st.file_uploader(
    "Carica uno o più file Excel (.xlsx)",
    accept_multiple_files=True,
    type=["xlsx"]
)

if uploaded_files:
    risultati = []
    
    with st.spinner("Elaborazione dei file in corso..."):
        for file in uploaded_files:
            try:
                # Leggiamo senza intestazione fissa per poter scansionare tutto il foglio
                df = pd.read_excel(file, header=None)
                header_row = None

                # 1️⃣ TROVA RIGA INTESTAZIONE COMPLETA
                for i in range(len(df)):
                    row = df.iloc[i].tolist()
                    row_clean = [str(x).strip() for x in row]
                    
                    if row_clean[:len(HEADER_ATTESO)] == HEADER_ATTESO:
                        header_row = i
                        break
                
                if header_row is None:
                    st.warning(f"⚠️ {file.name} ignorato: Intestazione corretta non trovata.")
                    continue

                # 2️⃣ PRENDI RIGA SUCCESSIVA E CONTROLLA COLONNA A
                dati_row = header_row + 1
                
                if dati_row >= len(df):
                    st.warning(f"⚠️ {file.name} ignorato: Riga dati mancante sotto l'intestazione.")
                    continue
                
                cella_A = df.iloc[dati_row, 0]
                if pd.notna(cella_A) and str(cella_A).strip() != "":
                    st.warning(f"⚠️ {file.name} ignorato: La colonna A nella riga dati non è vuota.")
                    continue

                # 3️⃣ ESTRAZIONE PUNT. B1d E %OS
                intestazione = [str(x).strip() for x in df.iloc[header_row].tolist()]
                
                punt_col = intestazione.index("Punt. B1d")
                os_col = intestazione.index("%OS (soglia propria)")
                
                punt_value = df.iloc[dati_row, punt_col]
                os_value = df.iloc[dati_row, os_col]

                # Arrotondamento sicuro
                try:
                    punt_value = round(float(punt_value), decimali) if pd.notna(punt_value) else None
                except ValueError:
                    pass # Lascia il valore originale se non è un numero
                
                try:
                    os_value = round(float(os_value), decimali) if pd.notna(os_value) else None
                except ValueError:
                    pass

                # 4️⃣ ESTRAZIONE METADATI (Data Inizio, Data Fine, Cliente)
                data_inizio = estrai_metadato(df, "Data Inizio")
                data_fine = estrai_metadato(df, "Data Fine")
                cliente = estrai_metadato(df, "Cliente")

                # 5️⃣ SALVATAGGIO RISULTATO
                risultati.append({
                    "File Origine": file.name,
                    "Cliente": cliente,
                    "Data Inizio": data_inizio,
                    "Data Fine": data_fine,
                    "Punt. B1d": punt_value,
                    "%OS (soglia propria)": os_value
                })

            except Exception as e:
                st.error(f"❌ Errore durante l'elaborazione di {file.name}: {e}")

    # --- 3. MOSTRA E SCARICA I RISULTATI ---
    if risultati:
        st.success(f"✅ Elaborazione completata per {len(risultati)} file validi!")
        
        df_finale = pd.DataFrame(risultati)
        st.dataframe(df_finale, use_container_width=True)

        # Creazione del file Excel in memoria
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_finale.to_excel(writer, index=False, sheet_name='Riepilogo Dati')
        
        st.download_button(
            label="📥 Scarica Riepilogo in Excel (.xlsx)",
            data=output.getvalue(),
            file_name="riepilogo_dati_estratti.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.info("Nessun dato valido estratto dai file caricati. Verifica i messaggi di avviso qui sopra.")
