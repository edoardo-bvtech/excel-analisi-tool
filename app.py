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

# 2️⃣ PRENDI TUTTE LE RIGHE DATI SOTTO L'INTESTAZIONE
                dati_df = df.iloc[header_row + 1:].reset_index(drop=True)
                if dati_df.empty:
                    st.warning(f"⚠️ {file.name} ignorato: Riga dati mancante sotto l'intestazione.")
                    continue

                # 3️⃣ ESTRAZIONE INTESTAZIONE E INDICI COLONNE
                intestazione = [str(x).strip() for x in df.iloc[header_row].tolist()]
                campi_obbligatori = [
                    "Punt. Reale", "Punt. B1d", "Circolati", "In Fascia",
                    "%OS (soglia propria)", "T Rit", "T Eff", "Fuori Fascia Per Esterne"
                ]

                missing = [c for c in campi_obbligatori if c not in intestazione]
                if missing:
                    st.warning(f"⚠️ {file.name} ignorato: Mancano le colonne {missing} nell'intestazione.")
                    continue

                col_idx = {c: intestazione.index(c) for c in campi_obbligatori}

                # 4️⃣ ESTRAZIONE METADATI (Data Inizio, Data Fine, Cliente)
                data_inizio = estrai_metadato(df, "Data Inizio")
                data_fine = estrai_metadato(df, "Data Fine")
                cliente = estrai_metadato(df, "Cliente")

                # 5️⃣ SALVATAGGIO RISULTATI RIGA PER RIGA
                for _, row in dati_df.iterrows():
                    # salta righe completamente vuote
                    if row.isna().all() or (pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == ""):
                        continue

                    def _round(val):
                        if pd.isna(val) or val == "":
                            return None
                        try:
                            return round(float(val), decimali)
                        except Exception:
                            return val

                    risultati.append({
                        "File Origine": file.name,
                        "Cliente": cliente,
                        "Data Inizio": data_inizio,
                        "Data Fine": data_fine,
                        "Punt. Reale": _round(row[col_idx["Punt. Reale"]]),
                        "Punt. B1d": _round(row[col_idx["Punt. B1d"]]),
                        "Circolati": _round(row[col_idx["Circolati"]]),
                        "In Fascia": _round(row[col_idx["In Fascia"]]),
                        "%OS (soglia propria)": _round(row[col_idx["%OS (soglia propria)"]]),
                        "T Rit": _round(row[col_idx["T Rit"]]),
                        "T Eff": _round(row[col_idx["T Eff"]]),
                        "Fuori Fascia Per Esterne": _round(row[col_idx["Fuori Fascia Per Esterne"]])
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
        st.info("Nessun dato valido estratto dai file caricati. Assicurati che i file contengano l'intestazione corretta e che la colonna A nella riga dati sia vuota.")
