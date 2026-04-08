import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Analisi Dati Excel", layout="wide")

st.title("📊 Analisi Excel - Estrazione Punt. B1d & %OS")

# --- 1. SELETTORE DECIMALI ---
st.markdown("### Impostazioni")
decimali = st.number_input(
    "Scegli il numero di decimali per i dati estratti:",
    min_value=0, max_value=6, value=2
)

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
                if "=" in testo:
                    return testo.split("=", 1)[1].strip()
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
    file_validi = 0

    with st.spinner("Elaborazione dei file in corso..."):
        for file in uploaded_files:
            try:
                df = pd.read_excel(file, header=None)
                header_row = None

                # 1️⃣ TROVA INTESTAZIONE
                for i in range(len(df)):
                    row = df.iloc[i].tolist()
                    row_clean = [str(x).strip() for x in row]

                    if row_clean[:len(HEADER_ATTESO)] == HEADER_ATTESO:
                        header_row = i
                        break

                if header_row is None:
                    st.warning(f"⚠️ {file.name} ignorato: Intestazione non trovata.")
                    continue

                # 2️⃣ DATI
                dati_df = df.iloc[header_row + 1:].reset_index(drop=True)
                if dati_df.empty:
                    st.warning(f"⚠️ {file.name} ignorato: Nessun dato.")
                    continue

                intestazione = [str(x).strip() for x in df.iloc[header_row].tolist()]

                campi = [
                    "Punt. Reale", "Punt. B1d", "Circolati", "In Fascia",
                    "%OS (soglia propria)", "T Rit", "T Eff", "Fuori Fascia Per Esterne"
                ]

                missing = [c for c in campi if c not in intestazione]
                if missing:
                    st.warning(f"⚠️ {file.name} ignorato: Mancano colonne {missing}")
                    continue

                col_idx = {c: intestazione.index(c) for c in campi}

                # 3️⃣ METADATI
                data_inizio = estrai_metadato(df, "Data Inizio")
                data_fine = estrai_metadato(df, "Data Fine")
                cliente = estrai_metadato(df, "Cliente")

                # 4️⃣ PRENDE SOLO LA PRIMA RIGA VALIDA
                row_valida = None
                for _, row in dati_df.iterrows():
                    if row.isna().all():
                        continue
                    if pd.isna(row[col_idx["Punt. Reale"]]):
                        continue
                    row_valida = row
                    break

                if row_valida is None:
                    st.warning(f"⚠️ {file.name} ignorato: Nessuna riga valida trovata.")
                    continue

                def _round(val):
                    if pd.isna(val) or val == "":
                        return None
                    try:
                        return round(float(val), decimali)
                    except:
                        return val

                risultati.append({
                    "File Origine": file.name,
                    "Cliente": cliente,
                    "Data Inizio": data_inizio,
                    "Data Fine": data_fine,
                    "Punt. Reale": _round(row_valida[col_idx["Punt. Reale"]]),
                    "Punt. B1d": _round(row_valida[col_idx["Punt. B1d"]]),
                    "Circolati": _round(row_valida[col_idx["Circolati"]]),
                    "In Fascia": _round(row_valida[col_idx["In Fascia"]]),
                    "%OS (soglia propria)": _round(row_valida[col_idx["%OS (soglia propria)"]]),
                    "T Rit": _round(row_valida[col_idx["T Rit"]]),
                    "T Eff": _round(row_valida[col_idx["T Eff"]]),
                    "Fuori Fascia Per Esterne": _round(row_valida[col_idx["Fuori Fascia Per Esterne"]])
                })

                file_validi += 1

            except Exception as e:
                st.error(f"❌ Errore su {file.name}: {e}")

    # --- OUTPUT ---
    if risultati:
        st.success(f"✅ File validi: {file_validi} | Righe generate: {len(risultati)}")

        df_finale = pd.DataFrame(risultati)
        st.dataframe(df_finale, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_finale.to_excel(writer, index=False, sheet_name='Riepilogo')

        st.download_button(
            label="📥 Scarica Excel",
            data=output.getvalue(),
            file_name="riepilogo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Nessun dato valido estratto.")
