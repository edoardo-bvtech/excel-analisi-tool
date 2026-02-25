import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Analisi Excel - Punt. B1d & %OS")

# Selettore numero decimali
decimali = st.number_input("Numero decimali", min_value=0, max_value=6, value=2)

HEADER_ATTESO = [
    "Viste", "Giorno", "Settimana", "Mese", "Anno",
    "Media Treni Circolati", "Punt. Reale", "Punt. B1d",
    "Punt. SB", "Punt. RFI", "Punt. IF",
    "Circolati", "In Fascia", "Fuori Fascia",
    "Fuori Fascia Per Esterne", "Fuori Fascia per RFI",
    "Fuori Fascia per Propria IF", "Fuori Fascia per Altre IF",
    "%OS (soglia propria)", "T Rit", "T Eff"
]

uploaded_files = st.file_uploader(
    "Carica file Excel",
    accept_multiple_files=True,
    type=["xlsx"]
)

def trova_valore_colonna_A(df, testo):
    for i in range(len(df)):
        valore = df.iloc[i, 0]
        if pd.notna(valore) and str(valore).strip().startswith(testo):
            # prende il valore nella colonna B (accanto)
            return df.iloc[i, 1]
    return None

if uploaded_files:
    risultati = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, header=None)

            header_row = None

            # 1️⃣ trova intestazione
            for i in range(len(df)):
                row = df.iloc[i].tolist()
                row_clean = [str(x).strip() for x in row]
                if row_clean[:len(HEADER_ATTESO)] == HEADER_ATTESO:
                    header_row = i
                    break

            if header_row is None:
                st.warning(f"{file.name} → intestazione non trovata")
                continue

            dati_row = header_row + 1

            if dati_row >= len(df):
                st.warning(f"{file.name} → riga dati mancante")
                continue

            # 2️⃣ controllo colonna A vuota
            cella_A = df.iloc[dati_row, 0]

            if pd.notna(cella_A) and str(cella_A).strip() != "":
                st.warning(f"{file.name} → riga dati non valida")
                continue

            intestazione = [str(x).strip() for x in df.iloc[header_row].tolist()]

            punt_col = intestazione.index("Punt. B1d")
            os_col = intestazione.index("%OS (soglia propria)")

            punt_value = df.iloc[dati_row, punt_col]
            os_value = df.iloc[dati_row, os_col]

            # arrotondamento
            if pd.notna(punt_value):
                punt_value = round(float(punt_value), decimali)

            if pd.notna(os_value):
                os_value = round(float(os_value), decimali)

            # 3️⃣ Estrazione dati da colonna A
            data_inizio = trova_valore_colonna_A(df, "Data Inizio")
            data_fine = trova_valore_colonna_A(df, "Data Fine")
            cliente = trova_valore_colonna_A(df, "Cliente")

            risultati.append({
                "File": file.name,
                "Cliente": cliente,
                "Data Inizio": data_inizio,
                "Data Fine": data_fine,
                "Punt. B1d": punt_value,
                "%OS (soglia propria)": os_value
            })

        except Exception as e:
            st.error(f"{file.name} → errore: {e}")

    if risultati:
        df_finale = pd.DataFrame(risultati)
        st.success("Elaborazione completata")
        st.dataframe(df_finale)

        # 🔥 Download Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_finale.to_excel(writer, index=False, sheet_name='Risultati')

        st.download_button(
            label="Scarica file Excel",
            data=output.getvalue(),
            file_name="risultato.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Nessun file valido elaborato.")
