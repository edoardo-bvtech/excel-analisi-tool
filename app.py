import streamlit as st
import pandas as pd

st.title("Analisi Excel - Punt. B1d & %OS")

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

if uploaded_files:
    risultati = []

    for file in uploaded_files:
        try:
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
                st.warning(f"{file.name} → intestazione non trovata")
                continue

            # 2️⃣ PRENDI RIGA SUCCESSIVA
            dati_row = header_row + 1

            if dati_row >= len(df):
                st.warning(f"{file.name} → riga dati mancante")
                continue

            # 3️⃣ CONTROLLO CELLA A DELLA RIGA DATI
            cella_A = df.iloc[dati_row, 0]

            if pd.notna(cella_A) and str(cella_A).strip() != "":
                st.warning(f"{file.name} → riga dati non valida (colonna A non vuota)")
                continue

            # 4️⃣ ESTRAZIONE COLONNE
            intestazione = df.iloc[header_row].tolist()
            intestazione_clean = [str(x).strip() for x in intestazione]

            punt_col = intestazione_clean.index("Punt. B1d")
            os_col = intestazione_clean.index("%OS (soglia propria)")

            punt_value = df.iloc[dati_row, punt_col]
            os_value = df.iloc[dati_row, os_col]

            risultati.append({
                "File": file.name,
                "Punt. B1d": punt_value,
                "%OS (soglia propria)": os_value
            })

        except Exception as e:
            st.error(f"{file.name} → errore: {e}")

    if risultati:
        df_finale = pd.DataFrame(risultati)
        st.success("Elaborazione completata")
        st.dataframe(df_finale)

        st.download_button(
            "Scarica risultato CSV",
            df_finale.to_csv(index=False),
            "risultato.csv",
            "text/csv"
        )
    else:
        st.info("Nessun file valido elaborato.")
