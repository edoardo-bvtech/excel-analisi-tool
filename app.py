import streamlit as st
import pandas as pd

st.title("Analisi Excel - Punt. B1d & %OS")

uploaded_files = st.file_uploader(
    "Carica file Excel",
    accept_multiple_files=True,
    type=["xlsx"]
)

if uploaded_files:
    risultati = []

    for file in uploaded_files:
        df = pd.read_excel(file, header=None)

        if pd.notna(df.iloc[0, 0]):
            st.warning(f"{file.name} scartato (A1 non vuota)")
            continue

        header_row = df[df.eq("Punt. B1d").any(axis=1)].index[0]
        dati_row = header_row + 1

        intestazione = df.iloc[header_row]

        punt_col = intestazione[intestazione == "Punt. B1d"].index[0]
        os_col = intestazione[intestazione == "%OS (soglia propria)"].index[0]

        risultati.append({
            "File": file.name,
            "Punt. B1d": df.iloc[dati_row, punt_col],
            "%OS": df.iloc[dati_row, os_col]
        })

    if risultati:
        df_finale = pd.DataFrame(risultati)
        st.dataframe(df_finale)

        st.download_button(
            "Scarica risultato",
            df_finale.to_csv(index=False),
            "risultato.csv",
            "text/csv"
        )
