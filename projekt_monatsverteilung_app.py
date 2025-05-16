
import streamlit as st
import pandas as pd
import calendar
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Projekt-Monatsverteilung", layout="wide")

# Datei-Upload
uploaded_file = st.file_uploader("📄 Excel-Datei mit Projekten hochladen", type=["xls", "xlsx", "xlsm"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Projekte")
    df = df.dropna(subset=["Projekt", "Beginn", "Ende", "Auftragssumme"])
    df["Beginn"] = pd.to_datetime(df["Beginn"])
    df["Ende"] = pd.to_datetime(df["Ende"])
    df["Jahr"] = df["Beginn"].dt.year
    phasen = ["Alle"] + sorted(df["Phase"].dropna().unique().tolist())
    jahre = ["Alle"] + sorted(df["Jahr"].dropna().unique().astype(str).tolist())

    col1, col2 = st.columns(2)
    selected_phase = col1.selectbox("🔍 Phase auswählen", phasen)
    selected_year = col2.selectbox("📅 Jahr auswählen", jahre)

    df_filtered = df.copy()
    if selected_phase != "Alle":
        df_filtered = df_filtered[df_filtered["Phase"] == selected_phase]
    if selected_year != "Alle":
        df_filtered = df_filtered[df_filtered["Beginn"].dt.year == int(selected_year)]

    if df_filtered.empty:
        st.warning("Keine Projekte für die gewählten Filter gefunden.")
    else:
        start = df_filtered["Beginn"].min()
        end = df_filtered["Ende"].max()
        monatsliste = pd.date_range(start=start, end=end, freq='MS').to_list()
        verteilung = pd.DataFrame(index=df_filtered["Projekt"], columns=[d.strftime("%Y-%m") for d in monatsliste])
        verteilung = verteilung.fillna(0.0)

        for _, row in df_filtered.iterrows():
            start_monat = row["Beginn"].replace(day=1)
            end_monat = row["Ende"].replace(day=1)
            anz_monate = ((end_monat.year - start_monat.year) * 12 + end_monat.month - start_monat.month + 1)
            if anz_monate <= 0:
                continue
            wert_pro_monat = row["Auftragssumme"] / anz_monate
            lauf = start_monat
            for _ in range(anz_monate):
                key = lauf.strftime("%Y-%m")
                if key in verteilung.columns:
                    verteilung.loc[row["Projekt"], key] += wert_pro_monat
                lauf = (lauf + pd.DateOffset(months=1)).replace(day=1)

        st.subheader("📊 Monatsverteilung")
        st.dataframe(verteilung.style.format("€ {:,.2f}"))

        st.bar_chart(verteilung.sum(axis=0))

        csv = verteilung.reset_index().to_csv(index=False).encode("utf-8")
        st.download_button("📥 CSV herunterladen", csv, "monatsverteilung.csv", "text/csv")
