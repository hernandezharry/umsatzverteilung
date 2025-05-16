
import streamlit as st
import pandas as pd
import calendar
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Projekt-Monatsverteilung", layout="wide")

# Datei-Upload
uploaded_file = st.file_uploader("📄 Excel-Datei mit Projekten hochladen", type=["xls", "xlsx", "xlsm"])

if uploaded_file:
    # Automatisch das passende Arbeitsblatt finden
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = None
    for name in xls.sheet_names:
        if "projekt" in name.lower():
            sheet_name = name
            break
    if not sheet_name:
        st.error("Kein Arbeitsblatt mit 'Projekt' im Namen gefunden.")
        st.stop()

    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
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
        monat_colnames = [d.strftime("%Y-%m") for d in monatsliste]

        # Basisstruktur für Verteilung
        phasen_farben = {}
        df_result = pd.DataFrame(0.0, index=monat_colnames, columns=[])
        grouped = df_filtered.groupby("Phase")

        for phase, group in grouped:
            phase_series = pd.Series(0.0, index=monat_colnames)
            for _, row in group.iterrows():
                start_monat = row["Beginn"].replace(day=1)
                end_monat = row["Ende"].replace(day=1)
                anz_monate = ((end_monat.year - start_monat.year) * 12 + end_monat.month - start_monat.month + 1)
                if anz_monate <= 0:
                    continue
                wert_pro_monat = row["Auftragssumme"] / anz_monate
                lauf = start_monat
                for _ in range(anz_monate):
                    key = lauf.strftime("%Y-%m")
                    if key in phase_series.index:
                        phase_series[key] += wert_pro_monat
                    lauf = (lauf + pd.DateOffset(months=1)).replace(day=1)
            df_result[phase] = phase_series

        st.subheader("📊 Monatsverteilung gestapelt nach Phase")
        st.dataframe(df_result.fillna(0.0).style.format("€ {:,.2f}"))

        # Gestapeltes Balkendiagramm mit Plotly
        import plotly.express as px
        import plotly.graph_objects as go

        fig = go.Figure()
        for phase in df_result.columns:
            fig.add_trace(go.Bar(
                name=phase,
                x=df_result.index,
                y=df_result[phase],
                hovertemplate=f"%{{x}}<br>{phase}: € %{{y:,.2f}}<extra></extra>"
            ))

        fig.update_layout(
            barmode='stack',
            title="Monatliche Auftragssumme – gestapelt nach Phase",
            xaxis_title="Monat",
            yaxis_title="Auftragssumme (€)",
            legend_title="Phase",
            hovermode="x unified"
        )
        st.plotly_chart(fig, use_container_width=True)

        csv = df_result.reset_index().rename(columns={"index": "Monat"}).to_csv(index=False).encode("utf-8")
        st.download_button("📥 CSV herunterladen", csv, "monatsverteilung_phasen.csv", "text/csv")
