
import streamlit as st
import pandas as pd
import calendar
from datetime import datetime
from io import BytesIO
import plotly.graph_objects as go
import plotly.io as pio

st.set_page_config(page_title="Projekt-Monatsverteilung", layout="wide")

uploaded_file = st.file_uploader("📄 Excel-Datei mit Projekten hochladen", type=["xls", "xlsx", "xlsm"])

if uploaded_file:
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

        df_matrix = pd.DataFrame(index=df_filtered["Projekt"], columns=monat_colnames).fillna(0.0)

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
                if key in df_matrix.columns:
                    df_matrix.loc[row["Projekt"], key] += wert_pro_monat
                lauf = (lauf + pd.DateOffset(months=1)).replace(day=1)

        st.subheader("📋 Verteilungstabelle")
        st.dataframe(df_matrix.style.format("€ {:,.2f}"))

        st.subheader("📊 Gestapeltes Balkendiagramm (nach Phase)")
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

        # Farbpalette festlegen
        farben = {
    'Ausführung': '#1f77b4',
    'Verhandlung': '#ff7f0e',
    'Angebotsbearbeitung': '#2ca02c',
    'Anfrage': '#d62728',
    'Marktbeobachtung': '#9467bd'
}

        fig = go.Figure()
        phasen_reihenfolge = ['Ausführung', 'Verhandlung', 'Angebotsbearbeitung', 'Anfrage', 'Marktbeobachtung']
phasen_sortiert = [p for p in phasen_reihenfolge if p in df_result.columns]
for i, phase in enumerate(phasen_sortiert):
            fig.add_trace(go.Bar(
                name=phase,
                x=df_result.index,
                y=df_result[phase],
                marker_color=farben.get(phase, '#999999'),
                hovertemplate=f"%{{x}}<br>{phase}: € %{{y:,.2f}}<extra></extra>"
            ))

        fig.update_layout(
            barmode='stack',
            title="Monatliche Auftragssumme – gestapelt nach Phase",
            xaxis_title="Monat",
            yaxis_title="Auftragssumme (€)",
            legend_title="Phase",
            hovermode="x unified",
            plot_bgcolor="white",
            paper_bgcolor="white"
        )

        st.plotly_chart(fig, use_container_width=True)

        # Export als PNG mit Layout
        img_bytes = pio.to_image(fig, format="png", width=1200, height=600, scale=2)
        st.download_button("🖼️ Diagramm als PNG herunterladen", data=img_bytes, file_name="diagramm_projektverteilung.png", mime="image/png")

        csv = df_result.reset_index().rename(columns={"index": "Monat"}).to_csv(index=False).encode("utf-8")
        st.download_button("📥 CSV herunterladen", csv, "monatsverteilung_phasen.csv", "text/csv")


# Zweite Visualisierung: Projekt-Kostenstruktur
st.subheader("📊 Projektkosten-Visualisierung")

kosten_spalten = ["Herstellkosten", "Gewährleistung", "Ergebnis"]
vorhandene_kosten = [col for col in kosten_spalten if col in df_filtered.columns]

if vorhandene_kosten:
    df_kosten = df_filtered[["Projekt", "Auftragssumme"] + vorhandene_kosten].copy()
# Prozentuelle Umrechnung
if "Ergebnis" in df_kosten.columns:
    df_kosten["Ergebnis"] = df_kosten["Ergebnis"] / 100 * df_kosten["Auftragssumme"]
if "Gewährleistung" in df_kosten.columns:
    df_kosten["Gewährleistung"] = df_kosten["Gewährleistung"] / 100 * df_kosten["Auftragssumme"]

    df_kosten = df_kosten.groupby("Projekt").sum(numeric_only=True)
    df_kosten = df_kosten[vorhandene_kosten]

    fig2 = go.Figure()
    farben_kosten = {
        "Herstellkosten": "#1f77b4",
        "Gewährleistung": "#ff7f0e",
        "Ergebnis": "#2ca02c"
    }

    for spalte in vorhandene_kosten:
        fig2.add_trace(go.Bar(
            x=df_kosten.index,
            y=df_kosten[spalte],
            name=spalte,
            marker_color=farben_kosten.get(spalte, "#999999"),
            hovertemplate=f"{spalte}: € %{{y:,.2f}}<extra></extra>"
        ))

    fig2.update_layout(
        barmode='group',
        title="Projektkosten nach Typ",
        xaxis_title="Projekt",
        yaxis_title="€ Betrag",
        legend_title="Kostenart",
        plot_bgcolor="white",
        paper_bgcolor="white",
        hovermode="x unified"
    )

    st.plotly_chart(fig2, use_container_width=True)

    # Download als PNG
    img_bytes_2 = pio.to_image(fig2, format="png", width=1200, height=600, scale=2)
    st.download_button("🖼️ Kosten-Diagramm als PNG herunterladen", data=img_bytes_2, file_name="projektkosten.png", mime="image/png")

    # CSV Export
    csv2 = df_kosten.reset_index().to_csv(index=False).encode("utf-8")
    st.download_button("📥 Kosten-Daten als CSV", data=csv2, file_name="projektkosten.csv", mime="text/csv")
else:
    st.info("Keine der Spalten 'Herstellkosten', 'Gewährleistung', 'Ergebnis' in den Daten gefunden.")
