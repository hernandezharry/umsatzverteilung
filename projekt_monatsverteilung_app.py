
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
from datetime import datetime

st.set_page_config(page_title="Projekt-Monatsverteilung", layout="wide")

uploaded_file = st.file_uploader("📄 Excel-Datei mit Projekten hochladen", type=["xls", "xlsx", "xlsm"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = next((s for s in xls.sheet_names if "projekt" in s.lower()), None)

    if not sheet_name:
        st.error("Kein Blatt mit 'Projekt' im Namen gefunden.")
        st.stop()

    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    df = df.dropna(subset=["Projekt", "Beginn", "Ende", "Auftragssumme"])
    df["Beginn"] = pd.to_datetime(df["Beginn"])
    df["Ende"] = pd.to_datetime(df["Ende"])
    df["Jahr"] = df["Beginn"].dt.year

    # Filter
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
        st.warning("Keine Daten für die Auswahl.")
        st.stop()

    # Monatsverteilung (Visualisierung 1)
    st.subheader("📋 Verteilungstabelle (Projekt × Monat)")
    start = df_filtered["Beginn"].min()
    end = df_filtered["Ende"].max()
    monate = pd.date_range(start=start, end=end, freq="MS").to_list()
    colnames = [d.strftime("%Y-%m") for d in monate]

    df_matrix = pd.DataFrame(index=df_filtered["Projekt"], columns=colnames).fillna(0.0)

    for _, row in df_filtered.iterrows():
        start_m = row["Beginn"].replace(day=1)
        end_m = row["Ende"].replace(day=1)
        anz = ((end_m.year - start_m.year) * 12 + end_m.month - start_m.month + 1)
        if anz <= 0: continue
        wert = row["Auftragssumme"] / anz
        lauf = start_m
        for _ in range(anz):
            key = lauf.strftime("%Y-%m")
            if key in df_matrix.columns:
                df_matrix.loc[row["Projekt"], key] += wert
            lauf = (lauf + pd.DateOffset(months=1)).replace(day=1)

    
    st.markdown(
        df_matrix.style
        .format("€ {:,.2f}")
        .set_table_attributes('style="color: white; font-size: 16px; background-color: #1e1e1e;"')
        .to_html(),
        unsafe_allow_html=True
    )
    

    csv = df_matrix.reset_index().to_csv(index=False).encode("utf-8")
    st.download_button("📥 CSV-Verteilung", data=csv, file_name="verteilung.csv", mime="text/csv")

    # Projektkosten (Visualisierung 2)
    kosten_cols = ["Herstellkosten", "Ergebnis", "Gewährleistung"]
    vorhanden = [c for c in kosten_cols if c in df_filtered.columns]

    if vorhanden:
        st.subheader("💰 Projektkosten-Vergleich")

        df_kosten = df_filtered[["Projekt", "Auftragssumme"] + vorhanden].copy()
        if "Ergebnis" in df_kosten.columns:
            df_kosten["Ergebnis"] = (df_kosten["Ergebnis"] / 100) * df_kosten["Auftragssumme"]
        if "Gewährleistung" in df_kosten.columns:
            df_kosten["Gewährleistung"] = (df_kosten["Gewährleistung"] / 100) * df_kosten["Auftragssumme"]

        df_kosten = df_kosten.groupby("Projekt").sum(numeric_only=True)[vorhanden]

        fig2 = go.Figure()
        farben = {
            "Herstellkosten": "#1f77b4",
            "Ergebnis": "#2ca02c",
            "Gewährleistung": "#ff7f0e"
        }

        for col in df_kosten.columns:
            fig2.add_trace(go.Bar(
                x=df_kosten.index,
                y=df_kosten[col],
                name=col,
                marker_color=farben.get(col, "#999999")
            ))

        fig2.update_layout(
            barmode="group",
            title="Projektkosten nach Typ",
            xaxis_title="Projekt",
            yaxis_title="Betrag (€)",
            font=dict(size=13, color="black"),
            height=500,
            plot_bgcolor="white",
            paper_bgcolor="white"
        )
        st.plotly_chart(fig2, use_container_width=True)

    # Gantt-Diagramm (Visualisierung 3)
    st.subheader("📅 Projekt-Zeitleiste (Gantt-Diagramm)")

    gantt_data = df_filtered.copy()
    gantt_data["Beginn"] = pd.to_datetime(gantt_data["Beginn"])
    gantt_data["Ende"] = pd.to_datetime(gantt_data["Ende"])
    gantt_data["Dauer"] = (gantt_data["Ende"] - gantt_data["Beginn"]).dt.days

    gantt_fig = go.Figure()
    for _, row in gantt_data.iterrows():
        text = f"{row['Projekt']}<br>Herstellkosten: {row.get('Herstellkosten', 'n.v.')}<br>Ergebnis: {row.get('Ergebnis', 'n.v.')}%<br>Gewährleistung: {row.get('Gewährleistung', 'n.v.')}%"
        gantt_fig.add_trace(go.Bar(
            x=[row["Dauer"]],
            y=[row["Projekt"]],
            base=row["Beginn"],
            orientation="h",
            hovertemplate=text + "<extra></extra>",
            marker_color="#1f77b4"
        ))

    gantt_fig.update_layout(
        title="Projektzeiträume",
        xaxis_title="Datum",
        yaxis_title="Projekt",
        xaxis=dict(type="date"),
        font=dict(size=13, color="black"),
        height=600,
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=False
    )

    st.plotly_chart(gantt_fig, use_container_width=True)
