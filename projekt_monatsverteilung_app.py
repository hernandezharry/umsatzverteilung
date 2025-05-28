
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime

st.title("📊 Übersicht Umsatzverteilung")

st.set_page_config(page_title="Projekt-Monatsverteilung", layout="wide")

uploaded_file = st.file_uploader("📄 Excel-Datei mit Projekten hochladen", type=["xls", "xlsx", "xlsm"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = next((s for s in xls.sheet_names if "projekt" in s.lower()), xls.sheet_names[0])
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    df = df.dropna(subset=["Projekt", "Beginn", "Ende", "Auftragssumme"])
    df["Beginn"] = pd.to_datetime(df["Beginn"], errors="coerce")
    df["Ende"] = pd.to_datetime(df["Ende"], errors="coerce")
    df = df.dropna(subset=["Beginn", "Ende"])
    df["Jahr"] = df["Beginn"].dt.year

    # === Filter ===
    phasen = ["Alle"] + sorted(df["Phase"].dropna().unique().tolist()) if "Phase" in df.columns else ["Alle"]
    jahre = ["Alle"] + sorted(df["Jahr"].dropna().unique().astype(str).tolist())

    col1, col2 = st.columns(2)
    selected_phase = col1.selectbox("🔍 Phase auswählen", phasen)
    selected_year = col2.selectbox("📅 Jahr auswählen", jahre)

    df_filtered = df.copy()
    if selected_phase != "Alle" and "Phase" in df_filtered.columns:
        df_filtered = df_filtered[df_filtered["Phase"] == selected_phase]
    if selected_year != "Alle":
        df_filtered = df_filtered[df_filtered["Beginn"].dt.year == int(selected_year)]

    if df_filtered.empty:
        st.warning("Keine Daten für die Auswahl.")
        st.stop()

    # === Monatsverteilung ===
    st.subheader("📋 Verteilungstabelle")
    start = df_filtered["Beginn"].min()
    end = df_filtered["Ende"].max()
    monate = pd.date_range(start=start, end=end, freq="MS")
    colnames = [d.strftime("%Y-%m") for d in monate]
    df_matrix = pd.DataFrame(index=df_filtered["Projekt"], columns=colnames).fillna(0.0)

    for _, row in df_filtered.iterrows():
        start_m = row["Beginn"].replace(day=1)
        end_m = row["Ende"].replace(day=1)
        anz = ((end_m.year - start_m.year) * 12 + end_m.month - start_m.month + 1)
        if anz <= 0:
            continue
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
    st.download_button("📥 CSV-Verteilung", data=df_matrix.reset_index().to_csv(index=False).encode("utf-8"),
                       file_name="verteilung.csv", mime="text/csv")

    # === Gestapeltes Balkendiagramm (nach Phase) ===
    st.subheader("📊 Monatsverteilung (gestapelt nach Phase)")
    df_result = pd.DataFrame(0.0, index=colnames, columns=[])
    grouped = df_filtered.groupby("Phase") if "Phase" in df_filtered.columns else {}

    for phase, group in grouped:
        phase_series = pd.Series(0.0, index=colnames)
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

    farben_phase = {
        "Ausführung": "#1f77b4",
        "Verhandlung": "#ff7f0e",
        "Angebotsbearbeitung": "#2ca02c",
        "Anfrage": "#d62728",
        "Marktbeobachtung": "#9467bd"
    }

    phasen_sortiert = ["Ausführung", "Verhandlung", "Angebotsbearbeitung", "Anfrage", "Marktbeobachtung"]
    phasen_final = [p for p in phasen_sortiert if p in df_result.columns] + [p for p in df_result.columns if p not in phasen_sortiert]

    fig = go.Figure()
    for phase in phasen_final:
        fig.add_trace(go.Bar(
            name=phase,
            x=df_result.index,
            y=df_result[phase],
            marker_color=farben_phase.get(phase, "#999999"),
            hovertemplate=f"%{{x}}<br>{phase}: € %{{y:,.2f}}<extra></extra>"
        ))

    fig.update_layout(
        barmode="stack",
        title="Monatliche Auftragssumme – gestapelt nach Phase",
        font=dict(size=13, color="white"),
        paper_bgcolor="#0e1117",
        plot_bgcolor="#0e1117",
        xaxis_title="Monat",
        xaxis=dict(tickfont=dict(color="white")),
        yaxis_title="Auftragssumme (€)",
        yaxis=dict(tickfont=dict(color="white")),
        legend=dict(font=dict(color="white")),
        height=500
    )
    st.plotly_chart(fig, use_container_width=True)

    # === Projektkosten (Herstellkosten, Ergebnis, Gewährleistung) ===
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
            font=dict(size=13, color="white"),
            paper_bgcolor="#0e1117",
            plot_bgcolor="#0e1117",
            xaxis_title="Projekt",
            xaxis=dict(tickfont=dict(color="white")),
            yaxis_title="Betrag (€)",
            yaxis=dict(tickfont=dict(color="white")),
            legend=dict(font=dict(color="white")),
            height=500
        )
        st.plotly_chart(fig2, use_container_width=True)

    # === Gantt-Diagramm ===
    st.subheader("📅 Projekt-Zeitleiste (Gantt-Diagramm)")
    gantt_fig = go.Figure()

    for _, row in df_filtered.iterrows():
        if row["Ende"] <= row["Beginn"]:
            continue
        gantt_fig.add_trace(go.Scatter(
            x=[row["Beginn"], row["Ende"]],
            y=[row["Projekt"], row["Projekt"]],
            mode="lines",
            line=dict(color="#1f77b4", width=10),
            hovertemplate=(
                f"{row['Projekt']}<br>"
                f"Start: {row['Beginn'].date()}<br>"
                f"Ende: {row['Ende'].date()}<br>"
                f"Summe: € {row['Auftragssumme']:,.2f}<extra></extra>"
            ),
            showlegend=False
        ))

    gantt_fig.update_layout(
        title="Projektzeiträume",
        font=dict(size=13, color="white"),
        paper_bgcolor="#0e1117",
        plot_bgcolor="#0e1117",
        xaxis_title="Datum",
        xaxis=dict(type="date", tickfont=dict(color="white")),
        yaxis_title="Projekt",
        yaxis=dict(tickfont=dict(color="white")),
        height=600
    )
    st.plotly_chart(gantt_fig, use_container_width=True)
