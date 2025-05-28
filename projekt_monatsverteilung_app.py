
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Projekt-Monatsverteilung & Gantt-Analyse")

uploaded_file = st.file_uploader("📁 Excel-Datei mit Projekt-Daten hochladen", type=["xls", "xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Projekte")

    df["Beginn"] = pd.to_datetime(df["Beginn"], errors="coerce")
    df["Ende"] = pd.to_datetime(df["Ende"], errors="coerce")
    df = df.dropna(subset=["Beginn", "Ende"])

    st.subheader("📅 Monatsverteilung")
    all_start = df["Beginn"].min()
    all_end = df["Ende"].max()
    month_range = pd.date_range(all_start, all_end, freq="MS")
    month_labels = [d.strftime("%Y-%m") for d in month_range]

    df_matrix = pd.DataFrame(0, index=df["Projekt"], columns=month_labels, dtype=float)

    for _, row in df.iterrows():
        proj = row["Projekt"]
        betrag = row["Auftragssumme"] if pd.notnull(row["Auftragssumme"]) else 0
        months = pd.date_range(row["Beginn"], row["Ende"], freq="MS")
        if len(months) == 0:
            continue
        share = betrag / len(months)
        for m in months:
            label = m.strftime("%Y-%m")
            if label in df_matrix.columns:
                df_matrix.loc[proj, label] += share

    st.dataframe(df_matrix.style.format("€ {:,.2f}"))

    st.subheader("📊 Gestapeltes Balkendiagramm (nach Phase)")

    df_plot = df.copy()
    df_plot["Monat"] = df_plot["Beginn"].dt.to_period("M").astype(str)
    df_plot = df_plot.groupby(["Monat", "Phase"]).agg({"Auftragssumme": "sum"}).reset_index()

    farben = {
        "Ausführung": "#0070C0",
        "Verhandlung": "#FFC000",
        "Angebotsbearbeitung": "#00B050",
        "Anfrage": "#9966FF",
        "Marktbeobachtung": "#BFBFBF"
    }

    fig = px.bar(df_plot, x="Monat", y="Auftragssumme", color="Phase", color_discrete_map=farben)
    fig.update_layout(
        barmode="stack",
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(color="black"),
        title="Projektvolumen pro Monat nach Phase",
        xaxis_title="Monat",
        yaxis_title="Betrag (€)",
        legend_title="Phase"
    )
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("📆 Gantt-Diagramm")

    df_gantt = df.copy()
    df_gantt = df_gantt.sort_values("Beginn")
    df_gantt["ID"] = df_gantt.index

    gantt_fig = go.Figure()
    for i, row in df_gantt.iterrows():
        gantt_fig.add_trace(go.Scatter(
            x=[row["Beginn"], row["Ende"]],
            y=[row["Projekt"], row["Projekt"]],
            mode="lines",
            line=dict(width=20, color=farben.get(row["Phase"], "gray")),
            showlegend=False
        ))

    gantt_fig.update_layout(
        title="Projektzeiträume",
        font=dict(size=13, color="black"),
        paper_bgcolor="white",
        plot_bgcolor="white",
        xaxis_title="Datum",
        yaxis=dict(
            title="Projekt",
            tickfont=dict(color="black")
        ),
        height=600
    )
    st.plotly_chart(gantt_fig, use_container_width=True)
