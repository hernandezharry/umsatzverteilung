
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(layout="wide")
st.title("📊 Umsatzverteilung & Projektanalyse")

uploaded_file = st.file_uploader("📂 Excel-Datei mit Projektliste hochladen", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Datumsfelder konvertieren
    df["Beginn"] = pd.to_datetime(df["Beginn"], errors="coerce")
    df["Ende"] = pd.to_datetime(df["Ende"], errors="coerce")

    st.subheader("📊 Gestapeltes Balkendiagramm (nach Phase)")

    phase_order = ["Ausführung", "Verhandlung", "Angebotsbearbeitung", "Anfrage", "Marktbeobachtung"]
    farben = {
        "Ausführung": "#0070C0",
        "Verhandlung": "#FFC000",
        "Angebotsbearbeitung": "#00B050",
        "Anfrage": "#9966FF",
        "Marktbeobachtung": "#BFBFBF"
    }

    df_bar = df.copy()
    df_bar["Phase"] = pd.Categorical(df_bar["Phase"], categories=phase_order, ordered=True)
    df_bar["Monat"] = df_bar["Beginn"].dt.to_period("M").astype(str)

    df_bar_grouped = df_bar.groupby(["Monat", "Phase"]).agg({"Auftragssumme": "sum"}).reset_index()

    fig = px.bar(
        df_bar_grouped,
        x="Monat",
        y="Auftragssumme",
        color="Phase",
        category_orders={"Phase": phase_order},
        color_discrete_map=farben
    )

    fig.update_layout(
        barmode="stack",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="white"),
        title="Projektvolumen pro Monat nach Phase",
        xaxis_title="Monat",
        yaxis_title="Auftragssumme (€)",
        legend_title="Phase",
        legend=dict(font=dict(color="white"))
    )

    st.plotly_chart(fig, use_container_width=True)

    st.subheader("📆 Gantt-Diagramm")

    df_gantt = df.copy()
    df_gantt = df_gantt[df_gantt["Beginn"].notna() & df_gantt["Ende"].notna()]

    df_gantt["Phase"] = pd.Categorical(df_gantt["Phase"], categories=phase_order, ordered=True)

    gantt_fig = px.timeline(
        df_gantt,
        x_start="Beginn",
        x_end="Ende",
        y="Projekt",
        color="Phase",
        color_discrete_map=farben,
        hover_data=["Phase", "Auftragssumme", "Ergebnis", "Gewährleistung"]
    )

    gantt_fig.update_yaxes(autorange="reversed", tickfont=dict(color="white"))
    gantt_fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="white"),
        xaxis_title="Zeitraum",
        yaxis_title="Projekt",
        title="Projektzeitplan (Gantt-Diagramm)",
        legend_title="Phase",
        legend=dict(font=dict(color="white"))
    )

    st.plotly_chart(gantt_fig, use_container_width=True)
