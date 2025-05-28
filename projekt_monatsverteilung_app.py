
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide")
st.title("📊 Projekt-Monatsverteilung & Gantt-Diagramm")

uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xlsm"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)

    st.subheader("📋 Eingelesene Projektdaten")
    st.dataframe(df)

    # 📆 Gantt-Diagramm mit Phasenfarben
    st.subheader("📆 Projekt-Zeitleiste (Gantt-Diagramm mit Phase-Farben)")

    farben = {
        "Ausführung": "#0070C0",
        "Verhandlung": "#FFC000",
        "Angebotsbearbeitung": "#00B050",
        "Anfrage": "#9966FF",
        "Marktbeobachtung": "#BFBFBF"
    }

    df_gantt = df[["Projekt", "Beginn", "Ende", "Phase"]].copy()
    df_gantt = df_gantt.dropna(subset=["Beginn", "Ende", "Phase"])

    df_gantt["Beginn"] = pd.to_datetime(df_gantt["Beginn"])
    df_gantt["Ende"] = pd.to_datetime(df_gantt["Ende"])

    fig_gantt = px.timeline(
        df_gantt,
        x_start="Beginn",
        x_end="Ende",
        y="Projekt",
        color="Phase",
        color_discrete_map=farben,
        title="Projektzeitleiste nach Phase"
    )

    fig_gantt.update_layout(
        height=600,
        xaxis_title="Zeitraum",
        yaxis_title="Projekt",
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(color="black"),
        legend_title="Phase"
    )

    fig_gantt.update_yaxes(autorange="reversed")
    st.plotly_chart(fig_gantt, use_container_width=True)
else:
    st.info("Bitte eine Excel-Datei mit Projektinformationen hochladen.")
