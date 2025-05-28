
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide")
st.title("📊 Projekt-Monatsverteilung & Visualisierung")

uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx", "xlsm"])
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.subheader("📋 Eingelesene Projektdaten")
        st.dataframe(df)

        # Erwartete Spalten
        erwartete_spalten = ["Projekt", "Beginn", "Ende", "Phase", "Auftragssumme", "Ergebnis", "Gewährleistung", "Herstellkosten"]
        fehlende_spalten = [sp for sp in erwartete_spalten if sp not in df.columns]
        if fehlende_spalten:
            st.error(f"❌ Fehlende Spalten: {fehlende_spalten}")
        else:
            df = df.dropna(subset=["Projekt", "Beginn", "Ende", "Phase"])

            farben = {
                "Ausführung": "#0070C0",
                "Verhandlung": "#FFC000",
                "Angebotsbearbeitung": "#00B050",
                "Anfrage": "#9966FF",
                "Marktbeobachtung": "#BFBFBF"
            }

            st.subheader("📆 Projekt-Zeitleiste (Gantt-Diagramm)")
            df["Beginn"] = pd.to_datetime(df["Beginn"])
            df["Ende"] = pd.to_datetime(df["Ende"])

            fig_gantt = px.timeline(
                df,
                x_start="Beginn",
                x_end="Ende",
                y="Projekt",
                color="Phase",
                color_discrete_map=farben,
                title="Projektzeitleiste nach Phase"
            )
            fig_gantt.update_layout(
                height=700,
                xaxis_title="Zeitraum",
                yaxis_title="Projekt",
                plot_bgcolor="white",
                paper_bgcolor="white",
                font=dict(color="black", size=14),
                legend_title="Phase"
            )
            fig_gantt.update_yaxes(autorange="reversed")
            st.plotly_chart(fig_gantt, use_container_width=True)

            st.subheader("📊 Gestapeltes Balkendiagramm (Monatsverteilung)")
            df["Monat"] = df["Beginn"].dt.to_period("M").astype(str)
            df_grouped = df.groupby(["Monat", "Phase"]).agg({"Auftragssumme": "sum"}).reset_index()

            fig_bar = px.bar(
                df_grouped,
                x="Monat",
                y="Auftragssumme",
                color="Phase",
                color_discrete_map=farben,
                title="Monatliche Auftragssummen nach Phase"
            )
            fig_bar.update_layout(
                barmode="stack",
                xaxis_title="Monat",
                yaxis_title="Auftragssumme (€)",
                font=dict(color="black", size=14),
                plot_bgcolor="white",
                paper_bgcolor="white"
            )
            st.plotly_chart(fig_bar, use_container_width=True)

            st.subheader("📉 Weitere Auswertung: Ergebnis & Gewährleistung")
            df_percent = df[["Projekt", "Ergebnis", "Gewährleistung"]].dropna()
            df_melted = df_percent.melt(id_vars="Projekt", value_vars=["Ergebnis", "Gewährleistung"],
                                        var_name="Kennzahl", value_name="Prozent")

            fig_percent = px.bar(
                df_melted,
                x="Projekt",
                y="Prozent",
                color="Kennzahl",
                barmode="group",
                title="Ergebnis & Gewährleistung je Projekt (%)"
            )
            fig_percent.update_layout(
                font=dict(color="black", size=14),
                plot_bgcolor="white",
                paper_bgcolor="white",
                xaxis_tickangle=45
            )
            st.plotly_chart(fig_percent, use_container_width=True)

    except Exception as e:
        st.error(f"Fehler beim Einlesen der Datei: {e}")
else:
    st.info("Bitte eine Excel-Datei hochladen.")
