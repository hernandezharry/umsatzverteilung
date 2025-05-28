
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime

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

    st.subheader("📋 Verteilungstabelle")
    start = df["Beginn"].min()
    end = df["Ende"].max()
    monate = pd.date_range(start=start, end=end, freq="MS")
    colnames = [d.strftime("%Y-%m") for d in monate]
    df_matrix = pd.DataFrame(index=df["Projekt"], columns=colnames).fillna(0.0)

    for _, row in df.iterrows():
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

    st.subheader("📅 Projekt-Zeitleiste (Gantt-Diagramm)")
    gantt_fig = go.Figure()

    for _, row in df.iterrows():
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
