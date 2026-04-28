# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide")

st.title("📊 Suivi des projets")

file = st.file_uploader("📂 Importer Excel brut", type=["xlsx"])

if file:

    df = pd.read_excel(file)

    # ==============================
    # RENOMMAGE PROPRE
    # ==============================
    df = df.rename(columns={
        "Nom de tâche": "Projet",
        "Créé par": "Responsable",
        "Progression": "Statut_brut",
        "Description": "Description"
    })

    # ==============================
    # AVANCEMENT (conversion texte → %)
    # ==============================
    def convert_progress(x):
        if isinstance(x, str):
            if "non démarr" in x.lower(): return 0
            if "en cours" in x.lower(): return 50
            if "termin" in x.lower(): return 100
        return 0

    df["Avancement"] = df["Statut_brut"].apply(convert_progress)

    # ==============================
    # STATUT
    # ==============================
    def statut(x):
        if x == 0: return "Non démarré"
        elif x == 100: return "Terminé"
        elif x < 40: return "En retard"
        else: return "En cours"

    df["Statut"] = df["Avancement"].apply(statut)

    # ==============================
    # CLEAN
    # ==============================
    df["Responsable"] = df["Responsable"].fillna("Non défini")

    df_clean = df[[
        "Projet",
        "Responsable",
        "Avancement",
        "Statut"
    ]]

    # ==============================
    # FILTRE RESPONSABLE
    # ==============================
    st.subheader("🎯 Filtre")

    chefs = df_clean["Responsable"].unique()
    chef = st.selectbox("Choisir un responsable", ["Tous"] + list(chefs))

    if chef != "Tous":
        dff = df_clean[df_clean["Responsable"] == chef]
    else:
        dff = df_clean

    # ==============================
    # KPI
    # ==============================
    st.subheader("📈 Indicateurs")

    col1, col2, col3 = st.columns(3)

    col1.metric("Projets", len(dff))
    col2.metric("Avancement moyen", f"{dff['Avancement'].mean():.0f}%")
    col3.metric("En retard", len(dff[dff["Statut"]=="En retard"]))

    # ==============================
    # GRAPHIQUES
    # ==============================
    st.subheader("📊 Graphiques")

    fig1 = px.pie(dff, names="Statut", hole=0.5)

    fig2 = px.bar(
        dff.groupby("Responsable")["Projet"].count().reset_index(),
        x="Responsable", y="Projet",
        title="Charge par responsable"
    )

    st.plotly_chart(fig1, use_container_width=True)
    st.plotly_chart(fig2, use_container_width=True)

    # ==============================
    # TABLEAU
    # ==============================
    st.subheader("📋 Tableau structuré")
    st.dataframe(dff, use_container_width=True)

    # ==============================
    # TABLEAU OUTLOOK
    # ==============================
    st.subheader("📧 Tableau pour Outlook")

    html = "<table style='border-collapse:collapse;font-family:Calibri;width:100%'>"
    html += "<tr style='background:#0A2463;color:white'>"

    for col in dff.columns:
        html += f"<th style='padding:8px;border:1px solid #ddd'>{col}</th>"

    html += "</tr>"

    for _, row in dff.iterrows():
        html += "<tr>"
        for val in row:
            html += f"<td style='padding:6px;border:1px solid #ddd'>{val}</td>"
        html += "</tr>"

    html += "</table>"

    st.text_area("Copier-coller dans Outlook", html, height=300)

    # ==============================
    # IA SIMPLE
    # ==============================
    st.subheader("🧠 Analyse automatique")

    total = len(dff)
    av = dff["Avancement"].mean()
    retard = len(dff[dff["Statut"]=="En retard"])

    st.write(f"""
- {total} projets suivis  
- Avancement moyen : {av:.0f}%  
- {retard} projets en retard  
""")