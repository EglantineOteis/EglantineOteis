# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import re

st.set_page_config(layout="wide")
st.title("📊 Suivi des projets")

file = st.file_uploader("📂 Importer Excel brut", type=["xlsx"])

# ==============================
# FONCTIONS INTELLIGENTES
# ==============================

def clean_responsable(x):
    if pd.isna(x):
        return "Non défini"
    # si plusieurs noms → prendre le premier
    return str(x).split(";")[0].strip()

def extract_avancement(row):
    # 1. depuis colonne progression
    prog = str(row.get("Progression", "")).lower()

    if "termin" in prog:
        return 100
    elif "cours" in prog:
        return 50
    elif "non" in prog:
        return 0

    # 2. fallback : chercher % dans description
    desc = str(row.get("Description", ""))
    match = re.search(r"(\d+)\s*%", desc)
    if match:
        return int(match.group(1))

    return 0

def get_statut(x):
    if x == 0:
        return "Non démarré"
    elif x == 100:
        return "Terminé"
    elif x < 40:
        return "En retard"
    else:
        return "En cours"

# ==============================
# TRAITEMENT
# ==============================

if file:

    df = pd.read_excel(file)

    # NORMALISATION colonnes
    df.columns = df.columns.str.strip()

    # mapping intelligent
    col_projet = [c for c in df.columns if "tâche" in c.lower() or "projet" in c.lower()][0]
    col_resp = [c for c in df.columns if "créé" in c.lower() or "responsable" in c.lower()][0]
    col_prog = [c for c in df.columns if "progress" in c.lower()][0]
    col_desc = [c for c in df.columns if "descript" in c.lower()][0]

    # ==============================
    # CLEAN
    # ==============================
    df["Projet"] = df[col_projet]
    df["Responsable"] = df[col_resp].apply(clean_responsable)
    df["Description"] = df[col_desc].fillna("")

    # AVANCEMENT INTELLIGENT
    df["Avancement"] = df.apply(extract_avancement, axis=1)

    df["Statut"] = df["Avancement"].apply(get_statut)

    df_clean = df[[
        "Projet",
        "Responsable",
        "Avancement",
        "Statut",
        "Description"
    ]]

    # ==============================
    # FILTRES
    # ==============================
    st.subheader("🎯 Filtres")

    col1, col2 = st.columns(2)

    with col1:
        resp_list = sorted(df_clean["Responsable"].unique())
        resp = st.selectbox("Responsable", ["Tous"] + resp_list)

    with col2:
        proj_list = sorted(df_clean["Projet"].dropna().unique())
        proj = st.selectbox("Projet", ["Tous"] + proj_list)

    dff = df_clean.copy()

    if resp != "Tous":
        dff = dff[dff["Responsable"] == resp]

    if proj != "Tous":
        dff = dff[dff["Projet"] == proj]

    # ==============================
    # KPI
    # ==============================
    st.subheader("📈 Indicateurs")

    c1, c2, c3 = st.columns(3)

    c1.metric("Projets", len(dff))
    c2.metric("Avancement moyen", f"{dff['Avancement'].mean():.0f}%")
    c3.metric("En retard", len(dff[dff["Statut"] == "En retard"]))

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
    # TABLEAU OUTLOOK PROPRE
    # ==============================
    st.subheader("📧 Tableau pour Outlook")

    html = "<table style='border-collapse:collapse;font-family:Calibri;width:100%'>"
    html += "<tr style='background:#0A2463;color:white'>"

    for col in dff.columns:
        html += f"<th style='border:1px solid #ddd;padding:8px'>{col}</th>"

    html += "</tr>"

    for _, row in dff.iterrows():
        html += "<tr>"
        for val in row:
            html += f"<td style='border:1px solid #ddd;padding:6px'>{val}</td>"
        html += "</tr>"

    html += "</table>"

    st.code(html, language="html")

    # ==============================
    # IA SIMPLE (FIABLE)
    # ==============================
    st.subheader("🧠 Analyse automatique")

    total = len(dff)
    av = dff["Avancement"].mean()
    retard = len(dff[dff["Statut"] == "En retard"])

    st.write(f"""
✔ {total} projets analysés  
✔ Avancement moyen : {av:.0f}%  
✔ {retard} projets en retard  

➡️ Lecture rapide :
- Si < 50% → portefeuille en difficulté  
- Si > 70% → bonne dynamique  
""")