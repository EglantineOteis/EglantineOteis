# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import re
import html

st.set_page_config(layout="wide")

st.title("📊 Suivi des projets")

file = st.file_uploader("📂 Importer Excel brut", type=["xlsx"])

# ==============================
# 🔧 PARSING TEXTE PROPRE
# ==============================
def parse_description(text):
    data = {
        "description": "",
        "admin": "",
        "intervenants": "",
        "remarques": ""
    }

    if pd.isna(text):
        return data

    text = str(text)

    patterns = {
        "description": r"Descriptif\s*:\s*(.*?)(?=Administration|Liste|Remarque|$)",
        "admin": r"Administration.*?:\s*(.*?)(?=Liste|Remarque|$)",
        "intervenants": r"Liste des intervenant.*?:\s*(.*?)(?=Remarque|$)",
        "remarques": r"Remarque.*?:\s*(.*)"
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            data[key] = match.group(1).strip()

    return data

# ==============================
# 🔧 NETTOYAGE RESPONSABLE
# ==============================
def clean_responsable(x):
    if pd.isna(x):
        return "Non défini"
    return str(x).split(";")[0].strip()

# ==============================
# 🔧 AVANCEMENT INTELLIGENT
# ==============================
def extract_avancement(row):

    val = str(row.get("Progression", "")).lower()

    if "termin" in val:
        return 100
    elif "cours" in val:
        return 50
    elif "non" in val:
        return 0

    # fallback %
    desc = str(row.get("Description brute", ""))
    match = re.search(r"(\d+)\s*%", desc)
    if match:
        return int(match.group(1))

    return 0

# ==============================
# 🔧 STATUT
# ==============================
def statut(x):
    if x == 0:
        return "Non démarré"
    elif x == 100:
        return "Terminé"
    elif x < 40:
        return "En retard"
    else:
        return "En cours"

# ==============================
# 🚀 TRAITEMENT
# ==============================
if file:

    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()

    # 🔍 détection colonnes
    col_projet = [c for c in df.columns if "tâche" in c.lower() or "projet" in c.lower()][0]
    col_resp = [c for c in df.columns if "créé" in c.lower() or "responsable" in c.lower()][0]
    col_prog = [c for c in df.columns if "progress" in c.lower()][0]
    col_desc = [c for c in df.columns if "descript" in c.lower()][0]

    # ==============================
    # 🧹 CLEAN DATA
    # ==============================
    df["Projet"] = df[col_projet]
    df["Responsable"] = df[col_resp].apply(clean_responsable)
    df["Description brute"] = df[col_desc].fillna("")

    # parsing
    parsed = df["Description brute"].apply(parse_description)

    df["Description"] = parsed.apply(lambda x: x["description"])
    df["Administration"] = parsed.apply(lambda x: x["admin"])
    df["Intervenants"] = parsed.apply(lambda x: x["intervenants"])
    df["Remarques"] = parsed.apply(lambda x: x["remarques"])

    # avancement
    df["Avancement"] = df.apply(extract_avancement, axis=1)

    df["Statut"] = df["Avancement"].apply(statut)

    df_clean = df[[
        "Projet",
        "Responsable",
        "Avancement",
        "Statut",
        "Description",
        "Remarques"
    ]]

    # ==============================
    # 🎯 FILTRES
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
    # 📈 KPI
    # ==============================
    st.subheader("📈 Indicateurs")

    c1, c2, c3 = st.columns(3)

    c1.metric("Projets", len(dff))

    if len(dff) > 0:
        c2.metric("Avancement moyen", f"{dff['Avancement'].mean():.0f}%")
    else:
        c2.metric("Avancement moyen", "0%")

    c3.metric("En retard", len(dff[dff["Statut"] == "En retard"]))

    # ==============================
    # 📊 GRAPHIQUES
    # ==============================
    st.subheader("📊 Graphiques")

    if len(dff) > 0:

        fig1 = px.pie(dff, names="Statut", hole=0.5)

        fig2 = px.bar(
            dff.groupby("Responsable")["Projet"].count().reset_index(),
            x="Responsable",
            y="Projet",
            title="Charge par responsable"
        )

        st.plotly_chart(fig1, use_container_width=True)
        st.plotly_chart(fig2, use_container_width=True)

    # ==============================
    # 📋 TABLEAU
    # ==============================
    st.subheader("📋 Tableau structuré")
    st.dataframe(dff, use_container_width=True)

    # ==============================
    # 📧 TABLEAU OUTLOOK PROPRE
    # ==============================
    st.subheader("📧 Tableau pour Outlook")

    html_table = "<table style='border-collapse:collapse;font-family:Calibri;width:100%'>"

    html_table += "<tr style='background:#0A2463;color:white'>"
    for col in dff.columns:
        html_table += f"<th style='border:1px solid #ddd;padding:8px'>{col}</th>"
    html_table += "</tr>"

    for _, row in dff.iterrows():
        html_table += "<tr>"
        for val in row:
            safe_val = html.escape(str(val))
            html_table += f"<td style='border:1px solid #ddd;padding:6px'>{safe_val}</td>"
        html_table += "</tr>"

    html_table += "</table>"

    st.code(html_table, language="html")

    # ==============================
    # 🧠 ANALYSE SIMPLE
    # ==============================
    st.subheader("🧠 Analyse automatique")

    if len(dff) > 0:
        av = dff["Avancement"].mean()
        retard = len(dff[dff["Statut"] == "En retard"])

        st.write(f"""
✔ {len(dff)} projets analysés  
✔ Avancement moyen : {av:.0f}%  
✔ {retard} projets en retard  

➡️ Lecture rapide :
- < 40% → alerte  
- > 70% → bonne dynamique  
""")