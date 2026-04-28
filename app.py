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
# 🔧 CLEAN RESPONSABLE
# ==============================
def clean_responsable(x):
    if pd.isna(x):
        return "Non défini"
    return str(x).split(";")[0].strip()

# ==============================
# 🔧 AVANCEMENT
# ==============================
def extract_avancement(row):
    txt = str(row.get("Progression", "")).lower()

    if "termin" in txt:
        return 100
    elif "cours" in txt:
        return 50
    elif "non" in txt:
        return 0

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
    # 🧹 CLEAN
    # ==============================
    df["Projet"] = df[col_projet]
    df["Responsable"] = df[col_resp].apply(clean_responsable)
    df["Description brute"] = df[col_desc].fillna("")

    # parsing texte
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
        resp = st.selectbox(
            "Responsable",
            ["Tous"] + sorted(df_clean["Responsable"].unique())
        )

    with col2:
        proj = st.selectbox(
            "Projet",
            ["Tous"] + sorted(df_clean["Projet"].dropna().unique())
        )

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
    c2.metric("Avancement moyen", f"{dff['Avancement'].mean():.0f}%" if len(dff) else "0%")
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
    # 📋 TABLEAU STYLE EXCEL
    # ==============================
    st.subheader("📋 Tableau structuré")

    def style_table(df):
        return df.style \
            .set_properties(**{
                "border": "1px solid #ddd",
                "padding": "6px"
            }) \
            .set_table_styles([
                {
                    "selector": "th",
                    "props": [
                        ("background-color", "#0A2463"),
                        ("color", "white"),
                        ("padding", "8px")
                    ]
                }
            ])

    st.write(style_table(dff))

    # ==============================
    # 📧 TABLEAU OUTLOOK
    # ==============================
    st.subheader("📧 Tableau pour Outlook")

    def clean_text(val):
        return html.escape(str(val)).replace("\n", "<br>")

    html_table = "<table style='border-collapse:collapse;font-family:Calibri;width:100%'>"

    html_table += "<tr style='background:#0A2463;color:white'>"
    for col in dff.columns:
        html_table += f"<th style='border:1px solid #ddd;padding:8px'>{col}</th>"
    html_table += "</tr>"

    for _, row in dff.iterrows():
        html_table += "<tr>"
        for val in row:
            html_table += f"<td style='border:1px solid #ddd;padding:6px'>{clean_text(val)}</td>"
        html_table += "</tr>"

    html_table += "</table>"

    st.code(html_table, language="html")

    # ==============================
    # 🧠 ANALYSE SIMPLE
    # ==============================
    st.subheader("🧠 Analyse automatique")

    if len(dff):
        av = dff["Avancement"].mean()
        retard = len(dff[dff["Statut"] == "En retard"])

        st.write(f"""
✔ {len(dff)} projets  
✔ Avancement moyen : {av:.0f}%  
✔ {retard} en retard  

➡️ Lecture :
- < 40% → risque  
- > 70% → OK  
""")