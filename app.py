# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Suivi projets", layout="wide")

# ==============================
# FONCTIONS
# ==============================

def detect_col(df, mots):
    for mot in mots:
        for col in df.columns:
            if mot in col:
                return col
    return None

def parse(txt):
    d = {"desc":"","rem":"","admin":""}
    if not isinstance(txt,str):
        return d

    for l in txt.split("\n"):
        l2 = l.lower()
        if "descriptif" in l2: d["desc"] = l.split(":")[-1].strip()
        if "remarque" in l2: d["rem"] = l.split(":")[-1].strip()
        if "administration" in l2: d["admin"] = l.split(":")[-1].strip()

    return d

def statut(x):
    if pd.isna(x): return "Inconnu"
    if x == 0: return "Non démarré"
    elif x == 100: return "Terminé"
    elif x < 40: return "En retard"
    else: return "En cours"

def generate_ai_summary(df):
    total = len(df)
    av = df["avancement"].mean()
    retard = len(df[df["statut"]=="En retard"])

    top = df.sort_values("avancement").head(3)

    txt = f"""
Synthèse des projets :

- {total} projets
- Avancement moyen : {av:.1f} %
- Projets en retard : {retard}

Points critiques :
"""

    for _, r in top.iterrows():
        txt += f"- {r.iloc[0]} ({r['avancement']}%)\n"

    return txt

def generate_html_table(df):
    html = """
    <table style="border-collapse:collapse;font-family:Calibri;font-size:11pt;width:100%">
    <tr style="background-color:#0A2463;color:white">
    """

    for col in df.columns:
        html += f'<th style="border:1px solid #ddd;padding:8px">{col}</th>'

    html += "</tr>"

    for _, row in df.iterrows():
        html += "<tr>"
        for val in row:
            html += f'<td style="border:1px solid #ddd;padding:6px">{val}</td>'
        html += "</tr>"

    html += "</table>"
    return html


# ==============================
# UI
# ==============================

st.title("📊 Suivi des projets")

file = st.file_uploader("📂 Importer Excel brut", type=["xlsx"])

if file:

    df = pd.read_excel(file)
    df.columns = df.columns.str.strip().str.lower()

    # ==============================
    # DETECTION COLONNES
    # ==============================
    col_p = detect_col(df, ["projet", "tache", "nom"])
    col_r = detect_col(df, ["responsable", "attrib", "assign"])
    col_a = detect_col(df, ["avancement", "progress", "%"])
    col_d = detect_col(df, ["descript", "comment", "detail"])

    if None in [col_p, col_r, col_a]:
        st.error("❌ Colonnes non reconnues")
        st.write(df.columns)
        st.stop()

    # ==============================
    # NETTOYAGE
    # ==============================
    df[col_a] = (
        df[col_a].astype(str)
        .str.replace("%","")
        .str.replace(",",".")
    )

    df[col_a] = pd.to_numeric(df[col_a], errors="coerce")

    df["statut"] = df[col_a].apply(statut)

    if col_d:
        parsed = df[col_d].apply(parse)
        df["description"] = parsed.apply(lambda x: x["desc"])
        df["remarques"] = parsed.apply(lambda x: x["rem"])
        df["admin"] = parsed.apply(lambda x: x["admin"])
    else:
        df["description"] = ""
        df["remarques"] = ""
        df["admin"] = ""

    df_clean = df[[col_p, col_r, col_a, "statut", "description", "remarques"]]
    df_clean.columns = ["Projet", "Responsable", "Avancement", "Statut", "Description", "Remarques"]

    # ==============================
    # FILTRE RESPONSABLE
    # ==============================
    st.subheader("🎯 Filtre")

    chefs = df_clean["Responsable"].dropna().unique()
    chef = st.selectbox("Choisir un responsable", ["Tous"] + list(chefs))

    if chef != "Tous":
        df_filtered = df_clean[df_clean["Responsable"] == chef]
    else:
        df_filtered = df_clean

    # ==============================
    # KPI
    # ==============================
    st.subheader("📈 Indicateurs")

    col1, col2, col3 = st.columns(3)

    col1.metric("Projets", len(df_filtered))
    col2.metric("Avancement moyen", f"{df_filtered['Avancement'].mean():.1f}%")
    col3.metric("En retard", len(df_filtered[df_filtered["Statut"]=="En retard"]))

    # ==============================
    # GRAPHIQUES
    # ==============================
    st.subheader("📊 Graphiques")

    fig1 = px.pie(df_filtered, names="Statut", hole=0.5)
    fig2 = px.histogram(df_filtered, x="Avancement", nbins=10)

    fig3 = px.bar(
        df_filtered.groupby("Responsable")["Projet"].count().reset_index(),
        x="Responsable", y="Projet",
        title="Charge par responsable"
    )

    st.plotly_chart(fig1, use_container_width=True)
    st.plotly_chart(fig2, use_container_width=True)
    st.plotly_chart(fig3, use_container_width=True)

    # ==============================
    # TABLEAU STRUCTURE
    # ==============================
    st.subheader("📋 Tableau structuré")
    st.dataframe(df_filtered, use_container_width=True)

    # ==============================
    # TABLEAU OUTLOOK
    # ==============================
    st.subheader("📧 Tableau pour Outlook (copier-coller)")

    html = generate_html_table(df_filtered)
    st.code(html, language="html")

    # ==============================
    # IA
    # ==============================
    st.subheader("🧠 Analyse automatique")

    st.text(generate_ai_summary(df_filtered))