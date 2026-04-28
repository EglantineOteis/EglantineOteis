# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import re

st.set_page_config(layout="wide")

# ==============================
# NORMALISATION COLONNES
# ==============================
def normalize_columns(df):
    df.columns = df.columns.str.lower().str.strip()
    mapping = {}

    for c in df.columns:
        if "tâche" in c or "task" in c:
            mapping[c] = "projet"
        elif "attrib" in c:
            mapping[c] = "responsable"
        elif "compartiment" in c:
            mapping[c] = "equipe"
        elif "progress" in c:
            mapping[c] = "progression"
        elif "description" in c:
            mapping[c] = "description_brut"

    df = df.rename(columns=mapping)
    return df

# ==============================
# AVANCEMENT INTELLIGENT
# ==============================
def clean_progress(x):
    if pd.isna(x):
        return 0
    x = str(x).lower()

    if "non" in x:
        return 0
    if "cours" in x:
        return 50
    if "term" in x:
        return 100

    return 0

# ==============================
# PARSE DESCRIPTION PROPRE
# ==============================
def parse_description(txt):
    if not isinstance(txt, str):
        return "", ""

    desc = ""
    rem = ""

    lines = txt.split("\n")

    for l in lines:
        l_low = l.lower()

        if "descriptif" in l_low:
            desc = l.split(":")[-1].strip()

        elif "remarque" in l_low:
            rem = l.split(":")[-1].strip()

    return desc, rem

# ==============================
# RESPONSABLE INTELLIGENT
# ==============================
def clean_responsable(row):

    # priorité 1 : attribué
    if "responsable" in row and pd.notna(row["responsable"]):
        return row["responsable"]

    # fallback : équipe
    if "equipe" in row:
        return row["equipe"]

    return "Non défini"

# ==============================
# TRANSFORMATION GLOBALE
# ==============================
def transform(df):

    df = normalize_columns(df)

    # sécurité colonnes
    for col in ["projet", "progression", "description_brut"]:
        if col not in df:
            df[col] = ""

    # nettoyage
    df["avancement"] = df["progression"].apply(clean_progress)
    df["responsable"] = df.apply(clean_responsable, axis=1)

    parsed = df["description_brut"].apply(parse_description)

    df["description"] = parsed.apply(lambda x: x[0])
    df["remarques"] = parsed.apply(lambda x: x[1])

    # statut
    def statut(x):
        if x == 0:
            return "Non démarré"
        elif x == 100:
            return "Terminé"
        else:
            return "En cours"

    df["statut"] = df["avancement"].apply(statut)

    df_clean = df[[
        "projet",
        "responsable",
        "avancement",
        "statut",
        "description",
        "remarques"
    ]]

    return df_clean

# ==============================
# TABLE HTML PRO (OUTLOOK)
# ==============================
def generate_html_table(df):

    html = """
    <table style="border-collapse:collapse;font-family:Calibri;width:100%">
    <tr style="background:#0A2463;color:white">
    <th style="padding:8px;border:1px solid #ddd">Projet</th>
    <th style="padding:8px;border:1px solid #ddd">Responsable</th>
    <th style="padding:8px;border:1px solid #ddd">Avancement</th>
    <th style="padding:8px;border:1px solid #ddd">Statut</th>
    <th style="padding:8px;border:1px solid #ddd">Description</th>
    </tr>
    """

    for _, r in df.iterrows():
        html += f"""
        <tr>
        <td style="padding:6px;border:1px solid #ddd">{r['projet']}</td>
        <td style="padding:6px;border:1px solid #ddd">{r['responsable']}</td>
        <td style="padding:6px;border:1px solid #ddd">{r['avancement']}%</td>
        <td style="padding:6px;border:1px solid #ddd">{r['statut']}</td>
        <td style="padding:6px;border:1px solid #ddd">{r['description']}</td>
        </tr>
        """

    html += "</table>"
    return html

# ==============================
# UI
# ==============================
st.title("📊 Suivi des projets intelligent")

file = st.file_uploader("📂 Importer Excel", type=["xlsx"])

if file:

    df_raw = pd.read_excel(file)
    df = transform(df_raw)

    # ==============================
    # FILTRES
    # ==============================
    col1, col2 = st.columns(2)

    responsables = ["Tous"] + sorted(df["responsable"].dropna().unique())
    projets = ["Tous"] + sorted(df["projet"].dropna().unique())

    resp_filter = col1.selectbox("Filtrer par responsable", responsables)
    proj_filter = col2.selectbox("Filtrer par projet", projets)

    df_filtered = df.copy()

    if resp_filter != "Tous":
        df_filtered = df_filtered[df_filtered["responsable"] == resp_filter]

    if proj_filter != "Tous":
        df_filtered = df_filtered[df_filtered["projet"] == proj_filter]

    # ==============================
    # KPI
    # ==============================
    st.subheader("📈 Indicateurs")

    c1, c2, c3 = st.columns(3)

    c1.metric("Projets", len(df_filtered))
    c2.metric("Avancement moyen", f"{df_filtered['avancement'].mean():.0f}%")
    c3.metric("En retard", len(df_filtered[df_filtered["statut"] == "En retard"]))

    # ==============================
    # GRAPHIQUES
    # ==============================
    st.subheader("📊 Graphiques")

    fig1 = px.pie(df_filtered, names="statut")
    fig2 = px.bar(df_filtered, x="responsable", y="avancement")

    st.plotly_chart(fig1, use_container_width=True)
    st.plotly_chart(fig2, use_container_width=True)

    # ==============================
    # TABLEAU PROPRE
    # ==============================
    st.subheader("📋 Tableau structuré")
    st.dataframe(df_filtered, use_container_width=True)

    # ==============================
    # TABLE MAIL PRO
    # ==============================
    st.subheader("📧 Tableau Outlook (copier-coller)")

    html_table = generate_html_table(df_filtered)

    st.code(html_table, language="html")

    st.markdown("👉 Copier ce code et coller directement dans Outlook")
