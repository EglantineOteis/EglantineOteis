# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import urllib.parse
import os

st.set_page_config(layout="wide")

# =========================
# 🔍 DETECTION INTELLIGENTE (ANTI-CRASH)
# =========================
def smart_detect(df):

    mapping = {
        "projet": None,
        "responsable": None,
        "avancement": None,
        "description": None
    }

    for col in df.columns:
        c = col.lower()

        if any(k in c for k in ["tâche", "tache", "nom"]):
            mapping["projet"] = col

        elif any(k in c for k in ["attribué", "responsable", "assign", "owner"]):
            mapping["responsable"] = col

        elif any(k in c for k in ["progress", "avancement", "%"]):
            mapping["avancement"] = col

        elif any(k in c for k in ["description", "comment"]):
            mapping["description"] = col

    return mapping


# =========================
# 🧠 PARSE TEXTE
# =========================
def parse_desc(txt):

    desc, rem = "", ""

    if not isinstance(txt, str):
        return desc, rem

    for l in txt.split("\n"):
        l_low = l.lower()

        if "descriptif" in l_low:
            desc = l.split(":",1)[-1].strip()

        if "remarque" in l_low:
            rem = l.split(":",1)[-1].strip()

    return desc, rem


# =========================
# 📊 HTML TABLE PRO
# =========================
def generate_html(df):

    html = "<table style='border-collapse:collapse;font-family:Calibri;width:100%'>"

    # header
    html += "<tr style='background:#0A2463;color:white'>"
    for col in df.columns:
        html += f"<th style='padding:10px;border:1px solid #ddd'>{col}</th>"
    html += "</tr>"

    # rows
    for i, (_, row) in enumerate(df.iterrows()):
        bg = "#f4f6fa" if i % 2 else "white"
        html += f"<tr style='background:{bg}'>"

        for col in df.columns:
            val = str(row[col]).replace("\n", "<br>")

            if col == "Avancement":
                html += f"<td style='padding:8px;border:1px solid #ddd;color:#E85D04;font-weight:bold'>{val}%</td>"
            else:
                html += f"<td style='padding:8px;border:1px solid #ddd'>{val}</td>"

        html += "</tr>"

    html += "</table>"
    return html


# =========================
# 📧 MAIL OUTLOOK
# =========================
def mail_link(df):

    html_table = generate_html(df)

    body = f"""
Bonjour,

Voici le suivi des projets :

{html_table}

Cordialement
"""

    return f"mailto:?subject=Suivi projets&body={urllib.parse.quote(body)}"


# =========================
# 🤖 IA SIMPLIFIÉE (safe)
# =========================
def ai_summary(df):

    total = len(df)
    av = df["Avancement"].mean()
    retard = len(df[df["Statut"] == "En retard"])

    return f"""
📊 Synthèse :

- {total} projets
- Avancement moyen : {av:.0f}%
- {retard} en retard

⚠️ Priorité :
Suivre les projets < 40%
"""


# =========================
# UI
# =========================
st.title("📊 Suivi des projets intelligent")

file = st.file_uploader("Importer Excel", type=["xlsx"])

if file:

    df_raw = pd.read_excel(file)
    df_raw.columns = df_raw.columns.str.lower().str.strip()

    st.write("Colonnes détectées :", df_raw.columns.tolist())

    # 🔍 IA détection
    mapping = smart_detect(df_raw)

    st.write("Mapping IA :", mapping)

    # 🚨 sécurité anti crash
    if None in mapping.values():
        st.error("❌ Impossible de détecter automatiquement les colonnes")
        st.stop()

    # extraction
    projet_col = mapping["projet"]
    resp_col = mapping["responsable"]
    prog_col = mapping["avancement"]
    desc_col = mapping["description"]

    # nettoyage
    df = pd.DataFrame()

    df["Projet"] = df_raw[projet_col]
    df["Responsable"] = df_raw[resp_col].astype(str).str.split(";").str[0]
    df["Avancement"] = pd.to_numeric(df_raw[prog_col], errors="coerce").fillna(0)

    parsed = df_raw[desc_col].apply(parse_desc)

    df["Description"] = parsed.apply(lambda x: x[0])
    df["Remarques"] = parsed.apply(lambda x: x[1])

    # statut
    def statut(x):
        if x == 0: return "Non démarré"
        elif x == 100: return "Terminé"
        elif x < 40: return "En retard"
        else: return "En cours"

    df["Statut"] = df["Avancement"].apply(statut)

    # =========================
    # FILTRES
    # =========================
    col1, col2 = st.columns(2)

    with col1:
        resp = st.selectbox("Responsable", ["Tous"] + sorted(df["Responsable"].dropna().unique()))

    with col2:
        proj = st.selectbox("Projet", ["Tous"] + sorted(df["Projet"].dropna().unique()))

    df_f = df.copy()

    if resp != "Tous":
        df_f = df_f[df_f["Responsable"] == resp]

    if proj != "Tous":
        df_f = df_f[df_f["Projet"] == proj]

    # =========================
    # KPI
    # =========================
    c1, c2, c3 = st.columns(3)

    c1.metric("Projets", len(df_f))
    c2.metric("Avancement", f"{df_f['Avancement'].mean():.0f}%")
    c3.metric("Retard", len(df_f[df_f["Statut"] == "En retard"]))

    # =========================
    # GRAPHIQUES
    # =========================
    st.subheader("Graphiques")

    st.plotly_chart(px.pie(df_f, names="Statut", hole=0.5), use_container_width=True)
    st.plotly_chart(px.bar(df_f.groupby("Responsable").size().reset_index(name="nb"),
                           x="Responsable", y="nb"), use_container_width=True)

    # =========================
    # TABLEAU
    # =========================
    st.subheader("Tableau")
    st.dataframe(df_f, use_container_width=True)

    # =========================
    # TABLEAU OUTLOOK
    # =========================
    st.subheader("Tableau Outlook")

    html = generate_html(df_f)
    st.markdown(html, unsafe_allow_html=True)

    # =========================
    # BOUTON MAIL
    # =========================
    link = mail_link(df_f)

    st.markdown(f"""
    <a href="{link}" target="_blank">
    <button style="background:#E85D04;color:white;padding:10px;border:none;border-radius:5px">
    📧 Ouvrir mail Outlook
    </button>
    </a>
    """, unsafe_allow_html=True)

    # =========================
    # IA
    # =========================
    st.subheader("🤖 Analyse automatique")
    st.text(ai_summary(df_f))