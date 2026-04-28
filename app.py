# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import urllib.parse

st.set_page_config(layout="wide")

# =========================
# EXTRACTION DES INFOS
# =========================
def parse_description(txt):
    desc, admin, inter, rem, av = "", "", "", "", ""

    if not isinstance(txt, str):
        return desc, admin, inter, rem, av

    for l in txt.split("\n"):
        l_low = l.lower()

        if "descriptif" in l_low:
            desc = l.split(":",1)[-1].strip()
        elif "administration" in l_low:
            admin = l.split(":",1)[-1].strip()
        elif "intervenant" in l_low:
            inter = l.split(":",1)[-1].strip()
        elif "remarque" in l_low:
            rem = l.split(":",1)[-1].strip()
        elif "avancement" in l_low:
            av = l.split(":",1)[-1].strip()

    return desc, admin, inter, rem, av

# =========================
# HTML TABLE PRO (Outlook)
# =========================
def generate_html_table(df):

    html = """
    <table style="border-collapse:collapse;font-family:Calibri;width:100%">
    """

    # HEADER
    html += "<tr style='background:#0A2463;color:white;'>"
    for col in df.columns:
        html += f"<th style='padding:10px;border:1px solid #ddd'>{col}</th>"
    html += "</tr>"

    # LIGNES
    for i, (_, row) in enumerate(df.iterrows()):
        bg = "#f4f6fa" if i % 2 else "white"
        html += f"<tr style='background:{bg}'>"

        for col in df.columns:
            val = str(row[col]).replace("\n", "<br>")

            if col == "Avancement":
                html += f"""
                <td style='padding:8px;border:1px solid #ddd;
                color:#E85D04;font-weight:bold;text-align:center'>
                {val}%
                </td>
                """
            else:
                html += f"<td style='padding:8px;border:1px solid #ddd'>{val}</td>"

        html += "</tr>"

    html += "</table>"
    return html

# =========================
# MAIL OUTLOOK
# =========================
def generate_mail_link(df):

    html_table = generate_html_table(df)

    body = f"""
    <html>
    <body>
    <p>Bonjour,</p>
    <p>Voici le suivi des projets :</p>
    {html_table}
    <br>
    <p>Cordialement</p>
    </body>
    </html>
    """

    subject = "Suivi des projets"

    return f"mailto:?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(body)}"

# =========================
# UI
# =========================
st.title("📊 Suivi des projets intelligent")

file = st.file_uploader("📂 Importer Excel", type=["xlsx"])

if file:

    df_raw = pd.read_excel(file)
    df_raw.columns = df_raw.columns.str.lower().str.strip()

    st.write("Colonnes détectées :", df_raw.columns.tolist())

    # =========================
    # COLONNES
    # =========================
    projet_col = [c for c in df_raw.columns if "tâche" in c or "tache" in c][0]
    resp_col = [c for c in df_raw.columns if "responsable" in c][0]
    desc_col = [c for c in df_raw.columns if "description" in c][0]
    prog_col = [c for c in df_raw.columns if "progress" in c][0]

    # =========================
    # PARSE
    # =========================
    parsed = df_raw[desc_col].apply(parse_description)

    df = pd.DataFrame()
    df["Projet"] = df_raw[projet_col]
    df["Responsable"] = df_raw[resp_col]
    df["Avancement"] = pd.to_numeric(df_raw[prog_col], errors="coerce").fillna(0)

    df["Description"] = parsed.apply(lambda x: x[0])
    df["Remarques"] = parsed.apply(lambda x: x[3])

    # STATUT
    def statut(x):
        if x == 0:
            return "Non démarré"
        elif x == 100:
            return "Terminé"
        elif x < 40:
            return "En retard"
        else:
            return "En cours"

    df["Statut"] = df["Avancement"].apply(statut)

    # =========================
    # FILTRES
    # =========================
    col1, col2 = st.columns(2)

    with col1:
        resp = st.selectbox("👤 Chef de projet", ["Tous"] + sorted(df["Responsable"].dropna().unique()))

    with col2:
        proj = st.selectbox("📁 Projet", ["Tous"] + sorted(df["Projet"].dropna().unique()))

    df_filtered = df.copy()

    if resp != "Tous":
        df_filtered = df_filtered[df_filtered["Responsable"] == resp]

    if proj != "Tous":
        df_filtered = df_filtered[df_filtered["Projet"] == proj]

    # =========================
    # KPI
    # =========================
    c1, c2, c3 = st.columns(3)

    c1.metric("Projets", len(df_filtered))
    c2.metric("Avancement moyen", f"{df_filtered['Avancement'].mean():.0f}%")
    c3.metric("En retard", len(df_filtered[df_filtered["Statut"] == "En retard"]))

    # =========================
    # GRAPHIQUES
    # =========================
    st.subheader("📊 Graphiques")

    fig1 = px.pie(df_filtered, names="Statut", hole=0.5)
    st.plotly_chart(fig1, use_container_width=True)

    fig2 = px.bar(df_filtered.groupby("Responsable").size().reset_index(name="nb"),
                  x="Responsable", y="nb")
    st.plotly_chart(fig2, use_container_width=True)

    # =========================
    # TABLEAU
    # =========================
    st.subheader("📋 Tableau structuré")
    st.dataframe(df_filtered, use_container_width=True)

    # =========================
    # TABLEAU OUTLOOK
    # =========================
    st.subheader("📧 Tableau Outlook")

    html_table = generate_html_table(df_filtered)

    st.markdown(html_table, unsafe_allow_html=True)

    # =========================
    # BOUTON MAIL
    # =========================
    link = generate_mail_link(df_filtered)

    st.markdown(f"""
    <a href="{link}" target="_blank">
        <button style="
            background:#E85D04;
            color:white;
            padding:12px;
            border:none;
            border-radius:6px;
            font-size:16px;
            cursor:pointer;">
            📧 Ouvrir le mail Outlook
        </button>
    </a>
    """, unsafe_allow_html=True)