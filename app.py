# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide")

# =========================
# DETECTION COLONNES
# =========================
def detect_columns(df):

    mapping = {"projet":None,"responsable":None,"avancement":None,"description":None}

    for col in df.columns:
        c = col.lower()

        if "tâche" in c:
            mapping["projet"] = col
        elif "responsable" in c:
            mapping["responsable"] = col
        elif "attribué" in c and mapping["responsable"] is None:
            mapping["responsable"] = col
        elif "progress" in c:
            mapping["avancement"] = col
        elif "description" in c:
            mapping["description"] = col

    return mapping


# =========================
# PARSE DESCRIPTION
# =========================
def parse_description(txt):

    desc, rem, av = "", "", None

    if not isinstance(txt, str):
        return desc, rem, av

    try:
        if "Descriptif" in txt:
            desc = txt.split("Descriptif :",1)[1].split("Administration")[0].strip()
    except: pass

    try:
        if "Remarques" in txt:
            rem = txt.split("Remarques :",1)[1].strip()
    except: pass

    try:
        if "Avancement" in txt:
            val = txt.split("Avancement :",1)[1].split("\n")[0]
            av = float(val.replace("%","").strip())
    except: pass

    return desc, rem, av


# =========================
# RESPONSABLE
# =========================
def clean_responsable(x):

    if pd.isna(x):
        return "Non défini"

    x = str(x)

    if ";" in x:
        return x.split(";")[0].strip()

    if "," in x:
        return x.split(",")[0].strip()

    return x.strip()


# =========================
# HTML TABLE EXCEL STYLE
# =========================
def generate_html(df):

    html = """
    <table style="border-collapse:collapse;font-family:Calibri;font-size:13px;width:100%">
    """

    # HEADER
    html += "<tr style='background:#0A2463;color:white'>"
    for col in df.columns:
        html += f"<th style='padding:8px;border:1px solid #d9d9d9'>{col}</th>"
    html += "</tr>"

    # LIGNES
    for i, (_, row) in enumerate(df.iterrows()):
        bg = "#ffffff" if i % 2 == 0 else "#f4f6fa"
        html += f"<tr style='background:{bg}'>"

        for col in df.columns:

            val = str(row[col]).replace("\n","<br>")

            if col == "Avancement":

                if row["Statut"] == "Début":
                    color = "#d62828"
                elif row["Statut"] == "Milieu":
                    color = "#f77f00"
                else:
                    color = "#2a9d8f"

                html += f"<td style='padding:6px;border:1px solid #ddd;color:{color};font-weight:bold;text-align:center'>{val}%</td>"

            else:
                html += f"<td style='padding:6px;border:1px solid #ddd'>{val}</td>"

        html += "</tr>"

    html += "</table>"

    return html


# =========================
# UI
# =========================
st.title("📊 Suivi des projets")

file = st.file_uploader("📂 Importer Excel", type=["xlsx"])

if file:

    df_raw = pd.read_excel(file)
    df_raw.columns = df_raw.columns.str.lower().str.strip()

    mapping = detect_columns(df_raw)

    if mapping["projet"] is None:
        st.error("❌ colonne 'Nom de tâche' introuvable")
        st.stop()

    df = pd.DataFrame()

    df["Projet"] = df_raw[mapping["projet"]]

    df["Responsable"] = df_raw[mapping["responsable"]].apply(clean_responsable)

    df["Avancement"] = pd.to_numeric(
        df_raw[mapping["avancement"]],
        errors="coerce"
    )

    parsed = df_raw[mapping["description"]].apply(parse_description)

    df["Description"] = parsed.apply(lambda x: x[0])
    df["Remarques"] = parsed.apply(lambda x: x[1])

    df["Avancement"] = df["Avancement"].fillna(
        parsed.apply(lambda x: x[2])
    ).fillna(0)

    # SUPPRESSION BRUIT
    df = df[~df["Projet"].str.lower().str.contains("compte rendu|tableau", na=False)]
    df = df.drop_duplicates(subset=["Projet"])

    # STATUT
    def statut(x):
        if x <= 30:
            return "Début"
        elif x <= 70:
            return "Milieu"
        else:
            return "Fin"

    df["Statut"] = df["Avancement"].apply(statut)

    # FILTRES
    col1, col2 = st.columns(2)

    resp = col1.selectbox("👤 Responsable", ["Tous"] + sorted(df["Responsable"].unique()))
    proj = col2.selectbox("📁 Projet", ["Tous"] + sorted(df["Projet"].unique()))

    df_f = df.copy()

    if resp != "Tous":
        df_f = df_f[df_f["Responsable"] == resp]

    if proj != "Tous":
        df_f = df_f[df_f["Projet"] == proj]

    # KPI
    c1, c2, c3 = st.columns(3)

    c1.metric("Projets", len(df_f))
    c2.metric("Avancement", f"{df_f['Avancement'].mean():.0f}%")
    c3.metric("Début", len(df_f[df_f["Statut"]=="Début"]))

    # GRAPHIQUES
    st.plotly_chart(px.pie(df_f, names="Statut"), use_container_width=True)

    # TABLEAU
    st.dataframe(df_f, use_container_width=True)

    # =========================
    # MAIL PRO
    # =========================
    st.subheader("📧 Mail prêt à envoyer")

    html_table = generate_html(df_f)

    st.markdown("### ✉️ 1. Ouvre Outlook")
    st.markdown("""
    <a href="https://outlook.office.com/mail/deeplink/compose" target="_blank">
        <button style="background:#0A2463;color:white;padding:10px;border:none;border-radius:5px">
        Ouvrir Outlook
        </button>
    </a>
    """, unsafe_allow_html=True)

    st.markdown("### 📋 2. Copie le tableau ci-dessous et colle-le dans ton mail")

    st.markdown(html_table, unsafe_allow_html=True)