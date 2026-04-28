# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import re
import html

st.set_page_config(layout="wide")

# ==============================
# NORMALISATION (ADAPTÉE PLANNER)
# ==============================
def normalize_columns(df):

    original = df.columns.tolist()
    df.columns = df.columns.str.lower().str.strip()

    mapping = {}

    for c in df.columns:

        # 🎯 SEULE colonne projet valide
        if "nom de tâche" in c:
            mapping[c] = "projet"

        elif "attribué" in c:
            mapping[c] = "responsable"

        elif "progression" in c:
            mapping[c] = "progression"

        elif "description" in c:
            mapping[c] = "description_brut"

        elif "compartiment" in c:
            mapping[c] = "equipe"

    df = df.rename(columns=mapping)

    # 🔥 supprimer doublons
    df = df.loc[:, ~df.columns.duplicated()]

    return df


# ==============================
# AVANCEMENT
# ==============================
def clean_progress(x):

    if pd.isna(x):
        return 0

    x = str(x).lower()

    if "non" in x:
        return 0
    elif "cours" in x:
        return 50
    elif "term" in x:
        return 100

    return 0


# ==============================
# PARSE DESCRIPTION PROPRE
# ==============================
def parse_description(txt):

    if not isinstance(txt, str):
        return "", ""

    def extract(label):
        match = re.search(rf"{label}\s*:\s*(.*?)(?=\n[A-Z]|$)", txt, re.IGNORECASE | re.DOTALL)
        return match.group(1).strip() if match else ""

    desc = extract("Descriptif")
    rem = extract("Remarque")

    return desc, rem


# ==============================
# RESPONSABLE
# ==============================
def clean_responsable(row):

    if pd.notna(row.get("responsable")):
        return row["responsable"]

    if pd.notna(row.get("equipe")):
        return row["equipe"]

    return "Non défini"


# ==============================
# TRANSFORMATION
# ==============================
def transform(df):

    df = normalize_columns(df)

    st.write("Colonnes détectées :", df.columns.tolist())

    # sécurité
    for col in ["projet", "progression", "description_brut"]:
        if col not in df:
            df[col] = ""

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
# HTML OUTLOOK
# ==============================
def generate_html_table(df):

    html_table = """
    <table style="
        border-collapse:collapse;
        font-family:Calibri;
        width:100%;
        font-size:13px;
    ">
    """

    # ======================
    # HEADER BLEU
    # ======================
    html_table += """
    <tr style="background:#0A2463;color:white;text-align:left;">
    """

    for col in df.columns:
        html_table += f"""
        <th style="
            padding:10px;
            border:1px solid #d9d9d9;
        ">{col}</th>
        """

    html_table += "</tr>"

    # ======================
    # LIGNES
    # ======================
    for i, (_, row) in enumerate(df.iterrows()):

        bg = "#ffffff" if i % 2 == 0 else "#f4f6fa"

        html_table += f"<tr style='background:{bg}'>"

        for col in df.columns:

            val = str(row[col]).replace("\n", "<br>")

            # 🎯 couleur spéciale pour avancement
            if col == "avancement":
                cell_style = """
                padding:8px;
                border:1px solid #ddd;
                color:#E85D04;
                font-weight:bold;
                text-align:center;
                """
                val = f"{val}%"

            else:
                cell_style = """
                padding:8px;
                border:1px solid #ddd;
                """

            html_table += f"<td style='{cell_style}'>{val}</td>"

        html_table += "</tr>"

    html_table += "</table>"

    return html_tablee


# ==============================
# UI
# ==============================
st.title("📊 Suivi des projets")

file = st.file_uploader("📂 Importer Excel", type=["xlsx"])

if file:

    df_raw = pd.read_excel(file)
    df = transform(df_raw)

    if "projet" not in df.columns:
        st.error("❌ Colonne 'Nom de tâche' introuvable")
        st.stop()

    # ==============================
    # FILTRES
    # ==============================
    col1, col2 = st.columns(2)

    responsables = ["Tous"] + sorted(df["responsable"].dropna().unique())
    projets = ["Tous"] + sorted(df["projet"].dropna().unique())

    resp_filter = col1.selectbox("Responsable", responsables)
    proj_filter = col2.selectbox("Projet", projets)

    dff = df.copy()

    if resp_filter != "Tous":
        dff = dff[dff["responsable"] == resp_filter]

    if proj_filter != "Tous":
        dff = dff[dff["projet"] == proj_filter]

    # ==============================
    # KPI
    # ==============================
    st.subheader("📈 Indicateurs")

    c1, c2, c3 = st.columns(3)

    c1.metric("Projets", len(dff))
    c2.metric("Avancement moyen", f"{dff['avancement'].mean():.0f}%")
    c3.metric("En cours", len(dff[dff["statut"] == "En cours"]))

    # ==============================
    # GRAPHIQUES
    # ==============================
    st.subheader("📊 Graphiques")

    st.plotly_chart(px.pie(dff, names="statut"), use_container_width=True)

    st.plotly_chart(
        px.bar(dff.groupby("responsable")["projet"].count().reset_index(),
               x="responsable", y="projet",
               title="Charge par responsable"),
        use_container_width=True
    )

    # ==============================
    # TABLEAU EXCEL
    # ==============================
    st.subheader("📋 Tableau structuré")
    st.dataframe(dff, use_container_width=True)

    # ==============================
    # TABLE OUTLOOK
    # ==============================
    st.subheader("📧 Tableau Outlook")

    html_table = generate_html_table(dff)

    st.code(html_table, language="html")
    st.markdown("👉 Copier-coller dans Outlook")