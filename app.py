# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import re
import streamlit.components.v1 as components

st.set_page_config(layout="wide")

# =========================
# DETECTION COLONNES
# =========================
def detect_columns(df):

    mapping = {
        "projet": None,
        "responsable": None,
        "avancement": None,
        "description": None
    }

    for col in df.columns:

        c = str(col).lower().strip()

        # NOM DE TACHE = PROJET
        if "tâche" in c or "tache" in c:
            mapping["projet"] = col

        # RESPONSABLE
        elif "responsable" in c:
            mapping["responsable"] = col

        elif "attribué" in c and mapping["responsable"] is None:
            mapping["responsable"] = col

        # AVANCEMENT
        elif "progress" in c:
            mapping["avancement"] = col

        # DESCRIPTION
        elif "description" in c:
            mapping["description"] = col

    return mapping


# =========================
# NETTOYAGE TEXTE
# =========================
def clean_text(txt):

    if pd.isna(txt):
        return ""

    txt = str(txt)

    # BRUITS EXCEL
    txt = txt.replace("_x000D_", " ")
    txt = txt.replace("_x000d_", " ")
    txt = txt.replace("**", " ")
    txt = txt.replace("\\n", " ")
    txt = txt.replace("\n", " ")

    # ESPACES
    txt = re.sub(r"\s+", " ", txt)

    # CARACTERES PARASITES
    txt = txt.replace("Â", "")
    txt = txt.replace("Ã", "")

    return txt.strip()


# =========================
# EXTRACTION DESCRIPTION
# =========================
def parse_description(txt):

    txt = clean_text(txt)

    desc = ""
    rem = ""
    av = None
    equipe = ""

    # DESCRIPTION
    try:
        if "Descriptif :" in txt:

            desc = txt.split("Descriptif :", 1)[1]

            stop_words = [
                "Administration",
                "Liste des intervenant",
                "Remarques",
                "Avancement"
            ]

            for sw in stop_words:
                if sw in desc:
                    desc = desc.split(sw)[0]

            desc = clean_text(desc)

    except:
        pass

    # REMARQUES
    try:
        if "Remarques :" in txt:

            rem = txt.split("Remarques :", 1)[1]

            stop_words = [
                "Descriptif",
                "Administration",
                "Liste des intervenant",
                "Avancement"
            ]

            for sw in stop_words:
                if sw in rem:
                    rem = rem.split(sw)[0]

            rem = clean_text(rem)

    except:
        pass

    # EQUIPE PROJET
    try:

        if "Liste des intervenant" in txt:

            equipe = txt.split("Liste des intervenant", 1)[1]

            stop_words = [
                "Remarques",
                "Avancement",
                "Administration"
            ]

            for sw in stop_words:
                if sw in equipe:
                    equipe = equipe.split(sw)[0]

            equipe = clean_text(equipe)

            equipe = equipe.replace("•", "")
            equipe = equipe.replace("-", "")

    except:
        pass

    # AVANCEMENT
    try:

        match = re.search(r"(\d+)\s*%", txt)

        if match:
            av = float(match.group(1))

    except:
        pass

    return desc, rem, av, equipe


# =========================
# RESPONSABLE
# =========================
def clean_responsable(x):

    if pd.isna(x):
        return "Non défini"

    x = clean_text(x)

    if ";" in x:
        x = x.split(";")[0]

    if "," in x:
        x = x.split(",")[0]

    x = x.strip()

    if x == "":
        return "Non défini"

    return x


# =========================
# STATUT
# =========================
def statut(x):

    if x <= 30:
        return "Début"

    elif x <= 70:
        return "Milieu"

    return "Fin"


# =========================
# TABLE HTML PRO
# =========================
def generate_html(df):

    html = """
    <table style="
        border-collapse:collapse;
        font-family:Calibri;
        font-size:13px;
        width:100%;
    ">
    """

    # HEADER
    html += """
    <tr style="
        background:#0A2463;
        color:white;
        font-weight:bold;
    ">
    """

    for col in df.columns:

        html += f"""
        <th style="
            border:1px solid #d9d9d9;
            padding:14px 10px;
            text-align:center;
            background:#0A2463;
            color:#F4A300;
            font-weight:bold;
            font-size:16px;
            font-family:Calibri;
        ">
            {col}
        </th>
        """

    html += "</tr>"

    # LIGNES
    for i, (_, row) in enumerate(df.iterrows()):

        bg = "#ffffff" if i % 2 == 0 else "#f4f6fa"

        html += f"<tr style='background:{bg}'>"

        for col in df.columns:

            val = clean_text(row[col])

            # AVANCEMENT COLORÉ
            if col == "Avancement":

                color = "#d62828"

                if row["Statut"] == "Milieu":
                    color = "#f77f00"

                elif row["Statut"] == "Fin":
                    color = "#2a9d8f"

                html += f"""
                <td style="
                    border:1px solid #d9d9d9;
                    padding:6px;
                    text-align:center;
                    color:{color};
                    font-weight:bold;
                ">
                {val}%
                </td>
                """

            else:

                html += f"""
                <td style="
                    border:1px solid #d9d9d9;
                    padding:6px;
                    vertical-align:top;
                ">
                {val}
                </td>
                """

        html += "</tr>"

    html += "</table>"

    return html


# =========================
# UI
# =========================
st.title("Suivi des projets")

file = st.file_uploader(
    "Importer Excel",
    type=["xlsx"]
)

if file:

    # =========================
    # LECTURE EXCEL
    # =========================
    df_raw = pd.read_excel(file)

    df_raw.columns = [
        str(c).lower().strip()
        for c in df_raw.columns
    ]

    mapping = detect_columns(df_raw)

    st.write("Colonnes détectées :", mapping)

    # =========================
    # VERIFICATION
    # =========================
    if mapping["projet"] is None:

        st.error(
            "Colonne 'Nom de tâche' introuvable"
        )

        st.stop()

    # =========================
    # DATAFRAME FINAL
    # =========================
    df = pd.DataFrame()

    # PROJET
    df["Projet"] = (
        df_raw[mapping["projet"]]
        .astype(str)
        .apply(clean_text)
    )

    # RESPONSABLE
    if mapping["responsable"]:

        df["Responsable"] = (
            df_raw[mapping["responsable"]]
            .apply(clean_responsable)
        )

    else:
        df["Responsable"] = "Non défini"

    # AVANCEMENT
    if mapping["avancement"]:

        df["Avancement"] = pd.to_numeric(
            df_raw[mapping["avancement"]],
            errors="coerce"
        )

    else:
        df["Avancement"] = None

    # DESCRIPTION
    if mapping["description"]:

        parsed = (
            df_raw[mapping["description"]]
            .apply(parse_description)
        )

        df["Description"] = (
            parsed.apply(lambda x: x[0])
        )

        df["Remarques"] = (
            parsed.apply(lambda x: x[1])
        )

        df["Équipe projet"] = (
            parsed.apply(lambda x: x[3])
        )

        avancement_desc = (
            parsed.apply(lambda x: x[2])
        )

        # COMPLETE AVANCEMENT
        df["Avancement"] = (
            df["Avancement"]
            .fillna(avancement_desc)
            .fillna(0)
        )

    else:

        df["Description"] = ""
        df["Remarques"] = ""
        df["Équipe projet"] = ""

        df["Avancement"] = (
            df["Avancement"]
            .fillna(0)
        )

    # =========================
    # NETTOYAGE FINAL
    # =========================
    df["Projet"] = df["Projet"].apply(clean_text)
    df["Description"] = df["Description"].apply(clean_text)
    df["Remarques"] = df["Remarques"].apply(clean_text)
    df["Équipe projet"] = df["Équipe projet"].apply(clean_text)

    # SUPPRESSION LIGNES INUTILES
    df = df[
        ~df["Projet"]
        .str.lower()
        .str.contains(
            "compte rendu|tableau de gestion",
            na=False
        )
    ]

    # VIDE
    df = df[df["Projet"] != ""]

    # DOUBLONS
    df = df.drop_duplicates(
        subset=["Projet"]
    )

    # AVANCEMENT
    df["Avancement"] = (
        pd.to_numeric(
            df["Avancement"],
            errors="coerce"
        )
        .fillna(0)
        .round(0)
    )

    # STATUT
    df["Statut"] = (
        df["Avancement"]
        .apply(statut)
    )

    # ORDRE COLONNES
    df = df[
        [
            "Projet",
            "Responsable",
            "Équipe projet",
            "Avancement",
            "Description",
            "Remarques",
            "Statut"
        ]
    ]

    # =========================
    # FILTRES
    # =========================
    col1, col2 = st.columns(2)

    with col1:

        resp = st.selectbox(
            "Responsable",
            ["Tous"] + sorted(
                df["Responsable"]
                .dropna()
                .unique()
                .tolist()
            )
        )

    with col2:

        proj = st.selectbox(
            "Projet",
            ["Tous"] + sorted(
                df["Projet"]
                .dropna()
                .unique()
                .tolist()
            )
        )

    # FILTRE
    df_f = df.copy()

    if resp != "Tous":
        df_f = df_f[
            df_f["Responsable"] == resp
        ]

    if proj != "Tous":
        df_f = df_f[
            df_f["Projet"] == proj
        ]

    # =========================
    # KPI
    # =========================
    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Projets",
        len(df_f)
    )

    c2.metric(
        "Avancement moyen",
        f"{df_f['Avancement'].mean():.0f}%"
    )

    c3.metric(
        "Projets début",
        len(df_f[df_f["Statut"] == "Début"])
    )

    # =========================
    # GRAPHIQUES
    # =========================
    st.subheader("Graphiques")

    fig = px.pie(
        df_f,
        names="Statut"
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

    # =========================
    # TABLEAU
    # =========================
    st.subheader("Tableau structuré")

    st.dataframe(
        df_f,
        use_container_width=True
    )

    # =========================
    # TABLEAU MAIL
    # =========================
    st.subheader("Mail prêt à envoyer")

    st.markdown("""
    <a href="https://outlook.office.com/mail/deeplink/compose" target="_blank">
        <button style="
            background:#0A2463;
            color:white;
            border:none;
            padding:10px 16px;
            border-radius:5px;
            cursor:pointer;
        ">
        Ouvrir Outlook
        </button>
    </a>
    """, unsafe_allow_html=True)

    st.markdown("""
    Copier le tableau ci-dessous puis coller directement dans Outlook.
    """)

    html_table = generate_html(df_f)

    components.html(
        html_table,
        height=1200,
        scrolling=True
    )