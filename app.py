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
        "description": None,
        "type_projet": None
    }

    for col in df.columns:

        c = str(col).lower().strip()

        # PROJET
        if "tâche" in c or "tache" in c:
            mapping["projet"] = col

        # RESPONSABLE
        elif "responsable" in c:
            mapping["responsable"] = col

        elif "attribué" in c and mapping["responsable"] is None:
            mapping["responsable"] = col

        # TYPE PROJET
        elif "compartiment" in c:
            mapping["type_projet"] = col

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

    txt = txt.replace("_x000D_", " ")
    txt = txt.replace("_x000d_", " ")
    txt = txt.replace("**", " ")
    txt = txt.replace("\\n", " ")
    txt = txt.replace("\n", " ")

    txt = re.sub(r"\s+", " ", txt)

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
    equipe = ""
    av = None

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

    # EQUIPE PROJET
    try:

        if "Liste des intervenant :" in txt:

            equipe = txt.split(
                "Liste des intervenant :", 1
            )[1]

            stop_words = [
                "Remarques",
                "Avancement",
                "Administration"
            ]

            for sw in stop_words:
                if sw in equipe:
                    equipe = equipe.split(sw)[0]

            equipe = clean_text(equipe)

            if equipe.startswith(":"):
                equipe = equipe[1:].strip()

            if equipe.startswith("-"):
                equipe = equipe[1:].strip()

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

    # AVANCEMENT
    try:

        match = re.search(r"(\d+)\s*%", txt)

        if match:
            av = float(match.group(1))

    except:
        pass

    return desc, rem, equipe, av


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
# TABLE HTML
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

            # AVANCEMENT
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

    # TYPE PROJET
    if mapping["type_projet"]:

        df["Type projet"] = (
            df_raw[mapping["type_projet"]]
            .astype(str)
            .apply(clean_text)
        )

    else:
        df["Type projet"] = "Non défini"

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
            parsed.apply(lambda x: x[2])
        )

        avancement_desc = (
            parsed.apply(lambda x: x[3])
        )

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
    # NETTOYAGE
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

    # =========================
    # FILTRES
    # =========================
    col1, col2, col3 = st.columns(3)

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

    with col3:

        type_proj = st.selectbox(
            "Type projet",
            ["Tous"] + sorted(
                df["Type projet"]
                .dropna()
                .unique()
                .tolist()
            )
        )

    # =========================
    # FILTRES DF
    # =========================
    df_f = df.copy()

    if resp != "Tous":
        df_f = df_f[
            df_f["Responsable"] == resp
        ]

    if proj != "Tous":
        df_f = df_f[
            df_f["Projet"] == proj
        ]

    if type_proj != "Tous":
        df_f = df_f[
            df_f["Type projet"] == type_proj
        ]

    # =========================
    # TRI
    # =========================
    df_f = df_f.sort_values(
        by="Type projet",
        ascending=True
    )

    # =========================
    # ORGANISATION COLONNES
    # =========================
    df_f = df_f[
        [
            "Type projet",
            "Projet",
            "Responsable",
            "Avancement",
            "Description",
            "Remarques",
            "Équipe projet",
            "Statut"
        ]
    ]

    # =========================
    # KPI
    # =========================
    st.metric(
        "Projets",
        len(df_f)
    )

    # =========================
    # MAIL
    # =========================
    st.subheader("Mail prêt à envoyer")

    st.markdown("""
    Copier le tableau ci-dessous puis coller directement dans Outlook.
    """)

    html_table = generate_html(df_f)

    # =========================
    # BOUTON COPIER
    # =========================
    copy_code = f"""
    <button onclick="
    navigator.clipboard.writeText(`{html_table}`);
    alert('Tableau copié');
    "
    style="
    background:#0A2463;
    color:white;
    border:none;
    padding:10px 16px;
    border-radius:5px;
    cursor:pointer;
    font-size:15px;
    ">
    Copier le tableau
    </button>
    """

    components.html(copy_code, height=60)

    # =========================
    # TABLEAU HTML
    # =========================
    components.html(
        html_table,
        height=1200,
        scrolling=True
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