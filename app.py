# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import re
import streamlit.components.v1 as components

st.set_page_config(layout="wide")

# =========================
# TITRE
# =========================
st.title("SUIVIS DES PROJETS")

# =========================
# DETECTION COLONNES
# =========================
def detect_columns(df):

    mapping = {
        "projet": None,
        "responsable": None,
        "description": None,
        "avancement": None,
        "compartiment": None
    }

    for col in df.columns:

        c = str(col).lower().strip()

        if "tâche" in c or "tache" in c:
            mapping["projet"] = col

        elif "attribué" in c or "responsable" in c:
            mapping["responsable"] = col

        elif "description" in c:
            mapping["description"] = col

        elif "progress" in c:
            mapping["avancement"] = col

        elif "compartiment" in c:
            mapping["compartiment"] = col

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
    txt = txt.replace("**", "")
    txt = txt.replace("\n", " ")
    txt = txt.replace("\\n", " ")

    txt = re.sub(r"\s+", " ", txt)

    return txt.strip()

# =========================
# DESCRIPTION
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

    # REMARQUES
    try:

        if "Remarques :" in txt:

            rem = txt.split("Remarques :", 1)[1]

            stop_words = [
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

        if "Liste des intervenant :" in txt:

            equipe = txt.split("Liste des intervenant :", 1)[1]

            stop_words = [
                "Remarques",
                "Avancement"
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

    if x == "":
        return "Non défini"

    return x

# =========================
# PHASE EN COURS DYNAMIQUE
# =========================
def get_phase_en_cours(done_value, phases_value):

    try:

        # ex : 5/6
        done = str(done_value).split("/")[0]
        done = int(done)

    except:
        return "Non définie"

    phases_text = clean_text(phases_value)

    if phases_text == "":
        return "Non définie"

    # séparation ;
    phases = [
        p.strip()
        for p in phases_text.split(";")
        if p.strip() != ""
    ]

    clean_phases = []

    for p in phases:

        # suppression dates
        p = re.sub(
            r"\d{2}/\d{2}/\d{4}",
            "",
            p
        )

        p = clean_text(p)

        clean_phases.append(p)

    # IMPORTANT
    # le tableau est inversé :
    # PRO est souvent au début
    clean_phases.reverse()

    # sécurité
    if done >= len(clean_phases):
        return "Terminée"

    return clean_phases[done]

# =========================
# TABLE HTML
# =========================
def generate_html(df):

    html = """
    <table style="
        border-collapse:collapse;
        width:100%;
        font-family:Calibri;
        font-size:14px;
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
            padding:12px;
            color:#F4A300;
            text-align:center;
            font-size:15px;
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

                phase = str(row["Phase en cours"])

                if phase in ["ACT", "PRO", "EXE"]:
                    color = "#f77f00"

                elif phase in ["DET", "AOR", "Terminée"]:
                    color = "#2a9d8f"

                html += f"""
                <td style="
                    border:1px solid #d9d9d9;
                    padding:8px;
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
                    padding:8px;
                    vertical-align:top;
                ">
                {val}
                </td>
                """

        html += "</tr>"

    html += "</table>"

    return html

# =========================
# IMPORT EXCEL
# =========================
file = st.file_uploader(
    "Importer Excel",
    type=["xlsx"]
)

if file:

    # =========================
    # LECTURE
    # =========================
    df_raw = pd.read_excel(file)

    df_raw.columns = [
        str(c).lower().strip()
        for c in df_raw.columns
    ]

    mapping = detect_columns(df_raw)

    # =========================
    # DATAFRAME
    # =========================
    df = pd.DataFrame()

    # COMPARTIMENT
    if mapping["compartiment"]:

        df["Compartiment"] = (
            df_raw[mapping["compartiment"]]
            .astype(str)
            .apply(clean_text)
        )

    else:

        df["Compartiment"] = "Non défini"

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

    # DESCRIPTION
    parsed = (
        df_raw[mapping["description"]]
        .apply(parse_description)
    )

    df["Description"] = parsed.apply(lambda x: x[0])
    df["Remarques"] = parsed.apply(lambda x: x[1])
    df["Équipe projet"] = parsed.apply(lambda x: x[2])

    # AVANCEMENT
    if mapping["avancement"]:

        df["Avancement"] = pd.to_numeric(
            df_raw[mapping["avancement"]],
            errors="coerce"
        )

    else:

        df["Avancement"] = None

    avancement_desc = parsed.apply(lambda x: x[3])

    df["Avancement"] = (
        df["Avancement"]
        .fillna(avancement_desc)
        .fillna(0)
        .round(0)
    )

    # =========================
    # PHASE EN COURS
    # =========================

    # colonne O
    done_column = df_raw.iloc[:, 14]

    # colonne P
    phases_column = df_raw.iloc[:, 15]

    df["Phase en cours"] = [
        get_phase_en_cours(done, phases)
        for done, phases in zip(done_column, phases_column)
    ]

    # =========================
    # NETTOYAGE
    # =========================
    df = df[
        ~df["Projet"]
        .str.lower()
        .str.contains(
            "compte rendu|tableau de gestion",
            na=False
        )
    ]

    df = df.drop_duplicates(
        subset=["Projet"]
    )

    # =========================
    # TRI
    # =========================
    df = df.sort_values(
        by="Compartiment",
        ascending=True
    )

    # =========================
    # FILTRES
    # =========================
    col1, col2, col3 = st.columns(3)
    
    with col1:
    
        compartiment = st.selectbox(
            "Compartiment",
            ["Tous"] + sorted(
                df["Compartiment"]
                .dropna()
                .unique()
                .tolist()
            )
        )
    
    with col2:
    
        resp = st.selectbox(
            "Responsable",
            ["Tous"] + sorted(
                df["Responsable"]
                .dropna()
                .unique()
                .tolist()
            )
        )
    
    with col3:
    
        phase = st.selectbox(
            "Phase en cours",
            ["Toutes"] + sorted(
                df["Phase en cours"]
                .dropna()
                .unique()
                .tolist()
            )
        )

    # =========================
    # FILTRE DF
    # =========================
    df_f = df.copy()

    if compartiment != "Tous":

        df_f = df_f[
            df_f["Compartiment"] == compartiment
        ]

    if resp != "Tous":

        df_f = df_f[
            df_f["Responsable"] == resp
        ]
    
    if phase != "Toutes":
    
        df_f = df_f[
            df_f["Phase en cours"] == phase
        ]

    # =========================
    # ORGANISATION COLONNES
    # =========================
    df_f = df_f[
        [
            "Compartiment",
            "Projet",
            "Responsable",
            "Avancement",
            "Phase en cours",
            "Description",
            "Remarques",
            "Équipe projet"
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
    Cliquez sur le bouton puis collez directement dans Outlook avec Ctrl+V.
    """)

    html_table = generate_html(df_f)

    copy_html = f"""
    <button onclick="
    navigator.clipboard.write([
    new ClipboardItem({{
    'text/html': new Blob(
    [document.getElementById('tableau_mail').innerHTML],
    {{type: 'text/html'}}
    )
    }})
    ]);
    alert('Tableau copié');
    "
    style="
    background:#0A2463;
    color:white;
    border:none;
    padding:10px 18px;
    border-radius:5px;
    cursor:pointer;
    font-size:15px;
    margin-bottom:15px;
    ">
    Copier le tableau
    </button>

    <div id="tableau_mail">
    {html_table}
    </div>
    """

    components.html(
        copy_html,
        height=1200,
        scrolling=True
    )

    # =========================
    # GRAPHIQUES
    # =========================
    st.subheader("Graphiques")

    # CAMEMBERT PHASE
    fig_phase = px.pie(
        df_f,
        names="Phase en cours",
        title="Répartition des projets par phase"
    )

    st.plotly_chart(
        fig_phase,
        use_container_width=True
    )

    # RESPONSABLE
    df_resp = (
        df_f["Responsable"]
        .value_counts()
        .reset_index()
    )

    df_resp.columns = [
        "Responsable",
        "Nombre"
    ]

    fig_resp = px.bar(
        df_resp,
        x="Responsable",
        y="Nombre",
        title="Répartition des projets par responsable",
        text="Nombre"
    )

    fig_resp.update_layout(
        xaxis_title="Responsable",
        yaxis_title="Nombre de projets"
    )

    st.plotly_chart(
        fig_resp,
        use_container_width=True
    )