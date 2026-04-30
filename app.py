# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import urllib.parse

st.set_page_config(layout="wide")

# =========================
# 🔍 DETECTION COLONNES
# =========================
def detect_columns(df):

    mapping = {
        "projet": None,
        "responsable": None,
        "avancement": None,
        "description": None
    }

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
# 🧠 PARSE TEXTE
# =========================
def parse_description(txt):

    desc = ""
    rem = ""
    av = None

    if not isinstance(txt, str):
        return desc, rem, av

    try:
        if "Descriptif" in txt:
            desc = txt.split("Descriptif :",1)[1].split("Administration")[0].strip()
    except:
        pass

    try:
        if "Remarques" in txt:
            rem = txt.split("Remarques :",1)[1].strip()
    except:
        pass

    try:
        if "Avancement" in txt:
            val = txt.split("Avancement :",1)[1].split("\n")[0]
            av = float(val.replace("%","").strip())
    except:
        pass

    return desc, rem, av


# =========================
# 👤 NETTOYAGE RESPONSABLE
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
# 📊 TABLE HTML
# =========================
def generate_html(df):

    html = "<table style='border-collapse:collapse;font-family:Calibri;width:100%'>"

    html += "<tr style='background:#0A2463;color:white'>"
    for col in df.columns:
        html += f"<th style='padding:10px;border:1px solid #ddd'>{col}</th>"
    html += "</tr>"

    for i, (_, row) in enumerate(df.iterrows()):
        bg = "#f4f6fa" if i % 2 else "white"
        html += f"<tr style='background:{bg}'>"

        for col in df.columns:
            val = str(row[col]).replace("\n", "<br>")

            if col == "Avancement":
                html += f"<td style='padding:8px;border:1px solid #ddd;color:#E85D04;font-weight:bold;text-align:center'>{val}%</td>"
            else:
                html += f"<td style='padding:8px;border:1px solid #ddd'>{val}</td>"

        html += "</tr>"

    html += "</table>"
    return html


# =========================
# 📧 MAIL (SAFE)
# =========================
def generate_mail(df):

    texte = "Bonjour,\n\nVoici le suivi des projets :\n\n"

    for _, r in df.iterrows():
        texte += f"- {r['Projet']} | {r['Responsable']} | {r['Avancement']}%\n"

    texte += "\nCordialement"

    return f"mailto:?subject={urllib.parse.quote('Suivi des projets')}&body={urllib.parse.quote(texte)}"


# =========================
# 🤖 IA SIMPLE PRO
# =========================
def ai_summary(df):

    total = len(df)
    av = df["Avancement"].mean()
    retard = len(df[df["Statut"] == "En retard"])

    critiques = df[df["Avancement"] < 40]["Projet"].head(5).tolist()

    txt = f"""
📊 Synthèse direction

- {total} projets
- Avancement moyen : {av:.0f}%
- {retard} projets en retard

⚠️ Projets critiques :
"""

    for p in critiques:
        txt += f"\n- {p}"

    txt += "\n\n👉 Recommandation : prioriser les projets < 40%"

    return txt


# =========================
# UI
# =========================
st.title("📊 Suivi des projets intelligent")

file = st.file_uploader("📂 Importer Excel", type=["xlsx"])

if file:

    df_raw = pd.read_excel(file)
    df_raw.columns = df_raw.columns.str.lower().str.strip()

    mapping = detect_columns(df_raw)

    if mapping["projet"] is None:
        st.error("❌ colonne 'Nom de tâche' introuvable")
        st.stop()

    df = pd.DataFrame()

    # ================= EXTRACTION
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

    # ================= NETTOYAGE
    df = df[~df["Projet"].str.lower().str.contains("compte rendu|tableau", na=False)]
    df = df.drop_duplicates(subset=["Projet"])

    # ================= STATUT
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

    # ================= FILTRES
    col1, col2 = st.columns(2)

    with col1:
        resp = st.selectbox("👤 Chef de projet", ["Tous"] + sorted(df["Responsable"].unique()))

    with col2:
        proj = st.selectbox("📁 Projet", ["Tous"] + sorted(df["Projet"].unique()))

    df_f = df.copy()

    if resp != "Tous":
        df_f = df_f[df_f["Responsable"] == resp]

    if proj != "Tous":
        df_f = df_f[df_f["Projet"] == proj]

    # ================= KPI
    c1, c2, c3 = st.columns(3)

    c1.metric("Projets", len(df_f))
    c2.metric("Avancement", f"{df_f['Avancement'].mean():.0f}%")
    c3.metric("En retard", len(df_f[df_f["Statut"] == "En retard"]))

    # ================= GRAPHIQUES
    st.subheader("📊 Graphiques")

    st.plotly_chart(px.pie(df_f, names="Statut", hole=0.5), use_container_width=True)

    st.plotly_chart(
        px.bar(df_f.groupby("Responsable").size().reset_index(name="nb"),
               x="Responsable", y="nb"),
        use_container_width=True
    )

    # ================= TABLE
    st.subheader("📋 Tableau structuré")
    st.dataframe(df_f, use_container_width=True)

    # ================= TABLE HTML
    st.subheader("📧 Tableau Outlook")

    html = generate_html(df_f)
    st.markdown(html, unsafe_allow_html=True)

    # ================= MAIL
    st.markdown(f"""
    <a href="{generate_mail(df_f)}">
        <button style="background:#E85D04;color:white;padding:12px;border:none;border-radius:6px">
        📧 Envoyer le mail
        </button>
    </a>
    """, unsafe_allow_html=True)

    # ================= IA
    st.subheader("🤖 Analyse automatique")
    st.text(ai_summary(df_f))