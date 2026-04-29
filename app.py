# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import urllib.parse

st.set_page_config(layout="wide")

# =========================
# 🔍 DETECTION ROBUSTE
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
# 🧠 PARSE TEXTE PROPRE
# =========================
def parse_desc(txt):

    data = {
        "desc": "",
        "rem": "",
        "av": None
    }

    if not isinstance(txt, str):
        return data

    for l in txt.split("\n"):

        if ":" not in l:
            continue

        key, val = l.split(":", 1)
        key = key.lower()
        val = val.strip()

        if "descriptif" in key:
            data["desc"] = val

        elif "remarque" in key:
            data["rem"] = val

        elif "avancement" in key:
            try:
                data["av"] = float(val.replace("%",""))
            except:
                pass

    return data


# =========================
# 📊 TABLE HTML PRO
# =========================
def generate_html(df):

    html = "<table style='border-collapse:collapse;font-family:Calibri;width:100%'>"

    # HEADER
    html += "<tr style='background:#0A2463;color:white'>"
    for col in df.columns:
        html += f"<th style='padding:10px;border:1px solid #ddd'>{col}</th>"
    html += "</tr>"

    # ROWS
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
# 📧 MAIL SAFE (SANS BUG)
# =========================
def outlook_mail(df):

    sujet = "Suivi des projets"

    texte = "Bonjour,\n\nVoici le suivi des projets.\n\n"

    for _, r in df.iterrows():
        texte += f"- {r['Projet']} | {r['Responsable']} | {r['Avancement']}%\n"

    texte += "\nCordialement"

    url = f"mailto:?subject={urllib.parse.quote(sujet)}&body={urllib.parse.quote(texte)}"

    return url


# =========================
# 🤖 IA PRO (SANS API)
# =========================
def ai_summary(df):

    total = len(df)
    av = df["Avancement"].mean()
    retard = len(df[df["Statut"] == "En retard"])

    critiques = df[df["Avancement"] < 40]["Projet"].head(5).tolist()

    txt = f"""
Synthèse direction :

- {total} projets en cours
- Avancement moyen : {av:.0f}%
- {retard} projets en retard

Points critiques :
"""

    for p in critiques:
        txt += f"\n- {p}"

    txt += "\n\nRecommandation : prioriser les projets en dessous de 40%."

    return txt


# =========================
# UI
# =========================
st.title("📊 Suivi des projets intelligent")

file = st.file_uploader("📂 Importer Excel", type=["xlsx"])

if file:

    df_raw = pd.read_excel(file)
    df_raw.columns = df_raw.columns.str.lower().str.strip()

    mapping = smart_detect(df_raw)

    if mapping["projet"] is None:
        st.error("❌ colonne projet introuvable")
        st.stop()

    df = pd.DataFrame()

    df["Projet"] = df_raw[mapping["projet"]]

    df["Responsable"] = (
        df_raw[mapping["responsable"]]
        .astype(str)
        .str.split(";")
        .str[0]
        .fillna("Non défini")
    )

    df["Avancement"] = pd.to_numeric(
        df_raw[mapping["avancement"]],
        errors="coerce"
    )

    # fallback via description
    parsed = df_raw[mapping["description"]].apply(parse_desc)

    df["Description"] = parsed.apply(lambda x: x["desc"])
    df["Remarques"] = parsed.apply(lambda x: x["rem"])

    df["Avancement"] = df["Avancement"].fillna(
        parsed.apply(lambda x: x["av"])
    ).fillna(0)

    # statut
    def statut(x):
        if x == 0: return "Non démarré"
        elif x == 100: return "Terminé"
        elif x < 40: return "En retard"
        return "En cours"

    df["Statut"] = df["Avancement"].apply(statut)

    # ================= FILTERS
    resp = st.selectbox("👤 Chef de projet", ["Tous"] + sorted(df["Responsable"].unique()))
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

    # ================= GRAPH
    st.plotly_chart(px.pie(df_f, names="Statut", hole=0.5), use_container_width=True)

    # ================= TABLE
    st.dataframe(df_f, use_container_width=True)

    # ================= TABLE HTML
    html = generate_html(df_f)
    st.markdown(html, unsafe_allow_html=True)

    # ================= MAIL BUTTON
    mail_link = outlook_mail(df_f)

    st.markdown(f"""
    <a href="{mail_link}">
        <button style="background:#E85D04;color:white;padding:12px;border:none;border-radius:6px">
        📧 Envoyer mail
        </button>
    </a>
    """, unsafe_allow_html=True)

    # ================= IA
    st.subheader("🤖 Analyse automatique")
    st.text(ai_summary(df_f))