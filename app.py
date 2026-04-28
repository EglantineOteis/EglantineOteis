# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px

# ======================
# CONFIG
# ======================
st.set_page_config(layout="wide")
st.title("📊 Suivi des projets")

# ======================
# IA (OPTIONNEL)
# ======================
USE_AI = False
try:
    from openai import OpenAI
    client = OpenAI()
    USE_AI = True
except:
    USE_AI = False

# ======================
# OUTILS
# ======================
def detect_col(df, mots):
    for mot in mots:
        for c in df.columns:
            if mot in c:
                return c
    return None

def parse(txt):
    d = {"desc":"","rem":"","admin":""}
    if not isinstance(txt,str):
        return d
    for l in txt.split("\n"):
        l2 = l.lower()
        if "descriptif" in l2: d["desc"] = l.split(":")[-1].strip()
        if "remarque" in l2: d["rem"] = l.split(":")[-1].strip()
        if "administration" in l2: d["admin"] = l.split(":")[-1].strip()
    return d

def statut(x):
    if pd.isna(x): return "Inconnu"
    if x == 0: return "Non démarré"
    if x == 100: return "Terminé"
    if x < 40: return "En retard"
    return "En cours"

def mail_table_html(df, col_p, col_r, col_a):
    html = """
    <table style="border-collapse:collapse;width:100%;font-family:Calibri;">
    <tr style="background:#0A2463;color:white;">
    <th>Projet</th><th>Responsable</th><th>Avancement</th><th>Statut</th><th>Remarques</th>
    </tr>
    """
    for _, r in df.iterrows():
        html += f"""
        <tr>
        <td style="border:1px solid #ddd;padding:6px;">{r[col_p]}</td>
        <td style="border:1px solid #ddd;padding:6px;">{r[col_r]}</td>
        <td style="border:1px solid #ddd;padding:6px;">{r[col_a]}%</td>
        <td style="border:1px solid #ddd;padding:6px;">{r['statut']}</td>
        <td style="border:1px solid #ddd;padding:6px;">{r['remarques']}</td>
        </tr>
        """
    html += "</table>"
    return html

# ======================
# UPLOAD
# ======================
file = st.file_uploader("📂 Importer Excel brut", type=["xlsx"])

if file:
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip().str.lower()

    col_p = detect_col(df, ["projet","tache","nom"])
    col_r = detect_col(df, ["responsable","attrib"])
    col_a = detect_col(df, ["avancement","%","progress"])
    col_d = detect_col(df, ["descript","detail","comment"])

    if not all([col_p, col_r, col_a, col_d]):
        st.error("❌ Colonnes non détectées")
        st.stop()

    df[col_a] = pd.to_numeric(df[col_a], errors="coerce")
    df["statut"] = df[col_a].apply(statut)

    parsed = df[col_d].apply(parse)
    df["description"] = parsed.apply(lambda x: x["desc"])
    df["remarques"] = parsed.apply(lambda x: x["rem"])
    df["admin"] = parsed.apply(lambda x: x["admin"])

    # KPI
    st.subheader("📈 Indicateurs")
    c1, c2, c3 = st.columns(3)
    c1.metric("Projets", len(df))
    c2.metric("Avancement moyen", f"{df[col_a].mean():.0f}%")
    c3.metric("En retard", len(df[df["statut"]=="En retard"]))

    # Graphs
    st.subheader("📊 Graphiques")
    g1, g2 = st.columns(2)

    with g1:
        st.plotly_chart(px.pie(df, names="statut"))

    with g2:
        st.plotly_chart(px.histogram(df, x=col_a))

    st.plotly_chart(
        px.bar(df.groupby(col_r)[col_p].count().reset_index(),
               x=col_r, y=col_p,
               title="Charge par responsable")
    )

    # Tableau
    st.subheader("📋 Tableau structuré")
    st.dataframe(df[[col_p, col_r, col_a, "statut","description","remarques"]])

    # Mail
    st.subheader("📧 Tableau pour Outlook")
    html_mail = mail_table_html(df, col_p, col_r, col_a)
    st.markdown("👉 Copier-coller directement dans Outlook")
    st.markdown(html_mail, unsafe_allow_html=True)

    # IA
    st.subheader("🤖 Analyse automatique")
    if USE_AI:
        if st.button("Analyser avec IA"):
            sample = df[[col_p, col_r, col_a, "statut"]].to_string()

            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {"role":"system","content":"Tu es directeur de projet."},
                    {"role":"user","content":f"""
Analyse ce tableau :
- risques
- synthèse
- actions
{sample}
"""}
                ]
            )

            st.success(response.choices[0].message.content)
    else:
        st.info("IA non activée")