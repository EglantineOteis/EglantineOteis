import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html

# === 1. Charger le fichier Excel ===
file_path = "MODELETYPE.xlsx"
df = pd.read_excel(file_path)

# === 2. Nettoyage (adapter aux noms de colonnes réels) ===
# Exemple attendu :
# "Nom de tâche", "Nom du compartiment", "Attribué à", "Avancement (%)"

df.columns = df.columns.str.strip()

# Convertir en % si nécessaire
if "Avancement (%)" in df.columns:
    df["Avancement (%)"] = pd.to_numeric(df["Avancement (%)"], errors="coerce")

# === 3. Graphiques ===

# Avancement moyen par compartiment
fig_bar = px.bar(
    df.groupby("Nom du compartiment")["Avancement (%)"].mean().reset_index(),
    x="Nom du compartiment",
    y="Avancement (%)",
    title="Avancement moyen par compartiment"
)

# Répartition des tâches par personne
fig_pie = px.pie(
    df,
    names="Attribué à",
    title="Répartition des tâches"
)

# Avancement global
avancement_global = df["Avancement (%)"].mean()

# === 4. Dashboard avec Dash ===
app = Dash(__name__)

app.layout = html.Div([
    html.H1("Dashboard Projet"),

    html.H2(f"Avancement global : {round(avancement_global, 1)} %"),

    dcc.Graph(figure=fig_bar),
    dcc.Graph(figure=fig_pie)
])

# === 5. Lancer ===
if __name__ == "__main__":
    app.run(debug=True)