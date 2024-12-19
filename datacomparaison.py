import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from io import BytesIO

# Configuration de l'application Streamlit
st.set_page_config(
    page_title="Analyse Acoustique Interactive",
    page_icon=":chart_with_upwards_trend:",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Liste des fréquences prédéfinies
frequencies = np.array([
    200, 250, 315, 400, 500, 630, 800, 1000,
    1250, 1600, 2000, 2500, 3150, 4000, 5000,
    6300, 8000, 10000
])

# Fonction pour charger des fichiers Excel
def load_excel(file):
    """
    Charge un fichier Excel, prenant en charge les formats .xls et .xlsx.
    """
    try:
        # Détecte le format du fichier
        if file.name.endswith(".xls"):
            df = pd.read_excel(file, sheet_name="Macro", engine="xlrd", header=None)
        else:
            df = pd.read_excel(file, sheet_name="Macro", engine="openpyxl", header=None)
        return df
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {e}")
        return None

# Fonction pour extraire des données
def extract_data(df):
    """
    Extrait les séries valides du DataFrame.
    """
    extracted_data = []

    for col_start, name_cell, values_range in [
        (0, "A1", "C3:C20"), (4, "E1", "G3:G20"),
        (8, "I1", "K3:K20"), (12, "M1", "O3:O20"),
        (16, "Q1", "S3:S20"), (20, "U1", "W3:W20")
    ]:
        try:
            # Lire le nom de la série
            name = df.iloc[0, col_start]

            # Lire les valeurs (en évitant les erreurs non numériques comme #DIV/0!)
            values = pd.to_numeric(
                df.iloc[2:20, col_start + 2], errors="coerce"
            ).values  # `coerce` transforme les erreurs en NaN

            # Vérifie si au moins une valeur est numérique
            if not np.isnan(values).all():
                extracted_data.append({
                    "name": name,  # Utilise uniquement le nom de l'échantillon
                    "values": values
                })
        except Exception:
            continue

    return extracted_data

# Téléchargement de fichiers
uploaded_files = st.sidebar.file_uploader(
    "Importez vos fichiers Excel (.xls ou .xlsx)",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

# Stocker toutes les séries extraites
all_series = []

# Charger les fichiers et extraire les données
if uploaded_files:
    for file in uploaded_files:
        st.subheader(f"Données extraites de : {file.name}")

        # Charger le fichier
        df = load_excel(file)
        if df is None:
            continue

        # Extraire les données valides
        extracted_data = extract_data(df)

        # Ajouter les séries au tableau global
        all_series.extend(extracted_data)

# Afficher les options de sélection si des séries valides existent
if all_series:
    # Liste des noms des séries disponibles
    series_names = [series["name"] for series in all_series]

    # Permettre à l'utilisateur de sélectionner les séries à afficher
    selected_series_names = st.sidebar.multiselect(
        "Choisissez les séries à afficher",
        options=series_names,
        default=series_names  # Par défaut, toutes les séries sont sélectionnées
    )

    # Filtrer les séries sélectionnées
    selected_series = [series for series in all_series if series["name"] in selected_series_names]

    # Générer un graphique unique pour toutes les séries sélectionnées
    fig, ax = plt.subplots(figsize=(12, 8))
    for series in selected_series:
        ax.plot(frequencies, series["values"], label=series["name"], marker="o")

    # Personnalisation du graphique
    ax.set_title("Courbes d'absorption (Séries sélectionnées)")
    ax.set_xscale("log")
    ax.set_xticks(frequencies)
    ax.get_xaxis().set_major_formatter(plt.ScalarFormatter())
    ax.set_xlabel("Fréquence (Hz)")
    ax.set_ylabel("Coefficient d'absorption")
    ax.legend(loc="best", fontsize=10)
    ax.grid(True, which="both", linestyle="--", linewidth=0.5)

    # Afficher le graphique dans Streamlit
    st.pyplot(fig)

    # Fonction pour enregistrer le graphique au format PDF
    def save_as_pdf(fig):
        pdf_buffer = BytesIO()
        fig.savefig(pdf_buffer, format="pdf")
        pdf_buffer.seek(0)
        return pdf_buffer

    # Fonction pour enregistrer le graphique au format JPEG
    def save_as_jpeg(fig):
        jpeg_buffer = BytesIO()
        fig.savefig(jpeg_buffer, format="jpeg", dpi=300)
        jpeg_buffer.seek(0)
        return jpeg_buffer

    # Ajouter des boutons de téléchargement
    st.download_button(
        label="Télécharger le graphique en PDF",
        data=save_as_pdf(fig),
        file_name="graphique_absorption.pdf",
        mime="application/pdf"
    )
    st.download_button(
        label="Télécharger le graphique en JPEG",
        data=save_as_jpeg(fig),
        file_name="graphique_absorption.jpeg",
        mime="image/jpeg"
    )
else:
    st.info("Veuillez importer au moins un fichier Excel contenant des séries valides pour commencer l'analyse.")
