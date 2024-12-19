import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

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
    try:
        extracted_data = []

        for col_start, name_cell, values_range in [
            (0, "A1", "C3:C20"), (4, "E1", "G3:G20"),
            (8, "I1", "K3:K20"), (12, "M1", "O3:O20"),
            (16, "Q1", "S3:S20"), (20, "U1", "W3:W20")
        ]:
            # Lire le nom de la série
            name = df.iloc[0, col_start]

            # Lire les valeurs (en évitant les erreurs non numériques comme #DIV/0!)
            values = pd.to_numeric(
                df.iloc[2:20, col_start + 2], errors="coerce"
            ).values  # `coerce` transforme les erreurs en NaN

            # Vérifie si au moins une valeur est numérique
            if not np.isnan(values).all():
                extracted_data.append({"name": name, "values": values})

        return extracted_data
    except Exception as e:
        st.error(f"Erreur lors de l'extraction des données : {e}")
        return []

# Téléchargement de fichiers
uploaded_files = st.sidebar.file_uploader(
    "Importez vos fichiers Excel (.xls ou .xlsx)",
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

# Afficher les courbes pour chaque fichier et série
if uploaded_files:
    for file in uploaded_files:
        st.subheader(f"Données extraites de : {file.name}")

        # Charger le fichier
        df = load_excel(file)
        if df is None:
            continue

        # Extraire les données valides
        extracted_data = extract_data(df)

        # Si aucune série valide
        if not extracted_data:
            st.warning(f"Aucune série valide trouvée dans {file.name}.")
            continue

        # Génération des graphiques pour chaque série
        fig, ax = plt.subplots(figsize=(10, 6))
        for series in extracted_data:
            ax.plot(frequencies, series["values"], label=series["name"], marker="o")

        # Personnalisation du graphique
        ax.set_title(f"Courbes d'absorption pour {file.name}")
        ax.set_xscale("log")
        ax.set_xticks(frequencies)
        ax.get_xaxis().set_major_formatter(plt.ScalarFormatter())
        ax.set_xlabel("Fréquence (Hz)")
        ax.set_ylabel("Coefficient d'absorption")
        ax.legend()
        ax.grid(True, which="both", linestyle="--", linewidth=0.5)

        # Afficher le graphique dans Streamlit
        st.pyplot(fig)
else:
    st.info("Veuillez importer au moins un fichier Excel pour commencer l'analyse.")
