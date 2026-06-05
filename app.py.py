import streamlit as st
import pandas as pd
from datetime import datetime
import io

# Configuration de la page
st.set_page_config(page_title="Évaluation CAP Conducteur d'Engins", layout="centered")

st.title("🚜 Grille d'Évaluation - CAP Conducteur d'Engins")
st.write("Formulaire d'évaluation en temps réel sur le terrain.")

# --- 1. INFORMATIONS GÉNÉRALES ---
st.header("📋 Informations Générales")
col1, col2 = st.columns(2)

with col1:
    nom_candidat = st.text_input("Nom & Prénom du Candidat")
    formateur = st.text_input("Nom du Formateur / Évaluateur")

with col2:
    date_eval = st.date_input("Date de l'évaluation", datetime.now())
    engin = st.selectbox(
        "Type d'engin (Recommandation R482)",
        [
            "Cat A : Mini-pelle, Moto-basculeur",
            "Cat B1 : Pelle hydraulique (> 6t)",
            "Cat C1 : Chargeuse, Chargeuse-pelleteuse",
            "Cat D : Bouteur, Trancheuse",
            "Cat E : Tombereau, Décapeuse",
            "Cat F : Chariot de manutention tout-terrain"
        ]
    )

st.divider()

# --- 2. GRILLE DE NOTATION ---
st.header("🎯 Barème et Compétences")
st.write("Notez chaque module de 0 à 20 :")

# Dictionnaire pour stocker les notes
notes = {}

# Module 1 : Prise de poste
st.subheader("1. Prise de poste & Vérifications")
notes["prise_de_poste"] = st.slider("Contrôles visuels, niveaux, EPI, environnement", 0, 20, 10)

# Module 2 : Conduite et Manœuvre
st.subheader("2. Conduite et Maîtrise de l'engin")
notes["conduite"] = st.slider("Souplesse, précision, respect des cycles de travail", 0, 20, 10)

# Module 3 : Sécurité (Critique)
st.subheader("3. Respect des Règles de Sécurité")
notes["securite"] = st.slider("Signalisation, gestes barrières, conduite défensive, piétons", 0, 20, 10)
if notes["securite"] < 10:
    st.warning("⚠️ Attention : Une note inférieure à 10 en sécurité peut être éliminatoire selon le référentiel.")

# Module 4 : Fin de poste
st.subheader("4. Fin de poste & Entretien")
notes["fin_de_poste"] = st.slider("Stationnement sécurisé, nettoyage, rapport d'anomalies", 0, 20, 10)

# Commentaires
commentaires = st.text_area("✍️ Commentaires et observations (points forts / axes d'amélioration)")

st.divider()

# --- 3. CALCUL DU RÉSULTAT ---
st.header("📊 Résultat Final")

# Calcul de la moyenne
moyenne = sum(notes.values()) / len(notes)

st.metric(label="Moyenne Générale", value=f"{moyenne:.2f} / 20")

# Validation automatique
if moyenne >= 10 and notes["securite"] >= 10:
    st.success("🎉 Avis de l'évaluateur : **FAVORABLE** (Aptitude validée)")
    statut = "Favorable"
else:
    st.error("❌ Avis de l'évaluateur : **INSUFFISANT** (À perfectionner)")
    statut = "Insuffisant"

st.divider()

# --- 4. EXPORT DES DONNÉES (VERSION EXCEL) ---
if st.button("💾 Enregistrer l'évaluation"):
    if not nom_candidat:
        st.error("⚠️ Veuillez entrer le nom du candidat avant de télécharger le fichier.")
    else:
        # Création du dictionnaire de données
        donnees_eval = {
            "Candidat": [nom_candidat],
            "Évaluateur": [formateur],
            "Date": [date_eval],
            "Engin": [engin],
            "Prise de poste": [notes["prise_de_poste"]],
            "Conduite": [notes["conduite"]],
            "Sécurité": [notes["securite"]],
            "Fin de poste": [notes["fin_de_poste"]],
            "Moyenne": [moyenne],
            "Statut": [statut],
            "Commentaires": [commentaires]
        }
        
        df = pd.DataFrame(donnees_eval)
        
        # Création d'un tampon mémoire pour le fichier Excel
        buffer = io.BytesIO()
        
        # Écriture du DataFrame dans le fichier Excel avec openpyxl
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Évaluation Candidat")
            
        # Reconfiguration du pointeur du fichier au début
        buffer.seek(0)
        
        # Bouton de téléchargement Excel adapté à l'iPhone
        st.download_button(
            label="📥 Télécharger le bilan au format Excel (.xlsx)",
            data=buffer,
            file_name=f"evaluation_{nom_candidat.replace(' ', '_')}_{date_eval}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )