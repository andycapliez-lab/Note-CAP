import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# Configuration de la page
st.set_page_config(page_title="Évaluation CAP Conducteur d'Engins", layout="centered")

st.title("🚜 Grille d'Évaluation - CAP Conducteur d'Engins")
st.write("Formulaire terrain connecté directement à Google Sheets.")

# --- CONNEXION SECURISÉE À GOOGLE SHEETS ---
# On crée la connexion avec le tableau en ligne
conn = st.connection("gsheets", type=GSheetsConnection)

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
notes = {}

st.subheader("1. Prise de poste & Vérifications")
notes["prise_de_poste"] = st.slider("Contrôles visuels, niveaux, EPI, environnement", 0, 20, 10)

st.subheader("2. Conduite et Maîtrise de l'engin")
notes["conduite"] = st.slider("Souplesse, précision, respect des cycles de travail", 0, 20, 10)

st.subheader("3. Respect des Règles de Sécurité")
notes["securite"] = st.slider("Signalisation, gestes barrières, conduite défensive, piétons", 0, 20, 10)
if notes["securite"] < 10:
    st.warning("⚠️ Attention : Note en sécurité inférieure à 10.")

st.subheader("4. Fin de poste & Entretien")
notes["fin_de_poste"] = st.slider("Stationnement sécurisé, nettoyage, rapport d'anomalies", 0, 20, 10)

commentaires = st.text_area("✍️ Commentaires et observations")

st.divider()

# --- 3. CALCUL DU RÉSULTAT ---
st.header("📊 Résultat Final")
moyenne = sum(notes.values()) / len(notes)
st.metric(label="Moyenne Générale", value=f"{moyenne:.2f} / 20")

if moyenne >= 10 and notes["securite"] >= 10:
    st.success("🎉 Avis : **FAVORABLE**")
    statut = "Favorable"
else:
    st.error("❌ Avis : **INSUFFISANT**")
    statut = "Insuffisant"

st.divider()

# --- 4. ENVOI EN DIRECT SUR GOOGLE SHEETS ---
if st.button("🚀 Valider et Envoyer l'évaluation"):
    if not nom_candidat:
        st.error("⚠️ Veuillez entrer le nom du candidat avant d'envoyer.")
    else:
        with st.spinner("Envoi des données sur Google Sheets..."):
            try:
                # 1. Lire les données existantes du Google Sheet pour ne pas les effacer
                donnees_existantes = conn.read()
                
                # 2. Préparer la nouvelle ligne
                nouvelle_ligne = pd.DataFrame([{
                    "Candidat": nom_candidat,
                    "Évaluateur": formateur,
                    "Date": str(date_eval),
                    "Engin": engin,
                    "Prise de poste": notes["prise_de_poste"],
                    "Conduite": notes["conduite"],
                    "Sécurité": notes["securite"],
                    "Fin de poste": notes["fin_de_poste"],
                    "Moyenne": round(moyenne, 2),
                    "Statut": statut,
                    "Commentaires": commentaires
                }])
                
                # 3. Fusionner l'ancien tableau avec la nouvelle ligne
                tableau_mis_a_jour = pd.concat([donnees_existantes, nouvelle_ligne], ignore_index=True)
                
                # 4. Renvoyer le tout sur Google Sheets
                conn.update(data=tableau_mis_a_jour)
                
                st.success(f"✅ L'évaluation de {nom_candidat} a bien été ajoutée au Google Sheet ! Vous pouvez quitter la page.")
                st.balloons()
                
            except Exception as e:
                st.error(f"Une erreur est survenue lors de la connexion à Google Sheets : {e}")