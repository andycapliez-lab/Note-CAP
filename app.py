import streamlit as st
import pandas as pd
from datetime import datetime
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuration de la page Streamlit
st.set_page_config(page_title="Évaluation CAP Conducteur d'Engins", layout="centered")

st.title("🚜 Grille d'Évaluation - CAP Conducteur d'Engins")
st.write("Formulaire terrain avec exportation Excel stylisée.")

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
    classe = st.selectbox(
        "Classe du Candidat",
        ["CAP 1 A", "CAP 1 B", "CAP 2 A", "CAP 2 B", "CAP 1AN"]
    )

st.divider()

# --- 2. GRILLE DE NOTATION ---
st.header("🎯 Barème et Compétences")
st.write("Notez chaque module de 0 à 20 :")

notes = {}

st.subheader("1. Prise de poste & Vérifications")
notes["prise_de_poste"] = st.slider("Contrôles visuels, niveaux, EPI, environnement", 0, 20, 10)

st.subheader("2. Conduite et Maîtrise de l'engin")
notes["conduite"] = st.slider("Souplesse, précision, respect des cycles de travail", 0, 20, 10)

st.subheader("3. Respect des Règles de Sécurité")
notes["securite"] = st.slider("Signalisation, conduite défensive, piétons", 0, 20, 10)
if notes["securite"] < 10:
    st.warning("⚠️ Attention : Une note inférieure à 10 en sécurité est éliminatoire.")

st.subheader("4. Fin de poste & Entretien")
notes["fin_de_poste"] = st.slider("Stationnement sécurisé, nettoyage, rapport d'anomalies", 0, 20, 10)

commentaires = st.text_area("✍️ Commentaires et observations (points forts / axes d'amélioration)")

st.divider()

# --- 3. CALCUL DU RÉSULTAT ---
st.header("📊 Résultat Final")
moyenne = sum(notes.values()) / len(notes)
st.metric(label="Moyenne Générale", value=f"{moyenne:.2f} / 20")

if moyenne >= 10 and notes["securite"] >= 10:
    st.success("🎉 Avis de l'évaluateur : **FAVORABLE** (Aptitude validée)")
    statut = "Favorable"
else:
    st.error("❌ Avis de l'évaluateur : **INSUFFISANT** (À perfectionner)")
    statut = "Insuffisant"

st.divider()

# --- 4. EXPORT EXCEL SOIGNÉ ---
if not nom_candidat:
    st.info("💡 Veuillez entrer le nom du candidat pour débloquer le bouton de téléchargement Excel.")
else:
    # Préparation des données
    donnees_eval = {
        "Candidat": [nom_candidat],
        "Classe": [classe],
        "Évaluateur": [formateur],
        "Date": [str(date_eval)],
        "Engin": [engin],
        "Prise de poste": [notes["prise_de_poste"]],
        "Conduite": [notes["conduite"]],
        "Sécurité": [notes["securite"]],
        "Fin de poste": [notes["fin_de_poste"]],
        "Moyenne": [round(moyenne, 2)],
        "Statut": [statut],
        "Commentaires": [commentaires]
    }
    
    df = pd.DataFrame(donnees_eval)
    
    buffer = io.BytesIO()
    
    # Écriture Excel et mise en forme avancée
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Évaluation")
        
        # Récupération de la feuille de calcul générée
        workbook = writer.book
        worksheet = writer.sheets["Évaluation"]
        
        # Définition des styles (Polices, Couleurs, Bordures)
        font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        font_body = Font(name="Arial", size=11, bold=False)
        fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Bleu pro
        
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        
        # 1. Styliser la ligne d'en-tête (Ligne 1)
        for cell in worksheet[1]:
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        
        # 2. Styliser les données (Ligne 2)
        for cell in worksheet[2]:
            cell.font = font_body
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")
            
            # Centrer les notes, les dates et le statut
            if cell.column_letter in ['B', 'D', 'F', 'G', 'H', 'I', 'J', 'K']:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
        # 3. Alerte Couleur sur la case Statut (Colonne K - 11ème colonne)
        cell_statut = worksheet.cell(row=2, column=11)
        if statut == "Favorable":
            cell_statut.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Vert clair
            cell_statut.font = Font(name="Arial", size=11, bold=True, color="375623")
        else:
            cell_statut.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # Rouge/Orange clair
            cell_statut.font = Font(name="Arial", size=11, bold=True, color="C65911")

        # 4. Ajuster automatiquement la largeur des colonnes
        for col in worksheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            # On donne un peu d'espace supplémentaire (marge de 4 caractères)
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    buffer.seek(0)
    
    # Bouton de téléchargement
    st.download_button(
        label="💾 Générer et Partager le fichier Excel Stylisé (.xlsx)",
        data=buffer,
        file_name=f"evaluation_{classe.replace(' ', '_')}_{nom_candidat.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
