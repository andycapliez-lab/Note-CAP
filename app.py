import streamlit as st
import pandas as pd
from datetime import datetime
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuration de la page Streamlit
st.set_page_config(page_title="Feuille de Notation - Conduite d'Engins", layout="centered")

st.title("🚜 Feuille de Notation - Conduite d'Engins")
st.write("Application terrain avec calcul automatique de la note sur 20.")

# --- 1. INFORMATIONS GÉNÉRALES ---
st.header("📋 Informations Générales")
col1, col2 = st.columns(2)

with col1:
    nom_candidat = st.text_input("Nom & Prénom du Candidat")
    formateur = st.text_input("Nom de l'Évaluateur")

with col2:
    date_eval = st.date_input("Date de l'évaluation", datetime.now())
    classe = st.selectbox(
        "Classe du Candidat",
        ["CAP 1 A", "CAP 1 B", "CAP 2 A", "CAP 2 B", "CAP 1AN", "Bac Pro", "BTS", "Adulte / Formation Continue"]
    )
    engin = st.selectbox(
        "Type d'engin",
        [
            "Pelle Hydraulique",
            "Mini-pelle",
            "Chargeuse",
            "Tombereau",
            "Compacteur",
            "Bouteur",
            "Chariot de chantier"
        ]
    )

st.divider()

# --- 2. GRILLE DE NOTATION ---
st.header("🎯 Barème détaillé")

# --- SECTION 1 : PRISE DE POSTE ---
st.subheader("1. Prise de poste (Note /20)")
p1_visuel = st.slider("• Contrôle visuel de l’engin (pneus, godet, etc.)", 0, 5, 5)
p1_niveaux = st.slider("• Vérification des niveaux (huile, carburant)", 0, 5, 5)
p1_securite = st.slider("• Inspection des équipements de sécurité (alarme, ceinture)", 0, 5, 5)
p1_docs = st.slider("• Conformité avec les documents (fiche d’entretien, check-list)", 0, 5, 5)
total_p1 = p1_visuel + p1_niveaux + p1_securite + p1_docs
st.markdown(f"**Sous-total Prise de poste : `{total_p1} / 20`**")

st.divider()

# --- SECTION 2 : CONDUITE ---
st.subheader("2. Conduite de l’engin (Note /40)")
p2_commandes = st.slider("• Maîtrise des commandes et fluidité des mouvements", 0, 10, 7)
p2_securite = st.slider("• Respect des consignes de sécurité (signalisation, zones à risque)", 0, 10, 7)
p2_precision = st.slider("• Précision dans l’exécution des tâches (terrassement, levage, etc.)", 0, 10, 7)
p2_terrain = st.slider("• Adaptabilité aux contraintes du terrain", 0, 10, 7)
total_p2 = p2_commandes + p2_securite + p2_precision + p2_terrain
st.markdown(f"**Sous-total Conduite : `{total_p2} / 40`**")

st.divider()

# --- SECTION 3 : FIN DE POSTE ---
st.subheader("3. Fin de poste (Note /20)")
p3_nettoyage = st.slider("• Nettoyage de l’engin et rangement des outils", 0, 5, 5)
p3_anomalies = st.slider("• Rapport des anomalies ou pannes", 0, 5, 5)
p3_stationnement = st.slider("• Stationnement en sécurité (frein, zone dédiée)", 0, 5, 5)
p3_fiche = st.slider("• Mise à jour des documents (fiche de poste)", 0, 5, 5)
total_p3 = p3_nettoyage + p3_anomalies + p3_stationnement + p3_fiche
st.markdown(f"**Sous-total Fin de poste : `{total_p3} / 20`**")

st.divider()

# --- SECTION 4 : COMPORTEMENT ---
st.subheader("4. Comportement général et professionnalisme (Note /20)")
p4_consignes = st.slider("• Respect des consignes orales et écrites", 0, 10, 8)
p4_attitude = st.slider("• Attitude professionnelle et communication", 0, 10, 8)
total_p4 = p4_consignes + p4_attitude
st.markdown(f"**Sous-total Comportement : `{total_p4} / 20`**")

st.divider()

# --- 3. CALCUL DU RÉSULTAT FINAL SUR 20 ---
st.header("📊 Résultat Final")
note_totale_100 = total_p1 + total_p2 + total_p3 + total_p4

# NOUVEAU : Calcul de la note sur 20
note_sur_20 = note_totale_100 / 5

# Affichage en grand de la note sur 20
st.metric(label="NOTE FINALE", value=f"{note_sur_20} / 20")

if note_sur_20 >= 10:
    st.success("🎉 **Validation réussie (Moyenne atteinte)**")
else:
    st.error("⚠️ **Insuffisant (En dessous de la moyenne)**")

commentaires = st.text_area("✍️ Appréciation générale (Observations)")

st.divider()

# --- 4. EXPORT EXCEL ULTRA-STYLISÉ ---
if not nom_candidat:
    st.info("💡 Veuillez entrer le nom du candidat pour débloquer le bouton de téléchargement Excel.")
else:
    donnees_eval = {
        "Candidat": [nom_candidat],
        "Classe": [classe],
        "Évaluateur": [formateur],
        "Date": [str(date_eval)],
        "Engin": [engin],
        "P1_Visuel (/5)": [p1_visuel],
        "P1_Niveaux (/5)": [p1_niveaux],
        "P1_Sécurité (/5)": [p1_securite],
        "P1_Docs (/5)": [p1_docs],
        "TOTAL_Prise_Poste (/20)": [total_p1],
        "P2_Commandes (/10)": [p2_commandes],
        "P2_Sécurité (/10)": [p2_securite],
        "P2_Précision (/10)": [p2_precision],
        "P2_Terrain (/10)": [p2_terrain],
        "TOTAL_Conduite (/40)": [total_p2],
        "P3_Nettoyage (/5)": [p3_nettoyage],
        "P3_Anomalies (/5)": [p3_anomalies],
        "P3_Stationnement (/5)": [p3_stationnement],
        "P3_Fiche (/5)": [p3_fiche],
        "TOTAL_Fin_Poste (/20)": [total_p3],
        "P4_Consignes (/10)": [p4_consignes],
        "P4_Attitude (/10)": [p4_attitude],
        "TOTAL_Comportement (/20)": [total_p4],
        "NOTE BRUTE (/100)": [note_totale_100],
        "NOTE FINALE (/20)": [note_sur_20], # La note finale est bien exportée sur 20 !
        "Appréciation générale": [commentaires]
    }
    
    df = pd.DataFrame(donnees_eval)
    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Évaluation Conduite")
        
        workbook = writer.book
        worksheet = writer.sheets["Évaluation Conduite"]
        
        font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        font_body = Font(name="Arial", size=10)
        fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
        )
        
        for cell in worksheet[1]:
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        
        for cell in worksheet[2]:
            cell.font = font_body
            cell.border = thin_border
            if isinstance(cell.value, (int, float)):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(vertical="center")
        
        # Coloration de la case "NOTE FINALE (/20)" (Colonne 25 ou 'Y')
        cell_note = worksheet.cell(row=2, column=25)
        cell_note.font = Font(name="Arial", size=12, bold=True)
        if note_sur_20 >= 10:
            cell_note.fill = PatternFill(start_color="C6EFCE", fill_type="solid") # Vert
        else:
            cell_note.fill = PatternFill(start_color="FFC7CE", fill_type="solid") # Rouge

        worksheet.row_dimensions[1].height = 28
        for col in worksheet.columns:
            col_letter = get_column_letter(col[0].column)
            if col[0].column in [1, 2, 3, 4, 5, 26]:
                worksheet.column_dimensions[col_letter].width = 24
            else:
                worksheet.column_dimensions[col_letter].width = 14
                
    buffer.seek(0)
    
    st.download_button(
        label="💾 Générer et Partager la Note (Excel)",
        data=buffer,
        file_name=f"Notation_{classe.replace(' ', '_')}_{nom_candidat.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
