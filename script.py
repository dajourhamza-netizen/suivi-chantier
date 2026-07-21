import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import os
import io

# --- Configuration de la page ---
st.set_page_config(
    page_title="Générateur de Fiches Chantier",
    page_icon="🖨️",
    layout="centered"
)

# --- Titre et En-tête ---
st.title("🖨️ Application Chantier - Générateur de Fiches")
st.markdown("---")

# --- Dossiers et Chemins ---
dossier_principal = r"C:\Users\lenovo\Desktop\application_chantier"
chemin_excel_defaut = r"C:\Users\lenovo\Desktop\application_chantier\suivi .xlsx"

# Variable pour stocker le dataframe
df = None

# --- Option 1 : Importer un fichier Excel ou utiliser celui par défaut ---
st.sidebar.header("📁 Source des Données")
source_excel = st.sidebar.radio(
    "Choisir la source :",
    ["Fichier par défaut (suivi .xlsx)", "Téléverser un nouveau fichier Excel"]
)

if source_excel == "Fichier par défaut (suivi .xlsx)":
    if os.path.exists(chemin_excel_defaut):
        try:
            df = pd.read_excel(chemin_excel_defaut)
            df = df.fillna("")
            st.sidebar.success("✅ Fichier Excel par défaut chargé avec succès !")
        except Exception as e:
            st.sidebar.error(f"❌ Erreur de lecture : {e}")
    else:
        st.sidebar.error("❌ Fichier Excel introuvable au chemin spécifié.")
else:
    fichier_upload = st.sidebar.file_uploader("Choisissez un fichier Excel", type=["xlsx", "xls"])
    if fichier_upload is not None:
        try:
            df = pd.read_excel(fichier_upload)
            df = df.fillna("")
            st.sidebar.success("✅ Fichier importé avec succès !")
        except Exception as e:
            st.sidebar.error(f"❌ Erreur lors de l'importation : {e}")

# --- Interface Principale ---
if df is not None:
    st.subheader("📋 Sélection de la fiche à générer")

    # Préparation de la liste d'affichage
    liste_choix = []
    for index, ligne in df.iterrows():
        nature = ligne.get('TITRE DE LA NATURE DES TRAVAUX', '???')
        partie = ligne.get("PARTIE D'OUVRAGE", '???')
        situation = ligne.get('SITUATION', '???')
        date = ligne.get('DATE', '')
        
        # Formatage de la date si nécessaire
        if isinstance(date, pd.Timestamp):
            date = date.strftime('%d/%m/%Y')
            
        texte_menu = f"Ligne {index + 1} ({date}) : {nature} | {partie} | {situation}"
        liste_choix.append(texte_menu)

    # Menu déroulant (Selectbox)
    choix_ligne = st.selectbox("Choisissez une ligne dans le tableau :", options=liste_choix)
    
    # Récupération de l'index de la ligne sélectionnée
    index_selection = liste_choix.index(choix_ligne)
    ligne_choisie = df.iloc[index_selection]

    st.markdown("---")
    
    # Bouton de génération
    if st.button("🖨️ Générer la fiche Word", type="primary", use_container_width=True):
        nom_modele = str(ligne_choisie.get('TITRE DE LA NATURE DES TRAVAUX', '')).strip()
        
        if nom_modele == "" or nom_modele == "???":
            st.error("❌ La colonne 'TITRE DE LA NATURE DES TRAVAUX' est vide pour cette ligne !")
        else:
            modele_word = os.path.join(dossier_principal, f"{nom_modele}.docx")
            
            if not os.path.exists(modele_word):
                st.error(f"❌ Le modèle Word '{nom_modele}.docx' est introuvable dans le dossier du chantier !")
            else:
                try:
                    doc = DocxTemplate(modele_word)
                    
                    # Contexte avec les balises Word
                    contexte = {
                        'NATURE': ligne_choisie.get('TITRE DE LA NATURE DES TRAVAUX', ''),
                        'REF': ligne_choisie.get('RÉFÉRENCE DE PROCÉDURE', ''),
                        'PARTIE': ligne_choisie.get("PARTIE D'OUVRAGE", ''),
                        'SITUATION': ligne_choisie.get('SITUATION', ''),
                        'PIECES': ligne_choisie.get('PIÈCES JOINTES', ''),
                        'DATE': ligne_choisie.get('DATE', '')
                    }
                    
                    doc.render(contexte)
                    
                    # Sauvegarde en mémoire pour le téléchargement direct
                    buffer = io.BytesIO()
                    doc.save(buffer)
                    buffer.seek(0)
                    
                    st.success("✅ La fiche a été générée avec succès !")
                    
                    # Bouton de téléchargement
                    nom_fichier_telechargement = f"Fiche_{nom_modele}_Ligne_{index_selection + 1}.docx"
                    st.download_button(
                        label="📥 Télécharger la fiche générée (Word)",
                        data=buffer,
                        file_name=nom_fichier_telechargement,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
                except Exception as e:
                    st.error(f"❌ Une erreur s'est produite lors de la génération : {e}")

else:
    st.info("💡 Veuillez charger un fichier Excel pour commencer.")