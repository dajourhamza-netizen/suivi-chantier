import streamlit as st
import pandas as pd
from datetime import datetime
import os
import re

# Imports OpenPyXL pour le style de tableau Excel
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Suivi de Chantier", page_icon="🏗️", layout="wide")

FILE_PATH = "suivi .xlsx"

# ==========================================
# COLONNES STANDARDS ET LIAISONS
# ==========================================
COLUMNS_TEMPLATE = [
    "DATE", "TITRE DE LA NATURE DES TRAVAUX", "PARTIE D'OUVRAGE", 
    "SITUATION", "ACTIVITÉ RÉALISÉE", "ÉSSAI/ CONTRÔLE RÉALISÉE", 
    "RÉFÉRENCE DE PROCÉDURE", "PIÈCES JOINTES"
]

LIAISONS = {
    "ARASE DE PST": {"procedure": "TER-PEX-05-00", "pieces": "* Fiche de suivi de la PST\n* Fiche de réception topographique\n* PVs laboratoire"},
    "ARASE DE TERRASSEMENT": {"procedure": "TER-PEX-03-00", "pieces": "* Fiche de contrôle des déblais\n* Fiche de réception topographique\n* PVs laboratoire"},
    "ASSISE DE REMBLAIS PURGE": {"procedure": "TER-PEX-04-00", "pieces": "* Fiche de réception de l'assise des remblais\n* Fiche de réception topographique\n* Fiche d'identification de la purge\n* PVs laboratoire"},
    "ASSISE DE REMBLAIS": {"procedure": "TER-PEX-04-00", "pieces": "* Fiche de réception de l'assise des remblais\n* Fiche de réception topographique\n* PVs laboratoire"},
    "ASSISE DE REMBLAIS CDF": {"procedure": "TER-PEX-04-00", "pieces": "* Fiche de réception de l'assise des remblais\n* Fiche de réception topographique\n* PVs laboratoire"},
    "ASSISE DE REMBLAIS CONTIGUS": {"procedure": "OVA-PEX-16-00", "pieces": "* Fiche de suivi des remblais contigus\n* Fiche de contrôle des remblais contigus\n* PVs laboratoire\n* Fiche de réception topographique"},
    "ASSISE DE REMBLAI DE FOUILLE": {"procedure": "OVA-PEX-04-00", "pieces": "* Fiche de suivi et de contrôle des fouilles et remblaiement de fouilles\n* PVs laboratoire"},
    "ASSISE DE REMBLAIS RENFORCE": {"procedure": "TER-PEX-13-00", "pieces": "* PV Manifold\n* PVs laboratoire\n* Fiche de réception topographique\n* Fiche de réception assise remblai renforcé"},
    "ASSISE DRAINANTE": {"procedure": "TER-PEX-13-00", "pieces": "* Fiche de réception topographique\n* PVs laboratoire\n* Fiche de contrôle de l'assise drainante"},
    "COUCHE DE FORME": {"procedure": "TER-PEX-09-00", "pieces": "* Fiche de suivi et de contrôle de la CDF\n* Fiche de réception topographique\n* PVs laboratoire"},
    "DÉCAPAGE": {"procedure": "TER-PEX-02-00", "pieces": "* Fiche de suivi et de contrôle du décapage\n* Fiche des sections à décaper\n* Fiche de réception topographique"},
    "DEGAGEMENT D'EMPRISE": {"procedure": "TER-PEX-01-00", "pieces": "* Fiche de suivi et de contrôle du dégagement des emprises\n* Fiche de réception topographique\n* Constat dégagement d'emprise"},
    "REMBLAIS": {"procedure": "TER-PEX-04-00", "pieces": "* Fiche de suivi et de contrôle des remblais\n* PVs laboratoire"},
    "REMBLAIS CDF": {"procedure": "TER-PEX-04-00", "pieces": "* Fiche de suivi et de contrôle des remblais\n* PVs laboratoire"},
    "REMBLAIS CONTIGUS": {"procedure": "OVA-PEX-16-00", "pieces": "* Fiche de suivi des remblais contigus\n* Fiche de contrôle des remblais contigus\n* PVs laboratoire\n* Fiche de réception topographique"},
    "REMBLAIS DE FOUILLE": {"procedure": "OVA-PEX-04-00", "pieces": "* Fiche de suivi et de contrôle des fouilles et remblaiement de fouilles\n* PVs laboratoire"},
    "REMBLAIS DE FOUILLS CDF": {"procedure": "OVA-PEX-04-00", "pieces": "* Fiche de suivi et de contrôle des fouilles et remblaiement de fouilles\n* PVs laboratoire"},
    "REMBLAIS RENFORCE": {"procedure": "TER-PEX-13-00", "pieces": "* Fiche de suivi des remblais renforcé\n* Fiche de contrôle des armatures Geostrap\n* Fiche de réception de pose des ecailles\n* PVs laboratoire"},
    "REMBLAIS PST": {"procedure": "TER-PEX-05-00", "pieces": "* Fiche de suivi et de contrôle des remblais PST\n* PVs laboratoire"}
}

# ==========================================
# FONCTIONS UTILITAIRES & STYLE EXCEL
# ==========================================
def save_to_excel_with_formatting(df_to_save, sheet_name):
    """Sauvegarde les données et applique le style officiel 'Tableau Excel' Orange"""
    try:
        # Créer le fichier s'il n'existe pas encore
        if not os.path.exists(FILE_PATH):
            with pd.ExcelWriter(FILE_PATH, engine="openpyxl") as writer:
                df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)

        with pd.ExcelWriter(FILE_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            # 1. Écrire les données brutes
            df_to_save.to_excel(writer, sheet_name=sheet_name, index=False)
            
            worksheet = writer.sheets[sheet_name]
            
            # 2. Ajuster automatiquement la largeur des colonnes
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    val_str = str(cell.value or "")
                    if len(val_str) > max_length:
                        max_length = len(val_str)
                worksheet.column_dimensions[column_letter].width = min(max_length + 4, 60)
            
            # 3. Créer le style 'Tableau Excel' officiel
            num_rows = max(len(df_to_save) + 1, 2)  # Au moins l'en-tête et une ligne
            num_cols = len(df_to_save.columns)
            
            if num_cols > 0:
                end_col_letter = get_column_letter(num_cols)
                table_ref = f"A1:{end_col_letter}{num_rows}"
                
                # Nettoyer le nom du chantier pour créer un identifiant de tableau valide
                clean_name = re.sub(r'\W+', '_', sheet_name)
                table_name = f"Tableau_{clean_name}"
                
                # Effacer d'éventuels anciens tableaux
                worksheet._tables.clear()
                
                tab = Table(displayName=table_name, ref=table_ref)
                
                # TableStyleMedium3 = Style Orange avec en-têtes colorés et lignes alternées
                style = TableStyleInfo(
                    name="TableStyleMedium3", 
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False
                )
                tab.tableStyleInfo = style
                worksheet.add_table(tab)

        return True, "✅ Fichier mis à jour avec le style Tableau Excel !"
    except PermissionError:
        return False, "❌ ERREUR : Le fichier Excel est OUVERT. Fermez Excel et réessayez !"
    except Exception as e:
        return False, f"❌ Erreur lors de la sauvegarde : {e}"

def init_excel():
    """Initialise le fichier avec une feuille principale si absent"""
    if not os.path.exists(FILE_PATH):
        df_empty = pd.DataFrame(columns=COLUMNS_TEMPLATE)
        save_to_excel_with_formatting(df_empty, "Chantier Principal")

def get_sheet_names():
    """Récupère la liste des chantiers (feuilles)"""
    if os.path.exists(FILE_PATH):
        try:
            xls = pd.ExcelFile(FILE_PATH)
            return xls.sheet_names
        except:
            return ["Chantier Principal"]
    return ["Chantier Principal"]

def load_data(sheet_name):
    """Charge les données du chantier sélectionné"""
    try:
        df = pd.read_excel(FILE_PATH, sheet_name=sheet_name)
        if 'DATE' in df.columns:
            df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce').dt.strftime('%d/%m/%Y')
        return df
    except Exception:
        return pd.DataFrame(columns=COLUMNS_TEMPLATE)

# Initialisation
init_excel()
chantiers_disponibles = get_sheet_names()

# ==========================================
# BARRE LATÉRALE : GESTION DES CHANTIERS
# ==========================================
with st.sidebar:
    st.header("🏢 Gestion des Chantiers")
    
    # 1. Sélection du chantier actif
    st.subheader("📌 Chantier Actif")
    chantier_actif = st.selectbox("Choisissez le chantier :", options=chantiers_disponibles)
    
    st.markdown("---")
    
    # 2. Créer un nouveau chantier
    st.subheader("➕ Ajouter un Chantier")
    nouveau_chantier = st.text_input("Nom du nouveau chantier :", placeholder="Ex: Tronçon B, Pont 2...")
    
    if st.button("➕ Créer la feuille", use_container_width=True):
        if not nouveau_chantier.strip():
            st.warning("Veuillez saisir un nom valide.")
        elif nouveau_chantier.strip() in chantiers_disponibles:
            st.warning("Ce chantier existe déjà !")
        else:
            df_empty = pd.DataFrame(columns=COLUMNS_TEMPLATE)
            success, msg = save_to_excel_with_formatting(df_empty, nouveau_chantier.strip())
            if success:
                st.success(f"✅ Chantier '{nouveau_chantier.strip()}' créé !")
                st.rerun()
            else:
                st.error(msg)

if not chantier_actif:
    st.stop()

# ==========================================
# PAGE PRINCIPALE
# ==========================================
st.title(f"🏗️ Suivi de Chantier : **{chantier_actif}**")
st.markdown("---")

df = load_data(chantier_actif)
tab1, tab2 = st.tabs(["📝 Nouvelle Saisie", "🔍 Historique & Filtres Excel"])

# ==========================================
# ONGLET 1 : FORMULAIRE DE SAISIE
# ==========================================
with tab1:
    st.subheader(f"📝 Saisie pour {chantier_actif}")
    
    col1, col2 = st.columns(2)
    
    with col1:
        date_saisie = st.date_input("🗓️ Date", value=datetime.today(), format="DD/MM/YYYY")
        
        nature_selectionnee = st.selectbox("📌 TITRE DE LA NATURE DES TRAVAUX", options=list(LIAISONS.keys()))
        info_liaison = LIAISONS.get(nature_selectionnee, {"procedure": "", "pieces": ""})
        
        partie_ouvrage = st.text_input("🧱 PARTIE D'OUVRAGE", placeholder="Ex: BRETELLE - A, BRANCHE 2...")
        situation = st.text_input("📍 SITUATION / PK", placeholder="Ex: DU PK 0+020 AU PK 0+300")
        
    with col2:
        activite = st.text_area("🚜 ACTIVITÉ RÉALISÉE", height=80, placeholder="Description des travaux...")
        
        essais_options = [
            "Aucun", "ESSAI À LA PLAQUE", "DENSITÉ", "ESSAI À LA PLAQUE + DENSITÉ",
            "TENEUR EN EAU", "IDENTIFICATION DES MATERIAUX", 
            "PRELEVEMENT AVANT COMPACTAGE", "PRELEVEMENT APRES COMPACTAGE", "PRELEVEMENT"
        ]
        essai = st.selectbox("🧪 ÉSSAI / CONTRÔLE RÉALISÉE", options=essais_options)
        procedure = st.text_input("📑 RÉFÉRENCE DE PROCÉDURE (Auto)", value=info_liaison["procedure"])
        pieces_jointes = st.text_area("📎 PIÈCES JOINTES (Auto)", value=info_liaison["pieces"], height=120)

    st.markdown("---")
    if st.button(f"💾 Enregistrer dans {chantier_actif}", type="primary", use_container_width=True):
        new_entry = {
            "DATE": date_saisie.strftime('%d/%m/%Y'),
            "TITRE DE LA NATURE DES TRAVAUX": nature_selectionnee,
            "PARTIE D'OUVRAGE": partie_ouvrage,
            "SITUATION": situation,
            "ACTIVITÉ RÉALISÉE": activite,
            "ÉSSAI/ CONTRÔLE RÉALISÉE": None if essai == "Aucun" else essai,
            "RÉFÉRENCE DE PROCÉDURE": procedure,
            "PIÈCES JOINTES": pieces_jointes
        }
        
        df_curr = pd.read_excel(FILE_PATH, sheet_name=chantier_actif)
        df_updated = pd.concat([df_curr, pd.DataFrame([new_entry])], ignore_index=True)
        
        success, msg = save_to_excel_with_formatting(df_updated, chantier_actif)
        if success:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

# ==========================================
# ONGLET 2 : FILTRES TYPE EXCEL & TABLEAU
# ==========================================
with tab2:
    st.subheader(f"📊 Données & Filtres : {chantier_actif}")
    
    if not df.empty:
        filtered_df = df.copy()

        with st.expander("🔻 **Panneau de Filtres Excel (Recherche, Date & Colonnes)**", expanded=True):
            # 1. Recherche rapide globale
            search_query = st.text_input("⚡ **Recherche rapide**", placeholder="Tapez pour chercher dans tout le tableau...")
            if search_query:
                mask = filtered_df.apply(lambda col: col.astype(str).str.contains(search_query, case=False, na=False, regex=False)).any(axis=1)
                filtered_df = filtered_df[mask]

            st.markdown("---")
            
            # 2. Filtre par date (Année / Mois / Jour)
            if 'DATE' in df.columns:
                st.markdown("##### 🗓️ **Filtre par Date (Année / Mois / Jour) :**")
                date_cols = st.columns(3)
                temp_date = pd.to_datetime(filtered_df['DATE'], format='%d/%m/%Y', errors='coerce')
                
                with date_cols[0]:
                    annees_uniques = sorted(temp_date.dt.year.dropna().unique().astype(int).tolist())
                    annees_selectionnees = st.multiselect("Année", options=annees_uniques, placeholder="Toutes")
                with date_cols[1]:
                    mois_uniques = sorted(temp_date.dt.month.dropna().unique().astype(int).tolist())
                    mois_noms = [str(m).zfill(2) for m in mois_uniques]
                    mois_selectionnes = st.multiselect("Mois", options=mois_noms, placeholder="Tous")
                with date_cols[2]:
                    jours_uniques = sorted(temp_date.dt.day.dropna().unique().astype(int).tolist())
                    jours_noms = [str(d).zfill(2) for d in jours_uniques]
                    jours_selectionnes = st.multiselect("Jour", options=jours_noms, placeholder="Tous")

                if annees_selectionnees:
                    filtered_df = filtered_df[temp_date.dt.year.isin(annees_selectionnees)]
                    temp_date = pd.to_datetime(filtered_df['DATE'], format='%d/%m/%Y', errors='coerce')
                if mois_selectionnes:
                    mois_int = [int(m) for m in mois_selectionnes]
                    filtered_df = filtered_df[temp_date.dt.month.isin(mois_int)]
                    temp_date = pd.to_datetime(filtered_df['DATE'], format='%d/%m/%Y', errors='coerce')
                if jours_selectionnes:
                    jours_int = [int(j) for j in jours_selectionnes]
                    filtered_df = filtered_df[temp_date.dt.day.isin(jours_int)]
            
            st.markdown("---")
            st.markdown("##### 📌 **Sélection par colonne :**")
            
            # 3. Filtres par colonnes
            filter_cols = st.columns(2)
            col_idx = 0
            for col_name in df.columns:
                if col_name == 'DATE':
                    continue
                unique_vals = sorted([str(v) for v in df[col_name].dropna().unique()])
                with filter_cols[col_idx % 2]:
                    selected = st.multiselect(
                        f"🔹 **{col_name}**", 
                        options=unique_vals, 
                        placeholder="Toutes les valeurs", 
                        key=f"excel_filter_{col_name}"
                    )
                    if selected:
                        filtered_df = filtered_df[filtered_df[col_name].astype(str).isin(selected)]
                col_idx += 1

        # Affichage du compteur de lignes
        st.caption(f"📊 **{len(filtered_df)}** ligne(s) affichée(s) sur **{len(df)}** au total")

        # Éditeur de données
        edited_df = st.data_editor(
            filtered_df,
            num_rows="dynamic",
            use_container_width=True,
            height=520,
            key="editeur_chantier_final"
        )
        
        st.markdown("---")
        if st.button(f"💾 Enregistrer les modifications dans Excel ({chantier_actif})", type="primary", use_container_width=True):
            df.update(edited_df)
            success, msg = save_to_excel_with_formatting(df, chantier_actif)
            if success:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)
    else:
        st.info("Ce chantier ne contient aucune donnée pour le moment. Allez dans l'onglet 'Nouvelle Saisie' pour commencer.")