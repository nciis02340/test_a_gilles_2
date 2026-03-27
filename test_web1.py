import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import io
import re

st.set_page_config(page_title="Extracteur Service (CSV Tab)", page_icon="📝")

# --- FONCTION DE RECHERCHE DANS LE TEXTE BRUT ---
def chercher_dans_texte(texte, mot_cle):
    # Cherche le mot clé et récupère ce qui suit sur la même ligne
    pattern = rf"{mot_cle}[:\s=]*(.*)"
    match = re.search(pattern, texte, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None

st.title("📝 Extracteur Service vers CSV")
st.write("Format de sortie : CSV avec séparateur TABULATION (idéal pour copier-coller dans Excel).")

# --- TA CONFIGURATION ---
CHAMPS_A_GARDER = [
    "serial number", "MODEL NO", "SYMPTOMS", "Problems Reported 1",
    "Problems Reported 2", "Work A", "Work B", "Work C",
    "Work 1", "Work 2", "WORK PERFORMED", "CUSTOMER",
    "location", "Serial #", "DATE RECEIVED"
]

fichiers_uploades = st.file_uploader(
    "Choisir les formulaires PDF", 
    type="pdf", 
    accept_multiple_files=True
)

if fichiers_uploades:
    toutes_les_donnees = []
    
    for fichier in fichiers_uploades:
        try:
            doc = fitz.open(stream=fichier.read(), filetype="pdf")
            infos = {"Nom_Fichier": fichier.name}
            
            # 1. Texte brut pour les PDF sans formulaire
            texte_complet = "\n".join([page.get_text() for page in doc])
            
            # 2. Extraction par formulaire (Widgets)
            for page in doc:
                for widget in page.widgets():
                    nom_brut = widget.field_name
                    if nom_brut.lower() in [c.lower() for c in CHAMPS_A_GARDER]:
                        infos[nom_brut] = widget.field_value

            # 3. Recherche textuelle pour les champs manquants
            for champ in CHAMPS_A_GARDER:
                if champ not in infos or not infos[champ]:
                    valeur_detectee = chercher_dans_texte(texte_complet, champ)
                    if valeur_detectee:
                        infos[champ] = valeur_detectee
            
            toutes_les_donnees.append(infos)
            
        except Exception as e:
            st.error(f"Erreur sur le fichier {fichier.name} : {e}")

    if toutes_les_donnees:
        df_final = pd.DataFrame(toutes_les_donnees)
        
        # Réorganisation des colonnes
        colonnes_presentes = [c for c in df_final.columns if c != "Nom_Fichier"]
        df_final = df_final[["Nom_Fichier"] + colonnes_presentes]

        st.success(f"✅ Analyse terminée sur {len(toutes_les_donnees)} fichiers.")
        st.dataframe(df_final)
        
        # --- PRÉPARATION DU CSV (SÉPARATEUR TABULATION) ---
        # On utilise l'encodage utf-16 pour une compatibilité parfaite avec Excel en double-clic
        csv_tab = df_final.to_csv(index=False, sep='\t', encoding='utf-16')
        
        st.download_button(
            label="📥 Télécharger le fichier CSV (Tabulation)",
            data=csv_tab,
            file_name="extraction_service.csv",
            mime="text/csv"
        )