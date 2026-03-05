import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import io

st.set_page_config(page_title="Extracteur Filtré", page_icon="🔍")

st.title("🔍 Extracteur PDF Ciblé")
st.write("Extraction spécifique : Serial Number, Customer, Location, Date.")

# --- CONFIGURATION : Modifie les noms ci-dessous pour qu'ils correspondent EXACTEMENT aux noms dans ton PDF ---
CHAMPS_A_GARDER = [
    "serial number",
    "MODEL NO", 
    "SYMPTOMS",
    "Problems Reported 1",
    "Problems Reported 2",
    "Work A", 
    "Work B", 
    "Work C",
    "Work 1",
    "Work 2", 
    "WORK PERFORMED", 
    "CUSTOMER",
    "location",
    "Serial #",
    "DATE RECEIVED"
    
]
# -------------------------------------------------------------------------------------------------------

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
            
            # On commence par noter le nom du fichier
            infos_filtrees = {"Nom_Fichier": fichier.name}
            
            # Extraction des champs
            for page in doc:
                for widget in page.widgets():
                    nom_brut = widget.field_name
                    # On vérifie si ce champ fait partie de notre liste (en minuscules pour éviter les erreurs)
                    if nom_brut.lower() in [c.lower() for c in CHAMPS_A_GARDER]:
                        infos_filtrees[nom_brut] = widget.field_value
            
            toutes_les_donnees.append(infos_filtrees)
            
        except Exception as e:
            st.error(f"Erreur sur le fichier {fichier.name} : {e}")

    if toutes_les_donnees:
        df_final = pd.DataFrame(toutes_les_donnees)
        
        # Optionnel : On réorganise les colonnes pour qu'elles soient dans l'ordre voulu
        # (Seulement si les colonnes existent dans le DataFrame)
        colonnes_presentes = [c for c in df_final.columns if c != "Nom_Fichier"]
        df_final = df_final[["Nom_Fichier"] + colonnes_presentes]

        st.success(f"✅ Analyse terminée sur {len(toutes_les_donnees)} fichiers.")
        st.dataframe(df_final)
        
        # Préparation Excel
        tampon_excel = io.BytesIO()
        with pd.ExcelWriter(tampon_excel, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Télécharger l'Excel filtré",
            data=tampon_excel.getvalue(),
            file_name="extraction_ciblee.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )