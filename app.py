import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Convertisseur de Rendez-vous", layout="centered")
st.title("🏥 Convertisseur de Fichiers Rendez-vous")
st.write("Déposez votre export brut pour générer un fichier Excel structuré.")

fichier_importe = st.file_uploader("Choisir le fichier brut (.xls, .csv, .txt)", type=["xls", "csv", "txt"])

if fichier_importe is not None:
    try:
        if fichier_importe.name.endswith('.xls'):
            try:
                df_brut = pd.read_excel(fichier_importe, header=None)
            except:
                df_brut = pd.read_csv(fichier_importe, header=None, sep=',', skip_blank_lines=False)
        else:
            df_brut = pd.read_csv(fichier_importe, header=None, sep=',', skip_blank_lines=False)

        donnees_propres = []
        patient_courant = {}

        for index, row in df_brut.iterrows():
            # Conversion forcée et sécurisée de chaque élément en texte (String)
            ligne_str = [str(x).strip() for x in row.tolist()]
            ligne_jointure = " ".join(ligne_str)

            if "Index patient" in ligne_jointure:
                if patient_courant and "Index patient" in patient_courant:
                    donnees_propres.append(patient_courant)
                patient_courant = {}
                valeurs = [x for x in ligne_str if x not in ['', 'nan', 'None'] and "Index patient" not in x and not x.startswith('*')]
                if valeurs:
                    try:
                        patient_courant["Index patient"] = int(float(valeurs[0]))
                    except:
                        patient_courant["Index patient"] = valeurs[0]

            elif "Nom patient" in ligne_jointure:
                valeurs = [x for x in ligne_str if x not in ['', 'nan', 'None'] and "Nom patient" not in x]
                if valeurs:
                    patient_courant["Nom patient"] = valeurs[0]

            elif "Rendez-vous" in ligne_jointure and "Rendez-vous N°:" not in ligne_jointure:
                valeurs = [x for x in ligne_str if x not in ['', 'nan', 'None'] and "Rendez-vous" not in x]
                if valeurs:
                    try:
                        date_excel = float(valeurs[0])
                        date_timestamp = pd.to_datetime(date_excel, unit='D', origin='1899-12-30')
                        patient_courant["Rendez-vous"] = date_timestamp.strftime('%d/%m/%Y %H:%M:%S')
                    except:
                        patient_courant["Rendez-vous"] = valeurs[0]

            elif "GYNC" in ligne_jointure or "CONSULTATION DU SPÉCIALISTE" in ligne_jointure:
                valeurs = [x for x in ligne_str if "CONSULTATION" in x or "GYNC" in x]
                if valeurs:
                    patient_courant["Consultation"] = valeurs[0]

            elif "CS :" in ligne_jointure or "Activité" in ligne_jointure:
                valeurs = [x for x in ligne_str if "CS :" in x]
                if valeurs:
                    patient_courant["Activité"] = valeurs[0]

            elif "GYNEC :" in ligne_jointure or "Specialité" in ligne_jointure:
                valeurs = [x for x in ligne_str if "GYNEC :" in x]
                if valeurs:
                    patient_courant["Specialité"] = valeurs[0]

            elif "Rendez-vous N°:" in ligne_jointure:
                cellule_rdv = [x for x in ligne_str if "Rendez-vous N°:" in x]
                if cellule_rdv:
                    texte = cellule_rdv[0]
                    match = re.search(r"pris le\s*:\s*([\d/]+\s+[\d:]+)", texte)
                    if match:
                        patient_courant["pris le"] = match.group(1)

        if patient_courant:
            donnees_propres.append(patient_courant)

        colonnes_cibles = ["Index patient", "Nom patient", "Rendez-vous", "Consultation", "Activité", "Specialité", "pris le"]
        df_final = pd.DataFrame(donnees_propres, columns=colonnes_cibles)

        st.success(f"Extraction réussie ! {len(df_final)} rendez-vous trouvés.")
        st.dataframe(df_final.head(10))

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Rendez-vous')
        
        st.download_button(
            label="📥 Télécharger le fichier Excel Converti",
            data=buffer.getvalue(),
            file_name="Rendez_vous_Formate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Une erreur est survenue lors du traitement : {e}")
