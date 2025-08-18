"""
Streamlit app for Alba2NWP

Fichier: src/alba2nwp/app.py
But: cette interface est volontairement simple — upload d'un fichier Excel/ODS/XLSB/CSV,
visualisation, conversion en CSV (UTF-8 with BOM) et téléchargement.

Dépendances minimales:
- streamlit
- pandas
- openpyxl
- pyxlsb (si vous voulez lire .xlsb)
- odfpy (si vous voulez lire .ods)

Lancer: `streamlit run src/alba2nwp/app.py`
"""

from pathlib import Path
import streamlit as st
import pandas as pd

from .convert import process_upload


# ----- Configuration -----
ALLOWED_EXTENSIONS = ("xls", "xlsx", "xlsm", "xlsb", "ods", "csv")
DEFAULT_ENCODING = "utf-8-sig"  # BOM helps Excel detect UTF-8 on Windows

# ----- Helpers -----

def get_extension(filename: str) -> str:
    return Path(filename).suffix.lower().lstrip('.')


from typing import Optional, Dict

def read_workbook(uploaded_file) -> Optional[Dict[str, pd.DataFrame]]:
    """Lit le fichier uploadé et renvoie un dict {sheet_name: DataFrame}.
    Pour un fichier CSV, la clé sera 'CSV'.
    Retourne None en cas d'erreur.
    """
    try:
        fname = uploaded_file.name
        ext = get_extension(fname)
        if ext == 'csv':
            # read_csv accepte file-like
            df = pd.read_csv(uploaded_file)
            return {"CSV": df}

        engine = None
        if ext == 'xlsb':
            engine = 'pyxlsb'

        # sheet_name=None lit toutes les feuilles et renvoie un dict
        sheets = pd.read_excel(uploaded_file, sheet_name=None, engine=engine)
        return sheets
    except Exception as e:
        st.error(f"Impossible de lire le fichier : {e}")
        return None


def df_to_bytes_csv(df: pd.DataFrame, encoding: str = DEFAULT_ENCODING, sep: str = ',') -> bytes:
    """Convertit un DataFrame en bytes CSV (prêt pour st.download_button)."""
    csv_str = df.to_csv(index=False, sep=sep)
    return csv_str.encode(encoding)


# ----- UI -----

st.set_page_config(page_title="Alba2NWP | Alba vers NW Publisher", layout="wide")

st.title("Alba2NWP")
st.subheader("Convertisseur de Alba vers NW Publisher")

st.markdown("---")

# Upload
uploaded = st.file_uploader(
    "Choisissez un fichier Excel/ODS (xls, xlsx, xlsm, xlsb, ods) ou CSV",
    type=list(ALLOWED_EXTENSIONS),
    accept_multiple_files=False,
)

st.info(f"Formats autorisés: {', '.join(ALLOWED_EXTENSIONS)}")

sheets = None
current_df = None
selected_sheet = None

if uploaded is not None:
    # Lecture du fichier
    sheets = read_workbook(uploaded)
    if sheets is None:
        st.stop()

    # Si plusieurs feuilles, proposer un selectbox
    if len(sheets) > 1:
        selected_sheet = st.selectbox("Feuille à visualiser:", list(sheets.keys()))
    else:
        selected_sheet = list(sheets.keys())[0]

    # Visualiser le DataFrame
    current_df = sheets[selected_sheet]

    # Réinitialiser l'index pour qu'il commence à 1
    current_df = current_df.reset_index(drop=True)
    current_df.index += 1

    st.markdown("**Aperçu du fichier importé**")
    st.dataframe(current_df, use_container_width=True)


    # Bouton pour convertir
if st.button("Convertir pour NW Publisher (CSV)"):
    st.markdown("---")

    if uploaded is None:
        st.error("Aucun fichier uploadé — veuillez d'abord importer un fichier.")
    else:
        try:
            # Utiliser la fonction process_upload du module de conversion.
            # Si tu importes from alba2nwp.convert_nwp import process_upload -> appeler avec nwp=True
            # Si tu importes from alba2nwp.convert import process_upload -> appeler avec nwp_rules={'target':'nwp'}
            # Exemple pour convert_nwp:
            result = process_upload(uploaded, sheet_name=selected_sheet, sep=',')

            # result est un dict: {'sheets': {name: DataFrame}, 'warnings': [...], 'output': bytes, 'output_name': str}
            # Récupérer le DataFrame transformé (s'il y a une seule feuille on prend la première)
            converted_sheets = result.get('sheets', {})
            if not converted_sheets:
                st.error("Aucune feuille transformée retournée.")
            else:
                # afficher warnings éventuels
                warnings = result.get('warnings', [])
                if warnings:
                    for w in warnings:
                        st.warning(w)

                # prendre la première feuille transformée pour affichage
                first_name = next(iter(converted_sheets.keys()))
                converted_df = converted_sheets[first_name]

                st.markdown("**Tableau converti (après normalisation pour NWP)**")
                st.dataframe(converted_df, use_container_width=True)

                st.success("Conversion terminée!")

                # Télécharger l'output (csv ou zip)
                out_bytes = result.get('output')
                out_name = result.get('output_name') or f"{Path(uploaded.name).stem}.csv"

                # Ajouter suffixe -Converted
                stem = Path(out_name).stem
                suffix = Path(out_name).suffix
                out_name = f"{stem} - Converted{suffix}"

                if out_bytes is not None:
                    st.download_button(
                        label="Télécharger le fichier converti",
                        data=out_bytes,
                        file_name=out_name,
                        mime="application/zip" if out_name.endswith('.zip') else "text/csv",
                        key="download-csv",
                    )
                    st.info(
                        "ℹ️ Pour choisir l’emplacement du fichier, activez dans votre navigateur l’option "
                        "**‘Toujours demander où enregistrer les fichiers’** dans les paramètres de téléchargement."
                    )
                else:
                    st.error("Erreur : aucun fichier à télécharger.")           

        except ValueError as ve:
            # erreurs attendues (p.ex. colonnes requises manquantes)
            st.error(f"Erreur de validation! Vérifiez que le jeu de données provient bien de la plateforme Alba.")
        except Exception as e:
            st.error(f"Erreur pendant la conversion en CSV! Veillez recommencer.")


else:
    st.write("Aucun fichier importé pour l'instant — importez un fichier pour commencer.")

# Footer / aides
st.markdown("---")
st.caption("Alba2NWP — Convertisseur de Alba vers NW Publisher. Pour plus d'options (délimiteur, encodage, multi-feuilles zip), demandez la fonctionnalité.")