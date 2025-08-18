"""
convert.py

Conversion et transformation d'un DataFrame / workbook uploadé en CSV compatible NWP
selon la spécification fournie par l'utilisateur.
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path
from typing import Dict, Optional, Any, List

import pandas as pd

from .utils import _build_col_map, _format_phone_to_north_american, _get_col, _get_extension, _norm_col_name, _parse_address_field, _title_case_safe

# ------------------------------------------------------------------
# Colonnes attendues (originales) et colonnes cibles (transformées)
# ------------------------------------------------------------------
REQUIRED_ORIGINAL_COLUMNS = [
    "Address_ID", "Territory_ID", "Language", "Status", "Name", "Suite", "Address",
    "City", "Province", "Postal_code", "Country", "Latitude", "Longitude",
    "Telephone", "Owner", "Notes", "Notes_private", "Account", "Created",
    "Modified", "Contacted", "Geocoded", "Territory_number", "Territory_description",
]

TARGET_COLUMNS = [
    "TerritoryID", "TerritoryNumber", "CategoryCode", "Category", "TerritoryAddressID",
    "ApartmentNumber", "Number", "Street", "Suburb", "PostalCode", "State", "Name",
    "Phone", "Type", "Status", "NotHomeAttempt", "Date1", "Date2", "Date3", "Date4",
    "Date5", "Language", "Latitude", "Longitude", "Notes", "NotesFromPublisher",
]


def read_workbook_from_filelike(uploaded_file) -> Dict[str, pd.DataFrame]:
    """
    Lit un file-like uploadé et retourne dict {sheet_name: DataFrame}.
    Supporte CSV & Excel (xls, xlsx, xlsm, xlsb) et ods (si support installé).
    """
    filename = getattr(uploaded_file, 'name', 'uploaded')
    ext = _get_extension(filename)

    uploaded_file.seek(0)
    if ext == 'csv':
        df = pd.read_csv(uploaded_file)
        return {"CSV": df}

    engine = None
    if ext == 'xlsb':
        engine = 'pyxlsb'  # requires pyxlsb installed

    # pandas détecte .ods s'il y a odfpy, sinon échouera
    uploaded_file.seek(0)
    sheets = pd.read_excel(uploaded_file, sheet_name=None, engine=engine)
    # cast keys to str
    return { str(k): v for k, v in sheets.items() }

# ------------------ Transformation principale ------------------

def transform_to_nwp(df: pd.DataFrame) -> pd.DataFrame:
    """
    Transforme le DataFrame source en DataFrame respectant TARGET_COLUMNS et règles NWP.
    Lève ValueError si des colonnes requises manquent (insensible à la casse).
    """
    if not isinstance(df, pd.DataFrame):
        raise ValueError("transform_to_nwp attend un pandas.DataFrame")

    # build normalized column map and check required columns
    cmap = _build_col_map(df)
    missing = []
    for req in REQUIRED_ORIGINAL_COLUMNS:
        if _norm_col_name(req) not in cmap:
            missing.append(req)
    if missing:
        raise ValueError(f"Colonnes requises manquantes: {missing}")

    # create a working copy
    d = df.copy()

    # Access original columns robustly
    col_Address = _get_col(d, "Address")
    col_Suite = _get_col(d, "Suite")
    col_City = _get_col(d, "City")
    col_Postal = _get_col(d, "Postal_code")
    col_Province = _get_col(d, "Province")
    col_Name = _get_col(d, "Name")
    col_Telephone = _get_col(d, "Telephone")
    col_Language = _get_col(d, "Language")
    col_Lat = _get_col(d, "Latitude")
    col_Lon = _get_col(d, "Longitude")
    col_Notes = _get_col(d, "Notes")

    # Parse address -> Number & Street
    addr_series = d[col_Address].astype(str).where(d[col_Address].notna(), None)
    parsed = addr_series.apply(_parse_address_field)
    # parsed is Series of tuples -> unzip safely
    numbers = parsed.apply(lambda t: t[0] if isinstance(t, tuple) else None)
    streets = parsed.apply(lambda t: t[1] if isinstance(t, tuple) else None)
    d['_NWP_Number'] = numbers
    d['_NWP_Street'] = streets

    # Phone format
    d['_NWP_Phone'] = d[col_Telephone].apply(_format_phone_to_north_american) if col_Telephone else None

    # Build output DataFrame with exact TARGET_COLUMNS order
    out = pd.DataFrame(index=d.index)

    out['TerritoryID'] = None
    out['TerritoryNumber'] = None
    out['CategoryCode'] = None
    out['Category'] = None
    out['TerritoryAddressID'] = None
    out['ApartmentNumber'] = d[col_Suite] if col_Suite else None
    out['Number'] = d['_NWP_Number']
    out['Street'] = d['_NWP_Street']
    out['Suburb'] = d[col_City] if col_City else None
    out['PostalCode'] = d[col_Postal] if col_Postal else None
    out['State'] = d[col_Province] if col_Province else None
    out['Name'] = d[col_Name] if col_Name else None
    out['Phone'] = d['_NWP_Phone'] if col_Telephone else None
    out['Type'] = None
    out['Status'] = None
    out['NotHomeAttempt'] = None
    out['Date1'] = None
    out['Date2'] = None
    out['Date3'] = None
    out['Date4'] = None
    out['Date5'] = None
    out['Language'] = d[col_Language] if col_Language else None
    out['Latitude'] = d[col_Lat] if col_Lat else None
    out['Longitude'] = d[col_Lon] if col_Lon else None
    out['Notes'] = d[col_Notes] if col_Notes else None
    out['NotesFromPublisher'] = None

    # Title-case (première lettre de chaque mot en majuscule) pour colonnes textuelles demandées
    title_cols = [
        'Street', 'Suburb', 'State', 'Name', 'ApartmentNumber', 'Category', 'CategoryCode', 'Type', 'Status'
    ]
    for c in title_cols:
        if c in out.columns:
            out[c] = out[c].apply(_title_case_safe)

    # Colonnes primaires pour le tri et la déduplication
    primary_cols = ["Street", "Number", "ApartmentNumber", "Suburb", "State"]

    # Construire un DataFrame auxiliaire pour trier
    sort_df = pd.DataFrame(index=out.index)
    for col in primary_cols:
        if col in out.columns:
            sort_df[col] = out[col].astype(str).where(out[col].notna(), '')

    if not sort_df.empty:
        # Trier selon l’ordre donné
        sorted_idx = sort_df.sort_values(
            by=primary_cols,
            na_position="last"
        ).index
        out = out.reindex(sorted_idx)

    # Supprimer doublons basés uniquement sur les colonnes primaires
    out = out.drop_duplicates(subset=primary_cols, keep="first")

    # Ensure final column order exactly matches TARGET_COLUMNS (and add missing None)
    for col in TARGET_COLUMNS:
        if col not in out.columns:
            out[col] = None
    out = out[TARGET_COLUMNS]

    # Réinitialiser l’index en commençant à 1
    out = out.reset_index(drop=True)
    out.index = out.index + 1

    return out  

# ------------------ Sérialisation CSV / ZIP ------------------

def df_to_csv_bytes(df: pd.DataFrame, sep: str = ',', encoding: str = 'utf-8-sig') -> bytes:
    """Convertit DataFrame en bytes CSV (BOM utf-8-sig pour compatibilité Windows/Excel)."""
    buf = io.StringIO()
    df.to_csv(buf, index=False, sep=sep)
    return buf.getvalue().encode(encoding)


def sheets_to_zip_bytes(sheets: Dict[str, pd.DataFrame], sep: str = ',', encoding: str = 'utf-8-sig') -> bytes:
    """Crée un zip en mémoire contenant chaque sheet.csv."""
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
        for name, df in sheets.items():
            safe_name = re.sub(r'[^0-9A-Za-z_\-\.]+', '_', str(name))
            csv_bytes = df_to_csv_bytes(df, sep=sep, encoding=encoding)
            zf.writestr(f"{safe_name}.csv", csv_bytes)
    mem.seek(0)
    return mem.getvalue()


# ------------------ Flux principal pour l'app ------------------
def process_upload(uploaded_file, sheet_name: Optional[str] = None, sep: str = ',') -> Dict[str, Any]:
    """
    Lit le fichier uploadé, applique la transformation NWP (transform_to_nwp) sur la
    feuille demandée (ou toutes les feuilles si None), et renvoie :
      {'sheets': {name: transformed_df}, 'warnings': [...], 'output': bytes, 'output_name': str}
    """
    # read workbook
    sheets = read_workbook_from_filelike(uploaded_file)

    # select sheet if requested
    if sheet_name:
        if sheet_name not in sheets:
            raise ValueError(f"Feuille '{sheet_name}' non trouvée dans le fichier.")
        sheets = {sheet_name: sheets[sheet_name]}

    transformed = {}
    warnings_list: List[str] = []

    for name, df in sheets.items():
        try:
            out_df = transform_to_nwp(df)
        except ValueError as e:
            # missing columns -> bubble up as warning and stop transformation for this sheet
            raise
        transformed[name] = out_df

    # single sheet -> CSV bytes, multiple -> zip
    if len(transformed) == 1:
        key = next(iter(transformed.keys()))
        out_bytes = df_to_csv_bytes(transformed[key], sep=sep)
        out_name = f"{Path(getattr(uploaded_file, 'name', 'output')).stem}.csv"
    else:
        out_bytes = sheets_to_zip_bytes(transformed, sep=sep)
        out_name = f"{Path(getattr(uploaded_file, 'name', 'output')).stem}_sheets.zip"

    return {
        "sheets": transformed,
        "warnings": warnings_list,
        "output": out_bytes,
        "output_name": out_name,
    }