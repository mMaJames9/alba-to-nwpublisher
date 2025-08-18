# ------------------ Helpers pour colonnes / lookup ------------------

from pathlib import Path
import re

import pandas as pd
from typing import Any, Dict, Tuple, Optional


def _norm_col_name(s: Optional[str]) -> str:
    """Normalise un nom de colonne pour comparaison (insensible à la casse/espaces/_)."""
    return re.sub(r'[\s_]+', '', str(s).strip()).lower()


def _build_col_map(df: pd.DataFrame) -> Dict[str, str]:
    """
    Renvoie un dict qui mappe nom normalisé -> nom original présent dans df.
    Ex: {'telephone': 'Telephone', 'postalcode': 'Postal_code', ...}
    """
    return { _norm_col_name(c): c for c in df.columns }


def _get_col(df: pd.DataFrame, name: str) -> Optional[str]:
    """
    Retourne le nom exact de la colonne dans df correspondant à 'name' (case-insensitive / underscores ignored).
    Retourne None si pas trouvé.
    """
    cmap = _build_col_map(df)
    return cmap.get(_norm_col_name(name))


# ------------------ Lecture workbook ------------------

def _get_extension(filename: str) -> str:
    return Path(filename).suffix.lower().lstrip('.')

# ------------------ Parsing / nettoyage spécifiques ------------------

def _parse_address_field(addr: Any) -> Tuple[Optional[str], Optional[str]]:
    """
    Sépare numéro et nom de la rue.
    Accepts: '142 Marrissa Ave.', '12A-14 Main St', 'No number', etc.
    Returns: (number_or_None, street_or_None)
    """
    if pd.isna(addr):
        return None, None
    s = str(addr).strip()
    # try leading number (may include letters or dash/slash)
    m = re.match(r'^\s*([0-9]+[A-Za-z0-9\-\/]*)\s+(.*)$', s)
    if m:
        number = m.group(1).strip()
        street = m.group(2).strip().rstrip('.')
        return number, street
    # no leading number: return None as number and full string as street
    return None, s.rstrip('.')


def _format_phone_to_north_american(phone_text: Any) -> Optional[str]:
    """
    Formate un numéro en +1 (AAA) PPP-SSSS uniquement si c'est clairement un NANP.
    - 10 chiffres -> formaté
    - 11 chiffres débutant par '1' -> on retire le 1, formaté
    - commence par '+' -> renvoyé tel quel (international)
    - sinon -> renvoyé inchangé (conformité stricte)
    """
    if phone_text is None or (isinstance(phone_text, float) and pd.isna(phone_text)) or pd.isna(phone_text):
        return None

    # Corrige le cas des floats venant d'Excel/CSV
    if isinstance(phone_text, float) and phone_text.is_integer():
        s = str(int(phone_text))
    else:
        s = str(phone_text).strip()

    # Numéro international explicite
    if s.startswith('+'):
        return s

    digits = re.sub(r"\D+", "", s)

    if len(digits) == 10:
        return f"+1 ({digits[0:3]}) {digits[3:6]}-{digits[6:10]}"

    if len(digits) == 11 and digits[0] == '1':
        d = digits[1:]
        return f"+1 ({d[0:3]}) {d[3:6]}-{d[6:10]}"

    return s if s else None


def _title_case_safe(x: Any) -> Any:
    """Met la première lettre de chaque mot en majuscule pour une chaîne."""
    if pd.isna(x):
        return x
    if isinstance(x, str):
        # .title() is simple and matches the user's request
        return x.title()
    return x