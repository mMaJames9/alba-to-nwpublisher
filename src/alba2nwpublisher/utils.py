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
    - Accepte: '142 Marrissa Ave.', '12A-14 Main St', 'No number', etc.
    - Retourne: (number_or_None, street_or_None)
    - Le numéro de rue est toujours normalisé sans décimales.
    """
    if pd.isna(addr):
        return None, None

    s = str(addr).strip()

    # Corrige le cas Excel/CSV: 142.0 → 142
    if re.match(r'^\d+\.0$', s):
        s = s[:-2]

    # Détection du numéro en tête
    m = re.match(r'^\s*([0-9]+[A-Za-z0-9\-\/]*)\s+(.*)$', s)
    if m:
        number = m.group(1).strip()
        street = m.group(2).strip().rstrip('.')

        # Si le numéro est un float déguisé → on force en entier
        if re.match(r'^\d+(\.0+)?$', number):
            number = str(int(float(number)))

        return number, street

    # Pas de numéro en tête → retourne tout comme rue
    return None, s.rstrip('.')


def _format_phone_to_north_american(phone_text: Any) -> Optional[str]:
    """
    Formate un numéro nord-américain (NANP) en AAA-PPP-SSSS.
    - 10 chiffres -> formaté
    - 11 chiffres débutant par '1' -> on retire le 1, formaté
    - commence par '+' -> renvoyé tel quel (international)
    - sinon -> renvoyé inchangé
    """
    if phone_text is None or (isinstance(phone_text, float) and pd.isna(phone_text)) or pd.isna(phone_text):
        return None

    # Corrige le cas des floats venant d'Excel/CSV
    if isinstance(phone_text, float) and phone_text.is_integer():
        s = str(int(phone_text))
    else:
        s = str(phone_text).strip()

    # Numéro international explicite → ne pas reformater
    if s.startswith('+'):
        return s

    digits = re.sub(r"\D+", "", s)

    if len(digits) == 10:
        return f"{digits[0:3]}-{digits[3:6]}-{digits[6:10]}"

    if len(digits) == 11 and digits[0] == '1':
        d = digits[1:]
        return f"{d[0:3]}-{d[3:6]}-{d[6:10]}"

    return s if s else None


def _title_case_safe(x: Any) -> Any:
    """
    Capitalisation "smart" :
      - conserve les acronymes (tout en majuscule, longueur <= 3 ou origine en MAJUSCULES)
      - gère les apostrophes (Hunter's -> Hunter's ; O'NEIL -> O'Neil)
      - gère les séparateurs '-', '/' (North-West -> North-West)
      - laisse les valeurs non-string inchangées (sauf NaN)
    """

    if x is None or (isinstance(x, float) and pd.isna(x)) or pd.isna(x):
        return x

    if not isinstance(x, str):
        # si c'est un nombre (ex: 123) on le retourne tel quel
        return x

    s = x.strip()
    if s == "":
        return s

    # helper : capitaliser un segment (hors séparateurs)
    def cap_segment(seg: str) -> str:
        # si c'était un acronyme en entrée (tout en majuscules et 2-3 lettres), on garde
        if seg.isupper() and len(seg) <= 3:
            return seg
        # si le segment est tout en majuscules et >3 lettres (ex: NASA) on conserve aussi
        if seg.isupper():
            return seg
        # sinon on capitalise première lettre et on met le reste en minuscules
        # (gère "mcDonald" mal — mais couvre la plupart des cas)
        return seg[:1].upper() + seg[1:].lower() if seg else seg

    # traiter token par token (garder les espaces tels quels)
    tokens = re.split(r'(\s+)', s)

    out_tokens = []
    for token in tokens:
        if token.isspace():
            out_tokens.append(token)
            continue

        # gérer les apostrophes, en conservant le séparateur et traitant la partie après correctement
        parts = re.split(r"(')", token)  # conserve l'apostrophe comme élément
        rebuilt_parts = []
        i = 0
        while i < len(parts):
            part = parts[i]
            if part == "'":
                # si apostrophe, regarder la partie suivante si elle existe
                next_part = parts[i+1] if i+1 < len(parts) else ""
                if next_part.lower() == 's':
                    # possessif : 's (toujours minuscule s)
                    rebuilt_parts.append("'s")
                    i += 2
                    continue
                else:
                    # autre apostrophe (p.ex. O'neill) : on capitalise la partie suivante
                    if next_part:
                        # traiter next_part avec séparateurs -/
                        subparts = re.split(r'([-/])', next_part)
                        cap_sub = ''.join(cap_segment(sp) if sp not in "-/" else sp for sp in subparts)
                        rebuilt_parts.append("'" + cap_sub)
                        i += 2
                        continue
                    else:
                        rebuilt_parts.append("'")
                        i += 1
                        continue
            else:
                # pas une apostrophe ; traiter les sous-séparateurs '-' et '/'
                subparts = re.split(r'([-/])', part)
                cap_subs = []
                for sp in subparts:
                    if sp in "-/":
                        cap_subs.append(sp)
                    else:
                        # si segment initial était tout en majuscule et court => acronyme
                        if sp.isupper() and len(sp) <= 3:
                            cap_subs.append(sp)
                        else:
                            cap_subs.append(cap_segment(sp))
                rebuilt_parts.append(''.join(cap_subs))
                i += 1

        out_tokens.append(''.join(rebuilt_parts))

    return ''.join(out_tokens)