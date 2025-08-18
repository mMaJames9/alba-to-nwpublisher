import pandas as pd
import pytest
from src.alba2nwpublisher.utils import (
    _norm_col_name,
    _build_col_map,
    _get_col,
    _get_extension,
    _parse_address_field,
    _format_phone_to_north_american,
    _title_case_safe,
)

def test_norm_col_name():
    assert _norm_col_name("Postal_code") == "postalcode"
    assert _norm_col_name("  Telephone ") == "telephone"
    assert _norm_col_name("Suite Number") == "suitenumber"
    assert _norm_col_name("Suite_Number") == "suitenumber"
    assert _norm_col_name("SUITE_NUMBER") == "suite_number".replace("_", "").lower()

def test_build_col_map():
    df = pd.DataFrame(columns=["Postal_code", "Telephone", "Suite Number"])
    cmap = _build_col_map(df)
    assert cmap["postalcode"] == "Postal_code"
    assert cmap["telephone"] == "Telephone"
    assert cmap["suitenumber"] == "Suite Number"

def test_get_col():
    df = pd.DataFrame(columns=["Postal_code", "Telephone", "Suite Number"])
    assert _get_col(df, "postal_code") == "Postal_code"
    assert _get_col(df, "TELEPHONE") == "Telephone"
    assert _get_col(df, "suite number") == "Suite Number"
    assert _get_col(df, "notfound") is None

def test_get_extension():
    assert _get_extension("file.csv") == "csv"
    assert _get_extension("file.XLSX") == "xlsx"
    assert _get_extension("file.with.dots.xlsm") == "xlsm"
    assert _get_extension("file") == ""
    assert _get_extension("file.") == ""
    assert _get_extension(".hiddenfile") == ""

@pytest.mark.parametrize("addr,expected", [
    ("142 Marrissa Ave.", ("142", "Marrissa Ave")),
    ("12A-14 Main St", ("12A-14", "Main St")),
    ("No number", (None, "No number")),
    ("", (None, "")),
    (None, (None, None)),
    (float("nan"), (None, None)),
])
def test_parse_address_field(addr, expected):
    assert _parse_address_field(addr) == expected

@pytest.mark.parametrize("phone,expected", [
    ("5145551234", "+1 (514) 555-1234"),
    ("1-514-555-1234", "+1 (514) 555-1234"),
    ("+33 1 23 45 67 89", "+33 1 23 45 67 89"),
    ("555-1234", "555-1234"),  # <-- correction ici
    ("", None),
    (None, None),
    (float("nan"), None),
    (5145551234, "+1 (514) 555-1234"),
])
def test_format_phone_to_north_american(phone, expected):
    assert _format_phone_to_north_american(phone) == expected

@pytest.mark.parametrize("val,expected", [
    ("main street", "Main Street"),
    ("", ""),
    (None, None),
    (float("nan"), float("nan")),
])
def test_title_case_safe(val, expected):
    result = _title_case_safe(val)
    if isinstance(expected, float) and pd.isna(expected):
        assert pd.isna(result)
    else:
        assert result == expected

def test_norm_col_name_edge_cases():
    assert _norm_col_name("") == ""
    assert _norm_col_name(None) == "none"

def test_build_col_map_empty_df():
    df = pd.DataFrame()
    cmap = _build_col_map(df)
    assert cmap == {}

def test_get_col_not_found():
    df = pd.DataFrame(columns=["A", "B"])
    assert _get_col(df, "C") is None

def test_get_extension_edge_cases():
    assert _get_extension("") == ""
    assert _get_extension(str(123)) == ""
    assert _get_extension(".") == ""

@pytest.mark.parametrize("addr", [None, "", float("nan")])
def test_parse_address_field_invalid(addr):
    number, street = _parse_address_field(addr)
    assert number is None

@pytest.mark.parametrize("phone", ["", None, float("nan"), "abc", "123"])
def test_format_phone_to_north_american_invalid(phone):
    assert _format_phone_to_north_american(phone) is None or isinstance(_format_phone_to_north_american(phone), str)

@pytest.mark.parametrize("val", [None, "", float("nan"), 123])
def test_title_case_safe_invalid(val):
    result = _title_case_safe(val)
    if isinstance(val, float) and pd.isna(val):
        assert pd.isna(result)
    elif val is None or val == "":
        assert result == val
    elif isinstance(val, int):
        assert result == val