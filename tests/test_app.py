import io
import pandas as pd
import pytest
from unittest.mock import patch, MagicMock
from app import get_extension, df_to_bytes_csv, read_workbook

# Python


def test_get_extension_basic():
    assert get_extension("file.csv") == "csv"
    assert get_extension("file.xlsx") == "xlsx"
    assert get_extension("file.XLSB") == "xlsb"
    assert get_extension("file.with.dots.xlsm") == "xlsm"
    assert get_extension("file") == ""
    assert get_extension("file.") == ""
    assert get_extension(".hiddenfile") == ""

def test_df_to_bytes_csv_default():
    df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
    result = df_to_bytes_csv(df)
    assert isinstance(result, bytes)
    # Should start with BOM for utf-8-sig
    assert result.startswith(b'\xef\xbb\xbf')
    # Should contain header and data
    assert b"A,B" in result
    assert b"1,x" in result

def test_df_to_bytes_csv_custom_sep():
    df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
    result = df_to_bytes_csv(df, sep=";")
    assert b"A;B" in result
    assert b"1;x" in result

def make_csv_file():
    # Simulate an uploaded CSV file-like object
    content = "col1,col2\n1,foo\n2,bar\n"
    file = io.StringIO(content)
    file.name = "test.csv"
    return file

def test_read_workbook_csv(monkeypatch):
    uploaded_file = make_csv_file()
    result = read_workbook(uploaded_file)
    assert isinstance(result, dict)
    assert "CSV" in result
    df = result["CSV"]
    assert list(df.columns) == ["col1", "col2"]
    assert df.iloc[0]["col2"] == "foo"

def test_read_workbook_xlsb(monkeypatch):
    # Simulate an uploaded xlsb file-like object
    fake_file = MagicMock()
    fake_file.name = "test.xlsb"
    # Patch pandas.read_excel to return a dict of DataFrames
    with patch("pandas.read_excel") as mock_read_excel:
        mock_read_excel.return_value = {"Sheet1": pd.DataFrame({"A": [1]})}
        result = read_workbook(fake_file)
        assert result is not None
        assert "Sheet1" in result
        assert isinstance(result["Sheet1"], pd.DataFrame)
        mock_read_excel.assert_called_with(fake_file, sheet_name=None, engine="pyxlsb")

def test_read_workbook_excel(monkeypatch):
    fake_file = MagicMock()
    fake_file.name = "test.xlsx"
    with patch("pandas.read_excel") as mock_read_excel:
        mock_read_excel.return_value = {"SheetA": pd.DataFrame({"X": [42]})}
        result = read_workbook(fake_file)
        assert result is not None
        assert "SheetA" in result
        mock_read_excel.assert_called_with(fake_file, sheet_name=None, engine=None)

def test_read_workbook_error(monkeypatch):
    fake_file = MagicMock()
    fake_file.name = "bad.csv"
    # Patch pandas.read_csv to raise an error
    with patch("pandas.read_csv", side_effect=Exception("fail")):
        # Patch st.error to avoid Streamlit side effects
        with patch("streamlit.error"):
            result = read_workbook(fake_file)
            assert result is None