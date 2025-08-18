import io
import pandas as pd
import pytest
from src.alba2nwpublisher.convert import (
    transform_to_nwp,
    df_to_csv_bytes,
    sheets_to_zip_bytes,
    process_upload,
    read_workbook_from_filelike,
    TARGET_COLUMNS,
    REQUIRED_ORIGINAL_COLUMNS,
)

def make_valid_df():
    # Minimal valid DataFrame with all required columns
    data = {col: [f"val_{i}_{col}" for i in range(2)] for col in REQUIRED_ORIGINAL_COLUMNS}
    # Add some realistic address/phone values
    data["Address"] = ["123 Main St", "456 Elm Rd"]
    data["Suite"] = ["A", "B"]
    data["City"] = ["Montreal", "Toronto"]
    data["Province"] = ["QC", "ON"]
    data["Postal_code"] = ["H1A1A1", "M2B2B2"]
    data["Telephone"] = ["514-555-1234", "416-555-5678"]
    data["Latitude"] = [str(45.5), str(43.7)]
    data["Longitude"] = [str(-73.6), str(-79.4)]
    data["Name"] = ["Smith", "Jones"]
    data["Language"] = ["fr", "en"]
    data["Notes"] = ["note1", "note2"]
    return pd.DataFrame(data)

def test_transform_to_nwp_valid():
    df = make_valid_df()
    out = transform_to_nwp(df)
    assert isinstance(out, pd.DataFrame)
    assert list(out.columns) == TARGET_COLUMNS
    assert len(out) == 2
    # Find the row with Number == "123"
    row = out[out["Number"] == "123"].iloc[0]
    assert str(row["Street"]).startswith("Main")
    assert "514" in str(row["Phone"])

def test_transform_to_nwp_missing_column():
    df = make_valid_df().drop(columns=["Address"])
    with pytest.raises(ValueError) as e:
        transform_to_nwp(df)
    assert "Colonnes requises manquantes" in str(e.value)

def test_df_to_csv_bytes_content():
    df = make_valid_df()
    out = transform_to_nwp(df)
    csv_bytes = df_to_csv_bytes(out)
    assert isinstance(csv_bytes, bytes)
    # Should start with BOM
    assert csv_bytes.startswith(b'\xef\xbb\xbf')
    # Should contain header
    assert b"TerritoryID" in csv_bytes

def test_sheets_to_zip_bytes_multiple():
    df1 = transform_to_nwp(make_valid_df())
    df2 = transform_to_nwp(make_valid_df())
    sheets = {"Sheet1": df1, "Sheet2": df2}
    zip_bytes = sheets_to_zip_bytes(sheets)
    assert isinstance(zip_bytes, bytes)
    # Check ZIP file signature
    assert zip_bytes[:2] == b'PK'

def test_read_workbook_from_filelike_csv():
    content = "Address_ID,Territory_ID,Language,Status,Name,Suite,Address,City,Province,Postal_code,Country,Latitude,Longitude,Telephone,Owner,Notes,Notes_private,Account,Created,Modified,Contacted,Geocoded,Territory_number,Territory_description\n1,2,en,active,Smith,A,123 Main St,Montreal,QC,H1A1A1,Canada,45.5,-73.6,514-555-1234,owner1,note1,priv1,acc1,2020,2021,yes,yes,001,desc\n"
    file = io.StringIO(content)
    file.name = "test.csv"
    sheets = read_workbook_from_filelike(file)
    assert "CSV" in sheets
    assert isinstance(sheets["CSV"], pd.DataFrame)
    assert sheets["CSV"].iloc[0]["Name"] == "Smith"

def test_process_upload_single_sheet():
    df = make_valid_df()
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    buf.seek(0)
    buf.name = "input.csv"
    result = process_upload(buf)
    assert "sheets" in result
    assert "output" in result
    assert result["output_name"].endswith(".csv")
    assert isinstance(result["output"], bytes)

def test_process_upload_sheet_not_found():
    df = make_valid_df()
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    buf.seek(0)
    buf.name = "input.csv"
    with pytest.raises(ValueError):
        process_upload(buf, sheet_name="NotFound")