import pytest

from bill_analysis import extraction


def test_extracts_legacy_fpl_labels_into_normalized_record(monkeypatch):
    text = """
    Account number: 12345-67890
    Rate: GSD-1 GENERAL
    Service to Mar 15, 2024
    Service days 30
    kWh Used 1,000
    Demand KW 10 1.0 25
    Non-fuel energy charge: $0.100000
    Fuel charge: $0.020000
    Demand charge: $10.00
    Base charge: $15.00
    Utility tax $5.00
    Gross receipts tax $2.00
    Late payment charge $3.00
    FPL SolarTogether credit −4.50
    """
    monkeypatch.setattr(extraction, "_read_pdf_text", lambda _: text)

    record = extraction.extract_bill_data(b"placeholder")

    assert record["Service month"] == 3
    assert record["Rate"] == "GSD-1 GENERAL"
    assert record["Total consumption kWh"] == 1000
    assert record["Demand kW"] == 10
    assert record["Usage"] == 25
    assert record["FPL SolarTogether credit"] == -4.5
    assert record["Total charge"] > 0


def test_missing_service_date_is_a_clear_extraction_error(monkeypatch):
    monkeypatch.setattr(extraction, "_read_pdf_text", lambda _: "Rate: GSD-1 GENERAL")

    with pytest.raises(extraction.MissingServiceDateError, match="Service to"):
        extraction.extract_bill_data(b"placeholder", filename="bad_bill.pdf")
