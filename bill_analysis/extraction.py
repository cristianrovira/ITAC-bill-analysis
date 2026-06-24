"""Readable, in-memory extraction of normalized values from FPL bill PDFs."""

from __future__ import annotations

import io
import re
from datetime import datetime
from pathlib import Path
from typing import Any, BinaryIO, Iterable

import pandas as pd
import pdfplumber

from .calculations import calculate_bill_fields
from .estimation import validate_unique_months
from .utils import DETAIL_FIELDS, round_monthly_data


class BillExtractionError(ValueError):
    """Base class for user-facing extraction failures."""


class UnreadablePDFError(BillExtractionError):
    pass


class MissingServiceDateError(BillExtractionError):
    pass


NUMBER = r"\(?[−+\-]?[\d,]+(?:\.\d+)?\)?"


FIELD_LABELS: dict[str, tuple[str, ...]] = {
    "Service days": ("Service days",),
    "kWh used": ("kWh Used", "kWh used"),
    "Total consumption kWh": ("Total Consumption KWH", "Total Comsuption KWH"),
    "Contract demand": ("Contract demand",),
    "On-peak kWh used": ("On-Peak kWh used", "On-peak kWh used"),
    "Off-peak kWh used": ("Off-peak kWh used", "Off-Peak kWh used"),
    "On-peak demand": ("On-peak demand", "On-Peak demand"),
    "Maximum demand": ("Maximum demand",),
    "Base charge": ("Base charge",),
    "Customer charge": ("Customer charge",),
    "Service charge": ("Service Charge",),
    "Gross receipts tax / Regulatory fee": (
        "Gross rec. tax/Regulatory fee",
        "Gross receipts tax/Regulatory fee",
    ),
    "Franchise charge": ("Franchise charge",),
    "Franchise fee": ("Franchise fee",),
    "Utility tax": ("Utility tax",),
    "Florida sales tax": ("Florida sales tax",),
    "County sales tax": ("County sales tax",),
    "Discretionary sales surtax": ("Discretionary sales surtax",),
    "Gross receipts tax": ("Gross receipts tax",),
    "Regulatory fee": (
        "Regulatory fee (State fee)",
        "Regulatoiy fee (State fee)",
        "Regulatory fee",
    ),
    "Taxes and charges": ("Taxes and charges",),
    "FPL SolarTogether charge": ("FPL SolarTogether charge",),
    "FPL SolarTogether credit": ("FPL SolarTogether credit",),
    "Power monitoring premium plus": (
        "Power monitoring-premium plus",
        "Power monitoring premium plus",
    ),
    "Late payment charge": ("Late payment charge",),
    "Total services and tax": ("Total services and tax", "Total Services and Tax"),
    "Total electric cost": ("Total electric cost",),
    "Total energy charge": ("Total energy charge",),
    "Total demand charge": (
        "Total Demand Charge",
        "Total demand charge",
    ),
    "Total charge": ("Total charge", "Total charges", "Amount due"),
}


def _filename(pdf_file: Any, supplied: str | None = None) -> str:
    return supplied or getattr(pdf_file, "name", None) or "uploaded PDF"


def _normalise_text(text: str) -> str:
    replacements = {
        "Euel": "Fuel",
        "Eeb": "Feb",
        "Regulatoiy": "Regulatory",
        "\u00a0": " ",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return "\n".join(re.sub(r"[ \t]+", " ", line).strip() for line in text.splitlines())


def _read_pdf_text(pdf_file: BinaryIO | bytes) -> str:
    source: BinaryIO | io.BytesIO
    if isinstance(pdf_file, bytes):
        source = io.BytesIO(pdf_file)
    else:
        source = pdf_file
        if hasattr(source, "seek"):
            source.seek(0)
    try:
        with pdfplumber.open(source) as pdf:
            pages = [(page.extract_text() or "") for page in pdf.pages]
    except Exception as exc:  # pdfplumber emits several parser-specific exception types
        raise UnreadablePDFError(f"The PDF could not be opened: {exc}") from exc
    finally:
        if hasattr(source, "seek"):
            source.seek(0)

    text = _normalise_text("\n".join(pages))
    if not text.strip():
        raise UnreadablePDFError("The PDF contains no extractable text; it may be scanned or damaged.")
    return text


def _parse_number(raw: str) -> float:
    text = raw.strip().replace("−", "-").replace(",", "").replace("$", "")
    negative = text.startswith("(") and text.endswith(")")
    value = float(text.strip("()"))
    return -abs(value) if negative else value


def _number_after_label(text: str, labels: Iterable[str]) -> float | None:
    for label in labels:
        pattern = rf"(?im)^.*?\b{re.escape(label)}\b\s*:?[ \t]*(?:\$[ \t]*)?({NUMBER})"
        match = re.search(pattern, text)
        if match:
            try:
                return _parse_number(match.group(1))
            except ValueError:
                continue
    return None


def _line_numbers(text: str, label: str) -> list[float]:
    match = re.search(rf"(?im)^.*?{re.escape(label)}[^\n]*$", text)
    if not match:
        return []
    suffix = match.group().split(label, 1)[-1]
    values: list[float] = []
    for raw in re.findall(NUMBER, suffix):
        try:
            values.append(_parse_number(raw))
        except ValueError:
            pass
    return values


def _service_date(text: str) -> datetime | None:
    month_pattern = r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
    patterns = (
        rf"Service\s+to\s+({month_pattern}\s+\d{{1,2}},\s+\d{{4}})",
        rf"Service\s+period[^\n]*?to\s+({month_pattern}\s+\d{{1,2}},\s+\d{{4}})",
    )
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            cleaned = re.sub(r"\s+", " ", match.group(1)).strip()
            for date_format in ("%b %d, %Y", "%B %d, %Y"):
                try:
                    return datetime.strptime(cleaned, date_format)
                except ValueError:
                    continue
    return None


def _rate(text: str) -> str:
    match = re.search(r"(?im)^.*?\bRate:\s*([^\n]+)$", text)
    if not match:
        return ""
    value = match.group(1).strip()
    for stopper in ("Service", "Account", "Meter"):
        value = value.split(stopper, 1)[0].strip()
    return value[:120]


def _account_number(text: str) -> str:
    match = re.search(r"(?im)\bAccount(?:\s+number|\s+no\.?|\s*#)?\s*:?\s*([\d-]{5,})", text)
    return match.group(1) if match else ""


def _section_rates(text: str, heading: str) -> tuple[float | None, float | None]:
    match = re.search(rf"(?is){re.escape(heading)}\s*:?(.*?)(?:\n\s*\n|Demand charge:|Base charge:|$)", text)
    section = match.group(1)[:600] if match else ""
    on_peak = _number_after_label(section, ("On-peak", "On-Peak"))
    off_peak = _number_after_label(section, ("Off-peak", "Off-Peak"))
    return on_peak, off_peak


def extract_bill_data(
    pdf_file: BinaryIO | bytes,
    account_number: str | None = None,
    filename: str | None = None,
) -> dict[str, Any]:
    """Extract one FPL bill into the normalized monthly schema."""
    name = _filename(pdf_file, filename)
    try:
        text = _read_pdf_text(pdf_file)
        service_date = _service_date(text)
        if service_date is None:
            raise MissingServiceDateError(f"{name}: no recognizable 'Service to' date was found.")

        record: dict[str, Any] = {
            "Account number": account_number or _account_number(text) or "Unspecified",
            "Rate": _rate(text),
            "Service to date": service_date,
            "Service month": service_date.month,
            "Service year": service_date.year,
            "Source file": Path(name).name,
            "Data source": "Actual",
            "Estimate method": "Actual uploaded bill",
            "Confidence": "High",
        }
        for field, labels in FIELD_LABELS.items():
            value = _number_after_label(text, labels)
            if value is not None:
                record[field] = value

        demand_values = _line_numbers(text, "Demand KW")
        if demand_values:
            record["Demand kW"] = demand_values[0]
            record["Usage"] = demand_values[-1]

        on_demand_values = _line_numbers(text, "On-peak demand")
        if on_demand_values:
            record["On-peak demand"] = on_demand_values[-1]

        record["Non-fuel energy rate"] = _number_after_label(text, ("Non-fuel energy charge", "Non-fuel"))
        record["Fuel rate"] = _number_after_label(text, ("Fuel charge", "Fuel"))
        record["Demand charge rate"] = _number_after_label(text, ("Demand charge", "Demand"))
        record["Maximum demand rate"] = _number_after_label(text, ("Maximum",))

        on_energy, off_energy = _section_rates(text, "Non-fuel energy charge")
        on_fuel, off_fuel = _section_rates(text, "Fuel charge")
        record["On-peak non-fuel energy rate"] = on_energy
        record["Off-peak non-fuel energy rate"] = off_energy
        record["On-peak fuel rate"] = on_fuel
        record["Off-peak fuel rate"] = off_fuel

        for field in DETAIL_FIELDS:
            record.setdefault(field, 0.0)
        return calculate_bill_fields(record)
    except BillExtractionError:
        raise
    except Exception as exc:
        raise BillExtractionError(f"{name}: extraction failed: {exc}") from exc


def extract_uploaded_bills(uploaded_files: Iterable[BinaryIO], account_number: str) -> pd.DataFrame:
    """Extract and validate all uploaded bills for one account."""
    records = [
        extract_bill_data(uploaded_file, account_number=account_number)
        for uploaded_file in uploaded_files
    ]
    if not records:
        raise BillExtractionError(f"{account_number}: no PDF files were uploaded.")
    frame = pd.DataFrame(records)
    validate_unique_months(frame)
    return round_monthly_data(frame.sort_values("Service month").reset_index(drop=True))
