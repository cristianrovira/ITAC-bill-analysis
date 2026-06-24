"""Shared constants and small, side-effect-free helpers."""

from __future__ import annotations

import calendar
import math
import re
from collections import Counter
from datetime import date, datetime
from typing import Any, Iterable

import pandas as pd

MONTH_NAMES = {month: calendar.month_name[month] for month in range(1, 13)}
MONTH_ABBREVIATIONS = {month: calendar.month_abbr[month] for month in range(1, 13)}

METADATA_FIELDS = [
    "Account number",
    "Rate",
    "Service to date",
    "Service month",
    "Service year",
    "Data source",
    "Estimate method",
    "Confidence",
]

KWH_FIELDS = {
    "kWh used",
    "Total consumption kWh",
    "On-peak kWh used",
    "Off-peak kWh used",
    "Non-TOU consumption kWh",
}

DEMAND_FIELDS = {
    "Demand kW",
    "Usage",
    "Contract demand",
    "On-peak demand",
    "Maximum demand",
    "Total demand",
}

CURRENCY_FIELDS = {
    "Energy charge",
    "Fuel charge",
    "Total energy charge",
    "Total demand charge",
    "Total electric cost",
    "Base charge",
    "Customer charge",
    "Service charge",
    "Franchise fee",
    "Franchise charge",
    "Utility tax",
    "Florida sales tax",
    "County sales tax",
    "Discretionary sales surtax",
    "Gross receipts tax",
    "Gross receipts tax / Regulatory fee",
    "Regulatory fee",
    "Taxes and charges",
    "FPL SolarTogether charge",
    "FPL SolarTogether credit",
    "Power monitoring premium plus",
    "Total services and tax",
    "Total charge",
    "Late payment charge",
}

RATE_FIELDS = {
    "Non-fuel energy rate",
    "Fuel rate",
    "Demand rate",
    "On-peak non-fuel energy rate",
    "Off-peak non-fuel energy rate",
    "On-peak fuel rate",
    "Off-peak fuel rate",
    "Demand charge rate",
    "Maximum demand rate",
    "Energy rate",
    "Total $/kWh cost",
}

DETAIL_FIELDS = [
    "Service days",
    "kWh used",
    "Total consumption kWh",
    "On-peak kWh used",
    "Off-peak kWh used",
    "Non-TOU consumption kWh",
    "Demand kW",
    "Usage",
    "Contract demand",
    "On-peak demand",
    "Maximum demand",
    "Total demand",
    "Non-fuel energy rate",
    "Fuel rate",
    "On-peak non-fuel energy rate",
    "Off-peak non-fuel energy rate",
    "On-peak fuel rate",
    "Off-peak fuel rate",
    "Demand charge rate",
    "Maximum demand rate",
    "Energy charge",
    "Fuel charge",
    "Total energy charge",
    "Total demand charge",
    "Total electric cost",
    "Base charge",
    "Customer charge",
    "Service charge",
    "Late payment charge",
    "Gross receipts tax / Regulatory fee",
    "Gross receipts tax",
    "Regulatory fee",
    "Franchise charge",
    "Franchise fee",
    "Utility tax",
    "Florida sales tax",
    "County sales tax",
    "Discretionary sales surtax",
    "Taxes and charges",
    "FPL SolarTogether charge",
    "FPL SolarTogether credit",
    "Power monitoring premium plus",
    "Total services and tax",
    "Total charge",
    "Energy rate",
    "Demand rate",
    "Total $/kWh cost",
]

NUMERIC_FIELDS = set(DETAIL_FIELDS) - {"Service days"}


def parse_number(value: Any, default: float | None = None) -> float | None:
    """Convert bill-like numeric text to a float, including credits and parentheses."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return default
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)

    text = str(value).strip()
    if not text:
        return default
    negative = text.startswith("(") and text.endswith(")")
    text = text.replace("−", "-").replace(",", "").replace("$", "")
    match = re.search(r"[-+]?\d*\.?\d+", text)
    if not match:
        return default
    number = float(match.group())
    return -abs(number) if negative else number


def numeric(value: Any) -> float:
    """Return a finite numeric value, treating missing data as zero."""
    parsed = parse_number(value, 0.0)
    return 0.0 if parsed is None or not math.isfinite(parsed) else parsed


def safe_divide(numerator: Any, denominator: Any, default: float = 0.0) -> float:
    denominator_value = numeric(denominator)
    if denominator_value == 0:
        return default
    return numeric(numerator) / denominator_value


def days_in_month(year: int, month: int) -> int:
    return calendar.monthrange(int(year), int(month))[1]


def infer_report_year(frame: pd.DataFrame, fallback: int | None = None) -> int:
    """Use the most common extracted service year, with a deterministic fallback."""
    fallback = fallback or datetime.now().year
    if "Service year" not in frame.columns:
        return fallback
    years = [int(value) for value in frame["Service year"].dropna() if str(value).isdigit()]
    return Counter(years).most_common(1)[0][0] if years else fallback


def month_names(months: Iterable[int], abbreviated: bool = False) -> list[str]:
    mapping = MONTH_ABBREVIATIONS if abbreviated else MONTH_NAMES
    return [mapping[int(month)] for month in sorted({int(month) for month in months})]


def month_list_text(months: Iterable[int]) -> str:
    names = month_names(months, abbreviated=True)
    return ", ".join(names) if names else "None"


def display_date(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    if isinstance(value, (date, datetime)):
        return value.strftime("%b %d, %Y")
    return str(value)


def round_monthly_data(frame: pd.DataFrame) -> pd.DataFrame:
    """Apply the report's rounding contract without mutating the input."""
    rounded = frame.copy()
    for field in KWH_FIELDS | DEMAND_FIELDS:
        if field in rounded.columns:
            rounded[field] = pd.to_numeric(rounded[field], errors="coerce").round(0)
    for field in CURRENCY_FIELDS:
        if field in rounded.columns:
            rounded[field] = pd.to_numeric(rounded[field], errors="coerce").round(2)
    for field in RATE_FIELDS:
        if field in rounded.columns:
            rounded[field] = pd.to_numeric(rounded[field], errors="coerce").round(6)
    if "Service days" in rounded.columns:
        rounded["Service days"] = pd.to_numeric(rounded["Service days"], errors="coerce").round(0)
    return rounded
