"""In-memory Excel report generation for actual and estimated bill data."""

from __future__ import annotations

import re
from io import BytesIO
from typing import Any

import pandas as pd

from .calculations import (
    build_consolidated_frame,
    build_recommendations,
    calculate_annual_operation_hours,
)
from .utils import (
    CURRENCY_FIELDS,
    DEMAND_FIELDS,
    DETAIL_FIELDS,
    KWH_FIELDS,
    MONTH_ABBREVIATIONS,
    RATE_FIELDS,
    display_date,
    month_list_text,
    numeric,
    safe_divide,
)

DISPLAY_LABELS = {
    "Data source": "Data Source",
    "Estimate method": "Estimate Method",
    "Service to date": "Service to Date",
    "Service days": "Service Days",
    "Total consumption kWh": "Total Consumption kWh",
    "On-peak kWh used": "On-Peak kWh Used",
    "Off-peak kWh used": "Off-Peak kWh Used",
    "Non-TOU consumption kWh": "Non-TOU Consumption kWh",
    "Demand kW": "Demand kW",
    "On-peak demand": "On-Peak Demand kW",
    "Maximum demand": "Maximum Demand kW",
    "Total demand": "Total Demand kW",
    "Non-fuel energy rate": "Non-Fuel Energy Rate ($/kWh)",
    "Fuel rate": "Fuel Rate ($/kWh)",
    "Demand rate": "Demand Rate ($/kW)",
    "On-peak non-fuel energy rate": "On-Peak Energy Rate ($/kWh)",
    "Off-peak non-fuel energy rate": "Off-Peak Energy Rate ($/kWh)",
    "On-peak fuel rate": "On-Peak Fuel Rate ($/kWh)",
    "Off-peak fuel rate": "Off-Peak Fuel Rate ($/kWh)",
    "Demand charge rate": "Demand Charge Rate ($/kW)",
    "Maximum demand rate": "Maximum Demand Rate ($/kW)",
    "Energy charge": "Energy Charge ($)",
    "Fuel charge": "Fuel Charge ($)",
    "Total energy charge": "Total Energy Charge ($)",
    "Total demand charge": "Total Demand Charge ($)",
    "Total electric cost": "Total Electric Cost ($)",
    "Base charge": "Base Charge ($)",
    "Customer charge": "Customer Charge ($)",
    "Service charge": "Service Charge ($)",
    "Late payment charge": "Late Payment Charge ($)",
    "Total services and tax": "Total Services and Tax ($)",
    "Total charge": "Total Charge ($)",
    "Energy rate": "Average Energy Rate ($/kWh)",
    "Total $/kWh cost": "Total $/kWh Cost",
}

REPORT_ROW_ORDER = [
    "Data source",
    "Estimate method",
    "Confidence",
    "Account number",
    "Rate",
    "Service to date",
] + DETAIL_FIELDS


def _annual_rate(field: str, frame: pd.DataFrame) -> float:
    numerator_denominator = {
        "Non-fuel energy rate": ("Energy charge", "Total consumption kWh"),
        "Fuel rate": ("Fuel charge", "Total consumption kWh"),
        "Energy rate": ("Total energy charge", "Total consumption kWh"),
        "Total $/kWh cost": ("Total energy charge", "Total consumption kWh"),
        "Demand rate": ("Total demand charge", "Total demand"),
    }
    if field in numerator_denominator:
        numerator, denominator = numerator_denominator[field]
        return safe_divide(frame[numerator].sum(), frame[denominator].sum())

    weights = {
        "On-peak non-fuel energy rate": "On-peak kWh used",
        "Off-peak non-fuel energy rate": "Off-peak kWh used",
        "On-peak fuel rate": "On-peak kWh used",
        "Off-peak fuel rate": "Off-peak kWh used",
        "Demand charge rate": "Total demand",
        "Maximum demand rate": "Maximum demand",
    }
    weight_field = weights.get(field)
    if not weight_field or field not in frame or weight_field not in frame:
        return 0.0
    denominator = frame[weight_field].map(numeric).sum()
    numerator = sum(numeric(rate) * numeric(weight) for rate, weight in zip(frame[field], frame[weight_field]))
    return safe_divide(numerator, denominator)


def build_annual_view(frame: pd.DataFrame) -> pd.DataFrame:
    """Transpose a complete monthly frame into the legacy-friendly annual layout."""
    monthly = frame.sort_values("Service month").set_index("Service month")
    rows: dict[str, list[Any]] = {}
    for field in REPORT_ROW_ORDER:
        values: list[Any] = []
        for month in range(1, 13):
            value = monthly.at[month, field] if field in monthly.columns else ""
            if field == "Service to date":
                value = display_date(value)
            values.append(value)

        if field in RATE_FIELDS:
            annual: Any = round(_annual_rate(field, frame), 6)
        elif field in (KWH_FIELDS | DEMAND_FIELDS | CURRENCY_FIELDS) or field == "Service days":
            annual = round(sum(numeric(value) for value in values), 2)
        else:
            annual = ""
        rows[DISPLAY_LABELS.get(field, field)] = values + [annual]

    columns = [MONTH_ABBREVIATIONS[month] for month in range(1, 13)] + ["Sum"]
    return pd.DataFrame.from_dict(rows, orient="index", columns=columns)


def build_estimation_notes(account_frames: dict[str, pd.DataFrame], warnings: list[str] | None = None) -> pd.DataFrame:
    notes: list[dict[str, Any]] = []
    warning_text = "; ".join(warnings or [])
    for account, frame in account_frames.items():
        actual_months = frame.loc[frame["Data source"] == "Actual", "Service month"].astype(int).tolist()
        estimated = frame.loc[frame["Data source"] == "Estimated"]
        missing_months = estimated["Service month"].astype(int).tolist()
        if estimated.empty:
            notes.append(
                {
                    "Account number": account,
                    "Uploaded months": month_list_text(actual_months),
                    "Missing months": "None",
                    "Estimated month": "None",
                    "Estimate method": "No estimation required",
                    "Confidence level": "High",
                    "Notes or warnings": warning_text,
                }
            )
            continue
        for _, row in estimated.sort_values("Service month").iterrows():
            notes.append(
                {
                    "Account number": account,
                    "Uploaded months": month_list_text(actual_months),
                    "Missing months": month_list_text(missing_months),
                    "Estimated month": MONTH_ABBREVIATIONS[int(row["Service month"])],
                    "Estimate method": row["Estimate method"],
                    "Confidence level": row["Confidence"],
                    "Notes or warnings": warning_text,
                }
            )
    return pd.DataFrame(notes)


def _safe_sheet_name(value: str, fallback: str) -> str:
    cleaned = re.sub(r"[\[\]:*?/\\]", "_", str(value)).strip()
    return (cleaned or fallback)[:31]


def _format_annual_sheet(workbook: Any, worksheet: Any, view: pd.DataFrame, frame: pd.DataFrame) -> None:
    header = workbook.add_format({"bold": True, "bg_color": "#005030", "font_color": "white", "align": "center"})
    label = workbook.add_format({"bold": True, "bg_color": "#DDE8E3"})
    estimated = workbook.add_format({"bg_color": "#FFF2CC"})
    actual = workbook.add_format({"bg_color": "#E2F0D9"})
    currency = workbook.add_format({"num_format": "$#,##0.00;[Red]-$#,##0.00"})
    number = workbook.add_format({"num_format": "#,##0"})
    rate = workbook.add_format({"num_format": "0.000000"})

    worksheet.freeze_panes(1, 1)
    worksheet.set_column(0, 0, 42, label)
    worksheet.set_column(1, 13, 14)
    worksheet.set_row(0, None, header)

    source_row = view.index.get_loc("Data Source") + 1
    for month_index, (_, row) in enumerate(frame.sort_values("Service month").iterrows(), start=1):
        worksheet.set_column(month_index, month_index, 14, estimated if row["Data source"] == "Estimated" else actual)
        worksheet.write(source_row, month_index, row["Data source"], estimated if row["Data source"] == "Estimated" else actual)

    for row_index, row_name in enumerate(view.index, start=1):
        if "($)" in row_name or row_name in {DISPLAY_LABELS.get(field, field) for field in CURRENCY_FIELDS}:
            worksheet.set_row(row_index, None, currency)
        elif "Rate" in row_name or "$/kWh" in row_name or "$/kW" in row_name:
            worksheet.set_row(row_index, None, rate)
        elif row_name in {DISPLAY_LABELS.get(field, field) for field in KWH_FIELDS | DEMAND_FIELDS}:
            worksheet.set_row(row_index, None, number)


def generate_excel_report(
    account_frames: dict[str, pd.DataFrame],
    working_hours: pd.DataFrame | None = None,
    warnings: list[str] | None = None,
) -> bytes:
    """Generate the complete annual workbook without writing uploaded data to disk."""
    if not account_frames:
        raise ValueError("No account data is available for reporting.")
    for account, frame in account_frames.items():
        months = set(frame["Service month"].astype(int))
        if months != set(range(1, 13)):
            raise ValueError(f"{account} does not contain a complete 12-month annual view.")

    consolidated = build_consolidated_frame(account_frames)
    recommendations = build_recommendations(
        consolidated,
        len(account_frames),
        calculate_annual_operation_hours(working_hours),
    )
    notes = build_estimation_notes(account_frames, warnings)
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="mmm d, yyyy") as writer:
        consolidated_view = build_annual_view(consolidated)
        consolidated_view.to_excel(writer, sheet_name="Consolidated")
        _format_annual_sheet(writer.book, writer.sheets["Consolidated"], consolidated_view, consolidated)

        used_names = {"Consolidated"}
        for index, (account, frame) in enumerate(account_frames.items(), start=1):
            candidate = _safe_sheet_name(f"Account_{index}_{account}", f"Account_{index}")
            while candidate in used_names:
                candidate = _safe_sheet_name(f"Account_{index}", f"Account_{index}")
            used_names.add(candidate)
            view = build_annual_view(frame)
            view.to_excel(writer, sheet_name=candidate)
            _format_annual_sheet(writer.book, writer.sheets[candidate], view, frame)

        notes.to_excel(writer, sheet_name="Estimation Notes", index=False)
        notes_sheet = writer.sheets["Estimation Notes"]
        notes_sheet.freeze_panes(1, 0)
        notes_sheet.autofilter(0, 0, max(len(notes), 1), max(len(notes.columns) - 1, 0))
        notes_sheet.set_column(0, 0, 18)
        notes_sheet.set_column(1, 2, 36)
        notes_sheet.set_column(3, 3, 18)
        notes_sheet.set_column(4, 4, 55)
        notes_sheet.set_column(5, 6, 24)

        if working_hours is not None and not working_hours.empty:
            working_hours.copy().to_excel(writer, sheet_name="Working Hours", index=False)
            writer.sheets["Working Hours"].set_column(0, len(working_hours.columns) - 1, 20)

        for sheet_name, recommendation in recommendations.items():
            safe_name = _safe_sheet_name(sheet_name, "Recommendation")
            recommendation.to_excel(writer, sheet_name=safe_name, index=False)
            writer.sheets[safe_name].set_column(0, 0, 46)
            writer.sheets[safe_name].set_column(1, 1, 90)

    return output.getvalue()
