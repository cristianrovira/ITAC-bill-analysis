"""Bill calculations and conservative recommendation screening calculations."""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any, Iterable

import pandas as pd

from .utils import (
    CURRENCY_FIELDS,
    DEMAND_FIELDS,
    DETAIL_FIELDS,
    KWH_FIELDS,
    RATE_FIELDS,
    numeric,
    round_monthly_data,
    safe_divide,
)


def _set_if_missing(record: dict[str, Any], field: str, value: float) -> None:
    if record.get(field) is None or (numeric(record.get(field)) == 0 and value != 0):
        record[field] = value


def calculate_bill_fields(values: dict[str, Any]) -> dict[str, Any]:
    """Fill derived bill values using the same relationships as the legacy app.

    Values printed on the bill take priority. Calculations are used only when a
    corresponding total is absent, which keeps the parser tolerant of rate layouts.
    """
    record = dict(values)

    consumption = numeric(record.get("Total consumption kWh")) or numeric(record.get("kWh used"))
    _set_if_missing(record, "Total consumption kWh", consumption)
    _set_if_missing(record, "kWh used", consumption)

    on_peak = numeric(record.get("On-peak kWh used"))
    off_peak = numeric(record.get("Off-peak kWh used"))
    if on_peak == 0 and off_peak > 0 and consumption >= off_peak:
        on_peak = consumption - off_peak
        record["On-peak kWh used"] = on_peak
    non_tou = max(consumption - on_peak - off_peak, 0)
    record["Non-TOU consumption kWh"] = non_tou

    demand_kw = max(numeric(record.get("Demand kW")), numeric(record.get("Usage")))
    contract_demand = numeric(record.get("Contract demand"))
    on_peak_demand = numeric(record.get("On-peak demand"))
    maximum_demand = numeric(record.get("Maximum demand"))
    total_demand = numeric(record.get("Total demand"))
    if total_demand == 0:
        total_demand = demand_kw + contract_demand
        if total_demand == 0:
            total_demand = on_peak_demand + maximum_demand
    _set_if_missing(record, "Demand kW", demand_kw)
    _set_if_missing(record, "Usage", demand_kw)
    _set_if_missing(record, "Total demand", total_demand)

    non_fuel_rate = numeric(record.get("Non-fuel energy rate"))
    fuel_rate = numeric(record.get("Fuel rate"))
    on_energy_rate = numeric(record.get("On-peak non-fuel energy rate"))
    off_energy_rate = numeric(record.get("Off-peak non-fuel energy rate"))
    on_fuel_rate = numeric(record.get("On-peak fuel rate"))
    off_fuel_rate = numeric(record.get("Off-peak fuel rate"))

    calculated_energy = (
        (on_peak * on_energy_rate)
        + (off_peak * off_energy_rate)
        + (non_tou * non_fuel_rate)
    )
    if calculated_energy == 0:
        calculated_energy = consumption * non_fuel_rate
    calculated_fuel = (
        (on_peak * on_fuel_rate)
        + (off_peak * off_fuel_rate)
        + (non_tou * fuel_rate)
    )
    if calculated_fuel == 0:
        calculated_fuel = consumption * fuel_rate

    _set_if_missing(record, "Energy charge", calculated_energy)
    _set_if_missing(record, "Fuel charge", calculated_fuel)
    energy_charge = numeric(record.get("Energy charge"))
    fuel_charge = numeric(record.get("Fuel charge"))
    _set_if_missing(record, "Total energy charge", energy_charge + fuel_charge)

    demand_rate = numeric(record.get("Demand charge rate")) or numeric(record.get("Demand rate"))
    maximum_rate = numeric(record.get("Maximum demand rate"))
    calculated_demand_charge = demand_kw * demand_rate
    if on_peak_demand or maximum_demand:
        calculated_demand_charge = on_peak_demand * demand_rate + maximum_demand * maximum_rate
    _set_if_missing(record, "Total demand charge", calculated_demand_charge)

    total_energy = numeric(record.get("Total energy charge"))
    total_demand_charge = numeric(record.get("Total demand charge"))
    _set_if_missing(record, "Total electric cost", total_energy + total_demand_charge)

    detailed_fees = sum(
        numeric(record.get(field))
        for field in (
            "Base charge",
            "Customer charge",
            "Service charge",
            "Franchise fee",
            "Franchise charge",
            "Utility tax",
            "Florida sales tax",
            "County sales tax",
            "Discretionary sales surtax",
            "Late payment charge",
            "FPL SolarTogether charge",
            "FPL SolarTogether credit",
            "Power monitoring premium plus",
        )
    )
    combined_receipts_fee = numeric(record.get("Gross receipts tax / Regulatory fee"))
    if combined_receipts_fee:
        detailed_fees += combined_receipts_fee
    else:
        detailed_fees += numeric(record.get("Gross receipts tax"))
        detailed_fees += numeric(record.get("Regulatory fee"))
    if detailed_fees == 0:
        detailed_fees = numeric(record.get("Taxes and charges"))
    _set_if_missing(record, "Total services and tax", detailed_fees)

    total_electric = numeric(record.get("Total electric cost"))
    services_and_tax = numeric(record.get("Total services and tax"))
    _set_if_missing(record, "Total charge", total_electric + services_and_tax)

    _set_if_missing(record, "Energy rate", safe_divide(total_energy, consumption))
    _set_if_missing(record, "Demand rate", safe_divide(total_demand_charge, total_demand))
    _set_if_missing(record, "Total $/kWh cost", safe_divide(total_energy, consumption))

    for field in DETAIL_FIELDS:
        record.setdefault(field, 0.0)
    return record


def calculate_frame(frame: pd.DataFrame) -> pd.DataFrame:
    records = [calculate_bill_fields(record) for record in frame.to_dict(orient="records")]
    return round_monthly_data(pd.DataFrame(records))


def build_consolidated_frame(account_frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Combine complete account frames and recalculate weighted rate rows."""
    if not account_frames:
        raise ValueError("At least one account is required for consolidation.")

    rows: list[dict[str, Any]] = []
    for month in range(1, 13):
        monthly = [
            frame.loc[frame["Service month"] == month].iloc[0]
            for frame in account_frames.values()
        ]
        row: dict[str, Any] = {
            "Account number": "Consolidated",
            "Service month": month,
            "Service year": int(max(numeric(item.get("Service year")) for item in monthly)),
            "Service to date": max(item.get("Service to date") for item in monthly),
            "Service days": max(numeric(item.get("Service days")) for item in monthly),
        }
        rates = sorted({str(item.get("Rate", "")).strip() for item in monthly if str(item.get("Rate", "")).strip()})
        row["Rate"] = rates[0] if len(rates) == 1 else "Multiple rates"

        for field in (KWH_FIELDS | DEMAND_FIELDS | CURRENCY_FIELDS) - {"Service days"}:
            row[field] = sum(numeric(item.get(field)) for item in monthly)

        estimated_accounts = [
            name for name, item in zip(account_frames, monthly) if item.get("Data source") == "Estimated"
        ]
        row["Data source"] = "Estimated" if estimated_accounts else "Actual"
        row["Estimate method"] = (
            "Includes estimates for " + ", ".join(estimated_accounts)
            if estimated_accounts
            else "All uploaded bills"
        )
        row["Confidence"] = "Low" if any(item.get("Confidence") == "Low" for item in monthly) else (
            "High" if not estimated_accounts else "Medium"
        )
        row = calculate_bill_fields(row)
        row["Non-fuel energy rate"] = safe_divide(row["Energy charge"], row["Total consumption kWh"])
        row["Fuel rate"] = safe_divide(row["Fuel charge"], row["Total consumption kWh"])
        row["Energy rate"] = safe_divide(row["Total energy charge"], row["Total consumption kWh"])
        row["Total $/kWh cost"] = row["Energy rate"]
        row["Demand rate"] = safe_divide(row["Total demand charge"], row["Total demand"])
        row["On-peak non-fuel energy rate"] = safe_divide(
            sum(numeric(item.get("On-peak non-fuel energy rate")) * numeric(item.get("On-peak kWh used")) for item in monthly),
            row["On-peak kWh used"],
        )
        row["Off-peak non-fuel energy rate"] = safe_divide(
            sum(numeric(item.get("Off-peak non-fuel energy rate")) * numeric(item.get("Off-peak kWh used")) for item in monthly),
            row["Off-peak kWh used"],
        )
        row["On-peak fuel rate"] = safe_divide(
            sum(numeric(item.get("On-peak fuel rate")) * numeric(item.get("On-peak kWh used")) for item in monthly),
            row["On-peak kWh used"],
        )
        row["Off-peak fuel rate"] = safe_divide(
            sum(numeric(item.get("Off-peak fuel rate")) * numeric(item.get("Off-peak kWh used")) for item in monthly),
            row["Off-peak kWh used"],
        )
        rows.append(row)

    return round_monthly_data(pd.DataFrame(rows).sort_values("Service month").reset_index(drop=True))


def calculate_annual_operation_hours(working_hours: pd.DataFrame | None) -> float:
    if working_hours is None or working_hours.empty:
        return 0.0
    weekly_hours = 0.0
    for row in working_hours.to_dict(orient="records"):
        if not row.get("Enabled", True):
            continue
        start, end = row.get("From"), row.get("To")
        if not start or not end:
            continue
        start_dt = datetime.combine(datetime.today(), start)
        end_dt = datetime.combine(datetime.today(), end)
        if end_dt <= start_dt:
            end_dt += timedelta(days=1)
        duration = (end_dt - start_dt).total_seconds() / 3600
        weekly_hours += duration * numeric(row.get("Days per week"))
    return weekly_hours * 52


def build_recommendations(
    consolidated: pd.DataFrame,
    account_count: int,
    annual_operation_hours: float = 0.0,
) -> dict[str, pd.DataFrame]:
    """Return transparent screening calculations derived from legacy recommendations."""
    annual_consumption = consolidated["Total consumption kWh"].sum()
    annual_late_fees = consolidated["Late payment charge"].sum()
    annual_demand_charge = consolidated["Total demand charge"].sum()
    average_demand_rate = safe_divide(annual_demand_charge, consolidated["Total demand"].sum())
    demand_reduction_kw = consolidated["Total demand"].sum() * 0.05
    demand_reduction_saving = demand_reduction_kw * average_demand_rate

    on_rate = (consolidated["On-peak non-fuel energy rate"] + consolidated["On-peak fuel rate"])
    off_rate = (consolidated["Off-peak non-fuel energy rate"] + consolidated["Off-peak fuel rate"])
    load_shift_saving = ((on_rate - off_rate).clip(lower=0) * consolidated["On-peak kWh used"] * 0.10).sum()
    meter_saving = annual_demand_charge * 0.10 if account_count > 1 else 0.0

    expected_average_demand = safe_divide(annual_consumption, annual_operation_hours)
    observed_average_demand = consolidated["Total demand"].mean()
    max_demand_saving = max(observed_average_demand - expected_average_demand, 0) * average_demand_rate
    if annual_operation_hours == 0:
        max_demand_saving = 0.0

    def sheet(title: str, formula: str, saving: float, note: str) -> pd.DataFrame:
        return pd.DataFrame(
            {
                "Category": [title, "Formula", "Estimated annual savings", "Applicability", "Notes"],
                "Value": [
                    "Screening recommendation",
                    formula,
                    round(saving, 2),
                    "Applicable" if saving > 0 else "Not applicable",
                    note,
                ],
            }
        )

    return {
        "Pay Electrical Bills On Time": sheet(
            "Avoid late-payment charges",
            "Annual late-payment charges",
            annual_late_fees,
            "Estimated months always use a zero late-payment charge.",
        ),
        "Load Factor": sheet(
            "Reduce monthly demand by 5%",
            "5% of annual monthly demand × weighted demand rate",
            demand_reduction_saving,
            "Preliminary screening estimate; verify operational feasibility.",
        ),
        "Expectation of Max Demand": sheet(
            "Compare observed and operation-hour-based average demand",
            "(Observed average demand − annual kWh / operating hours) × demand rate",
            max_demand_saving,
            "Requires valid working-hours input; otherwise marked not applicable.",
        ),
        "Change Rate Structure to GSD": sheet(
            "Shift 10% of on-peak consumption to off-peak",
            "10% on-peak kWh × positive on/off-peak rate difference",
            load_shift_saving,
            f"Meter-consolidation screening value: ${meter_saving:,.2f}.",
        ),
    }
