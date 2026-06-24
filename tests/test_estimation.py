from datetime import datetime

import pandas as pd
import pytest

from bill_analysis.estimation import estimate_missing_months


def actual(month, consumption, demand, charge, late=0.0, rate=0.1234567):
    return {
        "Account number": "A-1",
        "Rate": "GSD-1 GENERAL",
        "Service to date": datetime(2024, month, 15),
        "Service month": month,
        "Service year": 2024,
        "Service days": 30,
        "kWh used": consumption,
        "Total consumption kWh": consumption,
        "Demand kW": demand,
        "Usage": demand,
        "Total demand": demand,
        "Energy charge": charge,
        "Fuel charge": 0.0,
        "Total energy charge": charge,
        "Total demand charge": 0.0,
        "Total electric cost": charge,
        "Total services and tax": 0.0,
        "Total charge": charge,
        "Late payment charge": late,
        "Non-fuel energy rate": rate,
        "Energy rate": rate,
    }


def row_for(frame, month):
    return frame.loc[frame["Service month"] == month].iloc[0]


def test_estimates_one_month_from_surrounding_actual_months():
    frame = pd.DataFrame([actual(2, 100, 10, 20), actual(4, 300, 30, 60)])
    result = estimate_missing_months(frame)
    march = row_for(result, 3)
    assert march["Total consumption kWh"] == 200
    assert march["Total demand"] == 20
    assert march["Estimate method"] == "Interpolated from February and April"


def test_linearly_interpolates_consecutive_missing_months():
    frame = pd.DataFrame([actual(2, 100, 10, 20), actual(5, 400, 40, 80)])
    result = estimate_missing_months(frame)
    assert row_for(result, 3)["Total consumption kWh"] == 200
    assert row_for(result, 4)["Total consumption kWh"] == 300
    assert row_for(result, 3)["Estimate method"] == "Linear interpolation from February to May"


def test_missing_january_uses_december_and_february_wraparound():
    frame = pd.DataFrame([actual(12, 100, 10, 20), actual(2, 300, 30, 60)])
    january = row_for(estimate_missing_months(frame), 1)
    assert january["Total consumption kWh"] == 200
    assert january["Estimate method"] == "Interpolated from December and February"


def test_missing_december_uses_november_and_january_wraparound():
    frame = pd.DataFrame([actual(11, 100, 10, 20), actual(1, 300, 30, 60)])
    december = row_for(estimate_missing_months(frame), 12)
    assert december["Total consumption kWh"] == 200
    assert december["Estimate method"] == "Interpolated from November and January"


def test_actual_and_estimated_rows_are_labeled():
    frame = pd.DataFrame([actual(2, 100, 10, 20), actual(4, 300, 30, 60)])
    result = estimate_missing_months(frame)
    assert row_for(result, 2)["Data source"] == "Actual"
    assert row_for(result, 3)["Data source"] == "Estimated"
    assert row_for(result, 2)["Estimate method"] == "Actual uploaded bill"


def test_rounding_and_late_charge_rules_for_estimated_month():
    frame = pd.DataFrame(
        [
            actual(2, 100.2, 10.4, 1.111, late=5, rate=0.1234567),
            actual(4, 101.2, 11.4, 2.222, late=7, rate=0.2234567),
        ]
    )
    march = row_for(estimate_missing_months(frame), 3)
    assert march["Total consumption kWh"] == 101
    assert march["Total demand"] == 11
    assert march["Total charge"] == pytest.approx(1.67)
    assert march["Non-fuel energy rate"] == pytest.approx(0.173457)
    assert march["Late payment charge"] == 0


def test_estimated_service_days_use_calendar_days():
    frame = pd.DataFrame([actual(1, 100, 10, 20), actual(3, 300, 30, 60)])
    february = row_for(estimate_missing_months(frame, report_year=2024), 2)
    assert february["Service days"] == 29


def test_single_actual_month_uses_low_confidence_carry_forward():
    result = estimate_missing_months(pd.DataFrame([actual(6, 120, 12, 24)]))
    january = row_for(result, 1)
    assert january["Total consumption kWh"] == 120
    assert january["Confidence"] == "Low"
    assert january["Estimate method"] == "Single-month carry-forward, low confidence"
