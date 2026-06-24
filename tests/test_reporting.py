from datetime import datetime

import pandas as pd
import pytest

from bill_analysis.reporting import build_annual_view


def test_annual_view_sums_months_and_uses_weighted_energy_rate():
    rows = []
    for month in range(1, 13):
        consumption = 100 if month < 12 else 1000
        energy_charge = consumption * (0.10 if month < 12 else 0.20)
        rows.append(
            {
                "Account number": "A-1",
                "Rate": "GSD-1 GENERAL",
                "Service to date": datetime(2024, month, 15),
                "Service month": month,
                "Service year": 2024,
                "Service days": 30,
                "Data source": "Actual",
                "Estimate method": "Actual uploaded bill",
                "Confidence": "High",
                "Total consumption kWh": consumption,
                "Energy charge": energy_charge,
                "Fuel charge": 0,
                "Total energy charge": energy_charge,
                "Total demand": 0,
                "Total demand charge": 0,
                "Energy rate": energy_charge / consumption,
            }
        )
    frame = pd.DataFrame(rows)
    view = build_annual_view(frame)

    assert view.at["Total Consumption kWh", "Sum"] == 2100
    assert view.at["Average Energy Rate ($/kWh)", "Sum"] == pytest.approx(310 / 2100)
    assert view.at["Data Source", "Jan"] == "Actual"
    assert view.at["Estimate Method", "Jan"] == "Actual uploaded bill"
