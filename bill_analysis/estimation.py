"""Missing-month detection and deterministic surrounding-month estimation."""

from __future__ import annotations

import calendar
from datetime import datetime
from typing import Any, Iterable

import pandas as pd

from .calculations import calculate_frame
from .utils import (
    DETAIL_FIELDS,
    METADATA_FIELDS,
    MONTH_NAMES,
    NUMERIC_FIELDS,
    days_in_month,
    infer_report_year,
    numeric,
    round_monthly_data,
)


class MonthValidationError(ValueError):
    """Base error for invalid or ambiguous service-month data."""


class DuplicateMonthError(MonthValidationError):
    """Raised when more than one uploaded bill maps to the same month."""


def _coerce_months(months_or_frame: Iterable[int] | pd.DataFrame) -> list[int]:
    values = (
        months_or_frame["Service month"].tolist()
        if isinstance(months_or_frame, pd.DataFrame)
        else list(months_or_frame)
    )
    months: list[int] = []
    for value in values:
        try:
            month = int(value)
        except (TypeError, ValueError) as exc:
            raise MonthValidationError(f"Invalid service month: {value!r}") from exc
        if not 1 <= month <= 12:
            raise MonthValidationError(f"Service month must be between 1 and 12; received {month}.")
        months.append(month)
    return months


def validate_unique_months(months_or_frame: Iterable[int] | pd.DataFrame) -> list[int]:
    months = _coerce_months(months_or_frame)
    duplicates = sorted({month for month in months if months.count(month) > 1})
    if duplicates:
        names = ", ".join(MONTH_NAMES[month] for month in duplicates)
        raise DuplicateMonthError(f"Duplicate uploaded bill month(s): {names}.")
    return months


def detect_missing_months(months_or_frame: Iterable[int] | pd.DataFrame) -> list[int]:
    months = validate_unique_months(months_or_frame)
    return [month for month in range(1, 13) if month not in months]


def _forward_distance(start_month: int, end_month: int) -> int:
    distance = (end_month - start_month) % 12
    return distance if distance else 12


def _surrounding_months(month: int, actual_months: set[int]) -> tuple[int | None, int | None]:
    previous = next(
        (candidate for distance in range(1, 13) if (candidate := ((month - distance - 1) % 12) + 1) in actual_months),
        None,
    )
    following = next(
        (candidate for distance in range(1, 13) if (candidate := ((month + distance - 1) % 12) + 1) in actual_months),
        None,
    )
    return previous, following


def _interpolate_value(previous: Any, following: Any, fraction: float) -> float:
    return numeric(previous) + (numeric(following) - numeric(previous)) * fraction


def _actual_labels(frame: pd.DataFrame) -> pd.DataFrame:
    labeled = frame.copy()
    labeled["Data source"] = "Actual"
    labeled["Estimate method"] = "Actual uploaded bill"
    labeled["Confidence"] = "High"
    return labeled


def estimate_missing_months(
    actual_frame: pd.DataFrame,
    report_year: int | None = None,
) -> pd.DataFrame:
    """Return a complete January–December table using circular interpolation.

    Estimation occurs only after extraction has produced one normalized row per
    uploaded month. Actual rows are retained and labeled; estimated late fees are
    always zero.
    """
    if actual_frame.empty:
        raise MonthValidationError("At least one successfully extracted bill is required.")
    if "Service month" not in actual_frame.columns:
        raise MonthValidationError("Normalized bill data is missing the 'Service month' field.")

    actual_month_list = validate_unique_months(actual_frame)
    report_year = int(report_year or infer_report_year(actual_frame))
    actual = _actual_labels(actual_frame)
    actual["Service month"] = actual["Service month"].astype(int)
    actual_by_month = {int(row["Service month"]): row for _, row in actual.iterrows()}
    actual_months = set(actual_month_list)

    all_columns = list(dict.fromkeys(list(actual.columns) + METADATA_FIELDS + DETAIL_FIELDS))
    completed: list[dict[str, Any]] = [actual_by_month[month].to_dict() for month in sorted(actual_months)]

    if len(actual_months) == 1:
        anchor_month = next(iter(actual_months))
        anchor = actual_by_month[anchor_month]
        for month in detect_missing_months(actual_month_list):
            estimate = {column: anchor.get(column, 0.0) for column in all_columns}
            estimate.update(
                {
                    "Service month": month,
                    "Service year": report_year,
                    "Service to date": datetime(report_year, month, days_in_month(report_year, month)),
                    "Service days": days_in_month(report_year, month),
                    "Data source": "Estimated",
                    "Estimate method": "Single-month carry-forward, low confidence",
                    "Confidence": "Low",
                    "Late payment charge": 0.0,
                }
            )
            completed.append(estimate)
    else:
        for month in detect_missing_months(actual_month_list):
            previous_month, following_month = _surrounding_months(month, actual_months)
            if previous_month is None or following_month is None or previous_month == following_month:
                nearest_month = previous_month or following_month
                if nearest_month is None:
                    raise MonthValidationError(f"No source month is available for {MONTH_NAMES[month]}.")
                nearest = actual_by_month[nearest_month]
                estimate = {column: nearest.get(column, 0.0) for column in all_columns}
                method = "Nearest available month, low confidence"
                confidence = "Low"
            else:
                previous = actual_by_month[previous_month]
                following = actual_by_month[following_month]
                distance_to_month = _forward_distance(previous_month, month)
                total_distance = _forward_distance(previous_month, following_month)
                fraction = distance_to_month / total_distance
                estimate = {}
                for column in all_columns:
                    if column in NUMERIC_FIELDS:
                        estimate[column] = _interpolate_value(previous.get(column), following.get(column), fraction)
                    elif column in {"Account number", "Rate"}:
                        estimate[column] = previous.get(column) if fraction <= 0.5 else following.get(column)
                    else:
                        estimate[column] = previous.get(column, following.get(column))
                gap_size = total_distance - 1
                if gap_size == 1:
                    method = f"Interpolated from {MONTH_NAMES[previous_month]} and {MONTH_NAMES[following_month]}"
                else:
                    method = f"Linear interpolation from {MONTH_NAMES[previous_month]} to {MONTH_NAMES[following_month]}"
                confidence = "High"

            estimate.update(
                {
                    "Service month": month,
                    "Service year": report_year,
                    "Service to date": datetime(report_year, month, days_in_month(report_year, month)),
                    "Service days": days_in_month(report_year, month),
                    "Data source": "Estimated",
                    "Estimate method": method,
                    "Confidence": confidence,
                    "Late payment charge": 0.0,
                }
            )
            completed.append(estimate)

    result = pd.DataFrame(completed)
    result = calculate_frame(result)
    result = round_monthly_data(result)
    result["Service month"] = result["Service month"].astype(int)
    result["Service year"] = result["Service year"].astype(int)
    result["Service days"] = result["Service days"].astype(int)
    return result.sort_values("Service month").reset_index(drop=True)
