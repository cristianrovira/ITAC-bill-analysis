import pandas as pd
import pytest

from bill_analysis.estimation import DuplicateMonthError, detect_missing_months


def test_detects_missing_months():
    assert detect_missing_months([1, 2, 4, 12]) == [3, 5, 6, 7, 8, 9, 10, 11]


def test_detects_missing_months_from_normalized_frame():
    frame = pd.DataFrame({"Service month": list(range(1, 12))})
    assert detect_missing_months(frame) == [12]


def test_duplicate_months_are_rejected():
    with pytest.raises(DuplicateMonthError, match="February"):
        detect_missing_months([1, 2, 2, 3])
