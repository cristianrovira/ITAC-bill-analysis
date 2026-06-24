"""Core services for the ITAC Bill Analysis Tool.

PDF and Excel dependencies are imported lazily so estimation can be tested in
isolation and lightweight consumers do not need the full Streamlit environment.
"""

from .estimation import detect_missing_months, estimate_missing_months


def extract_bill_data(*args, **kwargs):
    from .extraction import extract_bill_data as implementation

    return implementation(*args, **kwargs)


def extract_uploaded_bills(*args, **kwargs):
    from .extraction import extract_uploaded_bills as implementation

    return implementation(*args, **kwargs)


def generate_excel_report(*args, **kwargs):
    from .reporting import generate_excel_report as implementation

    return implementation(*args, **kwargs)


__all__ = [
    "detect_missing_months",
    "estimate_missing_months",
    "extract_bill_data",
    "extract_uploaded_bills",
    "generate_excel_report",
]
