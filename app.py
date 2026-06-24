"""Streamlit entry point for the ITAC Bill Analysis Tool."""

from __future__ import annotations

from datetime import time
from pathlib import Path

import pandas as pd
import streamlit as st

from bill_analysis.estimation import DuplicateMonthError, detect_missing_months, estimate_missing_months
from bill_analysis.extraction import BillExtractionError, extract_bill_data
from bill_analysis.reporting import generate_excel_report
from bill_analysis.utils import MONTH_NAMES, month_list_text

CONFIRMATION_TEXT = "I understand that missing months will be estimated using nearby available months."


def _working_hours_inputs() -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    with st.expander("Working hours (used for recommendation screening)"):
        st.caption("Enable each operating shift and provide its typical schedule. Overnight shifts are supported.")
        defaults = [
            (True, time(7, 0), time(15, 0), 5),
            (False, time(15, 0), time(23, 0), 5),
            (False, time(23, 0), time(7, 0), 5),
        ]
        for index, (default_enabled, default_start, default_end, default_days) in enumerate(defaults, start=1):
            enabled_col, start_col, end_col, days_col = st.columns([1.2, 1, 1, 1])
            enabled = enabled_col.checkbox(f"Shift {index}", value=default_enabled, key=f"shift_{index}_enabled")
            start = start_col.time_input("From", value=default_start, key=f"shift_{index}_start", disabled=not enabled)
            end = end_col.time_input("To", value=default_end, key=f"shift_{index}_end", disabled=not enabled)
            days = days_col.number_input(
                "Days/week", min_value=0, max_value=7, value=default_days,
                key=f"shift_{index}_days", disabled=not enabled,
            )
            rows.append(
                {"Shift": f"Shift {index}", "Enabled": enabled, "From": start, "To": end, "Days per week": int(days)}
            )
    return pd.DataFrame(rows)


def _analyze_uploads(
    uploads_by_account: dict[str, list[object]],
) -> tuple[dict[str, pd.DataFrame], list[str]]:
    account_frames: dict[str, pd.DataFrame] = {}
    warnings: list[str] = []
    errors: list[str] = []

    for account, uploaded_files in uploads_by_account.items():
        records: list[dict[str, object]] = []
        if not uploaded_files:
            errors.append(f"{account}: upload at least one PDF bill.")
            continue
        for uploaded_file in uploaded_files:
            try:
                records.append(extract_bill_data(uploaded_file, account_number=account))
            except BillExtractionError as exc:
                errors.append(str(exc))
        if not records:
            continue

        frame = pd.DataFrame(records).sort_values("Service month").reset_index(drop=True)
        try:
            detect_missing_months(frame)
        except DuplicateMonthError as exc:
            errors.append(f"{account}: {exc}")
            continue

        service_years = sorted(frame["Service year"].dropna().astype(int).unique())
        if len(service_years) > 1:
            warnings.append(
                f"{account} contains service dates from multiple years ({', '.join(map(str, service_years))}); "
                "estimated dates use the most common service year."
            )
        account_frames[account] = frame

    if errors:
        raise BillExtractionError("\n".join(errors))
    return account_frames, warnings


def _summary_frame(account_frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows = []
    for account, frame in account_frames.items():
        uploaded = sorted(frame["Service month"].astype(int).tolist())
        missing = detect_missing_months(uploaded)
        rows.append(
            {
                "Account": account,
                "Uploaded bills": len(uploaded),
                "Uploaded months": month_list_text(uploaded),
                "Missing months": month_list_text(missing),
            }
        )
    return pd.DataFrame(rows)


def main() -> None:
    st.set_page_config(page_title="ITAC Bill Analysis Tool", page_icon="⚡", layout="wide")
    logo_path = Path(__file__).parent / "assets" / "Logo-University-of-Miami.jpg"
    if logo_path.exists():
        st.image(str(logo_path), width=260)

    st.title("ITAC Bill Analysis Tool")
    st.write(
        "Upload FPL electric bill PDFs to extract monthly usage and charges, identify missing months, "
        "and generate a labeled 12-month Excel bill analysis report."
    )

    working_hours = _working_hours_inputs()
    st.subheader("Accounts and bill uploads")
    number_of_accounts = int(st.number_input("Number of FPL accounts", min_value=1, max_value=25, value=1, step=1))

    uploads_by_account: dict[str, list[object]] = {}
    seen_account_names: set[str] = set()
    duplicate_account_name = False
    for index in range(1, number_of_accounts + 1):
        with st.container(border=True):
            account = st.text_input(
                f"Account {index} identifier", value=f"Account {index}", key=f"account_{index}_identifier",
                help="Use the FPL account number or a non-sensitive internal label.",
            ).strip() or f"Account {index}"
            if account in seen_account_names:
                duplicate_account_name = True
            seen_account_names.add(account)
            uploads = st.file_uploader(
                f"Upload PDF bills for {account}", type=["pdf"], accept_multiple_files=True,
                key=f"account_{index}_uploads",
            )
            uploads_by_account[account] = list(uploads or [])

    if duplicate_account_name:
        st.error("Each account identifier must be unique.")

    if st.button("Analyze uploaded bills", type="primary", disabled=duplicate_account_name):
        st.session_state.pop("report_bytes", None)
        try:
            with st.spinner("Reading and validating FPL bills..."):
                account_frames, warnings = _analyze_uploads(uploads_by_account)
            st.session_state["account_frames"] = account_frames
            st.session_state["analysis_warnings"] = warnings
        except BillExtractionError as exc:
            st.session_state.pop("account_frames", None)
            for message in str(exc).splitlines():
                st.error(message)

    account_frames = st.session_state.get("account_frames")
    warnings = st.session_state.get("analysis_warnings", [])
    if not account_frames:
        return

    st.subheader("Annual coverage review")
    for warning in warnings:
        st.warning(warning)
    st.dataframe(_summary_frame(account_frames), hide_index=True, use_container_width=True)

    has_missing_months = any(detect_missing_months(frame) for frame in account_frames.values())
    if has_missing_months:
        st.warning(
            "One or more accounts contain fewer than 12 uploaded months. Missing months must be estimated "
            "before a complete annual report can be generated."
        )
        confirmed = st.checkbox(CONFIRMATION_TEXT)
    else:
        st.success("All accounts contain 12 unique uploaded months. No estimation is required.")
        confirmed = True

    if st.button("Generate annual Excel report", disabled=not confirmed):
        try:
            complete_frames = {account: estimate_missing_months(frame) for account, frame in account_frames.items()}
            with st.spinner("Building the annual workbook..."):
                report = generate_excel_report(complete_frames, working_hours, warnings)
            st.session_state["report_bytes"] = report
            st.session_state["complete_frames"] = complete_frames
        except Exception as exc:
            st.error(f"Report generation failed: {exc}")

    report_bytes = st.session_state.get("report_bytes")
    if report_bytes:
        st.download_button(
            "Download annual bill analysis", data=report_bytes, file_name="ITAC_annual_bill_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary",
        )

        st.subheader("Generated monthly data")
        for account, frame in st.session_state.get("complete_frames", {}).items():
            with st.expander(account):
                preview_columns = [
                    "Service month", "Data source", "Estimate method", "Total consumption kWh",
                    "Total demand", "Total charge",
                ]
                preview = frame[preview_columns].copy()
                preview["Service month"] = preview["Service month"].map(MONTH_NAMES)
                st.dataframe(preview, hide_index=True, use_container_width=True)


if __name__ == "__main__":
    main()
