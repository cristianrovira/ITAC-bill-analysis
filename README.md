# ITAC Bill Analysis Tool

The ITAC Bill Analysis Tool is a Streamlit application for the University of Miami Industrial Assessment Center. It extracts monthly usage, demand, rate, fee, and tax data from FPL electric bill PDFs and produces a complete annual Excel bill analysis workbook for one or more accounts.

The original application is preserved unchanged in `legacy/billextraction_original.py`. The maintained application separates PDF extraction, calculations, missing-month estimation, reporting, and the Streamlit interface into focused modules.

## Run in GitHub Codespaces

From the repository root:

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Codespaces will display a forwarded-port link for the Streamlit application. Uploaded bills are processed in memory and are not permanently stored by the app.

## Analyze bills

1. Enter the number of FPL accounts.
2. Give each account a unique account number or internal label.
3. Optionally update the working-hours section used by recommendation screening.
4. Upload one or more monthly PDF bills for every account.
5. Select **Analyze uploaded bills**.
6. Review the uploaded-month and missing-month summary.
7. If any month is missing, acknowledge the estimation warning.
8. Generate and download the annual Excel report.

The app reports unreadable PDFs, missing service dates, duplicate months, empty accounts, and extraction failures before report generation.

## Missing-month estimation

Estimation is performed only after uploaded PDF data has been extracted into a normalized monthly table.

- A single missing month between two actual months is the midpoint of those months.
- Consecutive missing months use linear interpolation between the surrounding actual months.
- January and December use circular wraparound when the needed surrounding months are available.
- If only one actual month is available, its values are carried to the other months and marked low confidence.
- Estimated service days use the calendar days in the estimated month.
- Estimated late-payment charges are always zero.

Every monthly record is labeled `Actual` or `Estimated`. The annual sheets include `Data Source`, `Estimate Method`, and `Confidence` rows. The workbook's `Estimation Notes` sheet records uploaded months, missing months, the method used for each estimate, confidence, and warnings.

## Excel workbook

The workbook contains:

- `Consolidated`: a 12-month view across all accounts with annual sums and weighted rates.
- One annual detail sheet per account.
- `Estimation Notes`: an audit trail for actual and estimated coverage.
- `Working Hours`: the schedule used for recommendation screening.
- Recommendation sheets for late fees, load factor, maximum demand, and TOU/load-shifting screening.

Estimated months are shaded and remain visibly labeled. Annual rate values are weighted from their applicable usage or demand rather than summed.

## Run tests

```bash
python -m pytest
```

The tests cover missing-month detection, duplicate detection, isolated and consecutive interpolation, January/December wraparound, source labels, required rounding, calendar service days, and zero estimated late charges.

## Deploy to Streamlit Community Cloud

1. Push this repository to GitHub.
2. In Streamlit Community Cloud, create an app from the repository.
3. Select `app.py` as the entry point.
4. Use the repository's `requirements.txt` for dependencies.
5. Keep the repository and deployed app access restricted to approved ITAC users when handling client data.

No Windows-specific paths or local client-file dependencies are required.

## Known limitations

- PDFs must contain extractable text. Image-only scans require OCR before upload.
- FPL periodically changes bill layouts; new label variants may need to be added to `bill_analysis/extraction.py`.
- The parser cannot validate extracted values against an authoritative FPL data source. Review the workbook before using it in an assessment.
- Recommendation results are preliminary screening estimates inherited from the legacy workflow, not final engineering recommendations or tariff advice.
- A calendar annual view assumes at most one bill per account per service month. Multiple service years are flagged for review.

## Maintenance notes

- Keep `legacy/billextraction_original.py` unchanged as the historical reference.
- Add extraction label variants to `FIELD_LABELS` and focused parser tests rather than adding bill-specific logic to `app.py`.
- Keep estimation independent of PDF parsing; it should always accept normalized monthly data.
- Add any new numeric field to the appropriate rounding and report field sets in `bill_analysis/utils.py`.
- Preserve the `Data Source`, `Estimate Method`, and confidence audit trail when changing reports.
- Never commit client PDFs or generated workbooks.
