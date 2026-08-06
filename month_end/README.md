# Month-End Reconciliation

Streamlit app for importing CAD and USD Excel ledgers, normalizing transactions, and performing exact one-to-one reconciliation.

## Run

```text
pip install -r requirements.txt
streamlit run month_end.py
```

## Current behavior

- Scans every workbook sheet for a ledger header row.
- Accepts multiple CAD workbooks and multiple USD workbooks in one run.
- Normalizes to Plant, Currency, Date, Journal, Batch, References, Description, Original Amount, Converted Amount, Status, Match ID, and source metadata.
- Preserves amounts as `Decimal` values; no business rounding is applied.
- Converts CAD to USD using the user-provided FX rate.
- Performs exact one-to-one matching with only a negligible computational tolerance.
- Keeps unmatched CAD and USD pools ready for later one-to-many and many-to-one passes.
- Exposes normalized transactions, match results, and unmatched pools in the app and as CSV downloads.
