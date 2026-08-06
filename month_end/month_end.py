from __future__ import annotations

from decimal import Decimal, InvalidOperation

import streamlit as st

from reconciliation.engine import reconcile
from reconciliation.loaders import load_transactions
from reconciliation.models import ReconciliationConfig
from reconciliation.reports import matches_frame, transactions_frame


st.set_page_config(page_title="Month-End Reconciliation", layout="wide")
st.title("Month-End Reconciliation")
st.caption("Exact one-to-one matching with full precision after FX conversion. Near matches remain unmatched for review.")

with st.sidebar:
    st.header("Inputs")
    rate_text = st.text_input("CAD to USD FX rate", value="1")
    plant = st.text_input("Plant (optional)")
    tolerance_text = st.text_input("Floating-point tolerance", value="0.000001")
    cad_files = st.file_uploader("CAD ledgers", type=["xlsx", "xls"], accept_multiple_files=True, key="cad")
    usd_files = st.file_uploader("USD ledgers", type=["xlsx", "xls"], accept_multiple_files=True, key="usd")

try:
    rate = Decimal(rate_text)
    tolerance = Decimal(tolerance_text)
except InvalidOperation:
    st.error("FX rate and tolerance must be valid decimal numbers.")
    st.stop()

if not cad_files or not usd_files:
    st.info("Upload all CAD and USD files to begin. The loader scans every workbook sheet for ledger headers and ignores blank/balance/subtotal rows.")
    st.stop()

config = ReconciliationConfig(cad_to_usd_rate=rate, floating_tolerance=tolerance)
with st.spinner("Loading and matching ledgers..."):
    cad_transactions = [
        transaction
        for file in cad_files
        for transaction in load_transactions(file.getvalue(), "CAD", config, plant)
    ]
    usd_transactions = [
        transaction
        for file in usd_files
        for transaction in load_transactions(file.getvalue(), "USD", config, plant)
    ]
    result = reconcile(cad_transactions, usd_transactions, config)

if not cad_transactions or not usd_transactions:
    st.error("No transaction rows were detected in one or both uploads. Check that the selected files contain ledger detail rows, not only summary pages.")
    st.stop()

matched_count = len(result["matches"])
st.subheader("Summary")
cols = st.columns(4)
cols[0].metric("CAD transactions", len(cad_transactions))
cols[1].metric("USD transactions", len(usd_transactions))
cols[2].metric("Exact matches", matched_count)
cols[3].metric("Remaining unmatched", len(result["unmatched_cad"]) + len(result["unmatched_usd"]))

tab_all, tab_matches, tab_cad, tab_usd = st.tabs(["All transactions", "Exact matches", "Unmatched CAD", "Unmatched USD"])
with tab_all:
    st.dataframe(transactions_frame(result["transactions"]), use_container_width=True, hide_index=True)
with tab_matches:
    st.dataframe(matches_frame(result["matches"]), use_container_width=True, hide_index=True)
with tab_cad:
    st.dataframe(transactions_frame(result["unmatched_cad"]), use_container_width=True, hide_index=True)
with tab_usd:
    st.dataframe(transactions_frame(result["unmatched_usd"]), use_container_width=True, hide_index=True)

st.download_button("Download normalized transactions", transactions_frame(result["transactions"]).to_csv(index=False), "normalized_transactions.csv", "text/csv")
st.download_button("Download exact matches", matches_frame(result["matches"]).to_csv(index=False), "exact_matches.csv", "text/csv")
