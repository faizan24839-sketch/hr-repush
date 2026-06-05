import streamlit as st
import pandas as pd
from html.parser import HTMLParser
from io import StringIO, BytesIO

# ── page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="HR Repush Builder",
    page_icon="⬡",
    layout="centered",
)

# ── styling ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* dark industrial background */
.stApp {
    background-color: #0d0d0d;
    color: #e0e0e0;
}

h1, h2, h3 {
    font-family: 'IBM Plex Mono', monospace;
    color: #f0f0f0;
}

/* top header bar */
.header-bar {
    border-left: 3px solid #00ff88;
    padding: 0.4rem 1rem;
    margin-bottom: 2rem;
}
.header-bar h1 {
    font-size: 1.4rem;
    margin: 0;
    letter-spacing: 0.08em;
    color: #00ff88;
}
.header-bar p {
    margin: 0.2rem 0 0 0;
    font-size: 0.78rem;
    color: #666;
    font-family: 'IBM Plex Mono', monospace;
}

/* section labels */
.section-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    letter-spacing: 0.15em;
    color: #555;
    text-transform: uppercase;
    margin-bottom: 0.5rem;
}

/* file upload areas */
[data-testid="stFileUploader"] {
    border: 1px solid #222;
    border-radius: 4px;
    background: #111;
    padding: 0.5rem;
}

/* inputs */
[data-testid="stNumberInput"] input,
[data-testid="stTextInput"] input {
    background: #111 !important;
    border: 1px solid #333 !important;
    color: #e0e0e0 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    border-radius: 3px !important;
}

/* button */
.stButton > button {
    background: #00ff88 !important;
    color: #000 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-weight: 600 !important;
    border: none !important;
    border-radius: 3px !important;
    letter-spacing: 0.1em !important;
    padding: 0.6rem 2rem !important;
    width: 100%;
}
.stButton > button:hover {
    background: #00cc6a !important;
}

/* status boxes */
.status-ok {
    background: #001a0d;
    border: 1px solid #00ff88;
    border-radius: 4px;
    padding: 1rem 1.2rem;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.85rem;
    color: #00ff88;
    margin: 1rem 0;
}
.status-warn {
    background: #1a1400;
    border: 1px solid #ffcc00;
    border-radius: 4px;
    padding: 1rem 1.2rem;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.85rem;
    color: #ffcc00;
    margin: 1rem 0;
}
.status-err {
    background: #1a0000;
    border: 1px solid #ff4444;
    border-radius: 4px;
    padding: 1rem 1.2rem;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.85rem;
    color: #ff4444;
    margin: 1rem 0;
}

.mono {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.82rem;
    color: #aaa;
}

/* download button */
[data-testid="stDownloadButton"] > button {
    background: #111 !important;
    color: #00ff88 !important;
    border: 1px solid #00ff88 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-weight: 600 !important;
    border-radius: 3px !important;
    letter-spacing: 0.08em !important;
    width: 100%;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #001a0d !important;
}

/* divider */
hr {
    border-color: #1a1a1a;
    margin: 1.5rem 0;
}
</style>
""", unsafe_allow_html=True)

# ── helpers ─────────────────────────────────────────────────────────────────────

class HTMLTableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows, self.current_row, self.current_cell, self.in_cell = [], [], '', False

    def handle_starttag(self, tag, attrs):
        if tag in ('td', 'th'):
            self.in_cell = True
            self.current_cell = ''
        elif tag == 'tr':
            self.current_row = []

    def handle_endtag(self, tag):
        if tag in ('td', 'th'):
            self.current_row.append(self.current_cell.strip())
            self.in_cell = False
        elif tag == 'tr':
            if self.current_row:
                self.rows.append(self.current_row)

    def handle_data(self, data):
        if self.in_cell:
            self.current_cell += data


def read_source(file, receipt_number):
    df = pd.read_csv(file, sep='|', dtype=str)
    df.columns = df.columns.str.strip()
    df['ReceiptNumber'] = df['ReceiptNumber'].str.strip()
    df = df[df['ReceiptNumber'] == str(receipt_number)].copy()
    df['AmountApplied'] = pd.to_numeric(df['AmountApplied'], errors='coerce')
    return df


def read_receipt_application(file):
    content = file.read().decode('utf-8', errors='ignore')
    parser = HTMLTableParser()
    parser.feed(content)
    if not parser.rows:
        return pd.DataFrame()
    headers = parser.rows[0]
    df = pd.DataFrame(parser.rows[1:], columns=headers)
    return df


def read_claims_headers(file):
    df = pd.read_excel(file, dtype=str)
    df.columns = df.columns.str.strip()
    return df


def build_claims_lookup(claims_df):
    """Sheet2: CLAIM_NUMBER, AMOUNT, CUSTOMER_REF_NUMBER, RECEIPT_NUMBER + derived cols."""
    cols = ['CLAIM_NUMBER', 'AMOUNT', 'CUSTOMER_REF_NUMBER', 'RECEIPT_NUMBER']
    sheet2 = claims_df[cols].copy()
    sheet2['AMOUNT'] = pd.to_numeric(sheet2['AMOUNT'], errors='coerce')
    sheet2['AMOUNT_NEG'] = sheet2['AMOUNT'] * -1
    sheet2['CONCAT_KEY'] = sheet2['AMOUNT_NEG'].apply(
        lambda x: str(x) if x == int(x) else str(x)
    ) + sheet2['CUSTOMER_REF_NUMBER'].fillna('')
    return sheet2


def format_amount_for_key(val):
    """Match Google Sheets concat behaviour: drop .0 for whole numbers."""
    try:
        f = float(val)
        if f == int(f):
            return str(int(f))
        return str(f)
    except Exception:
        return str(val)


def process(source_df, receipt_app_df, claims_df, unapplied_amount):
    results = {}

    # ── Sheet2 lookup ─────────────────────────────────────────────────────────
    sheet2 = build_claims_lookup(claims_df)
    lookup_set = set(sheet2['CONCAT_KEY'].dropna())

    # ── split source: claims (blank TransactionNumber) vs invoices ────────────
    claim_mask = source_df['TransactionNumber'].isna() | (source_df['TransactionNumber'].str.strip() == '')
    claim_lines = source_df[claim_mask].copy()
    invoice_lines = source_df[~claim_mask].copy()

    # ── claims: build concat key and VLOOKUP ─────────────────────────────────
    claim_lines['SRC_CONCAT'] = claim_lines.apply(
        lambda r: format_amount_for_key(r['AmountApplied']) + str(r['CustomerReference']).strip()
        if pd.notna(r['CustomerReference']) else format_amount_for_key(r['AmountApplied']),
        axis=1
    )
    claim_lines['VLOOKUP'] = claim_lines['SRC_CONCAT'].apply(
        lambda k: k if k in lookup_set else None
    )
    pending_claims = claim_lines[claim_lines['VLOOKUP'].isna()].copy()

    # ── invoices: VLOOKUP TransactionNumber vs receipt application col B ──────
    app_ref_col = 'Application Reference'
    applied_refs = set(receipt_app_df[app_ref_col].astype(str).str.strip().tolist()) \
        if app_ref_col in receipt_app_df.columns else set()

    invoice_lines['VLOOKUP'] = invoice_lines['TransactionNumber'].apply(
        lambda t: t if str(t).strip() in applied_refs else None
    )
    pending_invoices = invoice_lines[invoice_lines['VLOOKUP'].isna()].copy()

    # ── drop helper columns before combining ─────────────────────────────────
    src_cols = [c for c in source_df.columns]
    pending_claims_clean = pending_claims[src_cols].copy()
    pending_invoices_clean = pending_invoices[src_cols].copy()

    # ── combine ───────────────────────────────────────────────────────────────
    repush = pd.concat([pending_claims_clean, pending_invoices_clean], ignore_index=True)

    # ── reconciliation ────────────────────────────────────────────────────────
    TOLERANCE = 0.02
    total = repush['AmountApplied'].sum()
    matched = abs(abs(total) - abs(unapplied_amount)) <= TOLERANCE
    dropped_rows = []
    recon_status = 'matched'

    if not matched:
        # drop claim rows from top one by one
        claim_indices = list(pending_claims_clean.index)
        working = repush.copy()
        for idx in range(len(pending_claims_clean)):
            working = working.iloc[1:].reset_index(drop=True)
            dropped_rows.append(pending_claims_clean.iloc[idx])
            new_total = working['AmountApplied'].sum()
            if abs(abs(new_total) - abs(unapplied_amount)) <= TOLERANCE:
                repush = working
                total = new_total
                matched = True
                recon_status = 'matched_after_drop'
                break
        if not matched:
            recon_status = 'unmatched'
            repush = pd.concat([pending_claims_clean, pending_invoices_clean], ignore_index=True)
            total = repush['AmountApplied'].sum()

    results['repush_df'] = repush
    results['total'] = total
    results['unapplied'] = unapplied_amount
    results['recon_status'] = recon_status
    results['dropped_rows'] = dropped_rows
    results['pending_claims_count'] = len(pending_claims_clean)
    results['pending_invoices_count'] = len(pending_invoices_clean)
    results['total_source_claims'] = len(claim_lines)
    results['total_source_invoices'] = len(invoice_lines)
    return results


def to_pipe_txt(df, receipt_number):
    buf = StringIO()
    df.to_csv(buf, sep='|', index=False)
    return buf.getvalue().encode('utf-8'), f"{receipt_number}_Reconstructed_1.txt"


# ── UI ──────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="header-bar">
  <h1>⬡ HR REPUSH BUILDER</h1>
  <p>HighRadius · Oracle AR · Receipt Reconstruction</p>
</div>
""", unsafe_allow_html=True)

# inputs
col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="section-label">Receipt Number</div>', unsafe_allow_html=True)
    receipt_number = st.text_input("", placeholder="e.g. 27173", label_visibility="collapsed")
with col2:
    st.markdown('<div class="section-label">Unapplied Amount</div>', unsafe_allow_html=True)
    unapplied_amount = st.number_input("", value=0.0, format="%.2f", label_visibility="collapsed")

st.markdown("---")

st.markdown('<div class="section-label">Source File (.txt pipe-delimited)</div>', unsafe_allow_html=True)
source_file = st.file_uploader("", type=["txt"], key="source", label_visibility="collapsed")

st.markdown('<div class="section-label">Receipt Application (.xls)</div>', unsafe_allow_html=True)
receipt_app_file = st.file_uploader("", type=["xls"], key="receipt_app", label_visibility="collapsed")

st.markdown('<div class="section-label">Open Claims Headers Extract (.xlsx)</div>', unsafe_allow_html=True)
claims_file = st.file_uploader("", type=["xlsx"], key="claims", label_visibility="collapsed")

st.markdown("---")

run = st.button("BUILD REPUSH FILE")

if run:
    errors = []
    if not receipt_number.strip():
        errors.append("Receipt number is required.")
    if unapplied_amount == 0.0:
        errors.append("Unapplied amount cannot be zero.")
    if not source_file:
        errors.append("Source file is missing.")
    if not receipt_app_file:
        errors.append("Receipt application file is missing.")
    if not claims_file:
        errors.append("Open claims headers file is missing.")

    if errors:
        for e in errors:
            st.markdown(f'<div class="status-err">✗ {e}</div>', unsafe_allow_html=True)
    else:
        with st.spinner("Processing..."):
            try:
                source_df = read_source(source_file, receipt_number.strip())
                if source_df.empty:
                    st.markdown(
                        f'<div class="status-err">✗ No rows found for receipt {receipt_number} in source file.</div>',
                        unsafe_allow_html=True
                    )
                    st.stop()

                receipt_app_df = read_receipt_application(receipt_app_file)
                claims_df = read_claims_headers(claims_file)
                results = process(source_df, receipt_app_df, claims_df, unapplied_amount)

                # ── summary stats ─────────────────────────────────────────────
                st.markdown("### Results")
                c1, c2, c3 = st.columns(3)
                c1.metric("Source rows", len(source_df))
                c2.metric("Pending claims", results['pending_claims_count'])
                c3.metric("Pending invoices", results['pending_invoices_count'])

                c4, c5 = st.columns(2)
                c4.metric("Repush rows", len(results['repush_df']))
                c5.metric("Repush total", f"{results['total']:,.2f}")

                # ── recon status ──────────────────────────────────────────────
                status = results['recon_status']

                if status == 'matched':
                    st.markdown(
                        f'<div class="status-ok">✓ RECONCILED — Repush total matches unapplied amount of {unapplied_amount:,.2f}</div>',
                        unsafe_allow_html=True
                    )

                elif status == 'matched_after_drop':
                    dropped_info = "<br>".join([
                        f"  Dropped: {r['CustomerReference']} | {r['AmountApplied']:,.2f} | {r.get('ClaimReason','')}"
                        for r in results['dropped_rows']
                    ])
                    st.markdown(
                        f'<div class="status-warn">⚠ RECONCILED AFTER DROPPING {len(results["dropped_rows"])} CLAIM ROW(S)<br><br>'
                        f'{dropped_info}<br><br>'
                        f'Repush total: {results["total"]:,.2f} | Unapplied: {unapplied_amount:,.2f}</div>',
                        unsafe_allow_html=True
                    )

                else:
                    st.markdown(
                        f'<div class="status-err">✗ UNRECONCILED — Could not match unapplied amount {unapplied_amount:,.2f} '
                        f'by dropping claim rows. Repush total: {results["total"]:,.2f} | '
                        f'Difference: {abs(abs(results["total"]) - abs(unapplied_amount)):,.2f}<br><br>'
                        f'File generated anyway. Review manually before uploading to SFTP.</div>',
                        unsafe_allow_html=True
                    )

                # ── preview ───────────────────────────────────────────────────
                with st.expander("Preview repush rows"):
                    st.dataframe(
                        results['repush_df'].head(50),
                        use_container_width=True,
                        hide_index=True
                    )

                # ── download ──────────────────────────────────────────────────
                file_bytes, filename = to_pipe_txt(results['repush_df'], receipt_number.strip())
                st.markdown("---")
                st.download_button(
                    label=f"↓ DOWNLOAD  {filename}",
                    data=file_bytes,
                    file_name=filename,
                    mime="text/plain"
                )

            except Exception as ex:
                st.markdown(
                    f'<div class="status-err">✗ Unexpected error: {str(ex)}</div>',
                    unsafe_allow_html=True
                )
                raise ex
