import streamlit as st
import pandas as pd
from html.parser import HTMLParser
from io import StringIO, BytesIO
import zipfile
import json

st.set_page_config(page_title="HR Repush Builder", page_icon="⬡", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.stApp { background-color: #0d0d0d; color: #e0e0e0; }
h1, h2, h3 { font-family: 'IBM Plex Mono', monospace; color: #f0f0f0; }
.header-bar { border-left: 3px solid #00ff88; padding: 0.4rem 1rem; margin-bottom: 2rem; }
.header-bar h1 { font-size: 1.4rem; margin: 0; letter-spacing: 0.08em; color: #00ff88; }
.header-bar p { margin: 0.2rem 0 0 0; font-size: 0.78rem; color: #666; font-family: 'IBM Plex Mono', monospace; }
.section-label { font-family: 'IBM Plex Mono', monospace; font-size: 0.7rem; letter-spacing: 0.15em; color: #555; text-transform: uppercase; margin-bottom: 0.5rem; }
.receipt-card { border: 1px solid #222; border-radius: 4px; background: #111; padding: 1rem 1.2rem; margin-bottom: 1rem; }
.receipt-card-title { font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; color: #00ff88; letter-spacing: 0.1em; margin-bottom: 0.8rem; }
[data-testid="stFileUploader"] { border: 1px solid #222; border-radius: 4px; background: #111; padding: 0.5rem; }
[data-testid="stNumberInput"] input, [data-testid="stTextInput"] input { background: #111 !important; border: 1px solid #333 !important; color: #e0e0e0 !important; font-family: 'IBM Plex Mono', monospace !important; border-radius: 3px !important; }
.stButton > button { background: #00ff88 !important; color: #000 !important; font-family: 'IBM Plex Mono', monospace !important; font-weight: 600 !important; border: none !important; border-radius: 3px !important; letter-spacing: 0.1em !important; padding: 0.6rem 2rem !important; width: 100%; }
.stButton > button:hover { background: #00cc6a !important; }
.status-ok { background: #001a0d; border: 1px solid #00ff88; border-radius: 4px; padding: 1rem 1.2rem; font-family: 'IBM Plex Mono', monospace; font-size: 0.85rem; color: #00ff88; margin: 1rem 0; }
.status-warn { background: #1a1400; border: 1px solid #ffcc00; border-radius: 4px; padding: 1rem 1.2rem; font-family: 'IBM Plex Mono', monospace; font-size: 0.85rem; color: #ffcc00; margin: 1rem 0; }
.status-err { background: #1a0000; border: 1px solid #ff4444; border-radius: 4px; padding: 1rem 1.2rem; font-family: 'IBM Plex Mono', monospace; font-size: 0.85rem; color: #ff4444; margin: 1rem 0; }
[data-testid="stDownloadButton"] > button { background: #111 !important; color: #00ff88 !important; border: 1px solid #00ff88 !important; font-family: 'IBM Plex Mono', monospace !important; font-weight: 600 !important; border-radius: 3px !important; letter-spacing: 0.08em !important; width: 100%; }
[data-testid="stDownloadButton"] > button:hover { background: #001a0d !important; }
hr { border-color: #1a1a1a; margin: 1.5rem 0; }
div[data-testid="stSelectbox"] > div > div { background: #111 !important; border: 1px solid #333 !important; color: #e0e0e0 !important; }
</style>
""", unsafe_allow_html=True)


# ── helpers ────────────────────────────────────────────────────────────────────

class HTMLTableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows, self.current_row, self.current_cell, self.in_cell = [], [], '', False
    def handle_starttag(self, tag, attrs):
        if tag in ('td', 'th'): self.in_cell = True; self.current_cell = ''
        elif tag == 'tr': self.current_row = []
    def handle_endtag(self, tag):
        if tag in ('td', 'th'): self.current_row.append(self.current_cell.strip()); self.in_cell = False
        elif tag == 'tr':
            if self.current_row: self.rows.append(self.current_row)
    def handle_data(self, data):
        if self.in_cell: self.current_cell += data


def read_source(file, receipt_number):
    df = pd.read_csv(file, sep='|', dtype=str)
    df.columns = df.columns.str.strip()
    df['ReceiptNumber'] = df['ReceiptNumber'].str.strip()
    df = df[df['ReceiptNumber'] == str(receipt_number).strip()].copy()
    df['AmountApplied'] = pd.to_numeric(df['AmountApplied'], errors='coerce')
    return df


def read_receipt_application(file):
    content = file.read().decode('utf-8', errors='ignore')
    parser = HTMLTableParser()
    parser.feed(content)
    if not parser.rows: return pd.DataFrame()
    headers = parser.rows[0]
    df = pd.DataFrame(parser.rows[1:], columns=headers)
    return df


def read_claims_headers(file):
    df = pd.read_excel(file, dtype=str)
    df.columns = df.columns.str.strip()
    return df


def build_claims_lookup(claims_df):
    cols = ['CLAIM_NUMBER', 'AMOUNT', 'CUSTOMER_REF_NUMBER', 'RECEIPT_NUMBER']
    sheet2 = claims_df[cols].copy()
    sheet2['AMOUNT'] = pd.to_numeric(sheet2['AMOUNT'], errors='coerce')
    sheet2['AMOUNT_NEG'] = sheet2['AMOUNT'] * -1
    sheet2['CONCAT_KEY'] = sheet2['AMOUNT_NEG'].apply(
        lambda x: str(int(x)) if x == int(x) else str(x)
    ) + sheet2['CUSTOMER_REF_NUMBER'].fillna('')
    return sheet2


def format_amount_for_key(val):
    try:
        f = float(val)
        if f == int(f): return str(int(f))
        return str(f)
    except: return str(val)


def get_pending_claims(source_df, claims_df):
    sheet2 = build_claims_lookup(claims_df)
    lookup_set = set(sheet2['CONCAT_KEY'].dropna())
    claim_mask = source_df['TransactionNumber'].isna() | (source_df['TransactionNumber'].str.strip() == '')
    claim_lines = source_df[claim_mask].copy()
    claim_lines['SRC_CONCAT'] = claim_lines.apply(
        lambda r: format_amount_for_key(r['AmountApplied']) + str(r['CustomerReference']).strip()
        if pd.notna(r['CustomerReference']) else format_amount_for_key(r['AmountApplied']), axis=1)
    claim_lines['VLOOKUP'] = claim_lines['SRC_CONCAT'].apply(lambda k: k if k in lookup_set else None)
    pending = claim_lines[claim_lines['VLOOKUP'].isna()].copy()
    src_cols = [c for c in source_df.columns]
    return pending[src_cols].copy()


def get_pending_invoices_standalone(source_df, receipt_app_df):
    invoice_mask = ~(source_df['TransactionNumber'].isna() | (source_df['TransactionNumber'].str.strip() == ''))
    invoice_lines = source_df[invoice_mask].copy()
    app_ref_col = 'Application Reference'
    applied_refs = set(receipt_app_df[app_ref_col].astype(str).str.strip().tolist()) if app_ref_col in receipt_app_df.columns else set()
    invoice_lines['VLOOKUP'] = invoice_lines['TransactionNumber'].apply(lambda t: t if str(t).strip() in applied_refs else None)
    pending = invoice_lines[invoice_lines['VLOOKUP'].isna()].copy()
    src_cols = [c for c in source_df.columns]
    return pending[src_cols].copy()


def reconcile_and_drop(pending_claims, pending_invoices, unapplied_amount):
    TOLERANCE = 0.02
    src_cols = list(pending_claims.columns) if len(pending_claims) > 0 else list(pending_invoices.columns)
    repush = pd.concat([pending_claims, pending_invoices], ignore_index=True)
    total = repush['AmountApplied'].sum()
    matched = abs(abs(total) - abs(unapplied_amount)) <= TOLERANCE
    dropped_rows = []
    recon_status = 'matched'

    if not matched and len(pending_claims) > 0:
        working = repush.copy()
        for idx in range(len(pending_claims)):
            working = working.iloc[1:].reset_index(drop=True)
            dropped_rows.append(pending_claims.iloc[idx])
            new_total = working['AmountApplied'].sum()
            if abs(abs(new_total) - abs(unapplied_amount)) <= TOLERANCE:
                repush = working; total = new_total; matched = True; recon_status = 'matched_after_drop'; break
        if not matched:
            recon_status = 'unmatched'
            repush = pd.concat([pending_claims, pending_invoices], ignore_index=True)
            total = repush['AmountApplied'].sum()

    return repush, total, recon_status, dropped_rows


def to_pipe_txt(df, receipt_number, iteration=1):
    buf = StringIO()
    df.to_csv(
        buf,
        sep='|',
        index=False,
        lineterminator='\r\n',
        float_format='%.2f',
    )
    return buf.getvalue().encode('utf-8'), f"{receipt_number}_Reconstructed_{iteration}.txt"


# ── receipt creation ──────────────────────────────────────────────────────────

def build_receipt_payloads(src_full, receipt_number):
    """Given the full source dataframe and a target receipt number, return a dict
    with parent payload, child payloads, and R2R lines for any DS child."""
    sub = src_full[src_full['ReceiptNumber'].str.strip() == str(receipt_number).strip()].copy()
    if sub.empty:
        return None

    h = sub.iloc[0]
    paying = h['CustomerAccountNumber'].strip()

    # children in order of first appearance
    seen = []
    for acct in sub['Item Account Number']:
        a = acct.strip() if isinstance(acct, str) else ''
        if a and a not in seen:
            seen.append(a)
    children = [a for a in seen if a != paying]

    def _amount_from_header(v):
        try: return float(v)
        except: return v

    parent = {
        "AccountingDate": h['AccountingDate'],
        "RemittanceBankAccountNumber": h['BankAccountNumber'],
        "ConversionRateType": h['ConversionRateType'],
        "Currency": h['CurrencyCode'],
        "CustomerName": h['CustomerName'],
        "ReceiptNumber": str(receipt_number).strip(),
        "Amount": _amount_from_header(h['ReceiptAmount']),
        "CustomerAccountNumber": paying,
        "BusinessUnit": h['BusinessUnit'],
        "StructuredPaymentReference": h['StructuredPaymentReference'],
        "ReceiptDate": h['ReceiptDate'],
        "ConversionDate": h['ConversionDate'],
    }

    child_payloads = []
    r2r_lines = []
    for idx, ch in enumerate(children, start=1):
        ch_rows = sub[sub['Item Account Number'].str.strip() == ch]
        ch_name = ch_rows['Item Customer Name'].iloc[0]
        child_rn = f"{str(receipt_number).strip()}_{idx}"
        child_payloads.append({
            "AccountingDate": h['AccountingDate'],
            "RemittanceBankAccountNumber": h['BankAccountNumber'],
            "ConversionRateType": h['ConversionRateType'],
            "Currency": h['CurrencyCode'],
            "CustomerName": ch_name,
            "ReceiptNumber": child_rn,
            "Amount": 0,
            "CustomerAccountNumber": ch,
            "BusinessUnit": h['BusinessUnit'],
            "StructuredPaymentReference": h['StructuredPaymentReference'],
            "ReceiptDate": h['ReceiptDate'],
            "ConversionDate": h['ConversionDate'],
        })
        if ch.endswith('DS'):
            child_amount = pd.to_numeric(ch_rows['AmountApplied'], errors='coerce').sum()
            r2r_lines.append({
                "BU": h['BusinessUnit'],
                "ReceiptNumber": str(receipt_number).strip(),
                "ChildReceiptNumber": child_rn,
                "ChildAmount": round(float(child_amount), 2),
                "Date": h['AccountingDate'],
            })

    return {"parent": parent, "children": child_payloads, "r2r": r2r_lines}


# ── parent/child core ──────────────────────────────────────────────────────────
# A claim line has a BLANK TransactionNumber and a POPULATED SNInvoiceNumber.
# A transaction line has a POPULATED TransactionNumber and a BLANK SNInvoiceNumber.
# The TransactionNumber on a txn line equals the SNInvoiceNumber on the claim lines
# it settles, which is the key that ties them together.

def _txn_mask(s):
    return ~(s['TransactionNumber'].isna() | (s['TransactionNumber'].str.strip() == ''))


def build_sum_sn(child_src):
    """sum_sn: from a child's CLAIM rows, group by SNInvoiceNumber, sum AmountApplied.
    Values stay negative. Used to derive each child transaction's pushable amount."""
    claims = child_src[child_src['TransactionNumber'].isna() | (child_src['TransactionNumber'].str.strip() == '')]
    if len(claims) == 0:
        return {}
    return claims.groupby(claims['SNInvoiceNumber'].str.strip())['AmountApplied'].sum().to_dict()


def build_child_pending(receipt_src, claims_df, sn_map, applied_refs, rn, acct):
    """Child file rows:
       claims  -> pending claims, original source columns/amounts, untouched
       txns    -> AmountApplied overwritten with (-1 * sum_sn[TransactionNumber]); 0 if no match
       Both VLOOKUP'd against the CHILD receipt application; NA rows kept.
       ReceiptNumber (col C), CustomerAccountNumber (col L), CustomerName (col K)
       stamped to child values (CustomerName pulled from Item Customer Name col Z)."""
    src_cols = list(receipt_src.columns)
    pending_claims = get_pending_claims(receipt_src, claims_df).copy()

    txn = receipt_src[_txn_mask(receipt_src)].copy()
    txn['AmountApplied'] = txn['TransactionNumber'].str.strip().map(lambda t: -sn_map.get(t, 0))
    txn['__vlk'] = txn['TransactionNumber'].apply(lambda t: t if str(t).strip() in applied_refs else None)
    pending_invoices = txn[txn['__vlk'].isna()][src_cols].copy()

    for d in (pending_claims, pending_invoices):
        if len(d) > 0:
            d['ReceiptNumber'] = rn
            d['CustomerAccountNumber'] = acct
            d['CustomerName'] = d['Item Customer Name']
    return pending_claims, pending_invoices


def build_parent_pending(receipt_src, claims_df, applied_refs, parent_rn, parent_acct,
                         child_entries, src_full):
    """Parent file rows:
       claims         -> parent's own pending claims, original columns/amounts
       own txns       -> parent's own transaction lines, ORIGINAL amounts, VLOOKUP vs parent RAP
       child residual -> for each child with txn rows: new O = original O + sum_sn[txn] (O+Q),
                         no dedup, VLOOKUP vs PARENT RAP, stamped with parent receipt/account.
       Claims-only children (no txn rows, e.g. DS accounts) are skipped here."""
    src_cols = list(receipt_src.columns)
    pending_claims = get_pending_claims(receipt_src, claims_df).copy()

    own = receipt_src[_txn_mask(receipt_src)].copy()
    own['__vlk'] = own['TransactionNumber'].apply(lambda t: t if str(t).strip() in applied_refs else None)
    frames = [own[own['__vlk'].isna()][src_cols].copy()]

    for ce in child_entries:
        sn_map, cacct = ce['sn_map'], ce['acct']
        ctxn = src_full[src_full['Item Account Number'].str.strip() == cacct].copy()
        ctxn = ctxn[_txn_mask(ctxn)].copy()
        if len(ctxn) == 0:
            continue  # claims-only child, no residual contribution to parent
        ctxn['__q'] = ctxn['TransactionNumber'].str.strip().map(lambda t: sn_map.get(t, 0))
        ctxn['AmountApplied'] = ctxn['AmountApplied'] + ctxn['__q']
        ctxn['__vlk'] = ctxn['TransactionNumber'].apply(lambda t: t if str(t).strip() in applied_refs else None)
        keep = ctxn[ctxn['__vlk'].isna()][src_cols].copy()
        keep['ReceiptNumber'] = parent_rn
        keep['CustomerAccountNumber'] = parent_acct
        frames.append(keep)

    pending_invoices = pd.concat(frames, ignore_index=True)
    if len(pending_claims) > 0:
        pending_claims['ReceiptNumber'] = parent_rn
        pending_claims['CustomerAccountNumber'] = parent_acct
    return pending_claims, pending_invoices


def zip_outputs(output_files):
    """Bundle (bytes, filename) tuples into a single zip for one-click download."""
    buf = BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file_bytes, filename in output_files:
            zf.writestr(filename, file_bytes)
    buf.seek(0)
    return buf.getvalue()


def recon_status_html(status, total, unapplied, dropped_rows, label=""):
    prefix = f"[{label}] " if label else ""
    if status == 'matched':
        return f'<div class="status-ok">✓ {prefix}RECONCILED — Total {total:,.2f} matches unapplied {unapplied:,.2f}</div>'
    elif status == 'matched_after_drop':
        dropped_info = "<br>".join([
            f"  Dropped: {r.get('CustomerReference','')} | {r['AmountApplied']:,.2f} | {r.get('ClaimReason','')}"
            for r in dropped_rows
        ])
        return f'<div class="status-warn">⚠ {prefix}RECONCILED AFTER DROPPING {len(dropped_rows)} CLAIM ROW(S)<br><br>{dropped_info}<br><br>Total: {total:,.2f} | Unapplied: {unapplied:,.2f}</div>'
    else:
        return f'<div class="status-err">✗ {prefix}UNRECONCILED — Total: {total:,.2f} | Unapplied: {unapplied:,.2f} | Gap: {abs(abs(total)-abs(unapplied)):,.2f}<br>File generated anyway. Review before uploading to SFTP.</div>'


# ── UI ─────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="header-bar">
  <h1>⬡ HR REPUSH BUILDER</h1>
  <p>HighRadius · Oracle AR · Receipt Reconstruction</p>
</div>
""", unsafe_allow_html=True)

# mode selector
st.markdown('<div class="section-label">Receipt Type</div>', unsafe_allow_html=True)
mode = st.radio("", ["Standalone", "Parent / Child", "Receipt Creation"], horizontal=True, label_visibility="collapsed")

st.markdown("---")

# shared inputs
st.markdown('<div class="section-label">Source File (.txt pipe-delimited)</div>', unsafe_allow_html=True)
source_file = st.file_uploader("", type=["txt"], key="source", label_visibility="collapsed")

if mode in ("Standalone", "Parent / Child"):
    st.markdown('<div class="section-label">Open Claims Headers Extract (.xlsx)</div>', unsafe_allow_html=True)
    claims_file = st.file_uploader("", type=["xlsx"], key="claims", label_visibility="collapsed")
else:
    claims_file = None

st.markdown("---")

# ── STANDALONE MODE ────────────────────────────────────────────────────────────
if mode == "Standalone":
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="section-label">Receipt Number</div>', unsafe_allow_html=True)
        receipt_number = st.text_input("", placeholder="e.g. 27173", key="sa_rn", label_visibility="collapsed")
    with col2:
        st.markdown('<div class="section-label">Unapplied Amount</div>', unsafe_allow_html=True)
        unapplied_amount = st.number_input("", value=0.0, format="%.2f", key="sa_ua", label_visibility="collapsed")

    st.markdown('<div class="section-label">Receipt Application (.xls)</div>', unsafe_allow_html=True)
    receipt_app_file = st.file_uploader("", type=["xls"], key="sa_rap", label_visibility="collapsed")

    st.markdown("---")
    run = st.button("BUILD REPUSH FILE")

    if run:
        errors = []
        if not receipt_number.strip(): errors.append("Receipt number required.")
        if unapplied_amount == 0.0: errors.append("Unapplied amount cannot be zero.")
        if not source_file: errors.append("Source file missing.")
        if not receipt_app_file: errors.append("Receipt application file missing.")
        if not claims_file: errors.append("Claims headers file missing.")

        for e in errors:
            st.markdown(f'<div class="status-err">✗ {e}</div>', unsafe_allow_html=True)

        if not errors:
            with st.spinner("Processing..."):
                try:
                    src = read_source(source_file, receipt_number.strip())
                    if src.empty:
                        st.markdown(f'<div class="status-err">✗ No rows found for receipt {receipt_number}.</div>', unsafe_allow_html=True)
                        st.stop()

                    rap = read_receipt_application(receipt_app_file)
                    clm = read_claims_headers(claims_file)

                    pending_claims = get_pending_claims(src, clm)
                    pending_invoices = get_pending_invoices_standalone(src, rap)
                    repush, total, status, dropped = reconcile_and_drop(pending_claims, pending_invoices, unapplied_amount)

                    st.markdown("### Results")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Source rows", len(src))
                    c2.metric("Pending claims", len(pending_claims))
                    c3.metric("Pending invoices", len(pending_invoices))
                    c4, c5 = st.columns(2)
                    c4.metric("Repush rows", len(repush))
                    c5.metric("Repush total", f"{total:,.2f}")

                    st.markdown(recon_status_html(status, total, unapplied_amount, dropped), unsafe_allow_html=True)

                    with st.expander("Preview repush rows"):
                        st.dataframe(repush.head(50), use_container_width=True, hide_index=True)

                    file_bytes, filename = to_pipe_txt(repush, receipt_number.strip())
                    st.markdown("---")
                    st.download_button(f"↓ DOWNLOAD  {filename}", file_bytes, filename, mime="text/plain")

                except Exception as ex:
                    st.markdown(f'<div class="status-err">✗ Error: {str(ex)}</div>', unsafe_allow_html=True)
                    raise ex

# ── PARENT / CHILD MODE ────────────────────────────────────────────────────────
elif mode == "Parent / Child":
    # init session state
    if 'pc_receipts' not in st.session_state:
        st.session_state.pc_receipts = [
            {'label': 'Parent', 'receipt_number': '', 'customer_account': '', 'unapplied': 0.0}
        ]

    st.markdown("### Receipts")
    st.markdown('<p style="font-size:0.8rem;color:#555;font-family:\'IBM Plex Mono\',monospace;">Add one entry per receipt as seen in Oracle. Children first, parent last.</p>', unsafe_allow_html=True)

    receipts = st.session_state.pc_receipts
    rap_files = {}

    for i, r in enumerate(receipts):
        label = r['label']
        st.markdown(f'<div class="receipt-card-title">{"▸ " + label.upper()}</div>', unsafe_allow_html=True)

        c1, c2, c3 = st.columns([2, 2, 2])
        with c1:
            receipts[i]['label'] = st.text_input("Label", value=r['label'], key=f"lbl_{i}", placeholder="e.g. Child 1")
        with c2:
            receipts[i]['receipt_number'] = st.text_input("Receipt Number", value=r['receipt_number'], key=f"rn_{i}", placeholder="e.g. 27173_1")
        with c3:
            receipts[i]['customer_account'] = st.text_input("Customer Account", value=r['customer_account'], key=f"ca_{i}", placeholder="e.g. COSTCA")

        receipts[i]['unapplied'] = st.number_input("Unapplied Amount", value=float(r['unapplied']), format="%.2f", key=f"ua_{i}")

        st.markdown(f'<div class="section-label">Receipt Application for {receipts[i]["label"]} (.xls)</div>', unsafe_allow_html=True)
        rap_files[i] = st.file_uploader("", type=["xls"], key=f"rap_{i}", label_visibility="collapsed")

        if i < len(receipts) - 1:
            st.markdown("<hr>", unsafe_allow_html=True)

    st.markdown("---")
    col_add, col_remove = st.columns(2)
    with col_add:
        if st.button("+ ADD CHILD"):
            n = len([r for r in receipts if 'Child' in r['label']])
            st.session_state.pc_receipts.append({'label': f'Child {n+1}', 'receipt_number': '', 'customer_account': '', 'unapplied': 0.0})
            st.rerun()
    with col_remove:
        if len(receipts) > 1 and st.button("- REMOVE LAST"):
            st.session_state.pc_receipts.pop()
            st.rerun()

    st.markdown("---")
    run_pc = st.button("BUILD ALL REPUSH FILES")

    if run_pc:
        errors = []
        if not source_file: errors.append("Source file missing.")
        if not claims_file: errors.append("Claims headers file missing.")
        for i, r in enumerate(receipts):
            if not r['receipt_number'].strip(): errors.append(f"{r['label']}: receipt number required.")
            if not r['customer_account'].strip(): errors.append(f"{r['label']}: customer account required.")
            if not rap_files.get(i): errors.append(f"{r['label']}: receipt application file missing.")

        for e in errors:
            st.markdown(f'<div class="status-err">✗ {e}</div>', unsafe_allow_html=True)

        if not errors:
            with st.spinner("Processing all receipts..."):
                try:
                    clm = read_claims_headers(claims_file)
                    # read source once — filter per receipt by Item Account Number (col Y)
                    src_full = pd.read_csv(source_file, sep='|', dtype=str)
                    src_full.columns = src_full.columns.str.strip()
                    src_full['AmountApplied'] = pd.to_numeric(src_full['AmountApplied'], errors='coerce')

                    # identify parent and children
                    parent = next((r for r in receipts if 'parent' in r['label'].lower()), receipts[-1])

                    # build a sum_sn map per child (from each child's CLAIM rows)
                    child_entries = []
                    for i, r in enumerate(receipts):
                        if r is parent:
                            continue
                        acct = r['customer_account'].strip()
                        csrc = src_full[src_full['Item Account Number'].str.strip() == acct].copy()
                        child_entries.append({
                            'idx': i, 'receipt': r, 'acct': acct,
                            'rn': r['receipt_number'].strip(),
                            'sn_map': build_sum_sn(csrc),
                        })

                    results = []  # each: dict with everything needed to render + download

                    for i, r in enumerate(receipts):
                        acct = r['customer_account'].strip()
                        rn = r['receipt_number'].strip()
                        unapplied = r['unapplied']
                        rap = read_receipt_application(rap_files[i])
                        app_ref_col = 'Application Reference'
                        applied_refs = (set(rap[app_ref_col].astype(str).str.strip().tolist())
                                        if app_ref_col in rap.columns else set())

                        receipt_src = src_full[src_full['Item Account Number'].str.strip() == acct].copy()

                        if r is parent:
                            pending_claims, pending_invoices = build_parent_pending(
                                receipt_src, clm, applied_refs, rn, acct, child_entries, src_full)
                        else:
                            ce = next(c for c in child_entries if c['idx'] == i)
                            pending_claims, pending_invoices = build_child_pending(
                                receipt_src, clm, ce['sn_map'], applied_refs, rn, acct)

                        n_pending = len(pending_claims) + len(pending_invoices)

                        # SKIP RULE: only when nothing is left to push after RAP matching.
                        # unapplied = 0.00 alone does NOT skip — a freshly-created receipt
                        # (nothing applied) also reads 0.00 and must still generate.
                        if n_pending == 0:
                            results.append({
                                'skipped': True, 'label': r['label'], 'rn': rn,
                                'reason': 'All lines already applied in Oracle — nothing to reconstruct.',
                            })
                            continue

                        repush, total, status, dropped = reconcile_and_drop(
                            pending_claims, pending_invoices, unapplied)

                        file_bytes, filename = to_pipe_txt(repush, rn)
                        results.append({
                            'skipped': False, 'label': r['label'], 'rn': rn,
                            'n_claims': len(pending_claims), 'n_invoices': len(pending_invoices),
                            'total': float(total), 'unapplied': float(unapplied),
                            'status': status, 'dropped': dropped,
                            'preview': repush.head(30).to_dict('records'),
                            'preview_cols': list(repush.columns),
                            'file_bytes': file_bytes, 'filename': filename,
                        })

                    # persist so download clicks (which rerun the app) don't wipe results
                    st.session_state.pc_results = results
                    st.session_state.pc_outputs = [
                        (r['file_bytes'], r['filename']) for r in results if not r['skipped']
                    ]

                except Exception as ex:
                    st.markdown(f'<div class="status-err">✗ Error: {str(ex)}</div>', unsafe_allow_html=True)
                    raise ex

    # ── render results from session state (survives download-triggered reruns) ──
    if st.session_state.get('pc_results'):
        st.markdown("### Results")
        for res in st.session_state.pc_results:
            if res['skipped']:
                st.markdown(
                    f'<div class="status-warn">⚠ {res["label"]} ({res["rn"]}) — skipped. {res["reason"]}</div>',
                    unsafe_allow_html=True)
                continue

            st.markdown(f"#### {res['label']} — {res['rn']}")
            c1, c2, c3 = st.columns(3)
            c1.metric("Claims", res['n_claims'])
            c2.metric("Invoices", res['n_invoices'])
            c3.metric("Total", f"{res['total']:,.2f}")

            st.markdown(
                recon_status_html(res['status'], res['total'], res['unapplied'],
                                  res['dropped'], res['label']),
                unsafe_allow_html=True)

            with st.expander(f"Preview — {res['label']}"):
                st.dataframe(pd.DataFrame(res['preview'], columns=res['preview_cols']),
                             use_container_width=True, hide_index=True)

        outputs = st.session_state.get('pc_outputs', [])
        if outputs:
            st.markdown("---")
            st.markdown("### Downloads")
            if len(outputs) > 1:
                st.download_button(
                    f"↓ DOWNLOAD ALL ({len(outputs)} files, .zip)",
                    zip_outputs(outputs), "repush_files.zip", mime="application/zip",
                    key="dl_zip_all")
            for file_bytes, filename in outputs:
                st.download_button(f"↓ DOWNLOAD  {filename}", file_bytes, filename,
                                   mime="text/plain", key=f"dl_{filename}")

# ── RECEIPT CREATION MODE ─────────────────────────────────────────────────────
elif mode == "Receipt Creation":
    st.markdown('<div class="section-label">Receipt Numbers (one per line)</div>', unsafe_allow_html=True)
    receipt_numbers_text = st.text_area(
        "",
        placeholder="4024385\n2000940653\n365367591",
        height=120,
        key="rc_rns",
        label_visibility="collapsed"
    )

    st.markdown("---")
    run_rc = st.button("BUILD RECEIPT PAYLOADS")

    if run_rc:
        errors = []
        if not source_file: errors.append("Source file missing.")
        rns = [r.strip() for r in receipt_numbers_text.splitlines() if r.strip()]
        if not rns: errors.append("At least one receipt number required.")

        for e in errors:
            st.markdown(f'<div class="status-err">✗ {e}</div>', unsafe_allow_html=True)

        if not errors:
            with st.spinner("Building payloads..."):
                try:
                    src_full = pd.read_csv(source_file, sep='|', dtype=str)
                    src_full.columns = src_full.columns.str.strip()

                    results = []
                    for rn in rns:
                        payload = build_receipt_payloads(src_full, rn)
                        if payload is None:
                            results.append({'rn': rn, 'found': False})
                        else:
                            results.append({'rn': rn, 'found': True, 'payload': payload})

                    st.session_state.rc_results = results

                except Exception as ex:
                    st.markdown(f'<div class="status-err">✗ Error: {str(ex)}</div>', unsafe_allow_html=True)
                    raise ex

    # render results (persists across reruns)
    if st.session_state.get('rc_results'):
        st.markdown("### Results")
        for res in st.session_state.rc_results:
            if not res['found']:
                st.markdown(
                    f'<div class="status-err">✗ Receipt {res["rn"]} not found in source file.</div>',
                    unsafe_allow_html=True)
                continue

            p = res['payload']
            st.markdown(f"#### Receipt {res['rn']}")
            n_children = len(p['children'])
            n_r2r = len(p['r2r'])

            c1, c2, c3 = st.columns(3)
            c1.metric("Paying Account", p['parent']['CustomerAccountNumber'])
            c2.metric("Children", n_children)
            c3.metric("DS / R2R", n_r2r)

            st.markdown('**Parent Payload**')
            st.code(json.dumps(p['parent'], indent=4), language='json')

            for i, ch in enumerate(p['children'], start=1):
                st.markdown(f'**Child {i} — {ch["CustomerAccountNumber"]}**')
                st.code(json.dumps(ch, indent=4), language='json')

            if p['r2r']:
                st.markdown('**R2R Lines**')
                for line in p['r2r']:
                    st.code(
                        f"BU: {line['BU']}\n"
                        f"Receipt Number: {line['ReceiptNumber']}\n"
                        f"Child Receipt Number: {line['ChildReceiptNumber']}\n"
                        f"Child Amount: {line['ChildAmount']:.2f}\n"
                        f"Date: {line['Date']}",
                        language='text')

            st.markdown("---")
