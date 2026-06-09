import streamlit as st
import pandas as pd
from html.parser import HTMLParser
from io import StringIO

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
    df.to_csv(buf, sep='|', index=False)
    return buf.getvalue().encode('utf-8'), f"{receipt_number}_Reconstructed_{iteration}.txt"


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
mode = st.radio("", ["Standalone", "Parent / Child"], horizontal=True, label_visibility="collapsed")

st.markdown("---")

# shared inputs
st.markdown('<div class="section-label">Source File (.txt pipe-delimited)</div>', unsafe_allow_html=True)
source_file = st.file_uploader("", type=["txt"], key="source", label_visibility="collapsed")

st.markdown('<div class="section-label">Open Claims Headers Extract (.xlsx)</div>', unsafe_allow_html=True)
claims_file = st.file_uploader("", type=["xlsx"], key="claims", label_visibility="collapsed")

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
else:
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
            if r['unapplied'] == 0.0: st.info(f"{r['label']} ({r['receipt_number']}) has unapplied amount of 0.00 — will be skipped.")
            if not rap_files.get(i): errors.append(f"{r['label']}: receipt application file missing.")

        for e in errors:
            st.markdown(f'<div class="status-err">✗ {e}</div>', unsafe_allow_html=True)

        if not errors:
            with st.spinner("Processing all receipts..."):
                try:
                    clm = read_claims_headers(claims_file)
                    # read source once — filter per receipt by customer account
                    src_full = pd.read_csv(source_file, sep='|', dtype=str)
                    src_full.columns = src_full.columns.str.strip()
                    src_full['AmountApplied'] = pd.to_numeric(src_full['AmountApplied'], errors='coerce')

                    # identify children and parent
                    children = [r for r in receipts if 'parent' not in r['label'].lower()]
                    parent = next((r for r in receipts if 'parent' in r['label'].lower()), receipts[-1])
                    parent_idx = receipts.index(parent)

                    # build child SNInvoiceNumber SUMIF lookup per child
                    # for each child, get unique SNInvoiceNumber -> summed AmountApplied
                    child_sn_sumifs = []  # list of dicts: {sn_number: summed_amount}
                    for i, r in enumerate(receipts):
                        if r == parent: continue
                        acct = r['customer_account'].strip()
                        child_src = src_full[
                            src_full['Item Account Number'].str.strip() == acct
                        ].copy()
                        # invoice lines: non-blank SNInvoiceNumber
                        inv_mask = child_src['SNInvoiceNumber'].notna() & (child_src['SNInvoiceNumber'].str.strip() != '')
                        inv_lines = child_src[inv_mask].copy()
                        # sumif per unique SNInvoiceNumber
                        sn_sumif = inv_lines.groupby('SNInvoiceNumber')['AmountApplied'].sum().to_dict()
                        child_sn_sumifs.append({'receipt': r, 'idx': i, 'sn_sumif': sn_sumif})

                    st.markdown("### Results")
                    output_files = []

                    # process each receipt
                    for i, r in enumerate(receipts):
                        if r['unapplied'] == 0.0:
                            st.markdown(f'<div class="status-warn">⚠ {r["label"]} ({r["receipt_number"]}) — unapplied amount is 0.00, skipping.</div>', unsafe_allow_html=True)
                            continue
                        acct = r['customer_account'].strip()
                        rn = r['receipt_number'].strip()
                        unapplied = r['unapplied']
                        rap = read_receipt_application(rap_files[i])
                        app_ref_col = 'Application Reference'
                        applied_refs = set(rap[app_ref_col].astype(str).str.strip().tolist()) if app_ref_col in rap.columns else set()

                        # filter source by customer account
                        receipt_src = src_full[
                            src_full['Item Account Number'].str.strip() == acct
                        ].copy()

                        # 1. pending claims
                        pending_claims = get_pending_claims(receipt_src, clm)

                        # 2. invoice lines
                        if r != parent:
                            # CHILD: filter non-blank SNInvoiceNumber, VLOOKUP vs receipt app
                            inv_mask = receipt_src['SNInvoiceNumber'].notna() & (receipt_src['SNInvoiceNumber'].str.strip() != '')
                            inv_lines = receipt_src[inv_mask].copy()
                            # get sumif for this child
                            child_entry = next(c for c in child_sn_sumifs if c['idx'] == i)
                            sn_sumif = child_entry['sn_sumif']
                            # unique SN numbers
                            unique_sns = list(sn_sumif.keys())
                            # for each row, use sumif amount (negated) as AmountApplied
                            inv_lines['AmountApplied'] = inv_lines['SNInvoiceNumber'].map(
                                lambda sn: -sn_sumif.get(sn, 0)
                            )
                            # deduplicate to unique SN rows
                            inv_lines = inv_lines.drop_duplicates(subset=['SNInvoiceNumber']).copy()
                            # vlookup vs child receipt app
                            inv_lines['VLOOKUP'] = inv_lines['TransactionNumber'].apply(
                                lambda t: t if str(t).strip() in applied_refs else None
                            )
                            pending_invoices = inv_lines[inv_lines['VLOOKUP'].isna()].copy()
                            src_cols = [c for c in receipt_src.columns]
                            pending_invoices = pending_invoices[[c for c in src_cols if c in pending_invoices.columns]].copy()
                            # fix receipt number and customer account
                            pending_invoices['ReceiptNumber'] = rn
                            pending_invoices['CustomerAccountNumber'] = acct
                            pending_claims['ReceiptNumber'] = rn
                            pending_claims['CustomerAccountNumber'] = acct

                        else:
                            # PARENT: two sets of invoice lines
                            # A) residual lines (have SNInvoiceNumber, shared with children)
                            inv_mask_sn = receipt_src['SNInvoiceNumber'].notna() & (receipt_src['SNInvoiceNumber'].str.strip() != '')
                            inv_lines_sn = receipt_src[inv_mask_sn].copy()

                            # calculate residual: original amount minus all children's sumif amounts for same SN
                            def get_residual(row):
                                sn = row['SNInvoiceNumber']
                                original = row['AmountApplied']
                                child_total = sum(c['sn_sumif'].get(sn, 0) for c in child_sn_sumifs)
                                return original + child_total  # original is negative, child_total is negative

                            inv_lines_sn['AmountApplied'] = inv_lines_sn.apply(get_residual, axis=1)
                            inv_lines_sn = inv_lines_sn.drop_duplicates(subset=['SNInvoiceNumber']).copy()
                            inv_lines_sn['VLOOKUP'] = inv_lines_sn['TransactionNumber'].apply(
                                lambda t: t if str(t).strip() in applied_refs else None
                            )
                            pending_sn = inv_lines_sn[inv_lines_sn['VLOOKUP'].isna()].copy()

                            # B) direct parent invoice lines (non-blank TransactionNumber, no SNInvoiceNumber filter)
                            inv_mask_txn = receipt_src['TransactionNumber'].notna() & (receipt_src['TransactionNumber'].str.strip() != '')
                            inv_lines_txn = receipt_src[inv_mask_txn].copy()
                            inv_lines_txn['VLOOKUP'] = inv_lines_txn['TransactionNumber'].apply(
                                lambda t: t if str(t).strip() in applied_refs else None
                            )
                            pending_txn = inv_lines_txn[inv_lines_txn['VLOOKUP'].isna()].copy()

                            src_cols = [c for c in receipt_src.columns]
                            pending_sn = pending_sn[[c for c in src_cols if c in pending_sn.columns]].copy()
                            pending_txn = pending_txn[[c for c in src_cols if c in pending_txn.columns]].copy()
                            pending_invoices = pd.concat([pending_sn, pending_txn], ignore_index=True)

                        repush, total, status, dropped = reconcile_and_drop(pending_claims, pending_invoices, unapplied)

                        st.markdown(f"#### {r['label']} — {rn}")
                        c1, c2, c3 = st.columns(3)
                        c1.metric("Claims", len(pending_claims))
                        c2.metric("Invoices", len(pending_invoices))
                        c3.metric("Total", f"{total:,.2f}")

                        st.markdown(recon_status_html(status, total, unapplied, dropped, r['label']), unsafe_allow_html=True)

                        with st.expander(f"Preview — {r['label']}"):
                            st.dataframe(repush.head(30), use_container_width=True, hide_index=True)

                        file_bytes, filename = to_pipe_txt(repush, rn)
                        output_files.append((file_bytes, filename))

                    st.markdown("---")
                    st.markdown("### Downloads")
                    for file_bytes, filename in output_files:
                        st.download_button(f"↓ DOWNLOAD  {filename}", file_bytes, filename, mime="text/plain", key=f"dl_{filename}")

                except Exception as ex:
                    st.markdown(f'<div class="status-err">✗ Error: {str(ex)}</div>', unsafe_allow_html=True)
                    raise ex
