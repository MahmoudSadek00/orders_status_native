"""
Logic for the standalone "Orders Status Check (Native Export)" Streamlit app (Aug 2026,
per Mahmoud -- kept as its own separate app/link, deliberately not a page inside the
consolidation tool or the other Orders Status Check app).

For Shopify's own native "Export orders" file (Orders page -> Export in Shopify itself
-- NOT an analytics report like Monthly POS Report / Sales overview, which the other two
tools use). Confirmed against a real sample file, Aug 2026:

1. One row per LINE ITEM, not one row per order. Order-level columns (Financial Status,
   Subtotal, Shipping, Taxes, Total, Created at, Cancelled at, Shipping address, ...) are
   populated ONLY on the FIRST row of each order and blank on every later line-item row
   for that same order (verified: non-blank 'Financial Status' row count matched the
   distinct order count exactly). So filtering to "order_value (Subtotal) column is
   non-blank" both finds the first row per order AND needs no groupby/sum across rows --
   unlike an analytics-report export, where an order's rows need summing.

2. Order price WITHOUT shipping = the 'Subtotal' column directly, no derivation needed.
   Verified: Total = Subtotal + Shipping when shipping was actually charged, and
   Total = Subtotal alone when a discount waived shipping (Shipping still shows the
   nominal/list rate either way) -- so Subtotal itself is unaffected and always means
   goods value, shipping excluded, already net of any per-order discount.

3. Cancelled orders: 'Cancelled at' (a real timestamp) is the complete signal -- non-
   blank for every genuinely cancelled order. Financial Status ('voided') alone missed
   21 of 169 real cancelled orders in the sample file (some cancelled orders carry some
   OTHER Financial Status, mostly 'pending').

4. 'Created at' already comes as an unambiguous ISO datetime with a timezone offset
   ('2026-03-10 12:29:49 +0400') -- no day-first/month-first guessing needed.

5. Country comes as a real ISO code ('AE' / 'OM'), normalized here to the raw sheets'
   own convention ('UAE' / 'OM'). Blank for POS/Draft orders (~57% in the sample --
   expected, not a data problem, same pattern seen in every other Shopify export format
   used across this whole project).

Why this app exists at all: the 3 raw order-tracking Google Sheets only ever get a row
once an order is physically WITH the shipping company -- the ops team's own workflow,
unchanged by any of this. That means Total Orders / Cancelled Orders on the Dashboard
were undercounted: a cancelled order, or one still awaiting shipment, never showed up
anywhere. This app reads the 3 raw sheets AND the staging "Orders" tab that already
feeds the Dashboard (live, read-only for the raw sheets) and returns only the orders
from an uploaded file that aren't logged ANYWHERE yet -- safe to append at the bottom of
the staging sheet without ever duplicating a row, on demand, whenever Mahmoud wants
fresher numbers.
"""
import io
import re
import time
import unicodedata
import datetime as dt

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly',
]

# The 3 raw order-tracking sheets (same ones the orders-sync project's clean.py reads) --
# duplicated here since this app deploys completely on its own, separate from that
# project and from the consolidation tool. Only the Reference Number column is needed
# here, just to know which orders are already logged.
RAW_SHEETS = {
    'uae_om': {'label': 'UAE & Oman', 'sheet_id': '17jK3xeRGuYrMrEVfkWbuwQEgO7LbcvpOq6KYPLQTtzU', 'tab': '2026', 'ref_col': 'Reference Number'},
    'gulf': {'label': 'Gulf (SA/QA/KW)', 'sheet_id': '1lTvEcsSJaOCSdEo00IxFxczqdUwc_oFo_xqzu5NfhAs', 'tab': '2026 orders', 'ref_col': 'Reference Number'},
    'iraq': {'label': 'Iraq', 'sheet_id': '1kLJvMn0rNvippmmdPSl-_SOdAlxfTD0FV17jO9rnZBc', 'tab': '2026 orders', 'ref_col': 'ReceiptNumber'},
}

ORDERS_TAB = 'Orders'

# Mahmoud's real staging sheet (confirmed, Aug 2026: docs.google.com/spreadsheets/d/
# 1dZMqtqvnxe6GspH0C10AvXECB74NP-ZjDG_BihMOkmg -- "Orders" tab). Baked in as a default
# so the only secret still needed is the service-account credential -- app.py's
# staging_spreadsheet_id secret, if set, overrides this (e.g. if the staging sheet ever
# moves).
STAGING_SPREADSHEET_ID_DEFAULT = '1dZMqtqvnxe6GspH0C10AvXECB74NP-ZjDG_BihMOkmg'

# default_country: fallback when Shipping Country is blank (POS/Draft orders -- no
# shipping address at all, not a data problem, same as every other Shopify format).
ORDER_STATUS_GROUPS = {
    'uae_om': {'default_country': 'UAE'},
    'gulf': {'default_country': 'SA'},
    'iraq': {'default_country': 'IQ'},
}

# Column-name guesses for this format.
NATIVE_EXPORT_DEFAULTS = {
    'ref_number': ['Name'],
    'order_date': ['Created at'],
    'order_value': ['Subtotal'],
    'country': ['Shipping Country'],
    'city': ['Shipping City'],
    'cancelled_at': ['Cancelled at'],
}

# Shopify's own 'Shipping country' field comes as a bare ISO code in this format ('AE'),
# unlike the analytics-report formats' full names ('United Arab Emirates') -- both
# handled here in case a similar file ever uses the long form.
SHOPIFY_COUNTRY_MAP = {
    'united arab emirates': 'UAE', 'uae': 'UAE', 'ae': 'UAE',
    'oman': 'OM', 'om': 'OM',
    'saudi arabia': 'SA', 'ksa': 'SA', 'sa': 'SA',
    'kuwait': 'KW', 'kw': 'KW',
    'qatar': 'QA', 'qa': 'QA',
    'iraq': 'IQ', 'iq': 'IQ',
}

ORDERS_HEADER = [
    'Source', 'Reference Number', 'Shipping Date', 'Order Date', 'Country', 'City',
    'Order Value', 'Status', 'New Customer Orders', 'Returning Customer Orders',
    'Needs Review', 'Review Reason', 'Last Synced At (UTC)',
]


# ---------------------------------------------------------------------------
# Small shared helpers (mirrors of the consolidation tool's engine.py / orders_status.py
# -- duplicated on purpose, this app has zero dependency on those separate projects).
# ---------------------------------------------------------------------------

def clean_display(s):
    """Strip invisible/combining Unicode characters and extra whitespace but keep
    original case/formatting -- used for values actually written into the output."""
    if s is None:
        return ''
    s = str(s)
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(ch for ch in s if not unicodedata.combining(ch) and (ch.isprintable() or ch == ' '))
    return re.sub(r'\s+', ' ', s).strip()


def clean_key(s):
    """Hidden-character-safe key for MATCHING two reference numbers -- uppercased, with
    a leading '#' stripped, so '#90127' / '90127' and similar variants all match."""
    d = clean_display(s).upper()
    if d.startswith('#'):
        d = d[1:]
    return d


def fix_mojibake(s):
    """Best-effort repair for Arabic text that was UTF-8 originally but got decoded as
    cp1252 somewhere upstream of the uploaded file (confirmed against a real sample --
    'Shipping City' values like 'Ø§Ù„Ø¹ÙŠÙ†' round-trip cleanly back to 'العين' via
    cp1252-encode/utf-8-decode). Only touches the string if the round-trip actually
    succeeds AND changes something -- an already-correct value (ASCII or already-proper
    Arabic) is returned untouched. The minority of values that don't cleanly round-trip
    are left as-is rather than guessed further."""
    if not s:
        return s
    try:
        fixed = s.encode('cp1252').decode('utf-8')
        return fixed if fixed != s else s
    except (UnicodeEncodeError, UnicodeDecodeError):
        return s


def read_any(file) -> pd.DataFrame:
    name = getattr(file, 'name', str(file)).lower()
    if name.endswith('.csv'):
        return pd.read_csv(file, dtype=str, keep_default_na=False)
    return pd.read_excel(file, dtype=str, keep_default_na=False)


def guess_column(columns, candidates):
    lower_map = {str(c).lower().strip(): c for c in columns}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


def default_mapping(columns, defaults):
    return {field: guess_column(columns, cands) for field, cands in defaults.items()}


def normalize_order_country(raw, default_country):
    """Returns (code, was_blank, was_unrecognized)."""
    s = clean_display(raw)
    if not s:
        return default_country, True, False
    key = s.lower()
    mapped = SHOPIFY_COUNTRY_MAP.get(key)
    if mapped:
        return mapped, False, False
    upper = s.upper()
    if upper in SHOPIFY_COUNTRY_MAP.values():
        return upper, False, False
    return default_country, False, True


def get_client(service_account_info):
    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
    return gspread.authorize(creds)


def _call_with_retry(fn, max_retries=4):
    """Runs fn() with short exponential backoff on transient Google API errors (429 rate
    limit, or a 5xx like the 503 'service currently unavailable' Mahmoud hit on a large
    real file, Aug 2026) -- a real, if uncommon, Sheets API hiccup, not a config problem.
    Anything else (403 permission denied, 404 not found, ...) is NOT transient and is
    raised immediately -- retrying it would just waste time on an error retrying can't
    fix."""
    last_err = None
    for attempt in range(max_retries):
        try:
            return fn()
        except gspread.exceptions.APIError as e:
            status = None
            resp = getattr(e, 'response', None)
            if resp is not None:
                status = getattr(resp, 'status_code', None)
            if status not in (429, 500, 502, 503, 504):
                raise
            last_err = e
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)  # 1s, 2s, 4s, 8s
    raise last_err


def _col_keys(ws, ref_col_name):
    """clean_key() of every non-blank value under ref_col_name in worksheet ws.

    Reads only the header row plus that ONE column -- not the whole sheet -- on purpose:
    an earlier version called ws.get_values() (the full sheet, every column) just to
    pull one column out of it, which on a large real sheet (Mahmoud's real merged file
    triggered this, Aug 2026 -- hundreds of thousands of source rows) is enough data to
    occasionally trip a 503 from Google's own API. Fetching only what's actually needed
    is both faster and far less likely to hit that limit; _call_with_retry above still
    covers the cases where it happens anyway."""
    header = _call_with_retry(lambda: ws.row_values(1))
    if not header or ref_col_name not in header:
        return set()
    idx = header.index(ref_col_name)  # 0-based position in the header row
    col_values = _call_with_retry(lambda: ws.col_values(idx + 1))  # gspread columns are 1-indexed
    out = set()
    for v in col_values[1:]:  # skip the header cell itself
        if v:
            k = clean_key(v)
            if k:
                out.add(k)
    return out


def read_existing_keys(gc, staging_spreadsheet_id):
    """Reads the 3 raw sheets AND the staging Orders tab live. Returns
    (keys, counts) -- keys: the union of every already-logged order's clean_key()
    across all 4 places; counts: per-source row counts for the on-screen caption."""
    keys = set()
    counts = {}
    for gkey, cfg in RAW_SHEETS.items():
        sh = _call_with_retry(lambda cfg=cfg: gc.open_by_key(cfg['sheet_id']))
        ws = _call_with_retry(lambda sh=sh, cfg=cfg: sh.worksheet(cfg['tab']))
        k = _col_keys(ws, cfg['ref_col'])
        counts[f'raw:{gkey}'] = len(k)
        keys |= k

    staging_sh = _call_with_retry(lambda: gc.open_by_key(staging_spreadsheet_id))
    try:
        orders_ws = _call_with_retry(lambda: staging_sh.worksheet(ORDERS_TAB))
    except gspread.WorksheetNotFound:
        raise RuntimeError(
            f"The staging sheet has no '{ORDERS_TAB}' tab yet -- run the main daily sync "
            f"(sync_orders.py, in the orders-sync project) at least once first, or create "
            f"the tab by hand with the right header row."
        )
    k = _col_keys(orders_ws, 'Reference Number')
    counts['staging'] = len(k)
    keys |= k
    return keys, counts


def classify_native_export_orders(shopify_df, target_key, mapping):
    """mapping: dict with keys ref_number / order_date / order_value / country / city /
    cancelled_at -- see NATIVE_EXPORT_DEFAULTS. order_value ('Subtotal') is populated
    only on each order's first row and blank on later line-item rows for the same order
    -- filtering to non-blank order_value both selects the first row per order and needs
    no sum across rows (see module docstring). order_value itself already excludes
    shipping, no derivation needed. cancelled_at is checked for non-blank, not a
    TRUE/1/YES flag (it's a real timestamp column, e.g. 'Cancelled at').

    Returns (rows, stats)."""
    ref_col = mapping.get('ref_number')
    date_col = mapping.get('order_date')
    value_col = mapping.get('order_value')
    country_col = mapping.get('country')
    city_col = mapping.get('city')
    cancelled_col = mapping.get('cancelled_at')

    work = shopify_df.copy()
    work = work[work[ref_col].astype(str).str.strip() != '']
    if value_col:
        # The one-row-per-order filter -- see docstring. Without an order_value mapping
        # there's no reliable way to isolate one row per order in this format at all.
        work = work[work[value_col].astype(str).str.strip() != '']
    work['_key'] = work[ref_col].map(clean_key)
    work = work[work['_key'] != '']

    default_country = ORDER_STATUS_GROUPS[target_key]['default_country']
    source_label = f"Shopify (unshipped) - {RAW_SHEETS[target_key]['label']}"
    rows = []
    blank_country_count = 0
    for _, r in work.iterrows():
        order_date = None
        if date_col:
            raw_date = r.get(date_col)
            if raw_date not in (None, ''):
                try:
                    order_date = pd.to_datetime(raw_date).date()
                except Exception:
                    order_date = None

        value = None
        if value_col:
            raw_val = r.get(value_col)
            try:
                value = float(str(raw_val).replace(',', '').strip())
            except (TypeError, ValueError):
                value = None

        country_raw = r.get(country_col, '') if country_col else ''
        country, was_blank, was_unrec = normalize_order_country(country_raw, default_country)
        if was_blank:
            blank_country_count += 1

        is_cancelled = bool(str(r.get(cancelled_col, '')).strip()) if cancelled_col else False

        rows.append({
            'key': r['_key'],
            'source_label': source_label,
            'ref_number': clean_display(r.get(ref_col, '')),
            'shipping_date': None,
            'order_date': order_date,
            'country': country,
            # fix_mojibake MUST run BEFORE clean_display: clean_display's NFKD-then-
            # strip-combining-marks step (needed elsewhere for real accented text)
            # treats several of the pseudo-Latin characters mojibake produces as
            # decomposable diacritics and deletes them, destroying the exact bytes the
            # cp1252 repair needs. Repairing first, then clean_display on the result
            # (proper Arabic by then, or the untouched original if repair failed) avoids
            # that entirely.
            'city': clean_display(fix_mojibake(r.get(city_col, ''))) if city_col else '',
            'order_value': value,
            'status': 'Cancelled' if is_cancelled else 'Pending',
            'new_customer': 0,
            'returning_customer': 0,
            'needs_review': was_unrec,
            'review_reason': (
                f"unrecognized Shipping country: {country_raw!r} -- defaulted to '{default_country}'"
                if was_unrec else ''
            ),
        })

    stats = {'orders_total': len(rows), 'blank_country_count': blank_country_count}
    return rows, stats


def filter_new(rows, existing_keys):
    """Only orders not already logged anywhere (see read_existing_keys) -- safe to
    append at the bottom of the staging sheet without ever duplicating a row."""
    return [r for r in rows if r['key'] and r['key'] not in existing_keys]


def fmt_date(d):
    if d is None:
        return ''
    if isinstance(d, dt.datetime):
        d = d.date()
    return d.isoformat()


def row_to_values(row, now_str):
    return [
        row['source_label'], row['ref_number'], fmt_date(row['shipping_date']),
        fmt_date(row['order_date']), row['country'], row['city'],
        (row['order_value'] if row['order_value'] is not None else ''),
        row['status'], row['new_customer'], row['returning_customer'],
        ('Yes' if row['needs_review'] else ''), row['review_reason'],
        now_str,
    ]


def rows_to_dataframe(rows):
    now_str = dt.datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
    data = [row_to_values(r, now_str) for r in rows]
    return pd.DataFrame(data, columns=ORDERS_HEADER)


def rows_to_excel_bytes(rows):
    df = rows_to_dataframe(rows)
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def append_to_staging(gc, staging_spreadsheet_id, rows):
    """Appends the given rows at the bottom of the live staging Orders tab. Returns the
    number of rows appended. Same retry wrapping as read_existing_keys/_col_keys, so a
    transient 429/5xx here (e.g. right after a big read) doesn't fail the whole add."""
    if not rows:
        return 0
    staging_sh = _call_with_retry(lambda: gc.open_by_key(staging_spreadsheet_id))
    orders_ws = _call_with_retry(lambda: staging_sh.worksheet(ORDERS_TAB))
    now_str = dt.datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
    values = [row_to_values(r, now_str) for r in rows]
    _call_with_retry(lambda: orders_ws.append_rows(values))
    return len(values)
