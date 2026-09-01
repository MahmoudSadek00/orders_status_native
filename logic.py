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

3. Cancelled orders: 'Cancelled at' (a real timestamp) being non-blank is the primary
   signal -- Financial Status ('voided') alone missed 21 of 169 real cancelled orders in
   the sample file (some cancelled orders carry some OTHER Financial Status, mostly
   'pending'). PLUS (Aug 2026, per Mahmoud, real production case: order #91780):
   Financial Status 'refunded' while Fulfillment Status ISN'T 'fulfilled' also counts as
   Cancelled -- money back before the order ever shipped, a real cancellation even
   though nobody clicked Shopify's own "Cancel order" button. A 'refunded' order that
   WAS fulfilled is left alone (shipped, then returned -- not a cancellation). See
   REFUND_CANCEL_STATUSES.

4. 'Created at' already comes as an unambiguous ISO datetime with a timezone offset
   ('2026-03-10 12:29:49 +0400') -- no day-first/month-first guessing needed.

5. Country comes as a real ISO code ('AE' / 'OM'), normalized here to the raw sheets'
   own convention ('UAE' / 'OM'). Blank for POS orders (~57% of orders in the sample --
   in-store, rung up and paid on the spot). POS orders are EXCLUDED outright via the
   'Source' column (Aug 2026, per Mahmoud) -- they never touch the shipping company, so
   they don't belong in an "orders still waiting on the shipping company" tracker at
   all. Confirmed against the real sample: every blank-Shipping-Country row was a POS
   row and vice versa, 1:1.

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
import zipfile
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

# Mahmoud's own new holding tab (Aug 2026) in the SAME staging spreadsheet as ORDERS_TAB
# -- he created it by hand. New Cancelled/Pending orders from this app now get appended
# here instead of straight into ORDERS_TAB, so he can review them separately from what
# the daily sync / other tools write. Still checked for de-duplication just like
# ORDERS_TAB (see read_existing_keys) and still cross-checked against the 3 raw sheets
# for staleness (see find_stale_not_shipped) -- same header row as ORDERS_HEADER below.
NOT_SHIPPED_TAB = 'Not Shipped'

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
    'source': ['Source'],
    'financial_status': ['Financial Status'],
    'fulfillment_status': ['Fulfillment Status'],
    'tags': ['Tags'],
}

# A manually-applied Tag containing "cancel" in any form/case (e.g. "Cancelled",
# "CANCEL - customer request", "cancel_by_agent") also counts as a real cancellation
# (Aug 2026, per Mahmoud) -- agents tag an order this way sometimes without ever using
# Shopify's own Cancel-order button (which is what sets Cancelled at) and without it
# necessarily being refunded yet either (which is what REFUND_CANCEL_STATUSES above
# catches). A simple substring match, not an exact-value match, since the tag text
# itself isn't standardized. ALSO tolerates typos in the word itself (confirmed real
# case, Aug 2026: a tag literally spelled "CANACEL") via a small edit-distance check on
# each individual word of the tag -- a plain substring check can never catch a misspelled
# word, since the exact letters "cancel" never appear anywhere in "canacel".
CANCEL_TAG_SUBSTRING = 'cancel'
CANCEL_TAG_MAX_EDIT_DISTANCE = 2
_CANCEL_TAG_WORD_LEN_RANGE = (len(CANCEL_TAG_SUBSTRING) - 3, len(CANCEL_TAG_SUBSTRING) + 3)


def _levenshtein(a: str, b: str) -> int:
    """Plain edit distance (insertions/deletions/substitutions), no external deps --
    small inputs only (single words), so the simple O(len(a)*len(b)) table is fine."""
    if a == b:
        return 0
    la, lb = len(a), len(b)
    if la == 0:
        return lb
    if lb == 0:
        return la
    prev = list(range(lb + 1))
    for i, ca in enumerate(a, start=1):
        cur = [i] + [0] * lb
        for j, cb in enumerate(b, start=1):
            cost = 0 if ca == cb else 1
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost)
        prev = cur
    return prev[lb]


def _tag_implies_cancelled(tags_val: str) -> bool:
    """tags_val must already be lowercased/cleaned. Fast path: plain substring (catches
    "cancelled", "cancel_by_agent", etc.). Fallback: each individual word (split on the
    usual tag separators) within a small edit distance of "cancel" -- catches a
    misspelled tag like "CANACEL" without over-matching unrelated short words (the word
    length is pre-filtered to stay close to "cancel"'s own length first)."""
    if not tags_val:
        return False
    if CANCEL_TAG_SUBSTRING in tags_val:
        return True
    lo, hi = _CANCEL_TAG_WORD_LEN_RANGE
    for word in re.split(r'[,;/|\s]+', tags_val):
        if lo <= len(word) <= hi and _levenshtein(word, CANCEL_TAG_SUBSTRING) <= CANCEL_TAG_MAX_EDIT_DISTANCE:
            return True
    return False

# A 'refunded' order whose Fulfillment Status ISN'T 'fulfilled' means the money went
# back to the customer without the order ever shipping -- functionally a cancellation,
# even though nobody clicked Shopify's own "Cancel order" button (which is what sets
# Cancelled at). Confirmed real case, Aug 2026 (order #91780, per Mahmoud): refunded +
# unfulfilled, Cancelled at blank -- came out Pending, which was wrong. A 'refunded'
# order that WAS fulfilled is a different situation entirely (shipped, then returned
# after the fact) and is left alone -- not a cancellation, and likely already logged
# elsewhere as shipped. 'partially_refunded' gets the exact same treatment (per
# Mahmoud, Aug 2026) -- same reasoning, same fulfillment_status guard: a
# partially-refunded order that's still unfulfilled is treated as Cancelled too, but
# one that's already fulfilled (shipped, then partially adjusted/returned) is left
# alone. Note this is naturally safe either way for an order that HAS actually
# shipped: filter_new drops it from the "new orders" list once it's found on a raw
# sheet, regardless of what status got assigned here.
REFUND_CANCEL_STATUSES = {'refunded', 'partially_refunded'}

# Shopify's own 'Source' column (own channel field, NOT this tool's own 'Source' output
# column) -- values seen in a real sample: 'pos' (in-store, sold and paid on the spot --
# never touches the shipping company, so it should never enter this "orders waiting on
# the shipping company" tracker at all), 'web', 'shopify_draft_order', and numeric
# third-party-app channel IDs. Confirmed against the real sample (Aug 2026, per Mahmoud):
# every single blank-Shipping-Country row was a 'pos' row and vice versa -- so excluding
# 'pos' here also naturally clears out the "blank country defaulted" noise from the
# count, not just double-counted in-store sales.
EXCLUDED_SOURCES = {'pos'}

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

# Countries GC doesn't actually serve at all, in any group (Aug 2026, per Mahmoud --
# Bahrain isn't a market here). Before this, a Bahrain shipping address just fell through
# to "unrecognized" and got silently defaulted into the group's default country (e.g.
# counted as Saudi under the Gulf group) with a needs-review flag -- technically visible,
# but still added to the sheet as if it belonged there. Skipped outright now instead of
# defaulted+flagged.
EXCLUDED_COUNTRY_NAMES = {'bahrain', 'bh'}

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


SUPPORTED_EXTS = ('.csv', '.xlsx', '.xls')


def _read_bytes(name: str, data: bytes, usecols=None, nrows=None) -> pd.DataFrame:
    # usecols as a callable (not a fixed list) never raises if a column is absent from a
    # particular part -- it just keeps whichever of the wanted names are actually there,
    # which matters across 70+ files that may not all share the exact same column set.
    col_filter = (lambda c: c in usecols) if usecols else None
    buf = io.BytesIO(data)
    if name.lower().endswith('.csv'):
        return pd.read_csv(buf, dtype=str, keep_default_na=False, usecols=col_filter, nrows=nrows)
    return pd.read_excel(buf, dtype=str, keep_default_na=False, usecols=col_filter, nrows=nrows)


def _iter_data_entries(files):
    """Yields (source_label, name, data_bytes) for every csv/xlsx/xls found across the
    given uploads -- diving into any .zip (at any folder depth inside it) -- without
    holding more than one entry's raw bytes at a time. Shared by read_headers/read_many
    so both walk the exact same set of files/entries."""
    for f in files:
        name = getattr(f, 'name', str(f))
        data = f.getvalue() if hasattr(f, 'getvalue') else f.read()
        lower = name.lower()
        if lower.endswith('.zip'):
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as zf:
                    for info in zf.infolist():
                        entry_name = info.filename
                        base = entry_name.rsplit('/', 1)[-1]
                        if info.is_dir() or not base or base.startswith('.') or '__MACOSX' in entry_name:
                            continue
                        if not base.lower().endswith(SUPPORTED_EXTS):
                            yield ('skip', f"{name} -> {entry_name}", "not a csv/xlsx/xls file")
                            continue
                        try:
                            with zf.open(info) as zf_entry:
                                entry_bytes = zf_entry.read()
                        except Exception as e:
                            yield ('skip', f"{name} -> {entry_name}", str(e))
                            continue
                        yield ('data', f"{name} -> {entry_name}", (base, entry_bytes))
            except Exception as e:
                yield ('skip', name, f"couldn't open as a zip: {e}")
        elif lower.endswith(SUPPORTED_EXTS):
            yield ('data', name, (name, data))
        else:
            yield ('skip', name, "unsupported file type")
        del data


def read_headers(files):
    """Cheaply peeks at just the column headers of the FIRST readable file/zip-entry
    among the uploads (Aug 2026, per Mahmoud -- with 70+ files/folders in one go, fully
    parsing every one of them just to populate the column-mapping dropdowns was what ran
    the app out of memory). Every part of the same "Export orders" run shares one schema,
    so one file's headers are enough to build the mapping UI; the full, multi-file read
    only happens afterwards in read_many, and only for the columns actually mapped."""
    for kind, label, payload in _iter_data_entries(files):
        if kind != 'data':
            continue
        base, data = payload
        try:
            df0 = _read_bytes(base, data, nrows=0)
            return df0.columns.tolist()
        except Exception:
            continue
    return []


def read_many(files, usecols=None, on_progress=None):
    """Reads any mix of uploaded csv/xlsx/xls files AND .zip files in one go (Aug 2026,
    per Mahmoud -- Shopify splits a big "Export orders" run into several
    orders_export_N.zip download links, all in ONE email, and across the Gulf country
    groups these pile up into dozens of downloaded zips/folders). Each zip can itself
    contain any number of csv/xlsx/xls files nested at any folder depth -- every part
    just gets dragged onto the uploader together, zipped or already-extracted, no manual
    unzip/re-foldering needed first.

    usecols, if given, restricts every file/entry to just those column names (see
    read_headers above) -- across 70+ files this is what keeps memory in check, since a
    native Shopify export can carry dozens of columns per row but this app only ever
    needs ~9 of them. on_progress, if given, is called after each entry as
    on_progress(done_count, total_count, label) -- purely cosmetic (e.g. a progress bar),
    no Streamlit dependency here.

    Returns (combined_df, stats) where stats has 'files_read' (source file names actually
    parsed, zip entries shown as "zipname -> entry path") and 'files_skipped'
    ((name, reason) pairs for anything that wasn't a recognized data file, e.g. a stray
    __MACOSX entry or a non-data attachment that slipped into the selection)."""
    import gc

    entries = list(_iter_data_entries(files))
    total = len(entries)
    frames = []
    files_read = []
    files_skipped = []
    for i, (kind, label, payload) in enumerate(entries, start=1):
        if kind == 'skip':
            files_skipped.append((label, payload))
        else:
            base, data = payload
            try:
                frames.append(_read_bytes(base, data, usecols=usecols))
                files_read.append(label)
            except Exception as e:
                files_skipped.append((label, str(e)))
        if on_progress:
            on_progress(i, total, label)
        if i % 10 == 0:
            gc.collect()

    stats = {'files_read': files_read, 'files_skipped': files_skipped}
    if not frames:
        return pd.DataFrame(), stats
    combined = pd.concat(frames, ignore_index=True, sort=False).fillna('')
    del frames
    gc.collect()
    return combined, stats


def read_any(file) -> pd.DataFrame:
    name = getattr(file, 'name', str(file))
    data = file.getvalue() if hasattr(file, 'getvalue') else file.read()
    return _read_bytes(name, data)


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
    """Reads the 3 raw sheets, the staging Orders tab, AND the staging Not Shipped tab,
    all live. Returns (keys, raw_keys, counts):
      - keys: the union of every already-logged order's clean_key() across all 5 places
        -- used to decide which uploaded orders are actually new (filter_new).
      - raw_keys: the union across ONLY the 3 raw sheets (not the 2 staging tabs) --
        used separately by find_stale_not_shipped to catch an order that was added to
        Not Shipped earlier and has SINCE shown up on a raw sheet for real.
      - counts: per-source row counts for the on-screen caption."""
    keys = set()
    raw_keys = set()
    counts = {}
    for gkey, cfg in RAW_SHEETS.items():
        sh = _call_with_retry(lambda cfg=cfg: gc.open_by_key(cfg['sheet_id']))
        ws = _call_with_retry(lambda sh=sh, cfg=cfg: sh.worksheet(cfg['tab']))
        k = _col_keys(ws, cfg['ref_col'])
        counts[f'raw:{gkey}'] = len(k)
        keys |= k
        raw_keys |= k

    staging_sh = _call_with_retry(lambda: gc.open_by_key(staging_spreadsheet_id))

    try:
        orders_ws = _call_with_retry(lambda: staging_sh.worksheet(ORDERS_TAB))
        k = _col_keys(orders_ws, 'Reference Number')
        counts['staging_orders'] = len(k)
        keys |= k
    except gspread.WorksheetNotFound:
        # Not fatal on its own -- Not Shipped is the tab this app actually writes to now
        # (see append_to_staging) -- but the Orders tab is still expected to exist since
        # the daily sync / other Orders Status Check app write into it.
        counts['staging_orders'] = 0

    try:
        not_shipped_ws = _call_with_retry(lambda: staging_sh.worksheet(NOT_SHIPPED_TAB))
        k = _col_keys(not_shipped_ws, 'Reference Number')
        counts['not_shipped'] = len(k)
        keys |= k
    except gspread.WorksheetNotFound:
        raise RuntimeError(
            f"The staging sheet has no '{NOT_SHIPPED_TAB}' tab yet -- create it by hand "
            f"first (same header row as the Orders tab works fine, or leave it "
            f"completely empty and this app will add the header itself on the first add)."
        )
    return keys, raw_keys, counts


def read_not_shipped_rows(gc, staging_spreadsheet_id):
    """Full rows (as dicts, keyed by header, PLUS a '_row_index' key -- the row's actual
    1-based row number in the sheet, e.g. 2 for the first data row right under the
    header) currently sitting in the staging 'Not Shipped' tab. A full read here is fine
    -- unlike the 3 raw sheets, this tab only ever holds what Mahmoud has manually added
    through this app, so it stays small. _row_index is what delete_not_shipped_rows
    needs to remove specific rows later; get_all_records() returns rows in the same
    order as the sheet with none skipped, so row i of the list is always sheet row
    i + 2."""
    staging_sh = _call_with_retry(lambda: gc.open_by_key(staging_spreadsheet_id))
    ws = _call_with_retry(lambda: staging_sh.worksheet(NOT_SHIPPED_TAB))
    records = _call_with_retry(lambda: ws.get_all_records())
    for i, row in enumerate(records, start=2):
        row['_row_index'] = i
    return records


def delete_not_shipped_rows(gc, staging_spreadsheet_id, row_indices):
    """Deletes the given physical row numbers (from read_not_shipped_rows' _row_index)
    from the staging 'Not Shipped' tab -- used to clean up rows find_stale_not_shipped
    flagged as since-actually-shipped (Aug 2026, per Mahmoud: a standalone cleanup, no
    file upload needed, since staleness only depends on the 3 raw sheets + this tab).
    Deletes from the BOTTOM row up so earlier deletions in the loop never shift the row
    numbers of ones still waiting to be deleted above them. Returns the count removed."""
    unique_indices = sorted(set(row_indices), reverse=True)
    if not unique_indices:
        return 0
    staging_sh = _call_with_retry(lambda: gc.open_by_key(staging_spreadsheet_id))
    ws = _call_with_retry(lambda: staging_sh.worksheet(NOT_SHIPPED_TAB))
    for idx in unique_indices:
        _call_with_retry(lambda idx=idx: ws.delete_rows(idx))
    return len(unique_indices)


def find_stale_not_shipped(not_shipped_rows, raw_keys):
    """Orders sitting in the Not Shipped tab (added there earlier as Cancelled/Pending)
    that have SINCE shown up on one of the 3 raw tracking sheets -- meaning the ops team
    actually shipped the order for real after Mahmoud logged it as not-shipped. Flags
    these so he can go remove/update them in Not Shipped, instead of leaving a
    contradictory status sitting there (counted Cancelled/Pending here, but the team
    already shipped it)."""
    stale = []
    for row in not_shipped_rows:
        ref = row.get('Reference Number', '')
        key = clean_key(ref)
        if key and key in raw_keys:
            stale.append(row)
    return stale


def classify_native_export_orders(shopify_df, target_key, mapping):
    """mapping: dict with keys ref_number / order_date / order_value / country / city /
    cancelled_at / source / financial_status / fulfillment_status / tags -- see
    NATIVE_EXPORT_DEFAULTS. order_value ('Subtotal') is populated only on each order's
    first row and blank on later line-item rows for the same order -- filtering to
    non-blank order_value both selects the first row per order and needs no sum across
    rows (see module docstring). order_value itself already excludes shipping, no
    derivation needed. An order counts as Cancelled if ANY of: cancelled_at is non-blank
    (a real timestamp, not a TRUE/1/YES flag); financial_status is 'refunded'/
    'partially_refunded' while fulfillment_status isn't 'fulfilled' (see
    REFUND_CANCEL_STATUSES); or tags contains "cancel" in any form/case (see
    CANCEL_TAG_SUBSTRING) -- none of these require anyone to have clicked Shopify's own
    "Cancel order" button.

    Returns (rows, stats)."""
    ref_col = mapping.get('ref_number')
    date_col = mapping.get('order_date')
    value_col = mapping.get('order_value')
    country_col = mapping.get('country')
    city_col = mapping.get('city')
    cancelled_col = mapping.get('cancelled_at')
    source_col = mapping.get('source')
    financial_status_col = mapping.get('financial_status')
    fulfillment_status_col = mapping.get('fulfillment_status')
    tags_col = mapping.get('tags')

    work = shopify_df.copy()
    work = work[work[ref_col].astype(str).str.strip() != '']
    if value_col:
        # The one-row-per-order filter -- see docstring. Without an order_value mapping
        # there's no reliable way to isolate one row per order in this format at all.
        work = work[work[value_col].astype(str).str.strip() != '']
    work['_key'] = work[ref_col].map(clean_key)
    work = work[work['_key'] != '']

    pos_excluded_count = 0
    if source_col and source_col in work.columns:
        # POS orders are rung up and paid in-store on the spot -- they never go through
        # the shipping company, so they don't belong in an "orders still waiting on the
        # shipping company" tracker at all (per Mahmoud, Aug 2026). Excluded outright
        # here rather than kept-and-flagged, since keeping them was inflating the "new
        # orders" count with orders that were never meant to be tracked here.
        is_pos = work[source_col].astype(str).str.strip().str.lower().isin(EXCLUDED_SOURCES)
        pos_excluded_count = int(is_pos.sum())
        work = work[~is_pos]

    default_country = ORDER_STATUS_GROUPS[target_key]['default_country']
    source_label = f"Shopify (unshipped) - {RAW_SHEETS[target_key]['label']}"
    rows = []
    blank_country_count = 0
    excluded_country_count = 0
    for _, r in work.iterrows():
        if country_col:
            country_raw_check = clean_display(r.get(country_col, '')).strip().lower()
            if country_raw_check in EXCLUDED_COUNTRY_NAMES:
                excluded_country_count += 1
                continue

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
        if not is_cancelled and financial_status_col:
            fin_status = str(r.get(financial_status_col, '')).strip().lower()
            if fin_status in REFUND_CANCEL_STATUSES:
                fulfil_status = (
                    str(r.get(fulfillment_status_col, '')).strip().lower()
                    if fulfillment_status_col else ''
                )
                if fulfil_status != 'fulfilled':
                    is_cancelled = True
        if not is_cancelled and tags_col:
            # clean_display strips invisible/combining Unicode noise first -- a Tags
            # cell with a stray hidden character sitting inside the word would otherwise
            # silently defeat the substring/fuzzy check in _tag_implies_cancelled.
            tags_val = clean_display(r.get(tags_col, '')).lower()
            if _tag_implies_cancelled(tags_val):
                is_cancelled = True

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

    stats = {
        'orders_total': len(rows),
        'blank_country_count': blank_country_count,
        'pos_excluded_count': pos_excluded_count,
        'excluded_country_count': excluded_country_count,
    }
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
    """Appends the given rows at the bottom of the live staging Not Shipped tab (Aug
    2026, per Mahmoud -- this app writes here now, not straight into Orders, so he can
    review this holding list separately). Returns the number of rows appended. Same
    retry wrapping as read_existing_keys/_col_keys, so a transient 429/5xx here (e.g.
    right after a big read) doesn't fail the whole add."""
    if not rows:
        return 0
    staging_sh = _call_with_retry(lambda: gc.open_by_key(staging_spreadsheet_id))
    ws = _call_with_retry(lambda: staging_sh.worksheet(NOT_SHIPPED_TAB))

    # If Mahmoud just created the tab by hand and it's still completely empty, seed the
    # header row first so the appended rows don't land as if they were the header.
    existing_header = _call_with_retry(lambda: ws.row_values(1))
    if not existing_header:
        _call_with_retry(lambda: ws.append_row(ORDERS_HEADER))

    now_str = dt.datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')
    values = [row_to_values(r, now_str) for r in rows]
    _call_with_retry(lambda: ws.append_rows(values))
    return len(values)
