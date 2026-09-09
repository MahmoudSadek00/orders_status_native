"""
Logic for the "Customer Support" section of this app (Sep 2026, per Mahmoud) -- appends
Agent Activity / Chats / Calls native export files into the live "Customer Support Raw
Data" Google Sheet (the same sheet CS Pulse reads from), without ever duplicating a row
already added on a previous run.

Kept in its own module, separate from logic.py, on purpose -- this has nothing to do
with orders. It DOES import a few data-type-agnostic helpers from logic.py (the
multi-file/zip reader, the Google client, the retry wrapper, the hidden-character-safe
string cleaners) rather than duplicating them, since those have no order-specific logic
in them at all.

Confirmed against real sample export files, Sep 2026:

  - Calls and Chats each carry a genuine unique ID column in their native export --
    'ID' for Calls, 'Conversation ID' for Chats -- and in BOTH cases the export's own
    columns are an EXACT match (same names, same count, same order) to the live sheet's
    own 'Calls' / 'Chats' tab headers (verified against a real export of the live sheet
    itself: zero columns different either direction). So for these two, dedup is just
    clean_key() of that one ID column, and rows are appended essentially as-is, aligned
    to the live tab's own header order (not just this file's own column list) in case
    someone's added/reordered a column on the sheet by hand since.

  - Agent Activity has no unique-per-row ID -- it's a state-change event log (native
    export columns: Agent Name, Timestamp, State). Its dedup key is Agent Name +
    Timestamp together. The live "Agents Activity" tab's Timestamp column mixes plain
    'DD-MM-YYYY HH:MM:SS' text (what a fresh native-export upload always writes) with
    some legacy Excel-auto-converted date cells from past manual imports (the same
    ambiguity cs_dashboard/logic.py's fix_activity_ts works around, on the CS Pulse
    dashboard side) -- _activity_ts_key below normalizes both forms to the same
    comparable string, so a state-change already logged in EITHER form is never
    re-added. New rows are appended via RAW value_input_option specifically so this app
    itself never creates another ambiguous auto-converted cell -- see append_new_rows.
"""
import datetime as dt

import gspread
import pandas as pd
from gspread.utils import ValueInputOption, ValueRenderOption

from logic import get_client, read_headers, read_many, clean_display, clean_key, _call_with_retry  # noqa: F401

# Same live sheet CS Pulse (the "GC CS Performance Report" dashboard) reads from.
CS_SPREADSHEET_ID_DEFAULT = '1Lz9OaWLpEM-m9w-5bxITTPKs9e00ZfuGtCl3m1Iicpw'

GOOGLE_SHEETS_EPOCH = dt.date(1899, 12, 30)  # same epoch cs_dashboard/logic.py uses

DATA_TYPES = {
    'agent_activity': {
        'label': 'Agent Activity',
        'tab': 'Agents Activity',
        'columns': ['Agent Name', 'Timestamp', 'State'],
        'key_columns': ['Agent Name', 'Timestamp'],
    },
    'calls': {
        'label': 'Calls',
        'tab': 'Calls',
        'columns': [
            'ID', 'Type', 'Agent', 'Created', 'Waiting Duration', 'Handling Duration',
            'Holding Duration', 'Ringing Duration', 'Duration', 'State',
            'Answering Machine Detected', 'Voicemail', 'Caller', 'Callee', 'Inputs',
            'Flags', 'Tags', 'Auto Tags', 'Sentiment', 'Call Summary',
            'Hangup Direction', 'Recording Links', 'Notes',
        ],
        'key_columns': ['ID'],
    },
    'chats': {
        'label': 'Chats',
        'tab': 'Chats',
        'columns': [
            'Conversation ID', 'DateTime Conversation Started',
            'DateTime Conversation Resolved', 'Contact ID', 'Assignee',
            'Number of Outgoing Messages', 'Number of Incoming Messages',
            'DateTime First Response', 'Conversation Category',
            'Closing Note Summary', 'Closed By', 'First Response Time',
            'Resolution Time', 'First Assignee', 'Closed By Source',
            'Closed By Team', 'Opened By Source', 'Opened By Channel',
            'First Assignment Timestamp', 'First Response By',
            'Last Assignment Timestamp', 'Last Assignee',
            'Time to First Assignment',
            'First Assignment to First Response Time',
            'Last Assignment to Response Time',
            'First Assignment to Close Time', 'Last Assignment to Close Time',
            'Average Response Time', 'Number of Assignments',
            'Number of Responses',
        ],
        'key_columns': ['Conversation ID'],
    },
}


def _id_key(v):
    """clean_key() of an ID cell, tolerant of it coming back as a float (335828368.0)
    instead of a string/int -- Sheets sometimes stores a plain digit string as a number
    after a manual paste/import. Returns None (never '') for anything blank, so two
    blanks are never treated as matching each other."""
    if v is None:
        return None
    if isinstance(v, float):
        if pd.isna(v):
            return None
        if v.is_integer():
            v = int(v)
    k = clean_key(v)
    return k or None


def _activity_ts_key(v):
    """Canonical 'YYYY-MM-DD HH:MM:SS' string for an Agents Activity Timestamp cell,
    whichever of the two forms it's stored in on the live sheet (see module docstring):
    a Sheets date-serial NUMBER (from a past auto-converted import, unambiguous -- no
    day/month guessing needed for a real serial value) or plain 'DD-MM-YYYY HH:MM:SS'
    TEXT (what this app itself always writes). Returns None if the cell can't be read as
    a date at all, rather than guessing."""
    if v is None or v == '':
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        try:
            ts = dt.datetime.combine(GOOGLE_SHEETS_EPOCH, dt.time()) + dt.timedelta(days=float(v))
        except (OverflowError, ValueError, OSError):
            return None
        return ts.strftime('%Y-%m-%d %H:%M:%S')
    s = clean_display(v)
    if not s:
        return None
    try:
        return dt.datetime.strptime(s, '%d-%m-%Y %H:%M:%S').strftime('%Y-%m-%d %H:%M:%S')
    except ValueError:
        pass
    parsed = pd.to_datetime(s, dayfirst=True, errors='coerce')
    if pd.isna(parsed):
        return None
    return parsed.strftime('%Y-%m-%d %H:%M:%S')


def row_key(data_type, row):
    """row: a dict/Series of column name -> value (works for both a freshly-uploaded
    file's row and a raw live-sheet cell row). Returns one string key, or None if this
    particular row can't be matched/deduped at all (missing whichever field(s) the key
    depends on) -- a None-keyed row is never treated as "new" on its own since it can
    never be told apart from another None-keyed row either; see filter_new."""
    if data_type == 'calls':
        return _id_key(row.get('ID'))
    if data_type == 'chats':
        return _id_key(row.get('Conversation ID'))
    if data_type == 'agent_activity':
        name = clean_key(row.get('Agent Name'))
        ts = _activity_ts_key(row.get('Timestamp'))
        if not name or not ts:
            return None
        return f"{name}||{ts}"
    raise ValueError(f"unknown data_type {data_type!r}")


def read_existing_keys(gc, spreadsheet_id, data_type):
    """Reads only the column(s) row_key() actually needs from the live tab -- not the
    whole sheet -- same reasoning as orders_status_native's own _col_keys (a large real
    sheet, and this CS one is large, can trip a 503 if asked to read every column just
    to check one or two of them). Returns (keys, existing_row_count). An empty/missing
    tab, or one that doesn't even have the key column(s) yet, is treated as "no rows
    logged yet" rather than raising -- append_new_rows still seeds a fresh header row
    the first time it appends to a genuinely empty tab."""
    cfg = DATA_TYPES[data_type]
    sh = _call_with_retry(lambda: gc.open_by_key(spreadsheet_id))
    try:
        ws = _call_with_retry(lambda: sh.worksheet(cfg['tab']))
    except gspread.WorksheetNotFound:
        return set(), 0

    header = _call_with_retry(lambda: ws.row_values(1))
    if not header:
        return set(), 0

    needed = cfg['key_columns']
    cols = {}
    for col_name in needed:
        if col_name not in header:
            return set(), 0  # tab exists but doesn't have the key column -- can't match against it
        idx = header.index(col_name)
        cols[col_name] = _call_with_retry(
            lambda idx=idx: ws.col_values(idx + 1, value_render_option=ValueRenderOption.unformatted)
        )

    n_rows = max((len(v) for v in cols.values()), default=1) - 1  # minus the header cell
    keys = set()
    for i in range(1, n_rows + 1):
        row = {col_name: (vals[i] if i < len(vals) else None) for col_name, vals in cols.items()}
        k = row_key(data_type, row)
        if k:
            keys.add(k)
    return keys, max(n_rows, 0)


def prepare_upload(data_type, files, on_progress=None):
    """Reads the uploaded native export file(s)/zip(s) (see logic.read_many -- handles
    multiple files and zips, memory-safe on large exports) restricted to just this data
    type's expected columns, strips stray whitespace from every cell, and drops fully
    blank rows. Returns (df, stats) where stats is logic.read_many's own
    files_read/files_skipped info."""
    cfg = DATA_TYPES[data_type]
    df, stats = read_many(files, usecols=set(cfg['columns']), on_progress=on_progress)
    if df.empty:
        return df, stats
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    non_blank = df.apply(lambda r: any(v for v in r), axis=1)
    df = df[non_blank].reset_index(drop=True)
    return df, stats


def dedupe_and_filter(data_type, df, existing_keys):
    """Drops rows already logged on the live sheet (existing_keys) AND rows duplicated
    WITHIN this same upload (e.g. two overlapping-date export files dragged on
    together) -- keeping the first occurrence of each new key. Rows with no usable key
    at all are dropped rather than guessed at."""
    seen = set(existing_keys)
    keep = []
    for _, row in df.iterrows():
        k = row_key(data_type, row)
        if k and k not in seen:
            keep.append(True)
            seen.add(k)
        else:
            keep.append(False)
    return df[keep].reset_index(drop=True)


def _chunked(seq, size):
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def append_new_rows(gc, spreadsheet_id, data_type, df):
    """Appends df (already deduped -- see dedupe_and_filter) to the bottom of the live
    tab, column-aligned to the TAB'S OWN live header order (not just this file's column
    list, in case a column's been added/reordered by hand on the sheet since). Written
    with RAW value_input_option specifically -- see module docstring on why (keeps
    Agent Activity's Timestamp text exactly as the native export wrote it, and never
    lets Sheets reinterpret/auto-convert any cell as a side effect of this app's own
    write). Sent in chunks of 500 rows per API call so one very large add can't build an
    oversized single request. Returns the number of rows appended."""
    cfg = DATA_TYPES[data_type]
    if df.empty:
        return 0
    sh = _call_with_retry(lambda: gc.open_by_key(spreadsheet_id))
    try:
        ws = _call_with_retry(lambda: sh.worksheet(cfg['tab']))
    except gspread.WorksheetNotFound:
        raise RuntimeError(
            f"The sheet has no '{cfg['tab']}' tab -- create it by hand first (header "
            f"row: {', '.join(cfg['columns'])}, or leave it completely empty and this "
            f"app will add that header itself on the first add)."
        )

    existing_header = _call_with_retry(lambda: ws.row_values(1))
    if not existing_header:
        _call_with_retry(lambda: ws.append_row(cfg['columns']))
        existing_header = cfg['columns']

    out = df.reindex(columns=existing_header, fill_value='')
    values = out.astype(str).values.tolist()
    for chunk in _chunked(values, 500):
        _call_with_retry(lambda chunk=chunk: ws.append_rows(chunk, value_input_option=ValueInputOption.raw))
    return len(values)
