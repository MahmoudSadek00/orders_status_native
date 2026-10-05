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
import time
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
        # Oct 2026: State added to the key -- the live tab has 60 real cases of the same
        # agent logging two different states in the same second (e.g. Offline then
        # Available at 01-08-2026 22:40:03). Agent + Timestamp alone treated the second
        # event as a duplicate and silently dropped it.
        'key_columns': ['Agent Name', 'Timestamp', 'State'],
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
        # Oct 2026: NOT Conversation ID on its own any more. The IDs are 16+ digit
        # numbers that have lost their last digits somewhere upstream (number
        # precision -- 39,300 of the live tab's IDs end in ...999996), so two different
        # chats started in the same second end up with the SAME ID: 432 such pairs
        # already exist on the live tab, each with a different Contact ID. Keyed on
        # Conversation ID alone, the second chat of every such pair is silently dropped
        # as a "duplicate". Contact ID + start time is unique on the live tab (one
        # contact can't open two chats in the same second) and doesn't depend on the
        # damaged digits. Rows missing either fall back to Conversation ID.
        'key_columns': ['Contact ID', 'DateTime Conversation Started'],
        'fallback_key_column': 'Conversation ID',
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
    # '335828368.0' -- what an ID cell turns into when an xlsx export stores it as a
    # number and it's then read as text. Without this it never matches '335828368'.
    if k.endswith('.0') and k[:-2].isdigit():
        k = k[:-2]
    return k or None


def _datetime_key(v):
    """Canonical 'YYYY-MM-DD HH:MM:SS' for a Chats start-time cell, whether it's
    stored as plain ISO text (what this app writes) or as a Sheets date serial number
    (a legacy auto-converted import). ISO text has no day/month ambiguity."""
    if v is None or v == '' or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        try:
            ts = dt.datetime.combine(GOOGLE_SHEETS_EPOCH, dt.time()) + dt.timedelta(days=float(v))
        except (OverflowError, ValueError, OSError):
            return None
        return (ts + dt.timedelta(microseconds=500000)).replace(microsecond=0).strftime('%Y-%m-%d %H:%M:%S')
    s = clean_display(v)
    if not s:
        return None
    try:  # fast path -- the native export's own format
        return dt.datetime.strptime(s, '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d %H:%M:%S')
    except ValueError:
        pass
    parsed = pd.to_datetime(s, errors='coerce')
    if pd.isna(parsed):
        return None
    return parsed.strftime('%Y-%m-%d %H:%M:%S')


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


def _activity_ts_candidates(v):
    """Like _activity_ts_key, but returns a SET of every canonical reading a legacy
    EXISTING sheet cell could plausibly be -- TWO candidates when the day and month are
    both <=12 and therefore genuinely ambiguous. This matters specifically for a live
    sheet that's accumulated rows from more than one source over time: some rows this
    app itself appends (always the literal, correct DD-MM-YYYY text, never ambiguous on
    that side), but older rows may have been entered by some other process that stored
    an ambiguous date the OTHER way round (day and month swapped) without anyone
    noticing, since e.g. '08-09-2026' and '09-08-2026' are both valid-looking dates.
    Used only when building the EXISTING-side key set (read_existing_keys) -- a freshly
    uploaded native-export row's OWN timestamp is never ambiguous, so upload-side
    matching (row_key, used by dedupe_and_filter) stays single-candidate on purpose."""
    if v is None or v == '' or isinstance(v, bool):
        return set()
    base = None
    if isinstance(v, (int, float)):
        try:
            base = dt.datetime.combine(GOOGLE_SHEETS_EPOCH, dt.time()) + dt.timedelta(days=float(v))
        except (OverflowError, ValueError, OSError):
            return set()
    else:
        s = clean_display(v)
        if not s:
            return set()
        try:
            base = dt.datetime.strptime(s, '%d-%m-%Y %H:%M:%S')
        except ValueError:
            parsed = pd.to_datetime(s, dayfirst=True, errors='coerce')
            if pd.isna(parsed):
                return set()
            base = parsed.to_pydatetime()

    out = {base.strftime('%Y-%m-%d %H:%M:%S')}
    if base.day <= 12 and base.month != base.day:
        try:
            swapped = base.replace(month=base.day, day=base.month)
            out.add(swapped.strftime('%Y-%m-%d %H:%M:%S'))
        except ValueError:
            pass
    return out


def row_key(data_type, row):
    """row: a dict/Series of column name -> value (works for both a freshly-uploaded
    file's row and a raw live-sheet cell row). Returns one string key, or None if this
    particular row can't be matched/deduped at all (missing whichever field(s) the key
    depends on) -- a None-keyed row is never treated as "new" on its own since it can
    never be told apart from another None-keyed row either; see filter_new."""
    if data_type == 'calls':
        return _id_key(row.get('ID'))
    if data_type == 'chats':
        contact = _id_key(row.get('Contact ID'))
        started = _datetime_key(row.get('DateTime Conversation Started'))
        if contact and started:
            return f"{contact}||{started}"
        cid = _id_key(row.get('Conversation ID'))
        return f"id:{cid}" if cid else None
    if data_type == 'agent_activity':
        name = clean_key(row.get('Agent Name'))
        ts = _activity_ts_key(row.get('Timestamp'))
        state = clean_key(row.get('State'))
        if not name or not ts:
            return None
        return f"{name}||{ts}||{state}"
    raise ValueError(f"unknown data_type {data_type!r}")


def _find_col(header, name):
    """Case/whitespace-tolerant column lookup -- returns the 0-based index of the first
    header cell matching `name` once both are stripped and casefolded, or None. Guards
    against a live header that's technically the same column but not a byte-for-byte
    match (stray trailing space, different case) silently making read_existing_keys
    think the whole tab is empty."""
    target = name.strip().casefold()
    for i, h in enumerate(header):
        if str(h).strip().casefold() == target:
            return i
    return None


def read_existing_keys(gc, spreadsheet_id, data_type):
    """Reads only the column(s) needed to de-duplicate from the live tab -- not the
    whole sheet -- same reasoning as orders_status_native's own _col_keys (a large real
    sheet, and this CS one is large, can trip a 503 if asked to read every column just
    to check one or two of them). Returns (keys, existing_row_count, sample) where
    sample is up to 3 (raw_row_values, computed_key(s)) pairs -- purely diagnostic, so a
    mismatch between what's actually on the sheet and what this app computes is visible
    on screen instead of having to be guessed at. An empty/missing tab, or one that
    doesn't even have the key column(s) yet, is treated as "no rows logged yet" rather
    than raising -- append_new_rows still seeds a fresh header row the first time it
    appends to a genuinely empty tab.

    Agent Activity gets EVERY plausible canonical reading of each existing row's
    Timestamp added to the key set (see _activity_ts_candidates) -- not just one -- to
    also catch a legacy row whose day/month may have been swapped by whatever process
    originally wrote it, long before this app existed. Calls/Chats have a genuine
    unique ID so there's no such ambiguity to account for."""
    cfg = DATA_TYPES[data_type]
    sh = _call_with_retry(lambda: gc.open_by_key(spreadsheet_id))
    try:
        ws = _call_with_retry(lambda: sh.worksheet(cfg['tab']))
    except gspread.WorksheetNotFound:
        return set(), 0, []

    header = _call_with_retry(lambda: ws.row_values(1))
    if not header:
        return set(), 0, []

    needed = list(cfg['key_columns'])
    if cfg.get('fallback_key_column'):
        needed.append(cfg['fallback_key_column'])
    cols = {}
    for col_name in needed:
        idx = _find_col(header, col_name)
        if idx is None:
            if col_name == cfg.get('fallback_key_column'):
                continue
            # Oct 2026: used to return "no rows logged" here, which made EVERY uploaded
            # row look new -- a renamed key column would have duplicated the whole tab.
            raise RuntimeError(
                f"The '{cfg['tab']}' tab has no '{col_name}' column, so this app can't "
                f"check for duplicates. Fix the header on the sheet first."
            )
        cols[col_name] = _call_with_retry(
            lambda idx=idx: ws.col_values(idx + 1, value_render_option=ValueRenderOption.unformatted)
        )

    n_rows = max((len(v) for v in cols.values()), default=1) - 1  # minus the header cell
    keys = set()
    sample = []
    for i in range(1, n_rows + 1):
        row = {col_name: (vals[i] if i < len(vals) else None) for col_name, vals in cols.items()}
        if data_type == 'agent_activity':
            name = clean_key(row.get('Agent Name'))
            state = clean_key(row.get('State'))
            ts_candidates = _activity_ts_candidates(row.get('Timestamp'))
            row_keys = {f"{name}||{ts}||{state}" for ts in ts_candidates} if name and ts_candidates else set()
            keys |= row_keys
            if len(sample) < 3:
                sample.append((row, sorted(row_keys)))
        else:
            k = row_key(data_type, row)
            row_keys = [k] if k else []
            if data_type == 'chats':
                # Also index the existing row by its Conversation ID, so an upload row
                # that's missing Contact ID / start time (and falls back to the ID) can
                # still be matched against it.
                cid = _id_key(row.get('Conversation ID'))
                if cid:
                    row_keys.append(f"id:{cid}")
            keys.update(row_keys)
            if len(sample) < 3:
                sample.append((row, row_keys))
    return keys, max(n_rows, 0), sample


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
    # Vectorized (Oct 2026) -- the old row-by-row apply() was very slow on a big
    # first-time upload of 100k+ rows.
    non_blank = (df != '').any(axis=1)
    df = df[non_blank].reset_index(drop=True)
    return df, stats


def dedupe_and_filter(data_type, df, existing_keys):
    """Drops rows already logged on the live sheet (existing_keys) AND rows duplicated
    WITHIN this same upload (e.g. two overlapping-date export files dragged on
    together) -- keeping the first occurrence of each new key. Rows with no usable key
    at all are dropped rather than guessed at."""
    seen = set(existing_keys)
    keep = []
    # to_dict('records') instead of iterrows() -- several times faster on a large
    # upload, same values.
    for row in df.to_dict('records'):
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
    # Oct 2026: 5,000 rows per call instead of 500. Google allows ~60 write requests a
    # minute, so a 120k-row first-time Chats upload at 500 rows/call needed 240 calls
    # and ran into the quota part-way. At 5,000 rows it's 24 calls (each well under
    # Google's request-size limit for this data). A short pause between calls keeps a
    # long upload under the per-minute quota; _call_with_retry still waits out any 429.
    done = 0
    for i, chunk in enumerate(_chunked(values, 5000)):
        if i:
            time.sleep(1.5)
        _call_with_retry(lambda chunk=chunk: ws.append_rows(chunk, value_input_option=ValueInputOption.raw))
        done += len(chunk)
    return done
