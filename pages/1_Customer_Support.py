import io
import json
import datetime as dt

import streamlit as st

from cs_logic import (
    DATA_TYPES, CS_SPREADSHEET_ID_DEFAULT, get_client, read_headers,
    read_existing_keys, stream_new_rows, check_cell_capacity, append_new_rows,
)

PREVIEW_ROWS = 1000        # rows shown on screen (the full set is still added/downloadable)
EXCEL_MAX_ROWS = 50_000    # above this the download is CSV only (Excel needs far more memory)


def _to_excel_bytes(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()

st.set_page_config(page_title="Customer Support", layout="wide")
st.title("Customer Support -- Agent Activity / Chats / Calls raw data")
st.caption(
    "Upload a native export (Agent Activity, Chats, or Calls) and add it to the "
    "\"Customer Support Raw Data\" sheet -- the same live sheet the CS Pulse dashboard "
    "reads from. Checks the target tab first and shows only the rows that aren't "
    "logged there yet, so re-running this on an export that overlaps a previous one "
    "(or dragging on two overlapping files at once) never duplicates a row. Calls are "
    "matched by their ID; Chats by Contact ID + start time (the Conversation ID loses "
    "its last digits, so two different chats can share one); Agent Activity by Agent + "
    "Timestamp + State."
)


def _load_creds_info():
    # st.secrets raises StreamlitSecretNotFoundError (not just a missing-key error) when
    # no secrets.toml exists on Streamlit Cloud AT ALL yet. Same handling as app.py.
    try:
        if 'gcp_service_account' in st.secrets:
            return dict(st.secrets['gcp_service_account'])
        if 'gcp_service_account_json' in st.secrets:
            raw = st.secrets['gcp_service_account_json']
            return json.loads(raw) if isinstance(raw, str) else dict(raw)
    except Exception:
        pass
    return None


def _load_cs_spreadsheet_id():
    # Falls back to the real CS Raw Data sheet (baked into cs_logic.py) if the secret
    # isn't set -- same convention as app.py's staging_spreadsheet_id override.
    try:
        return st.secrets.get('cs_spreadsheet_id') or CS_SPREADSHEET_ID_DEFAULT
    except Exception:
        return CS_SPREADSHEET_ID_DEFAULT


creds_info = _load_creds_info()
cs_spreadsheet_id = _load_cs_spreadsheet_id()

if not creds_info or not cs_spreadsheet_id:
    st.error(
        "This page needs Google Sheets access configured once first -- see the README's "
        "one-time setup section for exactly what to paste into this app's "
        "Settings -> Secrets on Streamlit Cloud. It reuses the SAME credential as the "
        "Orders Status Check page -- no new secret needed, as long as that service "
        "account also has Editor access to the Customer Support Raw Data sheet (see "
        "README)."
    )
    st.stop()


@st.cache_resource(show_spinner=False)
def _client():
    return get_client(creds_info)


try:
    gc = _client()
except Exception as e:
    st.error(f"Couldn't connect to Google Sheets with the configured credential: {e}")
    st.stop()

st.header("1. Data type")
data_type = st.selectbox(
    "Which native export is this?",
    options=list(DATA_TYPES.keys()),
    format_func=lambda k: DATA_TYPES[k]['label'],
)
cfg = DATA_TYPES[data_type]
st.caption(f"Goes into the \"{cfg['tab']}\" tab. Matched/de-duplicated by: {', '.join(cfg['key_columns'])}.")

st.header("2. File(s)")
files = st.file_uploader(
    f"{cfg['label']} native export -- csv, xls/xlsx, or .zip. Drag on more than one at "
    "once if you have several (e.g. several days' worth) -- overlapping rows across "
    "files, or against what's already on the sheet, are caught either way.",
    type=['csv', 'xlsx', 'xls', 'zip'],
    accept_multiple_files=True,
    key=f"cs_files_{data_type}",
)

if files:
    st.caption(f"{len(files)} upload(s) selected.")
    header_columns = read_headers(files)
    if not header_columns:
        st.error("Couldn't read column headers from any of the uploaded files -- check they're valid csv/xlsx/xls (or zips containing them).")
        st.stop()

    missing = [c for c in cfg['columns'] if c not in header_columns]
    extra = [c for c in header_columns if c not in cfg['columns']]
    if missing:
        st.warning(
            f"This file is missing {len(missing)} column(s) this app normally expects "
            f"for {cfg['label']}: {', '.join(missing)}. It'll still be read -- those "
            f"columns will just come through blank -- but double check this is really "
            f"a {cfg['label']} export."
            + (" The key column(s) needed to de-duplicate are still present." if not
               any(c in missing for c in cfg['key_columns']) else
               " ⚠️ At least one of the KEY column(s) needed to de-duplicate is missing "
               "-- rows here won't be matchable and will be dropped rather than risk "
               "adding a duplicate.")
        )
    if extra:
        st.caption(f"{len(extra)} column(s) in the file aren't used and will be ignored: {', '.join(extra)}.")

    st.header("3. Check & add")
    if st.button(f"Check against the live \"{cfg['tab']}\" tab", type="primary"):
        # Nov 2026: free the previous check's result before starting a new one.
        for k in ('cs_new_df', 'cs_new_df_type', 'cs_dl_bytes', 'cs_dl_name'):
            st.session_state.pop(k, None)
        # Live keys first, then the upload is read piece by piece and only NEW rows are
        # kept -- a 135 MB export used to be loaded whole and ran the app out of memory.
        try:
            with st.spinner(f"Reading the live \"{cfg['tab']}\" tab..."):
                existing_keys, existing_count, existing_sample = read_existing_keys(gc, cs_spreadsheet_id, data_type)
        except Exception as e:
            st.error(f"Couldn't read the sheet: {e}")
            st.stop()

        progress_bar = st.progress(0.0, text="Reading uploaded file(s)...")

        def _on_progress(done, total, label):
            frac = done / total if total else 1.0
            progress_bar.progress(frac, text=f"Read file {done}/{total}: {label}")

        with st.spinner("Reading the upload and checking each row (big files take a minute or two)..."):
            new_df, load_stats = stream_new_rows(data_type, files, existing_keys, on_progress=_on_progress)
        progress_bar.empty()
        del existing_keys

        if load_stats['files_skipped']:
            with st.expander(f"⚠️ {len(load_stats['files_skipped'])} item(s) skipped"):
                for nm, reason in load_stats['files_skipped']:
                    st.write(f"- {nm}: {reason}")
        rows_read = load_stats['rows_read']
        if rows_read == 0:
            st.error("No readable rows found in what was uploaded.")
            st.stop()

        n_dupe = rows_read - len(new_df)
        st.success(
            f"{existing_count:,} row(s) already in \"{cfg['tab']}\" right now. Of the "
            f"{rows_read:,} row(s) read from {len(load_stats['files_read'])} file(s): {n_dupe:,} already "
            f"logged (or duplicated within this same upload) -- skipped. "
            f"**{len(new_df):,} NEW row(s) to add.**"
        )

        with st.expander("Diagnostic -- sample of what was actually read off the sheet"):
            st.caption(
                "First 3 existing rows read from the sheet, and the matching key(s) "
                "computed from them. If this ever looks wrong (e.g. an empty key list "
                "for a row that clearly has data), that's the first thing to check."
            )
            for raw_row, computed_keys in existing_sample:
                st.write(f"Raw: {raw_row}  ->  key(s): {computed_keys or '(none -- unmatched)'}")
            if not existing_sample:
                st.write("(no existing rows found on the sheet)")

        st.session_state['cs_new_df'] = new_df
        st.session_state['cs_new_df_type'] = data_type

    new_df = st.session_state.get('cs_new_df')
    if new_df is not None and st.session_state.get('cs_new_df_type') == data_type:
        st.subheader(f"{len(new_df)} new row(s) ready")
        if len(new_df):
            # Nov 2026: only a preview goes to the browser, and the download file is
            # built only when asked for -- both used to cover every row on every rerun.
            if len(new_df) > PREVIEW_ROWS:
                st.caption(f"Showing the first {PREVIEW_ROWS:,} of {len(new_df):,} new rows.")
            st.dataframe(new_df.head(PREVIEW_ROWS), use_container_width=True)

            c1, c2 = st.columns(2)
            with c1:
                today = dt.date.today().strftime('%Y%m%d')
                as_excel = len(new_df) <= EXCEL_MAX_ROWS
                label = "Excel" if as_excel else "CSV"
                if st.session_state.get('cs_dl_bytes') is None:
                    if st.button(f"Prepare {label} download (for your own records)"):
                        with st.spinner(f"Building the {label} file..."):
                            if as_excel:
                                st.session_state['cs_dl_bytes'] = _to_excel_bytes(new_df)
                                st.session_state['cs_dl_name'] = f"cs_{data_type}_{today}.xlsx"
                            else:
                                st.session_state['cs_dl_bytes'] = new_df.to_csv(index=False).encode('utf-8-sig')
                                st.session_state['cs_dl_name'] = f"cs_{data_type}_{today}.csv"
                        st.rerun()
                    if not as_excel:
                        st.caption(f"More than {EXCEL_MAX_ROWS:,} rows, so the download is CSV (opens in Excel).")
                else:
                    name = st.session_state['cs_dl_name']
                    st.download_button(
                        f"Download {name}", data=st.session_state['cs_dl_bytes'], file_name=name,
                        mime=("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                              if name.endswith('.xlsx') else "text/csv"),
                    )
            with c2:
                if st.button(f"Add these {len(new_df)} row(s) to \"{cfg['tab']}\"", type="primary"):
                    try:
                        ok, used, needed = check_cell_capacity(gc, cs_spreadsheet_id, data_type, len(new_df))
                    except Exception:
                        ok, used, needed = True, 0, 0  # can't check -- don't block the add
                    if not ok:
                        st.error(
                            f"Not added: Google Sheets allows 10,000,000 cells per spreadsheet. "
                            f"This one already uses {used:,} and these rows need {needed:,} more. "
                            f"Delete unused tabs/empty rows/columns in the sheet, or move older "
                            f"chats to an archive spreadsheet, then try again."
                        )
                        st.stop()
                    try:
                        with st.spinner(f"Appending {len(new_df):,} rows to \"{cfg['tab']}\" (about 2 seconds per 5,000)..."):
                            n = append_new_rows(gc, cs_spreadsheet_id, data_type, new_df)
                        st.success(
                            f"{n} row(s) appended to \"{cfg['tab']}\". Re-run this check "
                            f"any time -- anything already added is recognized and "
                            f"skipped automatically, so running it twice on the same "
                            f"export is harmless."
                        )
                        for k in ('cs_new_df', 'cs_new_df_type', 'cs_dl_bytes', 'cs_dl_name'):
                            st.session_state.pop(k, None)
                    except Exception as e:
                        st.error(
                            f"Couldn't finish appending to \"{cfg['tab']}\": {e}. Some rows "
                            f"may already have been added. It's safe to press Check again "
                            f"and add the rest -- rows already added are skipped."
                        )
        else:
            st.info(f"Nothing new -- every row in this upload is already in \"{cfg['tab']}\".")
