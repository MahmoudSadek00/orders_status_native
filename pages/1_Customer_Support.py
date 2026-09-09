import io
import json
import datetime as dt

import streamlit as st

from cs_logic import (
    DATA_TYPES, CS_SPREADSHEET_ID_DEFAULT, get_client, read_headers, prepare_upload,
    read_existing_keys, dedupe_and_filter, append_new_rows,
)


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
    "(or dragging on two overlapping files at once) never duplicates a row. Calls and "
    "Chats are matched by their own unique ID column; Agent Activity (a state-change "
    "log with no ID of its own) is matched by Agent + Timestamp together."
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
        progress_bar = st.progress(0.0, text="Reading uploaded file(s)...")

        def _on_progress(done, total, label):
            frac = done / total if total else 1.0
            progress_bar.progress(frac, text=f"Reading file {done}/{total}: {label}")

        df, load_stats = prepare_upload(data_type, files, on_progress=_on_progress)
        progress_bar.empty()
        st.write(
            f"{len(df)} row(s) loaded from {len(load_stats['files_read'])} file(s) "
            f"across {len(files)} upload(s)."
        )
        if load_stats['files_skipped']:
            with st.expander(f"⚠️ {len(load_stats['files_skipped'])} item(s) skipped"):
                for nm, reason in load_stats['files_skipped']:
                    st.write(f"- {nm}: {reason}")
        if df.empty:
            st.error("No readable rows found in what was uploaded.")
            st.stop()

        try:
            with st.spinner(f"Reading the live \"{cfg['tab']}\" tab..."):
                existing_keys, existing_count = read_existing_keys(gc, cs_spreadsheet_id, data_type)
        except Exception as e:
            st.error(f"Couldn't read the sheet: {e}")
            st.stop()

        new_df = dedupe_and_filter(data_type, df, existing_keys)
        n_dupe = len(df) - len(new_df)

        st.success(
            f"{existing_count} row(s) already in \"{cfg['tab']}\" right now. Of the "
            f"{len(df)} row(s) in this upload: {n_dupe} already logged (or duplicated "
            f"within this same upload) -- skipped. **{len(new_df)} NEW row(s) to add.**"
        )

        st.session_state['cs_new_df'] = new_df
        st.session_state['cs_new_df_type'] = data_type

    new_df = st.session_state.get('cs_new_df')
    if new_df is not None and st.session_state.get('cs_new_df_type') == data_type:
        st.subheader(f"{len(new_df)} new row(s) ready")
        if len(new_df):
            st.dataframe(new_df, use_container_width=True)

            c1, c2 = st.columns(2)
            with c1:
                today = dt.date.today().strftime('%Y%m%d')
                st.download_button(
                    "Download Excel (for your own records)",
                    data=_to_excel_bytes(new_df),
                    file_name=f"cs_{data_type}_{today}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with c2:
                if st.button(f"Add these {len(new_df)} row(s) to \"{cfg['tab']}\"", type="primary"):
                    try:
                        with st.spinner(f"Appending to \"{cfg['tab']}\"..."):
                            n = append_new_rows(gc, cs_spreadsheet_id, data_type, new_df)
                        st.success(
                            f"{n} row(s) appended to \"{cfg['tab']}\". Re-run this check "
                            f"any time -- anything already added is recognized and "
                            f"skipped automatically, so running it twice on the same "
                            f"export is harmless."
                        )
                        st.session_state.pop('cs_new_df', None)
                    except Exception as e:
                        st.error(f"Couldn't append to \"{cfg['tab']}\": {e}")
        else:
            st.info(f"Nothing new -- every row in this upload is already in \"{cfg['tab']}\".")
