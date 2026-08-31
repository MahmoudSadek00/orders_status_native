import json
import datetime as dt

import pandas as pd
import streamlit as st

from logic import (
    RAW_SHEETS, ORDER_STATUS_GROUPS, NATIVE_EXPORT_DEFAULTS, STAGING_SPREADSHEET_ID_DEFAULT,
    NOT_SHIPPED_TAB, get_client, read_existing_keys, read_not_shipped_rows,
    find_stale_not_shipped, classify_native_export_orders, filter_new, rows_to_dataframe,
    rows_to_excel_bytes, append_to_staging, read_any, default_mapping,
)

st.set_page_config(page_title="Orders Status Check (Native Export)", layout="wide")
st.title("Orders Status Check -- Native \"Export orders\" file")
st.caption(
    "Standalone tool (Aug 2026, per Mahmoud) for Shopify's own native \"Export orders\" "
    "file (Orders page -> Export in Shopify itself -- one row per line item, with "
    "Subtotal / Financial Status / Cancelled at columns). Order Value here is that "
    "file's own Subtotal (goods only, shipping already excluded -- no subtraction "
    "needed). Upload a full export (every status, not just shipped) for ONE country "
    "group. This checks every order against the 3 raw tracking sheets AND the staging "
    f"Orders + \"{NOT_SHIPPED_TAB}\" tabs, and shows only the orders that aren't logged "
    "anywhere yet -- so Cancelled orders and orders still awaiting shipment finally get "
    f"counted, without ever duplicating a row that's already there. New orders get "
    f"added to the \"{NOT_SHIPPED_TAB}\" tab (not straight into Orders), and every run "
    f"also flags anything sitting in \"{NOT_SHIPPED_TAB}\" that has since actually shown "
    "up on a raw sheet -- meaning the team shipped it for real after it was logged here. "
    "The 3 raw sheets themselves are never touched -- read-only."
)


def _load_creds_info():
    # st.secrets raises StreamlitSecretNotFoundError (not just a missing-key error) when
    # no secrets.toml exists on Streamlit Cloud AT ALL yet -- e.g. right after deploying
    # this app for the first time, before the README's one-time setup step is done.
    try:
        if 'gcp_service_account' in st.secrets:
            return dict(st.secrets['gcp_service_account'])
        if 'gcp_service_account_json' in st.secrets:
            raw = st.secrets['gcp_service_account_json']
            return json.loads(raw) if isinstance(raw, str) else dict(raw)
    except Exception:
        pass
    return None


def _load_staging_id():
    # Falls back to Mahmoud's real staging sheet (baked into logic.py, Aug 2026) if the
    # secret isn't set -- the only secret actually required is the credential below. A
    # staging_spreadsheet_id secret, if set, still overrides this (e.g. if the staging
    # sheet ever moves).
    try:
        return st.secrets.get('staging_spreadsheet_id') or STAGING_SPREADSHEET_ID_DEFAULT
    except Exception:
        return STAGING_SPREADSHEET_ID_DEFAULT


creds_info = _load_creds_info()
staging_id = _load_staging_id()

if not creds_info or not staging_id:
    st.error(
        "This app needs Google Sheets access configured once first -- see the README's "
        "one-time setup section for exactly what to paste into this app's "
        "Settings -> Secrets on Streamlit Cloud."
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

st.header("1. Country group")
target_key = st.selectbox(
    "Which group is this Shopify export for?",
    options=list(ORDER_STATUS_GROUPS.keys()),
    format_func=lambda k: RAW_SHEETS[k]['label'],
)

st.header("2. Shopify \"Export orders\" file")
shopify_file = st.file_uploader(
    "Shopify native export (csv or xlsx) -- Orders page -> Export -- every status, not just shipped",
    type=['csv', 'xlsx', 'xls'],
)

if shopify_file is not None:
    shopify_df = read_any(shopify_file)
    st.write(f"{len(shopify_df)} rows loaded.")

    guessed = default_mapping(shopify_df.columns.tolist(), NATIVE_EXPORT_DEFAULTS)
    fields_needed = [
        'ref_number', 'order_date', 'order_value', 'country', 'city', 'cancelled_at',
        'source', 'financial_status', 'fulfillment_status',
    ]
    labels = {
        'ref_number': 'Reference / Order Number ("Name")', 'order_date': 'Order date ("Created at")',
        'order_value': 'Order value, excl. shipping ("Subtotal")', 'country': 'Shipping country',
        'city': 'Shipping city', 'cancelled_at': 'Cancelled-at column (non-blank = cancelled)',
        'source': 'Sales channel ("Source") -- excludes POS/in-store orders',
        'financial_status': 'Financial status (catches refunded-before-shipped as Cancelled too)',
        'fulfillment_status': 'Fulfillment status (used with Financial status above)',
    }
    mapping = {}
    cols = st.columns(3)
    for i, field in enumerate(fields_needed):
        with cols[i % 3]:
            options = ['(none)'] + shopify_df.columns.tolist()
            default = guessed.get(field)
            idx = options.index(default) if default in options else 0
            choice = st.selectbox(labels[field], options, index=idx, key=f'native_{target_key}_{field}')
            mapping[field] = None if choice == '(none)' else choice
    if not mapping.get('cancelled_at'):
        st.caption("No Cancelled-at column mapped -- every new order here will be added as Pending.")
    if not mapping.get('source'):
        st.warning(
            "No Source column mapped -- POS/in-store orders won't be excluded, which "
            "will inflate the count (they're sold and paid on the spot, never shipped)."
        )
    if not mapping.get('financial_status'):
        st.caption(
            "No Financial status column mapped -- a refunded-before-it-ever-shipped "
            "order won't be caught as Cancelled unless Cancelled-at is also set."
        )

    ready = bool(mapping.get('ref_number'))
    if not ready:
        st.info("Pick the Reference/Order Number column to continue.")
    if not mapping.get('order_value'):
        st.warning("No Order value column mapped -- Order Value will be written as blank for every new order.")

    st.header("3. Check & add")
    if st.button("Check against the raw sheets + staging", disabled=not ready, type="primary"):
        try:
            with st.spinner("Reading the 3 raw sheets and the staging Orders + Not Shipped tabs..."):
                existing_keys, raw_keys, counts = read_existing_keys(gc, staging_id)
        except Exception as e:
            st.error(f"Couldn't read the sheets: {e}")
            st.stop()

        st.caption(
            f"Already logged -- UAE & Oman: {counts.get('raw:uae_om', 0)}, "
            f"Gulf: {counts.get('raw:gulf', 0)}, Iraq: {counts.get('raw:iraq', 0)} "
            f"(raw sheets), {counts.get('staging_orders', 0)} row(s) in the staging "
            f"Orders tab, {counts.get('not_shipped', 0)} row(s) in "
            f"\"{NOT_SHIPPED_TAB}\" -- {len(existing_keys)} unique order(s) total once "
            f"de-duplicated."
        )

        try:
            with st.spinner(f"Checking \"{NOT_SHIPPED_TAB}\" for orders the team has since shipped..."):
                not_shipped_rows = read_not_shipped_rows(gc, staging_id)
                stale = find_stale_not_shipped(not_shipped_rows, raw_keys)
        except Exception as e:
            stale = []
            st.warning(f"Couldn't run the staleness check on \"{NOT_SHIPPED_TAB}\": {e}")

        if stale:
            with st.expander(
                f"⚠️ {len(stale)} order(s) in \"{NOT_SHIPPED_TAB}\" have since "
                f"shown up on a raw sheet -- the team shipped them for real, remove/"
                f"update them in \"{NOT_SHIPPED_TAB}\"",
                expanded=True,
            ):
                st.dataframe(pd.DataFrame(stale), use_container_width=True)

        rows, stats = classify_native_export_orders(shopify_df, target_key, mapping)
        # Back to tracking both Cancelled AND Pending (Aug 2026, per Mahmoud -- reverted
        # the Cancelled-only restriction; a brand-new Pending order just shows up as
        # Pending normally, no grace period held back for it).
        new_rows = filter_new(rows, existing_keys)
        n_canceled = sum(1 for r in new_rows if r['status'] == 'Cancelled')
        n_pending = len(new_rows) - n_canceled

        st.success(
            f"{stats['orders_total']} distinct order(s) in this export "
            f"({stats['pos_excluded_count']} POS/in-store order(s) excluded -- never "
            f"shipped, so never tracked here). {len(rows) - len(new_rows)} already "
            f"logged somewhere (a raw sheet, the staging sheet, or a previous run of "
            f"this tool) -- skipped, never duplicated. **{len(new_rows)} NEW order(s) "
            f"to add**: {n_canceled} Cancelled, {n_pending} Pending. "
            f"{stats['blank_country_count']} had no Shipping country on file despite "
            f"not being POS (Draft order, expected -- defaulted, not flagged)."
        )

        review = [r for r in new_rows if r['needs_review']]
        if review:
            with st.expander(f"Needs a look -- unrecognized Shipping country ({len(review)})"):
                st.dataframe(rows_to_dataframe(review), use_container_width=True)

        st.session_state['new_rows'] = new_rows
        st.session_state['new_rows_target_key'] = target_key

    new_rows = st.session_state.get('new_rows')
    if new_rows is not None and st.session_state.get('new_rows_target_key') == target_key:
        st.subheader(f"{len(new_rows)} new order(s) ready")
        if new_rows:
            preview_df = rows_to_dataframe(new_rows)
            st.dataframe(preview_df, use_container_width=True)

            c1, c2 = st.columns(2)
            with c1:
                today = dt.date.today().strftime('%Y%m%d')
                st.download_button(
                    "Download Excel (for your own records)",
                    data=rows_to_excel_bytes(new_rows),
                    file_name=f"orders_status_native_{target_key}_{today}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with c2:
                if st.button(f"Add these {len(new_rows)} row(s) to \"{NOT_SHIPPED_TAB}\"", type="primary"):
                    try:
                        with st.spinner(f"Appending to \"{NOT_SHIPPED_TAB}\"..."):
                            n = append_to_staging(gc, staging_id, new_rows)
                        st.success(
                            f"{n} row(s) appended to \"{NOT_SHIPPED_TAB}\". Re-run this "
                            f"check any time to catch any of them that the team has "
                            f"since actually shipped (see the staleness warning above)."
                        )
                        st.session_state.pop('new_rows', None)
                    except Exception as e:
                        st.error(f"Couldn't append to \"{NOT_SHIPPED_TAB}\": {e}")
        else:
            st.info("Nothing new -- every order in this export is already logged somewhere.")
