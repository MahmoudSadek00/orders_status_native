# Orders Status Check -- Native Export (standalone tool)

A separate, standalone Streamlit app -- its own link, its own deployment, not a page
inside the consolidation tool or the other Orders Status Check app (Aug 2026, per
Mahmoud: kept apart on purpose).

**What it's for:** Shopify's own native "Export orders" file (Orders page -> Export in
Shopify itself), NOT the analytics-report exports (Monthly POS Report / Sales overview)
the other two tools use. One row per LINE ITEM in this format, with `Subtotal` /
`Financial Status` / `Cancelled at` columns -- Order Value here is that file's own
`Subtotal` column directly: goods only, shipping already excluded, no subtraction
needed.

**What it solves:** the 3 raw order-tracking Google Sheets only ever get a row once an
order is physically WITH the shipping company -- that's just how the ops team works, and
nothing here changes that. But it means Total Orders / Cancelled Orders on the Dashboard
were undercounted: a cancelled order, or one still waiting to ship, never showed up
anywhere.

**How it works:** upload a full "Export orders" file (every status, not just shipped)
for ONE country group. The app:

1. Reads the 3 raw Google Sheets **live** (read-only) AND the staging "Orders" tab that
   already feeds the Dashboard, to see which orders are already logged, anywhere
   (including anything added by the other Orders Status Check app, or the daily sync --
   they all write into the same staging sheet, so nothing ever gets duplicated across
   tools).
2. Classifies every order in the uploaded file as **Cancelled** (via the `Cancelled at`
   column being non-blank -- more complete than Financial Status alone) or **Pending**.
3. Shows only the orders that AREN'T already logged anywhere -- safe to add without ever
   duplicating a row.
4. One click appends those new rows to the bottom of the staging Orders tab directly (or
   download an Excel copy for your own records).

**The 3 raw sheets are never written to by this** -- read-only. Re-run any time you want
fresher numbers; running it twice on the same export is harmless, since anything already
added on a previous run is recognized and skipped automatically.

A bonus fix baked in: Arabic city names in the real sample file came through
mis-encoded (garbled Latin-lookalike text like `Ø§Ù„Ø¹ÙŠÙ†`) -- this app repairs them back
to proper Arabic (`العين`) automatically wherever the encoding round-trips cleanly
(most of them); the rest are left as they came, not guessed further.

## Setup: deploying this as its own app

### Part 1 -- Google Sheets access

This app needs the SAME "robot" service account already set up for the daily sync in the
separate `orders-sync` project -- no new Google account or extra sharing needed, since
that robot already has Viewer access to the 3 raw sheets and Editor access to the
staging sheet. If you don't have that service account's `.json` key file handy anymore,
generate a fresh one for the same account: Google Cloud Console -> IAM & Admin -> Service
Accounts -> (the existing `orders-sync-bot` account) -> Keys tab -> Add Key -> Create new
key -> JSON.

### Part 2 -- Put the code on GitHub

1. Go to https://github.com and sign in (same account used for the other two tools).
2. **+ -> New repository**. Name it e.g. `orders-status-native`. Keep it **Private**
   (recommended). Click **Create repository**.
3. Upload every file from this folder: `app.py`, `logic.py`, `requirements.txt`,
   `README.md` (drag-and-drop or "uploading an existing file"). Commit.

### Part 3 -- Deploy on Streamlit Cloud

1. Go to https://share.streamlit.io and sign in (same account used for the other two
   Streamlit apps, or a new one -- either works).
2. Click **New app** -> pick the `orders-status-native` repository, branch `main`,
   main file path `app.py`. Click **Deploy**.
3. Once it's up, you'll have a NEW, SEPARATE link for this app (its own
   `*.streamlit.app` URL) -- different from the consolidation tool's link.

### Part 4 -- Add the secret

The staging sheet ID is already baked into `logic.py` (Mahmoud's real staging sheet,
confirmed Aug 2026) -- the only secret actually needed is the Google credential.

1. On this app's page (share.streamlit.io), click the **⋮** menu (top-right) ->
   **Settings** -> **Secrets**.
2. Paste in the following:

   ```toml
   gcp_service_account_json = '''
   PASTE_THE_FULL_CONTENT_OF_THE_SERVICE_ACCOUNT_JSON_KEY_FILE_HERE
   '''
   ```

   Open the `.json` key file (Part 1 above) with a text editor, select all, and paste
   its exact content between the `'''` lines, unchanged. **Important:** keep the `'''`
   (three single quotes) exactly as shown, not `"""` (three double quotes) --
   otherwise Streamlit "helpfully" converts the `\n` sequences inside the key text and
   corrupts it.

   (Optional: if the staging sheet ever moves, add
   `staging_spreadsheet_id = "THE_NEW_ID"` above the JSON block to override the
   built-in default -- not needed otherwise.)
3. Click **Save**. The app restarts automatically and should now work with no further
   setup needed. Nothing above is ever visible to anyone but you.

## Files

- **`logic.py`** -- all the cleaning/classification/Google Sheets logic, no Streamlit
  dependency (pure Python + pandas + gspread). Self-contained on purpose -- this app has
  zero dependency on the consolidation tool or the orders-sync project's code, even
  though the ideas and some config (the 3 raw sheet IDs, the column conventions) are
  shared between them.
- **`app.py`** -- the Streamlit UI: upload, column mapping, check, preview, add/download.
