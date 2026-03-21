import streamlit as st
import requests
import json
import csv
import io
import zipfile
from datetime import datetime

st.set_page_config(page_title="Supabase Backup", layout="wide")

# ─────────────────────────────────────────────
# PASSWORD GATE
# ─────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 Supabase Backup")
    pwd = st.text_input("Enter passcode to continue:", type="password")
    if st.button("Login"):
        if pwd == "RAINYEAR":
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("❌ Wrong passcode.")
    st.stop()

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
SUPABASE_URL = "https://lztrggttkgvgjouofibd.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imx6dHJnZ3R0a2d2Z2pvdW9maWJkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMyNDAwNzEsImV4cCI6MjA4ODgxNjA3MX0.tbHCQtGW21C2fXCEu2FGwlsXn4kGUWOGoOqjuYyiC7A"
HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
}

TABLES = [
    "categories",
    "customers",
    "delivery_addresses",
    "delivery_terms",
    "fattura_items",
    "fatture",
    "fatture_proforma",
    "offerte",
    "packing_lists",
    "payment_terms",
    "products",
    "transactions",
    "vat_codes",
    "vat_exemptions",
]

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def fetch_table(table: str) -> list:
    """Fetch all rows from a table, handling pagination."""
    all_rows = []
    limit = 1000
    offset = 0
    while True:
        r = requests.get(
            f"{SUPABASE_URL}/rest/v1/{table}",
            headers={**HEADERS, "Range-Unit": "items", "Range": f"{offset}-{offset + limit - 1}"},
            params={"select": "*"},
        )
        if not r.ok:
            return None, f"HTTP {r.status_code}: {r.text}"
        batch = r.json()
        if not isinstance(batch, list):
            return None, str(batch)
        all_rows.extend(batch)
        if len(batch) < limit:
            break
        offset += limit
    return all_rows, None

def rows_to_csv(rows: list) -> str:
    if not rows:
        return ""
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=rows[0].keys())
    writer.writeheader()
    writer.writerows(rows)
    return output.getvalue()

def rows_to_json(rows: list) -> str:
    return json.dumps(rows, indent=2, ensure_ascii=False, default=str)

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.title("🗄️ Supabase Backup")
st.caption(f"Database: `{SUPABASE_URL}`")

st.divider()

col_fmt, col_sel = st.columns([1, 2])
with col_fmt:
    fmt = st.radio("Export format", ["JSON", "CSV"], horizontal=True)
with col_sel:
    selected_tables = st.multiselect(
        "Tables to back up",
        options=TABLES,
        default=TABLES,
    )

st.divider()

if st.button("🚀 Run Backup", type="primary", use_container_width=True):
    if not selected_tables:
        st.warning("Please select at least one table.")
        st.stop()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_buffer = io.BytesIO()
    results = []

    progress = st.progress(0, text="Starting backup…")

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, table in enumerate(selected_tables):
            progress.progress((idx) / len(selected_tables), text=f"Fetching `{table}`…")

            rows, error = fetch_table(table)

            if error:
                results.append({"table": table, "status": "❌ Error", "rows": 0, "detail": error})
                continue

            ext = fmt.lower()
            content = rows_to_json(rows) if fmt == "JSON" else rows_to_csv(rows)
            filename = f"{timestamp}_{table}.{ext}"
            zf.writestr(filename, content)
            results.append({"table": table, "status": "✅ OK", "rows": len(rows), "detail": ""})

    progress.progress(1.0, text="Done!")

    # Summary table
    st.subheader("Backup Summary")
    total_rows = sum(r["rows"] for r in results)
    errors = [r for r in results if "Error" in r["status"]]

    col1, col2, col3 = st.columns(3)
    col1.metric("Tables backed up", f"{len(selected_tables) - len(errors)} / {len(selected_tables)}")
    col2.metric("Total rows", f"{total_rows:,}")
    col3.metric("Errors", len(errors))

    for r in results:
        if r["detail"]:
            st.error(f"`{r['table']}`: {r['detail']}")
        else:
            st.write(f"{r['status']} **{r['table']}** — {r['rows']:,} rows")

    # Download button
    zip_name = f"supabase_backup_{timestamp}.zip"
    zip_buffer.seek(0)
    st.divider()
    st.download_button(
        label=f"📦 Download Backup ({zip_name})",
        data=zip_buffer,
        file_name=zip_name,
        mime="application/zip",
        use_container_width=True,
    )
