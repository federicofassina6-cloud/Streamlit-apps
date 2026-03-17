import os
import copy
import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from datetime import date
import io
import requests

st.set_page_config(page_title="Packing List Generator", layout="wide")

def fmt_weight(n):
    """Italian weight format: 1.000,– or 1,25 (2 decimal places)"""
    try:
        f = float(n)
    except (TypeError, ValueError):
        return ""
    cents = round((f % 1) * 100)
    int_str = f"{int(f):,}".replace(",", ".")
    return f"{int_str},–" if cents == 0 else f"{int_str},{cents:02d}"

# ─────────────────────────────────────────────
# PASSWORD GATE
# ─────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 Packing List Generator")
    pwd = st.text_input("Enter passcode to continue:", type="password")
    if st.button("Login"):
        if pwd == "RAINYEAR":
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("❌ Wrong passcode.")
    st.stop()

# ─────────────────────────────────────────────
# SUPABASE
# ─────────────────────────────────────────────
SUPABASE_URL = "https://lztrggttkgvgjouofibd.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imx6dHJnZ3R0a2d2Z2pvdW9maWJkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMyNDAwNzEsImV4cCI6MjA4ODgxNjA3MX0.tbHCQtGW21C2fXCEu2FGwlsXn4kGUWOGoOqjuYyiC7A"
HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json"
}

def load_fatture():
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/fatture",
        headers=HEADERS,
        params={"select": "id,invoice_number,client_company,created_at,date_of_reference,address,zip,city,region,country",
                "order": "created_at.desc"}
    )
    try:
        data = response.json()
        if isinstance(data, list):
            return data
    except:
        pass
    return []

def load_fattura_items(fattura_id):
    """Load line items saved when the fattura was generated."""
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/fattura_items",
        headers=HEADERS,
        params={"fattura_id": f"eq.{fattura_id}",
                "select": "description,description_it,qty,net_weight_kg,dimensions",
                "order": "created_at.asc"}
    )
    try:
        data = response.json()
        if isinstance(data, list):
            return data
    except:
        pass
    return []

def load_pl_numbers():
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/packing_lists",
        headers=HEADERS,
        params={"select": "pl_number"}
    )
    year_2digit = date.today().strftime('%y')
    try:
        d = r.json()
        if isinstance(d, list):
            this_year = [x["pl_number"] for x in d if str(x.get("pl_number", "")).endswith(f"/{year_2digit}")]
            return len(this_year) + 1
    except:
        pass
    return 1

def save_pl_record(pl_number, client_company, date_of_reference=None):
    requests.post(
        f"{SUPABASE_URL}/rest/v1/packing_lists",
        headers={**HEADERS, "Prefer": "return=minimal"},
        json={"pl_number": pl_number, "client_company": client_company,
              "date_of_reference": date_of_reference}
    )

# ─────────────────────────────────────────────
# DOCX HELPERS  (identical to fattura app)
# ─────────────────────────────────────────────
def set_cell_text(cell, text, bold=False, italic=False, font_name="Verdana", font_size=10):
    tc = cell._tc
    paras = tc.findall(qn('w:p'))
    for extra_p in paras[1:]:
        tc.remove(extra_p)
    first_p = cell.paragraphs[0]
    for run in first_p.runs:
        run.text = ""
        rPr = run._r.find(qn('w:rPr'))
        if rPr is not None:
            run._r.remove(rPr)
    lines = text.split("\n")
    run = first_p.add_run(lines[0])
    run.bold = bold
    run.italic = italic
    run.font.name = font_name
    run.font.size = Pt(font_size)
    for line in lines[1:]:
        br = OxmlElement("w:br")
        run._r.addnext(br)
        run2 = first_p.add_run(line)
        run2.bold = bold
        run2.italic = italic
        run2.font.name = font_name
        run2.font.size = Pt(font_size)
        run = run2

def replace_in_paragraph(para, replacements):
    full_text = "".join(run.text for run in para.runs)
    changed = False
    for key, val in replacements.items():
        if key in full_text:
            full_text = full_text.replace(key, val)
            changed = True
    if changed and para.runs:
        para.runs[0].text = full_text
        for run in para.runs[1:]:
            run.text = ""

def delete_para(para):
    p = para._p
    p.getparent().remove(p)

# ─────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────
if "fatture_db" not in st.session_state:
    st.session_state.fatture_db = load_fatture()

# ─────────────────────────────────────────────
# SESSION STATE for gross weight overrides
# ─────────────────────────────────────────────
if "pl_gross_weights" not in st.session_state:
    st.session_state.pl_gross_weights = {}

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.title("📦 Packing List Generator")

# ── 1. PACKING LIST NUMBER ────────────────────
st.subheader("1. Packing List Number")
year_2digit  = date.today().strftime('%y')
next_pl_num  = load_pl_numbers()
suggested_pl = f"{next_pl_num:03d}/{year_2digit}"
pl_number    = st.text_input("Packing List Number (used in filename only)", value=suggested_pl)

# ── 2. LINK TO FATTURA ────────────────────────
st.subheader("2. Link to Fattura")

fatture = st.session_state.fatture_db
col_fat, col_fat_refresh = st.columns([5, 1])
with col_fat:
    if not fatture:
        st.warning("No fatture found in Supabase.")
        st.stop()
    fattura_labels = [
        f"{f['invoice_number']} — {f['client_company']} ({f['created_at'][:10]})"
        for f in fatture
    ]
    sel_fattura_idx = st.selectbox(
        "Select Fattura",
        range(len(fattura_labels)),
        format_func=lambda i: fattura_labels[i],
        key="fattura_picker"
    )
with col_fat_refresh:
    st.write("")
    if st.button("🔄", help="Reload fatture"):
        st.session_state.fatture_db = load_fatture()
        st.session_state.pl_gross_weights = {}
        st.rerun()

sel_fattura    = fatture[sel_fattura_idx]
fattura_id     = sel_fattura.get("id", "")
invoice_number = sel_fattura.get("invoice_number", "")
client_company = sel_fattura.get("client_company", "")
fat_date_raw   = sel_fattura.get("date_of_reference") or ""
try:
    fattura_date = date.fromisoformat(fat_date_raw[:10]).strftime("%d/%m/%Y")
except:
    fattura_date = fat_date_raw[:10]

st.caption(f"📄 Invoice: **{invoice_number}** | Client: **{client_company}** | Date: **{fattura_date}**")

# ── 3. CLIENT (auto from fattura, no picker) ──
st.subheader("3. Client")

company  = client_company
address  = sel_fattura.get("address", "") or ""
zip_code = sel_fattura.get("zip", "") or ""
city     = sel_fattura.get("city", "") or ""
region   = sel_fattura.get("region", "") or ""
country  = sel_fattura.get("country", "") or ""

col1, col2 = st.columns(2)
with col1:
    company  = st.text_input("Company",  value=company)
    address  = st.text_input("Address",  value=address)
    zip_code = st.text_input("ZIP",      value=zip_code)
with col2:
    city    = st.text_input("City",    value=city)
    region  = st.text_input("Region",  value=region,  placeholder="(optional)")
    country = st.text_input("Country", value=country)

include_attn = st.checkbox("Include 'To the attn. of' line?", value=False)
salutation = ""
full_name  = ""
if include_attn:
    col_s, col_n = st.columns([1, 3])
    with col_s:
        salutation = st.selectbox("Salutation", ["Mr.", "Ms.", "Dr.", "Messrs."])
    with col_n:
        full_name = st.text_input("Full Name (optional)")

# ── 4. ALL CONTAINED IN ───────────────────────
st.subheader("4. All contained in:")

if "container_options" not in st.session_state:
    st.session_state.container_options = [
        "One wooden crate [dimensions] cms",
        "One carton box [dimensions] cms",
        "One pallet [dimensions] cms",
    ]

container_choice = st.selectbox(
    "Container type",
    st.session_state.container_options + ["— add new —"]
)

if container_choice == "— add new —":
    new_container = st.text_input("New container description", placeholder="e.g. Two wooden crates [dimensions] cms")
    if new_container and st.button("➕ Add to list"):
        st.session_state.container_options.append(new_container)
        st.rerun()
    container_choice = new_container or ""

# If the chosen option contains [dimensions], show a dimensions input
crate_dimensions = ""
if "[dimensions]" in container_choice:
    crate_dimensions = st.text_input("Dimensions (cm)", value="", placeholder="e.g. 120 x 80 x 90")
    container_line = container_choice.replace("[dimensions]", crate_dimensions.strip()) if crate_dimensions.strip() else container_choice
else:
    container_line = container_choice

# ── 5. LINE ITEMS (from fattura) ──────────────
st.subheader("5. Line Items")

fattura_items = load_fattura_items(fattura_id)

if not fattura_items:
    st.warning("⚠️ No line items found for this fattura. Make sure you generated the fattura with the updated app so items are saved to Supabase.")
    valid_items = []
else:
    st.caption(f"✅ {len(fattura_items)} item(s) loaded from fattura {invoice_number}")

    # Show items with editable gross weight only
    valid_items = []
    for i, item in enumerate(fattura_items):
        desc    = item.get("description", "")
        desc_it = item.get("description_it", "")
        qty     = float(item.get("qty") or 0)
        nw      = float(item.get("net_weight_kg") or 0)
        dims    = item.get("dimensions") or ""

        # Gross weight: default = net weight, user can override
        gw_key = f"gw_{fattura_id}_{i}"
        if gw_key not in st.session_state.pl_gross_weights:
            st.session_state.pl_gross_weights[gw_key] = nw

        with st.container():
            c1, c2, c3, c4 = st.columns([3, 1, 2, 2])
            with c1:
                st.write(f"**{desc}**")
                if desc_it:
                    st.caption(f"🇮🇹 {desc_it}")
                if dims:
                    st.caption(f"📐 {dims}")
            with c2:
                st.write("**Qty**")
                st.write(f"{int(qty):,},0" if qty == int(qty) else str(qty))
            with c3:
                st.write("**Net Weight (kg)**")
                st.write(fmt_weight(nw) + " kg" if nw else "—")
            with c4:
                gross = st.number_input(
                    "Gross Weight (kg)", min_value=0.0,
                    value=float(st.session_state.pl_gross_weights[gw_key]),
                    step=0.01, format="%.2f", key=f"pl_gw_{i}"
                )
                st.session_state.pl_gross_weights[gw_key] = gross

            line_net   = qty * nw
            line_gross = qty * gross
            st.caption(f"Line net: {fmt_weight(line_net)} kg  |  Line gross: {fmt_weight(line_gross)} kg")
            st.divider()

            valid_items.append({
                "description": desc,
                "description_it": desc_it,
                "qty":         qty,
                "net_weight":  nw,
                "gross_weight": gross,
                "dimensions":  dims,
            })

total_net   = sum(it["qty"] * it["net_weight"]   for it in valid_items)
total_gross = sum(it["qty"] * it["gross_weight"] for it in valid_items)

col_nw, col_gw = st.columns(2)
with col_nw:
    st.markdown(f"### ⚖️ Total Net: {fmt_weight(total_net)} kg")
with col_gw:
    st.markdown(f"### ⚖️ Total Gross: {fmt_weight(total_gross)} kg")

# ── 6. DOCUMENT NAME ──────────────────────────
st.subheader("6. Document Name")
default_name = f"PackingList {pl_number.replace('/', '-')} {company}"
doc_name = st.text_input("File name (without .docx)", value=default_name)

# ── GENERATE ──────────────────────────────────
st.divider()
if st.button("📥 Generate Packing List", type="primary", use_container_width=True):
    if not company:
        st.warning("Please enter a company name.")
    elif not valid_items:
        st.warning("Please select a fattura with line items.")
    else:
        zip_city = f"{zip_code} {city}".strip()
        if region:
            zip_city += f", {region}"

        try:
            template_path = os.path.join(os.path.dirname(__file__), "packing_list_template.docx")
            doc = Document(template_path)
        except Exception as e:
            st.error(f"❌ Template not found: {e}")
            st.stop()

        # ── Header paragraphs ──
        header_replacements = {
            "[COMPANY NAME]": company.upper(),
            "[Address]":      address,
            "[Zip] [City], [Region]": zip_city,
            "[Country]":      country,
        }
        for para in doc.paragraphs:
            replace_in_paragraph(para, header_replacements)

        # Bold only company, everything else not bold
        for para in doc.paragraphs:
            full = "".join(r.text for r in para.runs)
            if company.upper() in full and full.strip() == company.upper():
                for run in para.runs:
                    run.bold = True
                    run.font.name = "Verdana"
                    run.font.size = Pt(10)
            elif full.strip() in ["", "Messrs.", "PACKING LIST",
                                   "Covering the shipment of:",
                                   "GOODS OF ITALIAN ORIGIN",
                                   "All contained in:"]:
                pass
            else:
                for run in para.runs:
                    if run.text.strip():
                        run.bold = False
                        run.font.name = "Verdana"
                        run.font.size = Pt(10)

        # Attn line — delete if not needed
        for para in doc.paragraphs:
            if "To the attn. of" in para.text:
                if include_attn and (salutation or full_name):
                    attn_text = f"To the attn. of {salutation} {full_name}".strip().replace("  ", " ")
                    replace_in_paragraph(para, {"To the attn. of [Sal.] [Full Name]": attn_text})
                    for run in para.runs:
                        run.bold = False
                        run.font.name = "Verdana"
                        run.font.size = Pt(10)
                else:
                    delete_para(para)
                break

        # Invoice ref, dimensions, weights
        other_replacements = {
            "[NNN/YY]":            invoice_number,
            "[DD/MM/YYYY]":        fattura_date,
            "[dimensions]":        crate_dimensions.strip() if crate_dimensions.strip() else "[dimensions]",
            "[sum of Net Weight]": fmt_weight(total_net),
        }
        for para in doc.paragraphs:
            replace_in_paragraph(para, other_replacements)
        # Replace the "All contained in:" line
        for para in doc.paragraphs:
            if "One wooden crate" in para.text or "[dimensions]" in para.text:
                replace_in_paragraph(para, {para.text.strip(): container_line})

        # ── Product table ──
        table    = doc.tables[0]
        MAX_ROWS = 15

        for row_idx in range(1, MAX_ROWS + 1):
            row      = table.rows[row_idx]
            cells    = row.cells
            item_idx = row_idx - 1

            if item_idx < len(valid_items):
                item = valid_items[item_idx]

                # Description cell: product name bold, dimensions below if available
                desc_cell  = cells[1]
                for para in desc_cell.paragraphs:
                    for run in para.runs:
                        run.text = ""
                first_para = desc_cell.paragraphs[0]
                r_name = first_para.add_run(item["description"])
                r_name.bold      = True
                r_name.font.name = "Verdana"
                r_name.font.size = Pt(10)
                if item.get("dimensions"):
                    new_p = copy.deepcopy(first_para._p)
                    desc_cell._tc.append(new_p)
                    dim_para = desc_cell.paragraphs[-1]
                    for run in dim_para.runs:
                        run.text = ""
                    r_dim = dim_para.add_run(item["dimensions"])
                    r_dim.bold      = False
                    r_dim.font.name = "Verdana"
                    r_dim.font.size = Pt(10)

                qty_val = item["qty"]
                qty_str = f"{int(qty_val)},0" if qty_val == int(qty_val) else f"{qty_val:.1f}".replace(".", ",")

                set_cell_text(cells[0], qty_str)
                set_cell_text(cells[2], "Kg")
                set_cell_text(cells[3], fmt_weight(item["net_weight"]))
                set_cell_text(cells[4], "Kg")
                set_cell_text(cells[5], fmt_weight(item["gross_weight"]))
            else:
                for cell in cells:
                    set_cell_text(cell, "")
                trPr = row._tr.find(qn('w:trPr'))
                if trPr is None:
                    trPr = OxmlElement('w:trPr')
                    row._tr.insert(0, trPr)
                existing_h = trPr.find(qn('w:trHeight'))
                if existing_h is not None:
                    trPr.remove(existing_h)
                trH = OxmlElement('w:trHeight')
                trH.set(qn('w:val'), '1')
                trH.set(qn('w:hRule'), 'exact')
                trPr.append(trH)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        save_pl_record(pl_number, company, date_of_reference=fat_date_raw[:10] if fat_date_raw else None)

        st.success(f"✅ Packing List {pl_number} ready!")
        st.download_button(
            label="📄 Download Word Document",
            data=buffer,
            file_name=f"{doc_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
