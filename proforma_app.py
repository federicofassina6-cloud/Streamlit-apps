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

# ─────────────────────────────────────────────
# PASSWORD GATE
# ─────────────────────────────────────────────
st.set_page_config(page_title="Proforma Generator", layout="wide")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 Proforma Generator")
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

def get_next_proforma_number():
    year_2digit = date.today().strftime('%y')
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/fatture_proforma",
        headers=HEADERS,
        params={"select": "proforma_number"}
    )
    try:
        existing = response.json()
        if isinstance(existing, list):
            this_year = [r for r in existing if str(r.get("proforma_number", "")).endswith(f"/{year_2digit}")]
            next_num = len(this_year) + 1
        else:
            next_num = 1
    except:
        next_num = 1
    return f"{next_num:03d}/{year_2digit}"

def save_proforma(proforma_number, client_company, total_amount, currency):
    requests.post(
        f"{SUPABASE_URL}/rest/v1/fatture_proforma",
        headers=HEADERS,
        json={
            "proforma_number": proforma_number,
            "client_company": client_company,
            "total_amount": total_amount,
            "currency": currency,
            "status": "not_sent"
        }
    )

def load_products():
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/products",
        headers=HEADERS,
        params={
            "select": "id,description,unit_price_client,unit_price_reseller,category",
            "order": "category.asc,created_at.asc"
        }
    )
    try:
        data = response.json()
        if isinstance(data, list):
            return data
    except:
        pass
    return []

def load_customers():
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/customers",
        headers=HEADERS,
        params={"select": "id,company_name,contact_name,email,phone,address,city,zip,country,notes",
                "order": "company_name.asc"}
    )
    try:
        data = response.json()
        if isinstance(data, list):
            return data
    except:
        pass
    return []

def load_delivery_terms():
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/delivery_terms",
        headers=HEADERS,
        params={"select": "term", "order": "created_at.asc"}
    )
    try:
        data = response.json()
        if isinstance(data, list):
            return [r["term"] for r in data]
    except:
        pass
    return []

def save_delivery_term(term):
    # Only save if it doesn't already exist
    check = requests.get(
        f"{SUPABASE_URL}/rest/v1/delivery_terms",
        headers=HEADERS,
        params={"term": f"eq.{term}", "select": "id"}
    )
    try:
        existing = check.json()
        if isinstance(existing, list) and len(existing) > 0:
            return
    except:
        pass
    requests.post(
        f"{SUPABASE_URL}/rest/v1/delivery_terms",
        headers=HEADERS,
        json={"term": term}
    )

def save_customer(company_name, contact_name, email, phone, address, city, zip_code, country, notes):
    # Check if customer already exists by company name
    check = requests.get(
        f"{SUPABASE_URL}/rest/v1/customers",
        headers=HEADERS,
        params={"company_name": f"eq.{company_name}", "select": "id"}
    )
    try:
        existing = check.json()
        if isinstance(existing, list) and len(existing) > 0:
            return  # already exists, skip
    except:
        pass
    requests.post(
        f"{SUPABASE_URL}/rest/v1/customers",
        headers=HEADERS,
        json={
            "company_name": company_name,
            "contact_name": contact_name,
            "email": email,
            "phone": phone,
            "address": address,
            "city": city,
            "zip": zip_code,
            "country": country,
            "notes": notes
        }
    )

def save_product(description, unit_price_client, unit_price_reseller, category):
    response = requests.post(
        f"{SUPABASE_URL}/rest/v1/products",
        headers=HEADERS,
        json={
            "description": description,
            "unit_price_client": unit_price_client,
            "unit_price_reseller": unit_price_reseller,
            "category": category
        }
    )
    return response.status_code in [200, 201]

# ─────────────────────────────────────────────
# ─────────────────────────────────────────────
# LOAD PRODUCTS FROM SUPABASE
# ─────────────────────────────────────────────
if "products_db" not in st.session_state:
    st.session_state.products_db = load_products()

# Load customers from Supabase (cached per session)
if "customers_db" not in st.session_state:
    st.session_state.customers_db = load_customers()

# Load delivery terms from Supabase (cached per session)
if "delivery_terms_db" not in st.session_state:
    st.session_state.delivery_terms_db = load_delivery_terms()

PRODUCTS = st.session_state.products_db

# Group by category for display
CATEGORIES = []
seen_cats = []
for p in PRODUCTS:
    cat = p.get("category") or "Other"
    if cat not in seen_cats:
        seen_cats.append(cat)
        CATEGORIES.append(cat)

# Build flat dropdown list with category separators
PRODUCT_OPTIONS = []   # list of dicts: {label, product_idx or None}
PRODUCT_NAMES   = ["— custom —"]
PRODUCT_MAP     = {}   # index in PRODUCT_NAMES → product dict

for cat in CATEGORIES:
    cat_products = [p for p in PRODUCTS if (p.get("category") or "Other") == cat]
    PRODUCT_NAMES.append(f"── {cat} ──")   # separator (not selectable)
    for p in cat_products:
        label = p["description"][:65] + ("…" if len(p["description"]) > 65 else "")
        PRODUCT_MAP[len(PRODUCT_NAMES)] = p
        PRODUCT_NAMES.append(label)

# ─────────────────────────────────────────────
# CATALOGUES
# ─────────────────────────────────────────────
CURRENCIES = ["EUR", "USD", "GBP", "CHF", "CNY", "RUB", "— custom —"]

HS_CODES = [
    "8453.9000",
    "8453.1000",
    "8466.9195",
    "8464.2019",
    "8451.9000",
    "8451.8030",
]
PAYMENT_OPTIONS = [
    "In advance by T/t transfer",
    "100% by T/T transfer at the order",
    "50% advance, 50% before shipment",
    "30 days from invoice date",
    "Letter of credit at sight",
]
DELIVERY_TERMS_OPTIONS = st.session_state.delivery_terms_db
DELIVERY_TIME_OPTIONS = [
    "2 weeks from payment receipt",
    "3 - 5 weeks from payment receipt",
    "4 - 6 weeks from payment receipt",
    "6 - 8 weeks from payment receipt",
    "To be confirmed",
]
PACKING_OPTIONS = [
    "Included, for air shipment",
    "Included with fumigated wooden crate, for air shipment",
    "Included with carton box",
    "Not included",
]
SHIPMENT_OPTIONS = [
    "By express courier",
    "By air",
    "By sea",
    "By road",
    "To be arranged by customer",
]

# ─────────────────────────────────────────────
# DOCX HELPERS
# ─────────────────────────────────────────────
def replace_in_paragraph(para, replacements):
    for key, val in replacements.items():
        full_text = "".join(run.text for run in para.runs)
        if key not in full_text:
            continue
        keeper_run = None
        for run in para.runs:
            if key in run.text or (run.text and run.text in key):
                keeper_run = run
                break
        if keeper_run is None:
            for run in para.runs:
                if run.bold:
                    keeper_run = run
                    break
        if keeper_run is None and para.runs:
            keeper_run = para.runs[-1]
        new_text = full_text.replace(key, val)
        if para.runs:
            para.runs[0].text = new_text
            if keeper_run and keeper_run != para.runs[0]:
                para.runs[0].bold = keeper_run.bold
                para.runs[0].italic = keeper_run.italic
                if keeper_run.font.name:
                    para.runs[0].font.name = keeper_run.font.name
                if keeper_run.font.size:
                    para.runs[0].font.size = keeper_run.font.size
            for run in para.runs[1:]:
                run.text = ""

def set_cell_text(cell, text, bold=False, italic=False, font_name="Verdana", font_size=10):
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ""
            rPr = run._r.find(qn('w:rPr'))
            if rPr is not None:
                run._r.remove(rPr)
    para = cell.paragraphs[0]
    run = para.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = font_name
    run.font.size = Pt(font_size)

def set_cell_borders(cell, top_val='nil', bottom_val='nil', left_val='nil', right_val='nil', top_double=False):
    tc = cell._tc
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    existing = tcPr.find(qn('w:tcBorders'))
    if existing is not None:
        tcPr.remove(existing)
    tcBorders = OxmlElement('w:tcBorders')
    for side, val in [('top', top_val), ('left', left_val), ('bottom', bottom_val), ('right', right_val)]:
        b = OxmlElement(f'w:{side}')
        if side == 'top' and top_double:
            b.set(qn('w:val'), 'double')
            b.set(qn('w:sz'), '6')
            b.set(qn('w:space'), '0')
            b.set(qn('w:color'), 'auto')
        else:
            b.set(qn('w:val'), val)
        tcBorders.append(b)
    tcPr.append(tcBorders)

def remove_row_borders(row):
    for cell in row.cells:
        set_cell_borders(cell)

def add_no_wrap(cell):
    tc = cell._tc
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    noWrap = OxmlElement('w:noWrap')
    tcPr.append(noWrap)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "line_items" not in st.session_state:
    st.session_state.line_items = [
        {"product_idx": 0, "description": "", "details": "", "qty": 1.0,
         "unit_price": 0.0, "price_type": "Cliente"}
    ]

def add_line():
    st.session_state.line_items.append(
        {"product_idx": 0, "description": "", "details": "", "qty": 1.0,
         "unit_price": 0.0, "price_type": "Cliente"}
    )

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.title("📄 Proforma Invoice Generator")

# ── 1. DATE & NUMBER ──
st.subheader("1. Date & Number")
col_d1, col_d2 = st.columns(2)
with col_d1:
    selected_date = st.date_input("Date", value=date.today())
with col_d2:
    proforma_number = get_next_proforma_number()
    st.metric("Proforma Number", proforma_number)

year_2digit = selected_date.strftime('%y')
formatted_date = selected_date.strftime('%d/%m/') + "\u2019" + year_2digit

# ── 2. CLIENT ──
st.subheader("2. Client")

# Customer picker
customers = st.session_state.customers_db
customer_names = ["— new customer —"] + [
    f"{c.get('company_name', '')} ({c.get('contact_name', '')})" for c in customers
]
col_cust, col_refresh = st.columns([5, 1])
with col_cust:
    selected_customer_idx = st.selectbox(
        "Pick existing customer or fill in manually below",
        range(len(customer_names)),
        format_func=lambda x: customer_names[x],
        key="customer_picker"
    )
with col_refresh:
    st.write("")
    if st.button("🔄", help="Reload customers from database"):
        st.session_state.customers_db = load_customers()
        st.rerun()

# Auto-fill if customer selected
if selected_customer_idx > 0:
    cust = customers[selected_customer_idx - 1]
    default_salutation = "Mr."
    default_full_name  = cust.get("contact_name", "")
    default_company    = cust.get("company_name", "")
    default_address    = cust.get("address", "")
    default_zip        = cust.get("zip", "")
    default_city       = cust.get("city", "")
    default_region     = ""
    default_country    = cust.get("country", "")
else:
    default_salutation = "Mr."
    default_full_name  = ""
    default_company    = ""
    default_address    = ""
    default_zip        = ""
    default_city       = ""
    default_region     = ""
    default_country    = ""

col1, col2 = st.columns([1, 3])
with col1:
    salutation = st.selectbox("Salutation", ["Mr.", "Ms.", "Dr.", "Messrs."],
                              index=["Mr.", "Ms.", "Dr.", "Messrs."].index(default_salutation))
with col2:
    full_name = st.text_input("Contact Full Name", value=default_full_name,
                              placeholder="e.g. John Smith")

company = st.text_input("Company Name", value=default_company,
                        placeholder="e.g. Vitrex s.r.o.")
address = st.text_input("Address", value=default_address,
                        placeholder="e.g. Zeyerova 1334")

col3, col4, col5 = st.columns(3)
with col3:
    zip_code = st.text_input("Zip", value=default_zip, placeholder="337 01")
with col4:
    city = st.text_input("City", value=default_city, placeholder="Shenzhen")
with col5:
    region = st.text_input("Region", value=default_region, placeholder="(optional)")

country = st.text_input("Country", value=default_country, placeholder="e.g. China")

# ── 3. CURRENCY & PRICE TYPE ──
st.subheader("3. Currency & Price Type")
col_cur, col_pt = st.columns(2)
with col_cur:
    currency_choice = st.selectbox("Currency (ISO)", CURRENCIES)
    if currency_choice == "— custom —":
        currency = st.text_input("Enter ISO currency code", placeholder="e.g. AED, BRL, INR")
    else:
        currency = currency_choice
with col_pt:
    global_price_type = st.radio(
        "Price type (applies to all products)",
        ["Cliente", "Rivenditore"],
        horizontal=True,
        key="global_price_type"
    )
    # When price type changes, update all line item prices
    if st.session_state.get("_last_price_type") != global_price_type:
        st.session_state["_last_price_type"] = global_price_type
        for item in st.session_state.line_items:
            item["price_type"] = global_price_type
            if item.get("product_idx", 0) > 0 and item.get("product_idx") in PRODUCT_MAP:
                pc = item.get("price_client", 0.0)
                pr = item.get("price_reseller", 0.0)
                item["unit_price"] = pc if global_price_type == "Cliente" else pr
        st.rerun()

# ── 4. LINE ITEMS ──
st.subheader("4. Line Items")
st.caption("Select from catalogue or choose '— custom —' to type manually.")

items_to_remove = []
needs_rerun = False
for i, item in enumerate(st.session_state.line_items):
    with st.container():
        c1, c2, c3, c4 = st.columns([3, 1.5, 1.5, 0.4])
        with c1:
            prod_idx = st.selectbox(
                f"Product Name #{i+1} (bold in document)",
                range(len(PRODUCT_NAMES)),
                format_func=lambda x: PRODUCT_NAMES[x],
                key=f"prod_{i}",
                index=item["product_idx"]
            )
            # Skip category separators
            if prod_idx > 0 and PRODUCT_NAMES[prod_idx].startswith("── "):
                prod_idx = item["product_idx"]

            if prod_idx != item["product_idx"]:
                item["product_idx"] = prod_idx
                if prod_idx > 0 and prod_idx in PRODUCT_MAP:
                    p = PRODUCT_MAP[prod_idx]
                    item["description"]    = p["description"]
                    item["price_client"]   = float(p.get("unit_price_client")   or 0)
                    item["price_reseller"] = float(p.get("unit_price_reseller") or 0)
                    new_price = item["price_client"] if global_price_type == "Cliente" else item["price_reseller"]
                    item["unit_price"] = new_price
                    item["price_type"] = global_price_type
                else:
                    item["description"]    = ""
                    item["unit_price"]     = 0.0
                    item["price_client"]   = 0.0
                    item["price_reseller"] = 0.0
                needs_rerun = True

            if prod_idx == 0:
                item["description"] = st.text_input(
                    "Custom Product Name",
                    value=item["description"],
                    key=f"desc_{i}",
                    placeholder="e.g. CHROMED STEEL ROLLER WW1300"
                )

            item["details"] = st.text_input(
                "Description / Specs (optional)",
                value=item.get("details", ""),
                key=f"details_{i}",
                placeholder="e.g. Dimensions (Length) × (Width) × (Height) (±0,1) mm. – blackish color"
            )

        with c2:
            item["qty"] = st.number_input(
                "Qty", min_value=0.0, value=float(item["qty"]),
                step=1.0, format="%.2f", key=f"qty_{i}"
            )
        with c3:
            # Read-only price display — cannot be modified
            st.write(f"**Unit Price ({currency})**")
            st.write(f"{item['unit_price']:.2f}")
        with c4:
            st.write("")
            st.write("")
            if st.button("🗑", key=f"del_{i}", help="Remove line"):
                items_to_remove.append(i)

        line_total = item["qty"] * item["unit_price"]
        st.caption(f"Line total: {currency} {line_total:.2f}")
        st.divider()

for i in sorted(items_to_remove, reverse=True):
    st.session_state.line_items.pop(i)
if items_to_remove or needs_rerun:
    st.rerun()

st.button("➕ Add Line Item", on_click=add_line)

grand_total = sum(item["qty"] * item["unit_price"] for item in st.session_state.line_items)
st.markdown(f"### 💰 Total: {currency} {grand_total:.2f}")

# ── 5. TERMS & CONDITIONS ──
st.subheader("5. Terms & Conditions")
col_t1, col_t2 = st.columns(2)
with col_t1:
    hs_code = st.selectbox("HS Code", HS_CODES + ["— custom —"])
    if hs_code == "— custom —":
        hs_code = st.text_input("Custom HS Code")

    payment = st.selectbox("Payment", PAYMENT_OPTIONS + ["— custom —"])
    if payment == "— custom —":
        payment = st.text_input("Custom Payment Terms")

    delivery_terms = st.selectbox("Delivery Terms", DELIVERY_TERMS_OPTIONS + ["— custom —"])
    if delivery_terms == "— custom —":
        delivery_terms = st.text_input("Custom Delivery Terms", placeholder="e.g. DAP Tokyo (JP)")
        if delivery_terms and delivery_terms not in DELIVERY_TERMS_OPTIONS:
            if st.button("💾 Save this delivery term", key="save_dt"):
                save_delivery_term(delivery_terms)
                st.session_state.delivery_terms_db = load_delivery_terms()
                st.success(f"✅ '{delivery_terms}' saved to database!")
                st.rerun()

    delivery_time = st.selectbox("Delivery Time", DELIVERY_TIME_OPTIONS + ["— custom —"])
    if delivery_time == "— custom —":
        delivery_time = st.text_input("Custom Delivery Time")

with col_t2:
    packing = st.selectbox("Packing", PACKING_OPTIONS + ["— custom —"])
    if packing == "— custom —":
        packing = st.text_input("Custom Packing")

    shipment = st.selectbox("Shipment", SHIPMENT_OPTIONS + ["— custom —"])
    if shipment == "— custom —":
        shipment = st.text_input("Custom Shipment")

# ── 6. DOCUMENT NAME ──
st.subheader("6. Document Name")
default_name = f"proforma {proforma_number.replace('/', '-')} {company}"
doc_name = st.text_input("File name (without .docx)", value=default_name)

# ── GENERATE ──
st.divider()
if st.button("📥 Generate Proforma Invoice", type="primary", use_container_width=True):
    if not company:
        st.warning("Please enter a company name.")
    elif not full_name:
        st.warning("Please enter a contact name.")
    elif not any(item["description"].strip() for item in st.session_state.line_items):
        st.warning("Please add at least one line item.")
    else:
        zip_city = f"{zip_code} {city}".strip()
        if region:
            zip_city += f", {region}"

        header_replacements = {
            f"Schio, [DD/MM/\u2019YY]": f"Schio, {formatted_date}",
            f"[DD/MM/\u2019YY]":        formatted_date,
            "[COMPANY NAME]":           company,
            "[Address]":                address,
            "[Zip] [City], [Region]":   zip_city,
            "[Country]":                country,
            "Mr./Ms. [Full Name]":      f"{salutation} {full_name}",
            "[Sal.]":                   salutation,
            "[Full Name]":              full_name,
            "[NNN/YY]":                 proforma_number,
        }

        try:
            template_path = os.path.join(os.path.dirname(__file__), "proforma_template_eng.docx")
            doc = Document(template_path)
        except Exception as e:
            st.error(f"❌ Template not found: {e}")
            st.stop()

        # Replace header paragraphs
        for para in doc.paragraphs:
            replace_in_paragraph(para, header_replacements)

        # Fix "To the attn. of" paragraph — entire line should NOT be bold
        for para in doc.paragraphs:
            if "To the attn. of" in "".join(r.text for r in para.runs):
                for run in para.runs:
                    run.bold = False

        # Paragraph 0: "Schio, " normal + date bold, Verdana 10
        date_para = doc.paragraphs[0]
        for run in date_para.runs:
            run.text = ""
            rPr = run._r.find(qn('w:rPr'))
            if rPr is not None:
                run._r.remove(rPr)
        date_para.clear()
        r1 = date_para.add_run("Schio, ")
        r1.bold = False
        r1.font.name = "Verdana"
        r1.font.size = Pt(10)
        r2 = date_para.add_run(formatted_date)
        r2.bold = True
        r2.font.name = "Verdana"
        r2.font.size = Pt(10)

        # Apply Verdana 10 to all other header paragraphs
        for para in doc.paragraphs:
            if para == date_para:
                continue
            for run in para.runs:
                run.font.name = "Verdana"
                run.font.size = Pt(10)

        # ── Product table (Table 0) — fill fixed 15 rows ──
        table = doc.tables[0]
        # Row 0 = header, rows 1-15 = data, row 16 = total
        MAX_ROWS = 15
        valid_items = [it for it in st.session_state.line_items if it["description"].strip()]

        for row_idx in range(1, MAX_ROWS + 1):
            row = table.rows[row_idx]
            cells = row.cells

            if row_idx - 1 < len(valid_items):
                item = valid_items[row_idx - 1]
                pos = row_idx * 10
                line_total = item["qty"] * item["unit_price"]
                qty_str   = f"{item['qty']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                price_str = f"{item['unit_price']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                total_str = f"{line_total:.2f},-"

                set_cell_text(cells[0], str(pos),  bold=False, italic=False)

                # Description cell: bold name line 1, details line 2
                desc_cell = cells[1]
                for para in desc_cell.paragraphs:
                    pPr = para._p.find(qn('w:pPr'))
                    if pPr is not None:
                        rPr_in_pPr = pPr.find(qn('w:rPr'))
                        if rPr_in_pPr is not None:
                            pPr.remove(rPr_in_pPr)
                    for run in para.runs:
                        run.text = ""
                        rPr = run._r.find(qn('w:rPr'))
                        if rPr is not None:
                            run._r.remove(rPr)
                first_para = desc_cell.paragraphs[0]
                r = first_para.add_run(item["description"])
                r.bold = True
                r.italic = False
                r.font.name = "Verdana"
                r.font.size = Pt(10)
                details = item.get("details", "").strip()
                if details:
                    new_p = copy.deepcopy(first_para._p)
                    desc_cell._tc.append(new_p)
                    second_para = desc_cell.paragraphs[-1]
                    for run in second_para.runs:
                        run.text = ""
                        rPr = run._r.find(qn('w:rPr'))
                        if rPr is not None:
                            run._r.remove(rPr)
                    dr = second_para.add_run(details)
                    dr.bold = False
                    dr.italic = False
                    dr.font.name = "Verdana"
                    dr.font.size = Pt(10)

                set_cell_text(cells[2], qty_str,   bold=False, italic=False)
                set_cell_text(cells[3], price_str, bold=False, italic=False)
                set_cell_text(cells[4], currency,  bold=False, italic=False)
                set_cell_text(cells[5], total_str, bold=False, italic=False)
            else:
                # Empty row — clear all cells
                for cell in cells:
                    set_cell_text(cell, "", bold=False, italic=False)

        # Total row (row 16) — fill: label | ISO | total
        total_row = table.rows[MAX_ROWS + 1]
        tcells = total_row.cells
        total_str   = f"{grand_total:.2f},-"
        total_label = f"TOTAL PRICE \u2013 {delivery_terms} -"
        # The first cell spans 4 cols — python-docx repeats it 4x in .cells
        # So: tcells[0]=tcells[1]=tcells[2]=tcells[3] = merged label cell
        #     tcells[4] = ISO column
        #     tcells[5] = total price column
        set_cell_text(tcells[0], total_label, bold=True, italic=False)
        set_cell_text(tcells[4], currency,    bold=True, italic=False)
        set_cell_text(tcells[5], total_str,   bold=True, italic=False)

        # Hide empty data rows by collapsing row height to 1 twip
        num_items = len(valid_items)
        for row_idx in range(num_items + 1, MAX_ROWS + 1):
            row = table.rows[row_idx]
            trPr = row._tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = OxmlElement('w:trPr')
                row._tr.insert(0, trPr)
            # Remove existing trHeight
            existing_h = trPr.find(qn('w:trHeight'))
            if existing_h is not None:
                trPr.remove(existing_h)
            trH = OxmlElement('w:trHeight')
            trH.set(qn('w:val'), '1')
            trH.set(qn('w:hRule'), 'exact')
            trPr.append(trH)

        # ── Terms table (Table 1) ──
        terms_table = doc.tables[1]
        terms_map = {
            0: hs_code,
            1: payment,
            4: delivery_terms,
            5: delivery_time,
            6: packing,
            7: shipment,
        }
        for row_idx, value in terms_map.items():
            if row_idx < len(terms_table.rows):
                set_cell_text(terms_table.rows[row_idx].cells[1], value,
                              bold=False, italic=False)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        save_proforma(proforma_number, company, grand_total, currency)

        # Save customer to Supabase if new (not already in DB)
        if company.strip():
            save_customer(
                company_name=company,
                contact_name=full_name,
                email="",
                phone="",
                address=address,
                city=city,
                zip_code=zip_code,
                country=country,
                notes=""
            )
            # Refresh customers cache
            st.session_state.customers_db = load_customers()

        st.success(f"✅ Proforma {proforma_number} ready! Total: {currency} {grand_total:.2f}")
        st.download_button(
            label="📄 Download Word Document",
            data=buffer,
            file_name=f"{doc_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
