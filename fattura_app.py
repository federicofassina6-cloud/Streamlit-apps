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

st.set_page_config(page_title="Fattura Generator", layout="wide")

def fmt_price(n):
    """Format number as European: 2.470,– """
    formatted = f"{n:,.2f}"
    formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    if formatted.endswith(",00"):
        formatted = formatted[:-3] + ",–"
    else:
        formatted = formatted + "–"
    return formatted

# ─────────────────────────────────────────────
# PASSWORD GATE
# ─────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 Fattura Generator")
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

def get_next_invoice_number():
    year_2digit = date.today().strftime('%y')
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/fatture",
        headers=HEADERS,
        params={"select": "invoice_number"}
    )
    try:
        existing = response.json()
        if isinstance(existing, list):
            this_year = [r for r in existing if str(r.get("invoice_number", "")).endswith(f"/{year_2digit}")]
            next_num = len(this_year) + 1
        else:
            next_num = 1
    except:
        next_num = 1
    return f"{next_num:03d}/{year_2digit}"

def save_fattura(invoice_number, client_company, total_amount, currency,
                  address="", zip_code="", city="", region="", country=""):
    requests.post(
        f"{SUPABASE_URL}/rest/v1/fatture",
        headers=HEADERS,
        json={
            "invoice_number": invoice_number,
            "client_company": client_company,
            "total_amount": total_amount,
            "currency": currency,
            "status": "not_sent",
            "address": address,
            "zip": zip_code,
            "city": city,
            "region": region,
            "country": country,
        }
    )

def load_products():
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/products",
        headers=HEADERS,
        params={"select": "id,description,description_eng,unit_price_client,unit_price_reseller,category",
                "order": "category.asc,created_at.asc"}
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
        params={"select": "id,company_name,contact_name,salutation,address,city,zip,country,vat_number",
                "order": "company_name.asc"}
    )
    try:
        data = response.json()
        if isinstance(data, list):
            return data
    except:
        pass
    return []

def load_delivery_addresses():
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/delivery_addresses",
        headers=HEADERS,
        params={"select": "*", "order": "company_name.asc"}
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

def save_delivery_address(company_name, street_address, zip_code, city, country):
    check = requests.get(
        f"{SUPABASE_URL}/rest/v1/delivery_addresses",
        headers=HEADERS,
        params={"company_name": f"eq.{company_name}", "select": "id"}
    )
    try:
        existing = check.json()
        if isinstance(existing, list) and len(existing) > 0:
            return
    except:
        pass
    requests.post(
        f"{SUPABASE_URL}/rest/v1/delivery_addresses",
        headers=HEADERS,
        json={"company_name": company_name, "street_address": street_address,
              "zip_code": zip_code, "city": city, "country": country}
    )

def load_vat_exemptions():
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/vat_exemptions",
        headers=HEADERS,
        params={"select": "code", "order": "created_at.asc"}
    )
    try:
        data = response.json()
        if isinstance(data, list):
            return [r["code"] for r in data if r.get("code")]
    except:
        pass
    return []

def save_vat_exemption(text):
    check = requests.get(
        f"{SUPABASE_URL}/rest/v1/vat_exemptions",
        headers=HEADERS,
        params={"code": f"eq.{text}", "select": "id"}
    )
    try:
        if isinstance(check.json(), list) and len(check.json()) > 0:
            return
    except:
        pass
    requests.post(
        f"{SUPABASE_URL}/rest/v1/vat_exemptions",
        headers=HEADERS,
        json={"code": text}
    )

# ─────────────────────────────────────────────
# DOCX HELPERS
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

def replace_in_table_cell(cell, replacements):
    for para in cell.paragraphs:
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

# ─────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────
if "products_db" not in st.session_state:
    st.session_state.products_db = load_products()
if "customers_db" not in st.session_state:
    st.session_state.customers_db = load_customers()
if "delivery_db" not in st.session_state:
    st.session_state.delivery_db = load_delivery_addresses()
if "delivery_terms_db" not in st.session_state:
    st.session_state.delivery_terms_db = load_delivery_terms()
if "vat_exemptions_db" not in st.session_state:
    st.session_state.vat_exemptions_db = load_vat_exemptions()

PRODUCTS = st.session_state.products_db
CATEGORIES = []
seen_cats = []
for p in PRODUCTS:
    cat = p.get("category") or "Other"
    if cat not in seen_cats:
        seen_cats.append(cat)
        CATEGORIES.append(cat)

PRODUCT_NAMES = ["— custom —"]
PRODUCT_MAP   = {}
for cat in CATEGORIES:
    cat_products = [p for p in PRODUCTS if (p.get("category") or "Other") == cat]
    for p in cat_products:
        eng = p.get("description_eng") or p["description"]
        label = eng[:60] + ("…" if len(eng) > 60 else "")
        PRODUCT_MAP[len(PRODUCT_NAMES)] = p
        PRODUCT_NAMES.append(label)

# ─────────────────────────────────────────────
# OPTIONS
# ─────────────────────────────────────────────
CURRENCIES = ["EUR", "USD", "GBP", "CHF", "CNY", "RUB", "— custom —"]
HS_CODES = ["8453.9000","8453.1000","8466.9195","8464.2019","8451.9000","8451.8030","— custom —"]
PAYMENT_OPTIONS = [
    "In advance by T/t transfer",
    "100% by T/T transfer at the order",
    "50% advance, 50% before shipment",
    "30 days from invoice date",
    "Letter of credit at sight",
    "— custom —"
]
DELIVERY_TIME_OPTIONS = [
    "2 weeks from payment receipt",
    "3 - 5 weeks from payment receipt",
    "4 - 6 weeks from payment receipt",
    "6 - 8 weeks from payment receipt",
    "To be confirmed",
    "— custom —"
]

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "fattura_line_items" not in st.session_state:
    st.session_state.fattura_line_items = [
        {"product_idx": 0, "description": "", "description_it": "", "details": "",
         "qty": 1.0, "unit_price": 0.0, "price_type": "Cliente"}
    ]

def add_line():
    st.session_state.fattura_line_items.append(
        {"product_idx": 0, "description": "", "description_it": "", "details": "",
         "qty": 1.0, "unit_price": 0.0, "price_type": "Cliente"}
    )

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.title("🧾 Fattura Generator")

# ── 1. DATE & NUMBER ──
st.subheader("1. Date & Invoice Number")
col_d1, col_d2 = st.columns(2)
with col_d1:
    selected_date = st.date_input("Date", value=date.today(), format="DD/MM/YYYY")
with col_d2:
    invoice_number = get_next_invoice_number()
    st.metric("Invoice Number", invoice_number)

formatted_date = selected_date.strftime('%d/%m/%Y')

# ── 2. CLIENT ──
st.subheader("2. Client")
customers = st.session_state.customers_db
customer_names = ["— new customer —"] + [
    f"{c.get('company_name', '')} ({c.get('contact_name', '')})" for c in customers
]
col_cust, col_refresh = st.columns([5, 1])
with col_cust:
    selected_customer_idx = st.selectbox(
        "Pick existing customer or fill in manually",
        range(len(customer_names)),
        format_func=lambda x: customer_names[x],
        key="cust_picker"
    )
with col_refresh:
    st.write("")
    if st.button("🔄", help="Reload customers"):
        st.session_state.customers_db = load_customers()
        st.rerun()

if selected_customer_idx > 0:
    cust = customers[selected_customer_idx - 1]
    default_company  = cust.get("company_name", "")
    default_address  = cust.get("address", "")
    default_zip      = cust.get("zip", "")
    default_city     = cust.get("city", "")
    default_country  = cust.get("country", "")
    default_vat      = cust.get("vat_number", "")
else:
    default_company = default_address = default_zip = ""
    default_city = default_country = default_vat = ""

company    = st.text_input("Company Name *", value=default_company)
address    = st.text_input("Address", value=default_address)
col3, col4, col5 = st.columns(3)
with col3:
    zip_code = st.text_input("Zip", value=default_zip)
with col4:
    city = st.text_input("City", value=default_city)
with col5:
    region = st.text_input("Region", placeholder="(optional)")
country    = st.text_input("Country", value=default_country)
vat_number = st.text_input("Tax ID / VAT code / Partita IVA", value=default_vat)

# ── 3. DELIVERY ADDRESS ──
st.subheader("3. Delivery Address")
delivery_addresses = st.session_state.delivery_db
delivery_names = ["— same as billing —", "— new address —"] + [
    f"{d.get('company_name', '')} — {d.get('city', '')}" for d in delivery_addresses
]
col_del, col_del_refresh = st.columns([5, 1])
with col_del:
    selected_delivery_idx = st.selectbox(
        "Select delivery address",
        range(len(delivery_names)),
        format_func=lambda x: delivery_names[x],
        key="delivery_picker"
    )
with col_del_refresh:
    st.write("")
    if st.button("🔄", key="reload_delivery", help="Reload delivery addresses"):
        st.session_state.delivery_db = load_delivery_addresses()
        st.rerun()

if selected_delivery_idx == 0:
    del_company = company
    del_address = address
    del_zip     = zip_code
    del_city    = city
    del_region  = region
    del_country = country
    st.caption("📦 Delivery address same as billing address")
elif selected_delivery_idx == 1:
    del_company = st.text_input("Delivery Company Name")
    del_address = st.text_input("Delivery Street Address")
    col_dz, col_dc = st.columns(2)
    with col_dz:
        del_zip = st.text_input("Delivery ZIP")
    with col_dc:
        del_city = st.text_input("Delivery City")
    del_region  = st.text_input("Delivery Region", placeholder="(optional)")
    del_country = st.text_input("Delivery Country")
    if del_company and st.button("💾 Save this delivery address"):
        save_delivery_address(del_company, del_address, del_zip, del_city, del_country)
        st.session_state.delivery_db = load_delivery_addresses()
        st.success(f"✅ '{del_company}' saved!")
        st.rerun()
else:
    d = delivery_addresses[selected_delivery_idx - 2]
    del_company = d.get("company_name", "")
    del_address = d.get("street_address", "")
    del_zip     = d.get("zip_code", "")
    del_city    = d.get("city", "")
    del_region  = ""
    del_country = d.get("country", "")
    st.caption(f"📦 {del_company} — {del_address}, {del_zip} {del_city}, {del_country}")

# ── 4. TERMS ──
st.subheader("4. Terms")
col_t1, col_t2 = st.columns(2)
with col_t1:
    delivery_terms_options = st.session_state.delivery_terms_db
    delivery_terms = st.selectbox("Delivery Terms", delivery_terms_options + ["— custom —"])
    if delivery_terms == "— custom —":
        delivery_terms = st.text_input("Custom Delivery Terms", placeholder="e.g. DAP Tokyo")

    payment = st.selectbox("Payment Terms", PAYMENT_OPTIONS)
    if payment == "— custom —":
        payment = st.text_input("Custom Payment Terms")

with col_t2:
    hs_code = st.selectbox("HS Code", HS_CODES)
    if hs_code == "— custom —":
        hs_code = st.text_input("Custom HS Code")

    vat_options_dynamic = ["— none —"] + st.session_state.vat_exemptions_db + ["— custom —"]
    vat_exemption_choice = st.selectbox("VAT Exemption", vat_options_dynamic)
    if vat_exemption_choice == "— custom —":
        vat_exemption = st.text_input("Custom VAT exemption text", placeholder="e.g. Art. 8 DPR 633/72")
        if vat_exemption and st.button("💾 Save this VAT exemption", key="save_vat"):
            save_vat_exemption(vat_exemption)
            st.session_state.vat_exemptions_db = load_vat_exemptions()
            st.success("✅ Saved!")
            st.rerun()
    elif vat_exemption_choice == "— none —":
        vat_exemption = ""
    else:
        vat_exemption = vat_exemption_choice

# ── 5. CURRENCY & PRICE TYPE ──
st.subheader("5. Currency & Price Type")
col_cur, col_pt = st.columns(2)
with col_cur:
    currency_choice = st.selectbox("Currency (ISO)", CURRENCIES)
    if currency_choice == "— custom —":
        currency = st.text_input("ISO currency code", placeholder="e.g. AED")
    else:
        currency = currency_choice
with col_pt:
    global_price_type = st.radio(
        "Price type", ["Cliente", "Rivenditore"], horizontal=True, key="fattura_price_type"
    )
    if st.session_state.get("_fattura_last_price_type") != global_price_type:
        st.session_state["_fattura_last_price_type"] = global_price_type
        for item in st.session_state.fattura_line_items:
            item["price_type"] = global_price_type
            if item.get("product_idx", 0) > 0 and item.get("product_idx") in PRODUCT_MAP:
                pc = item.get("price_client", 0.0)
                pr = item.get("price_reseller", 0.0)
                item["unit_price"] = pc if global_price_type == "Cliente" else pr
        st.rerun()

# ── 6. LINE ITEMS ──
st.subheader("6. Line Items")
st.caption("Select from catalogue or choose '— custom —' to type manually.")

items_to_remove = []
needs_rerun = False
for i, item in enumerate(st.session_state.fattura_line_items):
    with st.container():
        c1, c2, c3, c4 = st.columns([3, 1.5, 1.5, 0.4])
        with c1:
            prod_idx = st.selectbox(
                f"Product #{i+1}",
                range(len(PRODUCT_NAMES)),
                format_func=lambda x: PRODUCT_NAMES[x],
                key=f"fattura_prod_{i}",
                index=item["product_idx"]
            )
            if prod_idx != item["product_idx"]:
                item["product_idx"] = prod_idx
                if prod_idx > 0 and prod_idx in PRODUCT_MAP:
                    p = PRODUCT_MAP[prod_idx]
                    item["description"]    = p.get("description_eng") or p["description"]
                    item["description_it"] = p.get("description", "")
                    item["price_client"]   = float(p.get("unit_price_client") or 0)
                    item["price_reseller"] = float(p.get("unit_price_reseller") or 0)
                    item["unit_price"] = item["price_client"] if global_price_type == "Cliente" else item["price_reseller"]
                else:
                    item["description"] = ""
                    item["description_it"] = ""
                    item["unit_price"] = item["price_client"] = item["price_reseller"] = 0.0
                needs_rerun = True

            if prod_idx > 0 and prod_idx in PRODUCT_MAP:
                it_name = PRODUCT_MAP[prod_idx].get("description", "")
                if it_name:
                    st.caption(f"🇮🇹 {it_name}")

            if prod_idx == 0:
                item["description"] = st.text_input(
                    "Custom Product Name (EN)", value=item["description"], key=f"fattura_desc_{i}")
                item["description_it"] = st.text_input(
                    "Custom Product Name (IT)", value=item.get("description_it",""), key=f"fattura_desc_it_{i}")

            item["details"] = st.text_input(
                "Description / Specs (optional)", value=item.get("details", ""), key=f"fattura_details_{i}")

        with c2:
            item["qty"] = st.number_input(
                "Qty", min_value=0.0, value=float(item["qty"]),
                step=1.0, format="%.1f", key=f"fattura_qty_{i}")
        with c3:
            st.write(f"**Unit Price ({currency})**")
            st.write(fmt_price(item["unit_price"]))
        with c4:
            st.write("")
            st.write("")
            if st.button("🗑", key=f"fattura_del_{i}"):
                items_to_remove.append(i)

        line_total = item["qty"] * item["unit_price"]
        st.caption(f"Line total: {currency} {fmt_price(line_total)}")
        st.divider()

for i in sorted(items_to_remove, reverse=True):
    st.session_state.fattura_line_items.pop(i)
if items_to_remove or needs_rerun:
    st.rerun()

st.button("➕ Add Line Item", on_click=add_line)
grand_total = sum(item["qty"] * item["unit_price"] for item in st.session_state.fattura_line_items)
st.markdown(f"### 💰 Total: {currency} {fmt_price(grand_total)}")

# ── 7. DOCUMENT NAME ──
st.subheader("7. Document Name")
default_name = f"fattura {invoice_number.replace('/', '-')} {company}"
doc_name = st.text_input("File name (without .docx)", value=default_name)

# ── GENERATE ──
st.divider()
if st.button("📥 Generate Fattura", type="primary", use_container_width=True):
    if not company:
        st.warning("Please enter a company name.")
    elif not any(item["description"].strip() for item in st.session_state.fattura_line_items):
        st.warning("Please add at least one line item.")
    else:
        zip_city = f"{zip_code} {city}".strip()
        if region:
            zip_city += f", {region}"

        try:
            template_path = os.path.join(os.path.dirname(__file__), "fattura_template.docx")
            doc = Document(template_path)
        except Exception as e:
            st.error(f"❌ Template not found: {e}")
            st.stop()

        header_replacements = {
            "[COMPANY NAME]": company,
            "[Address]":      address,
            "[Zip] [City], [Region]": zip_city,
            "[Country]":      country,
        }
        for para in doc.paragraphs:
            replace_in_paragraph(para, header_replacements)

        # ── Table 0: Invoice details ──
        t0 = doc.tables[0]
        table0_replacements = {
            "[NNN/YY]":                 invoice_number,
            "[DD/MM/YYYY]":             formatted_date,
            "[Partita Iva/VAT Number]": vat_number or "—",
            "[Delivery terms]":         delivery_terms,
        }
        for row in t0.rows:
            for cell in row.cells:
                replace_in_table_cell(cell, table0_replacements)

        # ── Table 1: Payment, bank, delivery ──
        t1 = doc.tables[1]
        set_cell_text(t1.rows[0].cells[0],
                      f"PAYMENT TERMS:\n{payment}",
                      bold=False, font_name="Verdana", font_size=10)
        del_city_region = f"{del_zip} {del_city}".strip()
        if del_region:
            del_city_region += f", {del_region}"
        del_lines = ["DELIVERY PLACE OF THE GOODS:"]
        if del_company: del_lines.append(del_company)
        if del_address: del_lines.append(del_address)
        if del_city_region: del_lines.append(del_city_region)
        if del_country: del_lines.append(del_country)
        set_cell_text(t1.rows[2].cells[0], "\n".join(del_lines),
                      bold=False, font_name="Verdana", font_size=10)

        # ── Table 2: Products ──
        t2 = doc.tables[2]
        MAX_ROWS = 15
        valid_items = [it for it in st.session_state.fattura_line_items if it["description"].strip()]

        for row_idx in range(1, MAX_ROWS + 1):
            row   = t2.rows[row_idx]
            cells = row.cells

            if row_idx - 1 < len(valid_items):
                item       = valid_items[row_idx - 1]
                line_total = item["qty"] * item["unit_price"]
                qty_str    = f"{item['qty']:.1f}"
                price_str  = fmt_price(item["unit_price"])
                total_str  = fmt_price(line_total)

                set_cell_text(cells[0], qty_str, bold=False)

                desc_cell  = cells[1]
                for para in desc_cell.paragraphs:
                    for run in para.runs:
                        run.text = ""
                first_para = desc_cell.paragraphs[0]
                r_en = first_para.add_run(item["description"])
                r_en.bold = True
                r_en.font.name = "Verdana"
                r_en.font.size = Pt(10)
                details = item.get("details", "").strip()
                if details:
                    new_p2 = copy.deepcopy(first_para._p)
                    desc_cell._tc.append(new_p2)
                    det_para = desc_cell.paragraphs[-1]
                    for run in det_para.runs:
                        run.text = ""
                    r_det = det_para.add_run(details)
                    r_det.bold = False
                    r_det.font.name = "Verdana"
                    r_det.font.size = Pt(10)

                set_cell_text(cells[2], currency,  bold=False)
                set_cell_text(cells[3], price_str, bold=False)
                set_cell_text(cells[4], currency,  bold=False)
                set_cell_text(cells[5], total_str, bold=False)
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

        # ── Total row (row 16) ──
        total_row  = t2.rows[16]
        tcells     = total_row.cells
        total_label = f"TOTAL AMOUNT \u2013 {delivery_terms} \u2013"
        if vat_exemption:
            total_label += f"\n\n{vat_exemption}"
        set_cell_text(tcells[1], total_label, bold=True)
        set_cell_text(tcells[4], currency,    bold=True)
        set_cell_text(tcells[5], fmt_price(grand_total), bold=True)

        # ── HS Code row (row 17) ──
        hs_row = t2.rows[17]
        set_cell_text(hs_row.cells[1], f"HS code: {hs_code}", bold=False)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        save_fattura(invoice_number, company, grand_total, currency, address, zip_code, city, region, country)

        st.success(f"✅ Fattura {invoice_number} ready! Total: {currency} {fmt_price(grand_total)}")
        st.download_button(
            label="📄 Download Word Document",
            data=buffer,
            file_name=f"{doc_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
