import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from datetime import date
import io
import requests
import copy

st.set_page_config(page_title="Offerta Generator", layout="wide")

SUPABASE_URL = "https://lztrggttkgvgjouofibd.supabase.co"
SUPABASE_KEY = "sb_publishable_2kCkVA7G9VdPWiIXBGIFPw_O_0gfReQ"
HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
}

# ─── Italian price formatting ─────────────────────────────────────────────────
def format_price_it(value):
    """Format as Italian price: 1.000,– or 1.000,50"""
    try:
        f = float(value)
    except (ValueError, TypeError):
        return ""
    cents = round((f % 1) * 100)
    int_str = f"{int(f):,}".replace(",", ".")
    return f"{int_str},–" if cents == 0 else f"{int_str},{cents:02d}"


def parse_price_it(s):
    """Parse Italian-formatted price string back to float"""
    try:
        s = s.strip().rstrip("–").replace(".", "").replace(",", ".")
        return float(s)
    except Exception:
        return 0.0


# ─── Supabase helpers ─────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_contacts():
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/contacts?select=*&order=company.asc",
        headers=HEADERS,
    )
    return r.json() if r.ok else []


@st.cache_data(ttl=300)
def load_items():
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/items?select=*&order=name.asc",
        headers=HEADERS,
    )
    return r.json() if r.ok else []


@st.cache_data(ttl=300)
def load_offerta_numbers():
    yr = date.today().strftime("%y")
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/offerte?select=number&year=eq.{yr}&order=number.asc",
        headers=HEADERS,
    )
    return [row["number"] for row in r.json()] if r.ok else []


def save_offerta_record(number, year, company, doc_date):
    data = {"number": number, "year": year, "company": company, "date": doc_date}
    requests.post(
        f"{SUPABASE_URL}/rest/v1/offerte",
        headers={**HEADERS, "Prefer": "return=minimal"},
        json=data,
    )
    load_offerta_numbers.clear()


# ─── Docx helpers ────────────────────────────────────────────────────────────
def replace_run_text(para, old, new):
    """Replace placeholder text across runs in a paragraph."""
    full = "".join(r.text for r in para.runs)
    if old not in full:
        return False
    full = full.replace(old, new)
    if para.runs:
        para.runs[0].text = full
        for r in para.runs[1:]:
            r.text = ""
    return True


def set_para_bold(para, bold=True):
    for run in para.runs:
        if run.text.strip():
            run.bold = bold


def collapse_paragraph(para):
    """Make a paragraph invisible by collapsing its spacing and font size."""
    pPr = para._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "120")
    spacing.set(qn("w:lineRule"), "exact")
    existing = pPr.find(qn("w:spacing"))
    if existing is not None:
        pPr.remove(existing)
    pPr.append(spacing)
    for run in para.runs:
        run.font.size = Pt(1)
        run.font.color.rgb = None


def replace_cell_text(cell, old, new):
    """Replace placeholder in a table cell."""
    for para in cell.paragraphs:
        full = "".join(r.text for r in para.runs)
        if old in full:
            full = full.replace(old, new)
            if para.runs:
                para.runs[0].text = full
                for r in para.runs[1:]:
                    r.text = ""
            else:
                para.add_run(full)


def clear_product_row(row):
    """Clear all cell content in an unused product row."""
    for cell in row.cells:
        for para in cell.paragraphs:
            for run in para.runs:
                run.text = ""


# ─── Document generation ─────────────────────────────────────────────────────
def generate_offerta(
    lang,
    offer_number,
    doc_date,
    company,
    address,
    zip_code,
    city,
    region,
    country,
    include_attn,
    salutation,
    full_name,
    our_ref,
    products,
    delivery_terms,
    currency,
    hs_code,
    payment,
    delivery_time,
    packing,
    shipment,
    notes,
):
    template_path = (
        "offerta_template_eng.docx" if lang == "ENG" else "offerta_template_ita.docx"
    )
    doc = Document(template_path)
    paras = doc.paragraphs

    # Para 0: Date
    replace_run_text(paras[0], "[DD/MM/'YY]", doc_date)

    # Para 2: Company (bold)
    replace_run_text(paras[2], "[COMPANY NAME]", company.upper())
    set_para_bold(paras[2], True)

    # Para 3: Address (not bold)
    replace_run_text(paras[3], "[Address]", address)
    set_para_bold(paras[3], False)

    # Para 4: Zip City, Region (not bold)
    zip_city = f"{zip_code} {city}".strip()
    if region:
        zip_city += f", {region}"
    replace_run_text(paras[4], "[Zip] [City], [Region]", zip_city)
    set_para_bold(paras[4], False)

    # Para 5: Country (not bold)
    replace_run_text(paras[5], "[Country]", country)
    set_para_bold(paras[5], False)

    # Para 7: To the attn. of (optional)
    attn_para = paras[7]
    if include_attn and (salutation or full_name):
        attn_text = f"To the attn. of {salutation} {full_name}".strip().replace("To the attn. of  ", "To the attn. of ")
        replace_run_text(attn_para, "To the attn. of [Sal.] [Full Name]", attn_text)
        set_para_bold(attn_para, False)
    else:
        collapse_paragraph(attn_para)

    # Para 9: Offer number (bold)
    yr = doc_date.split("/")[-1][-2:] if "/" in doc_date else date.today().strftime("%y")
    num_str = f"{int(offer_number):03d}/{yr}"
    if lang == "ENG":
        full_line = f"OFFER NO...: {num_str}"
    else:
        full_line = f"OFFERTA Nr.: {num_str}"
    for run in paras[9].runs:
        run.text = ""
    if paras[9].runs:
        paras[9].runs[0].text = full_line
        paras[9].runs[0].bold = True
    else:
        paras[9].add_run(full_line).bold = True

    # Para 11: Our ref (ENG only)
    if lang == "ENG" and our_ref:
        replace_run_text(paras[11], "Description", our_ref)

    # Notes paragraph
    notes_placeholder = (
        "Notes and comments (ex. VAT excluded)" if lang == "ENG"
        else "Note e commenti (ex. IVA esclusa)"
    )
    for para in doc.paragraphs:
        if notes_placeholder in "".join(r.text for r in para.runs):
            replace_run_text(para, notes_placeholder, notes if notes else notes_placeholder)
            break

    # ── Product table ──
    prod_table = doc.tables[0]
    for idx in range(15):
        row = prod_table.rows[idx + 1]
        cells = row.cells
        if idx < len(products):
            p = products[idx]
            pos = str((idx + 1) * 10)
            desc = p.get("name", "")
            if p.get("description"):
                desc += f" {p['description']}"
            replace_cell_text(cells[0], cells[0].text, pos)
            replace_cell_text(cells[1], cells[1].text, desc)
            replace_cell_text(cells[2], cells[2].text, str(p.get("qty", "")))
            replace_cell_text(cells[3], cells[3].text, format_price_it(p.get("unit_price", 0)))
            replace_cell_text(cells[4], cells[4].text, p.get("currency", currency))
            replace_cell_text(cells[5], cells[5].text, format_price_it(p.get("total_price", 0)))
        else:
            clear_product_row(row)

    # Total row
    total_row = prod_table.rows[16]
    total_cells = total_row.cells
    total_sum = sum(p.get("total_price", 0) for p in products)
    total_label = f"TOTAL PRICE – {delivery_terms} –" if delivery_terms else "TOTAL PRICE –"
    replace_cell_text(total_cells[0], total_cells[0].text, total_label)
    replace_cell_text(total_cells[4], total_cells[4].text, currency)
    replace_cell_text(total_cells[5], total_cells[5].text, format_price_it(total_sum))

    # ── Terms table ──
    terms_table = doc.tables[1]
    terms_map = {
        "[HS code]": hs_code,
        "[Payment]": payment,
        "[Delivery terms]": delivery_terms,
        "[Delivery time]": delivery_time,
        "[Packing]": packing,
        "[Shipments]": shipment,
    }
    for row in terms_table.rows:
        for cell in row.cells:
            for placeholder, value in terms_map.items():
                if placeholder in cell.text:
                    replace_cell_text(cell, placeholder, value or "")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ─── Main App ────────────────────────────────────────────────────────────────
st.title("📄 Offerta Generator")

# ── Load data ──
contacts = load_contacts()
items = load_items()
existing_numbers = load_offerta_numbers()
current_year = date.today().strftime("%y")
next_num = (max(existing_numbers) + 1) if existing_numbers else 1

# ── Sidebar ──
with st.sidebar:
    st.header("Impostazioni")
    lang = st.selectbox("🌐 Lingua / Language", ["ENG", "ITA"])

    st.markdown("---")
    st.subheader("Numero Offerta")
    offer_num = st.number_input(
        f"N. Offerta /{current_year}",
        min_value=1, max_value=999, value=next_num, step=1,
    )

    num_ok = True
    if offer_num in existing_numbers:
        st.error(f"⛔ Offerta {offer_num:03d}/{current_year} esiste già!")
        num_ok = False
    elif offer_num > next_num:
        st.warning(f"⚠️ Salto nella numerazione. Prossimo atteso: {next_num:03d}")

    doc_date = st.date_input("Data", value=date.today(), format="DD/MM/YYYY")

    st.markdown("---")
    st.subheader("Valuta")
    currency = st.selectbox("Valuta", ["EUR", "USD", "GBP", "CHF"])

# ── Contact selection ──
st.subheader("👤 Cliente")

company_names = sorted(set(c.get("company", "") for c in contacts if c.get("company")))

col1, col2 = st.columns([2, 1])
with col1:
    selected_company = st.selectbox("Azienda *", options=["— Seleziona —"] + company_names)

contact_data = {}
if selected_company and selected_company != "— Seleziona —":
    for c in contacts:
        if c.get("company") == selected_company:
            contact_data = c
            break

with st.expander("📋 Dati cliente", expanded=True):
    c1, c2 = st.columns(2)
    with c1:
        address  = st.text_input("Indirizzo", value=contact_data.get("address", ""))
        zip_code = st.text_input("CAP",       value=contact_data.get("zip_code", ""))
        city     = st.text_input("Città",     value=contact_data.get("city", ""))
    with c2:
        region  = st.text_input("Regione/Stato", value=contact_data.get("region", ""))
        country = st.text_input("Paese",         value=contact_data.get("country", ""))

    include_attn = st.checkbox("📌 Includi 'To the attn. of'")
    salutation = ""
    full_name  = ""
    if include_attn:
        a1, a2 = st.columns([1, 3])
        with a1:
            salutation = st.text_input("Titolo (Mr./Ms./Dr.)", value="Mr.")
        with a2:
            full_name = st.text_input("Nome completo", value=contact_data.get("full_name", ""))

our_ref = ""
if lang == "ENG":
    our_ref = st.text_input("Our ref. / Description")

# ── Products ──
st.subheader("📦 Prodotti")

item_options = {it.get("name", ""): it for it in items if it.get("name")}
num_products = st.number_input("Numero righe prodotto", min_value=1, max_value=15, value=1)

products = []
for i in range(int(num_products)):
    st.markdown(f"**Riga {(i+1)*10}**")
    p_col1, p_col2, p_col3, p_col4 = st.columns([3, 1, 2, 2])

    with p_col1:
        item_names = ["— Seleziona prodotto —"] + sorted(item_options.keys())
        sel = st.selectbox("Prodotto", item_names, key=f"prod_{i}")

    item_data    = item_options.get(sel, {}) if sel != "— Seleziona prodotto —" else {}
    default_desc  = item_data.get("description", "")
    default_price = float(item_data.get("unit_price", 0.0) or 0.0)

    with p_col2:
        qty = st.number_input("Q.tà", min_value=0.0, value=1.0, step=1.0, key=f"qty_{i}")
    with p_col3:
        unit_price_str = st.text_input("P. Unità", value=format_price_it(default_price), key=f"uprice_{i}", placeholder="es. 1.000,–")
    unit_price  = parse_price_it(unit_price_str) if unit_price_str else 0.0
    total_price = round(unit_price * qty, 2)
    with p_col4:
        st.text_input("P. Totale", value=format_price_it(total_price), key=f"tprice_{i}", disabled=True)

    desc_input = st.text_input("Descrizione aggiuntiva (opzionale)", value=default_desc, key=f"desc_{i}")

    if sel and sel != "— Seleziona prodotto —":
        products.append({
            "name": sel,
            "description": desc_input,
            "qty": qty,
            "unit_price": unit_price,
            "total_price": total_price,
            "currency": currency,
        })

# ── Terms ──
st.subheader("📋 Condizioni")
t1, t2 = st.columns(2)
with t1:
    delivery_terms = st.text_input("Resa / Delivery terms", placeholder="es. EXW Schio")
    payment        = st.text_input("Pagamento / Payment",   placeholder="es. 30 gg d.f.f.m.")
    hs_code        = st.text_input("HS Code")
with t2:
    delivery_time = st.text_input("Consegna / Delivery time", placeholder="es. 8-10 weeks")
    packing       = st.text_input("Imballo / Packing",        placeholder="es. Export packing")
    shipment      = st.text_input("Spedizione / Shipment",    placeholder="es. By sea")

notes = st.text_area(
    "Note / Notes",
    placeholder="Note aggiuntive (es. VAT excluded)" if lang == "ENG" else "Note aggiuntive (es. IVA esclusa)",
    height=80,
)

# ── Generate ──
st.markdown("---")
company_ok  = selected_company and selected_company != "— Seleziona —"
products_ok = len(products) > 0

if not company_ok:
    st.warning("⚠️ Seleziona un'azienda cliente")
if not products_ok:
    st.warning("⚠️ Aggiungi almeno un prodotto")

if st.button("📄 Genera Offerta", disabled=not (num_ok and company_ok and products_ok), type="primary", use_container_width=True):
    with st.spinner("Generazione documento..."):
        date_str = doc_date.strftime("%d/%m/%y")
        buf = generate_offerta(
            lang=lang,
            offer_number=offer_num,
            doc_date=date_str,
            company=selected_company,
            address=address,
            zip_code=zip_code,
            city=city,
            region=region,
            country=country,
            include_attn=include_attn,
            salutation=salutation,
            full_name=full_name,
            our_ref=our_ref,
            products=products,
            delivery_terms=delivery_terms,
            currency=currency,
            hs_code=hs_code,
            payment=payment,
            delivery_time=delivery_time,
            packing=packing,
            shipment=shipment,
            notes=notes,
        )
        save_offerta_record(number=offer_num, year=current_year, company=selected_company, doc_date=date_str)

        filename = f"Offerta_{offer_num:03d}_{current_year}_{selected_company.replace(' ', '_')}.docx"
        st.success(f"✅ Offerta {offer_num:03d}/{current_year} generata!")
        st.download_button(
            label="⬇️ Scarica Offerta",
            data=buf,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
