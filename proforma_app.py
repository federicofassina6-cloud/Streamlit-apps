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

# ─────────────────────────────────────────────
# CATALOGUES
# ─────────────────────────────────────────────
PRODUCTS = [
    {"description": "CHROMED STEEL ENGRAVED ROLLER WW1300 with positive pattern for reverse rotation. Engraving 20C", "unit_price": 1750.0},
    {"description": "HARD RUBBER POSITIVE SLEEVE WW 300 for reverse rotation. Engraving type: 20C", "unit_price": 900.0},
    {"description": "HARD RUBBER POSITIVE SLEEVE WW 1300 for reverse rotation. Engraving type: 20C", "unit_price": 2150.0},
    {"description": "HARD RUBBER POSITIVE SLEEVE WW 2600 for reverse rotation. Engraving type: 20C", "unit_price": 3700.0},
    {"description": "HARD RUBBER POSITIVE SLEEVE WW 2600 for reverse rotation. Engraving type: 30CC for etching", "unit_price": 3700.0},
    {"description": "AIR MANDREL WW 300", "unit_price": 900.0},
    {"description": "AIR MANDREL WW 1300", "unit_price": 2500.0},
    {"description": "AIR MANDREL WW 2600", "unit_price": 4500.0},
    {"description": "STEEL DOCTOR BLADE WW300, dimensions 439x55x1,5 mm. with 30° bevelling", "unit_price": 48.0},
    {"description": "COMPOSITE MATERIAL (PLASTIC) DOCTOR BLADE WW1300, dimensions 1.384x57x0,9 mm.", "unit_price": 67.0},
    {"description": "COMPOSITE MATERIAL (PLASTIC) DOCTOR BLADE WW2600, dimensions 2.684x57x0,9 mm.", "unit_price": 110.0},
    {"description": "Complete doctor blade holder WW 1300", "unit_price": 1300.0},
    {"description": "Complete doctor blade holder WW 2600", "unit_price": 2100.0},
    {"description": "SIDE SEAL (green plastic dam)", "unit_price": 105.0},
    {"description": "FRONT SIDE SEAL (white plastic dam)", "unit_price": 34.0},
    {"description": "RH SIDE SEAL (white plastic dam)", "unit_price": 31.0},
    {"description": "LH SIDE SEAL (white plastic dam)", "unit_price": 31.0},
    {"description": "Split pins float valve 510/2 heavy, brass seat diam. 5 mm rod length 200 mm with 1/4 W thread (FARG)", "unit_price": 5.5},
    {"description": "Float valve in plastic with ball (FARG)", "unit_price": 0.5},
    {"description": "Motovario gearbox NMRV-P 063 7.5:1 PAM 120/19 slow shaft D25", "unit_price": 430.0},
    {"description": "Packing charges", "unit_price": 450.0},
    {"description": "CIF SZX airport charges", "unit_price": 1200.0},
    {"description": "Packing and DAP charges", "unit_price": 130.0},
    {"description": "Frame Tinter spares kit (section header, price=0)", "unit_price": 0.0},
    {"description": "Tinter 1300 spares kit (section header, price=0)", "unit_price": 0.0},
    {"description": "Tinter 2600 spares kit (section header, price=0)", "unit_price": 0.0},
]
PRODUCT_NAMES = ["— custom —"] + [
    p["description"][:65] + ("…" if len(p["description"]) > 65 else "")
    for p in PRODUCTS
]

CURRENCIES = ["EUR", "USD", "GBP", "CHF", "CNY", "RUB", "— custom —"]

HS_CODES = ["84.66.9195", "84.79.8998", "84.48.5900", "84.77.9000", "39.26.3000", "84.73.3000"]
PAYMENT_OPTIONS = [
    "In advance by T/t transfer",
    "100% by T/T transfer at the order",
    "50% advance, 50% before shipment",
    "30 days from invoice date",
    "Letter of credit at sight",
]
DELIVERY_TERMS_OPTIONS = [
    "DAP destination",
    "DAP Shenzhen (CN)",
    "DAP Perrysburg (USA)",
    "EXW Schio (Italy)",
    "CIF destination airport",
    "FCA Schio (Italy)",
]
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
    """Replace placeholders, preserving the formatting of the run that held the placeholder."""
    for key, val in replacements.items():
        # Find which run contains the key
        full_text = "".join(run.text for run in para.runs)
        if key not in full_text:
            continue
        # Find the run that contains (or starts) the placeholder
        keeper_run = None
        for run in para.runs:
            if key in run.text or (run.text and run.text in key):
                keeper_run = run
                break
        # If not found in a single run, find the bold one or last non-empty
        if keeper_run is None:
            for run in para.runs:
                if run.bold:
                    keeper_run = run
                    break
        if keeper_run is None and para.runs:
            keeper_run = para.runs[-1]
        # Merge all runs into keeper_run with replacement applied
        new_text = full_text.replace(key, val)
        if para.runs:
            para.runs[0].text = new_text
            # Copy formatting from keeper_run to runs[0]
            if keeper_run and keeper_run != para.runs[0]:
                para.runs[0].bold = keeper_run.bold
                para.runs[0].italic = keeper_run.italic
                if keeper_run.font.name:
                    para.runs[0].font.name = keeper_run.font.name
                if keeper_run.font.size:
                    para.runs[0].font.size = keeper_run.font.size
            for run in para.runs[1:]:
                run.text = ""

def set_cell_text(cell, text, bold=False, italic=False, font_name=None, font_size=None):
    """Clear cell and write text with explicit formatting."""
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ""
    para = cell.paragraphs[0]
    if para.runs:
        run = para.runs[0]
    else:
        run = para.add_run()
    run.text = text
    run.bold = bold
    run.italic = italic
    if font_name:
        run.font.name = font_name
    if font_size:
        run.font.size = Pt(font_size)

def remove_row_borders(row):
    """Remove all borders from a table row's cells."""
    for cell in row.cells:
        tc = cell._tc
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = OxmlElement('w:tcPr')
            tc.insert(0, tcPr)
        # Remove existing borders element if any
        existing = tcPr.find(qn('w:tcBorders'))
        if existing is not None:
            tcPr.remove(existing)
        # Add new borders element with all nil
        tcBorders = OxmlElement('w:tcBorders')
        for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{side}')
            border.set(qn('w:val'), 'nil')
            tcBorders.append(border)
        tcPr.append(tcBorders)

def remove_bottom_border_from_row(row):
    """Remove only the bottom border from all cells in a row."""
    for cell in row.cells:
        tc = cell._tc
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = OxmlElement('w:tcPr')
            tc.insert(0, tcPr)
        existing = tcPr.find(qn('w:tcBorders'))
        if existing is None:
            existing = OxmlElement('w:tcBorders')
            tcPr.append(existing)
        # Set bottom to nil
        bottom = existing.find(qn('w:bottom'))
        if bottom is not None:
            existing.remove(bottom)
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'nil')
        existing.append(bottom)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "line_items" not in st.session_state:
    st.session_state.line_items = [
        {"product_idx": 0, "description": "", "details": "", "qty": 1.0, "unit_price": 0.0}
    ]

def add_line():
    st.session_state.line_items.append(
        {"product_idx": 0, "description": "", "details": "", "qty": 1.0, "unit_price": 0.0}
    )

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.set_page_config(page_title="Proforma Generator", layout="wide")
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
# Curly apostrophe \u2019 to match template exactly
formatted_date = selected_date.strftime('%d/%m/') + "\u2019" + year_2digit

# ── 2. CLIENT ──
st.subheader("2. Client")
col1, col2 = st.columns([1, 3])
with col1:
    salutation = st.selectbox("Salutation", ["Mr.", "Ms.", "Dr."])
with col2:
    full_name = st.text_input("Contact Full Name", placeholder="e.g. John Smith")

company = st.text_input("Company Name", placeholder="e.g. Vitrex s.r.o.")
address = st.text_input("Address", placeholder="e.g. Zeyerova 1334")

col3, col4, col5 = st.columns(3)
with col3:
    zip_code = st.text_input("Zip", placeholder="337 01")
with col4:
    city = st.text_input("City", placeholder="Shenzhen")
with col5:
    region = st.text_input("Region", placeholder="(optional)")

country = st.text_input("Country", placeholder="e.g. China")

# ── 3. CURRENCY ──
st.subheader("3. Currency")
currency_choice = st.selectbox("Currency (ISO)", CURRENCIES)
if currency_choice == "— custom —":
    currency = st.text_input("Enter ISO currency code", placeholder="e.g. AED, BRL, INR")
else:
    currency = currency_choice

# ── 4. LINE ITEMS ──
st.subheader("4. Line Items")
st.caption("Select from catalogue or type a custom product name.")

items_to_remove = []
for i, item in enumerate(st.session_state.line_items):
    with st.container():
        c1, c2, c3, c4 = st.columns([3, 1.5, 1.5, 0.4])
        with c1:
            # Single dropdown — catalogue + custom option
            prod_idx = st.selectbox(
                f"Product Name #{i+1} (bold in document)",
                range(len(PRODUCT_NAMES)),
                format_func=lambda x: PRODUCT_NAMES[x],
                key=f"prod_{i}",
                index=item["product_idx"]
            )
            if prod_idx != item["product_idx"]:
                item["product_idx"] = prod_idx
                if prod_idx > 0:
                    item["description"] = PRODUCTS[prod_idx - 1]["description"]
                    item["unit_price"] = PRODUCTS[prod_idx - 1]["unit_price"]
                else:
                    item["description"] = ""
                    item["unit_price"] = 0.0
            # If custom, show text input to type name
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
            item["unit_price"] = st.number_input(
                f"Unit Price ({currency})", min_value=0.0,
                value=float(item["unit_price"]), step=0.01,
                format="%.2f", key=f"price_{i}"
            )
        with c4:
            st.write("")
            st.write("")
            if st.button("🗑", key=f"del_{i}", help="Remove line"):
                items_to_remove.append(i)

        line_total = item["qty"] * item["unit_price"]
        st.caption(f"Line total: {currency} {line_total:,.2f}")
        st.divider()

for i in sorted(items_to_remove, reverse=True):
    st.session_state.line_items.pop(i)
if items_to_remove:
    st.rerun()

st.button("➕ Add Line Item", on_click=add_line)

grand_total = sum(item["qty"] * item["unit_price"] for item in st.session_state.line_items)
st.markdown(f"### 💰 Total: {currency} {grand_total:,.2f}")

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
        delivery_terms = st.text_input("Custom Delivery Terms")

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

        # Use curly apostrophe to match template
        header_replacements = {
            f"Schio, [DD/MM/\u2019YY]": f"Schio, {formatted_date}",
            f"[DD/MM/\u2019YY]": formatted_date,
            "[COMPANY NAME]": company,
            "[Address]": address,
            "[Zip] [City], [Region]": zip_city,
            "[Country]": country,
            "Mr./Ms. [Full Name]": f"{salutation} {full_name}",
            "[Full Name]": full_name,
            "[NNN/YY]": proforma_number,
        }

        try:
            template_path = os.path.join(os.path.dirname(__file__), "proforma_template_eng.docx")
            doc = Document(template_path)
        except Exception as e:
            st.error(f"❌ Template not found: {e}")
            st.stop()

        # Replace header paragraphs (preserves bold/italic from template)
        for para in doc.paragraphs:
            replace_in_paragraph(para, header_replacements)

        # ── Product table (Table 0) ──
        table = doc.tables[0]

        # Remove all rows except header row
        while len(table.rows) > 1:
            tr = table.rows[-1]._tr
            tr.getparent().remove(tr)

        # Add line item rows
        valid_items = [it for it in st.session_state.line_items if it["description"].strip()]
        for idx, item in enumerate(valid_items):
            pos = (idx + 1) * 10
            line_total = item["qty"] * item["unit_price"]

            new_tr = copy.deepcopy(table.rows[0]._tr)
            table._tbl.append(new_tr)
            new_row = table.rows[-1]

            qty_str = f"{item['qty']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            price_str = f"{item['unit_price']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            total_str = f"{int(round(line_total)):,}".replace(",", ".") + ",-"

            cells = new_row.cells

            # Description cell: bold product name line 1, normal details line 2
            desc_cell = cells[1]
            # Clear all existing paragraphs content
            for para in desc_cell.paragraphs:
                for run in para.runs:
                    run.text = ""
                    # Clear all run XML properties so no inherited italic/bold
                    rPr = run._r.find(qn('w:rPr'))
                    if rPr is not None:
                        run._r.remove(rPr)

            first_para = desc_cell.paragraphs[0]
            r = first_para.add_run(item["description"])
            r.bold = True
            r.italic = False

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

            set_cell_text(cells[0], str(pos),   bold=False, italic=False)
            set_cell_text(cells[2], qty_str,    bold=False, italic=False)
            set_cell_text(cells[3], price_str,  bold=False, italic=False)
            # ISO currency: Verdana 10, not italic
            set_cell_text(cells[4], currency,   bold=False, italic=False, font_name="Verdana", font_size=10)
            set_cell_text(cells[5], total_str,  bold=False, italic=False)

            # No borders on data rows
            remove_row_borders(new_row)

        # Add total row — bold, not italic, double top border, no wrap
        new_tr = copy.deepcopy(table.rows[0]._tr)
        table._tbl.append(new_tr)
        total_row = table.rows[-1]
        tcells = total_row.cells
        total_str = f"{int(round(grand_total)):,}".replace(",", ".") + ",-"
        total_label = f"TOTAL PRICE \u2013 {delivery_terms} -"

        set_cell_text(tcells[0], total_label, bold=True, italic=False)
        set_cell_text(tcells[1], "",          bold=True, italic=False)
        set_cell_text(tcells[2], "",          bold=True, italic=False)
        set_cell_text(tcells[3], "",          bold=True, italic=False)
        set_cell_text(tcells[4], currency,    bold=True, italic=False, font_name="Verdana", font_size=10)
        set_cell_text(tcells[5], total_str,   bold=True, italic=False)

        # Remove all borders then add double top border + no bottom border on every cell
        for cell in total_row.cells:
            tc = cell._tc
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = OxmlElement('w:tcPr')
                tc.insert(0, tcPr)
            # Remove existing borders
            existing = tcPr.find(qn('w:tcBorders'))
            if existing is not None:
                tcPr.remove(existing)
            tcBorders = OxmlElement('w:tcBorders')
            # Double top border
            top = OxmlElement('w:top')
            top.set(qn('w:val'), 'double')
            top.set(qn('w:sz'), '6')
            top.set(qn('w:space'), '0')
            top.set(qn('w:color'), 'auto')
            tcBorders.append(top)
            # All other sides nil
            for side in ['left', 'bottom', 'right']:
                b = OxmlElement(f'w:{side}')
                b.set(qn('w:val'), 'nil')
                tcBorders.append(b)
            tcPr.append(tcBorders)
            # No wrap
            noWrap = OxmlElement('w:noWrap')
            tcPr.append(noWrap)

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
                set_cell_text(terms_table.rows[row_idx].cells[1], value, bold=False, italic=False)

        # Save to buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        # Save to Supabase
        save_proforma(proforma_number, company, grand_total, currency)

        st.success(f"✅ Proforma {proforma_number} ready! Total: {currency} {grand_total:,.2f}")
        st.download_button(
            label="📄 Download Word Document",
            data=buffer,
            file_name=f"{doc_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
