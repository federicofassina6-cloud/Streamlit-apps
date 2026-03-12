import streamlit as st
from docx import Document
from datetime import date
import io
import requests

SUPABASE_URL = "https://lztrggttkgvgjouofibd.supabase.co"
SUPABASE_KEY = "sb_publishable_2kCkVA7G9VdPWiIXBGIFPw_O_0gfReQ"

HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json"
}

def get_next_offer_number():
    year_2digit = date.today().strftime('%y')
    # Count existing offers this year
    response = requests.get(
        f"{SUPABASE_URL}/rest/v1/offerte",
        headers=HEADERS,
        params={"offer_number": f"like.%/{year_2digit}", "select": "offer_number"}
    )
    existing = response.json()
    next_num = len(existing) + 1
    return f"{next_num:03d}/{year_2digit}"

def save_offerta(offer_number, client_company):
    requests.post(
        f"{SUPABASE_URL}/rest/v1/offerte",
        headers=HEADERS,
        json={"offer_number": offer_number, "client_company": client_company}
    )

st.set_page_config(page_title="Offerta Generator", layout="centered")
st.title("📄 New Offerta")

st.subheader("1. Date")
selected_date = st.date_input("Select date", value=date.today(), label_visibility="collapsed")
year_2digit = selected_date.strftime('%y')
formatted_date = selected_date.strftime('%d/%m/') + "\u2019" + year_2digit

st.subheader("2. Customer")
col1, col2 = st.columns([1, 3])
with col1:
    salutation = st.selectbox("Salutation", ["Mr.", "Ms."])
with col2:
    full_name = st.text_input("Full Name", placeholder="e.g. Adrian Hlaváček")

company  = st.text_input("Company Name", placeholder="e.g. Vitrex s.r.o.")
address  = st.text_input("Address", placeholder="e.g. Zeyerova 1334")

col3, col4, col5 = st.columns(3)
with col3:
    zip_code = st.text_input("Zip", placeholder="337 01")
with col4:
    city = st.text_input("City", placeholder="Rokycany")
with col5:
    region = st.text_input("Region", placeholder="(optional)")

country = st.text_input("Country", placeholder="e.g. Czech Republic")

st.subheader("3. Product")
inventory      = st.text_input("Product Name", placeholder="e.g. ROYALMAC GLASSTINTER 1300")
inventory_desc = st.text_input("Product Description", placeholder="e.g. ROLLER PRINTER FOR FLOAT GLASS")

st.subheader("4. Offerta Name")
offer_number = get_next_offer_number()
default_name = f"offerta {selected_date.strftime('%m-%Y')} {company}"
offerta_name = st.text_input("Offerta Name", value=default_name)
st.caption(f"Offer number: **{offer_number}**")

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

st.divider()
if st.button("📥 Generate Offerta", type="primary", use_container_width=True):
    if not company:
        st.warning("Please enter a company name.")
    elif not full_name:
        st.warning("Please enter a contact name.")
    else:
        zip_city = f"{zip_code} {city}".strip()
        if region:
            zip_city += f", {region}"

        replacements = {
            "Schio, [DD/MM/\u2019YY]":    f"Schio, {formatted_date}",
            "[DD/MM/\u2019YY]":           formatted_date,
            "[COMPANY NAME]":             company,
            "[Address]":                  address,
            "[Zip] [City], [Region]":     zip_city,
            "[Country]":                  country,
            "Mr./Ms. [Full Name]":        f"{salutation} {full_name}",
            "[Full Name]":                full_name,
            "[NNN/YY]":                   offer_number,
            "[INVENTORY]":                inventory,
            "[INVENTORY DESCRIPTION]":    inventory_desc,
        }

        try:
            doc = Document("/workspaces/blank-app/offerta_template.docx")
        except:
            st.error("❌ Template not found.")
            st.stop()

        for para in doc.paragraphs:
            replace_in_paragraph(para, replacements)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        replace_in_paragraph(para, replacements)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        # Save to Supabase
        save_offerta(offer_number, company)

        st.success(f"✅ Offerta {offer_number} ready!")
        st.download_button(
            label="📄 Download Word Document",
            data=buffer,
            file_name=f"{offerta_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )