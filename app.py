import streamlit as st
import pdfplumber
import re
from docx import Document
from copy import deepcopy
from io import BytesIO
from datetime import datetime

# ==========================================
# PAGE CONFIG
# ==========================================
st.set_page_config(
    page_title="Nelumbo",
    page_icon="🪷",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# DARK UI
# ==========================================
st.markdown("""
<style>

html, body, [class*="css"] {
    background-color: #0f1117;
    color: white;
    font-family: 'Times New Roman';
}

.stApp {
    background-color: #0f1117;
}

h1,h2,h3,h4,h5,h6,p,label,div {
    color: white !important;
}

[data-testid="stFileUploader"] {
    border: 2px dashed #7c4dff;
    border-radius: 15px;
    padding: 20px;
    background-color: #1a1d29;
}

.stButton>button {
    background-color: #7c4dff;
    color: white;
    border-radius: 12px;
    border: none;
    padding: 12px 20px;
    font-size: 16px;
    font-weight: bold;
    width: 100%;
}

.stButton>button:hover {
    background-color: #9575ff;
}

.footer {
    text-align:center;
    font-size:13px;
    color:gray;
    padding-top:30px;
    padding-bottom:10px;
}

.logo {
    text-align:center;
    font-size:80px;
}

.title {
    text-align:center;
    font-size:42px;
    font-weight:bold;
    margin-bottom:0;
}

.subtitle {
    text-align:center;
    color:gray;
    margin-top:0;
    margin-bottom:30px;
    font-size:18px;
}

</style>
""", unsafe_allow_html=True)


# ==========================================
# HEADER
# ==========================================
col1, col2, col3 = st.columns([1,2,1])

with col2:

    st.image(
        "logo.png.png",
        width=180
    )

    st.markdown("""
    <div class="title">Nelumbo</div>
    <div class="subtitle">
        Patent AutoFill System
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# SIDEBAR
# ==========================================
st.sidebar.title("Nelumbo Dashboard")

form_type = st.sidebar.selectbox(
    "Select Form Type",
    [
        "FORM-1",
        "FORM-2",
        "FORM-3",
        "FORM-5",
        "CUSTOM TEMPLATE"
    ]
)

st.sidebar.info("Upload IASR / PCT PDF and Word Template")

# ==========================================
# COUNTRY MAP
# ==========================================
COUNTRY_MAP = {
    "CN": "People's Republic of China",
    "IN": "India",
    "US": "USA",
    "JP": "Japan",
    "KR": "Republic of Korea",
    "EP": "European Patent Office",
    "GB": "United Kingdom"
}

# ==========================================
# ADDRESS SPLITTER
# ==========================================
def split_address(addr):

    result = {
        "house_no": "",
        "street": "",
        "city": "",
        "state": "",
        "country": "",
        "pin": ""
    }

    # COUNTRY
    country_match = re.search(r'\(([A-Z]{2})\)', addr)

    country_code = ""

    if country_match:
        country_code = country_match.group(1)

    result["country"] = COUNTRY_MAP.get(country_code, country_code)

    # REMOVE COUNTRY CODE
    addr = re.sub(r'\([A-Z]{2}\)', '', addr).strip()

    # PIN CODE
    pin_match = re.search(r'([A-Za-z0-9 ]+)$', addr)

    if pin_match:

        pin = pin_match.group(1).strip()

        if len(pin) <= 12:

            result["pin"] = pin

            addr = addr[:addr.rfind(pin)].strip()

    # SPLIT BY COMMA
    parts = [p.strip() for p in addr.split(",")]

    parts = [p for p in parts if p]

    # HOUSE NO
    if len(parts) > 0:

        first = parts[0]

        if re.search(r'\d', first):

            result["house_no"] = first + ","

            parts = parts[1:]

    # STREET
    if len(parts) > 0:
        result["street"] = parts[0] + ","

    # CITY
    if len(parts) > 1:
        result["city"] = parts[1] + ","

    # STATE
    if len(parts) > 2:
        result["state"] = ", ".join(parts[2:]).strip()

    return result

# ==========================================
# SAFE STYLE PRESERVE REPLACER
# ==========================================
def replace_text_preserve_style(paragraph, key, value):

    if key not in paragraph.text:
        return

    full_text = "".join(run.text for run in paragraph.runs)

    if key not in full_text:
        return

    replaced_text = full_text.replace(key, value)

    first_run = paragraph.runs[0]

    first_run.text = replaced_text

    for run in paragraph.runs[1:]:
        run.text = ""

# ==========================================
# PDF DATA EXTRACTOR
# ==========================================
def extract_data(pdf_file):

    data = {}

    with pdfplumber.open(pdf_file) as pdf:
        text = " ".join([p.extract_text() or "" for p in pdf.pages])

    text = re.sub(r'\s+', ' ', text)

    def find(pattern):

        m = re.search(pattern, text, re.DOTALL)

        return m.group(1).strip() if m else ""

    # ======================================
    # TODAY DATE
    # ======================================
    today = datetime.today()

    day = today.strftime("%d")
    month = today.strftime("%B")
    year = today.strftime("%Y")

    suffix = "th"

    if day.endswith("1") and day != "11":
        suffix = "st"

    elif day.endswith("2") and day != "12":
        suffix = "nd"

    elif day.endswith("3") and day != "13":
        suffix = "rd"

    data["today_date_long"] = f"{day}{suffix} day of {month}, {year}"

    data["today_date_short"] = today.strftime("%B %d, %Y")

    # ======================================
    # BASIC DETAILS
    # ======================================
    data["application_no"] = find(r'Application Number:\s*(PCT/\S+)')

    date_match = re.search(
        r'Publication date.*?(\d{2}\.\d{2}\.\d{4})',
        text
    )

    data["publication_date"] = date_match.group(1) if date_match else ""

    data["title"] = find(r'Title \(EN\):\s*(.*?)\s\(')

    # ======================================
    # ABSTRACT
    # ======================================
    abstract_match = re.search(
        r'Abstract:\s*\(EN\):(.*?)(?:\([A-Z]{2}\):|Claims|Description)',
        text,
        re.DOTALL
    )

    data["abstract"] = abstract_match.group(1).strip() if abstract_match else ""

    # ======================================
    # APPLICANT
    # ======================================
    applicant_match = re.search(
        r'Applicant\(s\):(.*?)\[',
        text
    )

    if applicant_match:

        applicant_full = applicant_match.group(1).strip()

        if ";" in applicant_full:

            name, addr = applicant_full.split(";", 1)

            data["applicant_name"] = name.strip()

            data["applicant_address"] = addr.strip()

            split_addr = split_address(addr)

            data["app_house_no"] = split_addr["house_no"]
            data["app_street"] = split_addr["street"]
            data["app_city"] = split_addr["city"]
            data["app_state"] = split_addr["state"]
            data["app_country"] = split_addr["country"]
            data["app_pin"] = split_addr["pin"]

    # ======================================
    # INVENTORS
    # ======================================
    inventors = re.findall(
        r'([A-Z][^;]+);(.*?\([A-Z]{2}\))',
        text
    )

    data["inventors"] = []

    for name, addr in inventors:

        name = name.strip()

        addr = addr.strip()

        split_addr = split_address(addr)

        data["inventors"].append({
            "name": name,
            "house_no": split_addr["house_no"],
            "street": split_addr["street"],
            "city": split_addr["city"],
            "state": split_addr["state"],
            "country": split_addr["country"],
            "pin": split_addr["pin"]
        })

    # ======================================
    # PRIORITY
    # ======================================
    data["priority_no"] = find(r'(\d{12}\.\w)')

    priority_date_match = re.search(
        r'(\d{2}\.\d{2}\.\d{4})',
        text
    )

    data["priority_date"] = priority_date_match.group(1) if priority_date_match else ""

    data["priority_country"] = "China"

    return data

# ==========================================
# INVENTOR TABLE AUTO ROWS
# ==========================================
def fill_inventor_table(doc, inventors):

    for table in doc.tables:

        rows = table.rows

        for i, row in enumerate(rows):

            row_text = " ".join(cell.text for cell in row.cells)

            if "{{inv_name}}" in row_text:

                template_row = row

                for inv in inventors:

                    new_row = deepcopy(template_row)

                    replacements = {
                        "{{inv_name}}": inv["name"],
                        "{{house_no}}": inv["house_no"],
                        "{{street}}": inv["street"],
                        "{{city}}": inv["city"],
                        "{{state}}": inv["state"],
                        "{{country}}": inv["country"],
                        "{{pin}}": inv["pin"]
                    }

                    for cell in new_row.cells:

                        for para in cell.paragraphs:

                            for tag, value in replacements.items():

                                replace_text_preserve_style(
                                    para,
                                    tag,
                                    value
                                )

                    table._tbl.insert(i, new_row._tr)

                    i += 1

                table._tbl.remove(template_row._tr)

                return

# ==========================================
# DOC GENERATOR
# ==========================================
def generate_doc(template_file, data):

    doc = Document(template_file)

    # PARAGRAPH REPLACE
    for para in doc.paragraphs:

        for k, v in data.items():

            if isinstance(v, str):

                replace_text_preserve_style(
                    para,
                    f"{{{{{k}}}}}",
                    v
                )

    # TABLE REPLACE
    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                for para in cell.paragraphs:

                    for k, v in data.items():

                        if isinstance(v, str):

                            replace_text_preserve_style(
                                para,
                                f"{{{{{k}}}}}",
                                v
                            )

    # INVENTOR ROWS
    fill_inventor_table(doc, data["inventors"])

    output_stream = BytesIO()

    doc.save(output_stream)

    output_stream.seek(0)

    return output_stream

# ==========================================
# MAIN UI
# ==========================================
st.markdown("## 📂 Upload Files")

pdf_file = st.file_uploader(
    "Drag & Drop IASR / PCT PDF",
    type=["pdf"]
)

template_file = st.file_uploader(
    "Drag & Drop Word Template (.docx)",
    type=["docx"]
)

# ==========================================
# GENERATE BUTTON
# ==========================================
if st.button("🚀 Generate Patent Document"):

    if pdf_file and template_file:

        progress = st.progress(0)

        status = st.empty()

        status.text("Reading PDF...")
        progress.progress(20)

        data = extract_data(pdf_file)

        status.text("Extracting Data...")
        progress.progress(50)

        output = generate_doc(template_file, data)

        status.text("Generating Word File...")
        progress.progress(80)

        status.text("Completed Successfully")
        progress.progress(100)

        st.success("Patent document generated successfully.")

        st.download_button(
            label="⬇ Download Filled Document",
            data=output,
            file_name=f"{form_type}_Filled.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    else:

        st.error("Please upload both PDF and Word Template.")

# ==========================================
# FOOTER
# ==========================================
st.markdown("""
<div class="footer">
<hr>
<b>Design by Kamal Kant</b><br>
<i>Not recommended for convention and ordinary file</i>
</div>
""", unsafe_allow_html=True)
