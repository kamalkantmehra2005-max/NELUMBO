import streamlit as st
import pdfplumber
import re
from docx import Document
from copy import deepcopy
from io import BytesIO
from datetime import datetime

# ==============================
# PAGE CONFIG
# ==============================
st.set_page_config(
    page_title="Nelumbo",
    page_icon="🌸",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================
# DARK MODE UI
# ==============================
st.markdown("""
<style>

html, body, [class*="css"] {
    background-color: #0f1117;
    color: white;
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

# ==============================
# HEADER
# ==============================
st.markdown("""
<div class="logo">🌸</div>
<div class="title">Nelumbo</div>
<div class="subtitle">Patent AutoFill System</div>
""", unsafe_allow_html=True)

# ==============================
# SIDEBAR
# ==============================
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

st.sidebar.info("Upload IASR/PCT PDF and Word template.")

# ==============================
# ADDRESS SPLIT
# ==============================
def split_address(addr):

    parts = [p.strip() for p in addr.split(",")]

    result = {
        "house_no": "",
        "street": "",
        "city": "",
        "state": "",
        "country": "",
        "pin": ""
    }

    if len(parts) > 0:
        result["street"] = parts[0]

    if len(parts) > 1:
        result["city"] = parts[-3] if len(parts) >= 3 else ""

    if len(parts) > 0:
        last = parts[-1]

        pin_match = re.search(r'(\d{6})', last)
        country_match = re.search(r'\((\w+)\)', last)

        if pin_match:
            result["pin"] = pin_match.group(1)

        if country_match:
            result["country"] = country_match.group(1)

        result["state"] = re.sub(r'\d+', '', last).replace("(CN)", "").strip()

    return result

# ==============================
# PDF EXTRACTION
# ==============================
def extract_data(pdf_file):

    data = {}

    with pdfplumber.open(pdf_file) as pdf:
        text = " ".join([p.extract_text() or "" for p in pdf.pages])

    text = re.sub(r'\s+', ' ', text)

    def find(pattern):
        m = re.search(pattern, text, re.DOTALL)
        return m.group(1).strip() if m else ""

    # ==========================
    # DATE TAGS
    # ==========================
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

    # ==========================
    # BASIC DETAILS
    # ==========================
    data["application_no"] = find(r'Application Number:\s*(PCT/\S+)')
    data["publication_date"] = find(r'Publication date:\s*(.*?)\(')
    data["title"] = find(r'Title \(EN\):\s*(.*?)\s\(')

    # ==========================
    # ABSTRACT
    # ==========================
    abstract_match = re.search(
        r'Abstract:\s*(.*?)(?:Claims|Description|Drawings)',
        text,
        re.DOTALL
    )

    data["abstract"] = abstract_match.group(1).strip() if abstract_match else ""

    # ==========================
    # APPLICANT
    # ==========================
    applicant_full = find(r'Applicant\(s\):\s*(.*?)\(for')

    if ";" in applicant_full:
        name, addr = applicant_full.split(";", 1)

        data["applicant_name"] = name.strip()
        data["applicant_address"] = addr.strip()

        app_addr = split_address(addr)

        data["app_house_no"] = app_addr["house_no"]
        data["app_street"] = app_addr["street"]
        data["app_city"] = app_addr["city"]
        data["app_state"] = app_addr["state"]
        data["app_country"] = app_addr["country"]
        data["app_pin"] = app_addr["pin"]

    # ==========================
    # INVENTORS
    # ==========================
    inventors = re.findall(r'([A-Z]+,\s[A-Za-z]+);(.*?)\(CN\)', text)

    data["inventors"] = []

    for name, addr in inventors:

        full_addr = addr.strip() + " (CN)"

        split_addr = split_address(full_addr)

        data["inventors"].append({
            "name": name.strip(),
            "house_no": split_addr["house_no"],
            "street": split_addr["street"],
            "city": split_addr["city"],
            "state": split_addr["state"],
            "country": split_addr["country"],
            "pin": split_addr["pin"]
        })

    # ==========================
    # PRIORITY
    # ==========================
    data["priority_no"] = find(r'(\d{12}\.\w)')
    data["priority_date"] = find(r'(\d{2}\s\w+\s\d{4})')
    data["priority_country"] = "CN"

    return data

# ==============================
# INVENTOR TABLE
# ==============================
def fill_inventor_table(doc, inventors):

    for table in doc.tables:

        for row in table.rows:

            if "{{inv_name}}" in row.cells[0].text:

                template_row = row

                table._tbl.remove(row._tr)

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

                        for tag, value in replacements.items():
                            cell.text = cell.text.replace(tag, value)

                    table._tbl.append(new_row._tr)

                return

# ==============================
# DOCUMENT GENERATOR
# ==============================
def generate_doc(template_file, data):

    doc = Document(template_file)

    # Paragraph replace
    for para in doc.paragraphs:

        for k, v in data.items():

            if isinstance(v, str):
                para.text = para.text.replace(f"{{{{{k}}}}}", v)

    # Table replace
    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                for k, v in data.items():

                    if isinstance(v, str):
                        cell.text = cell.text.replace(f"{{{{{k}}}}}", v)

    # Dynamic inventor rows
    fill_inventor_table(doc, data["inventors"])

    output_stream = BytesIO()

    doc.save(output_stream)

    output_stream.seek(0)

    return output_stream

# ==============================
# DRAG & DROP UPLOAD
# ==============================
st.markdown("## 📂 Upload Files")

pdf_file = st.file_uploader(
    "Drag & Drop IASR / PCT PDF",
    type=["pdf"]
)

template_file = st.file_uploader(
    "Drag & Drop Word Template (.docx)",
    type=["docx"]
)

# ==============================
# GENERATE BUTTON
# ==============================
if st.button("🚀 Generate Patent Document"):

    if pdf_file and template_file:

        progress = st.progress(0)

        status = st.empty()

        status.text("Reading PDF...")
        progress.progress(20)

        data = extract_data(pdf_file)

        status.text("Processing Inventors...")
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

# ==============================
# FOOTER
# ==============================
st.markdown("""
<div class="footer">
<hr>
<b>Design by Kamal Kant</b><br>
<i>Not recommended for convention and ordinary file</i>
</div>
""", unsafe_allow_html=True)