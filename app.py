import streamlit as st
import pdfplumber
import re
from docx import Document
from docx.shared import Inches
from datetime import datetime
from copy import deepcopy
from io import BytesIO


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Nelumbo",
    page_icon="🌸",
    layout="wide"
)


# =========================================================
# CUSTOM CSS
# =========================================================

st.markdown("""
<style>

.stApp {
    background-color: #0f1117;
    color: white;
}

h1,h2,h3,h4,p,div,label {
    color:white !important;
}

.stButton>button {
    background:#7c4dff;
    color:white;
    border:none;
    border-radius:10px;
    padding:12px;
    width:100%;
    font-size:16px;
    font-weight:bold;
}

.stDownloadButton>button {
    background:#00c853;
    color:white;
    border:none;
    border-radius:10px;
    padding:12px;
    width:100%;
    font-size:16px;
    font-weight:bold;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# HEADER
# =========================================================

st.markdown("""
<div style='text-align:center;'>

<div style='font-size:80px;'>
🌸
</div>

<h1>
Nelumbo
</h1>

<p>
Patent Automation Tool
</p>

</div>
""", unsafe_allow_html=True)


# =========================================================
# ADDRESS SPLITTER
# =========================================================

def split_address(addr):

    parts = [p.strip() for p in addr.split(",")]

    return {

        "house": parts[0] if len(parts) > 0 else "",

        "street": parts[1] if len(parts) > 1 else "",

        "city": parts[2] if len(parts) > 2 else "",

        "state": parts[3] if len(parts) > 3 else "",

        "country": parts[4] if len(parts) > 4 else "",

        "pin": re.search(r'(\d{6})', addr).group(1)
        if re.search(r'(\d{6})', addr)
        else ""

    }


# =========================================================
# PDF EXTRACTION
# =========================================================

def extract_data(pdf_file):

    data = {}

    with pdfplumber.open(pdf_file) as pdf:

        text = " ".join([
            page.extract_text() or ""
            for page in pdf.pages
        ])

    text = re.sub(r'\s+', ' ', text)

    # =====================================================
    # APPLICANT
    # =====================================================

    applicant_match = re.search(
        r'Applicant\(s\):(.*?);(.*?)(?:\([A-Z]{2}\))',
        text
    )

    if applicant_match:

        applicant_name = applicant_match.group(1).strip()

        applicant_address = applicant_match.group(2).strip()

    else:

        applicant_name = ""
        applicant_address = ""

    data["applicant"] = applicant_name
    data["applicant_name"] = applicant_name

    app_addr = split_address(applicant_address)

    data["app_house_no"] = app_addr["house"]
    data["app_street"] = app_addr["street"]
    data["app_city"] = app_addr["city"]
    data["app_state"] = app_addr["state"]
    data["app_country"] = app_addr["country"]
    data["app_pin"] = app_addr["pin"]

    # =====================================================
    # TITLE
    # =====================================================

    title_match = re.search(
        r'Title.*?:\s*(.*?)\s{2,}',
        text
    )

    data["title"] = (
        title_match.group(1).strip()
        if title_match else ""
    )

    # =====================================================
    # PCT
    # =====================================================

    pct_match = re.search(
        r'PCT/[A-Z]{2}\d{4}/\d+',
        text
    )

    data["application_no"] = (
        pct_match.group(0)
        if pct_match else ""
    )

    # =====================================================
    # PUBLICATION DATE
    # =====================================================

    pub_match = re.search(
        r'Publication date:\s*(.*?)\(',
        text
    )

    data["publication_date"] = (
        pub_match.group(1).strip()
        if pub_match else ""
    )

    # =====================================================
    # PRIORITY
    # =====================================================

    priority_match = re.search(
        r'(\d{12}\.\w)',
        text
    )

    data["priority_no"] = (
        priority_match.group(1)
        if priority_match else ""
    )

    data["priority_country"] = "CN"

    data["priority_date"] = datetime.today().strftime(
        "%d %B %Y"
    )

    # =====================================================
    # INVENTOR SECTION
    # =====================================================

    inventor_section = re.search(
        r'\(72\)\s*Inventor\(s\):(.*?)\(74\)\s*Agent\(s\):',
        text,
        re.DOTALL
    )

    inventor_text = (
        inventor_section.group(1)
        if inventor_section else ""
    )

    inventor_pattern = re.findall(
        r'([A-Z][A-Z\s\-\,]+);(.*?)(?=[A-Z][A-Z\s\-\,]+;|$)',
        inventor_text
    )

    data["inventors"] = []

    inventor_names = []

    for idx, (name, address) in enumerate(
        inventor_pattern,
        start=1
    ):

        clean_name = name.strip()

        inventor_names.append(clean_name)

        split_addr = split_address(address)

        inventor_data = {

            "name": clean_name,

            "house": split_addr["house"],
            "street": split_addr["street"],
            "city": split_addr["city"],
            "state": split_addr["state"],
            "country": split_addr["country"],
            "pin": split_addr["pin"]

        }

        data["inventors"].append(
            inventor_data
        )

        data[f"inv_name_{idx}"] = inventor_data["name"]

        data[f"inv_house_{idx}"] = inventor_data["house"]
        data[f"inv_street_{idx}"] = inventor_data["street"]
        data[f"inv_city_{idx}"] = inventor_data["city"]
        data[f"inv_state_{idx}"] = inventor_data["state"]
        data[f"inv_country_{idx}"] = inventor_data["country"]
        data[f"inv_pin_{idx}"] = inventor_data["pin"]

    data["inventor_names"] = ", ".join(
        inventor_names
    )

    # =====================================================
    # DATES
    # =====================================================

    today = datetime.today()

    day = today.strftime("%d")
    month = today.strftime("%B")
    year = today.strftime("%Y")

    data["today_date_short"] = today.strftime(
        "%B %d, %Y"
    )

    data["today_date_long"] = (
        f"{day}th day of {month}, {year}"
    )

    return data


# =========================================================
# TAG REPLACEMENT
# =========================================================

def replace_text_preserve(paragraph, key, value):

    if key in paragraph.text:

        for run in paragraph.runs:

            if key in run.text:

                run.text = run.text.replace(
                    key,
                    value
                )


def replace_all_tags(doc, data):

    # Paragraphs
    for para in doc.paragraphs:

        for k, v in data.items():

            if isinstance(v, str):

                replace_text_preserve(
                    para,
                    f"{{{{{k}}}}}",
                    v
                )

    # Tables
    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                for para in cell.paragraphs:

                    for k, v in data.items():

                        if isinstance(v, str):

                            replace_text_preserve(
                                para,
                                f"{{{{{k}}}}}",
                                v
                            )


# =========================================================
# INVENTOR TABLE
# =========================================================

def add_inventor_block(table, inventor):

    # ==========================================
    # MAIN INVENTOR ROW
    # ==========================================

    row_cells = table.add_row().cells

    row_cells[0].text = inventor["name"]
    row_cells[1].text = "Unknown"
    row_cells[2].text = inventor["country"]

    # ==========================================
    # ADDRESS HEADER
    # ==========================================

    title_row = table.add_row().cells

    title_row[0].text = "Address of the Inventor"

    title_row[0].merge(title_row[1])
    title_row[0].merge(title_row[2])

    # ==========================================
    # ADDRESS TABLE
    # ==========================================

    address_row = table.add_row().cells

    nested = address_row[0].add_table(
        rows=6,
        cols=2
    )

    nested.style = "Table Grid"

    labels = [
        "House",
        "Street",
        "City",
        "State",
        "Country",
        "Pin Code"
    ]

    values = [
        inventor["house"],
        inventor["street"],
        inventor["city"],
        inventor["state"],
        inventor["country"],
        inventor["pin"]
    ]

    for i in range(6):

        nested.cell(i, 0).text = labels[i]
        nested.cell(i, 1).text = values[i]

    address_row[0].merge(address_row[1])
    address_row[0].merge(address_row[2])


# =========================================================
# DOCUMENT GENERATOR
# =========================================================

def generate_doc(template_file, data):

    doc = Document(template_file)

    # =====================================================
    # REPLACE TAGS
    # =====================================================

    replace_all_tags(doc, data)

    # =====================================================
    # INVENTOR TABLE DETECTION
    # =====================================================

    for table in doc.tables:

        found = False

        for row in table.rows:

            text = " ".join(
                cell.text for cell in row.cells
            ).lower()

            if "name in full" in text:

                found = True
                break

        if found:

            # REMOVE OLD SAMPLE ROWS
            while len(table.rows) > 2:

                tbl = table._tbl

                tbl.remove(
                    table.rows[-1]._tr
                )

            # ADD INVENTORS
            for inventor in data["inventors"]:

                add_inventor_block(
                    table,
                    inventor
                )

            break

    # =====================================================
    # SAVE OUTPUT
    # =====================================================

    output = BytesIO()

    doc.save(output)

    output.seek(0)

    return output


# =========================================================
# FILE UPLOADS
# =========================================================

st.markdown("## Upload Files")

pdf_file = st.file_uploader(
    "Upload PDF Information Sheet",
    type=["pdf"]
)

template_file = st.file_uploader(
    "Upload FORM-1 Template",
    type=["docx"]
)


# =========================================================
# GENERATE BUTTON
# =========================================================

if st.button("🚀 Generate Patent Document"):

    if pdf_file and template_file:

        progress = st.progress(0)

        progress.progress(20)

        data = extract_data(pdf_file)

        progress.progress(60)

        output = generate_doc(
            template_file,
            data
        )

        progress.progress(100)

        st.success(
            "Document Generated Successfully"
        )

        st.download_button(
            label="⬇ Download Filled FORM-1",
            data=output,
            file_name="FORM_1_FILLED.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    else:

        st.error(
            "Please upload both files"
        )


# =========================================================
# FOOTER
# =========================================================

st.markdown("""
<hr>

<div style='text-align:center;color:gray;'>

Design by Kamal Kant<br>
Not recommended for convention and ordinary file

</div>
""", unsafe_allow_html=True)
