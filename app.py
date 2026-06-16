import json
import re
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path

import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import Workbook


st.set_page_config(page_title="HYDAC Lead Agent", layout="wide")


HEADER = [
    "Referral","Brand","Product","ReceivedDateTime","FirstName","LastName",
    "ContactTitle","Email","Company","Address","County","City","State",
    "ZipCode","Country","LeadSource1","LeadSource2","LeadSource3",
    "LeadComments","PhoneSupplied","PhSuppliedExtension","PhoneResearched",
    "CSRName","PDF","DUNS","WebAddress","Linkedin_Title","Linkedin_Link",
    "SIC","NAICS","noOfEmployees","ParentName","LineOfBusiness","PQ",
    "Latitude","Longitude","DemoLead","ScreenReason","about_me","college_1",
    "college_1_degree","college_1_start","college_1_end","college_2",
    "college_2_degree","college_2_start","college_2_end","month_of_joining",
    "about_experience","searched_on_google","linkedin_city","linkedin_state",
    "linkedin_country"
]

INTERNAL_DOMAINS = [
    "hydacusa.com",
    "hydac.com",
    "hydac-interlynx.com",
]

BAD_WEBSITE_PATTERNS = [
    "aka.ms",
    "microsoft.com",
    "office.com",
    "outlook.com",
    "safelinks",
    "proofpoint",
    "mimecast",
]

SIGNATURE_IMAGE_NAMES = {
    "image.png", "image.jpg", "image.jpeg",
    "image001.png", "image001.jpg", "image001.jpeg",
    "image002.png", "image002.jpg", "image002.jpeg",
    "image003.png", "image003.jpg", "image003.jpeg",
    "image004.png", "image004.jpg", "image004.jpeg",
    "image005.png", "image005.jpg", "image005.jpeg",
    "image006.png", "image006.jpg", "image006.jpeg",
    "image007.png", "image007.jpg", "image007.jpeg",
    "image008.png", "image008.jpg", "image008.jpeg",
    "image009.png", "image009.jpg", "image009.jpeg",
    "image010.png", "image010.jpg", "image010.jpeg",
    "logo.png", "logo.jpg", "banner.png", "banner.jpg",
    "facebook.png", "linkedin.png", "twitter.png",
    "instagram.png", "youtube.png"
}

VALID_ATTACHMENT_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv",
    ".step", ".stp", ".igs", ".iges", ".dwg", ".dxf",
    ".zip", ".rar", ".7z",
    ".png", ".jpg", ".jpeg", ".tif", ".tiff"
}


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = str(text).replace("\r", "\n")
    text = re.sub(r"\xa0", " ", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def html_to_text(html: str) -> str:
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")
    return clean_text(soup.get_text("\n"))


def is_internal_email(email: str) -> bool:
    email = (email or "").lower()
    return any(domain in email for domain in INTERNAL_DOMAINS)


def extract_emails(text: str):
    emails = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text or "")
    unique = []
    for email in emails:
        if email not in unique:
            unique.append(email)
    return unique


def extract_external_block(text: str) -> str:
    """
    If an email has an external warning block, the real customer message often starts nearby.
    Return from EXTERNAL EMAIL marker onward, otherwise full text.
    """
    markers = [
        "EXTERNAL EMAIL",
        "External Email",
        "CAUTION:",
        "Caution:",
    ]
    for marker in markers:
        idx = text.find(marker)
        if idx != -1:
            return text[idx:]
    return text


def normalize_phone(phone: str) -> str:
    phone = (phone or "").strip()
    phone = phone.replace("+", "")
    phone = re.sub(r"[().\s]+", "", phone)
    phone = phone.replace("--", "-").strip("-")
    return phone


def looks_like_valid_attachment(filename: str) -> bool:
    if not filename:
        return False

    name = Path(filename).name
    low = name.lower()
    ext = Path(low).suffix

    if ext not in VALID_ATTACHMENT_EXTENSIONS:
        return False

    if low in SIGNATURE_IMAGE_NAMES:
        return False

    if re.fullmatch(r"image\d{0,3}\.(png|jpg|jpeg|gif)", low):
        return False

    if ext not in {".png", ".jpg", ".jpeg", ".gif"}:
        return True

    # For image attachments, keep only if the filename looks specific to the customer/request.
    if re.search(r"\d{4,}", name):
        return True

    if re.search(
        r"(pump|filter|drawing|label|plate|model|part|bieri|hydraulic|spec|quote|rfbn|serial|nameplate)",
        low,
    ):
        return True

    return False


def parse_msg(uploaded_file):
    import extract_msg

    with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    msg = extract_msg.Message(tmp_path)

    sender = msg.sender or ""
    subject = msg.subject or uploaded_file.name
    date = str(msg.date) if msg.date else ""

    body = clean_text(msg.body or "")
    if not body and getattr(msg, "htmlBody", None):
        body = html_to_text(msg.htmlBody)

    full_text = clean_text("\n".join([sender, subject, body]))
    external_focus_text = clean_text(extract_external_block(full_text))

    valid_attachments = []
    ignored_attachments = []

    for att in msg.attachments:
        fname = getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or ""
        if not fname:
            continue

        if looks_like_valid_attachment(fname):
            valid_attachments.append((fname, att))
        else:
            ignored_attachments.append(fname)

    return {
        "sender": sender,
        "subject": subject,
        "date": date,
        "body": body,
        "full_text": full_text,
        "external_focus_text": external_focus_text,
        "valid_attachments": valid_attachments,
        "ignored_attachments": ignored_attachments,
    }


def fallback_agent(parsed):
    """
    Minimal fallback when no OpenAI key is present.
    It is intentionally conservative and expects manual review.
    """
    focus = parsed["external_focus_text"]
    emails = [e for e in extract_emails(focus) if not is_internal_email(e)]
    email = emails[0] if emails else ""

    first = ""
    last = ""
    if email:
        name_guess = email.split("@")[0].replace(".", " ").replace("_", " ").title()
        parts = name_guess.split()
        first = parts[0] if parts else ""
        last = " ".join(parts[1:]) if len(parts) > 1 else ""

    return {
        "FirstName": first,
        "LastName": last,
        "ContactTitle": "",
        "Email": email,
        "Company": "",
        "Address": "",
        "City": "",
        "State": "",
        "ZipCode": "",
        "Country": "",
        "PhoneSupplied": "",
        "WebAddress": "",
        "LeadComments": focus[:1200],
        "Product": "",
        "Confidence": "Low",
        "AgentReason": "Fallback mode. Add OPENAI_API_KEY in Streamlit Secrets for agent understanding."
    }


def run_ai_agent(parsed, uploaded_name, valid_attachment_names):
    api_key = st.secrets.get("OPENAI_API_KEY", "")

    if not api_key:
        return fallback_agent(parsed)

    from openai import OpenAI

    client = OpenAI(api_key=api_key)

    email_text = parsed["full_text"][:22000]
    focus_text = parsed["external_focus_text"][:12000]

    prompt = f"""
You are a HYDAC lead-processing AGENT. You are not a simple extractor.

Your job:
Read the full email and identify the real external customer/requester and customer request.

CRITICAL RULES:
1. Do NOT select HYDAC employees, HYDAC Sales, internal forwarders, or HYDAC signatures as the lead.
2. If HYDAC forwarded the email, find the original outside customer.
3. If there is an EXTERNAL EMAIL / caution block, strongly prefer the customer info and request inside or below that block.
4. Ignore Microsoft security warnings, Outlook banners, confidentiality footers, and sender-identification links.
5. Ignore HYDAC phone numbers and HYDAC signature contact data.
6. Customer phone must come from the outside customer signature only. If only HYDAC phone is found, leave PhoneSupplied blank.
7. LeadComments must be only the actual customer request, not the full email trail.
8. If the customer request is clear and short, keep the exact wording as much as possible.
9. Split person name into FirstName and LastName.
10. PhoneSupplied format:
   - Remove +, spaces, parentheses, and periods.
   - Keep hyphens when useful.
   - Multiple customer numbers separated by " : "
   - Example: 43-31669151 : 43-6649614266
11. Use the customer signature block for company, title, address, city, state, zip, country, phone, website.
12. Leave blank if not available. Do not guess.
13. For BIERI example, Greg Jia / Klein International Inc. is the customer, not HYDAC or Brandon.
14. Company should not be inferred from Gmail unless written in the signature/body.
15. Product should contain product/model/part number only if clearly present.

Valid customer attachments already detected:
{valid_attachment_names}

Uploaded file name:
{uploaded_name}

Focused external/customer block:
{focus_text}

Full email text:
{email_text}

Return ONLY valid JSON with these exact keys:
FirstName, LastName, ContactTitle, Email, Company, Address, City, State, ZipCode, Country,
PhoneSupplied, WebAddress, LeadComments, Product, Confidence, AgentReason
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0,
            messages=[
                {
                    "role": "system",
                    "content": "You are a careful lead-processing agent. Return only valid JSON. Never guess unavailable fields."
                },
                {"role": "user", "content": prompt},
            ],
        )

        content = response.choices[0].message.content.strip()
        content = content.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
        data = json.loads(content)

        required = [
            "FirstName", "LastName", "ContactTitle", "Email", "Company",
            "Address", "City", "State", "ZipCode", "Country",
            "PhoneSupplied", "WebAddress", "LeadComments", "Product",
            "Confidence", "AgentReason"
        ]
        for key in required:
            data.setdefault(key, "")

        return data

    except Exception as e:
        data = fallback_agent(parsed)
        data["AgentReason"] = f"AI failed; fallback used. Error: {e}"
        return data


def make_excel(row):
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"

    ws.append(HEADER)
    ws.append([row.get(h, "") for h in HEADER])

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        ws.column_dimensions[col_letter].width = min(max_length + 2, 45)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def make_attachment_zip(valid_attachments):
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as z:
        for fname, att in valid_attachments:
            try:
                data = att.data
                if data:
                    z.writestr(Path(fname).name, data)
            except Exception:
                pass
    return output.getvalue()


st.title("HYDAC Lead Agent V1")
st.success("Agent is running.")

with st.expander("Agent status"):
    if st.secrets.get("OPENAI_API_KEY", ""):
        st.success("AI mode active: OPENAI_API_KEY detected.")
    else:
        st.warning("Fallback/manual mode: add OPENAI_API_KEY in Streamlit Secrets for best results.")

uploaded_file = st.file_uploader("Upload .msg file", type=["msg"])

if uploaded_file:
    try:
        parsed = parse_msg(uploaded_file)
        valid_names = [name for name, _ in parsed["valid_attachments"]]
        agent_data = run_ai_agent(parsed, uploaded_file.name, valid_names)

        st.subheader("Agent Review Panel")

        col1, col2 = st.columns(2)

        with col1:
            first = st.text_input("FirstName", value=agent_data.get("FirstName", ""))
            last = st.text_input("LastName", value=agent_data.get("LastName", ""))
            title = st.text_input("ContactTitle", value=agent_data.get("ContactTitle", ""))
            email = st.text_input("Email", value=agent_data.get("Email", ""))
            company = st.text_input("Company", value=agent_data.get("Company", ""))
            address = st.text_input("Address", value=agent_data.get("Address", ""))

        with col2:
            city = st.text_input("City", value=agent_data.get("City", ""))
            state = st.text_input("State", value=agent_data.get("State", ""))
            zip_code = st.text_input("ZipCode", value=agent_data.get("ZipCode", ""))
            country = st.text_input("Country", value=agent_data.get("Country", ""))
            phone = st.text_input("PhoneSupplied", value=agent_data.get("PhoneSupplied", ""))
            website = st.text_input("WebAddress", value=agent_data.get("WebAddress", ""))

        product = st.text_input("Product", value=agent_data.get("Product", ""))
        received = st.text_input("ReceivedDateTime", value=parsed["date"])

        pdf_base = uploaded_file.name.rsplit(".", 1)[0] + ".pdf"
        pdf_value = pdf_base + (("; Attachments: " + ", ".join(valid_names)) if valid_names else "")
        pdf = st.text_input("PDF", value=pdf_value)

        comments = st.text_area(
            "LeadComments / Customer Request",
            value=agent_data.get("LeadComments", ""),
            height=180
        )

        st.subheader("Agent Decision")
        st.write("Confidence:", agent_data.get("Confidence", ""))
        st.write("Reason:", agent_data.get("AgentReason", ""))

        st.subheader("Attachment Decision")
        st.write("Valid customer attachments:")
        st.write(valid_names or "None")

        st.write("Ignored signature/generic attachments:")
        st.write(parsed["ignored_attachments"] or "None")

        row = {h: "" for h in HEADER}
        row["Product"] = product
        row["ReceivedDateTime"] = received
        row["FirstName"] = first
        row["LastName"] = last
        row["ContactTitle"] = title
        row["Email"] = email
        row["Company"] = company
        row["Address"] = address
        row["City"] = city
        row["State"] = state
        row["ZipCode"] = zip_code
        row["Country"] = country
        row["LeadSource1"] = "Email"
        row["LeadComments"] = comments
        row["PhoneSupplied"] = phone
        row["PDF"] = pdf
        row["WebAddress"] = website

        excel_bytes = make_excel(row)

        st.download_button(
            "Download Reviewed Excel",
            excel_bytes,
            file_name="hydac_reviewed_lead.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if parsed["valid_attachments"]:
            zip_bytes = make_attachment_zip(parsed["valid_attachments"])
            st.download_button(
                "Download Valid Attachments ZIP",
                zip_bytes,
                file_name="valid_attachments.zip",
                mime="application/zip",
            )

        with st.expander("Agent Evidence: focused external/customer block"):
            st.text(parsed["external_focus_text"][:12000])

        with st.expander("Full extracted email text"):
            st.text(parsed["body"][:22000])

    except Exception as e:
        st.error("Agent processing failed.")
        st.exception(e)
