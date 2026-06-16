import re
import tempfile
from io import BytesIO
from pathlib import Path
import zipfile

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
    "mimecast",
    "proofpoint",
    "safelinks",
]

SIGNATURE_IMAGE_NAMES = {
    "image.png", "image.jpg", "image.jpeg",
    "image001.png", "image001.jpg", "image001.jpeg",
    "image002.png", "image002.jpg", "image002.jpeg",
    "image003.png", "image003.jpg", "image003.jpeg",
    "image004.png", "image004.jpg", "image004.jpeg",
    "image005.png", "image005.jpg", "image005.jpeg",
    "logo.png", "logo.jpg", "facebook.png", "linkedin.png",
    "twitter.png", "instagram.png", "youtube.png", "banner.png", "banner.jpg"
}

VALID_ATTACHMENT_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv",
    ".step", ".stp", ".igs", ".iges", ".dwg", ".dxf",
    ".zip", ".rar", ".7z", ".png", ".jpg", ".jpeg", ".tif", ".tiff"
}


def clean_text(text):
    if not text:
        return ""
    text = str(text).replace("\r", "\n")
    text = re.sub(r"\xa0", " ", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def html_to_text(html):
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")
    return clean_text(soup.get_text("\n"))


def is_internal_email(email):
    email = (email or "").lower()
    return any(d in email for d in INTERNAL_DOMAINS)


def split_name(name):
    name = clean_text(name)
    name = re.sub(r"<.*?>", "", name).strip(" '\"")
    if not name:
        return "", ""
    parts = name.split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def normalize_phone(phone):
    phone = (phone or "").strip()
    phone = phone.replace("+", "")
    phone = re.sub(r"[().\s]+", "", phone)
    phone = phone.replace("--", "-").strip("-")
    return phone


def extract_emails(text):
    found = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text or "")
    unique = []
    for e in found:
        if e not in unique:
            unique.append(e)
    return unique


def extract_websites(text):
    found = re.findall(r"(?:www\.|https?://)[A-Za-z0-9.\-]+\.[A-Za-z]{2,}(?:/[^\s<>()]*)?", text or "")
    out = []
    for site in found:
        site = site.rstrip(".,;)")
        site = site.replace("https://", "").replace("http://", "")
        low = site.lower()
        if any(bad in low for bad in BAD_WEBSITE_PATTERNS):
            continue
        if site not in out:
            out.append(site)
    return out


def extract_phones(text):
    candidates = re.findall(r"(?:\+?\d[\d\s().\-]{7,}\d)", text or "")
    out = []
    for c in candidates:
        n = normalize_phone(c)
        digits = re.sub(r"\D", "", n)
        if len(digits) >= 8 and n not in out:
            out.append(n)
    return out[:4]


def parse_from_lines(text):
    """
    Reads email trail and returns possible people from From: blocks.
    Prioritizes non-HYDAC emails.
    """
    candidates = []
    for line in (text or "").splitlines():
        line = line.strip()
        m = re.match(r"^(From|Von|De):\s*(.+)$", line, re.I)
        if not m:
            continue
        value = m.group(2).strip()
        emails = extract_emails(value)
        email = emails[0] if emails else ""
        name = value.split("<")[0].strip().strip('"')
        if not name and email:
            name = email.split("@")[0].replace(".", " ").title()
        if email:
            candidates.append({"name": name, "email": email, "source": line})
    return candidates


def infer_contact(parsed):
    text = parsed["full_text"]
    body = parsed["body"]
    candidates = parse_from_lines(text)

    # First choice: non-HYDAC From: candidate in trail.
    for c in candidates:
        if c["email"] and not is_internal_email(c["email"]):
            first, last = split_name(c["name"])
            return {
                "first": first,
                "last": last,
                "email": c["email"],
                "reason": "Original non-HYDAC sender found in email trail.",
                "confidence": "High",
            }

    # Second choice: first non-HYDAC email anywhere.
    for e in extract_emails(text):
        if not is_internal_email(e):
            name_guess = e.split("@")[0].replace(".", " ").replace("_", " ").title()
            first, last = split_name(name_guess)
            return {
                "first": first,
                "last": last,
                "email": e,
                "reason": "First non-HYDAC email found.",
                "confidence": "Medium",
            }

    # Last fallback: message sender.
    sender = parsed["sender"]
    emails = extract_emails(sender)
    email = emails[0] if emails else ""
    name = sender.split("<")[0].strip().strip('"')
    first, last = split_name(name)
    return {
        "first": first,
        "last": last,
        "email": email,
        "reason": "Fallback to latest sender.",
        "confidence": "Low",
    }


def extract_request(body):
    body = clean_text(body)

    # Remove Microsoft/security blocks where possible.
    remove_patterns = [
        r"Learn About Sender Identification.*",
        r"You don't often get email from.*",
        r"CAUTION:.*",
    ]
    for p in remove_patterns:
        body = re.sub(p, "", body, flags=re.I | re.S)

    # Split at common signature/history markers.
    markers = [
        "\nBest regards", "\nRegards", "\nThanks", "\nThank you",
        "\nMit freundlichen", "\nFrom:", "\nSent:", "\nVon:", "\nGesendet:",
    ]

    candidate = body
    for marker in markers:
        idx = candidate.find(marker)
        if idx > 20:
            candidate = candidate[:idx]
            break

    lines = [l.strip() for l in candidate.splitlines() if l.strip()]
    while lines and re.match(r"^(hi|hello|dear|good morning|good afternoon|guten tag|team)\b", lines[0], re.I):
        lines.pop(0)

    request = clean_text("\n\n".join(lines[:10]))

    # If latest body is just forward text, try to find question/request sentence in whole body.
    if len(request) < 20:
        req_lines = []
        for l in body.splitlines():
            if re.search(r"\b(quote|offer|price|availability|lead time|send|request|please|need|model|drawing|pump|filter)\b", l, re.I):
                req_lines.append(l.strip())
        request = clean_text("\n".join(req_lines[:8]))

    return request


def infer_company(text, email):
    domain = ""
    if email and "@" in email:
        domain = email.split("@", 1)[1].lower()

    # Avoid gmail/yahoo etc.
    free_domains = ["gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "icloud.com"]
    if domain and domain not in free_domains and not is_internal_email(email):
        return ""

    # Later this will become AI-assisted/manual reviewed.
    return ""


def looks_like_valid_attachment(filename):
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

    if re.search(r"\d{4,}", name):
        return True

    if re.search(r"(pump|filter|drawing|label|plate|model|part|bieri|hydraulic|spec|quote|rfbn)", low):
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
        "valid_attachments": valid_attachments,
        "ignored_attachments": ignored_attachments,
    }


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


st.title("HYDAC Lead Agent")
st.success("Step 4 running: Agent Review Panel enabled.")

uploaded_file = st.file_uploader("Upload MSG file", type=["msg"])

if uploaded_file:
    try:
        parsed = parse_msg(uploaded_file)
        contact = infer_contact(parsed)
        websites = extract_websites(parsed["full_text"])
        phones = extract_phones(parsed["full_text"])
        request = extract_request(parsed["body"])

        st.subheader("Agent Review Panel")

        col1, col2 = st.columns(2)

        with col1:
            first = st.text_input("FirstName", value=contact["first"])
            last = st.text_input("LastName", value=contact["last"])
            email = st.text_input("Email", value=contact["email"])
            company = st.text_input("Company", value=infer_company(parsed["full_text"], contact["email"]))
            title = st.text_input("ContactTitle", value="")
            phone = st.text_input("PhoneSupplied", value=" : ".join(phones))
            website = st.text_input("WebAddress", value=websites[0] if websites else "")

        with col2:
            received = st.text_input("ReceivedDateTime", value=parsed["date"])
            pdf_base = uploaded_file.name.rsplit(".", 1)[0] + ".pdf"
            valid_names = [name for name, _ in parsed["valid_attachments"]]
            pdf_value = pdf_base + (("; Attachments: " + ", ".join(valid_names)) if valid_names else "")
            pdf = st.text_input("PDF", value=pdf_value)
            lead_source = st.text_input("LeadSource1", value="Email")
            confidence = st.text_input("Agent Confidence", value=contact["confidence"])
            st.caption(contact["reason"])

        comments = st.text_area("LeadComments / Customer Request", value=request, height=180)

        st.subheader("Attachment Decision")
        st.write("Valid customer attachments:")
        st.write(valid_names or "None")
        st.write("Ignored signature/generic attachments:")
        st.write(parsed["ignored_attachments"] or "None")

        row = {h: "" for h in HEADER}
        row["ReceivedDateTime"] = received
        row["FirstName"] = first
        row["LastName"] = last
        row["ContactTitle"] = title
        row["Email"] = email
        row["Company"] = company
        row["LeadSource1"] = lead_source
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

        with st.expander("Agent Evidence: extracted email text"):
            st.text(parsed["body"][:15000])

    except Exception as e:
        st.error("Agent processing failed.")
        st.exception(e)
