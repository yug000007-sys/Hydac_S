import re
import tempfile
from io import BytesIO
from pathlib import Path
import zipfile

import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import Workbook

st.set_page_config(page_title="HYDAC Lead Processor", layout="wide")

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

SIGNATURE_IMAGE_NAMES = {
    "image.png", "image.jpg", "image.jpeg", "image001.png", "image001.jpg",
    "image002.png", "image002.jpg", "image003.png", "image003.jpg",
    "image004.png", "image004.jpg", "image005.png", "image005.jpg",
    "logo.png", "logo.jpg", "facebook.png", "linkedin.png", "twitter.png",
    "instagram.png", "youtube.png", "banner.png", "banner.jpg"
}

VALID_ATTACHMENT_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv",
    ".step", ".stp", ".igs", ".iges", ".dwg", ".dxf",
    ".zip", ".rar", ".7z", ".png", ".jpg", ".jpeg", ".tif", ".tiff"
}


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\r", "\n")
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def html_to_text(html: str) -> str:
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")
    return clean_text(soup.get_text("\n"))


def normalize_phone(phone: str) -> str:
    if not phone:
        return ""
    phone = phone.strip()
    phone = phone.replace("+", "")
    phone = re.sub(r"[().\s]+", "", phone)
    phone = phone.replace("--", "-")
    return phone


def get_name_parts(full_name: str):
    full_name = clean_text(full_name)
    if not full_name:
        return "", ""
    full_name = re.sub(r"<.*?>", "", full_name).strip()
    parts = full_name.split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def extract_email(text: str) -> str:
    emails = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text)
    ignored_domains = ["hydac"]
    for email in emails:
        if not any(domain in email.lower() for domain in ignored_domains):
            return email
    return emails[0] if emails else ""


def extract_website(text: str) -> str:
    sites = re.findall(r"(?:www\.|https?://)[A-Za-z0-9.\-]+\.[A-Za-z]{2,}(?:/[^\s]*)?", text)
    if not sites:
        return ""
    site = sites[0].rstrip(".,;)")
    site = site.replace("https://", "").replace("http://", "")
    return site


def extract_phones(text: str):
    candidates = re.findall(r"(?:\+?\d[\d\s().\-]{7,}\d)", text)
    cleaned = []
    for c in candidates:
        n = normalize_phone(c)
        digits = re.sub(r"\D", "", n)
        if len(digits) >= 8 and n not in cleaned:
            cleaned.append(n)
    return cleaned[:3]


def extract_customer_request(text: str) -> str:
    # Prefer text before signature/contact block and before forwarded history.
    stop_markers = [
        "\nBest regards", "\nRegards", "\nMit freundlichen", "\nThanks",
        "\nThank you", "\nFrom:", "\nSent:", "\nVon:", "\nGesendet:"
    ]
    body = text
    for marker in stop_markers:
        idx = body.find(marker)
        if idx > 40:
            body = body[:idx]
            break

    # Remove greeting lines if present.
    lines = [l.strip() for l in body.splitlines() if l.strip()]
    while lines and re.match(r"^(hi|hello|dear|good morning|good afternoon|guten tag)\b", lines[0], re.I):
        lines.pop(0)

    return clean_text("\n\n".join(lines[:8]))


def looks_like_valid_attachment(filename: str) -> bool:
    if not filename:
        return False

    name = Path(filename).name
    low = name.lower()
    ext = Path(low).suffix

    if ext not in VALID_ATTACHMENT_EXTENSIONS:
        return False

    # Generic signature images are ignored.
    if low in SIGNATURE_IMAGE_NAMES:
        return False

    # Generic image### pattern is normally signature/footer.
    if re.fullmatch(r"image\d{0,3}\.(png|jpg|jpeg|gif)", low):
        return False

    # Keep non-image document/CAD/zip files.
    if ext not in {".png", ".jpg", ".jpeg", ".gif"}:
        return True

    # Keep image attachments only when filename looks specific to request/product.
    # Examples: BIERI 3999534.png, pump_label.jpg, drawing_revA.png
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
    msg_sender = msg.sender or ""
    msg_subject = msg.subject or uploaded_file.name
    msg_date = str(msg.date) if msg.date else ""

    body = clean_text(msg.body or "")
    if not body and getattr(msg, "htmlBody", None):
        body = html_to_text(msg.htmlBody)

    full_text = clean_text("\n".join([msg_sender, msg_subject, body]))

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
        "subject": msg_subject,
        "date": msg_date,
        "sender": msg_sender,
        "body": body,
        "full_text": full_text,
        "valid_attachments": valid_attachments,
        "ignored_attachments": ignored_attachments,
        "msg": msg,
    }


def simple_extract_fields(parsed, uploaded_name):
    text = parsed["full_text"]
    body = parsed["body"]

    email = extract_email(text)
    website = extract_website(text)
    phones = extract_phones(text)

    # Try name from sender first.
    sender = parsed["sender"]
    sender_name = sender.split("<")[0].strip().strip('"')
    first, last = get_name_parts(sender_name)

    # If sender is HYDAC/internal, look for first email owner near body.
    if "hydac" in email.lower() or not first:
        first, last = "", ""

    row = {h: "" for h in HEADER}
    row["ReceivedDateTime"] = parsed["date"]
    row["FirstName"] = first
    row["LastName"] = last
    row["Email"] = email
    row["LeadSource1"] = "Email"
    row["LeadComments"] = extract_customer_request(body)
    row["PhoneSupplied"] = " : ".join(phones)
    row["WebAddress"] = website

    pdf_name = uploaded_name.rsplit(".", 1)[0] + ".pdf"
    attachment_names = [name for name, _ in parsed["valid_attachments"]]
    if attachment_names:
        row["PDF"] = pdf_name + "; Attachments: " + ", ".join(attachment_names)
    else:
        row["PDF"] = pdf_name

    return row


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


st.title("HYDAC Lead Processor")
st.success("Step 3 running: MSG parsing + attachment filtering enabled.")

uploaded_file = st.file_uploader("Upload MSG file", type=["msg"])

if uploaded_file:
    try:
        parsed = parse_msg(uploaded_file)
        row = simple_extract_fields(parsed, uploaded_file.name)

        st.subheader("Extracted Lead Row")
        st.json(row)

        st.subheader("Attachment Classification")
        st.write("Valid customer attachments:")
        st.write([name for name, _ in parsed["valid_attachments"]] or "None")

        st.write("Ignored signature/generic attachments:")
        st.write(parsed["ignored_attachments"] or "None")

        excel_bytes = make_excel(row)
        st.download_button(
            "Download Excel",
            excel_bytes,
            file_name="hydac_lead_output.xlsx",
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

        with st.expander("View extracted email text"):
            st.text(parsed["body"][:10000])

    except Exception as e:
        st.error("Processing failed.")
        st.exception(e)
