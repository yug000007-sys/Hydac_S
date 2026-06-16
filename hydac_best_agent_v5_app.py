import json
import re
import tempfile
import zipfile
from dataclasses import dataclass, asdict
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse
from typing import Dict, List, Optional, Tuple

import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import Workbook


st.set_page_config(page_title="HYDAC Best Agent V5", layout="wide")


HEADER = [
    "Referral", "Brand", "Product", "ReceivedDateTime", "FirstName", "LastName",
    "ContactTitle", "Email", "Company", "Address", "County", "City", "State",
    "ZipCode", "Country", "LeadSource1", "LeadSource2", "LeadSource3",
    "LeadComments", "PhoneSupplied", "PhSuppliedExtension", "PhoneResearched",
    "CSRName", "PDF", "DUNS", "WebAddress", "Linkedin_Title", "Linkedin_Link",
    "SIC", "NAICS", "noOfEmployees", "ParentName", "LineOfBusiness", "PQ",
    "Latitude", "Longitude", "DemoLead", "ScreenReason", "about_me", "college_1",
    "college_1_degree", "college_1_start", "college_1_end", "college_2",
    "college_2_degree", "college_2_start", "college_2_end", "month_of_joining",
    "about_experience", "searched_on_google", "linkedin_city", "linkedin_state",
    "linkedin_country"
]

INTERNAL_DOMAINS = {
    "hydacusa.com",
    "hydac.com",
    "hydac-interlynx.com",
}

INTERNAL_KEYWORDS = [
    "hydac", "hydac technology", "hydac international", "hydac usa",
    "brandon",  # learned example: Brandon signature/forwarder is not the customer
]

BAD_WEBSITE_PATTERNS = [
    "aka.ms", "microsoft.com", "office.com", "outlook.com", "safelinks",
    "proofpoint", "mimecast", "teams.microsoft", "sharepoint.com", "avast.com", "exclaimer.net"
]

FREE_EMAIL_DOMAINS = {
    "gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "aol.com",
    "icloud.com", "me.com", "live.com", "msn.com", "protonmail.com"
}

SIGNATURE_IMAGE_NAMES = {
    "image.png", "image.jpg", "image.jpeg", "image.gif",
    "image001.png", "image001.jpg", "image001.jpeg", "image001.gif",
    "image002.png", "image002.jpg", "image002.jpeg", "image002.gif",
    "image003.png", "image003.jpg", "image003.jpeg", "image003.gif",
    "image004.png", "image004.jpg", "image004.jpeg", "image004.gif",
    "image005.png", "image005.jpg", "image005.jpeg", "image005.gif",
    "image006.png", "image006.jpg", "image006.jpeg", "image006.gif",
    "image007.png", "image007.jpg", "image007.jpeg", "image007.gif",
    "image008.png", "image008.jpg", "image008.jpeg", "image008.gif",
    "image009.png", "image009.jpg", "image009.jpeg", "image009.gif",
    "image010.png", "image010.jpg", "image010.jpeg", "image010.gif",
    "logo.png", "logo.jpg", "logo.jpeg", "banner.png", "banner.jpg", "banner.jpeg",
    "facebook.png", "linkedin.png", "twitter.png", "instagram.png", "youtube.png",
}

VALID_ATTACHMENT_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv",
    ".step", ".stp", ".igs", ".iges", ".dwg", ".dxf",
    ".zip", ".rar", ".7z", ".png", ".jpg", ".jpeg", ".tif", ".tiff",
}

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".tif", ".tiff"}

REQUEST_STOP_PATTERNS = [
    r"^best regards", r"^kind regards", r"^regards", r"^thanks", r"^thank you",
    r"^sincerely", r"^mit freundlichen", r"^freundliche", r"^sent from my",
    r"^confidential", r"^disclaimer", r"^this e-mail", r"^this email",
    r"^e-?maily\s+z\s+adresy", r"you don't often receive email from",
    r"learnaboutsenderidentification", r"aka\.ms", r"microsoft",
    r"^from:", r"^sent:", r"^to:", r"^subject:", r"^-{2,}\s*original message",
]

COMPANY_SUFFIXES = [
    "inc", "inc.", "llc", "ltd", "ltd.", "corp", "corp.", "corporation", "company",
    "co.", "gmbh", "ag", "kg", "s.a.", "sa", "bv", "nv", "plc", "limited",
    "international", "industries", "systems", "technologies", "technology",
]

SECURITY_NOISE_PATTERNS = [
    r"(?i)e-?maily\s+z\s+adresy",
    r"(?i)nedost[áa]v[áa]te\s+\w*\s*často",
    r"(?i)you don't often receive email from",
    r"(?i)you do not often receive email from",
    r"(?i)learnaboutsenderidentification",
    r"(?i)aka\.ms",
    r"(?i)microsoft",
    r"(?i)external email",
    r"(?i)caution",
    r"(?i)sender identification",
    r"(?i)this message originated outside",
    r"(?i)this email originated outside",
]

COUNTRY_HINTS = {
    "usa": "USA", "u.s.a.": "USA", "united states": "USA",
    "czech republic": "Czech Republic", "cz": "Czech Republic",
    "germany": "Germany", "de": "Germany",
    "austria": "Austria", "at": "Austria",
}

US_STATE_NAMES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR", "california": "CA",
    "colorado": "CO", "connecticut": "CT", "delaware": "DE", "florida": "FL", "georgia": "GA",
    "hawaii": "HI", "idaho": "ID", "illinois": "IL", "indiana": "IN", "iowa": "IA",
    "kansas": "KS", "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD",
    "massachusetts": "MA", "michigan": "MI", "minnesota": "MN", "mississippi": "MS", "missouri": "MO",
    "montana": "MT", "nebraska": "NE", "nevada": "NV", "new hampshire": "NH", "new jersey": "NJ",
    "new mexico": "NM", "new york": "NY", "north carolina": "NC", "north dakota": "ND", "ohio": "OH",
    "oklahoma": "OK", "oregon": "OR", "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC",
    "south dakota": "SD", "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT",
    "virginia": "VA", "washington": "WA", "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
}

def is_noise_line(line: str) -> bool:
    line = line or ""
    if not line.strip():
        return True
    return any(re.search(pattern, line) for pattern in SECURITY_NOISE_PATTERNS)

PHONE_LABEL_RE = re.compile(
    r"(?i)(?:tel|telephone|phone|mobile|cell|direct|office|o|m|t|p)[:.\s-]*"
    r"(\+?\d[\d\s().\-/]{6,}\d)"
)

PHONE_GENERIC_RE = re.compile(r"(?<!\w)(\+?\d[\d\s().\-/]{7,}\d)(?!\w)")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
URL_RE = re.compile(r"(?i)\b(?:https?://|www\.)[^\s<>]+|\b[A-Za-z0-9.-]+\.(?:com|net|org|de|at|ca|in|co|io|cz|cy|us)\b")


@dataclass
class MessageBlock:
    index: int
    header: str
    body: str
    sender_name: str = ""
    sender_email: str = ""
    emails: List[str] = None
    phones: List[str] = None
    is_external_marker: bool = False
    is_internal: bool = False
    score: int = 0
    reasons: List[str] = None


@dataclass
class AttachmentDecision:
    filename: str
    decision: str
    reason: str
    attachment: object = None


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


def normalize_email(email: str) -> str:
    return (email or "").strip().strip("<>()[]{}.,;:'\"").lower()


def email_domain(email: str) -> str:
    email = normalize_email(email)
    return email.split("@", 1)[1] if "@" in email else ""


def is_internal_email(email: str) -> bool:
    domain = email_domain(email)
    return any(domain == d or domain.endswith("." + d) for d in INTERNAL_DOMAINS)


def is_internal_text(text: str) -> bool:
    low = (text or "").lower()
    if any(k in low for k in INTERNAL_KEYWORDS):
        return True
    return any(d in low for d in INTERNAL_DOMAINS)


def extract_emails(text: str) -> List[str]:
    found = [normalize_email(e) for e in EMAIL_RE.findall(text or "")]
    unique = []
    for email in found:
        if email and email not in unique:
            unique.append(email)
    return unique


def normalize_phone(phone: str) -> str:
    phone = (phone or "").strip()
    phone = phone.replace("+", "")
    phone = re.sub(r"[().\s]+", "", phone)
    phone = phone.replace("/", "-")
    phone = re.sub(r"-+", "-", phone)
    return phone.strip("-")



def extract_phones(text: str) -> List[str]:
    if not text:
        return []
    candidates = []
    # Prefer line-context extraction to avoid tracking IDs and product numbers.
    lines = (text or "").splitlines()
    for i, line in enumerate(lines):
        low = line.lower()
        prev = lines[i - 1].lower() if i else ""
        if "fax" in low:
            continue
        labelled = bool(re.search(r"(?i)\b(tel|telephone|phone|mobile|cell|direct|office|o|m|t|p)\s*[:.]", line))
        icon_labelled = bool(re.search(r"(?i)(phone|mobile|tel)\.(png|jpg|gif|svg)", prev + " " + low))
        if labelled or icon_labelled:
            for match in PHONE_GENERIC_RE.findall(line):
                candidates.append(match)
    # Fallback to explicit label regex only, not every random long number.
    for match in PHONE_LABEL_RE.findall(text):
        candidates.append(match)

    unique = []
    for phone in candidates:
        phone = normalize_phone(phone)
        digits = re.sub(r"\D", "", phone)
        if len(digits) < 7 or len(digits) > 18:
            continue
        # Reject obvious product/order numbers masquerading as phones.
        if re.fullmatch(r"20\d{5,7}", digits):
            continue
        if phone not in unique:
            unique.append(phone)
    return unique

def extract_external_focus(text: str) -> str:
    markers = ["EXTERNAL EMAIL", "External Email", "CAUTION:", "Caution:"]
    positions = [text.find(m) for m in markers if text.find(m) != -1]
    return text[min(positions):] if positions else text


def parse_sender_from_header(header: str) -> Tuple[str, str]:
    if not header:
        return "", ""
    email_match = EMAIL_RE.search(header)
    email = normalize_email(email_match.group(0)) if email_match else ""

    name = ""
    for line in header.splitlines():
        if re.match(r"(?i)^\s*from\s*:", line):
            value = re.sub(r"(?i)^\s*from\s*:\s*", "", line).strip()
            value = EMAIL_RE.sub("", value)
            value = value.replace("<", "").replace(">", "").strip(" -;,\t")
            name = value
            break
    return name, email


def split_email_thread(text: str) -> List[MessageBlock]:
    text = clean_text(text)
    if not text:
        return []

    # Split on common forwarded/original-message boundaries while preserving the boundary as header context.
    boundary = re.compile(
        r"(?im)(?=^\s*(?:-{2,}\s*)?(?:original message|forwarded message|from\s*:|sent\s*:|to\s*:|subject\s*:))"
    )
    raw_parts = [p.strip() for p in boundary.split(text) if p.strip()]

    # Merge tiny metadata-only fragments into the following/previous fragment.
    merged = []
    buffer = ""
    for part in raw_parts:
        if len(part) < 80 and re.search(r"(?im)^\s*(from|sent|to|subject)\s*:", part):
            buffer = (buffer + "\n" + part).strip()
            continue
        merged.append((buffer + "\n" + part).strip() if buffer else part)
        buffer = ""
    if buffer:
        merged.append(buffer)

    if not merged:
        merged = [text]

    blocks = []
    for idx, part in enumerate(merged):
        lines = part.splitlines()
        header_lines = []
        body_lines = []
        header_mode = True
        for line in lines:
            if header_mode and re.match(r"(?i)^\s*(from|sent|to|cc|subject|date)\s*:", line):
                header_lines.append(line)
            else:
                header_mode = False
                body_lines.append(line)
        header = "\n".join(header_lines)
        body = clean_text("\n".join(body_lines)) or part
        sender_name, sender_email = parse_sender_from_header(header or part[:500])
        emails = extract_emails(part)
        phones = extract_phones(part)
        is_external_marker = bool(re.search(r"(?i)external email|caution:", part[:1200]))
        is_internal = is_internal_email(sender_email) or is_internal_text(header[:1000])
        blocks.append(MessageBlock(
            index=idx,
            header=header,
            body=body,
            sender_name=sender_name,
            sender_email=sender_email,
            emails=emails,
            phones=phones,
            is_external_marker=is_external_marker,
            is_internal=is_internal,
            score=0,
            reasons=[],
        ))
    return blocks


def name_from_email(email: str) -> Tuple[str, str]:
    local = normalize_email(email).split("@", 1)[0]
    local = re.sub(r"[._\-]+", " ", local).strip()
    parts = [p for p in local.title().split() if p]
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])



def split_person_name(name: str, email: str = "") -> Tuple[str, str]:
    """Split names safely, including Outlook formats like 'Morton, Ken'."""
    name = clean_text(name)
    name = re.sub(r"[\"'<>]", "", name)
    name = re.sub(r"\([^)]*\)", "", name).strip(" -;,\t")
    if not name and email:
        return name_from_email(email)
    if "@" in name:
        return name_from_email(name)
    if "," in name:
        left, right = [x.strip() for x in name.split(",", 1)]
        if left and right:
            return right, left
    parts = [p for p in name.split() if p]
    if not parts and email:
        return name_from_email(email)
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def is_tracking_or_asset_url(url: str) -> bool:
    low = (url or "").lower()
    if any(bad in low for bad in BAD_WEBSITE_PATTERNS):
        return True
    if re.search(r"\.(png|jpg|jpeg|gif|svg|webp)(?:\?|$)", low):
        return True
    if re.search(r"/(imagesns|cms|logo|border|phone|mobile|email|location|social|facebook|linkedin|youtube|twitter|instagram)/", low):
        return True
    if "sig-email" in low or "signature" in low:
        return True
    return False


def decode_redirect_url(url: str) -> str:
    """Decode common signature redirectors but only return the embedded real URL."""
    raw = (url or "").strip("<>.,);] ")
    try:
        parsed = urlparse(raw if re.match(r"(?i)^https?://", raw) else "http://" + raw)
        qs = parse_qs(parsed.query)
        for key in ("url", "u", "target", "redirect"):
            if key in qs and qs[key]:
                return unquote(qs[key][0])
    except Exception:
        pass
    return raw


def pretty_company_from_domain(domain: str) -> str:
    domain = (domain or "").lower()
    root = domain.split(".", 1)[0]
    if root in {"gmail", "yahoo", "hotmail", "outlook", "icloud", "aol", "live", "msn"}:
        return ""
    known = {
        "fleetpride": "FleetPride",
        "mfcp": "MFCP",
        "mfcpinc": "MFCP",
        "andritz": "ANDRITZ AG",
        "chropynska": "Chropynska US",
        "cytanet": "",
    }
    return known.get(root, "")


def signature_lines(block_text: str) -> List[str]:
    lines = [l.strip(" |\t") for l in (block_text or "").splitlines() if l.strip()]
    start = 0
    for i, line in enumerate(lines):
        if re.search(r"(?i)^(best regards|kind regards|regards|sincerely|mfG|s pozdravem|mit freundlichen)", line):
            start = i + 1
            break
    return lines[start:start + 80]


def extract_signature_name(block_text: str, selected_email: str = "", header_name: str = "") -> str:
    first, last = split_person_name(header_name, selected_email)
    header_full = clean_text(" ".join([first, last]))
    if first and last and not is_internal_text(header_full):
        return header_full
    sig_lines = signature_lines(block_text)
    local_first, local_last = name_from_email(selected_email) if selected_email else ("", "")
    for line in sig_lines[:30]:
        cand = line.strip(" ,;|\t")
        if not cand or is_noise_line(cand) or EMAIL_RE.search(cand) or URL_RE.search(cand) or PHONE_GENERIC_RE.search(cand):
            continue
        if re.search(r"(?i)^(manager|engineer|buyer|purchasing|director|sales|systems designer|automation)$", cand):
            continue
        if re.search(r"(?i)^(good morning|hello|hi|dear)\b", cand):
            continue
        if re.search(r"(?i)(inc|llc|ltd|gmbh|ag|marine|technologies|technology|company|corporation|fleetpride|mfcp|andritz|chropynska)", cand):
            continue
        words = re.findall(r"[A-Za-zÀ-ž'-]+", cand)
        if 2 <= len(words) <= 4:
            if local_first and local_first.lower() in cand.lower():
                return cand.title() if cand.isupper() else cand
            if header_full and (header_full.split()[0].lower() in cand.lower() or header_full.lower() in cand.lower()):
                return cand.title() if cand.isupper() else cand
            # after a regards marker, a two-word personal line is often the sender's full name
            return cand.title() if cand.isupper() else cand
    return header_full



def extract_company_from_block(block_text: str, selected_email: str = "") -> str:
    lines = [l.strip(" |\t") for l in (block_text or "").splitlines() if l.strip()]
    sig = signature_lines(block_text)
    domain = email_domain(selected_email)
    domain_root = re.sub(r"[^a-z0-9]", "", domain.split(".", 1)[0].lower()) if domain else ""

    title_words = r"manager|engineer|buyer|purchasing|director|president|sales|maintenance|supervisor|specialist|procurement|designer|automation"

    def clean_company_candidate(line: str) -> str:
        line = re.sub(r"(?i)^(company|organization|firm)\s*[:\-]\s*", "", line).strip(" ,;-")
        return line

    def valid_company_line(line: str) -> bool:
        if not line or is_noise_line(line) or is_internal_text(line):
            return False
        if EMAIL_RE.search(line) or PHONE_GENERIC_RE.search(line) or URL_RE.search(line):
            return False
        if len(line) > 90 or len(line) < 2:
            return False
        if not re.search(r"[A-Za-zÀ-ž]", line) or re.fullmatch(r"[\W_]+", line):
            return False
        if re.search(r"(?i)^(from|sent|to|subject|tel|phone|mobile|fax|www|http)\b", line):
            return False
        if re.search(r"[_%=&]{3,}|tenantid|templateid|contentid|boundary|dkim|signature|received|postfix|amavis", line, re.I):
            return False
        if re.search(rf"(?i)^({title_words})$", line.strip()):
            return False
        if re.search(r"(?i)quote|quotation|request|project|filter|pump|qty|quantity|micron|flow rate|looking for|please|can you|good morning", line):
            return False
        # avoid personal names: two normal words with no company/domain signal
        if re.fullmatch(r"[A-Z][a-zÀ-ž'-]+\s+[A-Z][a-zÀ-ž'-]+", line.strip()):
            return False
        return True

    candidate_lines = sig + lines[:120]

    for line in candidate_lines:
        if is_noise_line(line):
            continue
        m = re.search(r"(?i)\b(?:company|organization|firm)\s*[:\-]\s*(.+)$", line)
        if m:
            cand = clean_company_candidate(m.group(1))
            if valid_company_line(cand):
                return cand

    for line in candidate_lines:
        cand = clean_company_candidate(line)
        low = cand.lower().strip()
        if not valid_company_line(cand):
            continue
        words = re.findall(r"[A-Za-z&.\-]+", low)
        if any(sfx in words or low.endswith(" " + sfx) for sfx in COMPANY_SUFFIXES):
            return cand

    if domain_root and domain_root not in {"gmail", "yahoo", "hotmail", "outlook", "aol", "icloud", "cytanet"}:
        for line in candidate_lines:
            cand = clean_company_candidate(line)
            cand_root = re.sub(r"[^a-z0-9]", "", cand.lower())
            if valid_company_line(cand) and (domain_root in cand_root or cand_root in domain_root):
                pretty = pretty_company_from_domain(domain)
                if pretty and (len(cand) < 8 or cand.lower().startswith(("chropy", "fleet", "mfcp"))):
                    return pretty
                return cand

    # Website/domain written in signature can support a conservative known-company fallback.
    pretty = pretty_company_from_domain(domain)
    if pretty:
        return pretty
    return ""


def extract_title_from_block(block_text: str) -> str:
    title_words = ["manager", "engineer", "buyer", "purchasing", "director", "president", "sales", "maintenance", "supervisor", "specialist", "procurement", "designer", "automation"]
    for line in signature_lines(block_text)[:40]:
        clean = line.strip(" |\t")
        low = clean.lower()
        if not (2 <= len(clean) <= 80):
            continue
        if is_noise_line(clean) or is_internal_text(clean) or EMAIL_RE.search(clean) or PHONE_GENERIC_RE.search(clean) or URL_RE.search(clean):
            continue
        if re.search(r"(?i)can you|please|request|quote|project|customer|requiring|looking for", clean):
            continue
        if any(w in low for w in title_words):
            return clean.title() if clean.isupper() else clean
    return ""


def extract_website_from_block(block_text: str) -> str:
    """Return a real customer website only; reject tracking/signature/image URLs."""
    for line in (block_text or "").splitlines():
        if is_noise_line(line):
            continue
        # If line is only an email, do not treat the email domain as a website.
        if EMAIL_RE.search(line) and not re.search(r"(?i)https?://|www\.", line):
            continue
        for raw in URL_RE.findall(line):
            decoded = decode_redirect_url(raw)
            clean = decoded.strip("<>.,);] ")
            if not clean:
                continue
            host = clean.lower().removeprefix("http://").removeprefix("https://").removeprefix("www.").split("/", 1)[0]
            if any(domain in host for domain in INTERNAL_DOMAINS):
                continue
            if host in FREE_EMAIL_DOMAINS:
                continue
            if "@" in clean:
                continue
            if is_tracking_or_asset_url(clean):
                continue
            return clean
    return ""


def sanitize_website(value: str, selected_email: str = "") -> str:
    """Post-process website values from deterministic logic or AI cleanup."""
    value = decode_redirect_url(value or "").strip("<>.,);] ")
    if not value:
        return ""
    low_host = value.lower().removeprefix("http://").removeprefix("https://").removeprefix("www.").split("/", 1)[0]
    if low_host in FREE_EMAIL_DOMAINS:
        return ""
    if selected_email and low_host == email_domain(selected_email):
        return ""
    if any(domain in low_host for domain in INTERNAL_DOMAINS):
        return ""
    if is_tracking_or_asset_url(value):
        return ""
    if re.search(r"\.(png|jpg|jpeg|gif|svg|webp)$", value.lower()):
        return ""
    return value

def extract_address_parts(block_text: str) -> Dict[str, str]:
    result = {"Address": "", "City": "", "State": "", "ZipCode": "", "Country": ""}
    lines = [l.strip(" |\t") for l in (block_text or "").splitlines() if l.strip()]

    def parse_city_state_zip(city_line: str):
        city_line = city_line.strip(" ,;")
        m = re.search(r"^(.+?),?\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$", city_line)
        if m:
            return m.group(1).strip(" ,"), m.group(2), m.group(3)
        m = re.search(r"^(.+?),\s*([A-Za-z ]+)\s+(\d{5}(?:-\d{4})?)$", city_line)
        if m:
            state_raw = m.group(2).strip().lower()
            return m.group(1).strip(" ,"), US_STATE_NAMES.get(state_raw, m.group(2).strip()), m.group(3)
        m = re.search(r"^(\d{4,6})\s+(.+)$", city_line)
        if m:
            return m.group(2).strip(" ,"), "", m.group(1)
        return "", "", ""

    for i, line in enumerate(lines[:120]):
        if is_noise_line(line) or is_internal_text(line) or EMAIL_RE.search(line):
            continue
        if re.search(r"\b\d{1,6}\s+[A-Za-z0-9 .'-]+\b(?:street|st\.?|road|rd\.?|ave|avenue|drive|dr\.?|lane|ln\.?|blvd|way|parkway|pkwy|platz|strasse|straße|square)\b", line, re.I):
            result["Address"] = line.strip(" ,;")
            for j in range(i + 1, min(i + 4, len(lines))):
                if is_noise_line(lines[j]) or EMAIL_RE.search(lines[j]):
                    continue
                city, state, zip_code = parse_city_state_zip(lines[j])
                if city or zip_code:
                    result["City"], result["State"], result["ZipCode"] = city, state, zip_code
                    break
            break

    for line in lines[:120]:
        low = line.lower().strip(" ,;")
        if low in COUNTRY_HINTS:
            result["Country"] = COUNTRY_HINTS[low]
            break
    if result["State"] in US_STATE_NAMES.values() and not result["Country"]:
        result["Country"] = "USA"
    return result


def score_blocks(blocks: List[MessageBlock]) -> List[MessageBlock]:
    for block in blocks:
        score = 0
        reasons = []
        external_emails = [e for e in block.emails if not is_internal_email(e)]
        if block.is_external_marker:
            score += 100
            reasons.append("Inside/near EXTERNAL EMAIL or caution block")
        if external_emails:
            score += 50
            reasons.append("Contains external email address")
        if block.sender_email and not is_internal_email(block.sender_email):
            score += 70
            reasons.append("Sender is external")
        if block.phones:
            score += 25
            reasons.append("Contains phone candidate")
        if re.search(r"(?i)quote|quotation|rfq|request|need|looking for|please|filter|pump|drawing|part|model|serial|price|availability", block.body):
            score += 35
            reasons.append("Contains request/product language")
        if block.is_internal:
            score -= 150
            reasons.append("Rejected/penalized: HYDAC/internal indicators")
        if re.search(r"(?i)confidentiality|disclaimer|microsoft|safelinks|proofpoint", block.body[:1200]):
            score -= 20
            reasons.append("Penalized: security/disclaimer noise")
        block.score = score
        block.reasons = reasons
    return sorted(blocks, key=lambda b: b.score, reverse=True)


def selected_customer_block(blocks: List[MessageBlock]) -> Optional[MessageBlock]:
    scored = score_blocks(blocks)
    for block in scored:
        if block.score > 0 and not block.is_internal and any(not is_internal_email(e) for e in block.emails):
            return block
    for block in scored:
        if any(not is_internal_email(e) for e in block.emails):
            return block
    return scored[0] if scored else None



def phones_owned_by_customer(block: Optional[MessageBlock]) -> List[str]:
    if not block:
        return []
    text = block.body or ""
    if is_internal_text(text[:1000]) and not block.is_external_marker:
        return []
    accepted = []
    lines = text.splitlines()
    for i, line in enumerate(lines):
        low = line.lower()
        prev = lines[i - 1].lower() if i else ""
        if is_internal_text(line) or is_noise_line(line) or "fax" in low:
            continue
        labelled = bool(re.search(r"(?i)\b(tel|telephone|phone|mobile|cell|direct|office|o|m|t|p)\s*[:.]", line))
        icon_labelled = bool(re.search(r"(?i)(phone|mobile|tel)\.(png|jpg|gif|svg)", prev + " " + low))
        if not labelled and not icon_labelled:
            continue
        for raw in PHONE_GENERIC_RE.findall(line):
            phone = normalize_phone(raw)
            digits = re.sub(r"\D", "", phone)
            if len(digits) < 7 or len(digits) > 18:
                continue
            if re.fullmatch(r"20\d{5,7}", digits):
                continue
            if phone not in accepted:
                accepted.append(phone)
    return accepted


def extract_request(block: Optional[MessageBlock]) -> str:
    if not block:
        return ""
    text = block.body
    text = re.sub(r"(?is)^.*?(?:external email|caution:).*?\n", "", text, count=1)
    lines = []
    started = False
    blank_after_start = 0

    request_start_re = re.compile(
        r"(?i)quote|quotation|rfq|request|need|looking for|please|can you|could you|"
        r"price|availability|supply|attached|filter|pump|part|model|serial|drawing|share price"
    )
    continuation_re = re.compile(
        r"(?i)^(?:qty|quantity|model|part|p/?n|pn|serial|s/n|sn|akp|rfbn|rf\d|cy\d|bieri)\b|"
        r"\b(?:qty|quantity)\s*[:#]?\s*\d+\b|"
        r"\b[A-Z]{2,8}[-/ ]?\d[0-9A-Z,./\- ]*\b|"
        r"^(?:flushing unit|return filter|backflush filter|in-line industrial filter|high-pressure)"
    )

    for raw in text.splitlines():
        line = raw.strip(" >\t")
        if not line:
            if started:
                blank_after_start += 1
                if blank_after_start > 2:
                    break
            continue
        if is_noise_line(line) or is_tracking_or_asset_url(line):
            continue
        if any(re.search(p, line, re.I) for p in REQUEST_STOP_PATTERNS):
            if started:
                break
            continue
        if EMAIL_RE.search(line) and len(line) < 100:
            continue
        if re.search(r"(?i)^(www\.|https?://|<https?://)", line):
            continue
        # Do not discard request/product lines that contain order numbers.
        if PHONE_GENERIC_RE.search(line) and len(line) < 80 and not continuation_re.search(line) and not request_start_re.search(line):
            continue

        if request_start_re.search(line):
            started = True
            blank_after_start = 0
            lines.append(line)
        elif started and continuation_re.search(line):
            blank_after_start = 0
            lines.append(line)
        elif started:
            if blank_after_start <= 1 and len(lines) < 10 and not is_internal_text(line):
                blank_after_start = 0
                lines.append(line)
            else:
                break
    if not lines:
        for para in re.split(r"\n\s*\n", text):
            para = clean_text(para)
            if len(para) < 15:
                continue
            if is_internal_text(para) or is_noise_line(para) or is_tracking_or_asset_url(para):
                continue
            if re.search(r"(?i)regards|confidential|from:|sent:|to:|subject:", para[:80]):
                continue
            lines = [l for l in para.splitlines()[:8] if not is_noise_line(l)]
            break
    return clean_text("\n".join(lines))[:1500]


def extract_product(text: str, attachment_names: List[str]) -> str:
    raw_joined = "\n".join([text or "", "\n".join(attachment_names or [])])
    # Remove URLs/tracking lines before model extraction.
    clean_lines = []
    for line in raw_joined.splitlines():
        if is_tracking_or_asset_url(line):
            continue
        line = re.sub(r"https?://\S+|www\.\S+", " ", line)
        line = re.sub(r"<[^>]+>", " ", line)
        if not is_noise_line(line):
            clean_lines.append(line)
    joined = clean_text("\n".join(clean_lines))

    found = []
    def add(value: str):
        value = clean_text(value).strip(" .,:;()")
        if not value:
            return
        if re.search(r"(?i)(content|exclaimer|sig-email|linkedin|youtube|facebook|imagesns|cms|border|location|mobile|logo|url=|tenantid|templateid|icular|boundary)", value):
            return
        if re.search(r"[a-z]", value) and not re.search(r"(?i)(filter|pump|flushing|industrial|high-pressure|in-line)", value):
            return
        if re.fullmatch(r"(?i)till\s+\d+|\d+\s*l/min", value):
            return
        if value not in found:
            found.append(value)

    # Full request/product phrases that contain item plus description.
    joined_one_line = re.sub(r"\s+", " ", joined)
    m = re.search(r"(?i)looking for\s*\(?\d*\)?\s*([0-9]{5,}\s+[A-Z0-9./ \-]+?\s+(?:high-pressure\s+)?in-line industrial filter)", joined_one_line)
    if m:
        add(m.group(1))
    m = re.search(r"(?i)\b(CY\d{6,}\s*/\s*Flushing unit)\b", joined)
    if m:
        add(m.group(1))

    model_patterns = [
        r"\bRF\d+(?:-\d+)?(?:-[A-Z0-9]+)+(?:-[A-Z0-9]+)*\b",
        r"\bRF\d+-\d+\b",
        r"\bRFBN/[A-Z0-9][A-Z0-9\-/.]*\b",
        r"\bAKP\d+[A-Z0-9,./\-]*\b",
        r"\bBIERI\s*\d{4,}\b",
        r"\b(?:part|model|serial|p/n|pn)\s*[:#-]?\s*([A-Z0-9][A-Z0-9\-/.]{3,})\b",
    ]
    for pattern in model_patterns:
        for m in re.finditer(pattern, joined, re.I):
            add(m.group(1) if m.lastindex else m.group(0))
    if found:
        return " / ".join(found[:4])

    if re.search(r"(?i)backflush filter", joined):
        return "Backflush Filter"
    if re.search(r"(?i)flushing unit", joined):
        return "Flushing unit"
    if re.search(r"(?i)hydraulic pump", joined):
        return "Hydraulic Pump"
    if re.search(r"(?i)return filter", joined):
        return "Return Filter"
    return ""

def decide_attachment(filename: str, attachment_obj=None) -> AttachmentDecision:
    if not filename:
        return AttachmentDecision(filename="", decision="Reject", reason="Blank attachment name", attachment=attachment_obj)
    name = Path(filename).name
    low = name.lower()
    ext = Path(low).suffix
    if ext not in VALID_ATTACHMENT_EXTENSIONS:
        return AttachmentDecision(name, "Reject", f"Unsupported extension {ext or '(none)'}", attachment_obj)
    if low in SIGNATURE_IMAGE_NAMES:
        return AttachmentDecision(name, "Reject", "Signature/logo/social image name", attachment_obj)
    if re.fullmatch(r"image\d{0,3}\.(png|jpg|jpeg|gif|tif|tiff)", low):
        return AttachmentDecision(name, "Reject", "Generic inline signature image", attachment_obj)
    if ext not in IMAGE_EXTENSIONS:
        return AttachmentDecision(name, "Keep", "Document/CAD/spreadsheet/archive customer attachment", attachment_obj)
    if re.search(r"\d{4,}", name):
        return AttachmentDecision(name, "Keep", "Image filename contains part/model number", attachment_obj)
    if re.search(r"(?i)(pump|filter|drawing|label|plate|model|part|bieri|hydraulic|spec|quote|rfbn|serial|nameplate)", name):
        return AttachmentDecision(name, "Keep", "Image filename has product/request keyword", attachment_obj)
    return AttachmentDecision(name, "Reject", "Image does not look like customer evidence", attachment_obj)


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

    full_text = clean_text("\n".join([f"From: {sender}", f"Subject: {subject}", body]))
    focus_text = clean_text(extract_external_focus(full_text))
    blocks = split_email_thread(full_text)

    attachment_decisions = []
    for att in msg.attachments:
        fname = getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or ""
        if fname:
            attachment_decisions.append(decide_attachment(fname, att))

    return {
        "sender": sender,
        "subject": subject,
        "date": date,
        "body": body,
        "full_text": full_text,
        "external_focus_text": focus_text,
        "blocks": blocks,
        "attachment_decisions": attachment_decisions,
    }


def build_reasoning(parsed, uploaded_name: str) -> Dict:
    blocks = parsed["blocks"]
    ranked = score_blocks(blocks)
    selected = selected_customer_block(blocks)
    valid_attachment_names = [d.filename for d in parsed["attachment_decisions"] if d.decision == "Keep"]
    ignored_attachment_names = [d.filename for d in parsed["attachment_decisions"] if d.decision == "Reject"]

    external_emails = []
    if selected:
        external_emails = [e for e in selected.emails if not is_internal_email(e)]
    selected_email = selected.sender_email if selected and selected.sender_email and not is_internal_email(selected.sender_email) else ""
    if not selected_email and external_emails:
        selected_email = external_emails[0]

    sender_name = selected.sender_name if selected else ""
    selected_text = selected.body if selected else ""
    signature_name = extract_signature_name(selected_text, selected_email, sender_name)
    first, last = split_person_name(signature_name, selected_email)

    phones = phones_owned_by_customer(selected)
    lead_comments = extract_request(selected)
    product = extract_product("\n".join([uploaded_name, parsed.get("subject", ""), lead_comments, selected_text]), valid_attachment_names)
    company = extract_company_from_block(selected_text, selected_email)
    title = extract_title_from_block(selected_text)
    website = sanitize_website(extract_website_from_block(selected_text), selected_email)
    address = extract_address_parts(selected_text)

    rejected = []
    for block in ranked:
        if block.is_internal:
            who = block.sender_email or block.sender_name or f"Block {block.index + 1}"
            rejected.append(f"{who}: HYDAC/internal indicators")

    confidence = "Low"
    if selected and selected.score >= 150 and selected_email and lead_comments:
        confidence = "High"
    elif selected and selected.score >= 80 and (selected_email or lead_comments):
        confidence = "Medium"

    evidence = {
        "SelectedCustomerBlock": selected.index + 1 if selected else "",
        "SelectedCustomerEmail": selected_email,
        "SelectedCustomerReasons": selected.reasons if selected else [],
        "SelectedCustomerNameEvidence": signature_name,
        "SelectedCustomerScore": selected.score if selected else 0,
        "RejectedInternalSenders": rejected,
        "PhoneEvidence": phones,
        "RequestEvidence": lead_comments,
        "AttachmentDecisions": [
            {"filename": d.filename, "decision": d.decision, "reason": d.reason}
            for d in parsed["attachment_decisions"]
        ],
        "RankedBlocks": [
            {
                "block": b.index + 1,
                "score": b.score,
                "sender_name": b.sender_name,
                "sender_email": b.sender_email,
                "is_internal": b.is_internal,
                "reasons": b.reasons,
                "preview": clean_text(b.body)[:500],
            }
            for b in ranked[:8]
        ],
    }

    agent_data = {
        "FirstName": first,
        "LastName": last,
        "ContactTitle": title,
        "Email": selected_email,
        "Company": company,
        "Address": address.get("Address", ""),
        "City": address.get("City", ""),
        "State": address.get("State", ""),
        "ZipCode": address.get("ZipCode", ""),
        "Country": address.get("Country", ""),
        "PhoneSupplied": " : ".join(phones),
        "WebAddress": website,
        "LeadComments": lead_comments,
        "Product": product,
        "Confidence": confidence,
        "AgentReason": "; ".join(selected.reasons if selected else ["No selected customer block"]),
        "Evidence": evidence,
        "ValidAttachmentNames": valid_attachment_names,
        "IgnoredAttachmentNames": ignored_attachment_names,
    }
    return agent_data


def run_ai_cleaner(parsed, uploaded_name: str, deterministic_data: Dict) -> Dict:
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        deterministic_data["AgentReason"] += " | AI cleaner not used: OPENAI_API_KEY missing."
        return deterministic_data

    from openai import OpenAI
    client = OpenAI(api_key=api_key)

    selected_block_num = deterministic_data.get("Evidence", {}).get("SelectedCustomerBlock")
    selected_block_text = ""
    for block in parsed["blocks"]:
        if selected_block_num and block.index + 1 == selected_block_num:
            selected_block_text = block.body[:8000]
            break

    prompt = f"""
You are HYDAC Best Agent V5. You are a reviewer, not a blind cleaner. Challenge deterministic fields when they are tracking links, signature images, warning banners, or product numbers misread as phones.

HYDAC RULES:
- Real external customer only, not latest sender.
- Reject HYDAC employees, HYDAC Sales, internal forwarders, HYDAC signatures, and Brandon signature data.
- If EXTERNAL EMAIL exists, prioritize the customer request inside/below that block.
- LeadComments must contain only the actual customer request, not the trail, not greetings, not signatures.
- PhoneSupplied must contain customer-owned phones only, formatted phone1 : phone2.
- Leave unavailable fields blank. Do not guess.
- Company must be written in body/signature. Do not infer company from Gmail/free email.
- Keep BIERI 3999534.png type attachments when they are product/part images.
- Reject tracking URLs such as avast.com/sig-email, exclaimer.net redirects, aka.ms, and image/logo URLs as WebAddress/Product.
- Never use warning banners as company/title.
- Do not use product/order numbers as phone numbers.

DETERMINISTIC PRE-ANALYSIS JSON:
{json.dumps({k: v for k, v in deterministic_data.items() if k != 'Evidence'}, ensure_ascii=False, indent=2)}

EVIDENCE JSON:
{json.dumps(deterministic_data.get('Evidence', {}), ensure_ascii=False, indent=2)}

SELECTED CUSTOMER BLOCK TEXT:
{selected_block_text}

Return ONLY valid JSON with these exact keys:
FirstName, LastName, ContactTitle, Email, Company, Address, City, State, ZipCode, Country,
PhoneSupplied, WebAddress, LeadComments, Product, Confidence, AgentReason
"""
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0,
            messages=[
                {"role": "system", "content": "Return only valid JSON. Never invent unavailable fields."},
                {"role": "user", "content": prompt},
            ],
        )
        content = response.choices[0].message.content.strip()
        content = content.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
        ai_data = json.loads(content)
        required = [
            "FirstName", "LastName", "ContactTitle", "Email", "Company", "Address", "City", "State",
            "ZipCode", "Country", "PhoneSupplied", "WebAddress", "LeadComments", "Product",
            "Confidence", "AgentReason",
        ]
        for key in required:
            ai_data.setdefault(key, deterministic_data.get(key, ""))
        ai_data["WebAddress"] = sanitize_website(ai_data.get("WebAddress", ""), ai_data.get("Email", deterministic_data.get("Email", "")))
        ai_data["Evidence"] = deterministic_data["Evidence"]
        ai_data["ValidAttachmentNames"] = deterministic_data["ValidAttachmentNames"]
        ai_data["IgnoredAttachmentNames"] = deterministic_data["IgnoredAttachmentNames"]
        ai_data["AgentReason"] = f"{ai_data.get('AgentReason', '')} | Deterministic evidence score: {deterministic_data.get('Evidence', {}).get('SelectedCustomerScore', '')}"
        return ai_data
    except Exception as exc:
        deterministic_data["AgentReason"] += f" | AI cleaner failed; deterministic result used. Error: {exc}"
        return deterministic_data


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


def make_attachment_zip(attachment_decisions: List[AttachmentDecision]):
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as z:
        for decision in attachment_decisions:
            if decision.decision != "Keep" or decision.attachment is None:
                continue
            try:
                data = decision.attachment.data
                if data:
                    z.writestr(Path(decision.filename).name, data)
            except Exception:
                pass
    return output.getvalue()


def build_pdf_field(uploaded_name: str, valid_names: List[str]) -> str:
    pdf_base = Path(uploaded_name).with_suffix(".pdf").name
    if valid_names:
        return pdf_base + "; Attachments: " + ", ".join(valid_names)
    return pdf_base


def render_evidence(agent_data: Dict):
    evidence = agent_data.get("Evidence", {})
    st.subheader("Agent Evidence and Confidence")
    c1, c2, c3 = st.columns(3)
    c1.metric("Confidence", agent_data.get("Confidence", ""))
    c2.metric("Selected Block", evidence.get("SelectedCustomerBlock", ""))
    c3.metric("Customer Score", evidence.get("SelectedCustomerScore", ""))

    st.write("**Selected customer reasons:**")
    st.write(evidence.get("SelectedCustomerReasons", []) or "None")

    st.write("**Rejected HYDAC/internal senders:**")
    st.write(evidence.get("RejectedInternalSenders", []) or "None")

    st.write("**Phone evidence accepted:**")
    st.write(evidence.get("PhoneEvidence", []) or "None")

    st.write("**Attachment decisions:**")
    st.dataframe(evidence.get("AttachmentDecisions", []), use_container_width=True)

    with st.expander("Ranked email blocks"):
        st.dataframe(evidence.get("RankedBlocks", []), use_container_width=True)


st.title("HYDAC Best Agent V5")
st.caption("Reasoning-first lead processor: thread parsing, customer ranking, phone ownership, attachment evidence, review, then Excel export.")

with st.expander("Agent status"):
    if st.secrets.get("OPENAI_API_KEY", ""):
        st.success("AI cleaner active: OPENAI_API_KEY detected. Deterministic evidence still controls the customer selection.")
    else:
        st.warning("Deterministic mode active. Add OPENAI_API_KEY only for final field cleanup, not for blind extraction.")

uploaded_file = st.file_uploader("Upload .msg file", type=["msg"])

if uploaded_file:
    try:
        parsed = parse_msg(uploaded_file)
        deterministic_data = build_reasoning(parsed, uploaded_file.name)
        agent_data = run_ai_cleaner(parsed, uploaded_file.name, deterministic_data)

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
        pdf_value = build_pdf_field(uploaded_file.name, agent_data.get("ValidAttachmentNames", []))
        pdf = st.text_input("PDF", value=pdf_value)
        comments = st.text_area("LeadComments / Customer Request", value=agent_data.get("LeadComments", ""), height=180)

        render_evidence(agent_data)

        row = {h: "" for h in HEADER}
        row.update({
            "Product": product,
            "ReceivedDateTime": received,
            "FirstName": first,
            "LastName": last,
            "ContactTitle": title,
            "Email": email,
            "Company": company,
            "Address": address,
            "City": city,
            "State": state,
            "ZipCode": zip_code,
            "Country": country,
            "LeadSource1": "Email",
            "LeadComments": comments,
            "PhoneSupplied": phone,
            "PDF": pdf,
            "WebAddress": website,
        })

        excel_bytes = make_excel(row)
        st.download_button(
            "Download Reviewed Excel",
            excel_bytes,
            file_name="hydac_reviewed_lead.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        kept_decisions = [d for d in parsed["attachment_decisions"] if d.decision == "Keep"]
        if kept_decisions:
            zip_bytes = make_attachment_zip(parsed["attachment_decisions"])
            st.download_button(
                "Download Valid Attachments ZIP",
                zip_bytes,
                file_name="valid_customer_attachments.zip",
                mime="application/zip",
            )

        with st.expander("Selected customer/request block"):
            selected_num = agent_data.get("Evidence", {}).get("SelectedCustomerBlock")
            for block in parsed["blocks"]:
                if block.index + 1 == selected_num:
                    st.text(block.body[:12000])
                    break

        with st.expander("Focused external/customer block"):
            st.text(parsed["external_focus_text"][:12000])

        with st.expander("Full extracted email text"):
            st.text(parsed["body"][:22000])

    except Exception as e:
        st.error("Agent processing failed.")
        st.exception(e)
