import json
import re
import tempfile
import zipfile
from dataclasses import dataclass, asdict
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import Workbook


st.set_page_config(page_title="HYDAC Best Agent V2", layout="wide")


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
    "proofpoint", "mimecast", "teams.microsoft", "sharepoint.com"
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
    "manufacturing", "solutions", "engineering", "group", "holdings", "services",
    "associates", "partners", "enterprises", "division", "global",
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
    r"(?i)(?:tel|telephone|phone|mobile|cell|direct|office|fax|m|t|p)[:.\t -]*"
    r"(\+?\d[\d \t().\-/]{6,}\d)"
)

PHONE_GENERIC_RE = re.compile(r"(?<!\w)(\+?\d[\d \t().\-/]{7,}\d)(?!\w)")
# Part/model numbers that look like 7-digit phone numbers: pure digit strings 6-8 chars
# used to exclude false phone matches
PART_NUMBER_RE = re.compile(r"^\d{6,8}$")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
URL_RE = re.compile(r"(?i)\b(?:https?://|www\.)[^\s<>]+|\b[A-Za-z0-9.-]+\.(?:com|net|org|de|at|ca|in|co|io)\b")


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


# Inline-image icon pattern used by many email clients as field decorators
# e.g. " <https://example.com/images/phone.png>	828-555-1234"
_ICON_PREFIX_RE = re.compile(
    r"^\s*<https?://[^\s>]+?(?:/images?/|/img/|/cms/|/static/|/icons?/|/logo|/media/)[^\s>]*>\s*[	 ]*",
    re.I
)

def strip_icon_prefix(line: str) -> str:
    """Remove leading inline-image icon URLs (common in HTML-signature plain-text renderings)."""
    return _ICON_PREFIX_RE.sub("", line)


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
    # Strip inline icon URL prefixes line by line before scanning
    text = "\n".join(strip_icon_prefix(l) for l in text.splitlines())
    labeled = set()
    candidates = []
    for match in PHONE_LABEL_RE.findall(text):
        n = normalize_phone(match)
        labeled.add(n)
        candidates.append((match, True))
    for match in PHONE_GENERIC_RE.findall(text):
        candidates.append((match, False))
    unique = []
    for phone, is_labeled in candidates:
        phone = normalize_phone(phone)
        digits = re.sub(r"\D", "", phone)
        if len(digits) < 7 or len(digits) > 18:
            continue
        # Reject bare 6-8 digit strings (likely part/model numbers) unless phone-labeled
        if not is_labeled and PART_NUMBER_RE.match(digits):
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
    name = clean_text(name)
    name = re.sub(r"[\"'<>]", "", name)
    name = re.sub(r"\([^)]*\)", "", name).strip(" -;\t")
    if not name and email:
        return name_from_email(email)
    if "@" in name:
        return name_from_email(name)
    # Handle "Last, First" format (common in Outlook/Exchange sender headers)
    if "," in name:
        parts = [p.strip() for p in name.split(",", 1) if p.strip()]
        if len(parts) == 2:
            return parts[1], parts[0]
        name = parts[0] if parts else name
    parts = [p for p in name.split() if p]
    if not parts and email:
        return name_from_email(email)
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def extract_company_from_block(block_text: str, selected_email: str = "") -> str:
    lines = [strip_icon_prefix(l).strip(" |\t") for l in (block_text or "").splitlines() if l.strip()]
    domain = email_domain(selected_email).split(".")[0] if selected_email else ""
    domain_root = re.sub(r"[^a-z0-9]", "", domain.lower())

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
        low = line.lower().strip()
        if re.search(r"(?i)^(from|sent|to|subject|tel|phone|mobile|fax|www)\b", line):
            return False
        if re.search(r"(?i)^(best regards|kind regards|thanks|thank you|automation)$", line):
            return False
        if re.search(r"(?i)quote|quotation|request|project|filter|pump|qty|quantity|micron|flow rate", line):
            return False
        return True

    # 1) Explicit company label wins.
    for line in lines[:120]:
        if is_noise_line(line):
            continue
        m = re.search(r"(?i)\b(?:company|organization|firm)\s*[:\-]\s*(.+)$", line)
        if m:
            cand = clean_company_candidate(m.group(1))
            if valid_company_line(cand):
                return cand

    # 2) Suffix-based company lines.
    for line in lines[:120]:
        cand = clean_company_candidate(line)
        low = cand.lower().strip()
        if not valid_company_line(cand):
            continue
        words = re.findall(r"[A-Za-z&.\-]+", low)
        if any(sfx in words or low.endswith(" " + sfx) for sfx in COMPANY_SUFFIXES):
            return cand

    # 3) Domain-root match in signature text. This catches names such as Chropynska US
    # from jan.schubert@chropynska.cz without ever using free-mail domains.
    if domain_root and domain_root not in {"gmail", "yahoo", "hotmail", "outlook", "aol", "icloud"}:
        for line in lines[:120]:
            cand = clean_company_candidate(line)
            cand_root = re.sub(r"[^a-z0-9]", "", cand.lower())
            if valid_company_line(cand) and domain_root in cand_root:
                return cand

    # 4) Derive company name from email domain as last resort (e.g. fleetpride.com -> FleetPride)
    # Only when domain is not a free-mail provider and no other source found.
    if domain_root and domain_root not in {"gmail", "yahoo", "hotmail", "outlook", "aol", "icloud", "me", "msn", "live", "protonmail"}:
        return domain_root.title()

    return ""


def extract_title_from_block(block_text: str) -> str:
    title_words = ["manager", "engineer", "buyer", "purchasing", "director", "president", "sales", "maintenance", "supervisor", "specialist", "procurement"]
    for line in (block_text or "").splitlines()[:80]:
        clean = line.strip(" |\t")
        low = clean.lower()
        if 2 <= len(clean) <= 80 and any(w in low for w in title_words) and not is_internal_text(clean):
            if not EMAIL_RE.search(clean) and not PHONE_GENERIC_RE.search(clean):
                return clean
    return ""


def extract_website_from_block(block_text: str) -> str:
    """Return a real customer website only. Never turn email domains such as gmail.com into websites."""
    candidates = []
    for raw_line in (block_text or "").splitlines():
        line = strip_icon_prefix(raw_line)
        # If the line contains an email address, URL_RE may match the email domain (for example gmail.com).
        # That is not a company website, so skip the whole line unless it also has an explicit http/www URL.
        if EMAIL_RE.search(line) and not re.search(r"(?i)https?://|www\.", line):
            continue
        for url in URL_RE.findall(line):
            raw = url.strip(".,);]<>")
            low = raw.lower().removeprefix("http://").removeprefix("https://").removeprefix("www.").split("/", 1)[0]
            if any(bad in raw.lower() for bad in BAD_WEBSITE_PATTERNS):
                continue
            if any(domain in low for domain in INTERNAL_DOMAINS):
                continue
            if low in FREE_EMAIL_DOMAINS:
                continue
            if "@" in raw:
                continue
            # Prefer explicit www. or http(s):// mentions — give them higher priority
            is_explicit = bool(re.search(r"(?i)https?://|www\.", url))
            # Skip image/asset paths (they contain /images, /img, /cms, /static etc.)
            is_asset = bool(re.search(r"(?i)/(?:images?|img|cms|static|assets?|icons?|logo|media)/", raw))
            if not is_asset:
                candidates.append((raw, is_explicit))
    # Return first explicit URL, fallback to first non-asset URL
    for url, explicit in candidates:
        if explicit:
            low = url.lower().removeprefix("http://").removeprefix("https://").removeprefix("www.").split("/", 1)[0]
            return low if low else url
    for url, _ in candidates:
        return url
    return ""


def sanitize_website(value: str, selected_email: str = "") -> str:
    """Post-process website values from deterministic logic or AI cleanup."""
    value = (value or "").strip().strip(".,);]")
    if not value:
        return ""
    low_host = value.lower().removeprefix("http://").removeprefix("https://").removeprefix("www.").split("/", 1)[0]
    if low_host in FREE_EMAIL_DOMAINS:
        return ""
    if selected_email and low_host == email_domain(selected_email):
        # Never infer WebAddress from the customer's email domain; it must be explicitly written as a website.
        return ""
    if any(bad in value.lower() for bad in BAD_WEBSITE_PATTERNS):
        return ""
    if any(domain in low_host for domain in INTERNAL_DOMAINS):
        return ""
    return value


def extract_address_parts(block_text: str) -> Dict[str, str]:
    result = {"Address": "", "City": "", "State": "", "ZipCode": "", "Country": ""}
    lines = [strip_icon_prefix(l).strip(" |\t") for l in (block_text or "").splitlines() if l.strip()]

    def parse_city_state_zip(city_line: str):
        city_line = city_line.strip(" ,;")
        # US: City, ST 12345 or City ST 12345
        m = re.search(r"^(.+?),?\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$", city_line)
        if m:
            return m.group(1).strip(" ,"), m.group(2), m.group(3)
        # US: City, State Name 12345
        m = re.search(r"^(.+?),\s*([A-Za-z ]+)\s+(\d{5}(?:-\d{4})?)$", city_line)
        if m:
            state_raw = m.group(2).strip().lower()
            return m.group(1).strip(" ,"), US_STATE_NAMES.get(state_raw, m.group(2).strip()), m.group(3)
        # European: 12345 City  or  1234 AB City (Dutch/German)
        m = re.search(r"^(\d{4,6})\s*(?:[A-Z]{2}\s+)?(.+)$", city_line)
        if m:
            return m.group(2).strip(" ,"), "", m.group(1)
        # City only
        if re.match(r"^[A-Za-z][A-Za-z .'-]{2,}$", city_line) and len(city_line) < 50:
            return city_line.strip(" ,"), "", ""
        return "", "", ""

    ADDRESS_RE = re.compile(
        r"(?i)(?:\b\d{1,6}\s+[A-Za-z0-9 .'\\-]+\b"
        r"(?:street|st\.?|road|rd\.?|avenue|ave\.?|drive|dr\.?|lane|ln\.?"
        r"|boulevard|blvd\.?|way|parkway|pkwy|court|ct\.?|circle|cir\.?"
        r"|platz|strasse|stra\u00dfe|square|sq\.?|trail|trl\.))"
        r"|(?:P\.?O\.?\s*Box\s*\d+)"
        r"|(?:Suite\s+\d+|Ste\.?\s*\d+|Unit\s+\d+|Apt\.?\s*\d+)"
    )

    for i, line in enumerate(lines[:150]):
        if is_noise_line(line) or is_internal_text(line) or EMAIL_RE.search(line):
            continue

        # Handle split address: bare street number on one line, rest on next
        # e.g. "8814" then "Dietz Ave. Hickory NC 28602"
        if re.match(r"^\d{1,6}$", line.strip()):
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                combined = line.strip() + " " + next_line
                if ADDRESS_RE.search(combined):
                    # Also handle split number+street where city/zip is inline on the street line
                    addr_part = combined
                    # Try to split street from city/state/zip on the same line
                    m = re.match(r"^(\d+\s+[A-Za-z0-9 .'\-]+?(?:street|st\.?|road|rd\.?|avenue|ave\.?|drive|dr\.?|lane|ln\.?|boulevard|blvd\.?|way|parkway|pkwy|court|ct\.?|platz|strasse|stra\u00dfe|square|sq\.?))\.?[,\s]+(.+)$", combined, re.I)
                    if m:
                        result["Address"] = m.group(1).strip(" ,;")
                        city, state, zip_code = parse_city_state_zip(m.group(2))
                        if city or zip_code:
                            result["City"], result["State"], result["ZipCode"] = city, state, zip_code
                    else:
                        result["Address"] = addr_part.strip(" ,;")
                    break

        if ADDRESS_RE.search(line):
            # Check if city/state/zip is inline on the same line (e.g. "123 Main St Hickory NC 28602")
            m = re.match(r"^(\d+\s+[A-Za-z0-9 .'\-]+?(?:street|st\.?|road|rd\.?|avenue|ave\.?|drive|dr\.?|lane|ln\.?|boulevard|blvd\.?|way|parkway|pkwy|court|ct\.?|platz|strasse|stra\u00dfe|square|sq\.?))\.?[,\s]+(.+)$", line, re.I)
            if m:
                result["Address"] = m.group(1).strip(" ,;")
                city, state, zip_code = parse_city_state_zip(m.group(2))
                if city or zip_code:
                    result["City"], result["State"], result["ZipCode"] = city, state, zip_code
            else:
                result["Address"] = line.strip(" ,;")
                # Look ahead up to 5 lines for city/state/zip
                for j in range(i + 1, min(i + 6, len(lines))):
                    nxt = lines[j]
                    if is_noise_line(nxt) or EMAIL_RE.search(nxt) or is_internal_text(nxt):
                        continue
                    # Secondary address line (suite/unit) — append to Address
                    if re.match(r"(?i)^(suite|ste|unit|apt|floor|fl)\b", nxt):
                        result["Address"] = result["Address"] + ", " + nxt.strip(" ,;")
                        continue
                    city, state, zip_code = parse_city_state_zip(nxt)
                    if city or zip_code:
                        result["City"], result["State"], result["ZipCode"] = city, state, zip_code
                        break
            break

    # Standalone zip fallback if address wasn't found but zip exists in text
    if not result["ZipCode"]:
        for line in lines[:150]:
            if is_noise_line(line) or is_internal_text(line):
                continue
            m = re.search(r"\b(\d{5}(?:-\d{4})?)\b", line)
            if m and not EMAIL_RE.search(line) and not re.search(r"\b\d{6,}\b", line):
                result["ZipCode"] = m.group(1)
                city, state, _ = parse_city_state_zip(re.sub(re.escape(m.group(1)), "", line))
                if city:
                    result["City"] = city
                if state:
                    result["State"] = state
                break

    COUNTRY_EXTENDED = {
        **COUNTRY_HINTS,
        "canada": "Canada",
        "uk": "UK", "united kingdom": "UK", "england": "UK",
        "france": "France",
        "netherlands": "Netherlands", "holland": "Netherlands",
        "poland": "Poland",
        "slovakia": "Slovakia",
        "hungary": "Hungary",
        "india": "India",
        "mexico": "Mexico",
        "brazil": "Brazil",
    }
    for line in lines[:150]:
        low = line.lower().strip(" ,;")
        if low in COUNTRY_EXTENDED:
            result["Country"] = COUNTRY_EXTENDED[low]
            break
        for hint, country_name in COUNTRY_EXTENDED.items():
            if re.search(r"\b" + re.escape(hint) + r"\b", low):
                result["Country"] = country_name
                break
        if result["Country"]:
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
        if re.search(r"(?i)quote|quotation|rfq|request|need|looking for|please|filter|pump|drawing|part|model|serial|price|availability|inquiry|specification|datasheet|replacement|spare", block.body):
            score += 35
            reasons.append("Contains request/product language")
        if re.search(r"(?i)\b(inc|llc|ltd|corp|gmbh|ag|kg|s\.a\.|bv|plc|co\.)\b", block.body):
            score += 20
            reasons.append("Contains company suffix")
        if re.search(r"\b\d{5}(?:-\d{4})?\b|\b[A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2}\b", block.body):
            score += 15
            reasons.append("Contains zip/postal code")
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


def phones_owned_by_customer(block: Optional[MessageBlock], all_blocks: Optional[List[MessageBlock]] = None) -> List[str]:
    if not block:
        return []
    text = block.body or ""
    if is_internal_text(text[:1000]) and not block.is_external_marker:
        return []
    phones = extract_phones(text)

    # Also pull phones from the top-2 non-internal blocks (they may be in the signature area
    # of a nearby block rather than the request body itself).
    if all_blocks:
        for b in all_blocks[:4]:
            if b.index == block.index or b.is_internal:
                continue
            for ph in extract_phones(b.body or ""):
                if ph not in phones:
                    phones.append(ph)

    # Reject phone candidates that only occur in internal-looking paragraphs.
    accepted = []
    all_text = text + "\n" + "\n".join(b.body or "" for b in (all_blocks or [])[:4] if not b.is_internal)
    paragraphs = re.split(r"\n\s*\n", all_text)
    for phone in phones:
        raw_digits = re.sub(r"\D", "", phone)
        owner_ok = False
        for para in paragraphs:
            para_digits = re.sub(r"\D", "", para)
            if raw_digits and raw_digits in para_digits:
                if not is_internal_text(para):
                    owner_ok = True
                break
        if owner_ok and phone not in accepted:
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
        r"price|availability|supply|attached|filter|pump|part|model|serial|drawing"
    )
    continuation_re = re.compile(
        r"(?i)^(?:qty|quantity|model|part|p/?n|pn|serial|s/n|sn|akp|rfbn|bieri)\b|"
        r"\b(?:qty|quantity)\s*[:#]?\s*\d+\b|"
        r"\b[A-Z]{2,8}[- ]?\d[0-9A-Z,./\-]*\b"
    )

    for raw in text.splitlines():
        line = raw.strip(" >\t")
        if not line:
            if started:
                blank_after_start += 1
                # Do not stop immediately: product/model/Qty often appears after a blank line.
                if blank_after_start > 2:
                    break
            continue
        if any(re.search(p, line, re.I) for p in REQUEST_STOP_PATTERNS):
            if started:
                break
            continue
        if EMAIL_RE.search(line) and len(line) < 80 and not request_start_re.search(line):
            continue
        # Skip phone-only lines, but NOT lines that contain a request (product number in context)
        if PHONE_GENERIC_RE.search(line) and len(line) < 80 and not request_start_re.search(line):
            continue

        if request_start_re.search(line):
            started = True
            blank_after_start = 0
            lines.append(line)
        elif started and continuation_re.search(line):
            blank_after_start = 0
            lines.append(line)
        elif started:
            # Keep short immediate continuation lines, but stop before signatures/trail noise.
            if blank_after_start <= 1 and len(lines) < 8 and not is_internal_text(line):
                blank_after_start = 0
                lines.append(line)
            else:
                break
    if not lines:
        # Conservative fallback: first non-signature, non-internal paragraph from selected customer block.
        for para in re.split(r"\n\s*\n", text):
            para = clean_text(para)
            if len(para) < 15:
                continue
            if is_internal_text(para):
                continue
            if re.search(r"(?i)regards|confidential|from:|sent:|to:|subject:", para[:80]):
                continue
            lines = para.splitlines()[:6]
            break
    return clean_text("\n".join(lines))[:1500]


def extract_product(text: str, attachment_names: List[str]) -> str:
    joined = "\n".join([text or "", "\n".join(attachment_names or [])])

    # HYDAC model patterns first. Prefer full model codes over generic specs like flow rate.
    model_patterns = [
        r"\bRF\d+(?:-\d+)?(?:-[A-Z0-9]+)+(?:-[A-Z0-9]+)*\b",
        r"\bRFBN\s*[A-Z0-9][A-Z0-9\-/.]*\b",
        r"\bAKP\d+[A-Z0-9,./\-]*\b",
        r"\bBIERI\s*\d{4,}\b",
        r"\b(?:part|model|serial|p/n|pn)\s*[:#-]?\s*([A-Z0-9][A-Z0-9\-/.]{3,})\b",
        # Bare 7-digit HYDAC part/order numbers (e.g. 2065706)
        r"\b(\d{7})\b",
        # HYDAC filter model strings like "DF ON 240 TE 10 B M 1.0/12"
        r"\b((?:DF|RF|MF|RKM|UDE|HF|HC|HCE|HS|EF|NF|VF)(?:\s+[A-Z0-9][A-Z0-9./]*){2,})\b",
    ]
    found = []
    for pattern in model_patterns:
        for m in re.finditer(pattern, joined, re.I):
            value = clean_text(m.group(1) if m.lastindex else m.group(0)).strip(" .,:;")
            if value and value.lower() not in {"till 150", "150 l/min"} and value not in found:
                found.append(value)
    if found:
        # Strip trailing noise words that crept in via broad patterns
        _NOISE_TAIL = re.compile(r"\s+(?:high|low|new|old|for|the|and|or|in|on|at|to|of)$", re.I)
        cleaned = [_NOISE_TAIL.sub("", f).strip() for f in found]
        # Include a short base model plus full configured model when both appear.
        return " / ".join(c for c in cleaned[:3] if c)

    if re.search(r"(?i)backflush filter", joined):
        return "Backflush Filter"
    if re.search(r"(?i)hydraulic pump", joined):
        return "Hydraulic Pump"
    if re.search(r"(?i)return filter", joined):
        return "Return Filter"
    if re.search(r"(?i)high.pressure.*filter|in.line.*filter|industrial filter", joined):
        return "High-Pressure In-Line Industrial Filter"
    if re.search(r"(?i)\bfilter\b", joined):
        return "Filter"
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
    first, last = split_person_name(sender_name, selected_email)
    selected_text = selected.body if selected else ""

    phones = phones_owned_by_customer(selected, blocks)
    lead_comments = extract_request(selected)
    product = extract_product("\n".join([lead_comments, selected_text]), valid_attachment_names)
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


def run_ai_cleaner(parsed, uploaded_name: str, deterministic_data: Dict, provider: str = "openai") -> Dict:
    openai_key = st.secrets.get("OPENAI_API_KEY", "")
    anthropic_key = st.secrets.get("ANTHROPIC_API_KEY", "")

    use_openai = provider == "openai" and openai_key
    use_claude = provider == "claude" and anthropic_key

    if not use_openai and not use_claude:
        missing = "OPENAI_API_KEY" if provider == "openai" else "ANTHROPIC_API_KEY"
        deterministic_data["AgentReason"] += f" | AI cleaner not used: {missing} missing."
        return deterministic_data

    selected_block_num = deterministic_data.get("Evidence", {}).get("SelectedCustomerBlock")
    selected_block_text = ""
    for block in parsed["blocks"]:
        if selected_block_num and block.index + 1 == selected_block_num:
            selected_block_text = block.body[:8000]
            break

    system_prompt = (
        "You are HYDAC Best Agent V2 — a precise lead extraction assistant. "
        "Return ONLY valid JSON. No preamble, no markdown fences, no explanation."
    )

    prompt = f"""
You are reviewing an inbound lead email for HYDAC, an industrial filtration and hydraulics company.
Deterministic pre-analysis has already selected the most likely external customer block.
Your job is to CLEAN and VERIFY the extracted fields — not re-select the customer.

STRICT RULES:
- Use ONLY information explicitly written in the email. Never guess or infer.
- FirstName / LastName: split the sender name cleanly. Remove titles (Mr/Mrs/Dr) and suffixes.
- ContactTitle: job title from signature only (e.g. "Purchasing Manager", "Maintenance Engineer").
- Email: the real external customer email. Never use a HYDAC domain.
- Company: must be explicitly written in the signature or body. If email is Gmail/Yahoo/free mail, leave Company blank unless the company name is written out.
- Address / City / State / ZipCode / Country: only if written. Never infer from email domain.
- PhoneSupplied: customer-owned phone(s) only, formatted as "phone1 : phone2". Exclude HYDAC employee phones.
- WebAddress: only if an explicit URL or www address is written. Never derive from email domain.
- LeadComments: the actual customer request/inquiry text only. Exclude greetings, signatures, forwarding trail, disclaimers. Keep product model numbers and quantities.
- Product: HYDAC model number or product category (e.g. "RF3-30", "Return Filter", "Hydraulic Pump"). Leave blank if none found.
- Confidence: "High" if email + LeadComments both extracted cleanly; "Medium" if one is uncertain; "Low" otherwise.
- AgentReason: one sentence explaining your key decision.

DETERMINISTIC PRE-ANALYSIS (use as baseline, correct errors):
{json.dumps({k: v for k, v in deterministic_data.items() if k != "Evidence"}, ensure_ascii=False, indent=2)}

EVIDENCE:
{json.dumps(deterministic_data.get("Evidence", {}), ensure_ascii=False, indent=2)}

SELECTED CUSTOMER BLOCK TEXT:
{selected_block_text}

Return ONLY valid JSON with these exact keys:
FirstName, LastName, ContactTitle, Email, Company, Address, City, State, ZipCode, Country,
PhoneSupplied, WebAddress, LeadComments, Product, Confidence, AgentReason
"""

    try:
        if use_openai:
            from openai import OpenAI
            client = OpenAI(api_key=openai_key)
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                temperature=0,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt},
                ],
            )
            content = response.choices[0].message.content.strip()
        else:
            import anthropic
            client = anthropic.Anthropic(api_key=anthropic_key)
            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=1000,
                system=system_prompt,
                messages=[{"role": "user", "content": prompt}],
            )
            content = response.content[0].text.strip()

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
        provider_label = "Claude" if use_claude else "GPT-4o-mini"
        ai_data["AgentReason"] = (
            f"{ai_data.get('AgentReason', '')} | "
            f"AI ({provider_label}) | "
            f"Deterministic evidence score: {deterministic_data.get('Evidence', {}).get('SelectedCustomerScore', '')}"
        )
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


st.title("HYDAC Best Agent V2")
st.caption("Reasoning-first lead processor: thread parsing, customer ranking, phone ownership, attachment evidence, review, then Excel export.")

with st.expander("Agent status & settings"):
    has_openai = bool(st.secrets.get("OPENAI_API_KEY", ""))
    has_claude = bool(st.secrets.get("ANTHROPIC_API_KEY", ""))

    provider_options = []
    if has_openai:
        provider_options.append("openai")
    if has_claude:
        provider_options.append("claude")

    if provider_options:
        provider_labels = {"openai": "OpenAI (GPT-4o-mini)", "claude": "Claude (claude-sonnet-4-6)"}
        selected_provider = st.radio(
            "AI cleaner provider",
            provider_options,
            format_func=lambda x: provider_labels[x],
            horizontal=True,
        )
        st.success(f"AI cleaner active: {provider_labels[selected_provider]}")
    else:
        selected_provider = "openai"
        st.warning("Deterministic mode active. Add OPENAI_API_KEY or ANTHROPIC_API_KEY to enable AI field cleanup.")

    auto_export = st.checkbox(
        "Auto-export without review panel when Confidence = High",
        value=False,
        help="Skips the manual review form for high-confidence extractions. A summary is still shown.",
    )

uploaded_file = st.file_uploader("Upload .msg file", type=["msg"])

if uploaded_file:
    try:
        parsed = parse_msg(uploaded_file)
        deterministic_data = build_reasoning(parsed, uploaded_file.name)
        agent_data = run_ai_cleaner(parsed, uploaded_file.name, deterministic_data, provider=selected_provider)

        confidence = agent_data.get("Confidence", "Low")
        confidence_color = {"High": "green", "Medium": "orange", "Low": "red"}.get(confidence, "gray")
        st.markdown(
            f"**Confidence:** :{confidence_color}[{confidence}]  &nbsp;|&nbsp; "
            f"**Block:** {agent_data.get('Evidence', {}).get('SelectedCustomerBlock', '?')}  &nbsp;|&nbsp; "
            f"**Score:** {agent_data.get('Evidence', {}).get('SelectedCustomerScore', '?')}"
        )

        if auto_export and confidence == "High":
            # Auto-export path: show a compact summary, skip the review form
            st.success("High-confidence extraction — auto-export enabled.")
            summary_cols = st.columns(3)
            summary_cols[0].markdown(
                f"**{agent_data.get('FirstName', '')} {agent_data.get('LastName', '')}**  \n"
                f"{agent_data.get('ContactTitle', '')}  \n"
                f"{agent_data.get('Email', '')}"
            )
            summary_cols[1].markdown(
                f"**{agent_data.get('Company', '')}**  \n"
                f"{agent_data.get('Address', '')}  \n"
                f"{agent_data.get('City', '')} {agent_data.get('State', '')} {agent_data.get('ZipCode', '')}  \n"
                f"{agent_data.get('Country', '')}"
            )
            summary_cols[2].markdown(
                f"**Phone:** {agent_data.get('PhoneSupplied', '')}  \n"
                f"**Web:** {agent_data.get('WebAddress', '')}  \n"
                f"**Product:** {agent_data.get('Product', '')}"
            )
            st.markdown(f"**Request:** {agent_data.get('LeadComments', '')}")

            row = {h: "" for h in HEADER}
            row.update({
                "Product": agent_data.get("Product", ""),
                "ReceivedDateTime": parsed["date"],
                "FirstName": agent_data.get("FirstName", ""),
                "LastName": agent_data.get("LastName", ""),
                "ContactTitle": agent_data.get("ContactTitle", ""),
                "Email": agent_data.get("Email", ""),
                "Company": agent_data.get("Company", ""),
                "Address": agent_data.get("Address", ""),
                "City": agent_data.get("City", ""),
                "State": agent_data.get("State", ""),
                "ZipCode": agent_data.get("ZipCode", ""),
                "Country": agent_data.get("Country", ""),
                "LeadSource1": "Email",
                "LeadComments": agent_data.get("LeadComments", ""),
                "PhoneSupplied": agent_data.get("PhoneSupplied", ""),
                "PDF": build_pdf_field(uploaded_file.name, agent_data.get("ValidAttachmentNames", [])),
                "WebAddress": agent_data.get("WebAddress", ""),
            })
            excel_bytes = make_excel(row)
            st.download_button(
                "Download Excel (auto-export)",
                excel_bytes,
                file_name="hydac_lead.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            render_evidence(agent_data)

        else:
            # Manual review path
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
