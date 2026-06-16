import json
import re
import tempfile
import zipfile
from dataclasses import dataclass, field, asdict
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import Workbook

st.set_page_config(page_title="HYDAC Best Agent V4", layout="wide")

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

INTERNAL_DOMAINS = {"hydacusa.com", "hydac.com", "hydac-interlynx.com"}
INTERNAL_WORDS = {
    "hydac", "hydac technology", "hydac international", "hydac usa", "hydac sales",
    "brandon", "interlynx"
}
FREE_EMAIL_DOMAINS = {
    "gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "aol.com", "icloud.com",
    "me.com", "live.com", "msn.com", "protonmail.com", "gmx.com", "mail.com"
}
BAD_NOISE_PATTERNS = [
    r"you don'?t often receive email from",
    r"e-mails? from .* don'?t.* often",
    r"e-maily z adresy",
    r"learnaboutsenderidentification",
    r"aka\.ms", r"microsoft", r"office365", r"outlook", r"safelinks",
    r"proofpoint", r"mimecast", r"external email", r"caution:", r"warning:",
    r"sender identification", r"confidential", r"disclaimer", r"virus-free",
]
BAD_WEBSITE_PATTERNS = ["aka.ms", "microsoft.com", "office.com", "outlook.com", "safelinks", "proofpoint", "mimecast", "teams.microsoft", "sharepoint.com"]
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
VALID_ATTACHMENT_EXTENSIONS = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv", ".step", ".stp", ".igs", ".iges", ".dwg", ".dxf", ".zip", ".rar", ".7z", ".png", ".jpg", ".jpeg", ".tif", ".tiff"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".tif", ".tiff"}
COMPANY_SUFFIXES = ["inc", "inc.", "llc", "ltd", "ltd.", "corp", "corp.", "corporation", "company", "co.", "gmbh", "ag", "kg", "s.a.", "sa", "bv", "nv", "plc", "limited", "international", "industries", "systems", "technologies", "technology", "group", "usa", "us"]
TITLE_WORDS = ["manager", "engineer", "buyer", "purchasing", "director", "president", "sales", "maintenance", "supervisor", "specialist", "procurement", "automation", "project", "technician", "coordinator", "owner"]
REQUEST_STOP_RE = re.compile(r"(?i)^(best regards|kind regards|regards|thanks|thank you|sincerely|mit freundlichen|freundliche|sent from my|from:|sent:|to:|subject:|-{2,}\s*original message)")
REQUEST_START_RE = re.compile(r"(?i)(quote|quotation|rfq|request|demand|need|looking for|please|can you|could you|price|availability|supply|attached|filter|pump|part|model|serial|drawing|project|qty|quantity|pcs)")
CONTINUATION_RE = re.compile(r"(?i)^(qty|quantity|model|part|p/?n|pn|serial|s/n|sn|akp|rf|rfbn|bieri|for our|something like|if you|please)\b|\b(qty|quantity)\s*[:#]?\s*\d+\b|\b[A-Z]{2,10}\d?[A-Z0-9]*(?:[-/,][A-Z0-9]+){1,}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
URL_RE = re.compile(r"(?i)\b(?:https?://|www\.)[^\s<>]+|\b[A-Za-z0-9.-]+\.(?:com|net|org|de|at|ca|in|co|cz|io|us|eu|uk)\b")
PHONE_LABEL_RE = re.compile(r"(?i)(?:tel|telephone|phone|mobile|cell|direct|office|fax|m|t|p)[:.\s-]*(\+?\d[\d\s().\-/]{6,}\d)")
PHONE_GENERIC_RE = re.compile(r"(?<!\w)(\+?\d[\d\s().\-/]{7,}\d)(?!\w)")

STATE_MAP = {
    "alabama":"AL","alaska":"AK","arizona":"AZ","arkansas":"AR","california":"CA","colorado":"CO","connecticut":"CT","delaware":"DE","florida":"FL","georgia":"GA","hawaii":"HI","idaho":"ID","illinois":"IL","indiana":"IN","iowa":"IA","kansas":"KS","kentucky":"KY","louisiana":"LA","maine":"ME","maryland":"MD","massachusetts":"MA","michigan":"MI","minnesota":"MN","mississippi":"MS","missouri":"MO","montana":"MT","nebraska":"NE","nevada":"NV","new hampshire":"NH","new jersey":"NJ","new mexico":"NM","new york":"NY","north carolina":"NC","north dakota":"ND","ohio":"OH","oklahoma":"OK","oregon":"OR","pennsylvania":"PA","rhode island":"RI","south carolina":"SC","south dakota":"SD","tennessee":"TN","texas":"TX","utah":"UT","vermont":"VT","virginia":"VA","washington":"WA","west virginia":"WV","wisconsin":"WI","wyoming":"WY"
}

@dataclass
class AttachmentDecision:
    filename: str
    decision: str
    reason: str
    attachment: object = None

@dataclass
class MessageUnit:
    index: int
    raw: str
    header: str
    body: str
    sender_name: str = ""
    sender_email: str = ""
    emails: List[str] = field(default_factory=list)
    phones: List[str] = field(default_factory=list)
    websites: List[str] = field(default_factory=list)
    external_marker: bool = False
    internal: bool = False
    noise_score: int = 0

@dataclass
class Participant:
    email: str
    name: str = ""
    first: str = ""
    last: str = ""
    company: str = ""
    title: str = ""
    phones: List[str] = field(default_factory=list)
    website: str = ""
    address: Dict[str, str] = field(default_factory=dict)
    source_blocks: List[int] = field(default_factory=list)
    evidence_lines: List[str] = field(default_factory=list)
    is_internal: bool = False
    score: int = 0
    reasons: List[str] = field(default_factory=list)
    request: str = ""
    product: str = ""


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


def domain_root(email: str) -> str:
    dom = email_domain(email)
    parts = dom.split(".")
    return parts[0] if parts else ""


def is_internal_email(email: str) -> bool:
    dom = email_domain(email)
    return any(dom == d or dom.endswith("." + d) for d in INTERNAL_DOMAINS)


def is_noise_line(line: str) -> bool:
    low = (line or "").lower()
    return any(re.search(p, low) for p in BAD_NOISE_PATTERNS)


def is_internal_text(text: str) -> bool:
    low = (text or "").lower()
    return any(k in low for k in INTERNAL_WORDS) or any(d in low for d in INTERNAL_DOMAINS)


def extract_emails(text: str) -> List[str]:
    out = []
    for e in EMAIL_RE.findall(text or ""):
        e = normalize_email(e)
        if e and e not in out:
            out.append(e)
    return out


def normalize_phone(phone: str) -> str:
    phone = (phone or "").strip().replace("+", "")
    phone = re.sub(r"[().\s]+", "", phone)
    phone = phone.replace("/", "-")
    phone = re.sub(r"-+", "-", phone)
    return phone.strip("-")


def extract_phones(text: str) -> List[str]:
    candidates = PHONE_LABEL_RE.findall(text or "") + PHONE_GENERIC_RE.findall(text or "")
    out = []
    for p in candidates:
        p = normalize_phone(p)
        digits = re.sub(r"\D", "", p)
        if 7 <= len(digits) <= 18 and p not in out:
            out.append(p)
    return out


def extract_websites(text: str, selected_email: str = "") -> List[str]:
    out = []
    for line in (text or "").splitlines():
        # Domains inside email addresses are not websites. Require explicit www/http if line has an email.
        if EMAIL_RE.search(line) and not re.search(r"(?i)https?://|www\.", line):
            continue
        for url in URL_RE.findall(line):
            clean = url.strip(".,);]>")
            host = clean.lower().removeprefix("http://").removeprefix("https://").removeprefix("www.").split("/", 1)[0]
            if "@" in clean:
                continue
            if host in FREE_EMAIL_DOMAINS:
                continue
            if selected_email and host == email_domain(selected_email):
                # Never infer website from email domain unless explicitly shown as www/http.
                if not clean.lower().startswith(("www.", "http://", "https://")):
                    continue
            if any(bad in clean.lower() for bad in BAD_WEBSITE_PATTERNS):
                continue
            if any(d in host for d in INTERNAL_DOMAINS):
                continue
            if clean not in out:
                out.append(clean)
    return out


def split_person_name(name: str, email: str = "") -> Tuple[str, str]:
    name = clean_text(name)
    name = re.sub(r"[\"'<>]", "", name)
    name = re.sub(r"\([^)]*\)", "", name).strip(" -;,\t")
    if not name or "@" in name:
        local = normalize_email(email or name).split("@", 1)[0]
        name = re.sub(r"[._\-]+", " ", local).title()
    parts = [p for p in name.split() if p]
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


def parse_sender(header: str) -> Tuple[str, str]:
    email_match = EMAIL_RE.search(header or "")
    email = normalize_email(email_match.group(0)) if email_match else ""
    name = ""
    for line in (header or "").splitlines():
        if re.match(r"(?i)^\s*from\s*:", line):
            val = re.sub(r"(?i)^\s*from\s*:\s*", "", line).strip()
            val = EMAIL_RE.sub("", val).replace("<", "").replace(">", "").strip(" -;,\t")
            name = val
            break
    return name, email


def extract_signature_window(text: str, email: str = "") -> str:
    lines = [l.rstrip() for l in (text or "").splitlines()]
    if not lines:
        return ""
    positions = []
    if email:
        for i, line in enumerate(lines):
            if email.lower() in line.lower():
                positions.append(i)
    for i, line in enumerate(lines):
        if re.search(r"(?i)^(best regards|kind regards|regards|thanks|thank you|sincerely|mit freundlichen|freundliche)", line.strip()):
            positions.append(i)
    if positions:
        start = max(0, min(positions) - 2)
        end = min(len(lines), max(positions) + 12)
        return clean_text("\n".join(lines[start:end]))
    # fallback: last 20 non-noise lines often contain signature
    clean_lines = [l for l in lines if l.strip() and not is_noise_line(l)]
    return clean_text("\n".join(clean_lines[-20:]))


def split_email_timeline(text: str) -> List[MessageUnit]:
    text = clean_text(text)
    if not text:
        return []
    # Preserve message boundaries. This intentionally over-splits; later participant logic recombines evidence.
    boundary = re.compile(r"(?im)(?=^\s*(?:-{2,}\s*)?(?:original message|forwarded message|from\s*:|sent\s*:|to\s*:|subject\s*:))")
    parts = [p.strip() for p in boundary.split(text) if p.strip()]
    if not parts:
        parts = [text]
    # Merge metadata fragments.
    merged, buf = [], ""
    for part in parts:
        if len(part) < 90 and re.search(r"(?im)^\s*(from|sent|to|subject)\s*:", part):
            buf = (buf + "\n" + part).strip()
            continue
        merged.append((buf + "\n" + part).strip() if buf else part)
        buf = ""
    if buf:
        merged.append(buf)
    units = []
    for idx, part in enumerate(merged):
        header_lines, body_lines, header_mode = [], [], True
        for line in part.splitlines():
            if header_mode and re.match(r"(?i)^\s*(from|sent|to|cc|bcc|subject|date)\s*:", line):
                header_lines.append(line)
            else:
                header_mode = False
                body_lines.append(line)
        header = "\n".join(header_lines)
        body = clean_text("\n".join(body_lines)) or part
        sender_name, sender_email = parse_sender(header or part[:700])
        emails = extract_emails(part)
        phones = extract_phones(part)
        websites = extract_websites(part)
        external_marker = bool(re.search(r"(?i)external email|caution:", part[:1500]))
        internal = is_internal_email(sender_email) or is_internal_text(header[:1000])
        noise_score = sum(1 for line in part.splitlines()[:30] if is_noise_line(line))
        units.append(MessageUnit(idx, part, header, body, sender_name, sender_email, emails, phones, websites, external_marker, internal, noise_score))
    return units


def line_score_as_company(line: str, selected_email: str = "") -> int:
    raw = line.strip(" |\t,;-")
    low = raw.lower()
    if not raw or len(raw) < 2 or len(raw) > 90:
        return -100
    if is_noise_line(raw) or is_internal_text(raw):
        return -100
    if EMAIL_RE.search(raw) or PHONE_GENERIC_RE.search(raw) or URL_RE.search(raw):
        return -50
    if re.search(r"(?i)^(tel|phone|mobile|fax|email|www|http|address|direct)", raw):
        return -40
    score = 0
    words = re.findall(r"[A-Za-z&.\-]+", low)
    if any(s in words or low.endswith(" " + s) for s in COMPANY_SUFFIXES):
        score += 50
    if selected_email and domain_root(selected_email) and domain_root(selected_email) not in {"gmail", "yahoo", "hotmail", "outlook", "icloud"}:
        if domain_root(selected_email).lower() in low:
            score += 70
    if re.search(r"\b(inc|llc|ltd|gmbh|ag|kg|corp|company|international|industries|technology|systems|usa|us)\b", low):
        score += 35
    # Company lines often use title case or all caps and are short.
    if raw[:1].isupper() and 1 <= len(raw.split()) <= 5:
        score += 15
    return score


def extract_company(signature: str, block_text: str, email: str = "") -> str:
    candidates = []
    for source, text in [("signature", signature), ("block", block_text)]:
        for line in (text or "").splitlines()[:100]:
            s = line_score_as_company(line, email)
            if s > 0:
                candidates.append((s + (30 if source == "signature" else 0), line.strip(" |\t,;-")))
    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        return candidates[0][1]
    return ""


def extract_title(signature: str) -> str:
    for line in (signature or "").splitlines()[:40]:
        clean = line.strip(" |\t,;-")
        low = clean.lower()
        if 2 <= len(clean) <= 80 and any(w in low for w in TITLE_WORDS):
            if not EMAIL_RE.search(clean) and not PHONE_GENERIC_RE.search(clean) and not is_noise_line(clean) and not is_internal_text(clean):
                return clean
    return ""


def extract_address(signature: str, block_text: str) -> Dict[str, str]:
    result = {"Address": "", "City": "", "State": "", "ZipCode": "", "Country": ""}
    lines = [l.strip(" |\t,;") for l in (signature or block_text or "").splitlines() if l.strip()]
    for i, line in enumerate(lines[:80]):
        if is_noise_line(line) or is_internal_text(line) or EMAIL_RE.search(line):
            continue
        if re.search(r"\b\d{1,6}\s+[A-Za-z0-9 .'-]+\b(?:street|st\.?|road|rd\.?|ave|avenue|drive|dr\.?|lane|ln\.?|blvd|way|parkway|pkwy|platz|strasse|straße|square)\b", line, re.I):
            result["Address"] = line.strip(" ,;")
            for j in range(i + 1, min(i + 4, len(lines))):
                city_line = lines[j].strip(" ,;")
                m = re.search(r"^(.+?),?\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)$", city_line)
                if m:
                    result["City"], result["State"], result["ZipCode"] = m.group(1).strip(" ,"), m.group(2), m.group(3)
                    break
                m = re.search(r"^(.+?),?\s+([A-Za-z ]+)\s+(\d{5}(?:-\d{4})?)$", city_line)
                if m and m.group(2).lower() in STATE_MAP:
                    result["City"], result["State"], result["ZipCode"] = m.group(1).strip(" ,"), STATE_MAP[m.group(2).lower()], m.group(3)
                    break
            # country can be next nearby line
            for j in range(i + 1, min(i + 6, len(lines))):
                if re.fullmatch(r"(?i)(usa|united states|us|canada|austria|germany|czech republic|cz|india|uk|united kingdom)", lines[j].strip()):
                    result["Country"] = lines[j].strip()
                    break
            break
    return result


def extract_request_from_unit(unit: MessageUnit, participant_email: str = "") -> str:
    text = unit.body or unit.raw
    # Remove top warning banner only, not the actual customer request.
    lines = []
    for line in text.splitlines():
        stripped = line.strip(" >\t")
        if is_noise_line(stripped):
            continue
        lines.append(stripped)
    started = False
    out = []
    blank_count = 0
    for line in lines:
        if not line:
            if started:
                blank_count += 1
                if blank_count > 2:
                    break
            continue
        if REQUEST_STOP_RE.search(line):
            if started:
                break
            continue
        if EMAIL_RE.search(line) and len(line) < 90:
            continue
        if PHONE_GENERIC_RE.search(line) and len(line) < 90:
            continue
        if is_internal_text(line):
            continue
        if REQUEST_START_RE.search(line):
            started = True
            blank_count = 0
            out.append(line)
        elif started and (CONTINUATION_RE.search(line) or len(out) < 8):
            blank_count = 0
            if not is_noise_line(line):
                out.append(line)
        elif started:
            break
    if not out:
        paras = re.split(r"\n\s*\n", "\n".join(lines))
        for para in paras:
            para = clean_text(para)
            if len(para) >= 15 and not is_internal_text(para) and not is_noise_line(para) and REQUEST_START_RE.search(para):
                out = para.splitlines()[:8]
                break
    return clean_text("\n".join(out))[:1600]


def extract_product(text: str, attachment_names: List[str]) -> str:
    joined = clean_text("\n".join([text or "", "\n".join(attachment_names or [])]))
    candidates = []
    patterns = [
        (100, r"\bRF\d+(?:[-/][A-Z0-9]+){1,}\b"),
        (100, r"\bRFBN\s*[A-Z0-9\-/.]*\b"),
        (100, r"\bAKP\d+[A-Z0-9,./\-]*\b"),
        (90, r"\bBIERI\s*\d{4,}\b"),
        (80, r"\b[A-Z]{2,10}\d?[A-Z0-9]*(?:[-/,][A-Z0-9]+){2,}\b"),
        (70, r"(?i)\b(?:part|model|serial|p/n|pn)\s*[:#-]?\s*([A-Z0-9][A-Z0-9\-/.]{3,})\b"),
    ]
    for base, pat in patterns:
        for m in re.finditer(pat, joined, re.I):
            val = m.group(1) if m.lastindex else m.group(0)
            val = clean_text(val).strip(" .,:;")
            if not val or re.search(r"(?i)^till\s+\d+$", val):
                continue
            candidates.append((base + len(val), val))
    # include short base model RF4-2 if longer model also exists
    base_models = []
    for _, val in candidates:
        m = re.match(r"^(RF\d+-\d+)", val, re.I)
        if m and m.group(1) not in base_models:
            base_models.append(m.group(1))
    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        vals = []
        for v in base_models + [v for _, v in candidates]:
            if v not in vals:
                vals.append(v)
        return " / ".join(vals[:3])
    if re.search(r"(?i)hydraulic pump", joined):
        return "Hydraulic Pump"
    if re.search(r"(?i)backflush filter", joined):
        return "Backflush Filter"
    if re.search(r"(?i)return filter", joined):
        return "Return Filter"
    return ""


def decide_attachment(filename: str, attachment_obj=None) -> AttachmentDecision:
    if not filename:
        return AttachmentDecision("", "Reject", "Blank attachment name", attachment_obj)
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
    messages = split_email_timeline(full_text)
    attachment_decisions = []
    for att in msg.attachments:
        fname = getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or ""
        if fname:
            attachment_decisions.append(decide_attachment(fname, att))
    return {"sender": sender, "subject": subject, "date": date, "body": body, "full_text": full_text, "messages": messages, "attachment_decisions": attachment_decisions}


def build_participant_ledger(messages: List[MessageUnit]) -> Dict[str, Participant]:
    ledger: Dict[str, Participant] = {}
    for unit in messages:
        all_emails = list(unit.emails)
        if unit.sender_email and unit.sender_email not in all_emails:
            all_emails.insert(0, unit.sender_email)
        for email in all_emails:
            if not email:
                continue
            p = ledger.setdefault(email, Participant(email=email, address={"Address":"", "City":"", "State":"", "ZipCode":"", "Country":""}))
            if unit.index + 1 not in p.source_blocks:
                p.source_blocks.append(unit.index + 1)
            if email == unit.sender_email and unit.sender_name:
                p.name = unit.sender_name
            if not p.name:
                # look for nearby name before email in lines
                for line in unit.raw.splitlines():
                    if email in line.lower():
                        candidate = EMAIL_RE.sub("", line).replace("<", "").replace(">", "").strip(" -;,\t")
                        candidate = re.sub(r"(?i)^(from|to|cc|email|e-mail)\s*:\s*", "", candidate).strip()
                        if candidate and not is_noise_line(candidate) and len(candidate) < 80:
                            p.name = candidate
                            break
            sig = extract_signature_window(unit.body, email)
            p.evidence_lines.extend([l for l in sig.splitlines()[:12] if l.strip() and l.strip() not in p.evidence_lines])
            if not p.first and not p.last:
                p.first, p.last = split_person_name(p.name, email)
            if not p.company:
                p.company = extract_company(sig, unit.body, email)
            if not p.title:
                p.title = extract_title(sig)
            if not p.website:
                ws = extract_websites(sig, email) or extract_websites(unit.body, email)
                p.website = ws[0] if ws else ""
            if not any(p.address.values()):
                p.address = extract_address(sig, unit.body)
            for ph in extract_phones(sig):
                if ph not in p.phones and not is_internal_text(sig):
                    p.phones.append(ph)
            if not p.request:
                req = extract_request_from_unit(unit, email)
                if req:
                    p.request = req
            p.is_internal = p.is_internal or is_internal_email(email) or is_internal_text(sig[:800])
    return ledger


def score_participants(ledger: Dict[str, Participant], messages: List[MessageUnit], attachment_names: List[str]) -> List[Participant]:
    request_by_block = {m.index + 1: extract_request_from_unit(m) for m in messages}
    for p in ledger.values():
        score = 0
        reasons = []
        if p.is_internal:
            score -= 300
            reasons.append("Rejected: HYDAC/internal email or signature")
        else:
            score += 50
            reasons.append("External email domain")
        if email_domain(p.email) in FREE_EMAIL_DOMAINS:
            score -= 20
            reasons.append("Free email domain; do not infer company/website")
        sender_blocks = [m for m in messages if m.sender_email == p.email]
        if sender_blocks:
            score += 60
            reasons.append("Appears as sender in message timeline")
        if any(m.external_marker for m in messages if m.index + 1 in p.source_blocks):
            score += 80
            reasons.append("Appears inside/near EXTERNAL EMAIL block")
        if p.request or any(request_by_block.get(b) for b in p.source_blocks):
            score += 70
            reasons.append("Near actual request language")
        if p.company:
            score += 35
            reasons.append("Company found near participant signature")
        if p.phones:
            score += 25
            reasons.append("Customer-owned phone found near participant signature")
        if p.website:
            score += 15
            reasons.append("Website explicitly written near participant")
        if p.name:
            score += 10
            reasons.append("Name found")
        # Penalize pure banner/security pseudo-candidates.
        if any(is_noise_line(l) for l in p.evidence_lines[:5]):
            score -= 60
            reasons.append("Penalized: security/banner noise near candidate")
        p.score = score
        p.reasons = reasons
        if not p.product:
            context = "\n".join([p.request] + [messages[b-1].body for b in p.source_blocks if 0 <= b-1 < len(messages)])
            p.product = extract_product(context, attachment_names)
    return sorted(ledger.values(), key=lambda x: x.score, reverse=True)


def choose_customer(candidates: List[Participant]) -> Optional[Participant]:
    for p in candidates:
        if p.score > 0 and not p.is_internal:
            return p
    return candidates[0] if candidates else None


def build_case(parsed, uploaded_name: str) -> Dict:
    kept = [d.filename for d in parsed["attachment_decisions"] if d.decision == "Keep"]
    ignored = [d.filename for d in parsed["attachment_decisions"] if d.decision == "Reject"]
    ledger = build_participant_ledger(parsed["messages"])
    ranked = score_participants(ledger, parsed["messages"], kept)
    selected = choose_customer(ranked)
    warnings = []
    if not selected:
        warnings.append("No customer candidate selected.")
    elif selected.score < 100:
        warnings.append("Low confidence: selected customer score below 100. Manual review required.")
    if selected and not selected.request:
        # Search selected source blocks for request, then global external message requests.
        for b in selected.source_blocks:
            if 0 <= b - 1 < len(parsed["messages"]):
                req = extract_request_from_unit(parsed["messages"][b - 1], selected.email)
                if req:
                    selected.request = req
                    break
        if not selected.request:
            for m in parsed["messages"]:
                if not m.internal:
                    req = extract_request_from_unit(m, selected.email if selected else "")
                    if req:
                        selected.request = req
                        warnings.append("Request came from another external block, not directly from selected signature block.")
                        break
    if selected and not selected.product:
        selected.product = extract_product(selected.request, kept)
    evidence = {
        "SelectedCustomer": asdict(selected) if selected else {},
        "RejectedCandidates": [asdict(p) for p in ranked if p.is_internal][:10],
        "CandidateRanking": [
            {"score": p.score, "email": p.email, "name": p.name, "company": p.company, "blocks": p.source_blocks, "internal": p.is_internal, "reasons": p.reasons}
            for p in ranked[:15]
        ],
        "MessageTimeline": [
            {"block": m.index + 1, "sender": m.sender_email or m.sender_name, "internal": m.internal, "external_marker": m.external_marker, "emails": m.emails, "phones": m.phones, "preview": clean_text(m.body)[:500]}
            for m in parsed["messages"][:20]
        ],
        "AttachmentDecisions": [{"filename": d.filename, "decision": d.decision, "reason": d.reason} for d in parsed["attachment_decisions"]],
        "Warnings": warnings,
    }
    if selected:
        data = {
            "FirstName": selected.first,
            "LastName": selected.last,
            "ContactTitle": selected.title,
            "Email": selected.email,
            "Company": selected.company,
            "Address": selected.address.get("Address", ""),
            "City": selected.address.get("City", ""),
            "State": selected.address.get("State", ""),
            "ZipCode": selected.address.get("ZipCode", ""),
            "Country": selected.address.get("Country", ""),
            "PhoneSupplied": " : ".join(selected.phones),
            "WebAddress": selected.website,
            "LeadComments": selected.request,
            "Product": selected.product,
            "Confidence": "High" if selected.score >= 190 and selected.request else "Medium" if selected.score >= 110 else "Low",
            "AgentReason": "; ".join(selected.reasons),
        }
    else:
        data = {k: "" for k in ["FirstName","LastName","ContactTitle","Email","Company","Address","City","State","ZipCode","Country","PhoneSupplied","WebAddress","LeadComments","Product"]}
        data.update({"Confidence":"Low", "AgentReason":"No selected customer"})
    data.update({"Evidence": evidence, "ValidAttachmentNames": kept, "IgnoredAttachmentNames": ignored})
    return data


def run_ai_reviewer(parsed, case_data: Dict) -> Dict:
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        case_data["AgentReason"] += " | AI reviewer not used: OPENAI_API_KEY missing."
        return case_data
    from openai import OpenAI
    client = OpenAI(api_key=api_key)
    prompt = f"""
You are HYDAC Best Agent V4, a case reviewer. You must challenge the deterministic result, not merely clean text.

Task:
1. Review candidate ranking, message timeline, selected customer, rejected HYDAC/internal candidates.
2. Correct the selected customer ONLY if evidence shows a different real external customer/requester.
3. Extract fields only from the chosen customer's message/signature.
4. LeadComments must be only the customer's actual request, not greetings, signatures, warnings, or email trail.
5. PhoneSupplied must contain customer-owned phones only, formatted phone1 : phone2.
6. WebAddress must be an explicitly written website only; never convert gmail.com/free email/email domain into website.
7. Company must come from signature/body evidence; never use Microsoft/Outlook warning text.
8. Product must prefer model/part numbers such as RF4-2, RF4-2-EPT2-NN-E-KN-3-16, AKP43-0,5-200-V-A, BIERI 3999534.
9. Leave unavailable fields blank.

DETERMINISTIC RESULT:
{json.dumps({k:v for k,v in case_data.items() if k != 'Evidence'}, ensure_ascii=False, indent=2)}

EVIDENCE:
{json.dumps(case_data.get('Evidence', {}), ensure_ascii=False, indent=2)}

Return ONLY valid JSON with exact keys:
FirstName, LastName, ContactTitle, Email, Company, Address, City, State, ZipCode, Country,
PhoneSupplied, WebAddress, LeadComments, Product, Confidence, AgentReason
"""
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0,
            messages=[
                {"role":"system", "content":"Return only valid JSON. You are a strict evidence-based reviewer. Never invent unavailable fields."},
                {"role":"user", "content": prompt},
            ],
        )
        content = response.choices[0].message.content.strip()
        content = content.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
        ai = json.loads(content)
        required = ["FirstName","LastName","ContactTitle","Email","Company","Address","City","State","ZipCode","Country","PhoneSupplied","WebAddress","LeadComments","Product","Confidence","AgentReason"]
        for key in required:
            ai.setdefault(key, case_data.get(key, ""))
        # safety post validation
        ws = extract_websites(ai.get("WebAddress", ""), ai.get("Email", ""))
        ai["WebAddress"] = ws[0] if ws else ""
        if ai.get("Company") and is_noise_line(ai["Company"]):
            ai["Company"] = case_data.get("Company", "")
        ai["Evidence"] = case_data["Evidence"]
        ai["ValidAttachmentNames"] = case_data["ValidAttachmentNames"]
        ai["IgnoredAttachmentNames"] = case_data["IgnoredAttachmentNames"]
        ai["AgentReason"] = f"{ai.get('AgentReason','')} | AI reviewer checked deterministic case."
        return ai
    except Exception as exc:
        case_data["AgentReason"] += f" | AI reviewer failed; deterministic case used. Error: {exc}"
        return case_data


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
    return pdf_base + ("; Attachments: " + ", ".join(valid_names) if valid_names else "")


def render_evidence(agent_data: Dict):
    evidence = agent_data.get("Evidence", {})
    st.subheader("Agent Evidence and Confidence")
    c1, c2, c3 = st.columns(3)
    c1.metric("Confidence", agent_data.get("Confidence", ""))
    selected = evidence.get("SelectedCustomer", {}) or {}
    c2.metric("Selected Score", selected.get("score", ""))
    c3.metric("Candidates", len(evidence.get("CandidateRanking", [])))
    warnings = evidence.get("Warnings", [])
    if warnings:
        st.warning(" | ".join(warnings))
    st.write("**Selected customer evidence:**")
    st.json(selected, expanded=False)
    st.write("**Candidate ranking:**")
    st.dataframe(evidence.get("CandidateRanking", []), use_container_width=True)
    st.write("**Message timeline:**")
    st.dataframe(evidence.get("MessageTimeline", []), use_container_width=True)
    st.write("**Attachment decisions:**")
    st.dataframe(evidence.get("AttachmentDecisions", []), use_container_width=True)
    with st.expander("Rejected HYDAC/internal candidates"):
        st.json(evidence.get("RejectedCandidates", []), expanded=False)


st.title("HYDAC Best Agent V4")
st.caption("Case-reasoning lead processor: participant ledger, message timeline, candidate challenge, evidence, review, then Excel export.")

with st.expander("Agent status"):
    if st.secrets.get("OPENAI_API_KEY", ""):
        st.success("AI reviewer active. Deterministic case evidence is reviewed/challenged before final fields.")
    else:
        st.warning("Deterministic case-reasoning mode active. Add OPENAI_API_KEY for reviewer challenge, not blind extraction.")

uploaded_file = st.file_uploader("Upload .msg file", type=["msg"])

if uploaded_file:
    try:
        parsed = parse_msg(uploaded_file)
        case_data = build_case(parsed, uploaded_file.name)
        agent_data = run_ai_reviewer(parsed, case_data)

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
        st.download_button("Download Reviewed Excel", excel_bytes, file_name="hydac_reviewed_lead.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if any(d.decision == "Keep" for d in parsed["attachment_decisions"]):
            zip_bytes = make_attachment_zip(parsed["attachment_decisions"])
            st.download_button("Download Valid Customer Attachments ZIP", zip_bytes, file_name="valid_customer_attachments.zip", mime="application/zip")

        with st.expander("Full extracted email text"):
            st.text(parsed["body"][:30000])
    except Exception as e:
        st.error("Agent processing failed.")
        st.exception(e)
