"""
HYDAC Lead Agent
AI-first: upload a .msg file → agent reads it → summary card → optional edit → Excel export.
"""

import json
import re
import tempfile
import zipfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional

import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import Workbook

st.set_page_config(page_title="HYDAC Lead Agent", layout="wide", page_icon="⚙️")

# ── Constants ──────────────────────────────────────────────────────────────────

EXCEL_HEADER = [
    "Referral", "Brand", "Product", "ReceivedDateTime", "FirstName", "LastName",
    "ContactTitle", "Email", "Company", "Address", "County", "City", "State",
    "ZipCode", "Country", "LeadSource1", "LeadSource2", "LeadSource3",
    "LeadComments", "Summary", "PhoneSupplied", "PhSuppliedExtension", "PhoneResearched",
    "CSRName", "PDF", "DUNS", "WebAddress", "Linkedin_Title", "Linkedin_Link",
    "SIC", "NAICS", "noOfEmployees", "ParentName", "LineOfBusiness", "PQ",
    "Latitude", "Longitude", "DemoLead", "ScreenReason", "about_me", "college_1",
    "college_1_degree", "college_1_start", "college_1_end", "college_2",
    "college_2_degree", "college_2_start", "college_2_end", "month_of_joining",
    "about_experience", "searched_on_google", "linkedin_city", "linkedin_state",
    "linkedin_country",
]

INTERNAL_DOMAINS = {"hydacusa.com", "hydac.com", "hydac-interlynx.com"}

SIGNATURE_IMAGE_NAMES = {
    "image.png","image.jpg","image.jpeg","image.gif",
    "logo.png","logo.jpg","logo.jpeg","banner.png","banner.jpg","banner.jpeg",
    "facebook.png","linkedin.png","twitter.png","instagram.png","youtube.png",
} | {f"image{str(i).zfill(3)}.{ext}"
     for i in range(1, 20) for ext in ("png","jpg","jpeg","gif")}

VALID_ATTACHMENT_EXTENSIONS = {
    ".pdf",".doc",".docx",".xls",".xlsx",".csv",
    ".step",".stp",".igs",".iges",".dwg",".dxf",
    ".zip",".rar",".7z",".png",".jpg",".jpeg",".tif",".tiff",
}

IMAGE_EXTENSIONS = {".png",".jpg",".jpeg",".gif",".tif",".tiff"}

BAD_WEBSITE_PATTERNS = [
    "aka.ms","microsoft.com","office.com","safelinks","proofpoint",
    "mimecast","teams.microsoft","sharepoint.com",
]

# ── MSG parsing helpers ────────────────────────────────────────────────────────

_ICON_PREFIX_RE = re.compile(
    r"^\s*<https?://[^\s>]+?(?:/images?/|/img/|/cms/|/static/|/icons?/|/logo|/media/)[^\s>]*>\s*[\t ]*",
    re.I,
)

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

def strip_icon_prefixes(text: str) -> str:
    """Remove inline icon image URL prefixes from every line (common in HTML signatures)."""
    return "\n".join(_ICON_PREFIX_RE.sub("", line) for line in text.splitlines())

def decide_attachment(filename: str, attachment_obj=None, subject: str = ""):
    if not filename:
        return {"filename": "", "decision": "Reject", "reason": "Blank filename"}
    name = Path(filename).name
    low  = name.lower()
    ext  = Path(low).suffix

    if ext not in VALID_ATTACHMENT_EXTENSIONS:
        return {"filename": name, "decision": "Reject", "reason": f"Unsupported extension {ext or '(none)'}"}

    # Explicit signature/logo names — always reject
    if low in SIGNATURE_IMAGE_NAMES:
        return {"filename": name, "decision": "Reject", "reason": "Signature/logo image"}

    # Plain image.ext with NO number (image.png, image.jpg) — signature logo
    if re.fullmatch(r"image[.](png|jpg|jpeg|gif|tif|tiff)", low):
        return {"filename": name, "decision": "Reject", "reason": "Signature/logo image"}

    # Numbered inline images: image004.jpg, image004(timestamp).jpg
    # These are always product photos embedded by email clients — keep unconditionally
    if re.match(r"image[0-9]", low) and ext in IMAGE_EXTENSIONS:
        return {"filename": name, "decision": "Keep", "reason": "Numbered inline product image", "obj": attachment_obj}

    # Non-image documents — always keep
    if ext not in IMAGE_EXTENSIONS:
        return {"filename": name, "decision": "Keep", "reason": "Document/CAD/spreadsheet attachment", "obj": attachment_obj}

    # Images with 4+ digit part/model number in filename
    if re.search(r"[0-9]{4,}", name):
        return {"filename": name, "decision": "Keep", "reason": "Image has part/model number", "obj": attachment_obj}

    # Images with product keywords
    if re.search(r"(?i)(pump|filter|drawing|label|plate|model|part|hydraulic|spec|quote|serial|bieri)", name):
        return {"filename": name, "decision": "Keep", "reason": "Image has product keyword", "obj": attachment_obj}

    return {"filename": name, "decision": "Reject", "reason": "Image not customer evidence"}

# ── Lead PDF generator ────────────────────────────────────────────────────────

def make_pdf_name(date_str: str) -> str:
    """Generate unique PDF name: hydac + MDDYYYY + HHMM from email date string."""
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})[\sT](\d{2}):(\d{2})", date_str or "")
    if m:
        year, month, day, hour, minute = m.groups()
        return f"hydac{int(month)}{day}{year}{hour}{minute}.pdf"
    import time
    t = time.localtime()
    return f"hydac{t.tm_mon}{t.tm_mday:02d}{t.tm_year}{t.tm_hour:02d}{t.tm_min:02d}.pdf"


def make_lead_pdf(kept_attachments: List[Dict], date_str: str = "") -> Optional[tuple]:
    """Bundle all kept images into a single PDF.
    Tries Pillow first; falls back to a raw JPEG-in-PDF wrapper (no extra packages).
    Returns (pdf_bytes, pdf_filename) or None if nothing to bundle."""
    import io as _io

    pdf_name = make_pdf_name(date_str)

    # Collect raw image bytes + filenames
    image_items = []
    for d in kept_attachments:
        try:
            data = d["obj"].data
            if not data:
                continue
            ext = Path(d["filename"]).suffix.lower()
            if ext in {".png", ".jpg", ".jpeg", ".gif", ".tif", ".tiff", ".bmp", ".webp"}:
                image_items.append((d["filename"], data))
        except Exception:
            continue

    if not image_items:
        return None

    # ── Try Pillow (best quality) ──────────────────────────────────────────
    try:
        from PIL import Image as PILImage
        images_for_pdf = []
        for fname, data in image_items:
            img = PILImage.open(_io.BytesIO(data)).convert("RGB")
            images_for_pdf.append(img)
        out = BytesIO()
        images_for_pdf[0].save(
            out, format="PDF", save_all=True,
            append_images=images_for_pdf[1:], resolution=150
        )
        return out.getvalue(), pdf_name
    except Exception:
        pass  # Pillow not installed or failed — use fallback

    # ── Fallback: wrap each JPEG directly into a minimal PDF ──────────────
    # Converts PNG/other to JPEG-compatible bytes first using struct
    try:
        def _to_jpeg_bytes(raw: bytes) -> Optional[bytes]:
            """Return JPEG bytes from any image format using Pillow if available,
            or return raw bytes if already JPEG."""
            if raw[:2] == b"\xff\xd8":
                return raw  # already JPEG
            try:
                from PIL import Image as _PI
                buf = _io.BytesIO()
                _PI.open(_io.BytesIO(raw)).convert("RGB").save(buf, format="JPEG", quality=85)
                return buf.getvalue()
            except Exception:
                return raw if raw[:2] == b"\xff\xd8" else None

        pdf_parts = []
        xref_offsets = []
        obj_count = 0

        # PDF header
        pdf_parts.append(b"%PDF-1.4\n")

        page_obj_ids = []
        image_obj_ids = []

        for fname, raw_data in image_items:
            jpg = _to_jpeg_bytes(raw_data)
            if not jpg:
                continue

            # Try to get image dimensions
            w, h = 595, 842  # A4 default
            try:
                from PIL import Image as _PI2
                im = _PI2.open(_io.BytesIO(raw_data))
                w, h = im.size
                # Scale to fit A4 width (595pt) if wider
                if w > 595:
                    h = int(h * 595 / w)
                    w = 595
            except Exception:
                pass

            obj_count += 1; img_id = obj_count
            obj_count += 1; page_id = obj_count

            image_obj_ids.append((img_id, jpg, w, h))
            page_obj_ids.append((page_id, img_id, w, h))

        if not image_obj_ids:
            return None

        obj_count += 1; pages_id = obj_count
        obj_count += 1; catalog_id = obj_count

        # Write image + page objects
        body = b""
        xrefs = {}

        for img_id, jpg, w, h in image_obj_ids:
            xrefs[img_id] = len(pdf_parts[0]) + len(body)
            body += (
                f"{img_id} 0 obj\n"
                f"<</Type /XObject /Subtype /Image /Width {w} /Height {h} "
                f"/ColorSpace /DeviceRGB /BitsPerComponent 8 "
                f"/Filter /DCTDecode /Length {len(jpg)}>>\nstream\n"
            ).encode() + jpg + b"\nendstream\nendobj\n"

        for page_id, img_id, w, h in page_obj_ids:
            xrefs[page_id] = len(pdf_parts[0]) + len(body)
            content_stream = f"q {w} 0 0 {h} 0 0 cm /Im{img_id} Do Q".encode()
            body += (
                f"{page_id} 0 obj\n"
                f"<</Type /Page /Parent {pages_id} 0 R "
                f"/MediaBox [0 0 {w} {h}] "
                f"/Resources <</XObject <</Im{img_id} {img_id} 0 R>>>> "
                f"/Contents {page_id+100} 0 R>>\nendobj\n"
            ).encode()
            # Content stream object
            xrefs[page_id + 100] = len(pdf_parts[0]) + len(body)
            body += (
                f"{page_id+100} 0 obj\n"
                f"<</Length {len(content_stream)}>>\nstream\n"
            ).encode() + content_stream + b"\nendstream\nendobj\n"

        # Pages object
        kids = " ".join(f"{pid} 0 R" for pid, *_ in page_obj_ids)
        xrefs[pages_id] = len(pdf_parts[0]) + len(body)
        body += (
            f"{pages_id} 0 obj\n"
            f"<</Type /Pages /Kids [{kids}] /Count {len(page_obj_ids)}>>\nendobj\n"
        ).encode()

        # Catalog
        xrefs[catalog_id] = len(pdf_parts[0]) + len(body)
        body += (
            f"{catalog_id} 0 obj\n"
            f"<</Type /Catalog /Pages {pages_id} 0 R>>\nendobj\n"
        ).encode()

        xref_pos = len(pdf_parts[0]) + len(body)
        all_ids = sorted(xrefs.keys())
        xref_table = f"xref\n0 1\n0000000000 65535 f \n{len(all_ids)+1} {max(all_ids)}\n"
        for oid in all_ids:
            xref_table += f"{xrefs[oid]:010d} 00000 n \n"

        trailer = (
            f"trailer\n<</Size {max(all_ids)+1} /Root {catalog_id} 0 R>>\n"
            f"startxref\n{xref_pos}\n%%EOF\n"
        )

        return pdf_parts[0] + body + xref_table.encode() + trailer.encode(), pdf_name

    except Exception:
        return None


def parse_msg(uploaded_file) -> Dict:
    import extract_msg
    with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    msg = extract_msg.Message(tmp_path)
    sender   = msg.sender or ""
    subject  = msg.subject or uploaded_file.name
    date     = str(msg.date) if msg.date else ""
    body     = clean_text(msg.body or "")
    if not body and getattr(msg, "htmlBody", None):
        body = html_to_text(msg.htmlBody)
    body = strip_icon_prefixes(body)
    full_text = clean_text("\n".join([f"From: {sender}", f"Subject: {subject}", body]))

    attachment_decisions = []
    kept_attachments = []
    for att in msg.attachments:
        fname = getattr(att, "longFilename", None) or getattr(att, "shortFilename", None) or ""
        if fname:
            d = decide_attachment(fname, att, subject=subject)
            attachment_decisions.append(d)
            if d["decision"] == "Keep":
                kept_attachments.append(d)

    # Generate lead PDF from kept attachments
    lead_pdf = make_lead_pdf(kept_attachments, date) if kept_attachments else None
    pdf_name = lead_pdf[1] if lead_pdf else ""

    return {
        "sender": sender,
        "subject": subject,
        "date": date,
        "body": body,
        "full_text": full_text,
        "attachment_decisions": attachment_decisions,
        "kept_attachments": kept_attachments,
        "valid_attachment_names": [d["filename"] for d in attachment_decisions if d["decision"] == "Keep"],
        "lead_pdf_bytes": lead_pdf[0] if lead_pdf else None,
        "lead_pdf_name": pdf_name,
    }

# ── Agent system prompt ────────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are the HYDAC Lead Agent. HYDAC is a US industrial filtration and hydraulics company.
You receive raw forwarded email threads and extract the external customer lead.

YOUR JOB:
Read the email thread carefully and identify the REAL external customer — the person who sent an inquiry TO HYDAC.
Ignore all HYDAC employees, internal forwarders, and anyone with a @hydacusa.com / @hydac.com / @hydac-interlynx.com email.
The customer is almost always the OLDEST message in the thread or the one below an "EXTERNAL EMAIL" marker.

EXTRACTION RULES (follow exactly):
- FirstName / LastName: split cleanly. Handle "Last, First" format. Remove titles (Mr/Mrs/Dr).
- ContactTitle: job title from signature only (e.g. "Purchasing Manager"). Leave blank if not found.
- Email: external customer email only. Never a HYDAC domain.
- Company: must be written in the email body or signature. Never infer from Gmail/free mail domains (gmail, yahoo, hotmail, etc.).
  If not explicitly written but a business email domain is present (e.g. fleetpride.com), derive the company name from the domain (e.g. "FleetPride").
- Address / City / State / ZipCode / Country: only what is written. Watch for split-line addresses (number on one line, street on next).
  Also watch for icon-prefixed lines like "<location icon> 8814\nDietz Ave. Hickory NC 28602".
- PhoneSupplied: customer phones only, formatted as "phone1 : phone2". Exclude HYDAC employee phones.
  Watch for icon-prefixed lines: "<phone icon>  828-328-1551" — the number after the icon is the phone.
- WebAddress: only explicit URLs or www. addresses written by the customer. Never derive from email domain.
  Watch for icon-prefixed lines: "<web icon>  www.example.com".
- LeadComments: copy the customer's request VERBATIM — do not paraphrase, summarise, or shorten anything.
  If the email contains a product/item table (description, part number, brand, quantity, measurements, reference numbers),
  format it using <br> line breaks and <b>bold</b> labels exactly like this example:
  "Please quote your best prices, delivery, weight, packing details for the listed items and send your offer per email.<br><br><b>Description</b> - FILTER ELEMENT 6 MICRON Measures: 40 MM 78 MM Height: 158 MM Ref: 1282779<br><b>Brand</b> - HYDAC<br><b>Part Number</b> - 0160 DN 006 BH4HC<br><b>Quantity</b> - 15 PCS<br><br>***PLEASE PROVIDE THE DATA SHEET OF THE PRODUCTS OFFERED***<br><br>Please don't forget to specify lead time, shipping point zip code, packaging dimensions and weight for freight quoting purposes."
  Remove ONLY: greetings (Dear/Hi/Good morning), sign-offs (Thank you/Regards/Best), the sender's signature block, and legal disclaimers.
  Keep everything else exactly as written including ALL caps text, asterisks, measurements, reference numbers, and special instructions.
- Product: HYDAC model number or product category. Examples: "2065706 / DF ON 240 TE 10 B M 1.0/12", "RF3-30", "Return Filter".
  7-digit numbers are often HYDAC part/order numbers. "DF ON ...", "RF ...", "RFBN ..." are model codes.
- Quantity: numeric quantity if specified (e.g. 7). Leave blank if not stated.
- Summary: write 2-4 sentences summarising the whole email thread for a sales rep. Include: who the customer is, what they are asking for, any urgency or context, and any references (PO numbers, existing orders, prior contact). Do NOT copy-paste the email — write it as a concise briefing.

WHAT TO IGNORE:
- "Best Regards, Brandon Huertas, Web Content Administrator" — this is an internal HYDAC forwarder, not the customer.
- Microsoft safety warnings ("You don't often get email from..."), disclaimer footers, confidentiality notices.
- Inline image icon URLs like <https://www.example.com/images/phone.png> — these are just decorators.

OUTPUT:
Return ONLY a valid JSON object with exactly these keys (leave unknown fields as empty string ""):
{
  "FirstName": "",
  "LastName": "",
  "ContactTitle": "",
  "Email": "",
  "Company": "",
  "Address": "",
  "City": "",
  "State": "",
  "ZipCode": "",
  "Country": "",
  "PhoneSupplied": "",
  "WebAddress": "",
  "LeadComments": "",
  "Summary": "2-4 sentence summary of the full email thread: who sent it, what they need, any context (urgency, existing relationship, references). Written for a sales rep who hasn't read the email.",
  "Product": "",
  "Quantity": "",
  "AgentReason": "one sentence explaining how you identified the customer"
}
No markdown, no preamble, no explanation outside the JSON."""

# ── AI call ────────────────────────────────────────────────────────────────────

def run_agent(email_text: str, provider: str, api_key: str = "") -> Dict:
    # Always re-read from secrets in case sidebar variable is stale
    if provider == "groq":
        api_key = api_key or st.secrets.get("GROQ_API_KEY", "")
    elif provider == "claude":
        api_key = api_key or st.secrets.get("ANTHROPIC_API_KEY", "")
    elif provider == "openai":
        api_key = api_key or st.secrets.get("OPENAI_API_KEY", "")
    elif provider == "gemini":
        api_key = api_key or st.secrets.get("GEMINI_API_KEY", "")

    user_prompt = f"""Extract the lead from this email thread:

{email_text[:12000]}"""

    raw = ""

    # ── Ollama (local, no API key) ─────────────────────────────────────────
    if provider == "ollama":
        import urllib.request
        payload = json.dumps({
            "model": "llama3.1:8b",
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": user_prompt},
            ],
            "stream": False,
            "options": {"temperature": 0},
        }).encode()
        req = urllib.request.Request(
            "http://localhost:11434/api/chat",
            data=payload,
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=120) as resp:
            data = json.loads(resp.read())
        raw = data["message"]["content"].strip()

    # ── Groq (free cloud, needs free key from console.groq.com) ───────────
    elif provider == "groq":
        from openai import OpenAI  # Groq is OpenAI-compatible
        client = OpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")
        response = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            temperature=0,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": user_prompt},
            ],
        )
        raw = response.choices[0].message.content.strip()

    # ── Gemini (free tier, needs free key from aistudio.google.com) ───────
    elif provider == "gemini":
        import google.generativeai as genai
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(
            model_name="gemini-1.5-flash",
            system_instruction=SYSTEM_PROMPT,
        )
        response = model.generate_content(user_prompt)
        raw = response.text.strip()

    # ── Claude ─────────────────────────────────────────────────────────────
    elif provider == "claude":
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1200,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_prompt}],
        )
        raw = response.content[0].text.strip()

    # ── OpenAI ─────────────────────────────────────────────────────────────
    elif provider == "openai":
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model="gpt-4o",
            temperature=0,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": user_prompt},
            ],
        )
        raw = response.choices[0].message.content.strip()

    else:
        raise ValueError(f"Unknown provider: {provider}")

    raw = raw.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
    return json.loads(raw)

# ── Excel export ───────────────────────────────────────────────────────────────

def make_excel(row: Dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    ws.append(EXCEL_HEADER)
    ws.append([row.get(h, "") for h in EXCEL_HEADER])
    for col in ws.columns:
        letter = col[0].column_letter
        width = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[letter].width = min(width + 2, 45)
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

def make_attachment_zip(kept: List[Dict]) -> Optional[bytes]:
    if not kept:
        return None
    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for d in kept:
            try:
                data = d["obj"].data
                if data:
                    z.writestr(Path(d["filename"]).name, data)
            except Exception:
                pass
    return out.getvalue()

# ── Research engine ────────────────────────────────────────────────────────────

RESEARCH_SYSTEM_PROMPT = """You are a B2B sales research assistant for HYDAC, an industrial filtration company.
Given web search results about a company and contact person, extract structured data.
Return ONLY valid JSON — no markdown, no preamble.

Fill every field you can find evidence for. Leave fields as empty string "" if not found.
Do NOT guess or invent data. Only use what the search results clearly show.

JSON keys to return:
{
  "LineOfBusiness": "industry/sector the company operates in (e.g. Fleet Services, Oil & Gas, Manufacturing)",
  "noOfEmployees": "employee count or range as a string (e.g. '500-1000' or '5,000+')",
  "ParentName": "parent company name if this is a subsidiary, else empty",
  "WebAddress": "company website if found and not already known",
  "SIC": "4-digit SIC code if found",
  "NAICS": "6-digit NAICS code if found",
  "Linkedin_Title": "contact's exact job title from LinkedIn",
  "Linkedin_Link": "full LinkedIn profile URL of the contact",
  "linkedin_city": "city from contact's LinkedIn location",
  "linkedin_state": "state/region from contact's LinkedIn location",
  "linkedin_country": "country from contact's LinkedIn location",
  "college_1": "first university/college attended",
  "college_1_degree": "degree from college_1",
  "college_1_start": "start year at college_1",
  "college_1_end": "end year at college_1",
  "college_2": "second university/college if any",
  "college_2_degree": "degree from college_2",
  "college_2_start": "start year at college_2",
  "college_2_end": "end year at college_2",
  "about_experience": "1-2 sentence summary of contact's career background from LinkedIn",
  "about_me": "company description in 1-2 sentences",
  "PhoneResearched": "company main phone number if found",
  "ResearchNotes": "2-3 sentences summarising what was found about the company and contact for the sales rep"
}"""


def web_search_tavily(query: str, api_key: str, max_results: int = 5) -> List[Dict]:
    """Search using Tavily API (free tier: 1000/month). Sign up at tavily.com."""
    import urllib.request, urllib.parse
    payload = json.dumps({
        "api_key": api_key,
        "query": query,
        "max_results": max_results,
        "search_depth": "basic",
    }).encode()
    req = urllib.request.Request(
        "https://api.tavily.com/search",
        data=payload,
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        data = json.loads(resp.read())
    return [
        {"title": r.get("title",""), "url": r.get("url",""), "body": r.get("content","")}
        for r in data.get("results", [])
    ]


def web_search_groq(queries: List[str], api_key: str) -> str:
    """Use Groq llama with web_search tool — searches happen server-side on Groq."""
    from openai import OpenAI
    client = OpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")
    combined = "\n".join(f"- {q}" for q in queries)
    response = client.chat.completions.create(
        model="llama-3.1-8b-instant",
        temperature=0,
        messages=[{
            "role": "user",
            "content": f"Search the web for each of these queries and return all findings as plain text:\n{combined}"
        }],
        tools=[{"type": "function", "function": {
            "name": "web_search",
            "description": "Search the web",
            "parameters": {"type": "object", "properties": {
                "query": {"type": "string"}
            }, "required": ["query"]}
        }}],
    )
    # Collect text from all content blocks
    texts = []
    for choice in [response.choices[0]]:
        if choice.message.content:
            texts.append(choice.message.content)
    return "\n".join(texts) or "No results from Groq search."


def format_search_results(results: List[Dict]) -> str:
    parts = []
    for i, r in enumerate(results, 1):
        parts.append(f"[{i}] {r['title']}\n{r['url']}\n{r['body']}")
    return "\n\n".join(parts)


def run_research(lead: Dict, provider: str, api_key: str = "", tavily_key: str = "") -> tuple:
    """Run web searches about the company and contact, then use AI to synthesise results.
    Returns (result_dict, search_debug_text)."""
    # Always re-read keys from secrets in case sidebar variable is stale
    if provider == "groq":
        api_key = api_key or st.secrets.get("GROQ_API_KEY", "")
    elif provider == "claude":
        api_key = api_key or st.secrets.get("ANTHROPIC_API_KEY", "")
    elif provider == "openai":
        api_key = api_key or st.secrets.get("OPENAI_API_KEY", "")
    elif provider == "gemini":
        api_key = api_key or st.secrets.get("GEMINI_API_KEY", "")
    tavily_key = tavily_key or st.secrets.get("TAVILY_API_KEY", "")

    company   = lead.get("Company", "").strip()
    firstname = lead.get("FirstName", "").strip()
    lastname  = lead.get("LastName", "").strip()
    full_name = f"{firstname} {lastname}".strip()
    website   = lead.get("WebAddress", "").strip()

    if not company and not full_name:
        return {"ResearchNotes": "Not enough information to research (no company or contact name)."}, ""

    # Build search queries
    queries = []
    if company:
        queries.append(f"{company} company industry employees headquarters")
        queries.append(f"{company} SIC NAICS industry code")
    if full_name and company:
        queries.append(f"{full_name} {company} LinkedIn")
    elif full_name:
        queries.append(f"{full_name} LinkedIn")
    if not website and company:
        queries.append(f"{company} official website")

    search_context = ""
    search_debug   = ""

    # ── Try Tavily first (works on Streamlit Cloud, free tier) ────────────
    tavily_key = tavily_key or st.secrets.get("TAVILY_API_KEY", "")
    if tavily_key:
        try:
            parts = []
            for q in queries:
                results = web_search_tavily(q, tavily_key, max_results=4)
                parts.append(f"=== {q.upper()} ===\n{format_search_results(results)}")
            search_context = "\n\n".join(parts)
            search_debug = f"[Tavily] {len(queries)} queries OK\n\n" + search_context
        except Exception as e:
            search_debug = f"[Tavily failed: {e}]"

    # ── Groq native search (no extra key needed if provider=groq) ─────────
    if not search_context and provider == "groq" and api_key:
        try:
            search_context = web_search_groq(queries, api_key)
            search_debug = f"[Groq web search]\n{search_context}"
        except Exception as e:
            search_debug += f"\n[Groq search failed: {e}]"

    # ── If still nothing, ask AI to reason from what it already knows ─────
    if not search_context:
        search_context = (
            f"No live search results available. "
            f"Use your training knowledge to fill what you can about: "
            f"{company} ({website}) and {full_name}. "
            f"Be conservative — only state what you are confident about."
        )
        search_debug += "\n[No live search — AI using training knowledge]"

    user_prompt = f"""Research target:
- Company: {company or "unknown"}
- Contact: {full_name or "unknown"}
- Known website: {website or "not known"}
- Known location: {lead.get("City","")} {lead.get("State","")} {lead.get("Country","")}

SEARCH RESULTS:
{search_context[:10000]}

Extract all fields you can find evidence for."""

    # ── Call AI to synthesise ──────────────────────────────────────────────
    raw = ""
    try:
        if provider == "ollama":
            import urllib.request
            payload = json.dumps({
                "model": "llama3.1:8b",
                "messages": [
                    {"role": "system", "content": RESEARCH_SYSTEM_PROMPT},
                    {"role": "user",   "content": user_prompt},
                ],
                "stream": False,
                "options": {"temperature": 0},
            }).encode()
            req = urllib.request.Request(
                "http://localhost:11434/api/chat",
                data=payload,
                headers={"Content-Type": "application/json"},
            )
            with urllib.request.urlopen(req, timeout=120) as resp:
                data = json.loads(resp.read())
            raw = data["message"]["content"].strip()

        elif provider == "groq":
            from openai import OpenAI
            client = OpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")
            response = client.chat.completions.create(
                model="llama-3.1-8b-instant",
                temperature=0,
                messages=[
                    {"role": "system", "content": RESEARCH_SYSTEM_PROMPT},
                    {"role": "user",   "content": user_prompt},
                ],
            )
            raw = response.choices[0].message.content.strip()

        elif provider == "gemini":
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(
                model_name="gemini-1.5-flash",
                system_instruction=RESEARCH_SYSTEM_PROMPT,
            )
            response = model.generate_content(user_prompt)
            raw = response.text.strip()

        elif provider == "claude":
            import anthropic
            client = anthropic.Anthropic(api_key=api_key)
            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=1200,
                system=RESEARCH_SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_prompt}],
            )
            raw = response.content[0].text.strip()

        elif provider == "openai":
            from openai import OpenAI
            client = OpenAI(api_key=api_key)
            response = client.chat.completions.create(
                model="gpt-4o",
                temperature=0,
                messages=[
                    {"role": "system", "content": RESEARCH_SYSTEM_PROMPT},
                    {"role": "user",   "content": user_prompt},
                ],
            )
            raw = response.choices[0].message.content.strip()

        raw = raw.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
        return json.loads(raw), search_context

    except Exception as e:
        return {"ResearchNotes": f"Research failed: {e}"}, search_context


# ── UI helpers ─────────────────────────────────────────────────────────────────

def field_row(label: str, value: str, key: str, multiline: bool = False) -> str:
    """Render a labelled field — display mode if not editing, input if editing."""
    if st.session_state.get("editing"):
        if multiline:
            return st.text_area(label, value=value, key=key, height=120)
        return st.text_input(label, value=value, key=key)
    else:
        st.markdown(f"**{label}**")
        st.markdown(value or "*—*")
        return value

# ── Main UI ────────────────────────────────────────────────────────────────────

st.title("⚙️ HYDAC Lead Agent")
st.caption("Upload a .msg file — the agent reads it and gives you a clean lead card ready to export.")

# Default values — overwritten inside sidebar block
api_key    = ""
tavily_key = st.secrets.get("TAVILY_API_KEY", "")

# Sidebar: API keys + provider
with st.sidebar:
    st.header("Settings")
    PROVIDER_LABELS = {
        "ollama": "🖥️ Ollama — Local (no key needed)",
        "groq":   "🆓 Groq — Free cloud (sign up at groq.com)",
        "gemini": "🆓 Gemini Flash (sign up at aistudio.google.com)",
        "claude": "Claude Sonnet 4.6 (paid)",
        "openai": "OpenAI GPT-4o (paid)",
    }
    provider = st.radio(
        "AI Provider",
        list(PROVIDER_LABELS.keys()),
        format_func=lambda x: PROVIDER_LABELS[x],
    )

    api_key = ""

    if provider == "ollama":
        st.info(
            "**No API key needed.**\n\n"
            "Make sure Ollama is running locally:\n"
            "1. Download from [ollama.com](https://ollama.com)\n"
            "2. Run: `ollama pull llama3.1:8b`\n"
            "3. Ollama starts automatically on port 11434"
        )

    elif provider == "groq":
        api_key = st.secrets.get("GROQ_API_KEY", "") or st.text_input(
            "Groq API Key", type="password",
            placeholder="gsk_...",
            help="Free at console.groq.com → API Keys",
        )
        if not api_key:
            st.info("Free key at [console.groq.com](https://console.groq.com) — no credit card")
        else:
            st.success("Groq key ready")

    elif provider == "gemini":
        api_key = st.secrets.get("GEMINI_API_KEY", "") or st.text_input(
            "Gemini API Key", type="password",
            placeholder="AIza...",
            help="Free at aistudio.google.com → Get API key",
        )
        if not api_key:
            st.info("Free key at [aistudio.google.com](https://aistudio.google.com)")
        else:
            st.success("Gemini key ready")

    elif provider == "claude":
        api_key = st.secrets.get("ANTHROPIC_API_KEY", "") or st.text_input(
            "Anthropic API Key", type="password", placeholder="sk-ant-...")
        if api_key:
            st.success("API key ready")
        else:
            st.warning("Enter your Anthropic API key")

    elif provider == "openai":
        api_key = st.secrets.get("OPENAI_API_KEY", "") or st.text_input(
            "OpenAI API Key", type="password", placeholder="sk-...")
        if api_key:
            st.success("API key ready")
        else:
            st.warning("Enter your OpenAI API key")

    st.divider()
    st.caption("Ollama & Groq are completely free. Gemini has a generous free tier.")

    st.subheader("🔍 Research (Web Search)")
    tavily_key = st.secrets.get("TAVILY_API_KEY", "") or st.text_input(
        "Tavily API Key (for Research)",
        type="password",
        placeholder="tvly-...",
        help="Free at tavily.com — 1000 searches/month. Required for the Research button.",
    )
    if tavily_key:
        st.success("Tavily ready — Research button enabled")
    else:
        if provider == "groq" and api_key:
            st.info("No Tavily key — Research will use Groq's knowledge instead. "
                    "For live web results get a free key at [tavily.com](https://tavily.com)")
        else:
            st.warning("Get a free key at [tavily.com](https://tavily.com) to enable live web research.")

# File upload
uploaded_file = st.file_uploader("Upload .msg file", type=["msg"])

# Ollama needs no key; all others need one
ready = (provider == "ollama") or bool(api_key)

if uploaded_file and not ready:
    st.warning("Add your API key in the sidebar to run the agent.")

if uploaded_file and ready:
    file_key = uploaded_file.name + str(uploaded_file.size)

    # Only re-run agent if a new file is uploaded
    if st.session_state.get("file_key") != file_key:
        st.session_state.file_key = file_key
        st.session_state.editing = False
        st.session_state.result = None
        st.session_state.parsed = None
        st.session_state.researched = False

        with st.spinner("Agent is reading the email…"):
            try:
                parsed = parse_msg(uploaded_file)
                st.session_state.parsed = parsed
                result = run_agent(parsed["full_text"], provider, api_key)
                # Inject fields the AI doesn't produce
                result["ReceivedDateTime"] = parsed["date"]
                result["LeadSource1"] = "Email"
                # Use auto-generated PDF name if we bundled attachments, else fallback
                if parsed.get("lead_pdf_name"):
                    result["PDF"] = parsed["lead_pdf_name"]
                else:
                    result["PDF"] = Path(uploaded_file.name).with_suffix(".pdf").name
                st.session_state.result = result
            except Exception as e:
                st.error(f"Agent failed: {e}")
                st.exception(e)

    result = st.session_state.get("result")
    parsed = st.session_state.get("parsed")

    if result:
        agent_reason = result.get("AgentReason", "")

        # ── Summary card ──────────────────────────────────────────────────────
        st.subheader("Lead Summary")
        if agent_reason:
            st.info(f"🤖 {agent_reason}")

        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown("#### Contact")
            name = f"{result.get('FirstName','')} {result.get('LastName','')}".strip()
            st.markdown(f"**{name or '—'}**")
            if result.get("ContactTitle"):
                st.markdown(result["ContactTitle"])
            st.markdown(result.get("Email") or "—")
            if result.get("PhoneSupplied"):
                st.markdown(f"📞 {result['PhoneSupplied']}")

        with col2:
            st.markdown("#### Company & Location")
            st.markdown(f"**{result.get('Company') or '—'}**")
            addr_parts = [
                result.get("Address",""),
                " ".join(filter(None,[result.get("City",""), result.get("State",""), result.get("ZipCode","")])),
                result.get("Country",""),
            ]
            for p in addr_parts:
                if p.strip():
                    st.markdown(p)
            if result.get("WebAddress"):
                st.markdown(f"🌐 {result['WebAddress']}")

        with col3:
            st.markdown("#### Request")
            if result.get("Product"):
                st.markdown(f"**Product:** {result['Product']}")
            if result.get("Quantity"):
                st.markdown(f"**Qty:** {result['Quantity']}")
            comments = result.get("LeadComments","")
            if comments:
                # Convert HTML tags to markdown for display
                display = re.sub(r"<br\s*/?>", "\n", comments, flags=re.I)
                display = re.sub(r"<b>(.*?)</b>", r"**\1**", display, flags=re.I)
                display = re.sub(r"<[^>]+>", "", display)
                display = re.sub(r"<[^>]+>", "", display)  # strip any remaining tags
                display = display[:500] + ("…" if len(display) > 500 else "")
                st.markdown(display)
            else:
                st.markdown("—")

        if result.get("Summary"):
            st.markdown("#### 📋 Email Summary")
            st.info(result["Summary"])

        st.divider()

        # ── Edit / Export row ─────────────────────────────────────────────────
        edit_col, export_col, zip_col = st.columns([1, 1, 1])

        with edit_col:
            if st.button("✏️ Edit fields" if not st.session_state.get("editing") else "✅ Done editing"):
                st.session_state.editing = not st.session_state.get("editing", False)
                st.rerun()

        # ── Edit form ─────────────────────────────────────────────────────────
        if st.session_state.get("editing"):
            with st.form("edit_form"):
                st.markdown("#### Edit Fields")
                c1, c2 = st.columns(2)
                with c1:
                    first    = st.text_input("First Name",     value=result.get("FirstName",""))
                    last     = st.text_input("Last Name",      value=result.get("LastName",""))
                    title    = st.text_input("Title",          value=result.get("ContactTitle",""))
                    email    = st.text_input("Email",          value=result.get("Email",""))
                    company  = st.text_input("Company",        value=result.get("Company",""))
                    phone    = st.text_input("Phone",          value=result.get("PhoneSupplied",""))
                    website  = st.text_input("Website",        value=result.get("WebAddress",""))
                with c2:
                    address  = st.text_input("Address",        value=result.get("Address",""))
                    city     = st.text_input("City",           value=result.get("City",""))
                    state    = st.text_input("State",          value=result.get("State",""))
                    zipcode  = st.text_input("Zip Code",       value=result.get("ZipCode",""))
                    country  = st.text_input("Country",        value=result.get("Country",""))
                    product  = st.text_input("Product",        value=result.get("Product",""))
                    quantity = st.text_input("Quantity",       value=result.get("Quantity",""))
                comments = st.text_area("Lead Comments / Request", value=result.get("LeadComments",""), height=150)
                summary  = st.text_area("Summary", value=result.get("Summary",""), height=100)

                if st.form_submit_button("💾 Save changes"):
                    result.update({
                        "FirstName": first, "LastName": last, "ContactTitle": title,
                        "Email": email, "Company": company, "PhoneSupplied": phone,
                        "WebAddress": website, "Address": address, "City": city,
                        "State": state, "ZipCode": zipcode, "Country": country,
                        "Product": product, "Quantity": quantity, "LeadComments": comments,
                        "Summary": summary,
                    })
                    st.session_state.result = result
                    st.session_state.editing = False
                    st.success("Saved!")
                    st.rerun()

        # ── Export buttons ────────────────────────────────────────────────────
        row = {h: "" for h in EXCEL_HEADER}
        row.update({k: result.get(k, "") for k in EXCEL_HEADER if k in result})
        excel_bytes = make_excel(row)

        with export_col:
            st.download_button(
                "📥 Download Excel",
                data=excel_bytes,
                file_name=f"hydac_lead_{Path(uploaded_file.name).stem}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        if parsed and parsed.get("lead_pdf_bytes"):
            with zip_col:
                st.download_button(
                    "📄 Download Lead PDF",
                    data=parsed["lead_pdf_bytes"],
                    file_name=parsed.get("lead_pdf_name", "lead.pdf"),
                    mime="application/pdf",
                )

        # ── Research section ──────────────────────────────────────────────────
        st.divider()
        res_col, _ = st.columns([1, 2])
        with res_col:
            research_btn = st.button(
                "🔍 Research this lead", use_container_width=True,
                help="Searches the web for company info, SIC/NAICS codes, employee count, LinkedIn profile and more.",
                key="research_btn",
            )

        if research_btn:
            with st.spinner("Searching the web and analysing results…"):
                try:
                    research, search_debug = run_research(result, provider, api_key, tavily_key=tavily_key)
                    st.session_state.search_debug = search_debug
                    # Merge all research fields into session result
                    RESEARCH_FIELDS = [
                        "ResearchNotes", "LineOfBusiness", "noOfEmployees", "ParentName",
                        "SIC", "NAICS", "WebAddress", "PhoneResearched",
                        "Linkedin_Title", "Linkedin_Link",
                        "linkedin_city", "linkedin_state", "linkedin_country",
                        "college_1", "college_1_degree", "college_1_start", "college_1_end",
                        "college_2", "college_2_degree", "college_2_start", "college_2_end",
                        "about_experience", "about_me",
                    ]
                    for k in RESEARCH_FIELDS:
                        if research.get(k):
                            result[k] = research[k]
                    st.session_state.result = result
                    st.session_state.researched = True
                except Exception as e:
                    st.error(f"Research failed: {e}")
                    st.exception(e)

        # Show research results — reads from session_state.result so persists across reruns
        if st.session_state.get("researched"):
            result = st.session_state.result  # always use latest
            st.markdown("### 🔍 Research Results")

            if result.get("ResearchNotes"):
                st.success(result["ResearchNotes"])

            r1, r2, r3 = st.columns(3)
            with r1:
                st.markdown("**🏢 Company**")
                if result.get("LineOfBusiness"):
                    st.markdown(f"**Industry:** {result['LineOfBusiness']}")
                if result.get("noOfEmployees"):
                    st.markdown(f"**Employees:** {result['noOfEmployees']}")
                if result.get("ParentName"):
                    st.markdown(f"**Parent:** {result['ParentName']}")
                if result.get("PhoneResearched"):
                    st.markdown(f"**Phone:** {result['PhoneResearched']}")
                if result.get("about_me"):
                    st.caption(result["about_me"])
            with r2:
                st.markdown("**📊 Industry Codes**")
                if result.get("SIC"):
                    st.markdown(f"**SIC:** `{result['SIC']}`")
                else:
                    st.markdown("SIC: *not found*")
                if result.get("NAICS"):
                    st.markdown(f"**NAICS:** `{result['NAICS']}`")
                else:
                    st.markdown("NAICS: *not found*")
                if result.get("WebAddress"):
                    st.markdown(f"**Web:** {result['WebAddress']}")
            with r3:
                st.markdown("**👤 Contact (LinkedIn)**")
                if result.get("Linkedin_Title"):
                    st.markdown(f"**Title:** {result['Linkedin_Title']}")
                if result.get("Linkedin_Link"):
                    st.markdown(f"[🔗 LinkedIn Profile]({result['Linkedin_Link']})")
                loc = " ".join(filter(None, [
                    result.get("linkedin_city",""),
                    result.get("linkedin_state",""),
                    result.get("linkedin_country",""),
                ]))
                if loc:
                    st.markdown(f"**Location:** {loc}")
                if result.get("about_experience"):
                    st.caption(result["about_experience"])

            if any(result.get(k) for k in ["college_1", "college_2"]):
                with st.expander("🎓 Education"):
                    for prefix in ["college_1", "college_2"]:
                        name = result.get(prefix, "")
                        if name:
                            deg   = result.get(f"{prefix}_degree", "")
                            start = result.get(f"{prefix}_start", "")
                            end   = result.get(f"{prefix}_end", "")
                            st.markdown(f"**{name}** — {deg} ({start}–{end})")

        # ── Debug expanders ───────────────────────────────────────────────────
        with st.expander("Raw email text (debug)"):
            st.text(parsed["full_text"][:15000] if parsed else "")

        if parsed:
            with st.expander("📎 Attachment debug"):
                st.write(f"**Total attachments:** {len(parsed.get('attachment_decisions', []))}")
                st.write(f"**Kept:** {len(parsed.get('kept_attachments', []))}")
                st.write(f"**PDF generated:** {bool(parsed.get('lead_pdf_bytes'))}")
                st.write(f"**PDF name:** {parsed.get('lead_pdf_name', 'none')}")
                for d in parsed.get("attachment_decisions", []):
                    st.write(f"- `{d['filename']}` -> **{d['decision']}** ({d['reason']})")

        if st.session_state.get("search_debug"):
            with st.expander("🔍 Raw search results (debug)"):
                st.text(st.session_state.search_debug[:8000])
