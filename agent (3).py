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

def decide_attachment(filename: str, attachment_obj=None):
    if not filename:
        return {"filename": "", "decision": "Reject", "reason": "Blank filename"}
    name = Path(filename).name
    low = name.lower()
    ext = Path(low).suffix
    if ext not in VALID_ATTACHMENT_EXTENSIONS:
        return {"filename": name, "decision": "Reject", "reason": f"Unsupported extension {ext or '(none)'}"}
    if low in SIGNATURE_IMAGE_NAMES:
        return {"filename": name, "decision": "Reject", "reason": "Signature/logo image"}
    if re.fullmatch(r"image\d{0,3}\.(png|jpg|jpeg|gif|tif|tiff)", low):
        return {"filename": name, "decision": "Reject", "reason": "Generic inline signature image"}
    if ext not in IMAGE_EXTENSIONS:
        return {"filename": name, "decision": "Keep", "reason": "Document/CAD/spreadsheet attachment", "obj": attachment_obj}
    if re.search(r"\d{4,}", name):
        return {"filename": name, "decision": "Keep", "reason": "Image has part/model number", "obj": attachment_obj}
    if re.search(r"(?i)(pump|filter|drawing|label|plate|model|part|hydraulic|spec|quote|serial)", name):
        return {"filename": name, "decision": "Keep", "reason": "Image has product keyword", "obj": attachment_obj}
    return {"filename": name, "decision": "Reject", "reason": "Image not customer evidence"}

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
            d = decide_attachment(fname, att)
            attachment_decisions.append(d)
            if d["decision"] == "Keep":
                kept_attachments.append(d)

    return {
        "sender": sender,
        "subject": subject,
        "date": date,
        "body": body,
        "full_text": full_text,
        "attachment_decisions": attachment_decisions,
        "kept_attachments": kept_attachments,
        "valid_attachment_names": [d["filename"] for d in attachment_decisions if d["decision"] == "Keep"],
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
- LeadComments: the customer's actual request/inquiry text only. No greetings, no signatures, no forwarding trail.
  Include product model numbers and quantities (e.g. "Looking for (7) 2065706 DF ON 240 TE 10 B M 1.0/12 high-pressure in-line industrial filter.")
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


def web_search(query: str, max_results: int = 5) -> List[Dict]:
    """Run a DuckDuckGo search and return results as list of {title, url, body}."""
    try:
        from duckduckgo_search import DDGS
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=max_results))
        return [{"title": r.get("title",""), "url": r.get("href",""), "body": r.get("body","")} for r in results]
    except Exception as e:
        return [{"title": "Search error", "url": "", "body": str(e)}]


def format_search_results(results: List[Dict]) -> str:
    parts = []
    for i, r in enumerate(results, 1):
        parts.append(f"[{i}] {r['title']}\n{r['url']}\n{r['body']}")
    return "\n\n".join(parts)


def run_research(lead: Dict, provider: str, api_key: str = "") -> Dict:
    """Run web searches about the company and contact, then use AI to synthesise results."""
    company   = lead.get("Company", "").strip()
    firstname = lead.get("FirstName", "").strip()
    lastname  = lead.get("LastName", "").strip()
    full_name = f"{firstname} {lastname}".strip()
    website   = lead.get("WebAddress", "").strip()

    if not company and not full_name:
        return {"ResearchNotes": "Not enough information to research (no company or contact name)."}

    # ── Run 4 targeted searches ────────────────────────────────────────────
    searches = {}

    if company:
        searches["company_overview"] = web_search(
            f"{company} company industry employees headquarters description", max_results=5
        )
        searches["company_codes"] = web_search(
            f"{company} SIC NAICS industry code", max_results=4
        )
        if not website:
            searches["company_website"] = web_search(
                f"{company} official website", max_results=3
            )

    if full_name and company:
        searches["linkedin_contact"] = web_search(
            f"{full_name} {company} LinkedIn", max_results=5
        )
    elif full_name:
        searches["linkedin_contact"] = web_search(
            f"{full_name} LinkedIn profile", max_results=5
        )

    # ── Format all results for the AI ─────────────────────────────────────
    context_parts = []
    for label, results in searches.items():
        context_parts.append(f"=== {label.upper().replace('_', ' ')} ===\n{format_search_results(results)}")
    search_context = "\n\n".join(context_parts)

    user_prompt = f"""Research target:
- Company: {company or "unknown"}
- Contact: {full_name or "unknown"}
- Known website: {website or "not known"}
- Known location: {lead.get("City","")} {lead.get("State","")} {lead.get("Country","")}

SEARCH RESULTS:
{search_context[:10000]}

Extract all fields you can find evidence for in the search results."""

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
        return json.loads(raw)

    except Exception as e:
        return {"ResearchNotes": f"Research failed: {e}"}


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
                result["PDF"] = Path(uploaded_file.name).with_suffix(".pdf").name
                if parsed["valid_attachment_names"]:
                    result["PDF"] += "; Attachments: " + ", ".join(parsed["valid_attachment_names"])
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
            st.markdown(comments[:400] + ("…" if len(comments) > 400 else "") if comments else "—")

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

        if parsed and parsed.get("kept_attachments"):
            zip_bytes = make_attachment_zip(parsed["kept_attachments"])
            if zip_bytes:
                with zip_col:
                    st.download_button(
                        "📎 Download Attachments ZIP",
                        data=zip_bytes,
                        file_name="attachments.zip",
                        mime="application/zip",
                    )

        # ── Research section ──────────────────────────────────────────────────
        st.divider()
        res_col, _ = st.columns([1, 2])
        with res_col:
            research_btn = st.button("🔍 Research this lead", use_container_width=True,
                help="Searches the web for company info, SIC/NAICS codes, employee count, LinkedIn profile and more.")

        if research_btn:
            with st.spinner("Researching company and contact…"):
                research = run_research(result, provider, api_key)
                # Merge research results into the lead result (don't overwrite existing values)
                for k, v in research.items():
                    if v and not result.get(k):
                        result[k] = v
                # Always update research-specific fields even if result had a value
                for k in ["ResearchNotes", "noOfEmployees", "ParentName", "SIC", "NAICS",
                          "LineOfBusiness", "Linkedin_Title", "Linkedin_Link",
                          "linkedin_city", "linkedin_state", "linkedin_country",
                          "college_1", "college_1_degree", "college_1_start", "college_1_end",
                          "college_2", "college_2_degree", "college_2_start", "college_2_end",
                          "about_experience", "about_me", "PhoneResearched"]:
                    if research.get(k):
                        result[k] = research[k]
                st.session_state.result = result
                st.session_state.researched = True
                st.rerun()

        if st.session_state.get("researched") and result.get("ResearchNotes"):
            st.markdown("### 🔍 Research Results")
            st.success(result["ResearchNotes"])

            r1, r2, r3 = st.columns(3)
            with r1:
                st.markdown("**Company**")
                if result.get("LineOfBusiness"):
                    st.markdown(f"Industry: {result['LineOfBusiness']}")
                if result.get("noOfEmployees"):
                    st.markdown(f"Employees: {result['noOfEmployees']}")
                if result.get("ParentName"):
                    st.markdown(f"Parent: {result['ParentName']}")
                if result.get("about_me"):
                    st.caption(result["about_me"])
            with r2:
                st.markdown("**Codes**")
                if result.get("SIC"):
                    st.markdown(f"SIC: `{result['SIC']}`")
                if result.get("NAICS"):
                    st.markdown(f"NAICS: `{result['NAICS']}`")
                if result.get("PhoneResearched"):
                    st.markdown(f"📞 {result['PhoneResearched']}")
            with r3:
                st.markdown("**Contact (LinkedIn)**")
                if result.get("Linkedin_Title"):
                    st.markdown(f"Title: {result['Linkedin_Title']}")
                if result.get("Linkedin_Link"):
                    st.markdown(f"[LinkedIn Profile]({result['Linkedin_Link']})")
                loc = " ".join(filter(None, [result.get("linkedin_city",""), result.get("linkedin_state",""), result.get("linkedin_country","")]))
                if loc:
                    st.markdown(f"📍 {loc}")
                if result.get("about_experience"):
                    st.caption(result["about_experience"])

            if any(result.get(k) for k in ["college_1","college_2"]):
                with st.expander("🎓 Education"):
                    for prefix in ["college_1", "college_2"]:
                        name = result.get(prefix,"")
                        if name:
                            deg   = result.get(f"{prefix}_degree","")
                            start = result.get(f"{prefix}_start","")
                            end   = result.get(f"{prefix}_end","")
                            st.markdown(f"**{name}** — {deg} ({start}–{end})")

        # ── Debug expander ────────────────────────────────────────────────────
        with st.expander("Raw email text (debug)"):
            st.text(parsed["full_text"][:15000] if parsed else "")


