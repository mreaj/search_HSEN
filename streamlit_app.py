"""
HSE Safety Notifications RAG Chatbot
─────────────────────────────────────────────────────────────────────────────
Stack  : Mistral AI · LangChain · ChromaDB (in-memory) · Streamlit
Source : OneDrive / SharePoint public shared-folder link  ← primary
         Manual upload (PDF / DOCX fallback)

Document parsing:
  • pdfplumber  — text + tables (markdown grids)
  • pytesseract — OCR for scanned/image-only pages
  • pdf2image   — rasterises pages for OCR
  • python-docx — paragraphs + table cells

Every answer highlights matched query keywords inline.
All answers cite their source document and page number.
"""

import os, io, re, time, logging, base64
from pathlib import Path
import streamlit as st

st.set_page_config(
    page_title="HSE Notifications Bot",
    page_icon="🦺",
    layout="wide",
    initial_sidebar_state="expanded",
)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ─── Constants ────────────────────────────────────────────────────────────────
MISTRAL_MODELS = {
    "Mistral Small (fast)": "mistral-small-latest",
    "Mistral Medium":       "mistral-medium-latest",
    "Mixtral 8x7B":         "open-mixtral-8x7b",
    "Mistral 7B (open)":    "open-mistral-7b",
}
EMBED_MODEL    = "sentence-transformers/all-MiniLM-L6-v2"
ANSWER_MODES   = {
    "Detailed":      "Give a thorough, detailed answer covering all relevant aspects.",
    "Concise":       "Give a concise, direct answer in 2-4 sentences.",
    "Bullet Points": "Structure your entire answer as clear bullet points.",
    "Action Items":  "Extract every actionable HSE requirement as a numbered checklist.",
}
MEMORY_TURNS   = 4
SUPPORTED_EXTS = {".pdf", ".docx", ".doc", ".txt"}

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600&family=JetBrains+Mono:wght@400;500&display=swap');
:root{
  --bg:#070a0f; --s1:#0d1219; --s2:#131b24; --s3:#18222e;
  --bd:#1c2838; --bd2:#253345;
  --g1:#00dfa8; --g2:#1a7aff;
  --tx:#ccd8e8; --mu:#4d6070; --mu2:#2d3f50;
  --ok:#22c55e; --wn:#f59e0b; --er:#ef4444;
}
*{box-sizing:border-box;}
html,body,[data-testid="stAppViewContainer"]{background:var(--bg)!important;color:var(--tx)!important;font-family:'DM Sans',sans-serif!important;}
[data-testid="stSidebar"]{background:var(--s1)!important;border-right:1px solid var(--bd)!important;}

/* header */
.hdr{display:flex;align-items:center;gap:14px;padding:18px 0 14px;border-bottom:1px solid var(--bd);margin-bottom:22px;}
.logo{width:44px;height:44px;border-radius:12px;background:linear-gradient(135deg,var(--g1),var(--g2));display:flex;align-items:center;justify-content:center;font-size:20px;flex-shrink:0;box-shadow:0 0 22px rgba(0,223,168,.2);}
.hdr-title{font-family:'Syne',sans-serif;font-size:22px;font-weight:800;color:var(--tx);margin:0;letter-spacing:-.02em;}
.hdr-sub{font-size:10px;color:var(--mu);letter-spacing:.12em;text-transform:uppercase;margin-top:2px;}

/* status badges */
.bdg{display:inline-flex;align-items:center;gap:5px;padding:3px 11px;border-radius:20px;font-size:11px;font-weight:600;}
.b-ok  {background:rgba(34,197,94,.08);color:var(--ok);border:1px solid rgba(34,197,94,.2);}
.b-wn  {background:rgba(245,158,11,.08);color:var(--wn);border:1px solid rgba(245,158,11,.2);}
.b-inf {background:rgba(26,122,255,.09);color:#5aaaff;border:1px solid rgba(26,122,255,.22);}

/* chat */
.chat{display:flex;flex-direction:column;gap:18px;padding:2px 0;}
.row-u,.row-b{display:flex;gap:11px;align-items:flex-start;animation:up .25s ease;}
.row-u{flex-direction:row-reverse;}
@keyframes up{from{opacity:0;transform:translateY(6px)}to{opacity:1}}
.av{width:34px;height:34px;border-radius:9px;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:15px;}
.av-u{background:var(--s2);border:1px solid var(--bd2);}
.av-b{background:linear-gradient(135deg,var(--g1),var(--g2));box-shadow:0 2px 10px rgba(0,223,168,.16);}
.bbl{max-width:78%;padding:13px 17px;border-radius:14px;font-size:14px;line-height:1.78;}
.bbl-u{background:var(--s2);border:1px solid var(--bd2);border-top-right-radius:4px;}
.bbl-b{background:var(--s1);border:1px solid var(--bd);border-top-left-radius:4px;}
.meta-row{display:flex;align-items:center;gap:7px;margin-top:6px;flex-wrap:wrap;}
.mt{font-size:11px;color:var(--mu);font-family:'JetBrains Mono',monospace;}
.cf{font-size:11px;padding:2px 8px;border-radius:10px;}
.cf-hi{background:rgba(34,197,94,.08);color:var(--ok);border:1px solid rgba(34,197,94,.2);}
.cf-me{background:rgba(245,158,11,.08);color:var(--wn);border:1px solid rgba(245,158,11,.2);}
.cf-lo{background:rgba(239,68,68,.08);color:var(--er);border:1px solid rgba(239,68,68,.2);}
.src-block{margin-top:9px;padding:9px 13px;background:var(--s2);border-left:3px solid var(--g1);border-radius:0 8px 8px 0;font-size:12px;}
.src-block strong{color:var(--g1);}
.src-tag{display:inline-flex;align-items:center;gap:3px;background:rgba(0,223,168,.06);border:1px solid rgba(0,223,168,.18);border-radius:5px;padding:2px 8px;margin:2px;font-family:'JetBrains Mono',monospace;font-size:11px;color:#00c9a0;}
/* keyword highlight */
.kw{background:rgba(0,223,168,.15);border-bottom:1.5px solid var(--g1);padding:0 2px;border-radius:3px;color:#00e8b5;font-weight:600;}
/* chunk viewer */
.c-lbl{font-size:11px;font-weight:700;color:var(--g1);margin-bottom:4px;font-family:'Syne',sans-serif;}
.c-box{background:var(--bg);border:1px solid var(--bd);border-radius:7px;padding:9px 12px;font-size:11.5px;font-family:'JetBrains Mono',monospace;color:var(--mu);white-space:pre-wrap;max-height:170px;overflow-y:auto;margin-bottom:7px;}

/* sidebar labels */
.sl{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--mu);margin-bottom:7px;display:block;}
.stat-r{display:flex;justify-content:space-between;margin-bottom:5px;}
.sk{font-size:12px;color:var(--mu);}
.sv{font-size:12px;font-weight:600;color:var(--tx);font-family:'JetBrains Mono',monospace;}
.scard{background:var(--s2);border:1px solid var(--bd);border-radius:11px;padding:13px;margin-bottom:10px;}

/* OneDrive link panel */
.od-panel{background:linear-gradient(135deg,rgba(26,122,255,.06),rgba(0,223,168,.04));border:1px solid rgba(26,122,255,.22);border-radius:12px;padding:15px;margin-bottom:10px;}
.od-title{font-family:'Syne',sans-serif;font-size:13px;font-weight:800;color:#5aaaff;margin-bottom:8px;display:flex;align-items:center;gap:6px;}
.od-hint{font-size:12px;color:var(--mu);line-height:1.65;}
.od-file{display:flex;align-items:center;padding:6px 9px;background:var(--s3);border:1px solid var(--bd);border-radius:7px;margin-bottom:4px;gap:8px;}
.od-fn{color:var(--tx);font-family:'JetBrains Mono',monospace;font-size:11px;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.od-sz{color:var(--mu);font-size:10px;white-space:nowrap;}
.pbdg{font-size:10px;padding:1px 6px;border-radius:5px;background:rgba(0,223,168,.07);border:1px solid rgba(0,223,168,.16);color:#00c49a;margin:1px;font-family:'JetBrains Mono',monospace;}

/* inputs */
.stTextInput>div>div>input{background:var(--s1)!important;border:1px solid var(--bd2)!important;border-radius:12px!important;color:var(--tx)!important;font-family:'DM Sans',sans-serif!important;font-size:14px!important;padding:13px 17px!important;transition:border-color .2s!important;}
.stTextInput>div>div>input:focus{border-color:var(--g1)!important;box-shadow:0 0 0 3px rgba(0,223,168,.07)!important;}
.stTextInput>div>div>input::placeholder{color:var(--mu)!important;}
.stButton>button{background:linear-gradient(135deg,var(--g1),var(--g2))!important;color:#030810!important;border:none!important;border-radius:10px!important;font-weight:700!important;font-family:'Syne',sans-serif!important;box-shadow:0 3px 14px rgba(0,223,168,.15)!important;transition:opacity .18s,transform .14s!important;}
.stButton>button:hover{opacity:.83!important;transform:translateY(-1px)!important;}
[data-testid="stSelectbox"]>div>div{background:var(--s2)!important;border-color:var(--bd)!important;color:var(--tx)!important;}
[data-testid="stExpander"]{background:var(--s2)!important;border:1px solid var(--bd)!important;border-radius:10px!important;}
[data-testid="stFileUploader"]{background:var(--s2)!important;border:1px dashed rgba(0,223,168,.2)!important;border-radius:12px!important;}

/* welcome card */
.welcome{border:1px solid var(--bd);border-radius:18px;padding:32px;text-align:center;background:var(--s1);margin:30px auto;max-width:600px;position:relative;overflow:hidden;}
.welcome::before{content:'';position:absolute;top:-90px;right:-90px;width:260px;height:260px;background:radial-gradient(circle,rgba(0,223,168,.06) 0%,transparent 68%);pointer-events:none;}
.wi{font-size:48px;margin-bottom:14px;}
.wt{font-family:'Syne',sans-serif;font-size:22px;font-weight:800;color:var(--tx);margin-bottom:9px;}
.wx{font-size:13.5px;color:var(--mu);line-height:1.68;}
.chip{display:inline-block;background:var(--s2);border:1px solid var(--bd2);border-radius:20px;padding:5px 13px;font-size:12px;color:var(--mu);margin:3px;}
.hint{background:rgba(0,223,168,.04);border:1px solid rgba(0,223,168,.15);border-radius:10px;padding:11px 14px;font-size:12px;color:#00b890;margin-bottom:9px;line-height:1.55;}
.apihint{background:rgba(245,158,11,.04);border:1px solid rgba(245,158,11,.14);border-radius:9px;padding:11px 14px;font-size:12px;color:#b89000;line-height:1.6;margin-top:5px;}
::-webkit-scrollbar{width:4px;}::-webkit-scrollbar-thumb{background:var(--bd2);border-radius:2px;}
hr{border-color:var(--bd)!important;}
#MainMenu,footer,header{visibility:hidden;}[data-testid="stToolbar"]{display:none;}
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  MICROSOFT GRAPH — AUTHENTICATED SHAREPOINT / ONEDRIVE FETCHER
# ═══════════════════════════════════════════════════════════════════════════════
#
#  Uses the Microsoft device-code OAuth flow (no Azure app registration needed
#  for most Microsoft 365 / work accounts — uses the well-known "Office" client).
#  Supports:
#    • Personal OneDrive  (graph.microsoft.com/v1.0/me/drive)
#    • SharePoint sites   (graph.microsoft.com/v1.0/sites/{site-id}/drives)
#    • Shared links       (graph.microsoft.com/v1.0/shares/{token}/driveItem)
#
#  Flow in the sidebar:
#    1. User clicks "Sign in with Microsoft"
#    2. App shows a device code + microsoft.com/devicelogin URL
#    3. User opens URL, enters code, signs in with their work account
#    4. App polls for the token, then lists / downloads files
# ─────────────────────────────────────────────────────────────────────────────

# Use the well-known Microsoft "Office" public client — works for any M365 tenant
# without requiring the admin to register an Azure app.
MS_CLIENT_ID   = "d3590ed6-52b3-4102-aeff-aad2292ab01c"   # Microsoft Office client
MS_SCOPES      = "Files.Read.All Sites.Read.All offline_access"
MS_DEVICE_URL  = "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode"
MS_TOKEN_URL   = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
GRAPH          = "https://graph.microsoft.com/v1.0"


def _is_likely_html(data: bytes) -> bool:
    try:
        sniff = data[:512].decode("utf-8", errors="ignore").lower()
        return "<html" in sniff or "<!doctype" in sniff or "<head" in sniff
    except Exception:
        return False


# ── Auth helpers ──────────────────────────────────────────────────────────────

def ms_start_device_flow() -> dict:
    import requests
    r = requests.post(MS_DEVICE_URL, data={"client_id": MS_CLIENT_ID, "scope": MS_SCOPES}, timeout=15)
    return r.json()


def ms_poll_token(device_code: str) -> dict | None:
    """Poll once. Returns token dict if ready, None if still pending, raises on hard error."""
    import requests
    r = requests.post(MS_TOKEN_URL, data={
        "client_id":   MS_CLIENT_ID,
        "grant_type":  "urn:ietf:params:oauth:grant-type:device_code",
        "device_code": device_code,
    }, timeout=15)
    d = r.json()
    if "access_token" in d:
        return d
    err = d.get("error", "")
    if err in ("authorization_pending", "slow_down"):
        return None   # still waiting — caller should retry
    raise RuntimeError(d.get("error_description", f"Auth error: {err}"))


def ms_refresh_token(refresh_tok: str) -> dict | None:
    import requests
    r = requests.post(MS_TOKEN_URL, data={
        "client_id":     MS_CLIENT_ID,
        "grant_type":    "refresh_token",
        "refresh_token": refresh_tok,
        "scope":         MS_SCOPES,
    }, timeout=15)
    d = r.json()
    return d if "access_token" in d else None


def _graph_get(path: str, token: str, params: dict = None) -> dict:
    import requests
    url = path if path.startswith("http") else f"{GRAPH}{path}"
    r   = requests.get(url, headers={"Authorization": f"Bearer {token}"},
                       params=params, timeout=30)
    return r.json()


def _graph_download(url: str, token: str) -> bytes:
    import requests
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"},
                     allow_redirects=True, timeout=90)
    r.raise_for_status()
    return r.content


# ── SharePoint URL parser ─────────────────────────────────────────────────────

def _parse_sharepoint_url(url: str) -> tuple[str | None, str | None]:
    """
    Extract (hostname, site_path) from a SharePoint URL.
    e.g. https://vestas.sharepoint.com/sites/GlobalQHSE-hub/HSEN%20Legacy/...
         → ("vestas.sharepoint.com", "/sites/GlobalQHSE-hub")
    """
    m = re.match(r"https?://([^/]+)((?:/sites/[^/?#]+)?)", url, re.IGNORECASE)
    if m:
        return m.group(1), m.group(2) or "/"
    return None, None


def _resolve_sharepoint_folder_path(url: str) -> str:
    """
    Extract the server-relative folder path from a SharePoint URL.
    e.g. .../sites/GlobalQHSE-hub/HSEN%20Legacy/Forms/... → /sites/GlobalQHSE-hub/HSEN Legacy
    Strips /Forms/... suffixes added by SharePoint's allitems.aspx view.
    """
    from urllib.parse import urlparse, unquote
    parsed = urlparse(url)
    path   = unquote(parsed.path)
    # Remove trailing /Forms/HSEN.aspx style view suffixes
    path   = re.sub(r"/Forms/[^/]*$", "", path)
    path   = re.sub(r"/[^/]+\.aspx$", "", path)
    return path.rstrip("/")


# ── Main fetcher ──────────────────────────────────────────────────────────────

def fetch_from_sharepoint(access_token: str, sp_url: str) -> tuple[list, list]:
    """
    Fetch all supported files from a SharePoint library folder using Graph API.
    sp_url: the full SharePoint URL the user pasted (site + folder path).
    Returns (file_pairs, errors).
    """
    results, errors = [], []

    hostname, site_path = _parse_sharepoint_url(sp_url)
    if not hostname:
        return [], ["Could not parse SharePoint URL — please paste the full URL from your browser."]

    folder_path = _resolve_sharepoint_folder_path(sp_url)

    # 1. Resolve site ID
    site_resp = _graph_get(f"/sites/{hostname}:{site_path}", access_token)
    if "error" in site_resp:
        return [], [f"Could not find SharePoint site '{site_path}' on {hostname}: "
                    f"{site_resp['error'].get('message', site_resp['error'])}"]
    site_id = site_resp["id"]

    # 2. List drives (document libraries) for the site
    drives_resp = _graph_get(f"/sites/{site_id}/drives", access_token)
    drives      = drives_resp.get("value", [])
    if not drives:
        return [], [f"No document libraries found on site {hostname}{site_path}"]

    # 3. Find the drive whose root path is a prefix of folder_path
    #    e.g. folder_path = /sites/GlobalQHSE-hub/HSEN Legacy
    target_drive = None
    subfolder    = ""
    for drive in drives:
        # Drive webUrl looks like: https://tenant.sharepoint.com/sites/site/LibraryName
        drive_web  = drive.get("webUrl", "")
        drive_path = _resolve_sharepoint_folder_path(drive_web)
        if folder_path.startswith(drive_path):
            target_drive = drive
            subfolder    = folder_path[len(drive_path):].lstrip("/")
            break

    if not target_drive:
        # Fallback: use the first drive and treat the whole path as subfolder
        target_drive = drives[0]
        drive_web    = target_drive.get("webUrl", "")
        drive_path   = _resolve_sharepoint_folder_path(drive_web)
        subfolder    = folder_path[len(drive_path):].lstrip("/") if folder_path.startswith(drive_path) else ""

    drive_id = target_drive["id"]

    # 4. List files in the target folder (or root if subfolder is empty)
    def list_folder(folder_rel_path: str):
        if folder_rel_path:
            url = f"{GRAPH}/drives/{drive_id}/root:/{folder_rel_path}:/children"
        else:
            url = f"{GRAPH}/drives/{drive_id}/root/children"

        while url:
            data = _graph_get(url, access_token)
            if "error" in data:
                errors.append(f"Could not list folder '{folder_rel_path}': "
                               f"{data['error'].get('message', data['error'])}")
                return
            for item in data.get("value", []):
                if "file" in item:
                    ext = Path(item["name"]).suffix.lower()
                    if ext not in SUPPORTED_EXTS:
                        continue
                    dl_url = item.get("@microsoft.graph.downloadUrl") or \
                             f"{GRAPH}/drives/{drive_id}/items/{item['id']}/content"
                    try:
                        content = _graph_download(dl_url, access_token)
                        if _is_likely_html(content):
                            errors.append(f"{item['name']}: download returned HTML — token may have expired")
                        else:
                            results.append((item["name"], content))
                    except Exception as e:
                        errors.append(f"Download failed for {item['name']}: {e}")
                elif "folder" in item:
                    sub = f"{folder_rel_path}/{item['name']}" if folder_rel_path else item["name"]
                    list_folder(sub)
            url = data.get("@odata.nextLink")

    list_folder(subfolder)
    return results, errors


def fetch_from_onedrive(access_token: str, folder_path: str = "") -> tuple[list, list]:
    """Fetch files from the signed-in user's personal OneDrive."""
    results, errors = [], []

    def list_folder(rel: str):
        url = (f"{GRAPH}/me/drive/root:/{rel}:/children" if rel
               else f"{GRAPH}/me/drive/root/children")
        while url:
            data = _graph_get(url, access_token)
            if "error" in data:
                errors.append(f"OneDrive error: {data['error'].get('message', data['error'])}")
                return
            for item in data.get("value", []):
                if "file" in item:
                    ext = Path(item["name"]).suffix.lower()
                    if ext not in SUPPORTED_EXTS:
                        continue
                    dl  = item.get("@microsoft.graph.downloadUrl", "")
                    try:
                        content = _graph_download(dl or
                            f"{GRAPH}/me/drive/items/{item['id']}/content", access_token)
                        if not _is_likely_html(content):
                            results.append((item["name"], content))
                    except Exception as e:
                        errors.append(f"Download failed {item['name']}: {e}")
                elif "folder" in item:
                    list_folder(f"{rel}/{item['name']}" if rel else item["name"])
            url = data.get("@odata.nextLink")

    list_folder(folder_path)
    return results, errors


# ═══════════════════════════════════════════════════════════════════════════════
#  DOCUMENT PARSERS
# ═══════════════════════════════════════════════════════════════════════════════

def _table_to_md(table: list) -> str:
    if not table:
        return ""
    rows = []
    for i, row in enumerate(table):
        cells = [str(c).strip() if c else "" for c in row]
        rows.append("| " + " | ".join(cells) + " |")
        if i == 0:
            rows.append("|" + "|".join(["---"] * len(cells)) + "|")
    return "\n".join(rows)


def _docx_table_to_md(table) -> str:
    rows = []
    for i, row in enumerate(table.rows):
        cells = [cell.text.strip() for cell in row.cells]
        rows.append("| " + " | ".join(cells) + " |")
        if i == 0:
            rows.append("|" + "|".join(["---"] * len(cells)) + "|")
    return "\n".join(rows)


def _ocr_page(pil_img) -> str:
    try:
        import pytesseract
        return pytesseract.image_to_string(pil_img, config="--psm 6")
    except Exception as e:
        logger.warning(f"OCR error: {e}")
        return ""


def parse_pdf(filename: str, data: bytes) -> tuple[list, dict]:
    import pdfplumber
    from langchain_core.documents import Document

    # Guard: reject HTML bytes masquerading as a PDF
    if _is_likely_html(data):
        raise ValueError(
            f"{filename}: received an HTML page instead of a PDF. "
            "The OneDrive link may require sign-in or has expired."
        )
    if len(data) < 64 or not data[:4] == b"%PDF":
        raise ValueError(
            f"{filename}: file does not start with %PDF — "
            f"first bytes: {data[:16]!r}. It may be password-protected or corrupted."
        )

    docs       = []
    stats      = {"pages": 0, "text_pages": 0, "tables": 0, "ocr_pages": 0}
    ocr_needed = []

    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            stats["pages"] = len(pdf.pages)
            for pn, page in enumerate(pdf.pages):
                parts = []
                raw   = page.extract_text() or ""
                try:
                    for t in page.extract_tables():
                        md = _table_to_md(t)
                        if md:
                            parts.append(f"\n[TABLE]\n{md}\n[/TABLE]\n")
                            stats["tables"] += 1
                except Exception:
                    pass   # table extraction is best-effort
                if raw.strip():
                    parts.insert(0, raw.strip())
                    stats["text_pages"] += 1
                else:
                    ocr_needed.append(pn)
                if parts:
                    docs.append(Document(
                        page_content="\n\n".join(parts),
                        metadata={"source_file": filename, "page": pn},
                    ))
    except ValueError:
        raise
    except Exception as e:
        raise ValueError(f"{filename}: pdfplumber failed — {e}") from e

    if ocr_needed:
        try:
            from pdf2image import convert_from_bytes
            imgs = convert_from_bytes(data, dpi=200)
            for pn in ocr_needed:
                if pn < len(imgs):
                    txt = _ocr_page(imgs[pn])
                    if txt.strip():
                        docs.append(Document(
                            page_content=f"[OCR]\n{txt.strip()}\n[/OCR]",
                            metadata={"source_file": filename, "page": pn, "ocr": True},
                        ))
                        stats["ocr_pages"] += 1
        except Exception as e:
            logger.warning(f"OCR pass failed for {filename}: {e}")
            stats["ocr_note"] = str(e)

    docs.sort(key=lambda d: d.metadata.get("page", 0))
    return docs, stats


def parse_docx(filename: str, data: bytes) -> tuple[list, dict]:
    from langchain_core.documents import Document

    if _is_likely_html(data):
        raise ValueError(
            f"{filename}: received an HTML page instead of a DOCX file. "
            "The OneDrive link may require sign-in or has expired."
        )
    # DOCX files are ZIP archives — check magic bytes
    if len(data) < 4 or data[:2] != b"PK":
        raise ValueError(
            f"{filename}: not a valid DOCX/ZIP file — "
            f"first bytes: {data[:8]!r}. May be corrupted or password-protected."
        )

    try:
        import docx as _docx
    except ImportError:
        try:
            from docx import Document as _DocxDocument   # noqa: F401
            import docx as _docx
        except ImportError:
            raise ImportError(
                "python-docx is not installed. Run: pip install python-docx"
            )

    try:
        doc   = _docx.Document(io.BytesIO(data))
        stats = {"paragraphs": 0, "tables": 0}
        parts = []

        for block in doc.element.body:
            tag = block.tag.split("}")[-1] if "}" in block.tag else block.tag
            if tag == "p":
                try:
                    para = _docx.text.paragraph.Paragraph(block, doc)
                    txt  = para.text.strip()
                    if txt:
                        parts.append(txt)
                        stats["paragraphs"] += 1
                except Exception:
                    pass
            elif tag == "tbl":
                try:
                    md = _docx_table_to_md(_docx.table.Table(block, doc))
                    if md:
                        parts.append(f"\n[TABLE]\n{md}\n[/TABLE]\n")
                        stats["tables"] += 1
                except Exception:
                    pass

        content = "\n\n".join(parts)
        docs    = ([Document(page_content=content, metadata={"source_file": filename})]
                   if content.strip() else [])
        return docs, stats
    except (ValueError, ImportError):
        raise
    except Exception as e:
        raise ValueError(f"{filename}: python-docx failed — {e}") from e


def parse_txt(filename: str, data: bytes) -> tuple[list, dict]:
    from langchain_core.documents import Document
    return [Document(page_content=data.decode("utf-8", errors="ignore"),
                     metadata={"source_file": filename})], {}


def parse_doc_legacy(filename: str, data: bytes) -> tuple[list, dict]:
    import docx2txt
    from langchain_core.documents import Document
    txt = docx2txt.process(io.BytesIO(data))
    return [Document(page_content=txt, metadata={"source_file": filename})], {"note": "legacy .doc"}


def dispatch_parse(filename: str, data: bytes) -> tuple[list, dict]:
    ext = Path(filename).suffix.lower()
    if ext == ".pdf":  return parse_pdf(filename, data)
    if ext == ".docx": return parse_docx(filename, data)
    if ext == ".doc":  return parse_doc_legacy(filename, data)
    if ext == ".txt":  return parse_txt(filename, data)
    return [], {"error": f"Unsupported: {ext}"}


# ═══════════════════════════════════════════════════════════════════════════════
#  VECTORSTORE
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_resource(show_spinner=False)
def build_index(_file_tuples: tuple, cache_ver: int):
    from langchain_text_splitters import RecursiveCharacterTextSplitter
    from langchain_community.vectorstores import Chroma
    from langchain_community.embeddings import HuggingFaceEmbeddings

    all_docs, all_stats = [], {}
    for fname, fbytes in _file_tuples:
        try:
            docs, stats = dispatch_parse(fname, fbytes)
            if docs:
                all_docs.extend(docs)
                all_stats[fname] = stats
                logger.info(f"Parsed {fname}: {len(docs)} doc(s), stats={stats}")
            else:
                all_stats[fname] = {"error": "Parser returned no content — file may be empty or image-only without OCR"}
                logger.warning(f"No content from {fname}")
        except Exception as e:
            err_msg = str(e)
            all_stats[fname] = {"error": err_msg}
            logger.warning(f"Parse failed {fname}: {err_msg}")

    if not all_docs:
        return None, 0, all_stats   # return stats so errors show in UI

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000, chunk_overlap=200,
        separators=["\n\n", "\n", ". ", "! ", "? "],
    )
    chunks = splitter.split_documents(all_docs)
    emb    = HuggingFaceEmbeddings(
        model_name=EMBED_MODEL,
        model_kwargs={"device": "cpu"},
        encode_kwargs={"normalize_embeddings": True},
    )
    vs = Chroma.from_documents(documents=chunks, embedding=emb)
    return vs, len(chunks), all_stats


# ═══════════════════════════════════════════════════════════════════════════════
#  RAG HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _mistral_key() -> str:
    try:
        return st.secrets["MISTRAL_API_KEY"]
    except Exception:
        return os.getenv("MISTRAL_API_KEY", "")


def format_sources(docs: list) -> str:
    seen, tags = set(), []
    for d in docs:
        name  = d.metadata.get("source_file", "Unknown")
        page  = d.metadata.get("page", "")
        ocr   = " OCR" if d.metadata.get("ocr") else ""
        label = name + (f" · p{page+1}" if page != "" else "") + ocr
        if label not in seen:
            seen.add(label)
            tags.append(f'<span class="src-tag">📄 {label}</span>')
    return "".join(tags)


def labeled_context(docs: list) -> str:
    parts = []
    for d in docs:
        src   = d.metadata.get("source_file", "Unknown")
        page  = d.metadata.get("page", "")
        ocr   = " [OCR]" if d.metadata.get("ocr") else ""
        label = src + (f" page {page+1}" if page != "" else "") + ocr
        parts.append(f"[Source: {label}]\n{d.page_content.strip()}")
    return "\n\n---\n\n".join(parts)


def get_confidence(answer: str, docs: list) -> str:
    low_phrases = [
        "not found", "not mention", "cannot find", "no information",
        "not available", "not in the document", "i don't", "i do not",
    ]
    if any(p in answer.lower() for p in low_phrases):
        return "low"
    if len({d.metadata.get("source_file", "") for d in docs}) >= 2 and len(answer) > 300:
        return "high"
    return "medium"


def conf_badge(level: str) -> str:
    L = {"high": "● High confidence", "medium": "◐ Medium confidence", "low": "○ Low confidence"}
    C = {"high": "cf-hi", "medium": "cf-me", "low": "cf-lo"}
    return f'<span class="cf {C.get(level, "cf-me")}">{L.get(level, "")}</span>'


def highlight_keywords(text: str, query: str) -> str:
    """Wrap query terms (>3 chars) in a highlight span."""
    words = sorted({w for w in re.split(r'\W+', query) if len(w) > 3},
                   key=len, reverse=True)
    for w in words:
        text = re.compile(re.escape(w), re.IGNORECASE).sub(
            lambda m: f'<span class="kw">{m.group()}</span>', text
        )
    return text


def build_memory(messages: list) -> str:
    pairs = [
        (messages[i], messages[i + 1])
        for i in range(len(messages) - 1)
        if messages[i]["role"] == "user" and messages[i + 1]["role"] == "assistant"
    ]
    if not pairs:
        return ""
    lines = ["Conversation so far:"]
    for u, a in pairs[-MEMORY_TURNS:]:
        lines += [f"Q: {u['content']}", f"A: {a['content'][:400]}..."]
    return "\n".join(lines)


def retrieve_docs(vs, question: str, files: list, k: int) -> list:
    """MMR retrieval, then fill gaps per source for cross-doc coverage."""
    retriever = vs.as_retriever(
        search_type="mmr",
        search_kwargs={"k": k, "fetch_k": k * 4, "lambda_mult": 0.65},
    )
    docs  = retriever.invoke(question)
    found = {d.metadata.get("source_file", "") for d in docs}
    for f in files:
        if f not in found:
            try:
                docs.extend(vs.similarity_search(question, k=2, filter={"source_file": f}))
                found.add(f)
            except Exception:
                pass
    return docs


def run_rag(vs, question: str, model: str, api_key: str, temp: float,
            k: int, mode: str, files: list, messages: list, scope: str) -> dict:
    from langchain_core.prompts import PromptTemplate
    from langchain_core.output_parsers import StrOutputParser
    from langchain_mistralai import ChatMistralAI

    llm = ChatMistralAI(model=model, api_key=api_key, temperature=temp, max_tokens=1536)

    docs = (vs.similarity_search(question, k=k, filter={"source_file": scope})
            if scope and scope != "All Documents"
            else retrieve_docs(vs, question, files, k))

    template = """\
You are an expert HSE (Health, Safety & Environment) assistant for an organisation.
Use ONLY the context below — sourced from official HSE notification documents.
Each chunk is labelled [Source: filename]. Cite the source filename for every fact you state.
Context may include tables in markdown format — read them carefully as structured data.
Context may include OCR-extracted text from scanned pages — treat it as authoritative.
If the question spans multiple documents, synthesise information from ALL of them.
If the answer cannot be found in the documents, say so clearly — never invent information.

{memory}

Context:
{context}

Question: {question}

Instruction: {mode}

Answer:\
"""
    chain  = PromptTemplate(
        input_variables=["context", "question", "memory", "mode"],
        template=template,
    ) | llm | StrOutputParser()

    answer = chain.invoke({
        "context":  labeled_context(docs),
        "question": question,
        "memory":   build_memory(messages),
        "mode":     ANSWER_MODES.get(mode, ANSWER_MODES["Detailed"]),
    })
    return {"result": answer, "source_documents": docs}


def generate_followups(answer: str, question: str, api_key: str, model: str) -> list:
    try:
        from langchain_mistralai import ChatMistralAI
        llm  = ChatMistralAI(model=model, api_key=api_key, temperature=0.4, max_tokens=120)
        resp = llm.invoke(
            f"Given this HSE Q&A:\nQ: {question}\nA: {answer[:400]}\n\n"
            "Suggest exactly 3 short follow-up questions a safety officer might ask next. "
            "Reply ONLY with the 3 questions, one per line, no numbering or bullets."
        )
        return [l.strip().lstrip("-• ") for l in resp.content.strip().split("\n") if l.strip()][:3]
    except Exception:
        return []


def render_parse_stats(stats: dict) -> str:
    lines = []
    for fname, s in stats.items():
        badges = []
        if "error" in s:
            badges.append('<span class="pbdg" style="color:var(--er)">⚠ error</span>')
        else:
            for key, icon in [("pages","📄"),("tables","📊"),("ocr_pages","🔍"),("paragraphs","¶")]:
                if s.get(key):
                    badges.append(f'<span class="pbdg">{icon} {s[key]} {key}</span>')
            if s.get("note"):
                badges.append(f'<span class="pbdg">ℹ {s["note"]}</span>')
        short = fname if len(fname) < 32 else fname[:29] + "…"
        lines.append(
            f'<div style="margin-bottom:5px">'
            f'<span style="font-size:11px;color:var(--mu)">{short}</span> {"".join(badges)}'
            f'</div>'
        )
    return "".join(lines)


def export_chat(messages: list) -> str:
    lines = ["HSE Notifications Bot — Chat Export", "=" * 50, ""]
    for m in messages:
        role = "You" if m["role"] == "user" else "HSE Assistant"
        lines += [f"[{role}]", m["content"]]
        if m.get("time"):
            lines.append(f"  Response time: {m['time']}")
        lines.append("")
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════════
defaults = {
    "messages":      [],
    "vectorstore":   None,
    "chunk_count":   0,
    "indexed_files": [],
    "index_ver":     0,
    "parse_stats":   {},
    "pending":       "",
    "fetched_files": [],
    # Microsoft auth
    "ms_token":       None,   # full token dict (access_token, refresh_token, …)
    "ms_device_flow": None,   # device-code flow dict while polling
    "ms_user":        "",     # display name once signed in
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ═══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style="padding:16px 0 10px">
      <div style="font-family:'Syne',sans-serif;font-size:18px;font-weight:800;color:#ccd8e8">🦺 HSE Bot</div>
      <div style="font-size:10px;color:#4d6070;letter-spacing:.1em;text-transform:uppercase;margin-top:3px">Configuration</div>
    </div>""", unsafe_allow_html=True)
    st.divider()

    # ── Mistral API Key ────────────────────────────────────────────────────────
    st.markdown('<span class="sl">Mistral AI API Key</span>', unsafe_allow_html=True)
    stored_key  = _mistral_key()
    if stored_key:
        st.markdown('<div style="font-size:12px;color:#22c55e;margin-bottom:8px">✅ Key loaded from secrets</div>',
                    unsafe_allow_html=True)
        mistral_key = stored_key
    else:
        mistral_key = st.text_input("mk", type="password",
                                    placeholder="Enter Mistral AI API key…",
                                    label_visibility="collapsed")
        st.markdown(
            '<div class="apihint">🔑 Free key at <a href="https://console.mistral.ai" '
            'target="_blank" style="color:#f59e0b;font-weight:700">console.mistral.ai</a></div>',
            unsafe_allow_html=True)
    st.divider()

    # ── Model & tuning ─────────────────────────────────────────────────────────
    st.markdown('<span class="sl">Model & Tuning</span>', unsafe_allow_html=True)
    model_lbl   = st.selectbox("Model", list(MISTRAL_MODELS.keys()), index=0,
                               label_visibility="collapsed")
    model_name  = MISTRAL_MODELS[model_lbl]
    temperature = st.slider("Temperature", 0.0, 1.0, 0.1, 0.05,
                            help="Lower = more precise and factual")
    top_k       = st.slider("Top-K Sources", 2, 12, 6,
                            help="Number of document chunks retrieved per question")
    st.divider()

    # ── Answer style ───────────────────────────────────────────────────────────
    st.markdown('<span class="sl">Answer Style</span>', unsafe_allow_html=True)
    answer_mode = st.radio("Style", list(ANSWER_MODES.keys()), index=0,
                           label_visibility="collapsed", horizontal=True)
    st.divider()

    # ── Document Source ────────────────────────────────────────────────────────
    st.markdown('<span class="sl">Document Source</span>', unsafe_allow_html=True)
    source_tab = st.radio("src", ["☁️ SharePoint / OneDrive", "📂 Upload Files"],
                          label_visibility="collapsed")

    # ════ MODE A: SharePoint / OneDrive with Microsoft sign-in ════════════════
    if source_tab == "☁️ SharePoint / OneDrive":

        # ── Already signed in ──────────────────────────────────────────────────
        if st.session_state.ms_token:
            tok = st.session_state.ms_token
            name_display = st.session_state.ms_user or "Microsoft Account"
            st.markdown(
                f'<div style="font-size:12px;color:#22c55e;margin-bottom:8px">'
                f'✅ Signed in as <strong>{name_display}</strong></div>',
                unsafe_allow_html=True)

            col_sig, col_clr = st.columns(2)
            with col_sig:
                if st.button("🚪 Sign Out", use_container_width=True):
                    st.session_state.ms_token       = None
                    st.session_state.ms_device_flow = None
                    st.session_state.ms_user        = ""
                    st.rerun()

            st.markdown("""
            <div class="od-panel">
              <div class="od-title">📂 SharePoint Folder URL</div>
              <div class="od-hint">
                Navigate to your HSE folder in SharePoint, then copy the URL
                from your browser address bar and paste it below.
              </div>
            </div>""", unsafe_allow_html=True)

            sp_url = st.text_input(
                "spurl",
                placeholder="https://company.sharepoint.com/sites/GlobalQHSE-hub/HSEN Legacy",
                label_visibility="collapsed",
                help="Paste the SharePoint or OneDrive folder URL from your browser",
            )

            is_onedrive = sp_url and "my.sharepoint.com" in sp_url
            is_sp       = sp_url and "sharepoint.com/sites" in sp_url

            if sp_url and st.button("⬇️  Fetch & Index Files", use_container_width=True):
                # Auto-refresh token if needed
                if tok.get("refresh_token"):
                    refreshed = ms_refresh_token(tok["refresh_token"])
                    if refreshed:
                        st.session_state.ms_token = refreshed
                        tok = refreshed

                access_token = tok["access_token"]

                with st.spinner("Browsing SharePoint library…"):
                    if is_onedrive:
                        file_pairs, errs = fetch_from_onedrive(access_token, "")
                    else:
                        file_pairs, errs = fetch_from_sharepoint(access_token, sp_url.strip())

                if errs:
                    for e in errs:
                        st.warning(f"⚠ {e}")

                if file_pairs:
                    st.session_state.fetched_files = [(n, len(b)) for n, b in file_pairs]
                    with st.spinner(f"Parsing & indexing {len(file_pairs)} file(s)…"):
                        st.session_state.index_ver += 1
                        vs, n, ps = build_index(tuple(file_pairs), st.session_state.index_ver)
                    st.session_state.parse_stats = ps
                    if vs:
                        st.session_state.vectorstore   = vs
                        st.session_state.chunk_count   = n
                        st.session_state.indexed_files = [p[0] for p in file_pairs]
                        st.success(f"✅ {n:,} chunks indexed from {len(file_pairs)} file(s)!")
                    else:
                        failed = {f: s for f, s in ps.items() if "error" in s}
                        if failed:
                            st.error("Parsing failed — see Parse Report below.")
                            for fname, s in failed.items():
                                st.error(f"📄 **{fname}**: {s['error']}")
                        else:
                            st.error("No text content extracted. Files may be image-only PDFs.")
                elif not errs:
                    st.warning("No PDF or DOCX files found in that folder.")

        # ── Device-code flow in progress ───────────────────────────────────────
        elif st.session_state.ms_device_flow:
            flow = st.session_state.ms_device_flow
            st.markdown(f"""
            <div class="od-panel">
              <div class="od-title">🔐 Sign in with Microsoft</div>
              <div class="od-hint">
                <strong>Step 1</strong> — Open this URL in your browser:<br>
                <a href="{flow.get('verification_uri','https://microsoft.com/devicelogin')}"
                   target="_blank"
                   style="color:#5aaaff;font-weight:700;word-break:break-all">
                  {flow.get('verification_uri','https://microsoft.com/devicelogin')}
                </a>
              </div>
              <div style="margin:10px 0 6px;font-size:12px;color:var(--mu)">
                <strong>Step 2</strong> — Enter this code:
              </div>
              <div style="background:var(--bg);border:1px solid var(--bd);border-radius:8px;
                          padding:10px;font-family:'JetBrains Mono',monospace;font-size:18px;
                          font-weight:700;color:var(--acc3);text-align:center;
                          letter-spacing:.16em;margin-bottom:10px">
                {flow.get('user_code','')}
              </div>
              <div class="od-hint">
                <strong>Step 3</strong> — Sign in with your Microsoft / work account,
                then click <strong>I've signed in</strong> below.
              </div>
            </div>""", unsafe_allow_html=True)

            col_check, col_cancel = st.columns(2)
            with col_check:
                if st.button("✅ I've signed in", use_container_width=True):
                    try:
                        tok = ms_poll_token(flow["device_code"])
                        if tok:
                            # Get display name
                            try:
                                import requests as _req
                                me = _req.get(f"{GRAPH}/me",
                                              headers={"Authorization": f"Bearer {tok['access_token']}"},
                                              timeout=10).json()
                                st.session_state.ms_user = me.get("displayName", me.get("userPrincipalName", ""))
                            except Exception:
                                pass
                            st.session_state.ms_token       = tok
                            st.session_state.ms_device_flow = None
                            st.success("✅ Signed in!")
                            st.rerun()
                        else:
                            st.warning("⏳ Not authenticated yet — finish signing in, then try again.")
                    except RuntimeError as e:
                        st.error(f"Authentication failed: {e}")
                        st.session_state.ms_device_flow = None
            with col_cancel:
                if st.button("✖ Cancel", use_container_width=True):
                    st.session_state.ms_device_flow = None
                    st.rerun()

        # ── Not yet started — show sign-in button ──────────────────────────────
        else:
            st.markdown("""
            <div class="od-panel">
              <div class="od-title">☁️ SharePoint / OneDrive</div>
              <div class="od-hint">
                Sign in with your Microsoft work account to access SharePoint
                document libraries and OneDrive folders directly.<br><br>
                No Azure app registration required — uses your existing
                Microsoft 365 credentials.
              </div>
            </div>""", unsafe_allow_html=True)

            if st.button("🔐 Sign in with Microsoft", use_container_width=True):
                with st.spinner("Starting sign-in…"):
                    try:
                        flow = ms_start_device_flow()
                        if "user_code" in flow:
                            st.session_state.ms_device_flow = flow
                            st.rerun()
                        else:
                            st.error(f"Could not start sign-in: {flow.get('error_description', flow)}")
                    except Exception as e:
                        st.error(f"Sign-in error: {e}")

        # Show fetched file list
        if st.session_state.fetched_files:
            st.markdown(
                f'<div style="font-size:12px;color:var(--mu);margin:8px 0 5px">'
                f'Last fetched — <strong style="color:var(--tx)">'
                f'{len(st.session_state.fetched_files)}</strong> file(s):</div>',
                unsafe_allow_html=True)
            for fname, fsize in st.session_state.fetched_files[:12]:
                sz = f"{fsize // 1024:,} KB" if fsize else ""
                st.markdown(
                    f'<div class="od-file">'
                    f'<span style="font-size:13px">📄</span>'
                    f'<span class="od-fn">{fname}</span>'
                    f'<span class="od-sz">{sz}</span>'
                    f'</div>',
                    unsafe_allow_html=True)
            if len(st.session_state.fetched_files) > 12:
                st.caption(f"…and {len(st.session_state.fetched_files) - 12} more")

    # ════ MODE B: Manual upload ════════════════════════════════════════════════
    else:
        st.markdown(
            '<div class="hint">📁 Upload PDF safety notices or Word reports directly.</div>',
            unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "files", accept_multiple_files=True,
            type=["pdf", "docx", "doc", "txt"],
            label_visibility="collapsed",
        )
        if uploaded and st.button("⚙️  Build Index", use_container_width=True):
            with st.spinner(f"Parsing {len(uploaded)} file(s)…"):
                st.session_state.index_ver += 1
                tuples = tuple((f.name, f.read()) for f in uploaded)
                vs, n, ps = build_index(tuples, st.session_state.index_ver)
            st.session_state.parse_stats   = ps
            st.session_state.fetched_files = []
            if vs:
                st.session_state.vectorstore   = vs
                st.session_state.chunk_count   = n
                st.session_state.indexed_files = [f.name for f in uploaded]
                st.success(f"✅ {n:,} chunks indexed!")
            else:
                failed = {f: s for f, s in ps.items() if "error" in s}
                if failed:
                    st.error("Parsing failed. See **Parse Report** below for details.")
                    for fname, s in failed.items():
                        st.error(f"📄 **{fname}**: {s['error']}")
                else:
                    st.error("No text content extracted. Files may be image-only PDFs (install Tesseract for OCR).")

    # ── Parse report ───────────────────────────────────────────────────────────
    if st.session_state.parse_stats:
        st.divider()
        st.markdown('<span class="sl">Parse Report</span>', unsafe_allow_html=True)
        st.markdown(
            f'<div style="background:var(--s2);border:1px solid var(--bd);border-radius:9px;padding:10px 13px">'
            f'{render_parse_stats(st.session_state.parse_stats)}</div>',
            unsafe_allow_html=True)

    # ── Scope selector ─────────────────────────────────────────────────────────
    scope_file = "All Documents"
    if st.session_state.vectorstore and len(st.session_state.indexed_files) > 1:
        st.divider()
        st.markdown('<span class="sl">Search Scope</span>', unsafe_allow_html=True)
        scope_file = st.selectbox(
            "Scope", ["All Documents"] + st.session_state.indexed_files,
            index=0, label_visibility="collapsed",
        )

    # ── Index stats ────────────────────────────────────────────────────────────
    if st.session_state.vectorstore:
        st.divider()
        st.markdown(f"""
        <div class="scard">
          <div class="sl">Index Stats</div>
          <div class="stat-r"><span class="sk">Chunks</span>
            <span class="sv">{st.session_state.chunk_count:,}</span></div>
          <div class="stat-r"><span class="sk">Files</span>
            <span class="sv">{len(st.session_state.indexed_files)}</span></div>
          <div class="stat-r"><span class="sk">Model</span>
            <span class="sv">{model_name}</span></div>
          <div class="stat-r"><span class="sk">Top-K</span>
            <span class="sv">{top_k}</span></div>
        </div>""", unsafe_allow_html=True)

        with st.expander("📄 Indexed files"):
            for f in st.session_state.indexed_files:
                st.markdown(f'<span class="src-tag">{f}</span>', unsafe_allow_html=True)

    st.divider()
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("🗑 Clear All", use_container_width=True):
            for k in ["messages", "indexed_files", "parse_stats", "pending", "fetched_files"]:
                st.session_state[k] = [] if k not in ("parse_stats", "pending") else ({} if k == "parse_stats" else "")
            st.session_state.vectorstore = None
            st.session_state.chunk_count = 0
            st.rerun()
    with col_b:
        if st.session_state.messages:
            st.download_button(
                "💾 Export",
                data=export_chat(st.session_state.messages),
                file_name="hse_chat.txt",
                mime="text/plain",
                use_container_width=True,
            )


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN CHAT AREA
# ═══════════════════════════════════════════════════════════════════════════════
is_ready  = bool(st.session_state.vectorstore) and bool(mistral_key)
scope_lbl = scope_file if scope_file != "All Documents" else "All Documents"
status_html = (
    f'<span class="bdg b-ok">● Ready · {scope_lbl}</span>'
    if is_ready else
    '<span class="bdg b-wn">● Waiting for documents</span>'
)

st.markdown(f"""
<div class="hdr">
  <div class="logo">🦺</div>
  <div>
    <div class="hdr-title">HSE Notifications Assistant</div>
    <div class="hdr-sub">OneDrive · PDF · DOCX · OCR · Keyword Search · Multi-Doc RAG</div>
  </div>
  <div style="margin-left:auto">{status_html}</div>
</div>""", unsafe_allow_html=True)

# ── Welcome screen ─────────────────────────────────────────────────────────────
if not st.session_state.messages:
    st.markdown("""
    <div class="welcome">
      <div class="wi">🦺</div>
      <div class="wt">Ask anything about your HSE notices</div>
      <div class="wx">
        Paste your OneDrive share link in the sidebar — the bot fetches and indexes
        all your PDF and Word HSE documents automatically.<br><br>
        Search by <strong style="color:#00e5b0">keyword</strong> or ask natural-language questions.
        Every answer cites its source document and highlights your search terms.
      </div>
      <div style="margin-top:18px">
        <span class="chip">PPE requirements for chemical handling</span>
        <span class="chip">Emergency evacuation procedure</span>
        <span class="chip">Incident reporting steps</span>
        <span class="chip">Hot work permit requirements</span>
        <span class="chip">Inspection checklist items</span>
        <span class="chip">Latest toolbox talk topics</span>
      </div>
    </div>""", unsafe_allow_html=True)

# ── Chat history ───────────────────────────────────────────────────────────────
else:
    st.markdown('<div class="chat">', unsafe_allow_html=True)
    for i, msg in enumerate(st.session_state.messages):
        if msg["role"] == "user":
            st.markdown(f"""
            <div class="row-u">
              <div class="av av-u">👤</div>
              <div class="bbl bbl-u">{msg["content"]}</div>
            </div>""", unsafe_allow_html=True)
        else:
            t_html = f'<span class="mt">⏱ {msg["time"]}</span>' if msg.get("time") else ""
            c_html = conf_badge(msg["confidence"]) if msg.get("confidence") else ""
            s_html = (f'<div class="src-block"><strong>Sources</strong><br>{msg["sources"]}</div>'
                      if msg.get("sources") else "")
            body   = highlight_keywords(msg["content"], msg.get("query", ""))
            st.markdown(f"""
            <div class="row-b">
              <div class="av av-b">🦺</div>
              <div style="max-width:78%">
                <div class="bbl bbl-b">{body}</div>
                <div class="meta-row">{t_html}{c_html}</div>
                {s_html}
              </div>
            </div>""", unsafe_allow_html=True)

            if msg.get("chunks"):
                with st.expander(f"🔍 View {len(msg['chunks'])} retrieved chunks"):
                    for j, ch in enumerate(msg["chunks"]):
                        pg  = f" · p{ch['page']+1}" if ch.get("page", "") != "" else ""
                        ocr = " 🔍OCR" if ch.get("ocr") else ""
                        st.markdown(
                            f'<div class="c-lbl">Chunk {j+1} — {ch["source"]}{pg}{ocr}</div>',
                            unsafe_allow_html=True)
                        st.markdown(
                            f'<div class="c-box">{ch["text"][:700]}</div>',
                            unsafe_allow_html=True)

            if msg.get("followups"):
                st.markdown("**💡 Suggested follow-ups:**")
                cols = st.columns(len(msg["followups"]))
                for ki, fu in enumerate(msg["followups"]):
                    with cols[ki]:
                        if st.button(fu, key=f"fu_{i}_{ki}", use_container_width=True):
                            st.session_state.pending = fu
                            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

# ── Input bar ─────────────────────────────────────────────────────────────────
st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
prefill = st.session_state.pending
if prefill:
    st.session_state.pending = ""

col_q, col_send = st.columns([6, 1])
with col_q:
    user_input = st.text_input(
        "q", label_visibility="collapsed",
        placeholder="Search HSE notifications… e.g. 'PPE for confined space entry'",
        value=prefill, key="user_input", disabled=not is_ready,
    )
with col_send:
    send = st.button("Send →", use_container_width=True, disabled=not is_ready)

if not is_ready:
    if not st.session_state.vectorstore:
        st.info("⬅  Paste your OneDrive share link in the sidebar and click **Fetch & Index Files** to begin.")
    else:
        st.info("⬅  Enter your Mistral AI API key in the sidebar.")

# ── Submit ────────────────────────────────────────────────────────────────────
if send and user_input.strip() and is_ready:
    question = user_input.strip()
    st.session_state.messages.append({"role": "user", "content": question})

    with st.spinner("🔍 Searching documents…"):
        t0     = time.time()
        result = run_rag(
            vs        = st.session_state.vectorstore,
            question  = question,
            model     = model_name,
            api_key   = mistral_key,
            temp      = temperature,
            k         = top_k,
            mode      = answer_mode,
            files     = st.session_state.indexed_files,
            messages  = st.session_state.messages,
            scope     = scope_file,
        )
        elapsed = time.time() - t0

    answer   = result["result"].strip()
    src_docs = result.get("source_documents", [])
    chunks   = [
        {
            "source": d.metadata.get("source_file", "?"),
            "page":   d.metadata.get("page", ""),
            "ocr":    d.metadata.get("ocr", False),
            "text":   d.page_content,
        }
        for d in src_docs
    ]
    fups = generate_followups(answer, question, mistral_key, model_name)

    st.session_state.messages.append({
        "role":       "assistant",
        "content":    answer,
        "query":      question,
        "sources":    format_sources(src_docs),
        "time":       f"{elapsed:.1f}s",
        "confidence": get_confidence(answer, src_docs),
        "chunks":     chunks,
        "followups":  fups,
    })
    st.rerun()
