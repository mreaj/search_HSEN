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
#  ONEDRIVE / SHAREPOINT — SHARED LINK FETCHER
# ═══════════════════════════════════════════════════════════════════════════════

def _encode_share_url(share_url: str) -> str:
    """Encode a share URL into the Graph API shares token format."""
    b64 = base64.urlsafe_b64encode(share_url.encode()).decode().rstrip("=")
    return f"u!{b64}"


def _is_likely_html(data: bytes) -> bool:
    """Return True if the bytes look like an HTML page rather than a binary file."""
    try:
        sniff = data[:512].decode("utf-8", errors="ignore").lower()
        return "<html" in sniff or "<!doctype" in sniff or "<head" in sniff
    except Exception:
        return False


def _download_file(url: str, session, timeout: int = 60) -> bytes | None:
    """Download a URL and return bytes, or None if it looks like HTML."""
    try:
        r = session.get(url, timeout=timeout, allow_redirects=True)
        r.raise_for_status()
        data = r.content
        if _is_likely_html(data):
            return None
        return data
    except Exception as e:
        logger.warning(f"Download error {url}: {e}")
        return None


def _try_graph_api(share_url: str, session) -> tuple[list, list, bool]:
    """
    Attempt to list / download files via the Graph API anonymous shares endpoint.
    Returns (file_pairs, errors, succeeded).
    """
    token    = _encode_share_url(share_url)
    api_base = f"https://graph.microsoft.com/v1.0/shares/{token}"
    results, errors = [], []

    try:
        r    = session.get(f"{api_base}/driveItem", timeout=20)
        item = r.json()
    except Exception as e:
        return [], [f"Graph API unreachable: {e}"], False

    if "error" in item:
        logger.info(f"Graph API error: {item['error'].get('message','')}")
        return [], [], False   # signal: try other strategies

    # ── Single-file share ──
    if "file" in item:
        ext = Path(item["name"]).suffix.lower()
        if ext not in SUPPORTED_EXTS:
            return [], [f"Unsupported file type: {ext}"], True
        dl = item.get("@microsoft.graph.downloadUrl", "")
        if dl:
            data = _download_file(dl, session)
            if data:
                results.append((item["name"], data))
            else:
                errors.append(f"Downloaded {item['name']} looks like HTML — link may have expired")
        return results, errors, True

    # ── Folder share ──
    children_url = f"{api_base}/driveItem/children"
    while children_url:
        try:
            cr   = session.get(children_url, timeout=20)
            data = cr.json()
        except Exception as e:
            errors.append(f"Could not list folder: {e}")
            break
        for child in data.get("value", []):
            if "file" not in child:
                continue
            ext = Path(child["name"]).suffix.lower()
            if ext not in SUPPORTED_EXTS:
                continue
            dl = child.get("@microsoft.graph.downloadUrl", "")
            if not dl:
                errors.append(f"No download URL for {child['name']}")
                continue
            content = _download_file(dl, session)
            if content:
                results.append((child["name"], content))
            else:
                errors.append(f"{child['name']}: download returned HTML (link may have expired)")
        children_url = data.get("@odata.nextLink")

    return results, errors, True


def _try_direct_download(share_url: str, session) -> tuple[list, list]:
    """
    Try the share URL itself as a direct file download.
    Works for single-file "download" links that end in ?download=1 etc.
    """
    results, errors = [], []
    try:
        # Append download=1 if not already present
        dl_url = share_url
        if "download=1" not in dl_url:
            sep    = "&" if "?" in dl_url else "?"
            dl_url = dl_url + sep + "download=1"

        r = session.get(dl_url, timeout=30, allow_redirects=True)
        r.raise_for_status()
        data = r.content

        if _is_likely_html(data):
            errors.append("Link redirected to a web page — not a direct file URL")
            return results, errors

        # Guess filename from Content-Disposition or URL
        cd    = r.headers.get("Content-Disposition", "")
        match = re.search(r'filename\*?=["\']?(?:UTF-8\'\')?([^"\'\n;]+)', cd, re.IGNORECASE)
        if match:
            fname = match.group(1).strip().strip('"\'')
        else:
            fname = Path(share_url.split("?")[0].rstrip("/")).name or "document"
        if not Path(fname).suffix:
            # Guess from Content-Type
            ct = r.headers.get("Content-Type", "")
            if "pdf" in ct:
                fname += ".pdf"
            elif "word" in ct or "openxml" in ct:
                fname += ".docx"

        ext = Path(fname).suffix.lower()
        if ext in SUPPORTED_EXTS:
            results.append((fname, data))
        else:
            errors.append(f"Downloaded file has unsupported type: {ext or 'unknown'}")
    except Exception as e:
        errors.append(f"Direct download failed: {e}")
    return results, errors


def fetch_shared_folder(share_url: str) -> tuple[list[tuple[str, bytes]], list[str]]:
    """
    Multi-strategy fetcher for OneDrive / SharePoint shared links.

    Strategy order:
      1. Microsoft Graph API anonymous shares  (works for most public OneDrive links)
      2. Direct download with ?download=1      (works for single-file links)
      3. Informative error with guidance

    Returns (file_pairs, errors)
    """
    import requests
    session = requests.Session()
    session.headers.update({"User-Agent": "HSEBot/1.0", "Accept": "application/json"})

    all_results, all_errors = [], []

    # ── Strategy 1: Graph API ──────────────────────────────────────────────────
    results, errors, graph_ok = _try_graph_api(share_url, session)
    all_results.extend(results)
    all_errors.extend(errors)

    if graph_ok and (results or errors):
        return all_results, all_errors   # Graph API handled it (even if 0 files found)

    # ── Strategy 2: Direct download ───────────────────────────────────────────
    logger.info("Graph API did not resolve link — trying direct download")
    results2, errors2 = _try_direct_download(share_url, session)
    all_results.extend(results2)
    all_errors.extend(errors2)

    if not all_results:
        all_errors.append(
            "Could not fetch files automatically. "
            "SharePoint organisation links often require sign-in. "
            "Please use the 'Upload Files' option instead: download the files from "
            "OneDrive manually and upload them here."
        )

    return all_results, all_errors


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
    "fetched_files": [],   # list of (name, size) from last OneDrive fetch
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
    source_tab = st.radio("src", ["🔗 OneDrive Share Link", "📂 Upload Files"],
                          label_visibility="collapsed")

    # ════ MODE A: OneDrive shared link ════════════════════════════════════════
    if source_tab == "🔗 OneDrive Share Link":
        st.markdown("""
        <div class="od-panel">
          <div class="od-title">☁ OneDrive / SharePoint</div>
          <div class="od-hint">
            In OneDrive, right-click your HSE folder → <strong>Share</strong> →
            set to <em>Anyone with the link can view</em> → copy the link and paste below.
          </div>
        </div>""", unsafe_allow_html=True)

        share_url = st.text_input(
            "url", placeholder="https://company.sharepoint.com/:f:/s/…",
            label_visibility="collapsed",
            help="Paste the public share link to your OneDrive/SharePoint folder",
        )

        if share_url and st.button("⬇️  Fetch & Index Files", use_container_width=True):
            with st.spinner("Connecting to OneDrive…"):
                file_pairs, errs = fetch_shared_folder(share_url.strip())

            if errs:
                for e in errs:
                    st.warning(f"⚠ {e}")

            if file_pairs:
                st.session_state.fetched_files = [(n, len(b)) for n, b in file_pairs]
                with st.spinner(f"Parsing & indexing {len(file_pairs)} file(s)…"):
                    st.session_state.index_ver += 1
                    vs, n, ps = build_index(tuple(file_pairs), st.session_state.index_ver)

                # Always store parse stats so errors appear in the Parse Report
                st.session_state.parse_stats = ps

                if vs:
                    st.session_state.vectorstore   = vs
                    st.session_state.chunk_count   = n
                    st.session_state.indexed_files = [p[0] for p in file_pairs]
                    st.success(f"✅ {n:,} chunks indexed from {len(file_pairs)} file(s)!")
                else:
                    # Show per-file errors from parse stats
                    failed = {f: s for f, s in ps.items() if "error" in s}
                    if failed:
                        st.error("Files were downloaded but parsing failed. See **Parse Report** below for details.")
                        for fname, s in failed.items():
                            st.error(f"📄 **{fname}**: {s['error']}")
                    else:
                        st.error("Files downloaded but no text content was extracted. They may be image-only PDFs without OCR support installed.")
            elif not errs:
                st.error("No supported files found at that link. Ensure the folder is public and contains PDF or DOCX files.")

        # Show fetched file list
        if st.session_state.fetched_files:
            st.markdown(f'<div style="font-size:12px;color:var(--mu);margin:8px 0 5px">Last fetched — <strong style="color:var(--tx)">{len(st.session_state.fetched_files)}</strong> file(s):</div>',
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
