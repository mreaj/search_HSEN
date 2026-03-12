import streamlit as st
import re
import io
import requests
import pandas as pd
import fitz  # PyMuPDF
from msal import ConfidentialClientApplication

# ======================
# ⚙️  CONFIGURATION — loaded from Streamlit Secrets (never hardcoded)
# Add these in: Streamlit Cloud → App → Settings → Secrets
# ======================
AZURE_CLIENT_ID     = st.secrets["AZURE_CLIENT_ID"]
AZURE_CLIENT_SECRET = st.secrets["AZURE_CLIENT_SECRET"]
AZURE_TENANT_ID     = st.secrets["AZURE_TENANT_ID"]
REDIRECT_URI        = st.secrets["REDIRECT_URI"]
SHAREPOINT_SITE_URL = st.secrets["SHAREPOINT_SITE_URL"]
SHAREPOINT_LIBRARY  = st.secrets.get("SHAREPOINT_LIBRARY", "Documents")

# ======================
# AUTH CONSTANTS
# ======================
AUTHORITY   = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
SCOPES      = ["Sites.Read.All", "Files.Read.All", "User.Read"]
GRAPH       = "https://graph.microsoft.com/v1.0"

# ======================
# PAGE CONFIG
# ======================
st.set_page_config(
    page_title="Document Search",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ======================
# CUSTOM CSS
# ======================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@300;400;500&display=swap');

    html, body, [class*="css"] { font-family: 'DM Mono', monospace; }

    .stApp { background: #f5f2eb; color: #1a1a1a; }

    [data-testid="stSidebar"] { background: #1a1a1a !important; border-right: none; }
    [data-testid="stSidebar"] * { color: #f5f2eb !important; }
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea {
        background: #2a2a2a !important; border: 1px solid #3a3a3a !important;
        color: #f5f2eb !important; font-family: 'DM Mono', monospace !important;
        border-radius: 2px !important;
    }
    [data-testid="stSidebar"] label {
        color: #666 !important; font-size: 0.68rem !important;
        text-transform: uppercase; letter-spacing: 0.1em;
    }

    .logo { font-family: 'Syne', sans-serif; font-weight: 800; font-size: 2.4rem;
        line-height: 1; color: #f5f2eb; letter-spacing: -0.03em; margin-bottom: 0.15rem; }
    .logo span { color: #e8c547; }
    .logo-sub { font-family: 'DM Mono', monospace; font-size: 0.62rem; color: #444;
        letter-spacing: 0.15em; text-transform: uppercase; margin-bottom: 2rem; }

    .main-title { font-family: 'Syne', sans-serif; font-weight: 800; font-size: 3rem;
        color: #1a1a1a; letter-spacing: -0.04em; line-height: 1; margin-bottom: 0.3rem; }
    .main-sub { font-family: 'DM Mono', monospace; font-size: 0.72rem;
        color: #999; letter-spacing: 0.04em; margin-bottom: 2rem; }

    .sec-label { font-family: 'DM Mono', monospace; font-size: 0.62rem; color: #888;
        text-transform: uppercase; letter-spacing: 0.12em; margin-bottom: 0.4rem; margin-top: 1rem; }

    .login-card { background: #fff; border: 1px solid #e0ddd5; border-radius: 6px;
        padding: 2.5rem 3rem; text-align: center; max-width: 420px;
        margin: 8rem auto 0; }
    .login-title { font-family: 'Syne', sans-serif; font-weight: 800; font-size: 1.6rem;
        color: #1a1a1a; letter-spacing: -0.03em; margin-bottom: 0.4rem; }
    .login-sub { font-family: 'DM Mono', monospace; font-size: 0.72rem;
        color: #999; margin-bottom: 1.8rem; line-height: 1.6; }

    .user-pill { background: #262626; border-radius: 3px; padding: 0.35rem 0.75rem;
        font-family: 'DM Mono', monospace; font-size: 0.65rem; color: #888;
        display: flex; align-items: center; gap: 0.4rem; margin-bottom: 1rem; }
    .user-dot { width: 6px; height: 6px; border-radius: 50%;
        background: #4caf50; flex-shrink: 0; }

    .result-card { background: #fff; border: 1px solid #e0ddd5; border-radius: 4px;
        padding: 1rem 1.2rem; margin-bottom: 0.55rem; position: relative; }
    .result-card::before { content: ''; position: absolute; left: 0; top: 0; bottom: 0;
        width: 3px; background: #e8c547; border-radius: 4px 0 0 4px; }
    .result-filename { font-family: 'DM Mono', monospace; font-size: 0.86rem;
        font-weight: 500; color: #1a1a1a; margin-bottom: 0.3rem; }
    .badge-row { display: flex; flex-wrap: wrap; gap: 0.25rem; margin-bottom: 0.45rem; }
    .badge { font-family: 'DM Mono', monospace; font-size: 0.62rem;
        padding: 0.12rem 0.4rem; border-radius: 2px; border: 1px solid; }
    .badge-ev { background: #fef9e0; border-color: #e8c547; color: #7a6200; }
    .badge-kw { background: #f0f0f0; border-color: #ccc; color: #555; }
    .badge-ft { background: #e8f5e9; border-color: #81c784; color: #2e7d32; }
    .dl-link { font-family: 'DM Mono', monospace; font-size: 0.72rem; color: #1a1a1a;
        text-decoration: none; border-bottom: 1px solid #e8c547; padding-bottom: 1px; }

    .stat-row { display: flex; gap: 0.8rem; margin-bottom: 1.3rem; flex-wrap: wrap; }
    .stat-pill { background: #1a1a1a; border-radius: 3px; padding: 0.45rem 0.9rem;
        display: flex; align-items: center; gap: 0.5rem; }
    .stat-num { font-family: 'Syne', sans-serif; font-weight: 700;
        font-size: 1.3rem; color: #e8c547; }
    .stat-lbl { font-family: 'DM Mono', monospace; font-size: 0.62rem;
        color: #777; text-transform: uppercase; letter-spacing: 0.08em; }

    .stButton > button { background: #1a1a1a !important; color: #f5f2eb !important;
        border: none !important; font-family: 'DM Mono', monospace !important;
        font-size: 0.78rem !important; letter-spacing: 0.08em !important;
        border-radius: 2px !important; }
    .stButton > button:hover { background: #333 !important; }

    .stDownloadButton > button { background: #e8c547 !important; color: #1a1a1a !important;
        border: none !important; font-family: 'DM Mono', monospace !important;
        font-size: 0.72rem !important; font-weight: 500 !important; border-radius: 2px !important; }

    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea {
        background: #fff !important; border: 1px solid #ddd !important;
        border-radius: 2px !important; font-family: 'DM Mono', monospace !important;
        font-size: 0.8rem !important; color: #1a1a1a !important; }
    .stTextInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus {
        border-color: #e8c547 !important; box-shadow: 0 0 0 2px rgba(232,197,71,0.15) !important; }

    label { font-family: 'DM Mono', monospace !important; font-size: 0.68rem !important;
        color: #888 !important; text-transform: uppercase; letter-spacing: 0.08em; }
    .stRadio label { font-size: 0.78rem !important; color: #1a1a1a !important;
        text-transform: none !important; letter-spacing: 0 !important; }
    .stCheckbox label { font-size: 0.78rem !important; color: #1a1a1a !important;
        text-transform: none !important; letter-spacing: 0 !important; }

    .stProgress > div > div { background: #e8c547 !important; }

    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    hr { border-color: #e0ddd5; margin: 1.3rem 0; }
</style>
""", unsafe_allow_html=True)

# ======================
# DEFAULTS
# ======================
DEFAULT_EVENT_NUMBERS = [
    "IN.0000092127", "IN.0000097889", "IN.0000133390",
    "IN.0000144100", "IN.0000220353", "IN.0000221017",
    "IN.0000263077", "IN.0000281719", "IN.0000312030",
    "IN.0000379870"
]
DEFAULT_KEYWORDS = [
    "fatal accident", "fatality", "serious motor vehicle incident",
    "serious road traffic accident", "electrocution incident",
    "fell inside the tower", "crushed by a frame", "plunged off a bridge",
    "trapped in the wreckage", "fatal injury", "under investigation",
    "mechanical completion", "guindaste", "Makro", "Maverick Creek wind farm"
]

# ======================
# SESSION STATE
# ======================
for k, v in [("token", None), ("user", None), ("results", []), ("searched", False)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ======================
# AUTH HELPERS
# ======================
def get_msal_app():
    return ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=AUTHORITY,
        client_credential=AZURE_CLIENT_SECRET,
    )

def get_auth_url() -> str:
    return get_msal_app().get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        state="streamlit_auth",
    )

def exchange_code_for_token(code: str):
    result = get_msal_app().acquire_token_by_authorization_code(
        code=code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    return result if "access_token" in result else None

def gh(token: str) -> dict:
    return {"Authorization": f"Bearer {token}"}

# ======================
# GRAPH API HELPERS
# ======================
@st.cache_data(show_spinner=False, ttl=300)
def get_site_and_drive(token: str):
    """Resolve SharePoint site → drive ID in one go."""
    # 1. Site ID
    parts = SHAREPOINT_SITE_URL.replace("https://", "").split("/", 1)
    host, path = parts[0], parts[1] if len(parts) > 1 else ""
    r = requests.get(f"{GRAPH}/sites/{host}:/{path}", headers=gh(token))
    r.raise_for_status()
    site_id = r.json()["id"]

    # 2. Drive ID — match by library name
    r = requests.get(f"{GRAPH}/sites/{site_id}/drives", headers=gh(token))
    r.raise_for_status()
    drives = r.json().get("value", [])
    drive_id = None
    for d in drives:
        if d.get("name", "").lower() == SHAREPOINT_LIBRARY.lower():
            drive_id = d["id"]
            break
    if not drive_id and drives:
        drive_id = drives[0]["id"]   # fallback: first drive

    return site_id, drive_id

def list_pdf_files(token: str, drive_id: str) -> list:
    files, url = [], f"{GRAPH}/drives/{drive_id}/root/children?$top=999"
    while url:
        r = requests.get(url, headers=gh(token))
        r.raise_for_status()
        data = r.json()
        for item in data.get("value", []):
            if item.get("name", "").lower().endswith(".pdf"):
                files.append({
                    "name":         item["name"],
                    "id":           item["id"],
                    "web_url":      item.get("webUrl", ""),
                    "download_url": item.get("@microsoft.graph.downloadUrl", ""),
                })
        url = data.get("@odata.nextLink")
    return files

def download_pdf(file_item: dict, token: str):
    dl = file_item.get("download_url")
    if dl:
        r = requests.get(dl, timeout=30)
        if r.status_code == 200:
            return r.content
    return None

# ======================
# PDF HELPER
# ======================
def extract_text(pdf_bytes: bytes) -> str:
    text = ""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            for page in doc:
                text += page.get_text("text") + "\n"
    except Exception:
        pass
    return text

# ======================
# HANDLE OAUTH CALLBACK
# ======================
query_params = st.query_params
if "code" in query_params and st.session_state.token is None:
    with st.spinner("Signing you in…"):
        result = exchange_code_for_token(query_params["code"])
    if result:
        st.session_state.token = result["access_token"]
        claims = result.get("id_token_claims", {})
        st.session_state.user = {
            "name":  claims.get("name", ""),
            "email": claims.get("preferred_username", ""),
        }
        st.query_params.clear()
        st.rerun()
    else:
        st.error("Authentication failed — please try signing in again.")
        st.query_params.clear()

# ======================
# NOT LOGGED IN
# ======================
if not st.session_state.token:
    auth_url = get_auth_url()

    # Must use window.top.location to break out of Streamlit iframe sandbox
    # A plain <a> tag or meta refresh gets blocked by the sandbox
    st.markdown(f'''
    <div class="login-card">
        <div class="login-title">Document Search</div>
        <div class="login-sub">
            Sign in with your Vestas Microsoft account<br>
            to search the SharePoint document library.
        </div>
        <button onclick="window.top.location.href=\'{auth_url}\'" style="
            display: inline-block;
            background: #1a1a1a;
            color: #f5f2eb;
            font-family: DM Mono, monospace;
            font-size: 0.82rem;
            letter-spacing: 0.07em;
            border: none;
            cursor: pointer;
            padding: 0.65rem 1.8rem;
            border-radius: 2px;
            margin-top: 0.5rem;
        ">Sign in with Microsoft →</button>
    </div>
    ''', unsafe_allow_html=True)
    st.stop()

# ======================
# SIDEBAR  (authenticated)
# ======================
with st.sidebar:
    st.markdown('<div class="logo">DOC<span>.</span><br>SCAN</div>', unsafe_allow_html=True)
    st.markdown('<div class="logo-sub">SharePoint PDF Search</div>', unsafe_allow_html=True)

    name  = st.session_state.user.get("name", "")
    email = st.session_state.user.get("email", "")
    st.markdown(
        f'<div class="user-pill"><div class="user-dot"></div>{name or email}</div>',
        unsafe_allow_html=True,
    )
    if st.button("Sign out", use_container_width=True):
        for k in ["token", "user", "results", "searched"]:
            st.session_state[k] = None if k in ("token", "user") else ([] if k == "results" else False)
        st.rerun()

    st.markdown("---")
    st.markdown(
        f'<div style="font-family:DM Mono,monospace;font-size:0.6rem;color:#555;">'
        f'SITE<br>'
        f'<span style="color:#888;word-break:break-all;">{SHAREPOINT_SITE_URL}</span><br><br>'
        f'LIBRARY<br>'
        f'<span style="color:#888;">{SHAREPOINT_LIBRARY}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown("---")
    max_files = st.slider("Max PDFs to scan", 10, 500, 100, 10)
    st.caption(f"Scanning up to **{max_files}** files per search")

# ======================
# MAIN
# ======================
st.markdown('<div class="main-title">Document<br>Search</div>', unsafe_allow_html=True)
st.markdown('<div class="main-sub">Scan SharePoint PDFs · event numbers · keywords · free text</div>', unsafe_allow_html=True)

col_left, col_right = st.columns([3, 2], gap="large")

with col_left:
    st.markdown('<div class="sec-label">Event Numbers — one per line</div>', unsafe_allow_html=True)
    event_numbers_raw = st.text_area("event_numbers", value="\n".join(DEFAULT_EVENT_NUMBERS),
                                      height=155, label_visibility="collapsed")
    st.markdown('<div class="sec-label">Description Keywords — one per line</div>', unsafe_allow_html=True)
    keywords_raw = st.text_area("keywords", value="\n".join(DEFAULT_KEYWORDS),
                                 height=155, label_visibility="collapsed")

with col_right:
    st.markdown('<div class="sec-label">Free Text Search</div>', unsafe_allow_html=True)
    free_text_query = st.text_input("free_text", placeholder="Any word or phrase...",
                                     label_visibility="collapsed")
    st.markdown('<div class="sec-label" style="margin-top:1.5rem;">Match Logic</div>', unsafe_allow_html=True)
    match_logic = st.radio("match_logic",
                            ["Match ANY criteria (OR)", "Match ALL criteria (AND)"],
                            label_visibility="collapsed")
    require_all = "AND" in match_logic

    st.markdown('<div class="sec-label" style="margin-top:1.5rem;">Export Format</div>', unsafe_allow_html=True)
    export_excel = st.checkbox("Excel (.xlsx)", value=True)
    export_csv   = st.checkbox("CSV (.csv)")

st.markdown("---")
run_btn = st.button("▶  Run Search")

# ======================
# SEARCH
# ======================
if run_btn:
    token = st.session_state.token

    ev_list   = [e.strip() for e in event_numbers_raw.splitlines() if e.strip()]
    kw_list   = [k.strip() for k in keywords_raw.splitlines()      if k.strip()]
    free_text = free_text_query.strip()

    ev_pat = re.compile(r"\b(" + "|".join(re.escape(n) for n in ev_list) + r")\b") if ev_list else None
    kw_pat = re.compile("|".join(re.escape(k) for k in kw_list), re.IGNORECASE)    if kw_list else None

    try:
        with st.spinner("Connecting to SharePoint…"):
            site_id, drive_id = get_site_and_drive(token)

        with st.spinner("Fetching file list…"):
            files = list_pdf_files(token, drive_id)
    except requests.HTTPError as e:
        st.error(f"SharePoint error: {e.response.status_code} — {e.response.text[:200]}")
        st.stop()

    if not files:
        st.warning("No PDF files found in that library.")
    else:
        total   = min(len(files), max_files)
        prog    = st.progress(0)
        status  = st.empty()
        results = []

        for i, f in enumerate(files[:total]):
            status.markdown(
                f'<span style="font-family:DM Mono,monospace;font-size:0.7rem;color:#999;">'
                f'[{i+1}/{total}] {f["name"]}</span>', unsafe_allow_html=True
            )
            prog.progress((i + 1) / total)

            pdf_bytes = download_pdf(f, token)
            if not pdf_bytes:
                continue

            text    = extract_text(pdf_bytes)
            m_ev    = list(set(ev_pat.findall(text)))                           if ev_pat    else []
            m_kw    = list(set(m.lower() for m in kw_pat.findall(text)))        if kw_pat    else []
            m_ft    = bool(re.search(re.escape(free_text), text, re.IGNORECASE)) if free_text else False

            if require_all:
                checks = []
                if ev_pat:    checks.append(bool(m_ev))
                if kw_pat:    checks.append(bool(m_kw))
                if free_text: checks.append(m_ft)
                matched = all(checks) if checks else False
            else:
                matched = bool(m_ev) or bool(m_kw) or m_ft

            if matched:
                results.append({
                    "File Name":             f["name"],
                    "SharePoint URL":        f["web_url"],
                    "Matched Event Numbers": ", ".join(m_ev),
                    "Matched Keywords":      ", ".join(m_kw),
                    "Free Text Match":       "Yes" if m_ft else "",
                    "_ev": m_ev, "_kw": m_kw, "_ft": m_ft,
                })

        prog.empty(); status.empty()
        st.session_state.results  = results
        st.session_state.searched = True

# ======================
# RESULTS
# ======================
results = st.session_state.results

if st.session_state.searched:
    st.markdown("---")

    if not results:
        st.markdown(
            '<div style="text-align:center;padding:3rem;font-family:DM Mono,monospace;'
            'font-size:0.82rem;color:#aaa;">No matching documents found.</div>',
            unsafe_allow_html=True,
        )
    else:
        n_ev = sum(1 for r in results if r["_ev"])
        n_kw = sum(1 for r in results if r["_kw"])
        n_ft = sum(1 for r in results if r["_ft"])

        st.markdown(
            f'<div class="stat-row">'
            f'<div class="stat-pill"><div class="stat-num">{len(results)}</div><div class="stat-lbl">Files matched</div></div>'
            f'<div class="stat-pill"><div class="stat-num">{n_ev}</div><div class="stat-lbl">Event hits</div></div>'
            f'<div class="stat-pill"><div class="stat-num">{n_kw}</div><div class="stat-lbl">Keyword hits</div></div>'
            f'<div class="stat-pill"><div class="stat-num">{n_ft}</div><div class="stat-lbl">Free text hits</div></div>'
            f'</div>', unsafe_allow_html=True,
        )

        export_df = pd.DataFrame([{k: v for k, v in r.items() if not k.startswith("_")} for r in results])
        ec1, ec2, _ = st.columns([1, 1, 6])
        if export_excel:
            buf = io.BytesIO()
            export_df.to_excel(buf, index=False)
            ec1.download_button("⬇ Excel", data=buf.getvalue(), file_name="results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if export_csv:
            ec2.download_button("⬇ CSV", data=export_df.to_csv(index=False).encode(),
                file_name="results.csv", mime="text/csv")

        st.markdown('<div class="sec-label" style="margin-top:1rem;margin-bottom:0.7rem;">Matched Files</div>', unsafe_allow_html=True)

        for r in results:
            badges  = "".join(f'<span class="badge badge-ev">{ev}</span>' for ev in r["_ev"])
            badges += "".join(f'<span class="badge badge-kw">{kw}</span>' for kw in r["_kw"])
            if r["_ft"]:
                badges += '<span class="badge badge-ft">free text ✓</span>'

            st.markdown(f"""
            <div class="result-card">
                <div class="result-filename">📄 {r["File Name"]}</div>
                <div class="badge-row">{badges}</div>
                <a class="dl-link" href="{r["SharePoint URL"]}" target="_blank">Open / Download ↗</a>
            </div>
            """, unsafe_allow_html=True)
