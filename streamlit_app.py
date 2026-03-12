import streamlit as st
import re
import io
import time
import requests
import pandas as pd
import fitz  # PyMuPDF
from msal import PublicClientApplication

# ======================
# CONFIGURATION
# Only CLIENT_ID and TENANT_ID needed — no secret, no redirect URI
# Add in Streamlit Cloud → App → Settings → Secrets
# ======================
AZURE_CLIENT_ID     = st.secrets["AZURE_CLIENT_ID"]
SHAREPOINT_SITE_URL = st.secrets["SHAREPOINT_SITE_URL"]
SHAREPOINT_LIBRARY  = st.secrets.get("SHAREPOINT_LIBRARY", "Documents")

# Use "organizations" so any work/school Microsoft account can sign in
# regardless of which tenant the app registration lives in
AUTHORITY = "https://login.microsoftonline.com/organizations"
SCOPES    = ["Sites.Read.All", "Files.Read.All", "User.Read"]
GRAPH     = "https://graph.microsoft.com/v1.0"

# ======================
# PAGE CONFIG
# ======================
st.set_page_config(
    page_title="Document Search",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ======================
# CUSTOM CSS
# ======================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Mono:wght@300;400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Mono', monospace; }
.stApp { background: #f5f2eb; color: #1a1a1a; }

[data-testid="stSidebar"] { background: #1a1a1a !important; }
[data-testid="stSidebar"] * { color: #d4d0c8 !important; }
[data-testid="stSidebar"] input, [data-testid="stSidebar"] textarea {
    background: #262626 !important; border: 1px solid #383838 !important;
    color: #d4d0c8 !important; font-family: 'DM Mono', monospace !important;
    border-radius: 2px !important; font-size: 0.8rem !important; }
[data-testid="stSidebar"] label {
    color: #555 !important; font-size: 0.62rem !important;
    text-transform: uppercase; letter-spacing: 0.1em; }

.logo { font-family: 'Syne', sans-serif; font-weight: 800; font-size: 2.2rem;
    line-height: 1; color: #f5f2eb; letter-spacing: -0.03em; }
.logo em { font-style: normal; color: #e8c547; }
.logo-sub { font-family: 'DM Mono', monospace; font-size: 0.58rem; color: #444;
    letter-spacing: 0.15em; text-transform: uppercase; margin: 0.25rem 0 2rem; }

.pg-title { font-family: 'Syne', sans-serif; font-weight: 800; font-size: 2.8rem;
    color: #1a1a1a; letter-spacing: -0.04em; line-height: 1; }
.pg-sub { font-family: 'DM Mono', monospace; font-size: 0.7rem; color: #999;
    letter-spacing: 0.04em; margin: 0.3rem 0 2rem; }

.sl { font-family: 'DM Mono', monospace; font-size: 0.6rem; color: #888;
    text-transform: uppercase; letter-spacing: 0.12em; margin: 1rem 0 0.35rem; }

/* Device code login card */
.dc-card { background: #fff; border: 1px solid #e0ddd5; border-radius: 6px;
    padding: 2rem 2.5rem; max-width: 500px; margin: 4rem auto; }
.dc-title { font-family: 'Syne', sans-serif; font-weight: 800; font-size: 1.5rem;
    color: #1a1a1a; letter-spacing: -0.03em; margin-bottom: 0.5rem; }
.dc-step { font-family: 'DM Mono', monospace; font-size: 0.72rem;
    color: #666; margin-bottom: 0.4rem; line-height: 1.7; }
.dc-code { font-family: 'DM Mono', monospace; font-size: 1.6rem; font-weight: 500;
    color: #1a1a1a; background: #f5f2eb; border: 2px solid #e8c547;
    border-radius: 4px; padding: 0.5rem 1.2rem; display: inline-block;
    letter-spacing: 0.15em; margin: 0.8rem 0; }
.dc-url { font-family: 'DM Mono', monospace; font-size: 0.8rem;
    color: #e8c547; text-decoration: none; border-bottom: 1px solid #e8c547; }

.user-pill { background: #262626; border-radius: 3px; padding: 0.35rem 0.75rem;
    font-family: 'DM Mono', monospace; font-size: 0.65rem; color: #888;
    display: flex; align-items: center; gap: 0.4rem; margin-bottom: 1rem; }
.user-dot { width: 6px; height: 6px; border-radius: 50%; background: #4caf50; flex-shrink: 0; }

.result-card { background: #fff; border: 1px solid #e0ddd5; border-radius: 4px;
    padding: 1rem 1.2rem; margin-bottom: 0.5rem; position: relative; }
.result-card::before { content: ''; position: absolute; left: 0; top: 0; bottom: 0;
    width: 3px; background: #e8c547; border-radius: 4px 0 0 4px; }
.result-filename { font-family: 'DM Mono', monospace; font-size: 0.85rem;
    font-weight: 500; color: #1a1a1a; margin-bottom: 0.3rem; }
.badge-row { display: flex; flex-wrap: wrap; gap: 0.25rem; margin-bottom: 0.45rem; }
.badge { font-family: 'DM Mono', monospace; font-size: 0.6rem;
    padding: 0.1rem 0.38rem; border-radius: 2px; border: 1px solid; }
.b-ev { background: #fef9e0; border-color: #e8c547; color: #7a6200; }
.b-kw { background: #f0f0f0; border-color: #ccc; color: #555; }
.b-ft { background: #e8f5e9; border-color: #81c784; color: #2e7d32; }
.dl-link { font-family: 'DM Mono', monospace; font-size: 0.7rem; color: #1a1a1a;
    text-decoration: none; border-bottom: 1px solid #e8c547; padding-bottom: 1px; }

.stat-row { display: flex; gap: 0.7rem; margin-bottom: 1.2rem; flex-wrap: wrap; }
.stat-pill { background: #1a1a1a; border-radius: 3px; padding: 0.4rem 0.85rem;
    display: flex; align-items: center; gap: 0.45rem; }
.stat-n { font-family: 'Syne', sans-serif; font-weight: 700; font-size: 1.25rem; color: #e8c547; }
.stat-l { font-family: 'DM Mono', monospace; font-size: 0.6rem;
    color: #777; text-transform: uppercase; letter-spacing: 0.08em; }

.stButton > button { background: #1a1a1a !important; color: #f5f2eb !important;
    border: none !important; font-family: 'DM Mono', monospace !important;
    font-size: 0.78rem !important; letter-spacing: 0.07em !important;
    border-radius: 2px !important; }
.stButton > button:hover { background: #333 !important; }
.stDownloadButton > button { background: #e8c547 !important; color: #1a1a1a !important;
    border: none !important; font-family: 'DM Mono', monospace !important;
    font-size: 0.72rem !important; font-weight: 500 !important; border-radius: 2px !important; }

.stTextInput > div > div > input, .stTextArea > div > div > textarea {
    background: #fff !important; border: 1px solid #ddd !important;
    border-radius: 2px !important; font-family: 'DM Mono', monospace !important;
    font-size: 0.8rem !important; color: #1a1a1a !important; }
.stTextInput > div > div > input:focus, .stTextArea > div > div > textarea:focus {
    border-color: #e8c547 !important; box-shadow: 0 0 0 2px rgba(232,197,71,0.15) !important; }
label { font-family: 'DM Mono', monospace !important; font-size: 0.65rem !important;
    color: #888 !important; text-transform: uppercase; letter-spacing: 0.08em; }
.stRadio label { font-size: 0.78rem !important; color: #1a1a1a !important;
    text-transform: none !important; letter-spacing: 0 !important; }
.stCheckbox label { font-size: 0.78rem !important; color: #1a1a1a !important;
    text-transform: none !important; letter-spacing: 0 !important; }
.stProgress > div > div { background: #e8c547 !important; }
#MainMenu { visibility: hidden; } footer { visibility: hidden; }
hr { border-color: #e0ddd5; margin: 1.2rem 0; }
</style>
""", unsafe_allow_html=True)

# ======================
# DEFAULTS
# ======================
DEFAULT_EVENT_NUMBERS = [
    "IN.0000092127", "IN.0000097889", "IN.0000133390",
    "IN.0000144100", "IN.0000220353", "IN.0000221017",
    "IN.0000263077", "IN.0000281719", "IN.0000312030",
    "IN.0000379870",
]
DEFAULT_KEYWORDS = [
    "fatal accident", "fatality", "serious motor vehicle incident",
    "serious road traffic accident", "electrocution incident",
    "fell inside the tower", "crushed by a frame", "plunged off a bridge",
    "trapped in the wreckage", "fatal injury", "under investigation",
    "mechanical completion", "guindaste", "Makro", "Maverick Creek wind farm",
]

# ======================
# SESSION STATE
# ======================
for k, v in [("token", None), ("user", None), ("results", []),
             ("searched", False), ("device_flow", None)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ======================
# MSAL PUBLIC CLIENT (device code — no secret needed)
# ======================
@st.cache_resource
def get_msal_app():
    return PublicClientApplication(
        AZURE_CLIENT_ID,
        authority=AUTHORITY,
    )

# ======================
# GRAPH HELPERS
# ======================
def gh(token):
    return {"Authorization": f"Bearer {token}"}

@st.cache_data(show_spinner=False, ttl=300)
def get_site_and_drive(token: str):
    parts    = SHAREPOINT_SITE_URL.replace("https://", "").split("/", 1)
    host     = parts[0]
    path     = parts[1] if len(parts) > 1 else ""
    r        = requests.get(f"{GRAPH}/sites/{host}:/{path}", headers=gh(token))
    r.raise_for_status()
    site_id  = r.json()["id"]

    r        = requests.get(f"{GRAPH}/sites/{site_id}/drives", headers=gh(token))
    r.raise_for_status()
    drives   = r.json().get("value", [])
    drive_id = next(
        (d["id"] for d in drives if d.get("name","").lower() == SHAREPOINT_LIBRARY.lower()),
        drives[0]["id"] if drives else None
    )
    return site_id, drive_id

def list_pdf_files(token, drive_id):
    files, url = [], f"{GRAPH}/drives/{drive_id}/root/children?$top=999"
    while url:
        r    = requests.get(url, headers=gh(token))
        r.raise_for_status()
        data = r.json()
        for item in data.get("value", []):
            if item.get("name","").lower().endswith(".pdf"):
                files.append({
                    "name":         item["name"],
                    "web_url":      item.get("webUrl", ""),
                    "download_url": item.get("@microsoft.graph.downloadUrl", ""),
                })
        url = data.get("@odata.nextLink")
    return files

def download_pdf(f, token):
    dl = f.get("download_url")
    if dl:
        r = requests.get(dl, timeout=30)
        if r.status_code == 200:
            return r.content
    return None

def extract_text(pdf_bytes):
    text = ""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            for page in doc:
                text += page.get_text("text") + "\n"
    except Exception:
        pass
    return text

# ======================
# SIDEBAR
# ======================
with st.sidebar:
    st.markdown('<div class="logo">DOC<em>.</em><br>SCAN</div>', unsafe_allow_html=True)
    st.markdown('<div class="logo-sub">SharePoint PDF Search</div>', unsafe_allow_html=True)

    if st.session_state.token:
        name  = st.session_state.user.get("name", "")
        email = st.session_state.user.get("email", "")
        st.markdown(
            f'<div class="user-pill"><div class="user-dot"></div>{name or email}</div>',
            unsafe_allow_html=True,
        )
        if st.button("Sign out", use_container_width=True):
            for k in ["token", "user", "results", "searched", "device_flow"]:
                st.session_state[k] = None if k in ("token","user","device_flow") \
                                       else ([] if k == "results" else False)
            st.rerun()
        st.markdown("---")
        max_files = st.slider("Max PDFs to scan", 10, 500, 100, 10)
    else:
        max_files = 100

# ======================
# NOT LOGGED IN — DEVICE CODE FLOW
# ======================
if not st.session_state.token:
    st.markdown('<div class="pg-title">Document<br>Search</div>', unsafe_allow_html=True)

    app = get_msal_app()

    # Step 1 — initiate device flow (only once)
    if st.session_state.device_flow is None:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            st.error(f"Could not start login: {flow.get('error_description', flow)}")
            st.stop()
        st.session_state.device_flow = flow

    flow      = st.session_state.device_flow
    user_code = flow["user_code"]
    verify_url = flow["verification_uri"]  # https://microsoft.com/devicelogin

    st.markdown(f"""
    <div class="dc-card">
        <div class="dc-title">Sign in to continue</div>
        <div class="dc-step">1. Open this URL in your browser:</div>
        <div style="margin-bottom:0.8rem;">
            <a class="dc-url" href="{verify_url}" target="_blank">{verify_url}</a>
        </div>
        <div class="dc-step">2. Enter this code:</div>
        <div class="dc-code">{user_code}</div>
        <div class="dc-step" style="margin-top:0.8rem; color:#aaa;">
            3. Then click <strong>Check sign-in</strong> below
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([2, 1.5, 2])
    with col2:
        check_btn = st.button("✓  Check sign-in", use_container_width=True)

    if check_btn:
        with st.spinner("Checking…"):
            result = app.acquire_token_by_device_flow(flow, exit_condition=lambda f: True)

        if "access_token" in result:
            st.session_state.token = result["access_token"]
            claims = result.get("id_token_claims", {})
            st.session_state.user  = {
                "name":  claims.get("name", ""),
                "email": claims.get("preferred_username", ""),
            }
            st.session_state.device_flow = None
            st.rerun()
        else:
            err = result.get("error", "")
            if err == "authorization_pending":
                st.warning("Not signed in yet — complete the sign-in in your browser, then click Check again.")
            elif err == "expired_token":
                st.session_state.device_flow = None
                st.warning("Code expired — refreshing…")
                st.rerun()
            else:
                st.error(f"Sign-in failed: {result.get('error_description', result)}")

    st.stop()

# ======================
# MAIN APP (authenticated)
# ======================
st.markdown('<div class="pg-title">Document<br>Search</div>', unsafe_allow_html=True)
st.markdown('<div class="pg-sub">Scan SharePoint PDFs · event numbers · keywords · free text</div>', unsafe_allow_html=True)

col_left, col_right = st.columns([3, 2], gap="large")

with col_left:
    st.markdown('<div class="sl">Event Numbers — one per line</div>', unsafe_allow_html=True)
    event_numbers_raw = st.text_area("ev", value="\n".join(DEFAULT_EVENT_NUMBERS),
                                      height=155, label_visibility="collapsed")
    st.markdown('<div class="sl">Description Keywords — one per line</div>', unsafe_allow_html=True)
    keywords_raw = st.text_area("kw", value="\n".join(DEFAULT_KEYWORDS),
                                 height=155, label_visibility="collapsed")

with col_right:
    st.markdown('<div class="sl">Free Text Search</div>', unsafe_allow_html=True)
    free_text_query = st.text_input("ft", placeholder="Any word or phrase…",
                                     label_visibility="collapsed")
    st.markdown('<div class="sl" style="margin-top:1.4rem;">Match Logic</div>', unsafe_allow_html=True)
    match_logic = st.radio("ml",
                            ["Match ANY criteria (OR)", "Match ALL criteria (AND)"],
                            label_visibility="collapsed")
    require_all = "AND" in match_logic
    st.markdown('<div class="sl" style="margin-top:1.4rem;">Export</div>', unsafe_allow_html=True)
    export_excel = st.checkbox("Excel (.xlsx)", value=True)
    export_csv   = st.checkbox("CSV (.csv)")

st.markdown("---")
run_btn = st.button("▶  Run Search")

# ======================
# SEARCH
# ======================
if run_btn:
    token     = st.session_state.token
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
        st.error(f"SharePoint error {e.response.status_code}: {e.response.text[:200]}")
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
                f'<span style="font-family:DM Mono,monospace;font-size:0.68rem;color:#999;">'
                f'[{i+1}/{total}] {f["name"]}</span>', unsafe_allow_html=True)
            prog.progress((i + 1) / total)

            pdf_bytes = download_pdf(f, token)
            if not pdf_bytes:
                continue

            text  = extract_text(pdf_bytes)
            m_ev  = list(set(ev_pat.findall(text)))                             if ev_pat    else []
            m_kw  = list(set(m.lower() for m in kw_pat.findall(text)))          if kw_pat    else []
            m_ft  = bool(re.search(re.escape(free_text), text, re.IGNORECASE))  if free_text else False

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
            unsafe_allow_html=True)
    else:
        n_ev = sum(1 for r in results if r["_ev"])
        n_kw = sum(1 for r in results if r["_kw"])
        n_ft = sum(1 for r in results if r["_ft"])

        st.markdown(
            f'<div class="stat-row">'
            f'<div class="stat-pill"><div class="stat-n">{len(results)}</div><div class="stat-l">matched</div></div>'
            f'<div class="stat-pill"><div class="stat-n">{n_ev}</div><div class="stat-l">event hits</div></div>'
            f'<div class="stat-pill"><div class="stat-n">{n_kw}</div><div class="stat-l">keyword hits</div></div>'
            f'<div class="stat-pill"><div class="stat-n">{n_ft}</div><div class="stat-l">free text</div></div>'
            f'</div>', unsafe_allow_html=True)

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

        st.markdown('<div class="sl" style="margin-top:1rem;margin-bottom:0.7rem;">Matched Files</div>',
                    unsafe_allow_html=True)

        for r in results:
            badges  = "".join(f'<span class="badge b-ev">{ev}</span>' for ev in r["_ev"])
            badges += "".join(f'<span class="badge b-kw">{kw}</span>' for kw in r["_kw"])
            if r["_ft"]:
                badges += '<span class="badge b-ft">free text ✓</span>'

            st.markdown(f"""
            <div class="result-card">
                <div class="result-filename">📄 {r["File Name"]}</div>
                <div class="badge-row">{badges}</div>
                <a class="dl-link" href="{r["SharePoint URL"]}" target="_blank">Open / Download ↗</a>
            </div>
            """, unsafe_allow_html=True)
