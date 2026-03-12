import streamlit as st
import re
import io
import requests
import pandas as pd
import fitz  # PyMuPDF

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
if "results" not in st.session_state:
    st.session_state.results = []
if "searched" not in st.session_state:
    st.session_state.searched = False

# ======================
# SHAREPOINT FUNCTIONS (No Auth — Public Library)
# ======================
def get_sharepoint_files(site_url: str, library_name: str):
    """List PDFs from a publicly accessible SharePoint document library."""
    site_url = site_url.rstrip("/")
    api_url = (
        f"{site_url}/_api/web/lists/getbytitle('{library_name}')/items"
        f"?$select=FileLeafRef,FileRef,ID"
        f"&$filter=substringof('.pdf',FileLeafRef)"
        f"&$top=5000"
    )
    headers = {"Accept": "application/json;odata=verbose"}
    try:
        resp = requests.get(api_url, headers=headers, timeout=30)
        if resp.status_code == 403:
            return [], "Access denied (403). The library may not be publicly accessible."
        if resp.status_code == 404:
            return [], f"Library '{library_name}' not found (404). Check the library name."
        resp.raise_for_status()
        items = resp.json().get("d", {}).get("results", [])
        return [
            {"name": it.get("FileLeafRef", ""), "server_relative_url": it.get("FileRef", "")}
            for it in items
        ], None
    except Exception as e:
        return [], str(e)


def download_pdf_bytes(site_url: str, server_relative_url: str):
    """Download a PDF file from a public SharePoint library."""
    domain = "/".join(site_url.rstrip("/").split("/")[:3])
    file_url = domain + server_relative_url
    try:
        resp = requests.get(file_url, timeout=30)
        resp.raise_for_status()
        return resp.content, None
    except Exception as e:
        return None, str(e)


def extract_text(pdf_bytes: bytes) -> str:
    text = ""
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            for page in doc:
                text += page.get_text("text") + "\n"
    except Exception:
        pass
    return text


def build_file_url(site_url: str, server_relative_url: str) -> str:
    domain = "/".join(site_url.rstrip("/").split("/")[:3])
    return domain + server_relative_url


# ======================
# SIDEBAR
# ======================
with st.sidebar:
    st.markdown('<div class="logo">DOC<span>.</span><br>SCAN</div>', unsafe_allow_html=True)
    st.markdown('<div class="logo-sub">SharePoint PDF Search</div>', unsafe_allow_html=True)

    site_url = st.text_input(
        "Site URL",
        placeholder="https://contoso.sharepoint.com/sites/YourSite"
    )
    library_name = st.text_input("Library Name", value="Documents", placeholder="e.g. Notifications")

    st.markdown("---")
    max_files = st.slider("Max PDFs to scan", 10, 500, 100, 10)
    st.caption(f"Scanning up to **{max_files}** files per search")

    st.markdown("---")
    st.markdown(
        '<span style="font-family:DM Mono,monospace;font-size:0.62rem;color:#444;line-height:1.6;">'
        'No login required.<br>Library must be set to<br>public / anonymous access.'
        '</span>',
        unsafe_allow_html=True
    )

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
    export_csv = st.checkbox("CSV (.csv)")

st.markdown("---")
run_btn = st.button("▶  Run Search")

# ======================
# SEARCH
# ======================
if run_btn:
    if not site_url.strip():
        st.error("Please enter the SharePoint Site URL in the sidebar.")
    else:
        event_numbers = [e.strip() for e in event_numbers_raw.splitlines() if e.strip()]
        keywords = [k.strip() for k in keywords_raw.splitlines() if k.strip()]
        free_text = free_text_query.strip()

        ev_pattern = re.compile(r"\b(" + "|".join(re.escape(n) for n in event_numbers) + r")\b") if event_numbers else None
        kw_pattern = re.compile("|".join(re.escape(k) for k in keywords), re.IGNORECASE) if keywords else None

        with st.spinner("Fetching file list from SharePoint..."):
            files, err = get_sharepoint_files(site_url.strip(), library_name.strip())

        if err:
            st.error(f"Error: {err}")
        elif not files:
            st.warning("No PDF files found in that library.")
        else:
            total = min(len(files), max_files)
            prog = st.progress(0)
            status = st.empty()
            results = []

            for i, f in enumerate(files[:total]):
                status.markdown(
                    f'<span style="font-family:DM Mono,monospace;font-size:0.7rem;color:#999;">'
                    f'[{i+1}/{total}] {f["name"]}</span>', unsafe_allow_html=True
                )
                prog.progress((i + 1) / total)

                pdf_bytes, _ = download_pdf_bytes(site_url.strip(), f["server_relative_url"])
                if not pdf_bytes:
                    continue

                text = extract_text(pdf_bytes)
                m_events = list(set(ev_pattern.findall(text))) if ev_pattern else []
                m_kws = list(set(m.lower() for m in kw_pattern.findall(text))) if kw_pattern else []
                m_free = bool(re.search(re.escape(free_text), text, re.IGNORECASE)) if free_text else False

                if require_all:
                    checks = []
                    if ev_pattern: checks.append(bool(m_events))
                    if kw_pattern: checks.append(bool(m_kws))
                    if free_text: checks.append(m_free)
                    matched = all(checks) if checks else False
                else:
                    matched = bool(m_events) or bool(m_kws) or m_free

                if matched:
                    results.append({
                        "File Name": f["name"],
                        "URL": build_file_url(site_url.strip(), f["server_relative_url"]),
                        "Matched Event Numbers": ", ".join(m_events),
                        "Matched Keywords": ", ".join(m_kws),
                        "Free Text Match": "Yes" if m_free else "",
                        "_ev": m_events, "_kw": m_kws, "_ft": m_free,
                    })

            prog.empty()
            status.empty()
            st.session_state.results = results
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
            unsafe_allow_html=True
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
            f'</div>', unsafe_allow_html=True
        )

        # Export buttons
        export_df = pd.DataFrame([{k: v for k, v in r.items() if not k.startswith("_")} for r in results])
        ec1, ec2, _ = st.columns([1, 1, 6])
        if export_excel:
            buf = io.BytesIO()
            export_df.to_excel(buf, index=False)
            ec1.download_button("⬇ Excel", data=buf.getvalue(),
                                file_name="search_results.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if export_csv:
            ec2.download_button("⬇ CSV", data=export_df.to_csv(index=False).encode(),
                                file_name="search_results.csv", mime="text/csv")

        st.markdown('<div class="sec-label" style="margin-top:1rem;margin-bottom:0.7rem;">Matched Files</div>', unsafe_allow_html=True)

        for r in results:
            badges = "".join(f'<span class="badge badge-ev">{ev}</span>' for ev in r["_ev"])
            badges += "".join(f'<span class="badge badge-kw">{kw}</span>' for kw in r["_kw"])
            if r["_ft"]:
                badges += '<span class="badge badge-ft">free text ✓</span>'

            st.markdown(f"""
            <div class="result-card">
                <div class="result-filename">📄 {r["File Name"]}</div>
                <div class="badge-row">{badges}</div>
                <a class="dl-link" href="{r["URL"]}" target="_blank">Open / Download ↗</a>
            </div>
            """, unsafe_allow_html=True)
