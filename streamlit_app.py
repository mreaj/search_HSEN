import os, io, re, streamlit as st
from pathlib import Path

st.set_page_config(page_title="HSE Notifications Bot", layout="wide")

# ───────── LOAD SECRETS ─────────
TENANT_ID     = st.secrets["TENANT_ID"]
CLIENT_ID     = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
MISTRAL_KEY   = st.secrets["MISTRAL_API_KEY"]

GRAPH = "https://graph.microsoft.com/v1.0"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

SUPPORTED_EXTS = {".pdf", ".docx"}

# ───────── SIMPLE WHITE UI ─────────
st.markdown("""
<style>
body { background:white; color:#1e293b; }
</style>
""", unsafe_allow_html=True)


# ───────── AUTH (NO LOGIN) ─────────
def get_token():
    import requests
    r = requests.post(TOKEN_URL, data={
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    })
    data = r.json()
    return data["access_token"]


def graph_get(url, token):
    import requests
    if not url.startswith("http"):
        url = GRAPH + url
    return requests.get(url, headers={"Authorization": f"Bearer {token}"}).json()


def download(url, token):
    import requests
    return requests.get(url, headers={"Authorization": f"Bearer {token}"}).content


# ───────── FETCH SHAREPOINT FILES ─────────
def get_files(token, sp_url):
    hostname = sp_url.split("/")[2]
    site_path = "/" + sp_url.split("/", 3)[3].split("/")[0] + "/" + sp_url.split("/", 4)[4]

    site = graph_get(f"/sites/{hostname}:{site_path}", token)
    site_id = site["id"]

    drives = graph_get(f"/sites/{site_id}/drives", token)["value"]
    drive_id = drives[0]["id"]

    items = graph_get(f"/drives/{drive_id}/root/children", token)["value"]

    files = []
    for item in items:
        if "file" in item:
            if Path(item["name"]).suffix.lower() in SUPPORTED_EXTS:
                data = download(item["@microsoft.graph.downloadUrl"], token)
                files.append((item["name"], data))

    return files


# ───────── PARSERS ─────────
from langchain_core.documents import Document

def parse_pdf(data):
    import pdfplumber
    docs = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for i, p in enumerate(pdf.pages):
            txt = p.extract_text() or ""
            if txt.strip():
                docs.append(Document(page_content=txt, metadata={"page": i}))
    return docs


def parse_docx(data):
    import docx
    doc = docx.Document(io.BytesIO(data))
    return [Document(page_content="\n".join(p.text for p in doc.paragraphs), metadata={})]


def parse(name, data):
    if name.endswith(".pdf"):
        return parse_pdf(data)
    if name.endswith(".docx"):
        return parse_docx(data)
    return []


# ───────── INDEX ─────────
@st.cache_resource
def build_index(files):
    from langchain_community.vectorstores import Chroma
    from langchain_community.embeddings import HuggingFaceEmbeddings
    from langchain_text_splitters import RecursiveCharacterTextSplitter

    docs = []
    for name, data in files:
        docs += parse(name, data)

    splitter = RecursiveCharacterTextSplitter(chunk_size=800, chunk_overlap=150)
    chunks = splitter.split_documents(docs)

    emb = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")
    return Chroma.from_documents(chunks, emb)


# ───────── RAG ─────────
def ask(vs, q):
    from langchain_mistralai import ChatMistralAI

    llm = ChatMistralAI(
        model="mistral-small-latest",
        api_key=MISTRAL_KEY
    )

    docs = vs.similarity_search(q, k=5)

    context = "\n\n".join(d.page_content for d in docs)

    prompt = f"""
You are an HSE assistant.

Use only the context below.

{context}

Question: {q}
"""

    return llm.invoke(prompt).content


# ───────── UI ─────────
st.title("🦺 HSE Notifications Assistant")

if "vs" not in st.session_state:
    st.session_state.vs = None

url = st.text_input("Paste SharePoint URL")

if st.button("Fetch & Index"):
    token = get_token()
    files = get_files(token, url)

    st.session_state.vs = build_index(files)
    st.success(f"✅ {len(files)} files indexed")


q = st.text_input("Ask a question")

if st.button("Ask") and st.session_state.vs:
    with st.spinner("Thinking..."):
        answer = ask(st.session_state.vs, q)
    st.write(answer)
