# 🦺 HSE Notifications Assistant

A RAG chatbot for your HSE (Health, Safety & Environment) documents.
Connects directly to an OneDrive / SharePoint shared folder link and lets your team
ask natural-language questions across all safety notices and reports.

---

## Features

| Capability | Detail |
|---|---|
| **OneDrive integration** | Paste a public shared folder link — files are fetched automatically |
| **Manual upload fallback** | Drag-and-drop PDF / DOCX files directly |
| **PDF parsing** | Text, embedded tables (as markdown grids), scanned pages (OCR) |
| **DOCX parsing** | Paragraphs + table cells, structure preserved |
| **Keyword highlighting** | Your search terms are highlighted in every answer |
| **Source citations** | Every fact is attributed to its source document and page |
| **Multi-doc synthesis** | Answers draw from all relevant documents simultaneously |
| **Conversation memory** | Remembers the last 4 Q&A turns for follow-up questions |
| **Follow-up suggestions** | 3 suggested follow-up questions after each answer |

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

> **Tesseract** (for scanned PDF OCR) must be installed separately:
> - **Ubuntu/Debian**: `sudo apt install tesseract-ocr`
> - **macOS**: `brew install tesseract`
> - **Windows**: [Download installer](https://github.com/UB-Mannheim/tesseract/wiki)

### 2. Get a Mistral AI API key

Sign up free at [console.mistral.ai](https://console.mistral.ai), create an API key.

### 3. Configure secrets (optional — recommended for deployment)

Create `.streamlit/secrets.toml`:

```toml
MISTRAL_API_KEY = "your-mistral-key-here"
```

Or set as environment variable: `export MISTRAL_API_KEY=your-key`

### 4. Run the app

```bash
streamlit run app.py
```

---

## OneDrive / SharePoint Setup

1. Navigate to your HSE documents folder in OneDrive or SharePoint
2. Click **Share** → set permissions to **"Anyone with the link can view"**
3. Click **Copy link**
4. Paste the link into the sidebar of the app → click **Fetch & Index Files**

The app will download all PDF and DOCX files from the folder and build a searchable index.

### Re-syncing documents

Click **Fetch & Index Files** again at any time to pull the latest files from OneDrive.
The index is rebuilt fresh each time.

---

## Deploy to Streamlit Community Cloud

1. Push this folder to a GitHub repository
2. Go to [share.streamlit.io](https://share.streamlit.io) → New app
3. Select your repo and set `app.py` as the entry point
4. Under **Advanced settings → Secrets**, add:
   ```toml
   MISTRAL_API_KEY = "your-mistral-key-here"
   ```
5. Deploy!

---

## Example Questions

- *What PPE is required for chemical handling?*
- *List all emergency evacuation steps*
- *What does the hot work permit require?*
- *What inspection items are in the safety checklist?*
- *Compare PPE requirements across all notices*
- *What are the incident reporting procedures?*

---

## Architecture

```
OneDrive Share Link
        │
        ▼
  Graph API (anonymous shares endpoint)
        │   downloads PDF / DOCX bytes
        ▼
  Document Parsers
  ├── pdfplumber  → text + tables
  ├── pdf2image + pytesseract  → OCR for scanned pages
  └── python-docx  → paragraphs + tables
        │
        ▼
  RecursiveCharacterTextSplitter  (1000 tokens, 200 overlap)
        │
        ▼
  HuggingFace all-MiniLM-L6-v2  (embeddings, CPU)
        │
        ▼
  ChromaDB  (in-memory vector store)
        │
        ▼  MMR retrieval + per-source gap fill
  Mistral AI  (LLM answer generation)
        │
        ▼
  Streamlit UI  (keyword highlight, source citations, follow-ups)
```
