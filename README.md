# SharePoint PDF Document Search — Streamlit App

A dark-themed, production-ready Streamlit app that scans a SharePoint document library for PDFs matching event numbers, description keywords, or free text.

---

## Features

- **Three search modes**: event numbers, keywords, or free text (or all combined)
- **Live SharePoint connection** via Azure AD App-only credentials
- **Download results** as Excel or CSV
- **Direct links** to matched files in SharePoint
- Match badge display showing exactly which terms were found

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Register an Azure AD App (one-time admin setup)

1. Go to **Azure Portal → App registrations → New registration**
2. Name it (e.g. `SharePoint Search App`), leave redirect URI blank
3. Go to **API Permissions → Add a permission → SharePoint → Application permissions**
4. Add `Sites.Read.All` → Grant admin consent
5. Go to **Certificates & Secrets → New client secret** — copy the value immediately
6. Note your **Application (client) ID** from the Overview page

### 3. Grant the app access to your SharePoint site

Option A — SharePoint Admin Center:
- Site settings → Site permissions → Share with the app

Option B — PowerShell:
```powershell
Connect-PnPOnline -Url https://contoso.sharepoint.com/sites/YourSite -Interactive
Grant-PnPAzureADAppSitePermission -AppId <your-client-id> -DisplayName "Search App" -Site <site-url> -Permissions Read
```

### 4. Run the app

```bash
streamlit run app.py
```

Then open http://localhost:8501 and enter your:
- **Site URL** — e.g. `https://contoso.sharepoint.com/sites/Notifications`
- **Client ID** — from Azure AD app registration
- **Client Secret** — from Azure AD app secret
- **Library Name** — the document library name (default: `Documents`)

---

## Deploying publicly (Streamlit Community Cloud)

1. Push this folder to a GitHub repo
2. Go to [share.streamlit.io](https://share.streamlit.io) and connect your repo
3. Set secrets in the Streamlit Cloud dashboard (optional — avoids entering credentials each time):

```toml
# .streamlit/secrets.toml (local dev only, never commit this)
SHAREPOINT_SITE_URL = "https://contoso.sharepoint.com/sites/YourSite"
SHAREPOINT_CLIENT_ID = "your-client-id"
SHAREPOINT_CLIENT_SECRET = "your-client-secret"
SHAREPOINT_LIBRARY = "Documents"
```

Then in `app.py`, pre-fill the sidebar inputs using `st.secrets` if available.

---

## File Structure

```
sharepoint_search_app/
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
└── README.md           # This file
```
