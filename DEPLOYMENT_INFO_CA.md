# Gemini for Microsoft 365 - GCP Deployment Reference

**Environment:** Production / Agentspace  
**GCP Project ID:** `agentspace-452714`  
**GCP Project Number:** `16933400417`  
**Deployment Region:** `us-central1`  
**Deployer Account:** `admin@caugusto.altostrat.com`  

---

## 🚀 Deployed Services & Live Endpoints

### 1. Backend Proxy (`geminiproxy`)
- **Cloud Run Service Name:** `askgemini-proxy`
- **Base URL:** [https://askgemini-proxy-mriilnqopa-uc.a.run.app](https://askgemini-proxy-mriilnqopa-uc.a.run.app)
- **Primary StreamAssist Endpoint:** `https://askgemini-proxy-mriilnqopa-uc.a.run.app/askGeminiEnterprise`
- **Direct Vertex AI Endpoint:** `https://askgemini-proxy-mriilnqopa-uc.a.run.app/askGemini`
- **Health Check:** `https://askgemini-proxy-mriilnqopa-uc.a.run.app/`

#### Environment Variables Configured:
| Variable | Value | Description |
| :--- | :--- | :--- |
| `GCP_PROJECT_ID` | `agentspace-452714` | Target GCP project |
| `GCP_REGION` | `us-central1` | Cloud Run execution region |
| `BACKEND_MODE` | `streamassist` | Enables Gemini Enterprise multi-source grounding |
| `GEMINI_ENTERPRISE_APP_ID` | `test1-agentspace_1741135345115` | Gemini Enterprise App / Engine ID |
| `GCP_LOCATION` | `global` | Discovery Engine collection location |
| `ENTERPRISE_COLLECTION_ID` | `default_collection` | Discovery Engine collection |
| `ENTERPRISE_ASSISTANT_ID` | `default_assistant` | Assistant ID with actions & connectors |

---

### 2. Frontend Add-in (`microsoft-addin`)
- **Cloud Run Service Name:** `gemini-frontend`
- **Base URL:** [https://gemini-frontend-mriilnqopa-uc.a.run.app](https://gemini-frontend-mriilnqopa-uc.a.run.app)
- **Taskpane UI:** `https://gemini-frontend-mriilnqopa-uc.a.run.app/taskpane.html?backend=streamassist`
- **Commands URL:** `https://gemini-frontend-mriilnqopa-uc.a.run.app/commands.html`
- **Assets URL:** `https://gemini-frontend-mriilnqopa-uc.a.run.app/assets/`

---

## 📁 Office 365 Manifest Files

For this environment, a dedicated manifest file was created:
- **Root Manifest File:** [`manifest-ca.xml`](manifest-ca.xml)
- **Add-in Subdirectory Copy:** [`microsoft-addin/manifest-ca.xml`](microsoft-addin/manifest-ca.xml)
- **Original Manifest (Unmodified):** [`microsoft-addin/manifest.xml`](microsoft-addin/manifest.xml)

---

## 🏢 Connected Gemini Enterprise Datastores & Connectors

The engine `test1-agentspace_1741135345115` (`Test1 Agentspace`) grounds responses across:
- 📁 **Google Drive:** `test-company-google-drive_1741135377980`
- ☁️ **Google Cloud Storage:** `switchcraft-gcs-ds_1760473790345_gcs_store`
- 📊 **BigQuery:** `switchcraft-bq-ds_1759766157182_web_data1`, `web_data2`, `mcp-bq_1784644851313_mcp_data`
- 🎫 **ServiceNow:** `servicenow-ds_1749067875905` (Knowledge, Incidents, Catalog, Attachments)
- ✉️ **Gmail Actions:** Send email action
- 📅 **Google Calendar Actions:** Create calendar events
- 🌐 **Web Grounding:** Google Search Grounding enabled

---

## 💻 How to Sideload into Microsoft 365

### Desktop (Word, PowerPoint, Excel on macOS / Windows):
1. Launch **Microsoft Word**, **PowerPoint**, or **Excel**.
2. Navigate to **Insert** > **Add-ins** > **My Add-ins**.
3. Click **Upload My Add-in** (or manage add-ins dropdown).
4. Select [`manifest-ca.xml`](manifest-ca.xml).
5. The **Gemini Enterprise (Agentspace)** button will appear in the **Home** tab ribbon.

### Web (Office 365 on Browser):
1. Open Word / PowerPoint / Excel on [office.com](https://www.office.com).
2. Go to **Insert** > **Add-ins** > **Upload My Add-in**.
3. Select [`manifest-ca.xml`](manifest-ca.xml).

---

## 🛠️ Redeployment & Maintenance Commands

### To Re-deploy Backend Proxy:
```bash
cd geminiproxy
gcloud run deploy askgemini-proxy \
  --source . \
  --region us-central1 \
  --project agentspace-452714 \
  --allow-unauthenticated \
  --set-env-vars GCP_PROJECT_ID=agentspace-452714,GCP_REGION=us-central1,BACKEND_MODE=streamassist,GEMINI_ENTERPRISE_APP_ID=test1-agentspace_1741135345115,GCP_LOCATION=global,ENTERPRISE_COLLECTION_ID=default_collection,ENTERPRISE_ASSISTANT_ID=default_assistant

gcloud run services update askgemini-proxy \
  --region=us-central1 \
  --no-invoker-iam-check \
  --project=agentspace-452714
```

### To Re-deploy Frontend Add-in:
```bash
cd ../microsoft-addin
gcloud run deploy gemini-frontend \
  --source . \
  --region us-central1 \
  --project agentspace-452714 \
  --allow-unauthenticated \
  --set-env-vars GEMINI_PROXY_URL=https://askgemini-proxy-mriilnqopa-uc.a.run.app/askGeminiEnterprise

gcloud run services update gemini-frontend \
  --region=us-central1 \
  --no-invoker-iam-check \
  --project=agentspace-452714
```
