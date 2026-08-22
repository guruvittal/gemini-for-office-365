# Gemini for Microsoft 365 - GCP Deployment Reference

**Environment:** Production / Agentspace  
**GCP Project ID:** `agentspace-452714`  
**GCP Project Number:** `16933400417`  
**Deployment Region:** `us-central1`  
**Deployer Account:** `admin@caugusto.altostrat.com`  

---

## 🚀 Deployed Services & Live Endpoints

### 1. Auth Gateway Proxy (`authproxy`)
- **Cloud Run Service Name:** `auth-proxy`
- **Base URL:** [https://auth-proxy-16933400417.us-central1.run.app](https://auth-proxy-16933400417.us-central1.run.app)
- **Health Endpoint:** `https://auth-proxy-16933400417.us-central1.run.app/health`
- **Interactive Swagger Docs:** `https://auth-proxy-16933400417.us-central1.run.app/docs`
- **Office 365 Add-in Endpoint:** `https://auth-proxy-16933400417.us-central1.run.app/askGeminiEnterprise`
- **Dedicated Service Account:** `gemini-office365-sa@agentspace-452714.iam.gserviceaccount.com`
- **Assigned IAM Roles:**
  - `roles/logging.logWriter` on Project `agentspace-452714` (Cloud Logging)
  - `roles/run.invoker` on Project `agentspace-452714` (Cloud Run invocation)
  - `roles/discoveryengine.editor` on Project `agentspace-wif` (Cross-project Discovery Engine access & ACL inspection)
- **Entra ID App ID (WIF Client ID):** `85fb5428-6249-4131-9eeb-f2436d5d4d8c`
- **Application ID URI:** `api://gemini-frontend-16933400417.us-central1.run.app/85fb5428-6249-4131-9eeb-f2436d5d4d8c`
- **Authorized Client Applications:**
  - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Office on the Web)
  - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Office on the Web)
  - `00000002-0000-0ff1-ce00-000000000000` (Office Desktop / Mac / Windows)
- **Downstream Backend:** `https://askgemini-proxy-16933400417.us-central1.run.app` (Invoked via Google S2S IAM token)
- **User Auth Mode:** `auto` (`USER_AUTH_MODE=auto` - dynamically queries Discovery Engine `aclConfig` to detect `GSUITE` vs `THIRD_PARTY` WIF)
- **Logging Mode:** Native GCP Structured JSON (`VERBOSE_LOGGING=true`)

---

### 2. Backend Proxy (`geminiproxy`)
- **Cloud Run Service Name:** `askgemini-proxy`
- **Base URL:** [https://askgemini-proxy-16933400417.us-central1.run.app](https://askgemini-proxy-16933400417.us-central1.run.app)
- **Dedicated Service Account:** `gemini-office365-sa@agentspace-452714.iam.gserviceaccount.com`
- **Security Posture:** Locked down with `--no-allow-unauthenticated` (Private S2S only; invokable by `gemini-office365-sa`)
- **Primary StreamAssist Endpoint:** `https://askgemini-proxy-16933400417.us-central1.run.app/askGeminiEnterprise`
- **Dynamic Grounding:** Automatically queries all datastores attached to the target Engine using the `toolsSpec.vertexAiSearchSpec: {}` payload wildcard workaround.
- **Direct Vertex AI Endpoint:** `https://askgemini-proxy-16933400417.us-central1.run.app/askGemini`
- **Health Check:** `https://askgemini-proxy-16933400417.us-central1.run.app/`

#### Environment Variables Configured:
| Variable | Value | Description |
| :--- | :--- | :--- |
| `GCP_PROJECT_ID` | `agentspace-wif` | Target GCP project hosting Gemini Enterprise instance |
| `GCP_REGION` | `us-central1` | Cloud Run execution region |
| `BACKEND_MODE` | `streamassist` | Enables Gemini Enterprise multi-source grounding |
| `GEMINI_ENTERPRISE_APP_ID` | `instance-demos1_1774616568648` | Gemini Enterprise App / Engine ID in `agentspace-wif` |
| `GCP_LOCATION` | `global` | Discovery Engine collection location |
| `ENTERPRISE_COLLECTION_ID` | `default_collection` | Discovery Engine collection |
| `ENTERPRISE_ASSISTANT_ID` | `default_assistant` | Assistant ID with actions & connectors |
| `ALLOW_SERVICE_ACCOUNT_FALLBACK` | `true` | Allows fallback when user token is absent while emitting structured warning logs |

---

### 3. Frontend Add-in (`microsoft-addin`)
- **Cloud Run Service Name:** `gemini-frontend`
- **Base URL:** [https://gemini-frontend-16933400417.us-central1.run.app](https://gemini-frontend-16933400417.us-central1.run.app)
- **Taskpane UI:** `https://gemini-frontend-16933400417.us-central1.run.app/taskpane.html?backend=streamassist`
- **Commands URL:** `https://gemini-frontend-16933400417.us-central1.run.app/commands.html`
- **Assets URL:** `https://gemini-frontend-16933400417.us-central1.run.app/assets/`
- **Hosted Manifest:** `https://gemini-frontend-16933400417.us-central1.run.app/manifest-ca.xml`

---

## 🔑 Identity Provider Auto-Discovery & Token Pass-Through Architecture

`auth-proxy` dynamically queries the Discovery Engine Access Control List configuration:
`GET https://{location}-discoveryengine.googleapis.com/v1/projects/{project_id}/locations/{location}/aclConfig`

1. **Google Identity / Cloud Identity Mode (`GSUITE`)**:
   - Discovered on project `agentspace-452714`.
   - Passes user token if provided or forwards verified identity context (`X-End-User-*`) to Discovery Engine.
2. **Workforce Identity Federation Mode (`THIRD_PARTY`)**:
   - Discovered on project `agentspace-wif` (`workforcePoolName: locations/global/workforcePools/ca-entra-id-oidc-pool`).
   - Dynamically exchanges the Microsoft Entra ID JWT with Google STS (`https://sts.googleapis.com/v1/token`) for a federated Google OAuth access token (`ya29.s...`) and passes it in `Authorization: Bearer <token>` to Discovery Engine `streamAssist`.
   - **Configured WIF Pool Providers**:
     - `entra-id-oidc-pool-provider`: Client ID `85fb5428-6249-4131-9eeb-f2436d5d4d8c` (General web SSO audience)
   - **Finding the WIF Client ID**:
     - You do **not** need to create a new App Registration. Modify the existing WIF App Registration.
     - Find the ID using the GCP Console: Go to **IAM & Admin** > **Workforce Identity Federation** > Select your pool > Select your provider > Find the **Client ID** under OIDC Settings.
     - Or using `gcloud`: `gcloud iam workforce-pools providers describe entra-id-oidc-pool-provider --workforce-pool="ca-entra-id-oidc-pool" --location="global" --format="value(oidc.clientId)"`
   - **Entra ID Manifest Configuration**:
     - Configure the WIF App Registration to trust the Office Add-in.
     - Application ID URI must match the `gemini-frontend` domain (e.g., `api://gemini-frontend-16933400417.us-central1.run.app/85fb5428-6249-4131-9eeb-f2436d5d4d8c`).
     - `"requestedAccessTokenVersion": 2` under `"api"` section in the Entra ID App Registration is required so Microsoft Entra ID emits v2.0 tokens matching Google STS requirements.
     - Add `email`, `upn`, and `preferred_username` as optional claims to the Access Token so Google WIF can map `google.subject`.
3. **Service Account Fallback & Licensing Enforcement**:
   - If an end user is unlicensed, Discovery Engine returns `HTTP 403 Permission Denied`.
   - If `ALLOW_SERVICE_ACCOUNT_FALLBACK=true`, `askgemini-proxy` logs a structured warning and uses Service Account credentials when user token is not supplied.

---

## 📁 Office 365 Manifest File

The repository maintains a single canonical manifest file:
- **Canonical Manifest File:** [`manifest-ca.xml`](manifest-ca.xml) (Repository Root)
- **Hosted Live Manifest URL:** `https://gemini-frontend-16933400417.us-central1.run.app/manifest-ca.xml`

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

---

## 💻 How to Sideload & Deploy to Microsoft 365

### Method 1: Automated macOS Local Sideloading (Recommended for Development / Testing)
Run the automated script to deploy [`manifest-ca.xml`](manifest-ca.xml) directly into the WEF cache for Word, PowerPoint, and Excel:
```bash
./scripts/sideload_mac.sh
```
1. Restart Microsoft Word, PowerPoint, or Excel.
2. Open a document/presentation/spreadsheet.
3. The **Gemini Assistant** button will be visible in the **Home** tab ribbon.
4. If not immediately visible, navigate to **Insert** > **Add-ins** > **My Add-ins** > **Developer Add-ins** and select **Gemini Enterprise (Agentspace)**.

---

### Method 2: Office on the Web (Word, PowerPoint, Excel in Browser)
1. Navigate to [office.com](https://www.office.com) or [onedrive.live.com](https://onedrive.live.com) and open a document in Word, PowerPoint, or Excel for Web.
2. Click **Insert** tab > **Add-ins** button (or `...` overflow menu > **Add-ins**).
3. In the Office Add-ins dialog, click **Manage My Add-ins** (or **My Add-ins** link in top right) > **Upload My Add-in**.
4. Browse and select [`manifest-ca.xml`](manifest-ca.xml).
5. The add-in loads immediately into the browser taskpane with full Entra ID SSO enabled.

---

### Method 3: Centralized Deployment for Organization (Microsoft 365 Admin Center)
> 📖 **Full Enterprise Deployment Guide:** See [MICROSOFT_365_ADMIN_CENTER_DEPLOYMENT.md](MICROSOFT_365_ADMIN_CENTER_DEPLOYMENT.md) for the complete walkthrough with permission models, Entra ID tenant consent, user assignment groups, lifecycle management, and troubleshooting.

To roll out the Gemini Enterprise add-in to all employees or specific security groups:
1. Sign in to the [Microsoft 365 Admin Center](https://admin.microsoft.com/) as a **Global Administrator** or **Exchange/App Administrator**.
2. Navigate to **Settings** > **Integrated apps** (or **Settings** > **Add-ins**).
3. Click **Upload custom apps** (or **Deploy Add-in**).
4. Select **Office Add-in** and choose **Provide link to manifest file**:
   ```text
   https://gemini-frontend-16933400417.us-central1.run.app/manifest-ca.xml
   ```
   *(or upload local file [`manifest-ca.xml`](manifest-ca.xml))*.
5. Set user assignment:
   - **Entire Organization**, **Specific users/groups** (e.g., "Gemini Enterprise Pilots"), or **Just me** (for piloting).
6. Click **Next** > **Finish deployment**. The add-in automatically appears in the Office ribbon across all desktop and web apps.

---

### Method 4: Windows Desktop Sideloading (Shared Folder Catalog)
1. Share a local folder on your network (e.g. `\\localhost\OfficeManifests`) with Read permissions.
2. Place [`manifest-ca.xml`](manifest-ca.xml) in that shared folder.
3. In Word/PowerPoint/Excel: Go to **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.
4. Enter the Catalog URL (`\\localhost\OfficeManifests`), click **Add catalog**, and check **Show in Menu**.
5. Restart Office. Go to **Insert** > **Add-ins** > **Shared Folder** > click **Gemini Enterprise (Agentspace)**.

---

## 🛠️ Redeployment & Maintenance Commands

### To Re-deploy Auth Proxy:
```bash
cd authproxy
gcloud run deploy auth-proxy \
  --source . \
  --project agentspace-452714 \
  --region us-central1 \
  --service-account auth-proxy-sa@agentspace-452714.iam.gserviceaccount.com \
  --allow-unauthenticated \
  --set-env-vars "\
MICROSOFT_ENTRA_APP_ID=85fb5428-6249-4131-9eeb-f2436d5d4d8c,\
DOWNSTREAM_BACKEND_URL=https://askgemini-proxy-16933400417.us-central1.run.app,\
GCP_PROJECT_ID=agentspace-wif,\
GCP_LOCATION=global,\
USER_AUTH_MODE=auto,\
REQUIRE_ENTRA_AUTH=true,\
VERBOSE_LOGGING=true"
```

### To Re-deploy Backend Proxy:
```bash
cd geminiproxy
gcloud run deploy askgemini-proxy \
  --source . \
  --project agentspace-452714 \
  --region us-central1 \
  --no-allow-unauthenticated \
  --set-env-vars "\
GCP_PROJECT_ID=agentspace-452714,\
GEMINI_ENTERPRISE_APP_ID=test1-agentspace_1741135345115,\
BACKEND_MODE=streamassist,\
GCP_LOCATION=global,\
ENTERPRISE_COLLECTION_ID=default_collection,\
ENTERPRISE_ASSISTANT_ID=default_assistant,\
ALLOW_SERVICE_ACCOUNT_FALLBACK=true"
```

### To Re-deploy Frontend Add-in:
```bash
cd ../microsoft-addin
gcloud run deploy gemini-frontend \
  --source . \
  --region us-central1 \
  --project agentspace-452714 \
  --allow-unauthenticated \
  --set-env-vars GEMINI_PROXY_URL=https://auth-proxy-16933400417.us-central1.run.app/askGeminiEnterprise
```
