# Gemini for Microsoft Office 365 (Word, PowerPoint, Excel)
**Author:** Sathya AG, Principal Architect, Google  
**License:** Apache-2.0  

An enterprise-ready Microsoft 365 Add-in and Google Cloud backend that integrates **Google Cloud Vertex AI (Gemini 2.5 Flash, Vertex AI Search Grounding, and Gemini 2.5 Flash Image / Nano Banana)** directly into **Microsoft Word**, **PowerPoint**, and **Excel**.

## 📽️ Demo Video


https://github.com/user-attachments/assets/7b361e71-13d4-48ba-81ce-41bf5d0c9e50



---

## 🌟 Key Capabilities

- **🏢 Enterprise Grounding & Zero Hallucination:** Directly grounded via **Vertex AI Search Datastores** over enterprise quarterly reports, policies, 10-K/10-Q SEC filings, and research documents in Google Cloud Storage.
- **🎨 Multimodal Visual Generation:** Generates high-resolution flat 2D vector charts and infographics via `gemini-2.5-flash-image` and natively embeds them into presentations and documents.
- **📄 Microsoft Word (`WordAdapter`):** Document summarization, risk analysis, text rewriting, and in-document `@gemini <prompt>` execution.
- **📊 Microsoft PowerPoint (`PPTAdapter`):** Automatic multi-slide executive deck generator with widescreen layouts, data tables, and native macOS WKWebView dual-pipeline visual injection.
- **📈 Microsoft Excel (`ExcelAdapter`):** Worksheet intelligence, cell range risk analysis, formula anomaly detection, and KPI metric cards.

---

> 📖 **Developer & Architecture Guide:** For a detailed breakdown of the new decoupled authentication architecture, Microsoft Entra ID SSO integration, Google Cloud S2S IAM security, and flow diagrams, see [DEVELOPER_ARCHITECTURE_GUIDE.md](file:///Users/caugusto/Documents/antigravity/retail-gemini-for-office-365/authproxy/DEVELOPER_ARCHITECTURE_GUIDE.md).

## 🏗️ Architecture Overview


```mermaid
graph TB
    subgraph ClientTier ["Microsoft 365 Client Tier (Desktop & Web)"]
        WordApp["Microsoft Word<br/>(WordAdapter)"]
        PPTApp["Microsoft PowerPoint<br/>(PPTAdapter)"]
        ExcelApp["Microsoft Excel<br/>(ExcelAdapter)"]
        TaskpaneUI["Add-in Taskpane Webview<br/>(Vanilla JS + Fluent UI CSS)"]
        
        WordApp <--> TaskpaneUI
        PPTApp <--> TaskpaneUI
        ExcelApp <--> TaskpaneUI
    end

    subgraph GCPInfrastructure ["Google Cloud Platform (genai-demo-catalog / us-central1)"]
        FrontendRun["Cloud Run: gemini-frontend<br/>(Nginx Container / Port 80)"]
        ProxyFunction["Cloud Function Gen 2: askGemini<br/>(Node.js 20 Microservice)"]
        
        TaskpaneUI -->|1. Loads Static Assets & JS Bundle| FrontendRun
        TaskpaneUI -->|2. HTTPS POST /askGemini JSON| ProxyFunction
    end

    subgraph VertexAIEngine ["Google Cloud Vertex AI Enterprise Backend"]
        SearchDS["Vertex AI Search Datastore<br/>(Enterprise 10-K/10-Q Docs in GCS)"]
        FlashModel["Gemini 2.5 Flash<br/>(Grounded Generative Text Model)"]
        ImageModel["Gemini 2.5 Flash Image<br/>(Nano Banana Visual Chart Generator)"]
        
        ProxyFunction -->|A. Grounded RAG Query| FlashModel
        FlashModel <-->|B. Semantic Retrieval & Citations| SearchDS
        ProxyFunction -->|C. Balanced Regex Extractor| ImageModel
    end

    subgraph OutputPipeline ["Client Rendering & Document Injection"]
        ProxyFunction -->|3. Grounded Text + Base64 PNGs| TaskpaneUI
        TaskpaneUI -->|Word.run / OOXML| WordApp
        TaskpaneUI -->|PowerPoint.run + setSelectedDataAsync| PPTApp
        TaskpaneUI -->|Excel.run| ExcelApp
    end

    style ClientTier fill:#e8f0fe,stroke:#1a73e8,stroke-width:2px;
    style GCPInfrastructure fill:#e6f4ea,stroke:#137333,stroke-width:2px;
    style VertexAIEngine fill:#fef7e0,stroke:#f9ab00,stroke-width:2px;
    style OutputPipeline fill:#fce8e6,stroke:#c5221f,stroke-width:2px;
```

> 📖 For full architectural deep dives, sequence diagrams, and macOS WKWebView rendering pipelines, see [ARCHITECTURE.md](ARCHITECTURE.md).

---

## 📂 Repository Structure

```
gemini-for-office-365/
│
├── README.md                                   # Main project overview & quick-start guide
├── ARCHITECTURE.md                             # Detailed architecture, grounding, visual & auth diagrams
├── DEPLOYMENT_INFO_CA.md                       # Active deployment environment & configuration reference
├── MICROSOFT_365_ADMIN_CENTER_DEPLOYMENT.md    # Microsoft 365 Admin Center enterprise distribution guide
├── manifest-ca.xml                             # Single canonical Office 365 Add-in XML Manifest
├── LICENSE                                     # Apache-2.0 License
├── .gitignore                                  # Root gitignore
│
├── authproxy/                                  # Tier 2: Microsoft Entra ID Auth Gateway (Python FastAPI)
│   ├── main.py                                 # JWT verification, IdP auto-discovery & S2S token minting
│   ├── requirements.txt                        # FastAPI, uvicorn, PyJWT, cryptography, google-auth
│   ├── Dockerfile                              # Python 3.11 Cloud Run container
│   ├── DEPLOYMENT_AND_ENTRA_GUIDE.md           # Step-by-step Entra ID & Cloud Run deployment guide
│   └── DEVELOPER_ARCHITECTURE_GUIDE.md         # Deep-dive architecture, sequence flows & developer guide
│
├── microsoft-addin/                            # Tier 1: Microsoft Office 365 Add-in (Word, PPT, Excel)
│   ├── package.json                            # Webpack, Babel, Office.js dependencies
│   ├── webpack.config.js                       # Production build & bundling configuration
│   ├── Dockerfile                              # Nginx web server for Cloud Run hosting
│   ├── nginx.conf                              # Nginx security & CORS headers
│   ├── assets/                                 # Office Ribbon and Taskpane branding icons
│   └── src/
│       ├── adapters/                           # WordAdapter, PPTAdapter, ExcelAdapter, HostAdapterFactory
│       ├── core/                               # authService (Office SSO), geminiClient, markdownParser
│       ├── taskpane/                           # taskpane.html, taskpane.css, taskpane.js
│       └── commands/                           # commands.html, commands.js
│
├── geminiproxy/                                # Tier 3: Core Inference Backend (Node.js Express)
│   ├── index.js                                # Discovery Engine streamAssist, Gemini 2.5 Flash, Grounding
│   ├── package.json                            # Dependencies (express, cors, google-auth-library)
│   ├── Dockerfile                              # Node.js 20 Cloud Run container
│   └── README.md                               # Backend configuration guide
│
└── scripts/
    └── sideload_mac.sh                         # macOS local development sideloading automation script
```

---

## 🚀 Quick Start & Deployment

For complete, detailed instructions on setting up Microsoft Entra ID, Google Cloud IAM, and deploying the microservices, consult:
- 📖 [DEVELOPER_ARCHITECTURE_GUIDE.md](authproxy/DEVELOPER_ARCHITECTURE_GUIDE.md)
- 🚀 [DEPLOYMENT_AND_ENTRA_GUIDE.md](authproxy/DEPLOYMENT_AND_ENTRA_GUIDE.md)
- 🏢 [MICROSOFT_365_ADMIN_CENTER_DEPLOYMENT.md](MICROSOFT_365_ADMIN_CENTER_DEPLOYMENT.md)
- 📋 [DEPLOYMENT_INFO_CA.md](DEPLOYMENT_INFO_CA.md)

### 1. Deploy Auth Gateway Proxy (`authproxy/`)
```bash
cd authproxy
gcloud run deploy auth-proxy \
  --source . \
  --project YOUR_GCP_PROJECT_ID \
  --region us-central1 \
  --allow-unauthenticated \
  --service-account auth-proxy-sa@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com \
  --set-env-vars "\
MICROSOFT_ENTRA_APP_ID=YOUR_MICROSOFT_ENTRA_CLIENT_ID,\
DOWNSTREAM_BACKEND_URL=https://askgemini-proxy-XXXXXXXX.us-central1.run.app,\
GCP_PROJECT_ID=YOUR_GCP_PROJECT_ID,\
GCP_LOCATION=global,\
USER_AUTH_MODE=auto,\
REQUIRE_ENTRA_AUTH=true,\
VERBOSE_LOGGING=true"
```

### 2. Deploy Backend Proxy (`geminiproxy/`)
```bash
cd ../geminiproxy
gcloud run deploy askgemini-proxy \
  --source . \
  --project YOUR_GCP_PROJECT_ID \
  --region us-central1 \
  --no-allow-unauthenticated \
  --set-env-vars "\
GCP_PROJECT_ID=YOUR_GCP_PROJECT_ID,\
GEMINI_ENTERPRISE_APP_ID=YOUR_GEMINI_ENTERPRISE_APP_ID,\
BACKEND_MODE=streamassist,\
GCP_LOCATION=global,\
ENTERPRISE_COLLECTION_ID=default_collection,\
ENTERPRISE_ASSISTANT_ID=default_assistant,\
ALLOW_SERVICE_ACCOUNT_FALLBACK=true"
```

### 3. Deploy Frontend Add-in (`microsoft-addin/`)
```bash
cd ../microsoft-addin
npm install
npm run build

gcloud run deploy gemini-frontend \
  --source . \
  --region us-central1 \
  --project YOUR_GCP_PROJECT_ID \
  --allow-unauthenticated
```

### 4. Sideload into Microsoft Office 365
- **macOS Quick Sideload**: Run `./scripts/sideload_mac.sh` and restart Word, PowerPoint, or Excel.
- **Office for Web**: Open document on [office.com](https://www.office.com), navigate to **Insert** > **Add-ins** > **Upload My Add-in**, and select `manifest-ca.xml`.
- **Microsoft 365 Admin Center**: Upload `manifest-ca.xml` under **Settings** > **Integrated apps** for tenant-wide deployment.


---

## 🔒 Security & Privacy
- **Zero API Keys in Client:** All client-side calls route through authenticated Google Cloud microservices.
- **Enterprise Isolation:** Documents, slides, and spreadsheet data stay strictly inside your Google Cloud tenant boundary.
- **CORS Configured:** Permissive for enterprise Office webviews while protecting internal execution pipelines.
