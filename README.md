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
        
        TaskpaneUI -->|"1. Loads Static Assets & JS Bundle"| FrontendRun
        TaskpaneUI -->|"2. HTTPS POST /askGemini JSON"| ProxyFunction
    end

    subgraph VertexAIEngine ["Google Cloud Vertex AI Enterprise Backend"]
        SearchDS["Vertex AI Search Datastore<br/>(Enterprise 10-K/10-Q Docs in GCS)"]
        FlashModel["Gemini 2.5 Flash<br/>(Grounded Generative Text Model)"]
        ImageModel["Gemini 2.5 Flash Image<br/>(Nano Banana Visual Chart Generator)"]
        
        ProxyFunction -->|"A. Grounded RAG Query"| FlashModel
        FlashModel ---|"B. Semantic Retrieval & Citations"| SearchDS
        ProxyFunction -->|"C. Balanced Regex Extractor"| ImageModel
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
├── manifest-wif.xml                            # Office 365 Add-in XML Manifest for WIF
├── manifest-gsuite.xml                         # Office 365 Add-in XML Manifest for GSuite / Cloud Identity
├── manifest.xml                                # Legacy / Base Add-in XML Manifest
├── LICENSE                                     # Apache-2.0 License
├── .gitignore                                  # Root gitignore
│
├── authproxy/                                  # Tier 2: Microsoft Entra ID Auth Gateway (Python FastAPI)
│   ├── main.py                                 # JWT verification, IdP auto-discovery & S2S token minting
│   ├── requirements.txt                        # FastAPI, uvicorn, PyJWT, cryptography, google-auth
│   ├── Dockerfile                              # Python 3.11 Cloud Run container
│   ├── README.md                               # Auth proxy service reference
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
    ├── generate_manifest.py                    # Interactive & CLI tool to generate custom Office XML manifests
    └── sideload_mac.sh                         # macOS local development sideloading automation script
```

---

## 📚 Documentation & Deployment Guides

For setup instructions, deployment steps, architecture deep-dives, and admin guides, consult the dedicated documentation files:

| Guide | Description |
| :--- | :--- |
| 📋 **[`DEPLOYMENT_INSTRUCTIONS.md`](DEPLOYMENT_INSTRUCTIONS.md)** | **Primary Deployment Runbook:** End-to-end first-time setup for **Track 1 (WIF)** and **Track 2 (GSuite)**, Microsoft Entra ID App Registration, live environment configuration, dual security boundary explanation, Google OAuth client setup, manifest customization reference, and full Cloud Run environment variables catalog. |
| 📖 **[`authproxy/DEVELOPER_ARCHITECTURE_GUIDE.md`](authproxy/DEVELOPER_ARCHITECTURE_GUIDE.md)** | **Architecture & Security Deep-Dive:** Token exchange flows (Entra ID JWT ➔ Google STS Workforce Pool ➔ Gemini Enterprise), Service-to-Service IAM authentication, and comprehensive error resolution matrix. |
| 🏢 **[`MICROSOFT_365_ADMIN_CENTER_DEPLOYMENT.md`](MICROSOFT_365_ADMIN_CENTER_DEPLOYMENT.md)** | **Centralized IT Admin Deployment:** Enterprise-wide rollout guide via Microsoft 365 Admin Center Integrated Apps. |
| 🏗️ **[`ARCHITECTURE.md`](ARCHITECTURE.md)** | **System Architecture:** Detailed client adapter lifecycle (`WordAdapter`, `PPTAdapter`, `ExcelAdapter`), multimodal visual generation pipeline, and document injection flows. |
| ⚙️ **[`geminiproxy/README.md`](geminiproxy/README.md)** | **Backend Proxy Reference:** Configuration, environment variables, and deployment for the Node.js Express inference backend. |
| 💻 **[`microsoft-addin/README.md`](microsoft-addin/README.md)** | **Office Add-in Frontend Reference:** Build, local development server, Nginx container packaging, and manifest sideloading. |

---

## 🔒 Security & Privacy

- **Zero API Keys in Client:** All client-side calls route through authenticated Google Cloud microservices using Microsoft Entra ID SSO tokens or Google OAuth.
- **Enterprise Isolation:** Documents, slides, and spreadsheet data stay strictly inside your Google Cloud tenant boundary.
- **Principle of Least Privilege:** Fine-grained IAM service accounts isolate gateway authentication from backend Gemini Enterprise execution.

