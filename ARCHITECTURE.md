# Architecture & System Design: Gemini for Microsoft 365
**Author:** Sathya AG, Principal Architect, Google  
**License:** Apache-2.0  

---

## 📑 Executive Summary

**Gemini for Microsoft 365** is an enterprise-grade AI integration platform that connects Microsoft Office client applications (**Word**, **PowerPoint**, and **Excel**) with **Google Cloud Vertex AI**.

The system combines:
1. **Factual Enterprise Grounding** via **Vertex AI Search Datastores** (RAG over enterprise 10-K/10-Q filings, policies, and internal docs).
2. **Generative Intelligence** via **Gemini 2.5 Flash** for multi-turn conversational reasoning and content synthesis.
3. **Multimodal Visual Synthesis** via **Gemini 2.5 Flash Image** (Nano Banana) for dynamic 2D flat vector financial charts and infographics.
4. **Native Host Adapters** for Microsoft Word (inline `@gemini` insertion, text transformation), PowerPoint (multi-slide executive decks with native macOS WKWebView dual-pipeline image rendering), and Excel (range analysis, anomaly detection, KPI metric cards).

> 📖 **Decoupled Auth & S2S System Architecture:** For a deep dive into the decoupled Microsoft Entra ID authentication gateway, Google Cloud Service-to-Service IAM model, and implementation blueprints, see [DEVELOPER_ARCHITECTURE_GUIDE.md](file:///Users/caugusto/Documents/antigravity/retail-gemini-for-office-365/authproxy/DEVELOPER_ARCHITECTURE_GUIDE.md) and [DEPLOYMENT_AND_ENTRA_GUIDE.md](file:///Users/caugusto/Documents/antigravity/retail-gemini-for-office-365/authproxy/DEPLOYMENT_AND_ENTRA_GUIDE.md).

---

## 🏛️ Enterprise Multi-Tier Architecture (Milestones 2 & 3)

The enterprise deployment separates concerns into dedicated, independently scalable, least-privilege tiers across **Microsoft 365**, **Microsoft Entra ID**, and **Google Cloud Platform**:

```mermaid
graph TB
    subgraph OfficeClientTier ["1. Microsoft 365 Office Client Tier"]
        WordApp["Microsoft Word<br/>(Desktop & Web)"]
        PPTApp["Microsoft PowerPoint<br/>(Desktop & Web)"]
        ExcelApp["Microsoft Excel<br/>(Desktop & Web)"]
        OfficeSSO["Office.js SSO Runtime<br/>(Office.auth.getAccessToken)"]
        TaskpaneUI["Add-in Taskpane UI<br/>(HTML5 / Vanilla JS / CSS)"]
        
        WordApp <--> TaskpaneUI
        PPTApp <--> TaskpaneUI
        ExcelApp <--> TaskpaneUI
        TaskpaneUI <--> OfficeSSO
    end

    subgraph IdentityTier ["2. Microsoft Entra ID Identity Tier"]
        EntraApp["Entra ID App Registration<br/>(App ID: b990d644-...)"]
        AppUri["Application ID URI<br/>api://gemini-frontend-...run.app/b990d644-..."]
        PreAuthApps["Pre-Authorized Office Clients<br/>• Web: ea5a67f6..., d3590ed6...<br/>• Desktop: 00000002-..."]
        
        OfficeSSO -->|1. Silent SSO Token Request| EntraApp
        EntraApp -->|2. Signed Microsoft JWT| OfficeSSO
    end

    subgraph GCPInfrastructure ["3. Google Cloud Platform (Project: agentspace-452714)"]
        FrontendRun["gemini-frontend (Cloud Run)<br/>Nginx Static Assets & Taskpane<br/>(Public HTTPS)"]
        
        subgraph AuthGatewayTier ["Auth & Token Translation Gateway"]
            AuthProxy["auth-proxy (Cloud Run)<br/>Python 3.11 / FastAPI<br/>Runtime SA: auth-proxy-sa"]
            IdpDiscovery["Dynamic IdP Auto-Discovery<br/>GET /v1/.../aclConfig"]
        end
        
        subgraph PrivateBackendTier ["Private Core Inference Backend (Private S2S Only)"]
            BackendProxy["askgemini-proxy (Cloud Run)<br/>Node.js 20 Express Microservice<br/>(--no-allow-unauthenticated)"]
        end
        
        TaskpaneUI -->|Loads Static Assets| FrontendRun
        TaskpaneUI -->|3. POST /askGeminiEnterprise + Entra Bearer JWT| AuthProxy
        AuthProxy <-->|4. Discovers IdP Type (GSUITE vs THIRD_PARTY)| IdpDiscovery
        AuthProxy -->|5. Forward Request + Google S2S IAM Token + User Claims| BackendProxy
    end

    subgraph GeminiEnterpriseTier ["4. Google Cloud Discovery Engine / Gemini Enterprise"]
        StreamAssist["Discovery Engine API<br/>POST /v1/.../engines/test1-agentspace/servingConfigs/default_search:streamAssist"]
        EnterpriseCorpus["Enterprise Grounding Corpus<br/>(Google Drive, GCS, Spanner, BigQuery)"]
        
        BackendProxy -->|6. StreamAssist Request with User Context| StreamAssist
        StreamAssist <-->|7. Semantic Grounding & Chunk Retrieval| EnterpriseCorpus
        StreamAssist -->>|8. Server-Sent Events (SSE) Stream| BackendProxy
        BackendProxy -->>|9. SSE Chunks| AuthProxy
        AuthProxy -->>|10. Stream to Client| TaskpaneUI
    end

    style OfficeClientTier fill:#e8f0fe,stroke:#1a73e8,stroke-width:2px;
    style IdentityTier fill:#f3e8fd,stroke:#7b1fa2,stroke-width:2px;
    style GCPInfrastructure fill:#e6f4ea,stroke:#137333,stroke-width:2px;
    style AuthGatewayTier fill:#e0f2f1,stroke:#00796b,stroke-width:2px;
    style PrivateBackendTier fill:#fff3e0,stroke:#e65100,stroke-width:2px;
    style GeminiEnterpriseTier fill:#fef7e0,stroke:#f9ab00,stroke-width:2px;
```

---

## 🔄 Clean End-to-End Flow

The sequence below illustrates the complete token acquisition, validation, service-to-service IAM minting, and streaming grounding lifecycle:

```mermaid
sequenceDiagram
    autonumber
    actor User as Enterprise User (AlexW@...)
    participant PPT as Microsoft Office Client (PowerPoint/Word/Excel)
    participant FE as gemini-frontend (Cloud Run)
    participant Entra as Microsoft Entra ID (IdP)
    participant AP as auth-proxy Gateway (Cloud Run)
    participant DiscEng as Discovery Engine aclConfig
    participant GP as askgemini-proxy (Private Cloud Run)
    participant GE as Gemini Enterprise (streamAssist)

    User->>PPT: Opens 'Gemini Assistant' in Office Ribbon
    PPT->>FE: Fetches taskpane.html, JS bundles & assets
    FE-->>PPT: Renders Taskpane UI (Status: Connecting...)
    
    rect rgb(240, 244, 255)
        Note over PPT,Entra: Phase 1: Microsoft 365 Office SSO Token Acquisition
        PPT->>PPT: Office.auth.getAccessToken({ forMSGraphAccess: false })
        PPT->>PPT: Validates SourceLocation domain == Resource URI domain (gemini-frontend...)
        PPT->>Entra: Requests access token for api://gemini-frontend-.../b990d644-...
        Entra-->>PPT: Returns signed Microsoft Entra ID JWT (claims: upn, email, tid, aud)
        PPT->>PPT: Updates UI header: AlexW@5m4qby.onmicrosoft.com [Logged in]
    end

    User->>PPT: Submits prompt: "What are our Q3 revenue targets?"
    
    rect rgb(240, 255, 240)
        Note over PPT,AP: Phase 2: Gateway Verification & IdP Auto-Discovery
        PPT->>AP: POST /askGeminiEnterprise with "Authorization: Bearer <Entra_JWT>"
        AP->>AP: Verifies Microsoft signature via JWKS & validates aud/tid
        AP->>DiscEng: GET /v1/projects/.../locations/global/aclConfig (cached)
        DiscEng-->>AP: Returns { idpConfig: { idpType: "GSUITE" } }
        AP->>AP: Attaches X-End-User-Email and X-End-User-ID claims
    end

    rect rgb(255, 248, 240)
        Note over AP,GP: Phase 3: Service-to-Service Google IAM Authentication
        AP->>AP: Fetches Google OIDC ID token for audience: https://askgemini-proxy...
        AP->>GP: POST /askGeminiEnterprise (Headers: Authorization: Bearer <Google_IAM_Token>)
        GP->>GP: Cloud Run IAM validates auth-proxy-sa has roles/run.invoker
    end

    rect rgb(255, 255, 240)
        Note over GP,GE: Phase 4: Gemini Enterprise Grounded Stream
        GP->>GE: POST /v1/.../servingConfigs/default_search:streamAssist (with user context)
        GE->>GE: Grounded retrieval across enterprise datastores
        GE-->>GP: Streams SSE chunks with answers, citations, and grounding references
        GP-->>AP: Streams SSE response
        AP-->>PPT: Streams SSE response
    end

    PPT->>PPT: Renders answer incrementally in Taskpane UI
    PPT->>User: Displays formatted markdown, citations, and Insert Slide action buttons
```

---

## 🔒 Key Architectural Principles & Invariants

### 1. The Office.js SSO Domain-Matching Rule
Microsoft Office Add-in SSO requires that the **`<SourceLocation>`** domain, the **`<WebApplicationInfo><Resource>`**, and the **Entra ID Application ID URI** match identically:
$$\text{Domain in } \langle\text{SourceLocation}\rangle \equiv \text{Domain in } \langle\text{Resource}\rangle \equiv \text{Entra ID Application ID URI}$$

* **Frontend Hosting URL:** `https://gemini-frontend-16933400417.us-central1.run.app/taskpane.html`
* **Manifest `<Resource>`:** `api://gemini-frontend-16933400417.us-central1.run.app/b990d644-e47b-4575-97b3-2067c488042b`
* **Entra ID Application ID URI:** `api://gemini-frontend-16933400417.us-central1.run.app/b990d644-e47b-4575-97b3-2067c488042b`

If the Application ID URI points to a backend or proxy domain (`auth-proxy`), Office detects a domain mismatch and immediately triggers **Office SSO Error 13007: Invalid resource Url specified in the manifest**.

### 2. Pre-Authorized Client Applications in Entra ID
To enable silent SSO without consent prompts across all Office platforms, the following 3 client IDs must be pre-authorized under **Expose an API**:
1. `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office on the Web - Word/PPT/Excel Online)
2. `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Office on the Web Companion / Outlook)
3. `00000002-0000-0ff1-ce00-000000000000` (Microsoft Office Desktop Client across macOS & Windows)

### 3. Decoupled Token Translation & Least Privilege
* **No Direct Internet Access to Core Backend:** `askgemini-proxy` is locked down with `--no-allow-unauthenticated`.
* **Runtime Service Account Identity:** `auth-proxy` runs as `auth-proxy-sa@agentspace-452714.iam.gserviceaccount.com` and only holds:
  - `roles/logging.logWriter` (Structured Cloud Logging)
  - `roles/discoveryengine.viewer` (Dynamic `aclConfig` auto-discovery)
  - `roles/run.invoker` on Cloud Run service `askgemini-proxy` (Private S2S communication)

---

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

---

## 🔍 Deep Dive: Vertex AI Search Datastore & Grounding

Enterprise financial analysis and corporate decision-making require **zero hallucination**, **deterministic factual accuracy**, and **strict citation traceability**. The platform implements Google Cloud's native Vertex AI Search Grounding architecture.

```mermaid
sequenceDiagram
    autonumber
    actor User as Corporate User (Word/PPT/Excel)
    participant Addin as Taskpane UI (Office.js)
    participant Proxy as Cloud Function (askGemini)
    participant VertexAI as Vertex AI API (VertexAI SDK)
    participant Datastore as Vertex AI Search Datastore (GCS)
    participant Flash as Gemini 2.5 Flash
    participant Imagen as Gemini 2.5 Flash Image

    User->>Addin: Enters prompt: "Compare Q2 Cloud vs Services performance with charts"
    Addin->>Proxy: POST /askGemini { prompt, history, sessionId, enableGrounding: true }
    
    rect rgb(240, 248, 255)
        Note over Proxy,Datastore: 1. Enterprise Grounding & Retrieval Phase
        Proxy->>VertexAI: getGenerativeModel with datastore tool definition
        VertexAI->>Datastore: Vector Semantic Search across indexed 10-Q/10-K filings
        Datastore-->>VertexAI: Relevant text chunks & document snippets
        VertexAI->>Flash: System prompt + User prompt + Retrieved ground truth chunks
        Flash-->>Proxy: Grounded Markdown text + ![Chart](image: prompt) triggers + Citations
    end

    rect rgb(255, 245, 238)
        Note over Proxy,Imagen: 2. Multimodal Visual Synthesis Phase
        Proxy->>Proxy: Balanced Parenthesis Regex extracts image prompt: "Clean 2D bar chart..."
        Proxy->>Imagen: generateContent(imagePrompt)
        Imagen-->>Proxy: Candidate inlineData (Base64 PNG)
        Proxy->>Proxy: Inline Base64 Data URI into HTML container
    end

    Proxy-->>Addin: JSON Response { result, citations, groundingMetadata, sessionId }
    
    rect rgb(245, 255, 245)
        Note over Addin,User: 3. Office Host Injection Phase
        Addin->>Addin: Parse Markdown, Tables, Visuals into Presentation/Document Model
        Addin->>User: Renders Title slide, Widescreen text layout, and Native 2D Charts
    end
```

### 1. Grounding Configuration & Tool Definition
In `gemini-o365-proxy/index.js`, grounding is configured via the Vertex AI Generative Model `tools` schema:

```javascript
const modelConfig = {
  model: 'gemini-2.5-flash',
  systemInstruction: {
    parts: [{ text: SYSTEM_INSTRUCTION_TEXT }]
  },
  generationConfig: {
    temperature: 0.4,
    maxOutputTokens: 8192
  }
};

if (enableGrounding && DATASTORE_ID) {
  modelConfig.tools = [
    {
      retrieval: {
        vertexAiSearch: {
          datastore: process.env.VERTEX_DATASTORE_ID // Format: projects/PROJECT_ID/locations/global/collections/default_collection/dataStores/DATASTORE_ID
        }
      }
    }
  ];
}
```

### 2. How Grounding Guarantees Precision
1. **Document Ingestion:** Unstructured enterprise files (SEC 10-K, 10-Q filings, earnings releases, board presentations) are ingested from Google Cloud Storage (`gs://...`) into the **Vertex AI Search Datastore**.
2. **Chunking & Hybrid Search:** Vertex AI Search creates dense vector embeddings alongside sparse lexical search indexes.
3. **Retrieval-Augmented Generation (RAG):** When a user asks a question, Vertex AI automatically queries the datastore, extracts top-k relevant passages, and feeds them into the Gemini 2.5 Flash attention context.
4. **Citation Extraction:** Grounding metadata, chunk references, and URI citations are returned alongside the response and rendered as verifiable footnotes in the Office taskpane.

---

## 🎨 Multimodal Visual Pipeline (gemini-2.5-flash-image)

When users request visual charts, graphs, or executive diagrams, the system utilizes a dual-model generation pipeline:

```mermaid
graph LR
    UserPrompt[User Prompt: 'Show revenue chart'] --> GeminiFlash[Gemini 2.5 Flash]
    GeminiFlash -->|Outputs Directive| Directive["![Chart](image: A modern 2D financial bar chart... ($94.5B)...)"]
    Directive --> BalancedParser[Balanced Parenthesis Parser]
    BalancedParser -->|Clean Prompt| NanoBanana[Gemini 2.5 Flash Image / Nano Banana]
    NanoBanana -->|Raw Base64 PNG| DataUriEmbedder[HTML & Base64 Inliner]
    DataUriEmbedder --> TaskpaneResponse[Taskpane / Host Document]
```

### Balanced Parenthesis Prompt Parser
Financial chart prompts frequently contain embedded currency and accounting numbers in parentheses (e.g. `Google Services ($94.54B)`). Standard regular expressions fail on nested parentheses. The proxy backend implements a stack-based counter parser:

```javascript
// Stack-based balanced parenthesis prompt extraction
const startMarkerRegex = /!\[([^\]]*)\]\((?:image:|image-prompt:|imagen:)\s*/gi;
while ((match = startMarkerRegex.exec(processed)) !== null) {
  let openCount = 1;
  let i = contentStartIndex;
  while (i < processed.length && openCount > 0) {
    if (processed[i] === '(') openCount++;
    else if (processed[i] === ')') openCount--;
    i++;
  }
  if (openCount === 0) {
    const prompt = processed.substring(contentStartIndex, i - 1).trim();
    // Generate high-resolution 2D image via Vertex AI
  }
}
```

---

## 🖥️ Client-Side Host Adapter Architecture

The add-in uses a polymorphic adapter pattern (`HostAdapterFactory`) to dynamically bind to the running Office application:

```mermaid
classDiagram
    class BaseAdapter {
        <<abstract>>
        +insertText(text)
        +getSelectedText()
        +insertCard(title, body)
        +insertVisualChart(base64)
    }
    class WordAdapter {
        +insertText(text)
        +insertHeading(text, level)
        +insertTable(headers, rows)
        +insertImage(base64)
        +insertInlineGeminiTrigger()
    }
    class PPTAdapter {
        +createExecutiveDeck(slides)
        +insertTitleSlide(title, subtitle)
        +insertContentSlide(title, bullets)
        +insertSplitVisualSlide(title, bullets, base64)
        +injectImageNativeMac(base64, bounds)
    }
    class ExcelAdapter {
        +getSelectedRangeValues()
        +insertMetricCard(kpiData)
        +highlightAnomalies(range)
        +insertSummarySheet(data)
    }

    BaseAdapter <|-- WordAdapter
    BaseAdapter <|-- PPTAdapter
    BaseAdapter <|-- ExcelAdapter
```

### 🍎 The macOS WKWebView Dual-Pipeline Solution (PowerPoint)

#### The Problem:
On macOS, PowerPoint executes Office add-ins inside a sandboxed `WKWebView`. Passing large base64 image data (>50KB) through `PowerPoint.run` shape APIs triggers an XPC buffer serialization failure in the macOS window server, silently dropping the image.

#### The Dual-Pipeline Architecture:
1. **Slide Creation & Geometry:** Inside `PowerPoint.run`, the slide, header, and left-hand text shapes (width `420px`) are created and committed.
2. **Slide Selection Bridge:** The engine reads `newSlide.id` and synchronizes the active viewport:
   ```javascript
   context.presentation.setSelectedSlides([newSlide.id]);
   await context.sync();
   ```
3. **Native Office Common API Injection:** The engine immediately calls the lower-level C++ Office Common API:
   ```javascript
   Office.context.document.setSelectedDataAsync(base64Image, {
     coercionType: Office.CoercionType.Image,
     imageLeft: 480,
     imageTop: 110,
     imageWidth: 440,
     imageHeight: 340
   });
   ```
This bypasses macOS XPC serialization constraints and provides 100% reliable image insertion across all macOS and Windows Office versions.

---

## 🔒 Security, Compliance & Deployment Model

```mermaid
graph TD
    subgraph EnterpriseBoundary ["Google Cloud Enterprise Boundary"]
        subgraph IAM ["Cloud IAM & Policies"]
            ServiceAccount["Cloud Run / GCF Service Account<br/>(roles/aiplatform.user)"]
        end
        
        subgraph Services ["Deployed Services"]
            Frontend["gemini-frontend (Cloud Run)<br/>Region: us-central1"]
            Backend["askGemini (Cloud Functions Gen 2)<br/>Region: us-central1"]
            DatastoreRes["Vertex AI Search Datastore<br/>Collection: default_collection"]
        end
    end
    
    OfficeClient["Microsoft 365 Client<br/>(Word / PPT / Excel)"] -->|HTTPS (TLS 1.3)| Frontend
    OfficeClient -->|HTTPS JSON (CORS)| Backend
    Backend -->|Native IAM Token| DatastoreRes
```

1. **Zero Secret Storage:** No API keys or static credentials reside in client code or manifests. Authentication to Vertex AI is handled through Google Cloud IAM Workload Identity and Service Account delegation.
2. **Data Residency:** All prompt tokens, document embeddings, and generated visuals remain strictly contained within the Google Cloud project (`genai-demo-catalog`, `us-central1`).
3. **CORS Hardened:** Strict CORS headers allow legitimate Microsoft 365 webview origins while blocking unauthenticated third-party scrapers.
