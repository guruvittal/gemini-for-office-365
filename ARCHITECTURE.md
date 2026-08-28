# Architecture & System Design: Gemini for Microsoft 365
**Author:** Carlos Augusto, Principal Architect, Google  
**License:** Apache-2.0  

---

## 📑 Executive Summary

**Gemini for Microsoft 365** is an enterprise-grade AI integration platform that connects Microsoft Office client applications (**Word**, **PowerPoint**, and **Excel**) with **Google Cloud Vertex AI**.

The system combines:
1. **Factual Enterprise Grounding** via **Vertex AI Search Datastores** (RAG over enterprise 10-K/10-Q filings, policies, and internal docs).
2. **Generative Intelligence** via **Gemini 2.5 Flash** for multi-turn conversational reasoning and content synthesis.
3. **Multimodal Visual Synthesis** via **Gemini 2.5 Flash Image** (Nano Banana) for dynamic 2D flat vector financial charts and infographics.
4. **Native Host Adapters** for Microsoft Word (inline `@gemini` insertion, text transformation), PowerPoint (multi-slide executive decks with native macOS WKWebView dual-pipeline image rendering), and Excel (range analysis, anomaly detection, KPI metric cards).

> 📖 **Decoupled Auth & S2S System Architecture:** For a deep dive into the decoupled Microsoft Entra ID authentication gateway, Google Cloud Service-to-Service IAM model, and implementation blueprints, see [DEVELOPER_ARCHITECTURE_GUIDE.md](DEVELOPER_ARCHITECTURE_GUIDE.md) and [DEPLOYMENT_INSTRUCTIONS.md](DEPLOYMENT_INSTRUCTIONS.md).

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
        GoogleDialog["Office Dialog API<br/>(google-auth.html / google-callback.html)"]
        
        WordApp <--> TaskpaneUI
        PPTApp <--> TaskpaneUI
        ExcelApp <--> TaskpaneUI
        TaskpaneUI <--> OfficeSSO
        TaskpaneUI <--> GoogleDialog
    end

    subgraph IdentityTier ["2. Microsoft Entra ID & Google Identity Tier"]
        EntraApp["Entra ID App Registration<br/>(App ID: e871aa77-...)"]
        AppUri["Application ID URI<br/>api://gemini-frontend-...run.app/e871aa77-..."]
        GoogleOAuth["Google Cloud OAuth 2.0 Web Client<br/>(Client ID: 497524937986-...apps.googleusercontent.com)"]
        
        OfficeSSO -->|1. Silent SSO Token Request| EntraApp
        EntraApp -->|2. Signed Microsoft JWT| OfficeSSO
        GoogleDialog -->|3-Legged Consent Flow| GoogleOAuth
        GoogleOAuth -->|Google User OAuth Token ya29...| GoogleDialog
    end

    subgraph GCPInfrastructure ["3. Google Cloud Platform (Project: agentspace-452714)"]
        FrontendRun["gemini-frontend (Cloud Run)<br/>Nginx Static Assets & Taskpane<br/>(Public HTTPS)"]
        
        subgraph AuthGatewayTier ["Auth & Token Translation Gateway"]
            AuthProxy["auth-proxy (Cloud Run)<br/>Python 3.11 / FastAPI<br/>Runtime SA: gemini-office365-sa"]
            ConfigEndpoint["Dynamic Config Gateway<br/>GET /api/config"]
            IdpDiscovery["Dynamic IdP Auto-Discovery<br/>GET /v1/.../aclConfig"]
        end
        
        subgraph PrivateBackendTier ["Private Core Inference Backend (Private S2S Only)"]
            BackendProxy["askgemini-proxy (Cloud Run)<br/>Node.js 20 Express Microservice<br/>(--no-allow-unauthenticated)"]
        end
        
        TaskpaneUI -->|"Loads Static Assets"| FrontendRun
        TaskpaneUI -->|"Fetches Dynamic Config"| ConfigEndpoint
        TaskpaneUI -->|"3. POST /askGeminiEnterprise + Entra Bearer JWT + X-End-User-Google-Token"| AuthProxy
        AuthProxy ---|"4. Discovers IdP Type (GSUITE vs THIRD_PARTY)"| IdpDiscovery
        AuthProxy -->|"5. Forward Request + Google S2S IAM Token + User Tokens"| BackendProxy
    end

    subgraph GeminiEnterpriseTier ["4. Google Cloud Discovery Engine / Gemini Enterprise"]
        StreamAssist["Discovery Engine API<br/>POST /v1/.../engines/gemini-enterprise-dummy-ap/assistants/default_assistant:streamAssist"]
        EnterpriseCorpus["Enterprise Grounding Corpus<br/>(Google Drive, GCS, Spanner, BigQuery)"]
        
        BackendProxy -->|"6. StreamAssist Request with User Context (Authorization: Bearer <ya29...>) + toolsSpec"| StreamAssist
        StreamAssist ---|"7. Semantic Grounding & User-Level Drive Search"| EnterpriseCorpus
        StreamAssist -.->|"8. Server-Sent Events (SSE) Stream"| BackendProxy
        BackendProxy -.->|"9. SSE Chunks"| AuthProxy
        AuthProxy -.->|"10. Stream to Client"| TaskpaneUI
    end

    style OfficeClientTier fill:#e8f0fe,stroke:#1a73e8,stroke-width:2px;
    style IdentityTier fill:#f3e8fd,stroke:#7b1fa2,stroke-width:2px;
    style GCPInfrastructure fill:#e6f4ea,stroke:#137333,stroke-width:2px;
    style AuthGatewayTier fill:#e0f2f1,stroke:#00796b,stroke-width:2px;
    style PrivateBackendTier fill:#fff3e0,stroke:#e65100,stroke-width:2px;
    style GeminiEnterpriseTier fill:#fef7e0,stroke:#f9ab00,stroke-width:2px;
```

---

## 🔄 End-to-End Execution Flows

The platform supports two distinct identity execution paths depending on enterprise IdP configuration.

### Track 1: Workforce Identity Federation (WIF) Silent Token Exchange

In Track 1, users experience seamless Single Sign-On with zero Google login prompts. Their Microsoft Entra ID JWT is exchanged on the fly for a Google STS federated token via RFC 8693.

```mermaid
sequenceDiagram
    autonumber
    actor User as Corporate User
    participant Office as Office Taskpane UI
    participant Entra as Microsoft Entra ID
    participant Auth as Cloud Run: auth-proxy
    participant STS as Google STS (sts.googleapis.com)
    participant Meta as GCP Metadata Server
    participant Backend as Cloud Run: askgemini-proxy
    participant Engine as Gemini / StreamAssist

    User->>Office: Submits prompt: "Summarize Q3 earnings"
    
    rect rgb(240, 248, 255)
        Note over Office,Entra: Phase 1: Silent Entra ID Token Acquisition
        Office->>Entra: Office.auth.getAccessToken({ forMSGraphAccess: false })
        Entra-->>Office: Returns Microsoft Entra ID JWT Bearer Token
    end

    rect rgb(255, 250, 235)
        Note over Office,Auth: Phase 2: Auth Gateway Validation & STS Exchange
        Office->>Auth: POST /askGeminiEnterprise<br/>Authorization: Bearer [Entra_JWT]<br/>Body: { prompt, history, sessionId }
        Auth->>Entra: Fetch/Match RS256 Public Key via JWKS cache
        Auth->>Auth: Verify signature, expiration (exp), audience (aud), tenant (tid)
        Auth->>Auth: Extract claims (user_id, email, name, tenant_id, oid)
        Auth->>STS: POST /v1/token (RFC 8693 Token Exchange)<br/>subject_token=[Entra_JWT], audience=[//iam.googleapis.com/workforcePools/...]
        STS-->>Auth: Returns Federated Google Access Token (ya29...)
    end

    rect rgb(235, 255, 235)
        Note over Auth,Backend: Phase 3: Google S2S IAM Token Exchange & Forwarding
        Auth->>Meta: GET /instance/service-accounts/default/identity?audience=https://askgemini-proxy-...
        Meta-->>Auth: Returns short-lived Google OIDC ID Token
        Auth->>Backend: POST /askGeminiEnterprise<br/>Authorization: Bearer [Google_ID_Token]<br/>Headers: X-End-User-Id, X-End-User-Email, X-End-User-Google-Token: ya29...<br/>Body: { prompt, sessionId, userPseudoId, authenticatedUser }
    end

    rect rgb(255, 240, 245)
        Note over Backend,Engine: Phase 4: Grounded Execution & Attribution
        Backend->>Engine: StreamAssist API Call with user attribution & STS token
        Engine-->>Backend: Grounded answer + citations
        Backend-->>Auth: HTTP 200 OK with AI result and citations
    end

    Auth-->>Office: HTTP 200 OK with authenticated AI payload
    Office-->>User: Renders formatted grounded response & citations
```

---

### Track 2: Cloud Identity 3-Legged Google User OAuth & S2S IAM

The sequence below illustrates the complete token acquisition, dynamic configuration retrieval, 3-legged Google user authorization, service-to-service IAM minting, and streaming grounding lifecycle:

```mermaid
sequenceDiagram
    autonumber
    actor User as Enterprise User (scim@...)
    participant PPT as Microsoft Office Client (PowerPoint/Word/Excel)
    participant FE as gemini-frontend (Cloud Run)
    participant Entra as Microsoft Entra ID (IdP)
    participant AP as auth-proxy Gateway (Cloud Run)
    participant Dialog as Office Dialog API (google-auth.html)
    participant GoogleAuth as accounts.google.com
    participant Callback as google-callback.html
    participant GP as askgemini-proxy (Private Cloud Run)
    participant GE as Gemini Enterprise (streamAssist)

    User->>PPT: Opens 'Gemini Assistant' in Office Ribbon
    PPT->>FE: Fetches taskpane.html, JS bundles & assets
    FE-->>PPT: Renders Taskpane UI (Status: Connecting...)
    
    rect rgb(240, 244, 255)
        Note over PPT,Entra: Phase 1: Microsoft 365 Office SSO Token Acquisition
        PPT->>PPT: Office.auth.getAccessToken({ forMSGraphAccess: false })
        PPT->>Entra: Requests access token for api://gemini-frontend-.../e871aa77-...
        Entra-->>PPT: Returns signed Microsoft Entra ID JWT (claims: upn, email, tid, aud)
        PPT->>PPT: Updates UI header: scim@jeansson.demo.altostrat.com [Logged in]
    end

    rect rgb(230, 245, 255)
        Note over PPT,AP: Phase 2: Dynamic Config & Google User Authorization (Cloud Identity Mode)
        PPT->>AP: GET /api/config
        AP-->>PPT: Returns { google_oauth_client_id: "497524937986-...", user_auth_mode: "cloud_identity" }
        User->>PPT: Clicks "📁 Connect Drive"
        PPT->>Dialog: Office.context.ui.displayDialogAsync(google-auth.html?client_id=...&login_hint=scim@...)
        Dialog->>GoogleAuth: Interactive Consent for drive.readonly + cloud-platform scopes
        User->>GoogleAuth: Grants Consent
        GoogleAuth->>Callback: Redirects to /google-callback.html#access_token=ya29...
        Callback->>PPT: Office.context.ui.messageParent({ google_token: "ya29..." })
        Callback->>Callback: Closes dialog
        PPT->>PPT: Caches token (Status: ✅ Drive Linked)
    end

    User->>PPT: Submits prompt: "Do you see documents referring api://...?"
    
    rect rgb(240, 255, 240)
        Note over PPT,AP: Phase 3: Gateway Verification & Header Pass-Through
        PPT->>AP: POST /askGeminiEnterprise (Headers: Authorization: Bearer <Entra_JWT>, X-End-User-Google-Token: ya29...)
        AP->>AP: Verifies Microsoft signature via JWKS & validates aud/tid
        AP->>AP: Attaches verified user claims (X-End-User-Email, X-End-User-ID) and forwards X-End-User-Google-Token
    end

    rect rgb(255, 248, 240)
        Note over AP,GP: Phase 4: Service-to-Service Google IAM Authentication
        AP->>AP: Fetches Google OIDC ID token for audience: https://askgemini-proxy...
        AP->>GP: POST /askGeminiEnterprise (Headers: Authorization: Bearer <Google_IAM_Token>, X-End-User-Google-Token: ya29...)
        GP->>GP: Cloud Run IAM validates gemini-office365-sa has roles/run.invoker
    end

    rect rgb(255, 255, 240)
        Note over GP,GE: Phase 5: Gemini Enterprise Grounded Stream
        GP->>GE: POST /v1alpha/.../assistants/default_assistant:streamAssist (Headers: Authorization: Bearer <ya29...>, toolsSpec: { vertexAiSearchSpec: {} })
        GE->>GE: Grounded retrieval across user-accessible Google Drive files
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
* **Runtime Service Account Identity:** `auth-proxy` runs as `gemini-office365-sa@agentspace-452714.iam.gserviceaccount.com` and only holds:
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

1. **Zero Secret Storage:** No API keys or static credentials reside in client code or manifests. Authentication to Vertex AI and Gemini Enterprise is handled through Google Cloud IAM Service Account delegation and user-consented OAuth tokens.
2. **Data Residency:** All prompt tokens, document embeddings, and generated visuals remain strictly contained within the Google Cloud project (`jeansson-gem-ent-ci`, `us-central1`).
3. **CORS Hardened:** Strict CORS headers allow legitimate Microsoft 365 webview origins while blocking unauthenticated third-party scrapers.

---

## ⚙️ Deployment Ordering & OAuth Client Dependencies

When deploying the solution for an environment with **Cloud Identity / Google Workspace Identity**:

```mermaid
flowchart LR
    Step1["1. Deploy Backend & Frontend<br/>(gemini-frontend, auth-proxy, geminiproxy)"] --> Step2["2. Obtain Frontend URL<br/>(https://gemini-frontend-...run.app)"]
    Step2 --> Step3["3. Create OAuth 2.0 Web Client<br/>(In Gemini Enterprise GCP Project)"]
    Step3 --> Step4["4. Set GOOGLE_OAUTH_CLIENT_ID<br/>(On auth-proxy Cloud Run service)"]
    Step4 --> Step5["5. Ready for Office 365 Users<br/>(Dynamic /api/config resolution)"]

    style Step1 fill:#e8f0fe,stroke:#1a73e8,stroke-width:2px;
    style Step2 fill:#fef7e0,stroke:#f9ab00,stroke-width:2px;
    style Step3 fill:#e6f4ea,stroke:#137333,stroke-width:2px;
    style Step4 fill:#f3e8fd,stroke:#7b1fa2,stroke-width:2px;
    style Step5 fill:#f0fdf4,stroke:#10b981,stroke-width:2px;
```

### Why Deployment Ordering Matters:
* **Redirect URI & Origin Dependency:** The Google OAuth 2.0 Web Client credentials must specify the exact JavaScript Origin (`https://gemini-frontend-...run.app`) and Redirect URI (`https://gemini-frontend-...run.app/google-callback.html`). These values are only known once the frontend Cloud Run service is deployed.
* **Dynamic Configuration Delivery:** By passing `GOOGLE_OAUTH_CLIENT_ID` to `auth-proxy` as an environment variable, the Office 365 add-in dynamically retrieves it at runtime via `/api/config`, preventing any static credential baking or build rebuilds when migrating across environments.

