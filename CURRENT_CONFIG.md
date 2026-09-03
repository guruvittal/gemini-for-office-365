# Current Live Deployment Configuration

This document captures the active runtime environment variables, service endpoints, and architectural settings configured across the deployed Google Cloud Run services for both authentication tracks.

---

## 1. Track 2: GSuite / Cloud Identity Track (`agentspace-452714`)

### A. `gemini-frontend`
* **Service URL:** `https://gemini-frontend-16933400417.us-central1.run.app`
* **Region:** `us-central1`
* **Service Account:** `gemini-office365-sa@agentspace-452714.iam.gserviceaccount.com`

| Variable Name | Value | Purpose |
| :--- | :--- | :--- |
| `GEMINI_PROXY_URL` | `https://auth-proxy-16933400417.us-central1.run.app/askGeminiEnterprise` | Points the taskpane frontend to the security gatekeeper proxy. |

---

### B. `auth-proxy`
* **Service URL:** `https://auth-proxy-16933400417.us-central1.run.app`
* **Region:** `us-central1`
* **Service Account:** `gemini-office365-sa@agentspace-452714.iam.gserviceaccount.com`

| Variable Name | Value | Purpose |
| :--- | :--- | :--- |
| `USER_AUTH_MODE` | `cloud_identity` | Configures the proxy to operate in Cloud Identity / GSuite 3-legged OAuth mode. |
| `REQUIRE_ENTRA_AUTH` | `true` | Enforces verification of Microsoft 365 Entra ID SSO tokens on all incoming taskpane calls. |
| `MICROSOFT_ENTRA_APP_ID` | `e871aa77-54f7-4310-a549-cad3b1edee4a` | Application (client) ID registered in Microsoft Entra ID. |
| `MICROSOFT_ENTRA_TENANT_ID` | `8ea14f5d-d857-4ceb-b0f5-7e27b174f795` | Directory (tenant) ID in Microsoft Entra ID. |
| `GOOGLE_OAUTH_CLIENT_ID` | `497524937986-66oh05fskrkufpv2he7fb00fmpd4nlt9.apps.googleusercontent.com` | Google OAuth 2.0 Web Client ID dynamically served to the frontend taskpane for 1-click Google Sign-In. |
| `DOWNSTREAM_BACKEND_URL` | `https://askgemini-proxy-16933400417.us-central1.run.app` | Internal downstream target service URL for authenticated requests. |
| `GE_GCP_PROJECT_ID` | `jeansson-gem-ent-ci` | Target Google Cloud project hosting the Discovery Engine / Gemini Enterprise engine. |
| `GE_GCP_LOCATION` | `us` | Regional location for Discovery Engine APIs. |
| `VERBOSE_LOGGING` | `true` | Enables rich audit and diagnostic trace logging. |

---

### C. `askgemini-proxy`
* **Service URL:** `https://askgemini-proxy-16933400417.us-central1.run.app`
* **Region:** `us-central1`
* **Service Account:** `gemini-office365-sa@agentspace-452714.iam.gserviceaccount.com`

| Variable Name | Value | Purpose |
| :--- | :--- | :--- |
| `BACKEND_MODE` | `streamassist` | Calls Discovery Engine `streamAssist` endpoint. |
| `GE_GCP_PROJECT_ID` | `jeansson-gem-ent-ci` | Target Google Cloud project containing the Gemini Enterprise engine. |
| `GE_GCP_LOCATION` | `us` | Regional location (`us`). |
| `STREAM_ASSIST_ENDPOINT_LOCATION` | `us` | Regional endpoint target (`us-discoveryengine.googleapis.com`). |
| `GEMINI_ENTERPRISE_APP_ID` | `gemini-enterprise-dummy-ap_1787693913560` | Gemini Enterprise App Engine ID. |
| `ENTERPRISE_COLLECTION_ID` | `default_collection` | Collection ID. |
| `ENTERPRISE_ASSISTANT_ID` | `default_assistant` | Assistant ID. |
| `ALLOW_SERVICE_ACCOUNT_FALLBACK` | `true` | Permits falling back to service account credentials when user tokens are omitted. |

---

## 2. Track 1: Workforce Identity Federation (WIF) Track (`agentspace-wif`)

### A. `gemini-frontend`
* **Service URL:** `https://gemini-frontend-1062675944253.us-central1.run.app`
* **Region:** `us-central1`
* **Service Account:** `gemini-office365-sa@agentspace-wif.iam.gserviceaccount.com`

| Variable Name | Value | Purpose |
| :--- | :--- | :--- |
| `GEMINI_PROXY_URL` | `https://auth-proxy-1062675944253.us-central1.run.app/askGeminiEnterprise` | Points the taskpane frontend to the WIF security gatekeeper proxy. |

---

### B. `auth-proxy`
* **Service URL:** `https://auth-proxy-1062675944253.us-central1.run.app`
* **Region:** `us-central1`
* **Service Account:** `gemini-office365-sa@agentspace-wif.iam.gserviceaccount.com`

| Variable Name | Value | Purpose |
| :--- | :--- | :--- |
| `USER_AUTH_MODE` | `auto` | Automatically detects and applies WIF token exchange via Google Cloud STS. |
| `WIF_AUDIENCE` | `//iam.googleapis.com/locations/global/workforcePools/ca-entra-id-oidc-pool/providers/entra-id-oidc-pool-provider` | Target Workforce Identity Federation Pool & Provider resource name. |
| `MICROSOFT_ENTRA_APP_ID` | `85fb5428-6249-4131-9eeb-f2436d5d4d8c` | Application (client) ID registered in Microsoft Entra ID. |
| `DOWNSTREAM_BACKEND_URL` | `https://askgemini-proxy-1062675944253.us-central1.run.app` | Internal downstream target service URL. |
| `GE_GCP_PROJECT_ID` | `agentspace-wif` | Target Google Cloud project hosting the Gemini Enterprise instance. |
| `GE_GCP_LOCATION` | `global` | Global Discovery Engine location. |
| `VERBOSE_LOGGING` | `true` | Enables audit and diagnostic trace logging. |

---

### C. `askgemini-proxy`
* **Service URL:** `https://askgemini-proxy-1062675944253.us-central1.run.app`
* **Region:** `us-central1`
* **Service Account:** `gemini-office365-sa@agentspace-wif.iam.gserviceaccount.com`

| Variable Name | Value | Purpose |
| :--- | :--- | :--- |
| `BACKEND_MODE` | `streamassist` | Calls Discovery Engine `streamAssist` endpoint. |
| `GE_GCP_PROJECT_ID` | `agentspace-wif` | Project hosting the Gemini Enterprise engine. |
| `GE_GCP_LOCATION` | `global` | Location for API requests (`global`). |
| `GEMINI_ENTERPRISE_APP_ID` | `instance-demos1_1774616568648` | Gemini Enterprise App Engine ID. |
| `ENTERPRISE_COLLECTION_ID` | `default_collection` | Collection ID. |
| `ENTERPRISE_ASSISTANT_ID` | `default_assistant` | Assistant ID. |
| `ALLOW_SERVICE_ACCOUNT_FALLBACK` | `true` | Permits falling back to service account credentials. |
