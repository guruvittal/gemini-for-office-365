# Microsoft Entra ID Setup & Cloud Run Deployment Guide for `auth-proxy`

This guide provides complete, step-by-step instructions for:
1. **Initial Microsoft Entra ID App Registration** to capture your Client ID.
2. **Deploying `auth-proxy` to Google Cloud Run** to obtain your public service URL.
3. **Configuring Entra ID Application ID URI & Scopes** using your Cloud Run URL and pre-authorizing Microsoft Office 365 client applications.
4. **Manually testing and verifying** authentication and user identity extraction.

---

## 1. Overview & Architecture

The `auth-proxy` service acts as an authenticated security gateway between Microsoft Office 365 Add-ins (PowerPoint, Excel, Word) and downstream backend services:

```
┌─────────────────────────────────────────────────────────┐
│               Microsoft 365 Client Host                 │
│              (PowerPoint, Excel, Word)                  │
│  - Calls Office.auth.getAccessToken() (Silent Request)  │
└────────────────────────────┬────────────────────────────┘
                             │ 1. Requests O365 JWT Token
                             ▼
┌─────────────────────────────────────────────────────────┐
│                   Microsoft Entra ID                    │
└────────────────────────────┬────────────────────────────┘
                             │ 2. Issues valid JWT Access Token
                             ▼
┌─────────────────────────────────────────────────────────┐
│              Office Add-in Web Taskpane                 │
└────────────────────────────┬────────────────────────────┘
                             │ 3. POST /askGeminiEnterprise
                             │    Header: "Authorization: Bearer <JWT>"
                             ▼
┌─────────────────────────────────────────────────────────┐
│                 auth-proxy (Cloud Run)                  │
│  - PyJWKClient: Dynamically fetches Microsoft JWKS keys │
│  - Verifies RS256 Signature & Expiration claim (exp)    │
│  - Validates Audience claim (aud matches Client ID/URI) │
│  - Extracts User Identity (preferred_username, email)   │
└─────────────────────────────────────────────────────────┘
```

---

## 2. Recommended Deployment Order

Because the Microsoft Entra **Application ID URI** requires the public domain of your Cloud Run service (e.g., `api://auth-proxy-xxxxx-uc.a.run.app/<CLIENT_ID>`), follow this sequence:

```
Step 1: Register App in Entra ID ➔ Capture Client ID
       │
Step 2: Deploy auth-proxy to Cloud Run with Client ID ➔ Capture Cloud Run URL
       │
Step 3: In Entra ID ➔ Set Application ID URI, Add Scopes & Pre-authorize Office
       │
Step 4: Verify & Test with cURL / Health probes
```

---

## 3. Phase 1: Microsoft Entra ID App Registration Strategy

Depending on how your organization synchronizes or federates identity with Google Cloud, your Entra ID App Registration strategy for this Office Add-in will differ.

### Path A: Google Cloud Identity / Google Workspace (Domain-Wide Delegation)
If your organization synchronizes Microsoft Entra ID users into Google Cloud Identity or Google Workspace (e.g., via Google Cloud Directory Sync), **you must create a new, dedicated App Registration** for the Office Add-in.

1. Navigate to the **[Microsoft Entra Admin Center](https://entra.microsoft.com/)** -> **App registrations**.
2. Click **+ New registration**:
   - **Name**: `GE Office 365 Assistant Add-On`
   - **Supported account types**: Select **`Multiple Entra ID tenants`** and choose **`Allow all tenants`**.
   - **Redirect URI**: Select **Single-page application (SPA)** and leave the URL blank.
3. Click **Register**.
4. **Capture Credentials**: Copy the **Application (client) ID**. You will use this as `YOUR_MICROSOFT_ENTRA_CLIENT_ID`.
5. *Continue to Step 3.2 below.*

### Path B: Google Cloud Workforce Identity Federation (WIF)
If your organization uses Google Cloud WIF to federate identities without syncing them to Cloud Identity, **you DO NOT need to create a new App Registration.** Instead, you will modify the existing App Registration that is already tied to your Google WIF Pool.

1. Navigate to the **[Microsoft Entra Admin Center](https://entra.microsoft.com/)** -> **App registrations**.
2. Locate and open the App Registration currently used by your Google WIF configuration.
3. **Capture Credentials**: Copy the **Application (client) ID**. You will use this as `YOUR_MICROSOFT_ENTRA_CLIENT_ID`.
4. *Continue to Step 3.2 below.*

---

### Step 3.2: Configure Basic API Permissions (For Both Paths)
1. Under the app's sidebar, click **API permissions**.
2. Click **+ Add a permission** -> **Microsoft Graph** -> **Delegated permissions**.
3. Ensure the following basic delegated permissions are added:
   - `openid` (Sign users in)
   - `profile` (View users' basic profile)
   - `email` (View users' email address)
   - `User.Read` (Sign in and read user profile)
4. Click **Add permissions**.
5. Click **Grant admin consent for [Your Tenant]** (if you have tenant admin rights).

---

## 4. Phase 2: Provision Service Account & Deploy `auth-proxy` to Cloud Run

The `auth-proxy` microservice acts as the single decoupled authentication and authorization gateway for all client add-ins. It runs under a dedicated, least-privilege Google Cloud Service Account (`auth-proxy-sa`) and forwards validated requests to the downstream Gemini/StreamAssist backend (`askgemini-proxy`) using Google Service-to-Service (S2S) IAM authentication.

```
┌─────────────────────────┐          ┌──────────────────────────────────────┐          ┌───────────────────────────────────────┐
│  Office 365 Add-in Task │          │         Cloud Run: auth-proxy        │          │       Cloud Run: askgemini-proxy      │
│  (PowerPoint / Excel)   │          │  (Runtime SA: auth-proxy-sa)         │          │  (--no-allow-unauthenticated)         │
└────────────┬────────────┘          └──────────────────┬───────────────────┘          └───────────────────┬───────────────────┘
             │                                          │                                                  │
             │ 1. POST /askGeminiEnterprise             │                                                  │
             │    Header: Bearer <Entra ID JWT>         │                                                  │
             │─────────────────────────────────────────>│                                                  │
             │                                          │ 2. Validate Entra ID JWT Claims                  │
             │                                          │ 3. Fetch Google OIDC ID Token from Metadata      │
             │                                          │                                                  │
             │                                          │ 4. Forward POST /askGeminiEnterprise             │
             │                                          │    Header: Bearer <Google S2S ID Token>          │
             │                                          │    Header: X-End-User-Id, X-End-User-Email       │
             │                                          │─────────────────────────────────────────────────>│
             │                                          │                                                  │ 5. Execute StreamAssist
             │                                          │                                                  │    with End-User Context
             │                                          │ 6. Response JSON                                 │
             │                                          │<─────────────────────────────────────────────────│
             │ 7. Authenticated JSON Response           │                                                  │
             │<─────────────────────────────────────────│                                                  │
```

### Prerequisites
- Google Cloud SDK (`gcloud`) installed and authenticated:
  ```bash
  gcloud auth login
  gcloud config set project YOUR_GCP_PROJECT_ID
  ```
- Enable required Google Cloud APIs:
  ```bash
  gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com iam.googleapis.com
  ```

### Step 4.1: Create Dedicated Service Account & IAM Roles
Create the dedicated service account and assign least-privilege roles for Cloud Logging and downstream Cloud Run invocation:

```bash
# 1. Create dedicated Service Account
gcloud iam service-accounts create auth-proxy-sa \
  --display-name="Auth Proxy Cloud Run Service Account" \
  --description="Dedicated runtime identity for auth-proxy microservice to invoke backend Cloud Run services securely"

# 2. Grant Cloud Logging Writer on the project
gcloud projects add-iam-policy-binding YOUR_GCP_PROJECT_ID \
  --member="serviceAccount:auth-proxy-sa@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com" \
  --role="roles/logging.logWriter"

# 3. Grant Discovery Engine Viewer on the project (Enables auto-discovery of Gemini Enterprise / Discovery Engine IdP aclConfig)
gcloud projects add-iam-policy-binding YOUR_GCP_PROJECT_ID \
  --member="serviceAccount:auth-proxy-sa@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com" \
  --role="roles/discoveryengine.viewer"

# 4. Grant Cloud Run Invoker on downstream askgemini-proxy
gcloud run services add-iam-policy-binding askgemini-proxy \
  --region=us-central1 \
  --member="serviceAccount:auth-proxy-sa@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com" \
  --role="roles/run.invoker"

# 5. Lock down downstream askgemini-proxy to private S2S traffic only
gcloud run services update askgemini-proxy \
  --region=us-central1 \
  --no-allow-unauthenticated
```

### Step 4.2: Deploy `auth-proxy` with Dedicated Service Account
From within the `authproxy/` directory:

```bash
gcloud run deploy auth-proxy \
  --source . \
  --project YOUR_GCP_PROJECT_ID \
  --region us-central1 \
  --platform managed \
  --allow-unauthenticated \
  --service-account auth-proxy-sa@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com \
  --set-env-vars "\
MICROSOFT_ENTRA_APP_ID=YOUR_MICROSOFT_ENTRA_CLIENT_ID,\
DOWNSTREAM_BACKEND_URL=https://askgemini-proxy-16933400417.us-central1.run.app,\
GCP_PROJECT_ID=YOUR_GCP_PROJECT_ID,\
GCP_LOCATION=global,\
USER_AUTH_MODE=auto,\
REQUIRE_ENTRA_AUTH=true,\
VERBOSE_LOGGING=true"
```

#### `auth-proxy` Environment Variable Reference

| Parameter | Type | Default | Description |
| :--- | :---: | :---: | :--- |
| `MICROSOFT_ENTRA_APP_ID` | String | *Required* | Entra ID (Azure AD) Application / Client ID (`b990d644-...`). |
| `DOWNSTREAM_BACKEND_URL` | URL | `""` | HTTPS URL of the private `askgemini-proxy` Cloud Run service. |
| `GCP_PROJECT_ID` | String | `agentspace-452714` | Target GCP project containing the Gemini Enterprise engine. |
| `GCP_LOCATION` | String | `global` | Location of Discovery Engine resources (`global`, `us`, `eu`). |
| `USER_AUTH_MODE` | String | `auto` | Token pass-through strategy: `auto` (inspects `aclConfig`), `cloud_identity`, `wif`, or `none`. |
| `WIF_AUDIENCE` | String | `""` | Explicit STS audience override (auto-discovered from `aclConfig` if left blank). |
| `WIF_PROVIDER_NAME` | String | `entra-id-oidc-pool-provider` | Workforce Identity Federation provider ID inside the workforce pool. |
| `REQUIRE_ENTRA_AUTH` | Boolean | `true` | When `true`, rejects unauthenticated requests with HTTP 401. Set `false` only for local dev. |
| `VERBOSE_LOGGING` | Boolean | `false` | When `true`, emits deep diagnostic JSON payload logs to Cloud Logging. |

---

### Step 4.3: Deploy `askgemini-proxy` Backend
From within the `geminiproxy/` directory:

```bash
gcloud run deploy askgemini-proxy \
  --source . \
  --project YOUR_GCP_PROJECT_ID \
  --region us-central1 \
  --no-allow-unauthenticated \
  --set-env-vars "\
GCP_PROJECT_ID=YOUR_GCP_PROJECT_ID,\
GEMINI_ENTERPRISE_APP_ID=test1-agentspace_1741135345115,\
BACKEND_MODE=streamassist,\
GCP_LOCATION=global,\
ENTERPRISE_COLLECTION_ID=default_collection,\
ENTERPRISE_ASSISTANT_ID=default_assistant,\
ALLOW_SERVICE_ACCOUNT_FALLBACK=true"
```

#### `askgemini-proxy` Environment Variable Reference

| Parameter | Type | Default | Description |
| :--- | :---: | :---: | :--- |
| `GCP_PROJECT_ID` | String | *Required* | Google Cloud Project ID hosting Discovery Engine & Vertex AI. |
| `GEMINI_ENTERPRISE_APP_ID` | String | *Required* | Gemini Enterprise Search / Assist Engine ID. |
| `BACKEND_MODE` | String | `streamassist` | Execution mode (`streamassist` for Discovery Engine, `vertex` for direct Vertex AI). |
| `GCP_LOCATION` | String | `global` | Discovery Engine collection/engine location. |
| `ENTERPRISE_COLLECTION_ID` | String | `default_collection` | Discovery Engine collection name. |
| `ENTERPRISE_ASSISTANT_ID` | String | `default_assistant` | Discovery Engine assistant resource ID. |
| `ALLOW_SERVICE_ACCOUNT_FALLBACK` | Boolean | `false` | When `true`, falls back to Service Account ADC if user token is absent, logging a prominent GCP `WARNING`. When `false`, strictly enforces end-user tokens (HTTP 403 on missing token). |

---

### Step 4.4: Capture Your Cloud Run Service URL
Once deployment completes, copy the Service URL provided in the output:
`https://auth-proxy-xxxxxxxxxx-uc.a.run.app`

Extract the hostname (e.g. `auth-proxy-xxxxxxxxxx-uc.a.run.app`) for use in Phase 3.

---

## 5. Phase 3: Microsoft Entra ID - Expose API & Authorize Office Clients

Return to your registered application in the Microsoft Entra Admin Center.

### Step 5.1: Set Application ID URI
1. In your app registration sidebar, navigate to **Expose an API**.
2. Next to **Application ID URI**, click **Add** (or **Set**).
3. Set the URI using your **Frontend Cloud Run domain** (`gemini-frontend`) and Client ID:
   ```text
   api://[YOUR_FRONTEND_CLOUD_RUN_HOSTNAME]/[YOUR_MICROSOFT_ENTRA_CLIENT_ID]
   ```
   > [!IMPORTANT]
   > Office SSO requires the domain in the Application ID URI to match the hosting domain of the Office Add-in frontend (`gemini-frontend`). Do not include a trailing slash.
   >
   > *Example*: `api://gemini-frontend-16933400417.us-central1.run.app/b990d644-e47b-4575-97b3-2067c488042b`
4. Click **Save**.

### Step 5.2: Define the `access_as_user` Scope
1. Under **Scopes defined by this API**, click **+ Add a scope**:
   - **Scope name**: `access_as_user`
   - **Who can consent**: **Admins and users**
   - **Admin consent display name**: `Access GE Office 365 Assistant on behalf of user`
   - **Admin consent description**: `Allows the Office Add-in to access the backend auth-proxy service under the identity of the signed-in corporate user.`
   - **User consent display name**: `Access GE Office 365 Assistant`
   - **User consent description**: `Allows the application to access GE Office 365 Assistant on your behalf.`
   - **State**: **Enabled**
2. Click **Add scope**.

### Step 5.3: Pre-Authorize Microsoft Office Client Applications
To enable seamless Single Sign-On (SSO) and prevent Office from presenting consent popups inside PowerPoint, Excel, and Word, pre-authorize the official Microsoft Office client applications:

1. Under **Expose an API** -> **Authorized client applications**, click **+ Add a client application**.
2. Add the following **3 official Microsoft Office Client Applications** (one by one), checking the box for your scope (`access_as_user`) for each:

   | Client ID | Platform / Application | Description |
   | :--- | :--- | :--- |
   | `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` | **Office on the Web** | Word, PowerPoint, Excel online ([office.com](https://www.office.com)) |
   | `d3590ed6-52b3-4102-aeff-aad2292ab01c` | **Office on the Web (WAC/Outlook)** | Web Application Companion & Office Web suite |
   | `00000002-0000-0ff1-ce00-000000000000` | **Office Desktop (Mac & Windows)** | Microsoft Office Desktop client suite (Word, PowerPoint, Excel) |

3. Under **API permissions** in the sidebar, click **Grant admin consent for [Tenant]** to ensure all users in the directory receive seamless SSO access without individual consent prompts.

> [!NOTE]
> **Client ID Coverage Breakdown**:
> 
> | Application / Platform | Covered by Pre-Authorization | Notes |
> | :--- | :---: | :--- |
> | **PowerPoint (Desktop Mac & Windows)** | ✅ **Yes** | Covered by `00000002-0000-0ff1-ce00-000000000000`. Silent SSO supported. |
> | **Excel (Desktop Mac & Windows)** | ✅ **Yes** | Covered by `00000002-0000-0ff1-ce00-000000000000`. Silent SSO supported. |
> | **Word (Desktop Mac & Windows)** | ✅ **Yes** | Covered by `00000002-0000-0ff1-ce00-000000000000`. Silent SSO supported. |
> | **Office on the Web (office.com)** | ✅ **Yes** | Covered by `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` and `d3590ed6-52b3-4102-aeff-aad2292ab01c`. |
> | **New Outlook (Windows, Mac, & Web)** | ✅ **Yes** | Modern Outlook client uses unified Office runtime. |
> 
> ---
> 
> ### Step 5.4: Enforce Token Version 2 in Entra ID Manifest (CRITICAL for Google STS / WIF)
> By default, Microsoft Entra ID issues v1.0 tokens (`iss: https://sts.windows.net/{tenant_id}/`). When using Google Cloud Workforce Identity Federation (WIF) or Google STS, Google strictly verifies OIDC tokens against the v2.0 endpoint (`iss: https://login.microsoftonline.com/{tenant_id}/v2.0`).
> 
> To configure Entra ID to emit v2.0 tokens for the Office Add-in:
> 
> 1. In the Entra ID App Registration sidebar, click **Manifest**.
> 2. Locate the `"api"` section and set `"requestedAccessTokenVersion"` to **`2`**:
>    ```json
>    "api": {
>        "acceptMappedClaims": null,
>        "knownClientApplications": [],
>        "requestedAccessTokenVersion": 2,
>        "oauth2PermissionScopes": [ ... ],
>        "preAuthorizedApplications": [ ... ]
>    }
>    ```
>    *(Alternatively, in older manifest schemas, locate top-level `"accessTokenAcceptedVersion"` and set to `2`).*
> 3. Click **Save**.
> 
> > [!IMPORTANT]
> > If `"requestedAccessTokenVersion"` is left as `null` or `1`, Google STS will reject token exchange with HTTP 400: `invalid_grant: The issuer in ID Token https://sts.windows.net/... does not match the expected ones: https://login.microsoftonline.com/.../v2.0`. Setting this property to `2` ensures immediate compatibility with Google STS and Gemini Enterprise licensing.
>
> ---
>
> ### Step 5.5: Add Optional Claims to the Access Token (Required for Identity Mapping)
> The Google Cloud WIF Pool is configured to extract the user's identity from the `email` claim in the Microsoft token to map to `google.subject`. Microsoft Entra ID does not include this claim in access tokens by default.
> 1. Navigate to **Token configuration** in the left menu.
> 2. Click **Add optional claim**.
> 3. Select **Access** as the token type.
> 4. Check the boxes for **email**, **upn**, and **preferred_username**.
> 5. Click **Add**. (If prompted to turn on Microsoft Graph email permissions, accept).

---

## 6. Phase 4: Update the Office Add-in Manifest (XML)

Once your Entra ID App Registration is fully configured, you must link the Microsoft Office client to it by modifying your Add-in's XML manifest file (e.g., `manifest-ca.xml`).

1. Open `manifest-ca.xml` in your code editor.
2. Locate the `<WebApplicationInfo>` section near the bottom of the file.
3. Update the `<Id>` tag with your Entra ID Application (client) ID.
4. Update the `<Resource>` tag with the Application ID URI you set in Step 5.1.

**Example Modification:**
```xml
    <WebApplicationInfo>
      <!-- Replace with your Microsoft Entra ID Client ID -->
      <Id>85fb5428-6249-4131-9eeb-f2436d5d4d8c</Id>
      
      <!-- Replace with your Application ID URI (must exactly match Entra ID) -->
      <Resource>api://gemini-frontend-16933400417.us-central1.run.app/85fb5428-6249-4131-9eeb-f2436d5d4d8c</Resource>
      
      <Scopes>
        <Scope>access_as_user</Scope>
      </Scopes>
    </WebApplicationInfo>
```

> [!NOTE]
> If the `Id` and `Resource` in the XML do not perfectly match the Microsoft Entra ID configuration, the silent SSO call `Office.auth.getAccessToken()` will fail with error 13003 or 13005.

---

## 7. Phase 5: Verification & Manual Testing

### Test 1: Verify Cloud Run Health
```bash
curl https://auth-proxy-xxxxxxxxxx-uc.a.run.app/health
```

**Expected Response**:
```json
{
  "status": "healthy",
  "service": "auth-proxy",
  "version": "1.0.0",
  "require_entra_auth": true,
  "entra_app_id_configured": true
}
```

---

### Test 2: Unauthenticated Request (Expecting 401 Unauthorized)
Send a request without an `Authorization` header:
```bash
curl -i -X POST https://auth-proxy-xxxxxxxxxx-uc.a.run.app/askGeminiEnterprise \
  -H "Content-Type: application/json" \
  -d '{"prompt": "Generate sales summary"}'
```

**Expected Response**:
```http
HTTP/1.1 401 Unauthorized
WWW-Authenticate: Bearer
Content-Type: application/json

{"detail":"Missing Authorization header with Bearer token."}
```

---

### Test 3: Invalid / Forged Token Request (Expecting 401 Unauthorized)
```bash
curl -i -X POST https://auth-proxy-xxxxxxxxxx-uc.a.run.app/askGeminiEnterprise \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.invalid.payload" \
  -d '{"prompt": "Generate sales summary"}'
```

**Expected Response**:
```http
HTTP/1.1 401 Unauthorized
WWW-Authenticate: Bearer error="invalid_token"
Content-Type: application/json

{"detail":"Invalid token: ..."}
```

---

### Test 4: Authenticated Request with Real Entra ID Token (Browser Flow)

Follow these steps to acquire a real corporate JWT access token using your browser and test the live `auth-proxy` service:

#### Step A: Configure Redirect URI in Microsoft Entra Admin Center
1. Navigate to **Microsoft Entra ID** -> **App registrations** -> `GE Office 365 Assistant Add-On` -> **Authentication**.
2. Click **Add Redirect URI** (or **Add a platform**).
3. In the right-hand panel, select **Single-page application**.
4. In the configuration window:
   - **Redirect URI**: `https://jwt.ms`
   - Under **Implicit grant and hybrid flows**:
     - Check **Access tokens (used for implicit flows)**
     - Check **ID tokens (used for implicit and hybrid flows)**
5. Click **Save** (or **Configure**).

---

#### Step B: Log In via Browser
1. Construct your authorization URL by replacing the `<YOUR_MICROSOFT_ENTRA_CLIENT_ID>` and `<YOUR_CLOUD_RUN_SERVICE_HOSTNAME>` placeholders below:

```text
https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=<YOUR_MICROSOFT_ENTRA_CLIENT_ID>&response_type=token&redirect_uri=https://jwt.ms&scope=api://<YOUR_CLOUD_RUN_SERVICE_HOSTNAME>/<YOUR_MICROSOFT_ENTRA_CLIENT_ID>/access_as_user%20openid%20profile%20email&response_mode=fragment&nonce=678910
```

> **Placeholders to replace**:
> - `<YOUR_MICROSOFT_ENTRA_CLIENT_ID>`: The Application (client) ID obtained in Phase 1 (e.g. `dc1eb951-0ad7-4147-9af6-5d7c7f853447`).
> - `<YOUR_CLOUD_RUN_SERVICE_HOSTNAME>`: The Cloud Run service domain obtained in Phase 2 (e.g. `auth-proxy-xxxxxxxxxx-uc.a.run.app`).

2. Paste the constructed URL into your web browser.
3. Sign in with your corporate Microsoft account (and accept permissions if prompted).
4. Once authenticated, Microsoft redirects your browser to `https://jwt.ms`.

---

#### Step C: Copy the Access Token
1. On the `jwt.ms` page, copy the entire encoded token string directly from the top text box labeled **"Enter token below (it never leaves your browser):"** (the full colored string starting with `eyJ...`).

---

#### Step D: Test with `curl`
In your terminal, replace `<YOUR_CLOUD_RUN_SERVICE_HOSTNAME>` with your Cloud Run host and run:

```bash
TOKEN="<PASTE_COPIED_TOKEN_HERE>"

curl -X POST https://<YOUR_CLOUD_RUN_SERVICE_HOSTNAME>/askGeminiEnterprise \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer $TOKEN" \
  -d '{
    "prompt": "Create an executive summary for Q3 sales report",
    "sessionId": "session_001"
  }'
```

**Expected Response**:
```json
{
  "result": "[Auth-Proxy Verified] Hello John Doe! Your request was authenticated via Microsoft Entra ID. User ID captured: 'john.doe@contoso.com'. Prompt received: 'Create an executive summary for Q3 sales report'.",
  "sessionId": "session_001",
  "authenticatedUser": {
    "userId": "john.doe@contoso.com",
    "email": "john.doe@contoso.com",
    "name": "John Doe",
    "tenantId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
  },
  "citations": [],
  "backendMode": "auth-proxy-isolated"
}
```

---

### Test 5: Inspecting User Claims via `/api/auth/me`
```bash
curl -X GET https://<YOUR_CLOUD_RUN_SERVICE_HOSTNAME>/api/auth/me \
  -H "Authorization: Bearer $TOKEN"
```

**Expected Response**:
```json
{
  "authenticated": true,
  "user_id": "john.doe@contoso.com",
  "name": "John Doe",
  "email": "john.doe@contoso.com",
  "tenant_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "oid": "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy",
  "sub": "zzzzzzzz-zzzz-zzzz-zzzz-zzzzzzzzzzzz",
  "roles": [],
  "scopes": ["access_as_user"]
}
```

## 7. API Reference & Available Endpoints

FastAPI provides built-in interactive OpenAPI Swagger documentation for all routes:
- **Swagger UI**: `https://<YOUR_CLOUD_RUN_SERVICE_HOSTNAME>/docs`
- **ReDoc UI**: `https://<YOUR_CLOUD_RUN_SERVICE_HOSTNAME>/redoc`
- **OpenAPI Schema**: `https://<YOUR_CLOUD_RUN_SERVICE_HOSTNAME>/openapi.json`

### Endpoints Overview

| Method | Path | Auth Required? | Purpose |
| :--- | :--- | :---: | :--- |
| `GET` | **`/`** | ❌ No (Public) | Cloud Run root health check probe |
| `GET` | **`/health`** | ❌ No (Public) | Dedicated container health & config probe |
| `GET` | **`/api/auth/me`** | 🔒 **Yes** (Bearer Token) | Returns full decoded user identity and claims |
| `POST` | **`/api/auth/validate`** | 🔒 **Yes** (Bearer Token) | Explicit token validation check |
| `POST` | **`/askGeminiEnterprise`** | 🔒 **Yes** (Bearer Token) | Primary Office Add-in chat & generation endpoint |
| `POST` | **`/api/gemini/chat`** | 🔒 **Yes** (Bearer Token) | Alias to `/askGeminiEnterprise` |
| `GET` | **`/docs`** | ❌ No (Public) | Interactive Swagger UI API documentation |
| `GET` | **`/redoc`** | ❌ No (Public) | Alternative ReDoc API documentation |

---

### Endpoint Details

#### 1. `GET /` and `GET /health`
- **Authentication**: Public
- **Description**: Returns container status, service name, version, and auth configuration flags.

#### 2. `GET /api/auth/me`
- **Authentication**: `Authorization: Bearer <token>`
- **Description**: Decodes and returns authenticated Microsoft Entra ID user identity claims (`user_id`, `name`, `email`, `tenant_id`, `oid`, `sub`, `roles`, `scopes`).

#### 3. `POST /api/auth/validate`
- **Authentication**: `Authorization: Bearer <token>`
- **Description**: Validates the Bearer token without executing downstream actions. Returns `200 OK` if valid or `401 Unauthorized` if expired or invalid.

#### 4. `POST /askGeminiEnterprise` (and alias `/api/gemini/chat`)
- **Authentication**: `Authorization: Bearer <token>`
- **Description**: Decoupled endpoint for Office 365 add-ins (PowerPoint, Excel, Word). Validates the user's Entra ID JWT, captures user identity, and returns structured response.
- **Request Body**:
  ```json
  {
    "prompt": "Summarize Q3 sales performance",
    "history": [],
    "sessionId": "session_001",
    "enableGrounding": true
  }
  ```

---

## 8. Configuration Reference

| Variable | Required | Default | Description |
| :--- | :---: | :---: | :--- |
| `MICROSOFT_ENTRA_APP_ID` | **Yes** | *None* | Entra ID Client ID (GUID) |
| `REQUIRE_ENTRA_AUTH` | No | `false` | When `true`, strictly validates JWT on all requests |
| `VERBOSE_LOGGING` | No | `false` | When `true`, logs detailed request contexts, latency, and full decoded user claims payloads |
| `ENTRA_JWKS_URL` | No | `https://login.microsoftonline.com/common/discovery/v2.0/keys` | Microsoft Public Key Endpoint |
| `PORT` | No | `8080` | Port for Cloud Run / HTTP server |
| `LOG_LEVEL` | No | `INFO` *(or `DEBUG` if `VERBOSE_LOGGING=true`)* | Logging level (`DEBUG`, `INFO`, `WARNING`, `ERROR`) |

---

## 9. Viewing Native GCP Cloud Logging

The service formats all logs into **GCP Structured JSON** format (`severity`, `message`, `structured_context`, `httpRequest`, `labels`).

### Viewing Logs via `gcloud` CLI:
```bash
# Read live streaming logs
gcloud logging tail "resource.type=cloud_run_revision AND resource.labels.service_name=auth-proxy"

# Read recent logs formatted with severity and message
gcloud logging read "resource.type=cloud_run_revision AND resource.labels.service_name=auth-proxy" \
  --limit 50 \
  --format="table(timestamp, severity, jsonPayload.message, jsonPayload.structured_context.user_id)"
```

### Viewing Logs in Google Cloud Console:
1. Navigate to **Google Cloud Console > Logging > Logs Explorer**.
2. Query filter:
   ```text
   resource.type="cloud_run_revision"
   resource.labels.service_name="auth-proxy"
   ```
3. Expand any log entry to filter by `jsonPayload.structured_context.user_context` or `jsonPayload.structured_context.user_id`.


