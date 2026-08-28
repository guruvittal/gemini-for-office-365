# Microsoft Entra ID Setup & Cloud Run Deployment Guide for `auth-proxy`
**Author:** Carlos Augusto, Principal Architect, Google  
**License:** Apache-2.0  

> [!NOTE]
> **Consolidated Master Guide Available:**  
> This guide has been consolidated into the unified master runbook: **[`DEPLOYMENT_INSTRUCTIONS.md`](../DEPLOYMENT_INSTRUCTIONS.md)**.  
> Please refer to **[`DEPLOYMENT_INSTRUCTIONS.md`](../DEPLOYMENT_INSTRUCTIONS.md)** for the complete, step-by-step first-time deployment instructions for both Track 1 (WIF) and Track 2 (GSuite), Entra ID app registration, manifest XML customizations, cross-project permissions, and full environment variables catalog.

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

Because `auth-proxy` requires the downstream backend URL of `askgemini-proxy`, Microsoft Entra's **Application ID URI** requires the public domain of `gemini-frontend` (e.g., `api://gemini-frontend-xxxxx.us-central1.run.app/<CLIENT_ID>`), and the **Google OAuth 2.0 Web Client** (for Cloud Identity mode) requires the frontend origin and redirect URI, follow this sequence:

```
Step 1: Register App in Entra ID ➔ Capture Entra Client ID
       │
Step 2: Deploy Backend Proxy (askgemini-proxy) ➔ Capture Backend URL
       │
Step 3: Deploy Frontend Add-in (gemini-frontend) ➔ Capture Frontend URL
       │
Step 4: Create Google OAuth 2.0 Web Client (GCP Console) ➔ Set Frontend URL & Callback URI ➔ Capture Google Client ID
       │
Step 5: Deploy auth-proxy to Cloud Run with Entra Client ID, Backend URL & GOOGLE_OAUTH_CLIENT_ID
       │
Step 6: In Entra ID ➔ Set Application ID URI, Add Scopes & Pre-authorize Office
       │
Step 7: Update Office Manifest (manifest-wif.xml / manifest-gsuite.xml) & Verify in Office 365
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

## 4. Phase 2: Provision Service Account & Deploy Microservices to Cloud Run

The microservices run under a unified, least-privilege Google Cloud Service Account (`gemini-office365-sa`). `auth-proxy` acts as the single decoupled authentication and authorization gateway for all client add-ins and forwards validated requests to the downstream Gemini/StreamAssist backend (`askgemini-proxy`) using Google Service-to-Service (S2S) IAM authentication.

```
┌─────────────────────────┐          ┌──────────────────────────────────────┐          ┌───────────────────────────────────────┐
│  Office 365 Add-in Task │          │         Cloud Run: auth-proxy        │          │       Cloud Run: askgemini-proxy      │
│  (PowerPoint / Excel)   │          │  (Runtime SA: gemini-office365-sa)   │          │  (--no-allow-unauthenticated)         │
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
  gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com iam.googleapis.com discoveryengine.googleapis.com
  ```

### Step 4.1: Create Dedicated Service Account & IAM Roles
Create the dedicated service account and assign least-privilege roles for Cloud Logging, Discovery Engine, and downstream Cloud Run invocation:

```bash
# 1. Create dedicated Service Account
gcloud iam service-accounts create gemini-office365-sa \
  --display-name="Gemini Office 365 Unified Service Account" \
  --description="Dedicated runtime identity for Gemini for Office 365 services"

# 2. Grant Cloud Logging Writer on the project
gcloud projects add-iam-policy-binding YOUR_GCP_PROJECT_ID \
  --member="serviceAccount:gemini-office365-sa@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com" \
  --role="roles/logging.logWriter"

# 3. Grant Discovery Engine Viewer on the project (Enables auto-discovery of Gemini Enterprise / Discovery Engine IdP aclConfig)
gcloud projects add-iam-policy-binding YOUR_GCP_PROJECT_ID \
  --member="serviceAccount:gemini-office365-sa@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com" \
  --role="roles/discoveryengine.viewer"

# 4. Grant Cloud Run Invoker on downstream askgemini-proxy
gcloud run services add-iam-policy-binding askgemini-proxy \
  --region=us-central1 \
  --member="serviceAccount:gemini-office365-sa@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com" \
  --role="roles/run.invoker"

# 5. Lock down downstream askgemini-proxy to private S2S traffic only
gcloud run services update askgemini-proxy \
  --region=us-central1 \
  --no-allow-unauthenticated
```

### Step 4.2: Deploy `askgemini-proxy` Backend First
From within the `geminiproxy/` directory:

```bash
cd geminiproxy
gcloud run deploy askgemini-proxy \
  --source . \
  --project YOUR_GCP_PROJECT_ID \
  --region us-central1 \
  --service-account gemini-office365-sa@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com \
  --allow-unauthenticated \
  --set-env-vars "\
GE_GCP_PROJECT_ID=YOUR_GEMINI_ENTERPRISE_PROJECT_ID,\
GEMINI_ENTERPRISE_APP_ID=YOUR_GEMINI_ENTERPRISE_APP_ID,\
BACKEND_MODE=streamassist,\
GE_GCP_LOCATION=global,\
ENTERPRISE_COLLECTION_ID=default_collection,\
ENTERPRISE_ASSISTANT_ID=default_assistant,\
ALLOW_SERVICE_ACCOUNT_FALLBACK=true"
```

> **Capture Backend URL**: Once deployed, copy the backend URL (e.g. `https://askgemini-proxy-XXXXXXXX.us-central1.run.app`) to supply as `DOWNSTREAM_BACKEND_URL` in the next step.

#### `askgemini-proxy` Environment Variable Reference

| Parameter | Type | Required / Optional | Default Value | Example Value | Description & Impact |
| :--- | :---: | :---: | :---: | :--- | :--- |
| `GE_GCP_PROJECT_ID` | String | **Required** | `process.env.GCP_PROJECT_ID` | `agentspace-wif` / `jeansson-gem-ent-ci` | Google Cloud Project ID hosting Discovery Engine & Vertex AI. |
| `GEMINI_ENTERPRISE_APP_ID` | String | **Required** *(in `streamassist` mode)* | `""` | `instance-demos1_1774616568648` / `gemini-enterprise-dummy-ap_1787693913560` | Gemini Enterprise Search / Assist Engine ID. |
| `BACKEND_MODE` | String | Optional | `streamassist` | `streamassist` | Execution mode (`streamassist` for Discovery Engine, `vertex` for direct Vertex AI). |
| `GE_GCP_LOCATION` | String | Optional | `global` | `global` / `us` | Discovery Engine collection/engine location (`global`, `us`, `eu`). |
| `STREAM_ASSIST_ENDPOINT_LOCATION` | String | Optional | `global` | `global` / `us` | API endpoint subdomain routing prefix (`global` or regional e.g. `us`). |
| `ENTERPRISE_COLLECTION_ID` | String | Optional | `default_collection` | `default_collection` | Discovery Engine collection name. |
| `ENTERPRISE_ASSISTANT_ID` | String | Optional | `default_assistant` | `default_assistant` | Discovery Engine assistant resource ID. |
| `ALLOW_SERVICE_ACCOUNT_FALLBACK` | Boolean | Optional | `false` | `false` (WIF) / `true` (GSuite) | When `true`, falls back to Service Account ADC if user token is absent. |
| `GCP_REGION` | String | Optional | `us-central1` | `us-central1` | GCP region for direct Vertex AI client (used in `vertex` mode). |
| `GEMINI_MODEL` | String | Optional | `gemini-2.5-flash` | `gemini-2.5-flash` | Generative text foundation model name for direct Vertex AI calls. |
| `GEMINI_IMAGE_MODEL` | String | Optional | `gemini-2.5-flash-image` | `gemini-2.5-flash-image` | Multimodal visual generation model for charts and diagrams. |
| `VERTEX_DATASTORE_ID` | String | Optional | `""` | `""` | Full resource name of Vertex AI Search Datastore in direct `vertex` mode. |
| `PORT` | Integer | Optional | `8080` | `8080` | Container listen port (injected by Cloud Run runtime). |

---

### Step 4.3: Deploy `auth-proxy` with Dedicated Service Account
From within the `authproxy/` directory:

```bash
cd ../authproxy
gcloud run deploy auth-proxy \
  --source . \
  --project YOUR_GCP_PROJECT_ID \
  --region us-central1 \
  --platform managed \
  --allow-unauthenticated \
  --service-account gemini-office365-sa@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com \
  --set-env-vars "\
MICROSOFT_ENTRA_APP_ID=YOUR_MICROSOFT_ENTRA_CLIENT_ID,\
DOWNSTREAM_BACKEND_URL=https://askgemini-proxy-XXXXXXXX.us-central1.run.app,\
GE_GCP_PROJECT_ID=YOUR_GEMINI_ENTERPRISE_PROJECT_ID,\
GE_GCP_LOCATION=global,\
USER_AUTH_MODE=auto,\
REQUIRE_ENTRA_AUTH=true,\
VERBOSE_LOGGING=true"
```

#### `auth-proxy` Environment Variable Reference

| Parameter | Type | Required / Optional | Default Value | Example Value | Description & Impact |
| :--- | :---: | :---: | :---: | :--- | :--- |
| `MICROSOFT_ENTRA_APP_ID` | String | **Required** | `""` | `85fb5428-6249-4131-9eeb-f2436d5d4d8c` / `e871aa77-54f7-4310-a549-cad3b1edee4a` | Entra ID (Azure AD) Application / Client ID. |
| `DOWNSTREAM_BACKEND_URL` | URL | **Required** | `""` | `https://askgemini-proxy-1062675944253.us-central1.run.app` | HTTPS URL of the private `askgemini-proxy` Cloud Run service. |
| `GE_GCP_PROJECT_ID` | String | Optional | `agentspace-452714` | `agentspace-wif` / `jeansson-gem-ent-ci` | Target GCP project containing the Gemini Enterprise engine. |
| `GE_GCP_LOCATION` | String | Optional | `global` | `global` / `us` | Location of Discovery Engine resources (`global`, `us`, `eu`). |
| `USER_AUTH_MODE` | String | Optional | `auto` | `auto` / `wif` / `cloud_identity` | Authentication mode: `auto` (auto-detects `GSUITE` vs `THIRD_PARTY`), `wif`, or `cloud_identity`. |
| `MICROSOFT_ENTRA_TENANT_ID` | String | Optional | `""` *(any tenant)* | `464c0986-459d-42b9-a68d-aff41ccd3b16` / `8ea14f5d-d857-4ceb-b0f5-7e27b174f795` | Enforces single-tenant locking. Tokens from other tenants are rejected with `401`. |
| `GOOGLE_OAUTH_CLIENT_ID` | String | Optional | `""` | `497524937986-66oh05fskrkufpv2he7fb00fmpd4nlt9.apps.googleusercontent.com` | Google Cloud OAuth 2.0 Web Client ID for 3-legged login in GSuite mode. |
| `WIF_AUDIENCE` | String | Optional | `""` *(auto-detected)* | `//iam.googleapis.com/locations/global/workforcePools/ca-entra-id-oidc-pool/providers/entra-id-oidc-pool-provider` | Explicit Google STS Workforce Pool provider audience string. |
| `WIF_PROVIDER_NAME` | String | Optional | `entra-id-oidc-pool-provider` | `entra-id-oidc-pool-provider` | Provider resource name under the workforce pool. |
| `REQUIRE_ENTRA_AUTH` | Boolean | Optional | `false` | `false` (WIF auto) / `true` (GSuite) | When `true`, rejects unauthenticated requests with HTTP 401. |
| `VERBOSE_LOGGING` | Boolean | Optional | `false` | `true` | When `true`, emits deep diagnostic JSON payload logs to Cloud Logging. |
| `LOG_LEVEL` | String | Optional | `INFO` *(or `DEBUG` if verbose)* | `DEBUG` | Logging level (`DEBUG`, `INFO`, `WARNING`, `ERROR`). |
| `DOWNSTREAM_TIMEOUT` | Integer | Optional | `300` | `300` | HTTP timeout in seconds when awaiting responses from `askgemini-proxy`. |
| `ENTRA_JWKS_URL` | URL | Optional | `https://login.microsoftonline.com/common/discovery/v2.0/keys` | `https://login.microsoftonline.com/common/discovery/v2.0/keys` | Public JWKS endpoint URL for downloading Microsoft RS256 signing certificates. |
| `PORT` | Integer | Optional | `8080` | `8080` | Container listen port (injected by Cloud Run runtime). |

---

### Step 4.4: Capture Your Auth Proxy Cloud Run Service URL
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
---

### Step 5.4: Enforce Token Version 2 in Entra ID Manifest (CRITICAL for Google STS / WIF)
By default, Microsoft Entra ID issues v1.0 tokens (`iss: https://sts.windows.net/{tenant_id}/`). When using Google Cloud Workforce Identity Federation (WIF) or Google STS, Google strictly verifies OIDC tokens against the v2.0 endpoint (`iss: https://login.microsoftonline.com/{tenant_id}/v2.0`).

To configure Entra ID to emit v2.0 tokens for the Office Add-in:

1. In the Entra ID App Registration sidebar, click **Manifest**.
2. Locate the `"api"` section and set `"requestedAccessTokenVersion"` to **`2`**:
   ```json
   "api": {
       "acceptMappedClaims": null,
       "knownClientApplications": [],
       "requestedAccessTokenVersion": 2,
       "oauth2PermissionScopes": [ ... ],
       "preAuthorizedApplications": [ ... ]
   }
   ```
   *(Alternatively, in older manifest schemas, locate top-level `"accessTokenAcceptedVersion"` and set to `2`).*
3. Click **Save**.

> [!IMPORTANT]
> If `"requestedAccessTokenVersion"` is left as `null` or `1`, Google STS will reject token exchange with HTTP 400: `invalid_grant: The issuer in ID Token https://sts.windows.net/... does not match the expected ones: https://login.microsoftonline.com/.../v2.0`. Setting this property to `2` ensures immediate compatibility with Google STS and Gemini Enterprise licensing.

---

### Step 5.5: Configure Optional Claims (`email`, `upn`, `preferred_username`, `groups`)
When Microsoft Office issues an access token via silent SSO (`Office.auth.getAccessToken`), Entra ID **omits user profile claims (`email`, `upn`, `preferred_username`) by default** unless they are configured as **Optional Claims**.

* **For WIF**: Google Cloud STS attribute mapping requires `email` (e.g., `google.subject = assertion.email.lowerAscii()`). If `email` is absent from the token, Google STS throws an attribute mapping error (`Could not obtain a value for google.subject from the given credential`), blocking the user's Gemini Enterprise license.
* **For Cloud Identity**: `auth-proxy` extracts `preferred_username`, `email`, and `upn` from the JWT claims to forward verified `X-End-User-Email` and `X-End-User-Id` headers.

You can configure Optional Claims using either the **Azure Portal UI** or by **editing the App Registration Manifest**:

#### Method A: Via Azure Portal UI (Token Configuration)
1. In your App Registration sidebar, click **Token configuration**.
2. Click **+ Add optional claim**.
3. Select **Access** as the token type.
4. Check the boxes for:
   - `email`
   - `upn`
   - `preferred_username`
   - `groups` *(optional, if using group-based access control)*
5. Click **Add**. If prompted to turn on Microsoft Graph email permissions, click **Turn on** / **Accept**.

#### Method B: Via App Registration Manifest JSON
1. In your App Registration sidebar, click **Manifest**.
2. Locate the `"optionalClaims"` section and populate `"accessToken"`:
   ```json
   "optionalClaims": {
       "accessToken": [
           {
               "name": "email",
               "source": null,
               "essential": false,
               "additionalProperties": []
           },
           {
               "name": "upn",
               "source": null,
               "essential": false,
               "additionalProperties": []
           },
           {
               "name": "preferred_username",
               "source": null,
               "essential": false,
               "additionalProperties": []
           },
           {
               "name": "groups",
               "source": null,
               "essential": false,
               "additionalProperties": []
           }
       ],
       "idToken": [],
       "saml2Token": []
   }
   ```
3. Click **Save**.

---

## 6. Phase 4: Update the Office Add-in Manifest (XML)

Once your Entra ID App Registration is fully configured, you must link the Microsoft Office client to it by configuring your Add-in's XML manifest metadata file ([`manifest-wif.xml`](manifest-wif.xml) or [`manifest-gsuite.xml`](manifest-gsuite.xml)).

### Step 6.1: Field-by-Field Manifest Modification Reference

| XML Element / Attribute | What to Update | Why It Must Be Updated | Example Value |
| :--- | :--- | :--- | :--- |
| `<Id>` *(top-level GUID)* | Unique Add-in GUID | Identifies the add-in in the Office catalog. If deploying both WIF and GSuite or running multiple environments, each manifest **must have a distinct GUID** to prevent ribbon conflicts. | `6b9f4a12-8923-4d32-bb15-99d9b89e9001` |
| `<ProviderName>` | Organization Name | Displayed in Microsoft 365 Admin Center and the Add-in info dialog. | `Google Cloud Architecture Team` |
| `<DisplayName DefaultValue="...">` | Display Label | The button title on the Word/PowerPoint/Excel Home ribbon. | `Gemini Assistant (WIF)` |
| `<IconUrl>` / `<HighResolutionIconUrl>` | URL to icon assets | HTTPS URLs pointing to your `gemini-frontend` Cloud Run service (`icon-32.png` and `icon-80.png`). | `https://gemini-frontend-1062675944253.us-central1.run.app/assets/icon-32.png` |
| `<SourceLocation DefaultValue="...">` | Taskpane HTML entrypoint | The exact HTTPS URL of `taskpane.html` on `gemini-frontend`. Office loads this in the webview. | `https://gemini-frontend-1062675944253.us-central1.run.app/taskpane.html` |
| `<AppDomains>` | Whitelisted origins | Office blocks unauthorized network calls. Must include your `gemini-frontend` domain. | `<AppDomain>https://gemini-frontend-1062675944253.us-central1.run.app</AppDomain>` |
| `<bt:Url id="msg.Taskpane.Url" DefaultValue="...">` | Taskpane URL resource string | Resource string reference for Office UI. Must match `<SourceLocation>`. | `https://gemini-frontend-1062675944253.us-central1.run.app/taskpane.html` |
| `<WebApplicationInfo><Id>` | Microsoft Entra App ID | Links the Office client to your Entra ID App. Office sends this to Entra ID for SSO. | `85fb5428-6249-4131-9eeb-f2436d5d4d8c` |
| `<WebApplicationInfo><Resource>` | Application ID URI | Must **exactly match** the Application ID URI configured in Entra ID (Step 5.1). Office checks this against token audience. | `api://gemini-frontend-1062675944253.us-central1.run.app/85fb5428-6249-4131-9eeb-f2436d5d4d8c` |
| `<WebApplicationInfo><Scopes><Scope>` | OAuth Permission Scope | **CRITICAL: Must have exactly ONE scope: `access_as_user`**. Having 0 or >1 scopes will fail manifest validation. | `<Scope>access_as_user</Scope>` |

### Step 6.2: Example `<WebApplicationInfo>` Snippet

```xml
    <WebApplicationInfo>
      <!-- Replace with your Microsoft Entra ID Client ID -->
      <Id>85fb5428-6249-4131-9eeb-f2436d5d4d8c</Id>
      
      <!-- Replace with your Application ID URI (must exactly match Step 5.1) -->
      <Resource>api://gemini-frontend-1062675944253.us-central1.run.app/85fb5428-6249-4131-9eeb-f2436d5d4d8c</Resource>
      
      <Scopes>
        <Scope>access_as_user</Scope>
      </Scopes>
    </WebApplicationInfo>
```

> [!WARNING]
> If `<Resource>` does not exactly match the Application ID URI set in Step 5.1, or if `<Id>` does not match the Application (Client) ID, Office silent SSO (`Office.auth.getAccessToken()`) will fail with error `13003` or `13005`.

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

## 8. Configuration Reference & Live Environment Examples

### Live Environment Examples Matrix

| Setting | 🅰️ WIF Deployment (`agentspace-wif`) | 🅱️ GSuite Deployment (`agentspace-452714`) |
| :--- | :--- | :--- |
| `GE_GCP_PROJECT_ID` | `agentspace-wif` | `jeansson-gem-ent-ci` |
| `GE_GCP_LOCATION` | `global` | `us` |
| `MICROSOFT_ENTRA_APP_ID` | `85fb5428-6249-4131-9eeb-f2436d5d4d8c` | `e871aa77-54f7-4310-a549-cad3b1edee4a` |
| `MICROSOFT_ENTRA_TENANT_ID` | `464c0986-459d-42b9-a68d-aff41ccd3b16` (`5m4qby.onmicrosoft.com`) | `8ea14f5d-d857-4ceb-b0f5-7e27b174f795` (GSuite M365 Tenant) |
| `USER_AUTH_MODE` | `auto` (or `wif`) | `cloud_identity` |
| `GOOGLE_OAUTH_CLIENT_ID` | *None* | `497524937986-66oh05fskrkufpv2he7fb00fmpd4nlt9.apps.googleusercontent.com` |
| `WIF_AUDIENCE` | `//iam.googleapis.com/locations/global/workforcePools/ca-entra-id-oidc-pool/providers/entra-id-oidc-pool-provider` | *None* |
| Office Manifest | [`manifest-wif.xml`](../manifest-wif.xml) | [`manifest-gsuite.xml`](../manifest-gsuite.xml) |

### Parameter Reference

| Variable | Type | Required / Optional | Default | Description |
| :--- | :---: | :---: | :---: | :--- |
| `MICROSOFT_ENTRA_APP_ID` | String | **Required** | *None* | Entra ID (Azure AD) Client Application GUID. |
| `MICROSOFT_ENTRA_TENANT_ID` | String | Optional | *None* | Entra ID Directory (Tenant) GUID. |
| `DOWNSTREAM_BACKEND_URL` | URL | **Required** | `""` | HTTPS endpoint of the downstream `askgemini-proxy` Cloud Run service. |
| `GE_GCP_PROJECT_ID` | String | **Required** | *None* | Target GCP project containing the Discovery Engine instance. (Legacy `GCP_PROJECT_ID` supported as fallback). |
| `GE_GCP_LOCATION` | String | Optional | `global` | Discovery Engine collection/engine location (`global`, `us`, `eu`). |
| `USER_AUTH_MODE` | String | Optional | `auto` | Identity strategy: `auto` (detects `GSUITE` vs `THIRD_PARTY`), `cloud_identity`, `wif`, or `none`. |
| `WIF_AUDIENCE` | String | Optional *(WIF only)* | `""` | Explicit STS audience override URL. Leave blank in WIF mode to auto-discover via `aclConfig`. Not used for Cloud Identity. |
| `WIF_PROVIDER_NAME` | String | Optional *(WIF only)* | `entra-id-oidc-pool-provider` | Workforce pool provider ID name. Not used for Cloud Identity. |
| `REQUIRE_ENTRA_AUTH` | Boolean | Optional | `true` | When `true`, strictly enforces Entra ID JWT validation (returns 401 on missing/invalid token). |
| `VERBOSE_LOGGING` | Boolean | Optional | `false` | When `true`, logs detailed request contexts, latency, and full decoded user claims payloads. |
| `ENTRA_JWKS_URL` | URL | Optional | `https://login.microsoftonline.com/common/discovery/v2.0/keys` | Microsoft Public Key Endpoint for JWKS verification. |
| `PORT` | Integer | Optional | `8080` | Port for Cloud Run / HTTP server. |
| `LOG_LEVEL` | String | Optional | `INFO` *(or `DEBUG` if `VERBOSE_LOGGING=true`)* | Logging level (`DEBUG`, `INFO`, `WARNING`, `ERROR`). |

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

---

## 10. Cross-Project Deployment (Cross-Project Gemini Enterprise)

When the **Cloud Run microservices** (`auth-proxy`, `askgemini-proxy`, `gemini-frontend`) are deployed in one GCP project (e.g., `PROJECT_A`), but the **Gemini Enterprise (Discovery Engine) instance** resides in a different GCP project (e.g., `PROJECT_B`), cross-project IAM access must be granted.

### Option A: Configure Cross-Project Access via `gcloud` (Recommended)

Run the following commands as an **Owner** or **IAM Admin** on the **Gemini Enterprise target project (Project B)**:

```bash
# Set your target project and Cloud Run service account
TARGET_GEMINI_PROJECT="YOUR_GEMINI_ENTERPRISE_PROJECT_ID"
CLOUD_RUN_SERVICE_ACCOUNT="YOUR_SERVICE_ACCOUNT@YOUR_CLOUD_RUN_PROJECT_ID.iam.gserviceaccount.com"

# 1. Grant Discovery Engine Editor access on the Gemini Enterprise project
gcloud projects add-iam-policy-binding "${TARGET_GEMINI_PROJECT}" \
  --member="serviceAccount:${CLOUD_RUN_SERVICE_ACCOUNT}" \
  --role="roles/discoveryengine.editor"

# 2. Grant Service Usage Consumer permission on the Gemini Enterprise project
gcloud projects add-iam-policy-binding "${TARGET_GEMINI_PROJECT}" \
  --member="serviceAccount:${CLOUD_RUN_SERVICE_ACCOUNT}" \
  --role="roles/serviceusage.serviceUsageConsumer"
```

### Option B: Configure via Google Cloud Console
1. Navigate to the **Gemini Enterprise GCP Project** in the [Google Cloud Console](https://console.cloud.google.com/).
2. Go to **IAM & Admin** ➔ **IAM** ➔ Click **+ Grant Access**.
3. **New principals**: Enter the Cloud Run service account (`YOUR_SERVICE_ACCOUNT@YOUR_CLOUD_RUN_PROJECT_ID.iam.gserviceaccount.com`).
4. **Assign roles**:
   - `Discovery Engine Editor` (`roles/discoveryengine.editor`)
   - `Service Usage Consumer` (`roles/serviceusage.serviceUsageConsumer`)
5. Click **Save**.

### Cloud Run Service Configuration for Cross-Project:
Set the following environment variables on the `askgemini-proxy` and `auth-proxy` Cloud Run services in Project A:
- `GE_GCP_PROJECT_ID`: Target project ID hosting Gemini Enterprise (e.g., `YOUR_GEMINI_ENTERPRISE_PROJECT_ID`)
- `GEMINI_ENTERPRISE_APP_ID`: Target Engine/App ID (e.g., `YOUR_GEMINI_ENTERPRISE_APP_ID`)
- `GE_GCP_LOCATION`: Location of collection/engine resource (`global`, `us`, or `eu`)
- `STREAM_ASSIST_ENDPOINT_LOCATION`: Regional API endpoint prefix (`global`, `us`, or `eu`)



