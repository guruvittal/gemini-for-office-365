# Google Workspace Domain-Wide Delegation (DWD) Setup Guide
**Integration:** Gemini Enterprise for Microsoft Office 365  
**Target:** Google Drive End-User Access & Grounding  

---

## Overview

To allow the **Gemini Enterprise for Microsoft Office 365** integration to search and ground responses on end-users' personal and shared Google Drive documents, the **Google Workspace Super Administrator** must configure Domain-Wide Delegation (DWD) for the backend service account.

Once configured, the Cloud Run backend automatically mints short-lived Google OAuth tokens for authenticated enterprise users (e.g. `user@yourdomain.com`) on the fly.

---

## Prerequisites

* **Role Required:** Google Workspace Super Administrator
* **Console URL:** [https://admin.google.com](https://admin.google.com)
* **Estimated Time:** 2 minutes (One-time setup)

---

## Step-by-Step Instructions

### 1. Open Domain-Wide Delegation in Google Workspace

1. Navigate to **[admin.google.com](https://admin.google.com)** and log in with your **Google Workspace Super Admin** account.
2. In the left navigation menu, go to:
   **Security** ➔ **Access and data control** ➔ **API controls**
   *(Direct URL: `https://admin.google.com/ac/owl/domainwidedelegation`)*
3. Locate the **Domain-wide delegation** pane at the bottom of the page.
4. Click **Manage Domain-Wide Delegation**.
5. Click **Add new**.

---

### 2. Enter Service Account Client ID and Scopes

In the **Add a new client ID** dialog, fill in the following exact configuration:

#### **Client ID:**
```text
110472144010421785275
```
> **Note:** This is the globally unique OAuth 2 Client ID belonging to the dedicated backend service account (`gemini-office365-sa@agentspace-452714.iam.gserviceaccount.com`).

#### **OAuth Scopes (comma-separated):**
```text
https://www.googleapis.com/auth/cloud-platform,https://www.googleapis.com/auth/drive.readonly
```

| Scope | Purpose |
| :--- | :--- |
| `https://www.googleapis.com/auth/cloud-platform` | Allows querying Discovery Engine / Gemini Enterprise APIs with user context |
| `https://www.googleapis.com/auth/drive.readonly` | Read-only access to user's Google Drive files for search & grounding |

---

### 3. Authorize

1. Click **Authorize**.
2. Confirm that **Client ID `110472144010421785275`** appears in the active API clients table with the status **Enabled**.

---

## Security & Architecture Summary

* **Least Privilege:** The Service Account is granted **read-only** access to Drive files. It cannot modify, delete, or share user files.
* **Per-User ACL Enforcement:** Google Workspace strictly enforces each user's individual file access permissions. Users will only see documents they already have access to.
* **Cross-Project / Cross-Org Support:** Domain-Wide Delegation operates at the Google identity layer. The Service Account in the proxy project is securely authorized to access Workspace user data across organizational boundaries.
* **Token Lifecycle:** Google user access tokens are minted on-demand, encrypted in transit, cached in-memory for 1 hour max, and never written to disk or databases.

---
*Generated for Gemini Enterprise Microsoft Office 365 Integration*
