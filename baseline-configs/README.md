# Baseline Cloud Run Configuration Snapshot

**Snapshot Tag:** `ca-snapshot-wif-baseline`  
**Snapshot Date:** August 25, 2026  
**Primary GCP Project:** `agentspace-452714` (Project Number: `16933400417`)  
**Target Gemini Enterprise Project:** `agentspace-wif` (Engine: `instance-demos1_1774616568648`)  
**Region:** `us-central1`  
**Unified Service Account:** `gemini-office365-sa@agentspace-452714.iam.gserviceaccount.com`

---

## 📦 Snapshot Inventory

| Service | File | Description |
| :--- | :--- | :--- |
| **Backend Proxy** | [`askgemini-proxy-baseline.yaml`](askgemini-proxy-baseline.yaml) | Node.js 20 Express backend connecting to Discovery Engine StreamAssist on `agentspace-wif` |
| **Auth Gateway** | [`auth-proxy-baseline.yaml`](auth-proxy-baseline.yaml) | Python 3.11 FastAPI gateway handling Entra ID SSO token validation & Google S2S proxying |
| **Frontend Web App** | [`gemini-frontend-baseline.yaml`](gemini-frontend-baseline.yaml) | Nginx container hosting Office 365 add-in assets (PowerPoint, Word, Excel) |

---

## 🔄 How to Rollback to Baseline

### Method 1: Run the One-Click Rollback Script
```bash
./scripts/rollback_to_baseline.sh
```

### Method 2: Manual Declarative Restore via `gcloud`
```bash
# 1. Restore Backend Proxy
gcloud run services replace baseline-configs/askgemini-proxy-baseline.yaml \
  --project=agentspace-452714 \
  --region=us-central1

# 2. Restore Auth Gateway
gcloud run services replace baseline-configs/auth-proxy-baseline.yaml \
  --project=agentspace-452714 \
  --region=us-central1

# 3. Restore Frontend Add-in
gcloud run services replace baseline-configs/gemini-frontend-baseline.yaml \
  --project=agentspace-452714 \
  --region=us-central1
```

### Method 3: Restore Codebase via Git Tag
```bash
git checkout ca-snapshot-wif-baseline
```
