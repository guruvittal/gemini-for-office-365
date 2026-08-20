# Argolis Cloud Run Deployment Containers for Gemini for Office 365

This directory contains containerized backend and frontend services configured for deployment on the **Argolis GCP environment** (`vertexsearch-447722`).

---

## 📁 Directory Structure

* **[`backend/`](backend/)**: Node.js 22 Vertex AI & Gemini Enterprise Proxy microservice (`gemini-proxy`).
  * `index.js`: Microservice handling `/askGemini` and `/askGeminiEnterprise` requests.
  * `Dockerfile`: Container definition for Node.js 22 slim.
  * `package.json`: Dependencies (@google-cloud/vertexai, express, cors, etc.).
* **[`frontend/`](frontend/)**: Office 365 Add-in webview static server container (`gemini-frontend`).
  * `Dockerfile`: Multi-stage build (Node 22 build -> Nginx Alpine runner).
  * `nginx.conf`: Nginx routing and cache control configuration.
  * `src/`: Taskpane UI and Word/PowerPoint/Excel host adapters.
  * `manifest.xml`: Office Add-in manifest configuration.
* **[`deploy.sh`](deploy.sh)**: Automated shell script to build and deploy both containers to Cloud Run.

---

## 🚀 Quick Deployment

Run the automated deployment script:
```bash
cd /home/guruvittal/ge_o365/gemini-for-office-365/guruvittal_argolis_containers
./deploy.sh
```

Or deploy each container manually:

### 1. Deploy Backend Proxy
```bash
cd /home/guruvittal/ge_o365/gemini-for-office-365/guruvittal_argolis_containers/backend
gcloud run deploy gemini-proxy \
  --source=. \
  --region=us-central1 \
  --platform=managed \
  --allow-unauthenticated \
  --set-env-vars GCP_PROJECT_ID=vertexsearch-447722,GCP_REGION=us-central1,GEMINI_MODEL=gemini-2.5-flash,GEMINI_IMAGE_MODEL=gemini-2.5-flash-image \
  --project=vertexsearch-447722
```

### 2. Deploy Frontend Add-in
```bash
cd /home/guruvittal/ge_o365/gemini-for-office-365/guruvittal_argolis_containers/frontend
gcloud run deploy gemini-frontend \
  --source=. \
  --region=us-central1 \
  --platform=managed \
  --allow-unauthenticated \
  --port=8080 \
  --project=vertexsearch-447722
```
