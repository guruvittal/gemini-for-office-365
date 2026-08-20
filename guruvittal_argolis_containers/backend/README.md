# Gemini Enterprise Backend Proxy for Microsoft 365
**Author:** Sathya AG, Principal Architect, Google  
**License:** Apache-2.0  

A scalable, enterprise-grade Google Cloud Function (Gen 2) and Cloud Run microservice that bridges Microsoft 365 (Word, PowerPoint, Excel) with Google Cloud Vertex AI generative models.

---

## 🌟 Capabilities

- **Grounded Enterprise Intelligence:** Direct integration with Vertex AI Search Datastores for factual, enterprise-grounded document analysis and Q&A over quarterly reports, policies, and research filings.
- **Multimodal Visual Chart Generation:** Automatically generates high-resolution, flat 2D vector financial charts and infographics via `gemini-2.5-flash-image` (Nano Banana) and embeds them as base64 images.
- **Multi-Turn Session Memory:** In-memory caching and session context management for continuous multi-turn conversations.
- **Host-Optimized Output:** Built-in system instructions tuned specifically for Microsoft Word documents, PowerPoint widescreen presentations, and Excel spreadsheets.

---

## 🏗️ Architecture

```mermaid
graph LR
    O365[Microsoft 365 Add-in (Word / PPT / Excel)] -->|HTTPS POST JSON| Proxy[Gemini O365 Proxy Cloud Function]
    
    subgraph Google Cloud Vertex AI
        Proxy -->|Grounded Search| Search[Vertex AI Search Datastore]
        Proxy -->|Generative Text| Flash[Gemini 2.5 Flash]
        Proxy -->|Visual Charts| Imagen[Gemini 2.5 Flash Image]
    end
```

---

## 🚀 Quick Start & Local Setup

### 1. Prerequisites
- [Node.js](https://nodejs.org/) (v20+ recommended)
- [Google Cloud SDK (`gcloud`)](https://cloud.google.com/sdk/docs/install)
- A Google Cloud Project with Vertex AI API enabled (`aiplatform.googleapis.com`)

### 2. Environment Configuration
Copy the template configuration file:
```bash
cp .env.example .env
```

Configure the environment variables in `.env`:
```bash
GCP_PROJECT_ID=your-google-cloud-project-id
GCP_REGION=us-central1
VERTEX_DATASTORE_ID=projects/YOUR_PROJECT/locations/global/collections/default_collection/dataStores/YOUR_DATASTORE
GEMINI_MODEL=gemini-2.5-flash
GEMINI_IMAGE_MODEL=gemini-2.5-flash-image
```

### 3. Install Dependencies & Run Locally
```bash
npm install
npm start
```
The server will start locally at `http://localhost:8080`.

---

## ☁️ Deployment to Google Cloud

### Deploy as Cloud Function (Gen 2)
```bash
gcloud functions deploy askGemini \
  --gen2 \
  --runtime=nodejs20 \
  --region=us-central1 \
  --source=. \
  --entry-point=askGemini \
  --trigger-http \
  --no-allow-unauthenticated \
  --set-env-vars GCP_PROJECT_ID=YOUR_PROJECT_ID,VERTEX_DATASTORE_ID=YOUR_DATASTORE_ID \
  --project=YOUR_PROJECT_ID
```

### Allow Public Invocation (if using direct CORS from Add-in)
```bash
gcloud run services update askgemini \
  --region=us-central1 \
  --no-invoker-iam-check \
  --project=YOUR_PROJECT_ID
```

---

## 🔒 Security & Privacy Best Practices
- **No Hardcoded Secrets:** All project IDs, regions, and datastore configurations are strictly managed via environment variables.
- **Enterprise Grounding:** Only authenticated enterprise Vertex AI datastores are queried.
- **CORS Configured:** Permissive for enterprise add-in webviews while protecting internal execution pipelines.
