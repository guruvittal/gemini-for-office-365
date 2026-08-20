#!/bin/bash
set -e

# ==============================================================================
# Gemini for Office 365 - Cloud Run Deployment Script
# ==============================================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ID="vertexsearch-447722"
REGION="us-central1"
BACKEND_SERVICE_NAME="gemini-proxy"
FRONTEND_SERVICE_NAME="gemini-frontend"

echo "============================================================"
echo " Deploying Gemini for Office 365 to Google Cloud Run"
echo " Project: ${PROJECT_ID}"
echo " Region:  ${REGION}"
echo "============================================================"

# Ensure gcloud configuration
gcloud config set project "${PROJECT_ID}"
gcloud config set run/region "${REGION}"

# ------------------------------------------------------------------------------
# 1. Deploy Backend Proxy Container (geminiproxy)
# ------------------------------------------------------------------------------
echo ""
echo ">>> [1/3] Deploying Backend Proxy Service (${BACKEND_SERVICE_NAME})..."
cd "${SCRIPT_DIR}/geminiproxy"

gcloud run deploy "${BACKEND_SERVICE_NAME}" \
  --source=. \
  --region="${REGION}" \
  --platform=managed \
  --allow-unauthenticated \
  --set-env-vars GCP_PROJECT_ID="${PROJECT_ID}",GCP_REGION="${REGION}",GEMINI_MODEL="gemini-2.5-flash",GEMINI_IMAGE_MODEL="gemini-2.5-flash-image",BACKEND_MODE="streamassist" \
  --project="${PROJECT_ID}"

BACKEND_URL=$(gcloud run services describe "${BACKEND_SERVICE_NAME}" --region="${REGION}" --project="${PROJECT_ID}" --format="value(status.url)")
echo " Backend Service deployed successfully at: ${BACKEND_URL}"

# ------------------------------------------------------------------------------
# 2. Deploy Frontend Add-in Container (microsoft-addin)
# ------------------------------------------------------------------------------
echo ""
echo ">>> [2/3] Deploying Frontend Add-in Service (${FRONTEND_SERVICE_NAME})..."
cd "${SCRIPT_DIR}/microsoft-addin"

# Build and deploy frontend container
gcloud run deploy "${FRONTEND_SERVICE_NAME}" \
  --source=. \
  --region="${REGION}" \
  --platform=managed \
  --allow-unauthenticated \
  --port=8080 \
  --project="${PROJECT_ID}"

FRONTEND_URL=$(gcloud run services describe "${FRONTEND_SERVICE_NAME}" --region="${REGION}" --project="${PROJECT_ID}" --format="value(status.url)")
echo " Frontend Service deployed successfully at: ${FRONTEND_URL}"

# ------------------------------------------------------------------------------
# 3. Summary & Office Sideloading Instructions
# ------------------------------------------------------------------------------
echo ""
echo "============================================================"
echo " Deployment Complete!"
echo " Backend URL:  ${BACKEND_URL}"
echo " Frontend URL: ${FRONTEND_URL}"
echo "============================================================"
echo "To sideload into Word, PowerPoint, or Excel:"
echo "1. Verify microsoft-addin/manifest.xml points to ${FRONTEND_URL}/"
echo "2. Open Office > Insert > Add-ins > My Add-ins > Upload My Add-in"
echo "3. Upload microsoft-addin/manifest.xml (or manifest_gemini_enterprise.xml)"
echo "============================================================"
