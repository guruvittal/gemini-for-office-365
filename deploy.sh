#!/bin/bash
set -e

# ==============================================================================
# Gemini for Office 365 - Cloud Run Deployment Script
# Supports 3-Tier Enterprise Architecture (Frontend -> Auth Proxy -> Backend Proxy)
# ==============================================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Configurable Parameters (override via environment variables or flags)
PROJECT_ID="${PROJECT_ID:-$(gcloud config get-value project 2>/dev/null || echo "vertexsearch-447722")}"
REGION="${REGION:-us-central1}"
DEPLOY_MODE="${DEPLOY_MODE:-three-tier}"  # "three-tier" (recommended) or "two-tier" (legacy)
TRACK="${TRACK:-cloud_identity}"          # "cloud_identity" (Track 2) or "wif" (Track 1)

# Service Names
BACKEND_SERVICE_NAME="${BACKEND_SERVICE_NAME:-askgemini-proxy}"
AUTH_PROXY_SERVICE_NAME="${AUTH_PROXY_SERVICE_NAME:-auth-proxy}"
FRONTEND_SERVICE_NAME="${FRONTEND_SERVICE_NAME:-gemini-frontend}"

# Discovery Engine & Enterprise App Configuration
GE_GCP_PROJECT_ID="${GE_GCP_PROJECT_ID:-${PROJECT_ID}}"
GE_GCP_LOCATION="${GE_GCP_LOCATION:-global}"
STREAM_ASSIST_ENDPOINT_LOCATION="${STREAM_ASSIST_ENDPOINT_LOCATION:-${GE_GCP_LOCATION}}"
ENTERPRISE_APP_ID="${GEMINI_ENTERPRISE_APP_ID:-new-ge-app_1780069391112}"
ALLOW_SERVICE_ACCOUNT_FALLBACK="${ALLOW_SERVICE_ACCOUNT_FALLBACK:-true}"

# Microsoft Entra ID & Google Identity Configuration
MICROSOFT_ENTRA_APP_ID="${MICROSOFT_ENTRA_APP_ID:-}"
MICROSOFT_ENTRA_TENANT_ID="${MICROSOFT_ENTRA_TENANT_ID:-}"
GOOGLE_OAUTH_CLIENT_ID="${GOOGLE_OAUTH_CLIENT_ID:-}"
WIF_AUDIENCE="${WIF_AUDIENCE:-}"

echo "============================================================"
echo " Deploying Gemini for Office 365 to Google Cloud Run"
echo " Project:       ${PROJECT_ID}"
echo " Region:        ${REGION}"
echo " Mode:          ${DEPLOY_MODE}"
echo " Identity Track:${TRACK}"
echo " Discovery App: ${ENTERPRISE_APP_ID}"
echo "============================================================"

# Ensure gcloud configuration
gcloud config set project "${PROJECT_ID}"
gcloud config set run/region "${REGION}"

# Ensure service account exists if needed
SA_NAME="gemini-office365-sa"
SA_EMAIL="${SA_NAME}@${PROJECT_ID}.iam.gserviceaccount.com"
if ! gcloud iam service-accounts describe "${SA_EMAIL}" --project="${PROJECT_ID}" >/dev/null 2>&1; then
  echo "Service account ${SA_EMAIL} not found, checking if default compute engine SA can be used or creating..."
  gcloud iam service-accounts create "${SA_NAME}" \
    --display-name="Gemini for Office 365 Runtime Service Account" \
    --project="${PROJECT_ID}" 2>/dev/null || true
fi

if [ "${DEPLOY_MODE}" = "three-tier" ]; then
  # ------------------------------------------------------------------------------
  # Step 1: Deploy Backend Proxy Service (geminiproxy / askgemini-proxy)
  # ------------------------------------------------------------------------------
  echo ""
  echo ">>> [1/3] Deploying Backend Proxy Service (${BACKEND_SERVICE_NAME})..."
  cd "${SCRIPT_DIR}/geminiproxy"

  gcloud run deploy "${BACKEND_SERVICE_NAME}" \
    --source=. \
    --region="${REGION}" \
    --project="${PROJECT_ID}" \
    --platform=managed \
    --allow-unauthenticated \
    --set-env-vars "\
GE_GCP_PROJECT_ID=${GE_GCP_PROJECT_ID},\
GE_GCP_LOCATION=${GE_GCP_LOCATION},\
STREAM_ASSIST_ENDPOINT_LOCATION=${STREAM_ASSIST_ENDPOINT_LOCATION},\
GEMINI_ENTERPRISE_APP_ID=${ENTERPRISE_APP_ID},\
BACKEND_MODE=streamassist,\
ENTERPRISE_COLLECTION_ID=default_collection,\
ENTERPRISE_ASSISTANT_ID=default_assistant,\
ALLOW_SERVICE_ACCOUNT_FALLBACK=${ALLOW_SERVICE_ACCOUNT_FALLBACK}" \
    --quiet

  BACKEND_URL=$(gcloud run services describe "${BACKEND_SERVICE_NAME}" --region="${REGION}" --project="${PROJECT_ID}" --format="value(status.url)")
  echo "✅ Backend Service deployed at: ${BACKEND_URL}"

  # ------------------------------------------------------------------------------
  # Step 2: Deploy Auth Gateway (authproxy / auth-proxy)
  # ------------------------------------------------------------------------------
  echo ""
  echo ">>> [2/3] Deploying Auth Gateway Service (${AUTH_PROXY_SERVICE_NAME})..."
  cd "${SCRIPT_DIR}/authproxy"

  AUTH_ENV_VARS="GE_GCP_PROJECT_ID=${GE_GCP_PROJECT_ID},GE_GCP_LOCATION=${GE_GCP_LOCATION},DOWNSTREAM_BACKEND_URL=${BACKEND_URL},USER_AUTH_MODE=${TRACK},VERBOSE_LOGGING=true"
  if [ -n "${MICROSOFT_ENTRA_APP_ID}" ]; then
    AUTH_ENV_VARS="${AUTH_ENV_VARS},MICROSOFT_ENTRA_APP_ID=${MICROSOFT_ENTRA_APP_ID}"
  fi
  if [ -n "${MICROSOFT_ENTRA_TENANT_ID}" ]; then
    AUTH_ENV_VARS="${AUTH_ENV_VARS},MICROSOFT_ENTRA_TENANT_ID=${MICROSOFT_ENTRA_TENANT_ID}"
  fi
  if [ -n "${GOOGLE_OAUTH_CLIENT_ID}" ]; then
    AUTH_ENV_VARS="${AUTH_ENV_VARS},GOOGLE_OAUTH_CLIENT_ID=${GOOGLE_OAUTH_CLIENT_ID}"
  fi
  if [ -n "${WIF_AUDIENCE}" ]; then
    AUTH_ENV_VARS="${AUTH_ENV_VARS},WIF_AUDIENCE=${WIF_AUDIENCE}"
  fi

  gcloud run deploy "${AUTH_PROXY_SERVICE_NAME}" \
    --source=. \
    --region="${REGION}" \
    --project="${PROJECT_ID}" \
    --platform=managed \
    --allow-unauthenticated \
    --set-env-vars "${AUTH_ENV_VARS}" \
    --quiet

  AUTH_PROXY_URL=$(gcloud run services describe "${AUTH_PROXY_SERVICE_NAME}" --region="${REGION}" --project="${PROJECT_ID}" --format="value(status.url)")
  echo "✅ Auth Gateway deployed at: ${AUTH_PROXY_URL}"

  # ------------------------------------------------------------------------------
  # Step 3: Deploy Frontend Add-in Service (microsoft-addin)
  # ------------------------------------------------------------------------------
  echo ""
  echo ">>> [3/3] Deploying Frontend Add-in Service (${FRONTEND_SERVICE_NAME})..."
  cd "${SCRIPT_DIR}/microsoft-addin"

  gcloud run deploy "${FRONTEND_SERVICE_NAME}" \
    --source=. \
    --region="${REGION}" \
    --project="${PROJECT_ID}" \
    --platform=managed \
    --allow-unauthenticated \
    --set-env-vars "GEMINI_PROXY_URL=${AUTH_PROXY_URL}/askGeminiEnterprise" \
    --port=8080 \
    --quiet

  FRONTEND_URL=$(gcloud run services describe "${FRONTEND_SERVICE_NAME}" --region="${REGION}" --project="${PROJECT_ID}" --format="value(status.url)")
  echo "✅ Frontend Service deployed at: ${FRONTEND_URL}"

else
  # Legacy 2-Tier Deployment
  echo ""
  echo ">>> [1/2] Deploying 2-Tier Backend Service (${BACKEND_SERVICE_NAME})..."
  cd "${SCRIPT_DIR}/geminiproxy"

  gcloud run deploy "${BACKEND_SERVICE_NAME}" \
    --source=. \
    --region="${REGION}" \
    --project="${PROJECT_ID}" \
    --platform=managed \
    --allow-unauthenticated \
    --set-env-vars GCP_PROJECT_ID="${PROJECT_ID}",GCP_REGION="${REGION}",GEMINI_MODEL="gemini-2.5-flash",GEMINI_IMAGE_MODEL="gemini-2.5-flash-image",BACKEND_MODE="streamassist",GEMINI_ENTERPRISE_APP_ID="${ENTERPRISE_APP_ID}",ALLOW_SERVICE_ACCOUNT_FALLBACK=true \
    --quiet

  BACKEND_URL=$(gcloud run services describe "${BACKEND_SERVICE_NAME}" --region="${REGION}" --project="${PROJECT_ID}" --format="value(status.url)")
  echo "✅ Backend Service deployed at: ${BACKEND_URL}"

  echo ""
  echo ">>> [2/2] Deploying 2-Tier Frontend Service (${FRONTEND_SERVICE_NAME})..."
  cd "${SCRIPT_DIR}/microsoft-addin"

  gcloud run deploy "${FRONTEND_SERVICE_NAME}" \
    --source=. \
    --region="${REGION}" \
    --project="${PROJECT_ID}" \
    --platform=managed \
    --allow-unauthenticated \
    --set-env-vars "GEMINI_PROXY_URL=${BACKEND_URL}/askGeminiEnterprise" \
    --port=8080 \
    --quiet

  FRONTEND_URL=$(gcloud run services describe "${FRONTEND_SERVICE_NAME}" --region="${REGION}" --project="${PROJECT_ID}" --format="value(status.url)")
  echo "✅ Frontend Service deployed at: ${FRONTEND_URL}"
fi

# ------------------------------------------------------------------------------
# Summary & Next Steps
# ------------------------------------------------------------------------------
echo ""
echo "============================================================"
echo " Deployment Complete!"
echo " Backend Service:    ${BACKEND_URL}"
if [ "${DEPLOY_MODE}" = "three-tier" ]; then
echo " Auth Proxy Gateway: ${AUTH_PROXY_URL}"
fi
echo " Frontend Service:   ${FRONTEND_URL}"
echo "============================================================"
echo ""
echo "Manifest Generation:"
echo "To generate your custom Microsoft Office manifest, run:"
echo "  python3 scripts/generate_manifest.py \\"
echo "    --track ${TRACK} \\"
echo "    --frontend-url ${FRONTEND_URL} \\"
if [ "${DEPLOY_MODE}" = "three-tier" ]; then
echo "    --auth-proxy-url ${AUTH_PROXY_URL} \\"
fi
if [ -n "${MICROSOFT_ENTRA_APP_ID}" ]; then
echo "    --entra-app-id ${MICROSOFT_ENTRA_APP_ID} \\"
fi
echo "    --output manifest-custom.xml"
echo ""
echo "To sideload in Office:"
echo "1. In Word/PowerPoint/Excel: Insert > Add-ins > My Add-ins > Upload My Add-in"
echo "2. Select your generated manifest XML"
echo "============================================================"
