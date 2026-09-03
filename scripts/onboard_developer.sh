#!/bin/bash
# ==============================================================================
# Gemini for Microsoft 365 — Multi-Developer Onboarding Automation Script
#
# Automates:
# 1. Appending Developer Frontend URL to Microsoft Entra ID Application ID URIs
# 2. Granting Developer Service Account IAM roles on Central Gemini Enterprise Project
# 3. Granting Developer User email viewer access on Gemini Enterprise Project
#
# Author: Carlos Augusto, Principal Architect, Google
# ==============================================================================

set -e

GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${BLUE}====================================================================${NC}"
echo -e "${BLUE}  Gemini for Microsoft 365 — Developer Onboarding Tool              ${NC}"
echo -e "${BLUE}====================================================================${NC}"

usage() {
  echo -e "\nUsage: $0 [options]"
  echo "Options:"
  echo "  --dev-project-id       Developer GCP Project ID (e.g. dev-alex-sandbox)"
  echo "  --dev-project-num      Developer GCP Project Number (e.g. 987654321012)"
  echo "  --dev-email            Developer Google Email (e.g. alex@yourdomain.com)"
  echo "  --central-ge-project   Central Gemini Enterprise GCP Project (default: jeansson-gem-ent-ci)"
  echo "  --entra-app-id         Central Microsoft Entra App ID (default: e871aa77-54f7-4310-a549-cad3b1edee4a)"
  echo "  --region               GCP Region for Cloud Run (default: us-central1)"
  echo "  -h, --help             Show this help message"
  exit 1
}

# Default values
CENTRAL_GE_PROJECT="jeansson-gem-ent-ci"
ENTRA_APP_ID="e871aa77-54f7-4310-a549-cad3b1edee4a"
REGION="us-central1"
DEV_PROJECT_ID=""
DEV_PROJECT_NUM=""
DEV_EMAIL=""

while [[ "$#" -gt 0 ]]; do
  case $1 in
    --dev-project-id) DEV_PROJECT_ID="$2"; shift ;;
    --dev-project-num) DEV_PROJECT_NUM="$2"; shift ;;
    --dev-email) DEV_EMAIL="$2"; shift ;;
    --central-ge-project) CENTRAL_GE_PROJECT="$2"; shift ;;
    --entra-app-id) ENTRA_APP_ID="$2"; shift ;;
    --region) REGION="$2"; shift ;;
    -h|--help) usage ;;
    *) echo -e "${RED}Unknown parameter: $1${NC}"; usage ;;
  esac
  shift
done

if [[ -z "${DEV_PROJECT_ID}" ]]; then
  read -p "Enter Developer GCP Project ID: " DEV_PROJECT_ID
fi

if [[ -z "${DEV_PROJECT_NUM}" ]]; then
  echo -e "${YELLOW}Attempting to lookup project number for ${DEV_PROJECT_ID}...${NC}"
  DEV_PROJECT_NUM=$(gcloud projects describe "${DEV_PROJECT_ID}" --format="value(projectNumber)" 2>/dev/null || true)
  if [[ -z "${DEV_PROJECT_NUM}" ]]; then
    read -p "Enter Developer GCP Project Number: " DEV_PROJECT_NUM
  else
    echo -e "${GREEN}✓ Found Project Number: ${DEV_PROJECT_NUM}${NC}"
  fi
fi

# dev-email is optional (can be granted via Google Group at the domain level)
if [[ -z "${DEV_EMAIL}" ]]; then
  echo -e "${YELLOW}ℹ Developer user email not provided. Skipping individual user IAM (ensure group:domain level access is configured).${NC}"
fi

DEV_SA="gemini-office365-sa@${DEV_PROJECT_ID}.iam.gserviceaccount.com"
DEV_FRONTEND_HOST="gemini-frontend-${DEV_PROJECT_NUM}.${REGION}.run.app"
NEW_ENTRA_URI="api://${DEV_FRONTEND_HOST}/${ENTRA_APP_ID}"

echo -e "\n${BLUE}--- Target Configuration ---${NC}"
echo "• Developer Project ID      : ${DEV_PROJECT_ID}"
echo "• Developer Project Number  : ${DEV_PROJECT_NUM}"
echo "• Developer Service Account : ${DEV_SA}"
echo "• Developer Frontend Host   : ${DEV_FRONTEND_HOST}"
echo "• Central GE Project ID     : ${CENTRAL_GE_PROJECT}"
echo "• Microsoft Entra App ID    : ${ENTRA_APP_ID}"
echo "• New Application ID URI    : ${NEW_ENTRA_URI}"
echo "----------------------------"

# 1. Update Microsoft Entra ID Application ID URIs
echo -e "\n${YELLOW}[1/3] Updating Microsoft Entra ID Application ID URIs via Azure CLI...${NC}"
if command -v az &> /dev/null; then
  CURRENT_URIS=$(az ad app show --id "${ENTRA_APP_ID}" --query "identifierUris" -o tsv 2>/dev/null || true)
  if echo "${CURRENT_URIS}" | grep -q "${DEV_FRONTEND_HOST}"; then
    echo -e "${GREEN}✓ Application ID URI already registered in Entra ID.${NC}"
  else
    echo "Appending ${NEW_ENTRA_URI} to Entra ID App ${ENTRA_APP_ID}..."
    az ad app update --id "${ENTRA_APP_ID}" --identifier-uris ${CURRENT_URIS} "${NEW_ENTRA_URI}"
    echo -e "${GREEN}✓ Successfully updated Entra ID Application ID URIs!${NC}"
  fi
else
  echo -e "${RED}Azure CLI ('az') not found in PATH.${NC}"
  echo -e "${YELLOW}Please add the following URI manually in Azure Portal (App registrations -> Expose an API):${NC}"
  echo "  ${NEW_ENTRA_URI}"
fi

# 2. Grant Cross-Project IAM on Central Gemini Enterprise Project
echo -e "\n${YELLOW}[2/3] Granting IAM permissions on Central Gemini Enterprise Project (${CENTRAL_GE_PROJECT})...${NC}"

echo "Granting roles/discoveryengine.editor to ${DEV_SA}..."
gcloud projects add-iam-policy-binding "${CENTRAL_GE_PROJECT}" \
  --member="serviceAccount:${DEV_SA}" \
  --role="roles/discoveryengine.editor" \
  --condition=None --quiet

echo "Granting roles/serviceusage.serviceUsageConsumer to ${DEV_SA}..."
gcloud projects add-iam-policy-binding "${CENTRAL_GE_PROJECT}" \
  --member="serviceAccount:${DEV_SA}" \
  --role="roles/serviceusage.serviceUsageConsumer" \
  --condition=None --quiet

if [[ -n "${DEV_EMAIL}" ]]; then
  echo "Granting roles/discoveryengine.viewer to end-user ${DEV_EMAIL}..."
  gcloud projects add-iam-policy-binding "${CENTRAL_GE_PROJECT}" \
    --member="user:${DEV_EMAIL}" \
    --role="roles/discoveryengine.viewer" \
    --condition=None --quiet
fi
echo -e "${GREEN}✓ Successfully configured GCP IAM permissions!${NC}"

# 3. Reminder for Google OAuth Client ID Whitelist
echo -e "\n${YELLOW}[3/3] Google OAuth Web Client Whitelist Reminder (Track 2):${NC}"
echo "Please ensure the following URIs are whitelisted in your Google Cloud OAuth Client ID"
echo "(APIs & Services -> Credentials in project '${CENTRAL_GE_PROJECT}'):"
echo "  • JavaScript Origin : https://${DEV_FRONTEND_HOST}"
echo "  • Redirect Callback : https://${DEV_FRONTEND_HOST}/google-callback.html"

echo -e "\n${GREEN}====================================================================${NC}"
echo -e "${GREEN}  ✓ DEVELOPER ONBOARDING COMPLETE!                                  ${NC}"
echo -e "${GREEN}====================================================================${NC}"
