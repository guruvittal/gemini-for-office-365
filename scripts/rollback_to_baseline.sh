#!/usr/bin/env bash
# ==============================================================================
# Rollback Script: Restore Baseline WIF & Agentspace-452714 Cloud Run Services
# Tag Snapshot: ca-snapshot-wif-baseline
# ==============================================================================

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
CONFIG_DIR="${REPO_ROOT}/baseline-configs"

PROJECT_ID="agentspace-452714"
REGION="us-central1"

echo "======================================================================"
echo " Starting Rollback to Baseline (Snapshot: ca-snapshot-wif-baseline)   "
echo " Target Project: ${PROJECT_ID} (Region: ${REGION})                    "
echo "======================================================================"

# 1. Verify baseline files exist
for svc in "auth-proxy" "askgemini-proxy" "gemini-frontend"; do
  file="${CONFIG_DIR}/${svc}-baseline.yaml"
  if [[ ! -f "${file}" ]]; then
    echo "ERROR: Baseline configuration file not found: ${file}"
    exit 1
  fi
done

# 2. Restore askgemini-proxy (Backend Proxy)
echo ""
echo "[1/3] Restoring Backend Proxy (askgemini-proxy)..."
gcloud run services replace "${CONFIG_DIR}/askgemini-proxy-baseline.yaml" \
  --project="${PROJECT_ID}" \
  --region="${REGION}"

# 3. Restore auth-proxy (Auth Gateway)
echo ""
echo "[2/3] Restoring Auth Gateway Proxy (auth-proxy)..."
gcloud run services replace "${CONFIG_DIR}/auth-proxy-baseline.yaml" \
  --project="${PROJECT_ID}" \
  --region="${REGION}"

# 4. Restore gemini-frontend (Frontend Web App)
echo ""
echo "[3/3] Restoring Frontend Add-in (gemini-frontend)..."
gcloud run services replace "${CONFIG_DIR}/gemini-frontend-baseline.yaml" \
  --project="${PROJECT_ID}" \
  --region="${REGION}"

echo ""
echo "======================================================================"
echo " SUCCESS: All Cloud Run services restored to baseline snapshot!       "
echo "======================================================================"
echo "Live Service URLs:"
echo " - auth-proxy:       https://auth-proxy-16933400417.us-central1.run.app"
echo " - askgemini-proxy:  https://askgemini-proxy-16933400417.us-central1.run.app"
echo " - gemini-frontend:  https://gemini-frontend-16933400417.us-central1.run.app"
echo "======================================================================"
