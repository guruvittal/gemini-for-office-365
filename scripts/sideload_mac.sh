#!/usr/bin/env bash
# ==============================================================================
# Gemini Enterprise for Microsoft 365 - macOS Sideloading Helper Script
# ==============================================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
MANIFEST_FILE="${1:-manifest-wif.xml}"
MANIFEST_PATH="$REPO_ROOT/$MANIFEST_FILE"

if [ ! -f "$MANIFEST_PATH" ]; then
  echo "❌ Error: Manifest file not found at $MANIFEST_PATH"
  exit 1
fi

echo "🚀 Installing Gemini Enterprise Office 365 Add-in ($MANIFEST_FILE) on macOS..."

APPS=("com.microsoft.Word" "com.microsoft.Powerpoint" "com.microsoft.Excel")

for app in "${APPS[@]}"; do
  WEF_DIR="$HOME/Library/Containers/$app/Data/Documents/wef"
  echo "📁 Target directory: $WEF_DIR"
  mkdir -p "$WEF_DIR"
  rm -rf "$WEF_DIR"/*
  cp -f "$MANIFEST_PATH" "$WEF_DIR/gemini-enterprise-manifest.xml"
  echo "   ✅ Manifest copied to $WEF_DIR/gemini-enterprise-manifest.xml"
done

echo ""
echo "✨ Sideloading complete!"
echo "👉 Next steps:"
echo "   1. Restart Microsoft Word, PowerPoint, or Excel."
echo "   2. Navigate to the 'Home' tab."
echo "   3. Click 'Gemini (WIF)' in the ribbon."
