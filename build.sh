#!/usr/bin/env bash
set -euo pipefail

# MailStudio build script
# Usage:
#   ./build.sh          — macOS DMG + ZIP (default)
#   ./build.sh --dir    — unpackaged app bundle (fastest, for quick testing)
#   ./build.sh --win    — Windows NSIS installer (requires Wine on macOS)
#   ./build.sh --all    — all platforms

cd "$(dirname "$0")"

MODE="${1:-}"
START=$(date +%s)

echo "→ Checking syntax…"
npm run check

echo "→ Installing dependencies…"
npm install --prefer-offline 2>/dev/null || npm install

echo "→ Building…"
case "$MODE" in
  --dir)  npm run dist:dir ;;
  --win)  npm exec -- electron-builder --win ;;
  --all)  npm exec -- electron-builder --mac --win ;;
  *)      npm run dist:mac ;;
esac

END=$(date +%s)
echo ""
echo "✓ Done in $((END - START))s — output in dist/"
echo ""
ls -lh dist/*.dmg dist/*.zip dist/*.exe dist/*.AppImage 2>/dev/null || ls -lh dist/
echo ""

# Open dist/ in Finder so you can grab the file immediately
if [[ "$(uname)" == "Darwin" ]]; then
  open dist/
fi
