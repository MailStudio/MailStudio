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

echo "→ Installing dependencies…"
if [[ -f package-lock.json ]]; then
  npm ci --prefer-offline
else
  npm install --prefer-offline
fi

echo "→ Checking syntax…"
npm run check

if [[ "${MAILSTUDIO_SKIP_TESTS:-}" == "1" ]]; then
  echo "→ Skipping tests (MAILSTUDIO_SKIP_TESTS=1)…"
else
  echo "→ Running tests…"
  npm test
fi

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

# Keep the wrapper automation-friendly by default. Set MAILSTUDIO_OPEN_DIST=1
# for the old convenience behavior of opening the output folder on macOS.
if [[ "${MAILSTUDIO_OPEN_DIST:-}" == "1" && "$(uname)" == "Darwin" ]]; then
  open dist/
fi
