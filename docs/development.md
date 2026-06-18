# Development & builds

## Getting started

```bash
npm install
npm start          # run in development
npm run check      # syntax-check every source file
npm test           # run the unit tests (settings normalization + URL sanitization)
npm run dist:mac   # build a .dmg and .zip for macOS
npm run dist       # build for the current platform (mac / win / linux)
```

There's also a convenience wrapper, `./build.sh` (`--dir` for a fast unpackaged
build, `--win` / `--all` for other targets), which syntax-checks first. Set
`MAILSTUDIO_OPEN_DIST=1` if you want it to open the `dist/` folder when done on
macOS.

Continuous integration runs `npm run check` and `npm test` on every push and pull
request via [`.github/workflows/ci.yml`](../.github/workflows/ci.yml).

## Cross-platform builds

| Platform | Output | Notes |
|---|---|---|
| **macOS** | `.dmg`, `.zip` | Ad-hoc signed (see Gatekeeper note below) |
| **Windows** | `.exe` (NSIS installer) | Build on Windows |
| **Linux** | `.AppImage`, `.deb` | Build on Linux |

electron-builder packages for the host OS, so build each platform's artifact on
that platform (or in CI). The macOS-only icon/signing step in
[`scripts/after-pack.js`](../scripts/after-pack.js) automatically no-ops on Windows
and Linux.

## Auto-update

Packaged builds check **GitHub Releases** on launch and every six hours, download
new versions in the background, and post a native notification prompting a restart
(handled by [`src/main/updater.js`](../src/main/updater.js) via `electron-updater`).
There's also a **Check for Updates…** item in the app menu.

To publish the next update:

1. Bump `version` in `package.json` and push a matching tag (e.g. `v0.3.0`).
   [`release.yml`](../.github/workflows/release.yml) builds the Windows and Linux
   artifacts plus update metadata (`latest*.yml`) and uploads them to a draft
   GitHub Release. Build the macOS artifact locally with `npm run dist:mac` and
   upload it to the same draft release.
2. Publish the release. Installed apps pick it up on their next check.

(Manual alternative: `npm run release` with a `GH_TOKEN` that can create releases —
builds the current platform only.)

> Auto-update is a no-op in development. On macOS, Squirrel.Mac only applies updates
> to a **notarized** build — until the app is notarized, macOS users update by
> downloading the new `.dmg`. Windows and Linux auto-update work with the current
> builds.

## Installing on macOS (Gatekeeper)

The build is code-signed but not notarized, so first launch is blocked.

**System Settings (macOS 13+):**
1. Drag **MailStudio** to `/Applications` and double-click (it will be blocked)
2. **System Settings → Privacy & Security → Open Anyway**
3. Launch again and confirm — remembered after that

**Terminal:**
```bash
xattr -dr com.apple.quarantine "/Applications/MailStudio.app"
```

## Signing in to Microsoft

Microsoft may prompt for a passkey or security key inside MailStudio:

- **Touch ID / platform passkey** — not supported (Chromium needs Apple's
  restricted browser entitlement, granted only to notarized browsers)
- **Hardware security keys (YubiKey, etc.)** — supported
- **Recommended** — use the back arrow or _"Sign in another way"_ to switch to
  Microsoft Authenticator push or a one-time code; both render in-page and work

Because all Microsoft tabs share one persistent session, you sign in **once** and
everything is authenticated.

## Project layout

```
src/
  main/
    main.js              — window management, BrowserViews, IPC, notifications
    settings-store.js    — persistent settings (services, theme, scratch, notif)
    secure-store.js      — safeStorage vault for OAuth tokens
    oauth.js             — PKCE engine + provider definitions
    api-feeds.js         — Graph + Asana API fetchers
    connections.js       — token lifecycle + feed accessors
    updater.js           — GitHub Releases auto-update (electron-updater)
    panel-preload.js     — sandboxed bridge between panel renderer and main
    service-preload.js   — title/URL reporter for web view tabs
  renderer/
    panel.html/css/js    — sidebar UI (tabs, feeds, settings, focus, scratchpad)
    menu.html/css/js     — tray dropdown menu
assets/
  icon.png
  trayTemplate.png       — menu bar icon (inactive)
  trayUnreadTemplate.png — menu bar icon (unread)
```
