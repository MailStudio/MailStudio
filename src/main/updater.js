// Auto-update via electron-updater. Update artifacts are pulled from the
// project's GitHub Releases (configured under `build.publish` in package.json).
//
// This degrades gracefully on every axis:
//   - In development (unpackaged) there are no artifacts, so it no-ops.
//   - On macOS, Squirrel.Mac only applies updates to a code-signed + notarized
//     build; for an ad-hoc-signed build electron-updater raises an error, which
//     we swallow rather than letting it surface to the user.
//   - If electron-updater itself is missing or fails to load, auto-update is
//     simply disabled.
let autoUpdater = null
try {
  ;({ autoUpdater } = require('electron-updater'))
} catch {
  // electron-updater unavailable — auto-update disabled.
}

const SIX_HOURS_MS = 6 * 60 * 60 * 1000

function isSupported(app) {
  return Boolean(autoUpdater && app && app.isPackaged)
}

function log(err) {
  console.error('[updater]', err && err.message ? err.message : err)
}

// Wire up periodic checks. Call once after the app is ready.
function init({ app } = {}) {
  if (!isSupported(app)) return
  autoUpdater.autoDownload = true
  autoUpdater.on('error', log)
  check()
  const timer = setInterval(check, SIX_HOURS_MS)
  if (typeof timer.unref === 'function') timer.unref()
}

// Check now and, if an update is available, download it and post a native
// notification prompting the user to restart. Safe to call manually (e.g. from
// a "Check for Updates…" menu item) — no-ops when unsupported.
function check() {
  if (!autoUpdater) return
  try {
    // checkForUpdatesAndNotify returns a promise that rejects on failure — the
    // sync try/catch alone would leave that rejection unhandled.
    autoUpdater.checkForUpdatesAndNotify()?.catch((e) => log(e))
  } catch (err) {
    log(err)
  }
}

module.exports = { init, check, isSupported }
