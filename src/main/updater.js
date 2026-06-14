// Auto-update via electron-updater. Update metadata + artifacts are pulled from
// the project's GitHub Releases (configured under `build.publish` in package.json).
//
// Two install strategies, chosen by platform:
//   - Windows / Linux: electron-updater downloads the artifact and installs it on
//     restart (NSIS / AppImage), so we auto-download and post a click-to-restart
//     notification when it's ready.
//   - macOS: Squirrel.Mac only applies updates to a code-signed + notarized build.
//     Our macOS build is ad-hoc-signed, so quitAndInstall() would fail. Instead we
//     DON'T download — we just detect the new version (a cheap metadata fetch) and
//     open the GitHub release page so the user can grab the .dmg and drag it in.
//
// This degrades gracefully: in development (unpackaged) there are no artifacts so
// it no-ops, and if electron-updater is missing it's simply disabled.
let autoUpdater = null
try {
  ;({ autoUpdater } = require('electron-updater'))
} catch {
  // electron-updater unavailable — auto-update disabled.
}

const { dialog, shell, Notification } = require('electron')

const SIX_HOURS_MS = 6 * 60 * 60 * 1000
const RELEASES_URL = 'https://github.com/MailStudio/MailStudio/releases/latest'

// macOS can't silently self-install an unsigned build, so it gets the
// notify-and-link flow instead of download-and-install.
const MANUAL_INSTALL = process.platform === 'darwin'

let appRef = null
// Versions we've already surfaced this session, so the 6-hour re-check doesn't
// nag for a release the user has already been told about.
const promptedVersions = new Set()
// True while a user-initiated check is running, so we can report "up to date" or
// a failure — feedback the silent periodic check intentionally stays quiet about.
let manualCheckPending = false

function isSupported(app) {
  const a = app || appRef
  return Boolean(autoUpdater && a && a.isPackaged)
}

function log(err) {
  console.error('[updater]', err && err.message ? err.message : err)
}

function openReleasesPage() {
  shell.openExternal(RELEASES_URL).catch(log)
}

// macOS: a newer version exists. Offer to open the download page rather than
// attempting a Squirrel.Mac install that an unsigned build can't complete.
function promptManualUpdate(info) {
  const version = info && info.version ? info.version : null
  if (version && promptedVersions.has(version)) return
  if (version) promptedVersions.add(version)
  dialog
    .showMessageBox({
      type: 'info',
      buttons: ['Download', 'Later'],
      defaultId: 0,
      cancelId: 1,
      title: 'Update Available',
      message: version ? `MailStudio ${version} is available.` : 'A MailStudio update is available.',
      detail: 'Opening the download page in your browser. Download the new version, then drag MailStudio into your Applications folder to replace this copy.'
    })
    .then(({ response }) => {
      if (response === 0) openReleasesPage()
    })
    .catch(log)
}

// Windows / Linux: the artifact is downloaded and ready — prompt a restart.
function notifyDownloaded(info) {
  if (!Notification.isSupported()) return
  const version = info && info.version ? info.version : null
  const name = version ? `MailStudio ${version}` : 'A MailStudio update'
  const n = new Notification({
    title: 'Update Ready',
    body: `${name} will be installed when you restart. Click to restart now.`
  })
  n.on('click', () => {
    try {
      autoUpdater.quitAndInstall()
    } catch (err) {
      log(err)
    }
  })
  n.show()
}

// Wire up periodic checks. Call once after the app is ready.
function init({ app } = {}) {
  appRef = app || null
  if (!isSupported()) return

  autoUpdater.autoDownload = !MANUAL_INSTALL
  autoUpdater.autoInstallOnAppQuit = !MANUAL_INSTALL

  autoUpdater.on('error', (err) => {
    if (manualCheckPending) {
      manualCheckPending = false
      dialog
        .showMessageBox({
          type: 'warning',
          buttons: ['OK'],
          title: 'Update Check Failed',
          message: 'Could not check for updates.',
          detail: (err && err.message) || String(err)
        })
        .catch(() => {})
    }
    log(err)
  })

  autoUpdater.on('update-not-available', () => {
    if (!manualCheckPending) return
    manualCheckPending = false
    dialog
      .showMessageBox({
        type: 'info',
        buttons: ['OK'],
        title: 'No Updates',
        message: 'You\u2019re up to date.',
        detail: `MailStudio ${appRef.getVersion()} is the latest version.`
      })
      .catch(() => {})
  })

  if (MANUAL_INSTALL) {
    autoUpdater.on('update-available', (info) => {
      manualCheckPending = false
      promptManualUpdate(info)
    })
  } else {
    autoUpdater.on('update-downloaded', (info) => {
      manualCheckPending = false
      notifyDownloaded(info)
    })
  }

  check()
  const timer = setInterval(check, SIX_HOURS_MS)
  if (typeof timer.unref === 'function') timer.unref()
}

// Check for an update. `manual: true` (from the "Check for Updates…" menu item)
// surfaces "up to date"/error feedback the silent periodic check suppresses, and
// falls back to opening the releases page when updates can't be checked in-app
// (e.g. an unpackaged dev build).
function check({ manual = false } = {}) {
  if (!isSupported()) {
    if (manual) openReleasesPage()
    return
  }
  if (manual) manualCheckPending = true
  try {
    autoUpdater.checkForUpdates()?.catch((e) => {
      if (manual) manualCheckPending = false
      log(e)
    })
  } catch (err) {
    if (manual) manualCheckPending = false
    log(err)
  }
}

module.exports = { init, check, isSupported }
