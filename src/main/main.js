const path = require('path')
const os = require('os')
const {
  app,
  BrowserView,
  BrowserWindow,
  ipcMain,
  Menu,
  Notification,
  Tray,
  nativeImage,
  session,
  shell
} = require('electron')
const store = require('./settings-store')
const connections = require('./connections')
const updater = require('./updater')

// macOS 26 (Tahoe) ships as Darwin 26+. On that OS we enable native Liquid
// Glass vibrancy instead of a solid window background.
const IS_GLASS_MODE = process.platform === 'darwin' && parseInt(os.release().split('.')[0], 10) >= 26

const APP_NAME = 'MailStudio'
const MICROSOFT_SESSION_PARTITION = 'persist:mailstudio'
const ASANA_SESSION_PARTITION = 'persist:mailstudio-asana'
const WINDOW_SIZE = { width: 1280, height: 860, minWidth: 720, minHeight: 480 }
const SIDEBAR = { expanded: 280, rail: 76 }
const TOPBAR_HEIGHT = 46
const MENU_SIZE = { width: 328, height: 460 }
const FEED_REFRESH_MS = 25000
// Background view lifecycle: prewarm hidden views so every tab's notifications
// and sidebar stay live (not just the first tab opened), then hibernate idle
// views whose feed is API-backed (or which have no feed) to cap memory. Scrape-
// feed and Teams views are never hibernated — they ARE the notification source
// and must stay resident.
const MAX_LOADED_VIEWS = 6             // cap on simultaneously-resident web views
const PREWARM_BOOT_DELAY_MS = 3000     // let the active tab settle before prewarming
const PREWARM_SETTLE_MS = 5000         // gap between staggered prewarm loads
const PREWARM_TIMEOUT_MS = 20000       // give up on a stalled load, move to next
const HIDDEN_SCRAPE_MS = 50000         // min gap between scrapes of a hidden view
const HIBERNATE_IDLE_MS = 15 * 60 * 1000 // idle time before an eligible view sleeps
const REAPER_MS = 60000                // hibernation sweep interval
const REPO_URL = 'https://github.com/MailStudio/MailStudio'

// Derived at runtime from the real Chromium UA so the platform token and the
// Chrome version stay correct as Electron updates. Only the Electron and
// app-name tokens are stripped — they make Microsoft/Google treat the app as
// an embedded browser and block sign-in.
let appUserAgent = null
function getAppUserAgent() {
  if (!appUserAgent) {
    appUserAgent = session.defaultSession
      .getUserAgent()
      .replace(/\sElectron\/\S+/i, '')
      .replace(/\s(mailstudio|MailStudio|MailStudio)\/\S+/i, '')
      .replace(/\s{2,}/g, ' ')
  }
  return appUserAgent
}

// Trusted base domains that may load in-app, matched by suffix so first-party
// subdomains (login.microsoftonline.com, teams.microsoft.com, *.cloud.microsoft,
// etc.) are covered without an exhaustive list. Custom pinned sites are NOT
// granted here — they're allowed only by their exact host (allowedExactHosts).
const TRUSTED_BASE_DOMAINS = [
  'office.com',
  'office365.com',
  'microsoft.com',
  'microsoftonline.com',
  'live.com',
  'cloud.microsoft',
  'sharepoint.com',
  'onenote.com',
  'planner.microsoft.com',
  '1drv.ms',
  'asana.com',
  'asanausercontent.com'
]

let settings = store.normalize(null)
let allowedExactHosts = new Set()
const configuredPartitions = new Set()

// Open links externally only for safe schemes — never let a page hand us a
// file:, javascript:, smb:, or custom-scheme URL to launch.
function openExternalSafe(target) {
  try {
    const protocol = new URL(target).protocol
    if (protocol === 'http:' || protocol === 'https:' || protocol === 'mailto:') {
      shell.openExternal(target)
    }
  } catch {
    /* ignore malformed URLs */
  }
}

const serviceViews = new Map()
const loadedServiceKeys = new Set()
const serviceState = {}
const serviceFeeds = {}
// key → ms timestamp this view was last the active/split pane (drives hibernation)
const lastActiveAt = {}
// key → ms timestamp of the last scrape (throttles hidden-view background scrapes)
const lastScrapeAt = {}
// Staggered prewarm queue: keys waiting to load one-at-a-time in the background so
// the app never spawns every renderer at once (the "too many windows" crash).
const prewarmQueue = []
let prewarmActive = false
let reaperTimer = null

let activeServiceKey = 'mail'
// Split view: when this holds exactly two service keys, both are shown side by
// side (splitKeys[0] = left pane, splitKeys[1] = right). Empty array means the
// normal single-view mode keyed by activeServiceKey. Built via Cmd/Ctrl-click in
// the sidebar; activeServiceKey stays one of the two as the "focused" pane that
// nav/compose/feed actions target.
let splitKeys = []
// Split layout: orientation of the divider and the fraction of the content area
// the first pane (left or top) gets. 'vertical' = side by side (vertical
// divider); 'horizontal' = stacked (horizontal divider). In-memory like
// splitKeys — these reset on relaunch. splitDragging is true while the user
// drags the divider, during which the web views are detached so the renderer
// can track the mouse and paint a live preview.
let splitOrientation = 'vertical'
let splitRatio = 0.5
let splitDragging = false
const SPLIT_GUTTER = 8
const SPLIT_RATIO_MIN = 0.15
const SPLIT_RATIO_MAX = 0.85
let sidebarCollapsed = false
let settingsOpen = false
// True while the onboarding/connections sheet is open. Like firstBoot, this
// detaches the service BrowserView so the centered sheet (DOM) isn't hidden
// behind an on-top web view.
let onboardingOpen = false
let tray = null
let panelWindow = null
let menuWindow = null
let feedTimer = null
let lastMailNotificationCount = 0
let lastTeamsNotificationCount = 0
let teamsNotifReady = false
const notifiedEmailIds = new Set()
let mailNotifReady = false
// Asana: ids already seen, so only genuinely new assigned tasks notify.
const notifiedTaskIds = new Set()
let asanaNotifReady = false
// Calendar: event ids already reminded, so each upcoming event fires once.
const remindedEventIds = new Set()
// Per-service snooze: serviceKey → expiresAtMs
const serviceSnooze = new Map()
// True after any Microsoft service successfully authenticates (one-time pre-warm)
let servicesPrewarmed = false
// First-boot welcome/SSO screen: when active the BrowserView is detached so the
// renderer owns the whole content area.
let firstBoot = false
// Set when the user chooses "Sign in" so the next authenticated MS navigation
// triggers a background pre-warm of the other Microsoft tabs.
let wantPrewarm = false

/* ---------- Service helpers ---------- */
function getServices() {
  return settings.services
}

function visibleServices() {
  return settings.services.filter((service) => service.visible)
}

function findService(key) {
  return settings.services.find((service) => service.key === key)
}

function ensureServiceState(service) {
  if (!serviceState[service.key]) {
    serviceState[service.key] = {
      key: service.key,
      label: service.label,
      title: service.label,
      unreadCount: 0,
      href: service.url,
      visible: false
    }
  } else {
    serviceState[service.key].label = service.label
  }
  if (!serviceFeeds[service.key] && service.feed) {
    // gen is bumped whenever the feed's source of truth changes (API connect/
    // disconnect) so in-flight polls from the old source can be discarded.
    serviceFeeds[service.key] = { kind: service.feed, state: 'loading', items: [], gen: 0 }
  }
}

function buildAllowedHosts() {
  // Exact hosts for pinned sites (built-ins are covered by TRUSTED_BASE_DOMAINS).
  const hosts = new Set()
  for (const service of settings.services) {
    for (const value of [service.url, service.home]) {
      try {
        hosts.add(new URL(value).hostname)
      } catch {
        /* ignore */
      }
    }
  }
  allowedExactHosts = hosts
}

function isTrustedDomain(host) {
  return TRUSTED_BASE_DOMAINS.some((base) => host === base || host.endsWith(`.${base}`))
}

function isSnoozed(serviceKey) {
  const exp = serviceSnooze.get(serviceKey)
  if (!exp) return false
  if (Date.now() >= exp) { serviceSnooze.delete(serviceKey); return false }
  return true
}

// Services that fire notifications and can therefore be snoozed individually.
const SNOOZABLE_SERVICES = ['mail', 'teams', 'calendar', 'asana']

// Mute/unmute a tab's in-page audio to match its snooze state. Snoozing already
// suppresses our native Notification sounds (isSnoozed guards fireNotification);
// this additionally silences sounds the web app plays itself (e.g. the Teams or
// OWA new-message chime). setAudioMuted is whole-tab, so a snoozed Teams loses
// call audio too — acceptable while snoozed.
function applyServiceMute(serviceKey) {
  const view = serviceViews.get(serviceKey)
  if (view && !view.webContents.isDestroyed()) {
    view.webContents.setAudioMuted(isSnoozed(serviceKey))
  }
}

// Single entry point for changing a service's snooze: updates the map, mutes the
// tab, and refreshes the UI. expiresAtMs <= 0 resumes (clears) the snooze.
function setSnooze(serviceKey, expiresAtMs) {
  if (expiresAtMs > 0) serviceSnooze.set(serviceKey, expiresAtMs)
  else serviceSnooze.delete(serviceKey)
  applyServiceMute(serviceKey)
  pushSnapshot()
}

// True when every snoozable, visible service is currently snoozed.
function isGlobalSnoozed() {
  const keys = visibleServices().map((s) => s.key).filter((k) => SNOOZABLE_SERVICES.includes(k))
  return keys.length > 0 && keys.every((k) => isSnoozed(k))
}

// Global snooze: snooze every snoozable service and mute every tab's audio.
// expiresAtMs <= 0 resumes all and unmutes all.
function setGlobalSnooze(expiresAtMs) {
  for (const service of getServices()) {
    if (SNOOZABLE_SERVICES.includes(service.key)) {
      if (expiresAtMs > 0) serviceSnooze.set(service.key, expiresAtMs)
      else serviceSnooze.delete(service.key)
    }
  }
  // Mute reflects each service's resulting snooze state across all open views.
  for (const key of serviceViews.keys()) applyServiceMute(key)
  pushSnapshot()
}

// Each script resolves to true only when it actually clicked a compose
// control, so compose() can fall back to the deep-link URL. The Asana one is
// async: open the omnibutton "Create" menu, then pick its "Task" entry.
const COMPOSE_SCRIPTS = {
  mail: `(() => { try { const el = document.querySelector('[data-icon-name="NewMail"],[aria-label="New message"],[aria-label*="New email" i],[title*="New mail" i],button[aria-label*="New mail" i]'); if (!el) return false; el.click(); return true } catch (e) { return false } })()`,
  calendar: `(() => { try { const el = document.querySelector('[data-icon-name="CalculatorAddition"],[aria-label="New event"],button[aria-label*="New event" i]'); if (!el) return false; el.click(); return true } catch (e) { return false } })()`,
  todo: `(() => { try { const el = document.querySelector('[aria-label="Add to-do"],button[aria-label*="New task" i],[data-automationid*="NewTask"]'); if (!el) return false; el.click(); return true } catch (e) { return false } })()`,
  asana: `(() => { try {
    const btn = document.querySelector('[aria-label="Create"],[data-testid="omnibutton"],.Omnibutton,[class*="Omnibutton" i] button');
    if (!btn) return false;
    btn.click();
    return new Promise((resolve) => setTimeout(() => {
      try {
        const entry = Array.from(document.querySelectorAll('[role="menuitem"],[class*="MenuItem" i]'))
          .find((el) => /^task\\b/i.test((el.textContent || '').trim()));
        if (entry) entry.click();
        resolve(true);
      } catch (err) { resolve(true); }
    }, 350));
  } catch (e) { return false } })()`
}

const COMPOSE_DEEPLINK = {
  mail: 'https://outlook.office.com/mail/deeplink/compose',
  calendar: 'https://outlook.office.com/calendar/deeplink/compose',
  todo: 'https://to-do.office.com/tasks/'
}

// Open a compose surface in the owning tab. 'todo' follows the Tasks setting
// (Microsoft To Do by default, Asana when selected — falling back to whichever
// task tab is visible). Try the in-page "new" button first; if it can't be
// found, fall back to the compose deep-link URL.
function compose(kind) {
  let target = kind
  if (kind === 'todo') {
    const preferred = settings.taskProvider === 'asana' ? 'asana' : 'todo'
    const fallback = preferred === 'asana' ? 'todo' : 'asana'
    const preferredService = findService(preferred)
    target = preferredService && preferredService.visible ? preferred : fallback
  }
  const service = findService(target)
  if (!service || !service.visible) return
  showPanelWindow()
  showService(target)
  const view = serviceViews.get(target)
  if (!view || view.webContents.isDestroyed()) return
  const script = COMPOSE_SCRIPTS[target]
  if (!script) return
  setTimeout(() => {
    view.webContents
      .executeJavaScript(script, true)
      .then((ok) => {
        if (!ok && COMPOSE_DEEPLINK[target]) {
          view.webContents.loadURL(COMPOSE_DEEPLINK[target])
        }
      })
      .catch(() => {
        if (COMPOSE_DEEPLINK[target]) view.webContents.loadURL(COMPOSE_DEEPLINK[target])
      })
  }, 250)
}

// Push a one-off event (not a full snapshot) to the panel renderer — used for
// find-in-page open/results and other transient UI signals.
function sendPanelEvent(data) {
  if (panelWindow && !panelWindow.isDestroyed()) {
    showPanelWindow()
    panelWindow.webContents.send('panel:event', data)
  }
}

function isMicrosoftLoginHost(host) {
  return (
    host === 'login.microsoftonline.com' ||
    host === 'login.live.com' ||
    host === 'login.microsoft.com' ||
    host.endsWith('.login.microsoftonline.com')
  )
}

// Once one Microsoft tab is authenticated, eagerly load the other VISIBLE
// built-in Microsoft tabs in the background. They share the Microsoft session
// partition, so silent SSO signs them all in without a second prompt. Hidden
// Office-suite tabs are left lazy — they'll SSO silently whenever first opened.
function prewarmMicrosoftServices() {
  if (servicesPrewarmed) return
  servicesPrewarmed = true
  // After MS auth, spread the session to the other Microsoft tabs — but through
  // the staggered queue so they don't all spin up renderers at once.
  const keys = []
  for (const service of settings.services) {
    if (service.builtin && service.visible && service.key !== 'asana' && !loadedServiceKeys.has(service.key)) {
      keys.push(service.key)
    }
  }
  enqueuePrewarm(keys)
}

// A view needs to stay resident because it is the live source of its
// notifications: Teams (unread parsed from the tab title) and any feed still
// served by DOM scrape (no API connected). API-backed feeds keep polling without
// a view, so those are safe to hibernate.
function needsLiveView(key) {
  if (key === 'teams') return true
  const feed = serviceFeeds[key]
  if (feed && !connections.feedIsLive(feed.kind)) return true
  return false
}

// Recreate a view that was hibernated (or never built), reusing the retained
// serviceState/serviceFeeds cache. No-op if a live view already exists.
function ensureServiceView(key) {
  const existing = serviceViews.get(key)
  if (existing && !existing.webContents.isDestroyed()) return existing
  const service = findService(key)
  if (!service) return null
  return createServiceView(service)
}

// Tear down an idle view's renderer to reclaim memory while keeping its cached
// feed items and state so the sidebar still shows them. The view is rebuilt
// lazily (ensureServiceView) the next time the tab is shown. Unlike
// destroyServiceView this preserves serviceState/serviceFeeds.
function hibernateServiceView(key) {
  if (layoutKeys().includes(key)) return // never sleep an on-screen pane
  const view = serviceViews.get(key)
  if (!view) return
  try {
    view.webContents.destroy()
  } catch {
    /* ignore */
  }
  serviceViews.delete(key)
  loadedServiceKeys.delete(key)
  delete lastScrapeAt[key]
  if (serviceState[key]) serviceState[key].visible = false
}

// Queue keys for staggered background loading and start the pump.
function enqueuePrewarm(keys) {
  for (const key of keys) {
    if (loadedServiceKeys.has(key)) continue
    if (prewarmQueue.includes(key)) continue
    prewarmQueue.push(key)
  }
  pumpPrewarm()
}

// Load the next queued view, one at a time. Non-essential (hibernatable) views
// respect MAX_LOADED_VIEWS; essential notification-source views bypass the cap so
// their notifications always work. Stalled loads time out so the queue drains.
function pumpPrewarm() {
  if (prewarmActive) return
  while (prewarmQueue.length && (loadedServiceKeys.has(prewarmQueue[0]) || !serviceViews.has(prewarmQueue[0]))) {
    prewarmQueue.shift()
  }
  if (!prewarmQueue.length) return
  let idx = 0
  if (loadedServiceKeys.size >= MAX_LOADED_VIEWS) {
    idx = prewarmQueue.findIndex((key) => needsLiveView(key))
    if (idx === -1) return // only non-essential keys left — wait for a freed slot
  }
  const key = prewarmQueue.splice(idx, 1)[0]
  const view = serviceViews.get(key)
  const service = findService(key)
  if (!view || !service || view.webContents.isDestroyed()) {
    pumpPrewarm()
    return
  }
  prewarmActive = true
  let advanced = false
  const advance = () => {
    if (advanced) return
    advanced = true
    prewarmActive = false
    setTimeout(pumpPrewarm, PREWARM_SETTLE_MS)
  }
  view.webContents.once('did-finish-load', advance)
  ensureServiceLoaded(key)
  setTimeout(advance, PREWARM_TIMEOUT_MS)
}

// Boot-time prewarm: load notification-source views first (so Teams and any
// scrape feed start firing without being clicked), then the rest for sidebar
// fill, all staggered through the queue.
function startPrewarmQueue() {
  const essential = []
  const rest = []
  for (const service of visibleServices()) {
    if (loadedServiceKeys.has(service.key)) continue
    if (needsLiveView(service.key)) essential.push(service.key)
    else rest.push(service.key)
  }
  enqueuePrewarm([...essential, ...rest])
}

// Hibernate views that are off-screen, idle past the threshold, and safe to
// sleep (API-backed feed or no feed). Frees a slot, then lets any deferred
// prewarm proceed.
function reapHibernation() {
  if (firstBoot || onboardingOpen) return
  const now = Date.now()
  const attached = new Set(layoutKeys())
  for (const key of [...loadedServiceKeys]) {
    if (attached.has(key)) continue
    if (needsLiveView(key)) continue
    if (now - (lastActiveAt[key] || 0) < HIBERNATE_IDLE_MS) continue
    hibernateServiceView(key)
  }
  pumpPrewarm()
}

function startReaperTimer() {
  if (reaperTimer) clearInterval(reaperTimer)
  reaperTimer = setInterval(reapHibernation, REAPER_MS)
}

function finishFirstBoot() {
  if (firstBoot) {
    firstBoot = false
    settings = store.save({ ...settings, firstBoot: false })
  }
}

function isQuietHours() {
  const { quietStart, quietEnd } = (settings.notif || {})
  if (!quietStart || !quietEnd) return false
  const [sh, sm = 0] = quietStart.split(':').map(Number)
  const [eh, em = 0] = quietEnd.split(':').map(Number)
  const now = new Date()
  const nowMins = now.getHours() * 60 + now.getMinutes()
  const startMins = sh * 60 + sm
  const endMins = eh * 60 + em
  // Handle overnight spans (e.g. 22:00–08:00)
  return startMins <= endMins
    ? nowMins >= startMins && nowMins < endMins
    : nowMins >= startMins || nowMins < endMins
}

function fireNotification(opts, onClick) {
  if (!Notification.isSupported()) return
  const n = new Notification(opts)
  if (onClick) n.on('click', onClick)
  n.show()
}

function diffAndNotifyMail(items, unreadCount) {
  if (unreadCount === 0) {
    notifiedEmailIds.clear()
    lastMailNotificationCount = 0
    mailNotifReady = false
    return
  }

  if (!mailNotifReady) {
    // First scrape after launch: mark all current emails as seen without notifying.
    for (const item of items) { if (item.id) notifiedEmailIds.add(item.id) }
    capBaselineSet(notifiedEmailIds)
    lastMailNotificationCount = unreadCount
    mailNotifReady = true
    return
  }

  const notif = settings.notif || {}
  const countWentUp = unreadCount > lastMailNotificationCount
  lastMailNotificationCount = unreadCount

  if (!countWentUp) return
  if (!settings.onboarded) return
  if (!notif.mail) return
  if (isQuietHours()) return
  if (isSnoozed('mail')) return
  if (panelWindow && panelWindow.isFocused() && serviceFeeds[activeServiceKey]?.kind === 'mail') return

  const newItems = items.filter((item) => item.id && !notifiedEmailIds.has(item.id))

  // Mark all visible emails as seen to prevent re-notification next cycle.
  for (const item of items) { if (item.id) notifiedEmailIds.add(item.id) }
  capBaselineSet(notifiedEmailIds)

  if (newItems.length === 0) {
    // Count went up but scraper found no new IDs — fire a generic fallback.
    fireNotification(
      { title: APP_NAME, body: 'New email in your inbox.' },
      () => { showPanelWindow(); showService('mail') }
    )
    return
  }

  if (newItems.length >= 4) {
    const senders = [...new Set(newItems.map((i) => i.sender).filter(Boolean))].slice(0, 4)
    fireNotification(
      { title: `${newItems.length} new emails`, body: senders.join(', ') },
      () => { showPanelWindow(); showService('mail') }
    )
    return
  }

  for (let i = 0; i < newItems.length; i++) {
    const item = newItems[i]
    const subject = item.subject || '(no subject)'
    const body = notif.preview && item.preview ? `${subject}\n${item.preview}` : subject
    const rowIdx = item.rowIdx
    fireNotification(
      {
        title: item.sender || 'New email',
        body,
        silent: i > 0 // only the first notification in a batch plays a sound
      },
      () => {
        showPanelWindow()
        showService('mail')
        const view = serviceViews.get('mail')
        if (view && !view.webContents.isDestroyed() && typeof rowIdx === 'number') {
          const idx = Math.trunc(rowIdx)
          setTimeout(() => {
            view.webContents
              .executeJavaScript(
                `(() => { const rows = Array.from(document.querySelectorAll('div[role="option"], div[role="listitem"]')); const r = rows[${idx}]; if (r) { r.click(); r.scrollIntoView({ block: 'nearest' }); } })()`,
                true
              )
              .catch(() => {})
          }, 400)
        }
      }
    )
  }
}

function maybeNotifyTeams(unreadCount) {
  // Teams is a SPA whose title transiently drops the "(N)" prefix during
  // in-app navigation, parsing as 0 — never rebase the baseline downward on a
  // 0 reading or every flap re-notifies the same backlog.
  if (unreadCount === 0) return
  if (!teamsNotifReady) {
    // First real count after launch establishes the baseline silently.
    teamsNotifReady = true
    lastTeamsNotificationCount = unreadCount
    return
  }
  const prev = lastTeamsNotificationCount
  lastTeamsNotificationCount = unreadCount
  if (unreadCount <= prev) return
  if (!settings.onboarded) return
  const notif = settings.notif || {}
  if (!notif.teams) return
  if (isQuietHours()) return
  if (isSnoozed('teams')) return
  if (panelWindow && panelWindow.isFocused() && activeServiceKey === 'teams') return
  const delta = unreadCount - prev
  fireNotification(
    {
      title: delta === 1 ? 'New Teams message' : `${delta} new Teams messages`,
      body: `${unreadCount} unread`
    },
    () => { showPanelWindow(); showService('teams') }
  )
}

// Asana: notify on newly assigned tasks. Mirrors the mail diff — the first poll
// after launch only establishes a baseline (no notifications for the backlog).
function diffAndNotifyAsana(items) {
  if (!asanaNotifReady) {
    for (const item of items) { if (item.id) notifiedTaskIds.add(item.id) }
    capBaselineSet(notifiedTaskIds)
    asanaNotifReady = true
    return
  }
  const newItems = items.filter((item) => item.id && !notifiedTaskIds.has(item.id))
  for (const item of items) { if (item.id) notifiedTaskIds.add(item.id) }
  capBaselineSet(notifiedTaskIds)
  if (!newItems.length) return
  if (!settings.onboarded) return
  const notif = settings.notif || {}
  if (!notif.asana) return
  if (isQuietHours()) return
  if (isSnoozed('asana')) return

  const open = () => { showPanelWindow(); showService('asana') }
  if (newItems.length === 1) {
    fireNotification({ title: 'New Asana task', body: newItems[0].name || 'Task assigned to you' }, open)
  } else {
    fireNotification(
      { title: `${newItems.length} new Asana tasks`, body: newItems.slice(0, 3).map((t) => t.name).filter(Boolean).join(', ') },
      open
    )
  }
}

// Baseline Sets live for the whole tray-resident session; cap them so weeks of
// uptime can't grow them unboundedly. Sets iterate in insertion order, so
// dropping from the front removes the oldest ids first.
const BASELINE_SET_MAX = 2000
function capBaselineSet(set) {
  while (set.size > BASELINE_SET_MAX) set.delete(set.values().next().value)
}

// Calendar: remind shortly before an event starts. Each event fires once, and
// only within the reminder window so we don't surprise-notify the whole agenda.
const CALENDAR_REMINDER_MS = 5 * 60 * 1000
function maybeNotifyCalendar(items) {
  if (!settings.onboarded) return
  const notif = settings.notif || {}
  if (!notif.calendar) return
  if (isQuietHours()) return
  if (isSnoozed('calendar')) return
  const now = Date.now()
  for (const item of items) {
    if (!item.id || !item.startIso || remindedEventIds.has(item.id)) continue
    const start = new Date(String(item.startIso).replace(/(\.\d+)?$/, '')).getTime()
    if (Number.isNaN(start)) continue
    const delta = start - now
    if (delta <= 0 || delta > CALENDAR_REMINDER_MS) continue
    remindedEventIds.add(item.id)
    capBaselineSet(remindedEventIds)
    const mins = Math.max(1, Math.round(delta / 60000))
    fireNotification(
      { title: 'Upcoming event', body: `${item.title || 'Event'} starts in ${mins} min` },
      () => { showPanelWindow(); showService('calendar') }
    )
  }
}

// Scrape ids (DOM fingerprints) and API ids (Graph/Asana gids) never match, so
// whenever a provider's source of truth switches — connect or disconnect — the
// per-item notification baselines must reset. The next poll from the new source
// then re-baselines silently instead of notifying for the entire backlog.
function resetNotificationBaselines(provider) {
  const kinds = connections.PROVIDER_FEEDS[provider] || []
  if (kinds.includes('mail')) {
    notifiedEmailIds.clear()
    mailNotifReady = false
    lastMailNotificationCount = 0
  }
  if (kinds.includes('calendar')) {
    remindedEventIds.clear()
  }
  if (kinds.includes('asana')) {
    notifiedTaskIds.clear()
    asanaNotifReady = false
  }
}

// Hard cutover between sources: drop the provider's cached feed items along
// with the baselines so the next refresh repopulates entirely from the new
// source — API after connect (scrapers stay disabled while connected), scrape
// after disconnect — instead of stale items lingering until replaced.
function resetProviderFeeds(provider) {
  resetNotificationBaselines(provider)
  for (const key of Object.keys(serviceFeeds)) {
    if (connections.providerForFeed(serviceFeeds[key].kind) === provider) {
      serviceFeeds[key].state = 'loading'
      serviceFeeds[key].items = []
      // Invalidate any in-flight poll from the previous source: its result
      // would repopulate the feed and re-baseline notifications with ids the
      // new source will never produce (one spurious "N new emails" burst).
      serviceFeeds[key].gen = (serviceFeeds[key].gen || 0) + 1
    }
  }
}

function sidebarWidth() {
  // Settings live inside the sidebar, so it must be full width while open
  // (regardless of the collapse state) — and the web view stays visible.
  if (settingsOpen) {
    return SIDEBAR.expanded
  }
  if (!sidebarCollapsed) {
    return SIDEBAR.expanded
  }
  return settings.collapseMode === 'rail' ? SIDEBAR.rail : 0
}

/* ---------- Navigation ---------- */
function canGoBack(webContents) {
  try {
    return webContents.navigationHistory.canGoBack()
  } catch {
    return typeof webContents.canGoBack === 'function' ? webContents.canGoBack() : false
  }
}

function canGoForward(webContents) {
  try {
    return webContents.navigationHistory.canGoForward()
  } catch {
    return typeof webContents.canGoForward === 'function' ? webContents.canGoForward() : false
  }
}

function navigate(direction) {
  const webContents = serviceViews.get(activeServiceKey)?.webContents
  if (!webContents) {
    return
  }

  if (direction === 'reload') {
    webContents.reload()
    return
  }

  const history = webContents.navigationHistory
  if (direction === 'back') {
    if (history && typeof history.goBack === 'function') {
      if (history.canGoBack()) history.goBack()
    } else if (webContents.canGoBack()) {
      webContents.goBack()
    }
  } else if (direction === 'forward') {
    if (history && typeof history.goForward === 'function') {
      if (history.canGoForward()) history.goForward()
    } else if (webContents.canGoForward()) {
      webContents.goForward()
    }
  }
}

function goHome() {
  const service = findService(activeServiceKey)
  const webContents = serviceViews.get(activeServiceKey)?.webContents
  if (service && webContents) {
    loadedServiceKeys.add(activeServiceKey)
    webContents.loadURL(service.home || service.url)
  }
}

// Load a service's URL on first activation so that sign-in on one tab sets
// the SSO cookie before the next tab loads, letting Microsoft's silent auth
// carry the session across without requiring a second sign-in.
function ensureServiceLoaded(key) {
  if (loadedServiceKeys.has(key)) return
  const view = serviceViews.get(key)
  const service = findService(key)
  if (view && service && !view.webContents.isDestroyed()) {
    view.webContents.loadURL(service.url)
    loadedServiceKeys.add(key)
  }
}

function parseUnreadCount(title) {
  const match = title.match(/^\((\d+)\)/)
  return match ? Number.parseInt(match[1], 10) : 0
}

function isAllowedHost(urlString) {
  try {
    const host = new URL(urlString).hostname
    return isTrustedDomain(host) || allowedExactHosts.has(host)
  } catch {
    return false
  }
}

// Main-process loadURL() bypasses the will-navigate hook that enforces the
// allowlist for page navigations, so renderer-supplied deep links (feed item
// webLink/taskUrl) must be validated here before they can steer a view that
// carries the authenticated session. Disallowed links open externally.
function loadInServiceView(view, serviceKey, url) {
  let allowed = false
  try {
    const protocol = new URL(url).protocol
    allowed = (protocol === 'http:' || protocol === 'https:') && isAllowedHost(url)
  } catch {
    allowed = false
  }
  if (!allowed) {
    openExternalSafe(url)
    return
  }
  try {
    loadedServiceKeys.add(serviceKey)
    view.webContents.loadURL(url)
  } catch {
    /* ignore */
  }
}

// Maps a URL to the core-suite tab that owns it, recognising the various
// Microsoft host variants so a deep link from Teams (or anywhere) lands on the
// right tab. Only visible tabs claim ownership — a hidden tab's links fall
// through to the in-app/current-view fallback instead of surfacing a tab the
// user removed from the sidebar. Returns the configured service object, or null.
function coreServiceForUrl(urlString) {
  let parsed
  try {
    parsed = new URL(urlString)
  } catch {
    return null
  }
  const host = parsed.hostname
  const route = parsed.pathname
  const find = (key) =>
    settings.services.find((service) => service.key === key && service.builtin && service.visible)

  const isOutlookHost =
    host === 'outlook.office.com' ||
    host === 'outlook.office365.com' ||
    host === 'outlook.live.com' ||
    host.endsWith('.outlook.office.com') ||
    host.endsWith('.outlook.com')

  if (isOutlookHost && /\/calendar/i.test(route)) return find('calendar')
  if (isOutlookHost) return find('mail') // /mail, /owa, deep links — Mail owns the rest
  if (host === 'to-do.office.com' || host === 'to-do.microsoft.com' || host.endsWith('.todo.microsoft.com')) {
    return find('todo')
  }
  if (host === 'app.asana.com' || host.endsWith('.asana.com')) return find('asana')
  if (
    host === 'teams.microsoft.com' ||
    host.endsWith('.teams.microsoft.com') ||
    host === 'teams.live.com' ||
    host === 'teams.cloud.microsoft' ||
    host.endsWith('.teams.cloud.microsoft')
  ) {
    return find('teams')
  }
  if (
    host === 'planner.microsoft.com' ||
    host.endsWith('.planner.microsoft.com') ||
    host === 'tasks.office.com' ||
    host === 'planner.cloud.microsoft'
  ) {
    return find('planner')
  }
  // Office documents live on SharePoint/OneDrive hosts; the /:w:/-style path
  // segment marks the app, so a shared doc link lands on its app tab first and
  // falls back to the storage tab (OneDrive/SharePoint) when that tab is hidden.
  if (host.endsWith('.sharepoint.com') || host === 'onedrive.live.com' || host === '1drv.ms') {
    const docMatch = route.match(/\/:([wxpo]):\//i)
    if (docMatch) {
      const appTab = { w: 'word', x: 'excel', p: 'powerpoint', o: 'onenote' }[docMatch[1].toLowerCase()]
      const found = find(appTab)
      if (found) return found
    }
    if (host === 'onedrive.live.com' || host === '1drv.ms' || host.endsWith('-my.sharepoint.com')) {
      return find('onedrive') || find('sharepoint')
    }
    return find('sharepoint')
  }
  if (host === 'onenote.com' || host.endsWith('.onenote.com')) return find('onenote')
  // office.com / Microsoft 365 home: /launch/<app> deep links go to the app tab,
  // anything else to the Office home tab.
  if (
    host === 'office.com' ||
    host === 'www.office.com' ||
    host === 'm365.cloud.microsoft' ||
    host === 'microsoft365.com' ||
    host === 'www.microsoft365.com'
  ) {
    const appMatch = route.match(/\/launch\/(word|excel|powerpoint|onenote|onedrive|sharepoint)/i)
    if (appMatch) {
      const found = find(appMatch[1].toLowerCase())
      if (found) return found
    }
    return find('office')
  }
  // New per-app hosts (word.cloud.microsoft and friends).
  const cloudApp = host.match(/^(word|excel|powerpoint|onenote|onedrive|planner)\.cloud\.microsoft$/)
  if (cloudApp) return find(cloudApp[1])
  return null
}

function resolveServiceByUrl(urlString) {
  const core = coreServiceForUrl(urlString)
  if (core) {
    return core
  }
  // Fall back to host + path-prefix match (covers custom pinned sites).
  try {
    const target = new URL(urlString)
    return (
      settings.services.find((service) => {
        if (!service.visible) return false
        const current = new URL(service.url)
        return target.hostname === current.hostname && target.pathname.startsWith(current.pathname)
      }) || null
    )
  } catch {
    return null
  }
}

// Silently move to the tab that owns this URL and, for core-suite tabs, load
// the exact deep-linked page into that tab's view.
function routeToService(targetService, url) {
  const view = serviceViews.get(targetService.key)
  if (view && targetService.builtin) {
    try {
      // Mark loaded BEFORE calling loadURL so ensureServiceLoaded (called from
      // attachServiceView) doesn't override this deep-link with the home URL.
      loadedServiceKeys.add(targetService.key)
      if (view.webContents.getURL() !== url) {
        view.webContents.loadURL(url)
      }
    } catch {
      /* ignore */
    }
  }
  showService(targetService.key)
}

// In-app-first link opening for clicks that originate outside the service web
// views (panel, dropdown, help links): the tab that owns the URL wins; any
// other trusted Microsoft/Asana page loads in the active tab (loadInServiceView
// re-checks the allowlist); everything else goes to the default browser.
function openLinkInApp(url) {
  const internalService = resolveServiceByUrl(url)
  if (internalService) {
    routeToService(internalService, url)
    return
  }
  const view = serviceViews.get(activeServiceKey)
  if (isAllowedHost(url) && view && !view.webContents.isDestroyed()) {
    loadInServiceView(view, activeServiceKey, url)
    showService(activeServiceKey)
    return
  }
  openExternalSafe(url)
}

function getTrayIconPath(unread) {
  return path.join(__dirname, '..', '..', 'assets', unread > 0 ? 'trayUnreadTemplate.png' : 'trayTemplate.png')
}

/* ---------- Menus ---------- */
function buildAppMenu() {
  // ⌘1–9 jump straight to the Nth visible tab.
  const visible = visibleServices()
  const tabAccelerators = visible.slice(0, 9).map((service, index) => ({
    label: `Show ${service.label}`,
    accelerator: `CmdOrCtrl+${index + 1}`,
    click: () => showService(service.key)
  }))

  const template = [
    {
      label: APP_NAME,
      submenu: [
        { label: 'Open MailStudio', click: () => togglePanelWindow(true) },
        { label: 'Check for Updates…', click: () => updater.check() },
        { type: 'separator' },
        { role: 'quit' }
      ]
    },
    {
      // Standard macOS editing shortcuts (Cmd+A/C/V/X/Z). Without this menu,
      // those keystrokes never reach the focused web view.
      label: 'Edit',
      submenu: [
        { role: 'undo' },
        { role: 'redo' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { role: 'pasteAndMatchStyle' },
        { role: 'delete' },
        { role: 'selectAll' }
      ]
    },
    {
      label: 'Compose',
      submenu: [
        { label: 'New Email', accelerator: 'CmdOrCtrl+N', click: () => compose('mail') },
        { label: 'New Event', accelerator: 'CmdOrCtrl+Shift+E', click: () => compose('calendar') },
        { label: 'New Task', accelerator: 'CmdOrCtrl+Shift+T', click: () => compose('todo') }
      ]
    },
    {
      label: 'Go',
      submenu: [
        {
          label: 'Find in Page…',
          accelerator: 'CmdOrCtrl+F',
          click: () => sendPanelEvent({ type: 'open-search' })
        },
        { type: 'separator' },
        { label: 'Back', accelerator: 'CmdOrCtrl+[', click: () => navigate('back') },
        { label: 'Forward', accelerator: 'CmdOrCtrl+]', click: () => navigate('forward') },
        { label: 'Home', accelerator: 'CmdOrCtrl+Shift+H', click: () => goHome() },
        ...(tabAccelerators.length ? [{ type: 'separator' }, ...tabAccelerators] : [])
      ]
    },
    {
      label: 'View',
      submenu: [
        { role: 'reload' },
        { role: 'forceReload' },
        { role: 'togglefullscreen' },
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' }
      ]
    },
    { label: 'Window', role: 'windowMenu' },
    {
      role: 'help',
      submenu: [
        { label: 'Report an Issue…', click: () => openExternalSafe(`${REPO_URL}/issues/new?template=bug_report.yml`) },
        { label: 'View on GitHub', click: () => openExternalSafe(REPO_URL) },
        { type: 'separator' },
        { label: `Version ${app.getVersion()} (beta)`, enabled: false }
      ]
    }
  ]

  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

// Microsoft services share one session so SSO carries across Outlook, Teams,
// Calendar, To Do, and Office. Asana is connected in the product, but it does
// not need Microsoft cookies, so it gets its own built-in partition. Each
// custom pinned site gets its own partition as well.
function partitionFor(service) {
  if (!service.builtin) return `persist:mailstudio-site-${service.key}`
  if (service.key === 'asana') return ASANA_SESSION_PARTITION
  return MICROSOFT_SESSION_PARTITION
}

function partitionForProvider(provider) {
  return provider === 'asana' ? ASANA_SESSION_PARTITION : MICROSOFT_SESSION_PARTITION
}

function configurePartition(partitionName) {
  if (configuredPartitions.has(partitionName)) {
    return
  }
  configuredPartitions.add(partitionName)

  const partitionSession = session.fromPartition(partitionName)

  partitionSession.setPermissionCheckHandler((_webContents, permission, requestingOrigin) => {
    if (!isAllowedHost(requestingOrigin)) {
      return false
    }
    return permission === 'notifications' || permission === 'clipboard-sanitized-write'
  })

  partitionSession.setPermissionRequestHandler((_webContents, permission, callback, details) => {
    if (!isAllowedHost(details.requestingUrl)) {
      callback(false)
      return
    }
    callback(permission === 'notifications' || permission === 'clipboard-sanitized-write')
  })

  partitionSession.setUserAgent(getAppUserAgent())
}

function configureSession() {
  configurePartition(MICROSOFT_SESSION_PARTITION)
  configurePartition(ASANA_SESSION_PARTITION)
}

/* ---------- Service views ---------- */
function inSplit() {
  return splitKeys.length === 2
}

// The service key(s) that should currently fill the content area: both panes in
// split mode, otherwise just the active tab. Filtered to keys that actually have
// a live view so a half-built split can never leave a pane dangling.
function layoutKeys() {
  if (inSplit()) {
    return splitKeys.filter((key) => serviceViews.has(key))
  }
  return serviceViews.has(activeServiceKey) ? [activeServiceKey] : []
}

// Remove every web view from the panel window. Used for the first-boot/onboarding
// overlays and as the base of each re-layout. Views are managed with add/remove
// (never setBrowserView) so single view and split view — which needs two views
// attached at once — share one code path.
function detachAllViews() {
  if (!panelWindow || panelWindow.isDestroyed()) {
    return
  }
  if (typeof panelWindow.getBrowserViews === 'function') {
    for (const view of panelWindow.getBrowserViews()) {
      panelWindow.removeBrowserView(view)
    }
  } else {
    panelWindow.setBrowserView(null)
  }
}

// Lay out the content area from the current state (single tab or two-pane split).
// Idempotent: callers fire it on every resize, sidebar toggle, and tab change.
function attachServiceView() {
  if (!panelWindow || panelWindow.isDestroyed()) {
    return
  }

  // First-boot welcome and the onboarding sheet own the whole content area —
  // keep the on-top web views off so the centered DOM is visible and clickable.
  // While the divider is being dragged, the views are also off so the renderer
  // can track the mouse across the whole content area and paint a live preview.
  if (firstBoot || onboardingOpen || splitDragging) {
    detachAllViews()
    return
  }

  // Recreate any hibernated view we're about to show BEFORE layoutKeys() filters
  // it out for lacking a live view, so a slept tab wakes instead of going blank.
  const intended = inSplit() ? splitKeys.slice() : (activeServiceKey ? [activeServiceKey] : [])
  for (const key of intended) {
    ensureServiceView(key)
  }

  const keys = layoutKeys()
  if (!keys.length) {
    detachAllViews()
    return
  }
  // Stamp these as freshly active so the reaper's idle clock restarts for them.
  for (const key of keys) {
    lastActiveAt[key] = Date.now()
  }
  const targetViews = keys.map((key) => serviceViews.get(key)).filter(Boolean)

  // Detach any view that shouldn't be on screen (the previous single tab, or the
  // other pane after leaving split) before sizing the survivors.
  if (typeof panelWindow.getBrowserViews === 'function') {
    for (const view of panelWindow.getBrowserViews()) {
      if (!targetViews.includes(view)) {
        panelWindow.removeBrowserView(view)
      }
    }
  }
  for (const key of keys) {
    ensureServiceLoaded(key)
  }
  for (const view of targetViews) {
    if (!panelWindow.getBrowserViews().includes(view)) {
      panelWindow.addBrowserView(view)
    }
  }

  const [width, height] = panelWindow.getContentSize()
  const offset = Math.round(sidebarWidth())
  const contentWidth = Math.max(1, Math.round(width) - offset)
  const contentHeight = Math.max(1, Math.round(height) - TOPBAR_HEIGHT)
  // Manual resize handling keeps bounds exact; autoResize stays OFF so a view
  // can't drift past the right edge (an over-wide view clips its own content
  // left/right — the Asana edge bug).
  if (targetViews.length === 1) {
    targetViews[0].setBounds({ x: offset, y: TOPBAR_HEIGHT, width: contentWidth, height: contentHeight })
    targetViews[0].setAutoResize({ width: false, height: false })
    return
  }

  // Split: two panes separated by a gutter that reveals the window background as
  // a draggable divider. The two sizes + gutter add up to exactly the content
  // span so neither pane overshoots. splitRatio is the first pane's fraction.
  const gutter = SPLIT_GUTTER
  const ratio = Math.min(SPLIT_RATIO_MAX, Math.max(SPLIT_RATIO_MIN, splitRatio))
  if (splitOrientation === 'horizontal') {
    // Stacked: first pane on top, second below, horizontal divider.
    const topHeight = Math.max(1, Math.round((contentHeight - gutter) * ratio))
    const bottomHeight = Math.max(1, contentHeight - gutter - topHeight)
    targetViews[0].setBounds({ x: offset, y: TOPBAR_HEIGHT, width: contentWidth, height: topHeight })
    targetViews[1].setBounds({ x: offset, y: TOPBAR_HEIGHT + topHeight + gutter, width: contentWidth, height: bottomHeight })
  } else {
    // Side by side: first pane on the left, second on the right, vertical divider.
    const leftWidth = Math.max(1, Math.round((contentWidth - gutter) * ratio))
    const rightWidth = Math.max(1, contentWidth - gutter - leftWidth)
    targetViews[0].setBounds({ x: offset, y: TOPBAR_HEIGHT, width: leftWidth, height: contentHeight })
    targetViews[1].setBounds({ x: offset + leftWidth + gutter, y: TOPBAR_HEIGHT, width: rightWidth, height: contentHeight })
  }
  targetViews[0].setAutoResize({ width: false, height: false })
  targetViews[1].setAutoResize({ width: false, height: false })
}

function createServiceView(service) {
  const partition = partitionFor(service)
  configurePartition(partition)

  const view = new BrowserView({
    webPreferences: {
      preload: path.join(__dirname, 'service-preload.js'),
      sandbox: true,
      contextIsolation: true,
      nodeIntegration: false,
      webSecurity: true,
      partition,
      additionalArguments: [`--service-key=${service.key}`]
    }
  })

  view.webContents.setUserAgent(getAppUserAgent())
  view.webContents.setMaxListeners(20)

  view.webContents.setWindowOpenHandler(({ url }) => {
    // Microsoft auth/consent popups (used by Teams, SharePoint, etc.) MUST open
    // as a real child window — denying them and kicking to the browser breaks
    // the in-app sign-in flow. The child inherits this view's session partition.
    let popupHost = ''
    try {
      popupHost = new URL(url).hostname
    } catch {
      /* ignore */
    }
    const isAuthPopup =
      isMicrosoftLoginHost(popupHost) ||
      popupHost === 'login.microsoft.com' ||
      popupHost.endsWith('.microsoftonline.com') ||
      popupHost === 'login.live.com' ||
      popupHost.endsWith('.b2clogin.com')
    if (isAuthPopup) {
      return {
        action: 'allow',
        overrideBrowserWindowOptions: {
          autoHideMenuBar: true,
          webPreferences: {
            partition: partitionFor(service),
            sandbox: true,
            contextIsolation: true,
            nodeIntegration: false,
            webSecurity: true
          }
        }
      }
    }
    // A popup/new-window link that belongs to a tab: open it there silently.
    const internalService = resolveServiceByUrl(url)
    if (internalService) {
      routeToService(internalService, url)
      return { action: 'deny' }
    }
    // Trusted Microsoft/Asana page with no owning tab: keep it in-app in the
    // tab that spawned it rather than bouncing to the default browser.
    if (isAllowedHost(url)) {
      loadInServiceView(view, service.key, url)
      return { action: 'deny' }
    }
    openExternalSafe(url)
    return { action: 'deny' }
  })

  view.webContents.on('will-navigate', (event, url) => {
    const internalService = resolveServiceByUrl(url)
    if (internalService) {
      // A link that belongs to a different tab → hand it off silently.
      if (internalService.key !== service.key) {
        event.preventDefault()
        routeToService(internalService, url)
      }
      return
    }
    if (!isAllowedHost(url)) {
      event.preventDefault()
      openExternalSafe(url)
    }
  })

  view.webContents.on('page-title-updated', (_event, title) => {
    const state = serviceState[service.key]
    if (!state) return
    const prevCount = state.unreadCount || 0
    state.title = title || service.label
    const titleCount = parseUnreadCount(state.title)
    // When Microsoft is connected, the Graph API owns the mail unread count —
    // don't let the OWA tab title clobber it. The Asana badge is always
    // feed-owned (task count from scrape or API; Asana titles carry no count).
    const feedOwnsCount =
      (service.key === 'mail' && connections.feedIsLive('mail')) || service.key === 'asana'
    if (!feedOwnsCount) {
      state.unreadCount = titleCount
    }

    // Mail: a rising title count means new email — refresh the feed promptly
    // (API or scrape) for fresh notification content, faster than the poll.
    if (serviceFeeds[service.key]?.kind === 'mail' && titleCount > prevCount) {
      setTimeout(() => refreshFeed(service.key), 1500)
    }
    // Teams: notify directly from title count (no DOM scrape needed).
    if (service.key === 'teams' && state.unreadCount !== prevCount) {
      maybeNotifyTeams(state.unreadCount)
    }
    pushSnapshot()
  })

  view.webContents.on('did-navigate', (_event, url) => {
    if (serviceState[service.key]) serviceState[service.key].href = url
    // Reached an authenticated Microsoft app page → spread the session to the
    // other Microsoft tabs via silent SSO.
    if (wantPrewarm && service.builtin && service.key !== 'asana') {
      try {
        const host = new URL(url).hostname
        if (isTrustedDomain(host) && !isMicrosoftLoginHost(host)) {
          wantPrewarm = false
          setTimeout(prewarmMicrosoftServices, 1200)
        }
      } catch {
        /* ignore */
      }
    }
    pushSnapshot()
  })

  view.webContents.on('did-navigate-in-page', (_event, url, isMainFrame) => {
    if (!isMainFrame) return
    if (serviceState[service.key]) serviceState[service.key].href = url
    pushSnapshot()
  })

  view.webContents.on('did-finish-load', () => {
    // A reload resets the webContents mute flag, so re-assert it if still snoozed.
    applyServiceMute(service.key)
    view.webContents.insertCSS(`
      ::-webkit-scrollbar { width: 11px; height: 11px; }
      ::-webkit-scrollbar-thumb { background: rgba(140, 150, 165, 0.45); border-radius: 999px; }
    `)
    if (serviceFeeds[service.key]) {
      // First scrape attempt — OWA/Asana render asynchronously so a short delay
      // is needed before the email rows/task cards exist in the DOM.
      setTimeout(() => refreshFeed(service.key), 2500)
      // Second attempt for slow connections or heavy pages; if the first already
      // populated the cache this is a no-op (cache rule: good data wins).
      setTimeout(() => refreshFeed(service.key), 8000)
    }
  })

  view.webContents.on('found-in-page', (_event, result) => {
    if (result.finalUpdate) {
      sendPanelEvent({ type: 'find-result', activeMatchOrdinal: result.activeMatchOrdinal, matches: result.matches })
    }
  })

  serviceViews.set(service.key, view)
  return view
}

function destroyServiceView(key) {
  const view = serviceViews.get(key)
  if (view) {
    try {
      view.webContents.destroy()
    } catch {
      /* ignore */
    }
    serviceViews.delete(key)
  }
  loadedServiceKeys.delete(key)
  delete serviceState[key]
  delete serviceFeeds[key]
}

function syncServiceViews() {
  buildAllowedHosts()

  for (const service of settings.services) {
    ensureServiceState(service)
    if (!serviceViews.has(service.key)) {
      createServiceView(service)
    }
  }

  // Tear down views for services no longer configured.
  for (const key of [...serviceViews.keys()]) {
    if (!findService(key)) {
      destroyServiceView(key)
    }
  }

  // Keep the active service valid and visible.
  const active = findService(activeServiceKey)
  if (!active || !active.visible) {
    const fallback = visibleServices()[0]
    activeServiceKey = fallback ? fallback.key : (settings.services[0] && settings.services[0].key)
  }

  // Drop split panes that were hidden or removed; collapse to single view when
  // fewer than two valid panes remain, and keep the focus on a surviving pane.
  if (splitKeys.length) {
    splitKeys = splitKeys.filter((key) => {
      const svc = findService(key)
      return svc && svc.visible && serviceViews.has(key)
    })
    if (splitKeys.length < 2) {
      splitKeys = []
    } else if (!splitKeys.includes(activeServiceKey)) {
      activeServiceKey = splitKeys[splitKeys.length - 1]
    }
  }
}

function showService(serviceKey, options = {}) {
  const { reveal = true } = options
  const service = findService(serviceKey)
  if (!service) {
    return
  }
  settingsOpen = false
  splitKeys = [] // a plain tab switch always drops back to a single view
  splitDragging = false // clear any stuck divider-drag so the view re-attaches
  activeServiceKey = serviceKey
  attachServiceView()
  if (reveal && panelWindow && !panelWindow.isVisible()) {
    showPanelWindow()
  }
  updateVisibleStates()
  pushSnapshot()
}

// Cmd/Ctrl-click handler: toggle a service into/out of the side-by-side split.
// Seeds from the active tab, so the first such click pairs the active tab with
// the clicked one. Capped at two panes ("two windows"); a third modifier-click
// walks the split by dropping the oldest pane.
function splitSelect(serviceKey) {
  const service = findService(serviceKey)
  if (!service || !service.visible) {
    return
  }

  const selection = inSplit() ? splitKeys.slice() : [activeServiceKey]
  const existing = selection.indexOf(serviceKey)

  if (existing !== -1) {
    // Toggling a current pane off → fall back to a single view of the other one.
    selection.splice(existing, 1)
    showService(selection[selection.length - 1] || serviceKey)
    return
  }

  selection.push(serviceKey)
  if (selection.length < 2) {
    // No active tab to pair with (degenerate) — just show the clicked one.
    showService(serviceKey)
    return
  }

  splitKeys = selection.slice(-2)
  settingsOpen = false
  splitDragging = false
  activeServiceKey = serviceKey // the just-clicked pane takes nav/compose focus
  attachServiceView()
  if (panelWindow && !panelWindow.isVisible()) {
    showPanelWindow()
  }
  updateVisibleStates()
  pushSnapshot()
}

// Ensure a service is shown as a split pane (without the toggle-off behaviour of
// splitSelect) and focus it. Used when Cmd/Ctrl-clicking a feed item so it opens
// beside the current view instead of replacing it. Pairs with the active tab, or
// adds into an existing split, walking off the oldest pane.
function splitOpen(serviceKey) {
  const service = findService(serviceKey)
  if (!service || !service.visible) {
    return
  }
  if (inSplit()) {
    if (!splitKeys.includes(serviceKey)) {
      splitKeys = [splitKeys[1], serviceKey]
    }
  } else if (activeServiceKey !== serviceKey) {
    splitKeys = [activeServiceKey, serviceKey]
  }
  // (activeServiceKey === serviceKey with no split: nothing to pair — stays single.)
  settingsOpen = false
  splitDragging = false
  activeServiceKey = serviceKey
  attachServiceView()
  if (panelWindow && !panelWindow.isVisible()) {
    showPanelWindow()
  }
  updateVisibleStates()
  pushSnapshot()
}

// Reveal a feed item's owning service — in a split pane when split is requested
// (Cmd/Ctrl-click), otherwise as the single active view.
function revealFeedTarget(serviceKey, split) {
  if (split) {
    splitOpen(serviceKey)
  } else {
    showService(serviceKey)
  }
}

function updateVisibleStates() {
  const onScreen = Boolean(panelWindow && panelWindow.isVisible() && !settingsOpen)
  const shownKeys = new Set(inSplit() ? splitKeys : [activeServiceKey])
  for (const service of settings.services) {
    if (serviceState[service.key]) {
      serviceState[service.key].visible = onScreen && shownKeys.has(service.key)
    }
  }
}

/* ---------- Live feeds (scraped from logged-in web views) ---------- */

// rowIdx is the position in the full querySelectorAll result so we can click
// the right row back when the user taps a mail item in the sidebar.
const MAIL_SCRAPE = `(() => {
  try {
    // Not signed in: the view is parked on a Microsoft login page. Report it
    // so the sidebar shows "Sign in to continue" instead of an empty inbox.
    if (/^(login|account)\\./i.test(location.hostname) || document.querySelector('input[name="loginfmt"]')) {
      return { state: 'login', items: [] };
    }
    // OWA mixes icon-font glyphs (Private Use Area) and zero-width characters
    // into row text — they render as □ squares, so strip them everywhere.
    const strip = (s) => s.replace(/[\\u200B-\\u200D\\u2060\\uFEFF\\uFFFD\\uE000-\\uF8FF]/g, '').trim();
    const rows = Array.from(document.querySelectorAll('div[role="option"], div[role="listitem"], [data-convid]'));
    const items = [];
    const seen = new Set();
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const label = strip(row.getAttribute('aria-label') || '');
      const text = (row.innerText || '').split('\\n').map(strip).filter(Boolean);
      // Primary unread checks first; only fall back to the bold-text marker
      // OWA paints on unread rows (selectable rows only, to limit false hits).
      let unread = /\\bunread\\b/i.test(label)
        || row.querySelector('[aria-label*="Unread" i]')
        || row.querySelector('span[class*="unread" i]');
      if (!unread && row.hasAttribute('aria-selected')) {
        const spans = row.querySelectorAll('span');
        for (let s = 0; s < spans.length && s < 12; s++) {
          const w = getComputedStyle(spans[s]).fontWeight;
          if (w === 'bold' || parseInt(w, 10) >= 600) { unread = true; break; }
        }
      }
      if (!unread) continue;
      // Today detection from the row's own stamp: today's mail shows a bare
      // time ("9:30 AM"), older mail a day/date prefix.
      const hasTime = text.some(t => /^\\d{1,2}:\\d{2}(\\s*(AM|PM))?$/i.test(t));
      const hasOlderStamp = text.some(t => /^(Yesterday|Mon|Tue|Wed|Thu|Fri|Sat|Sun)\\b/i.test(t));
      const today = text.some(t => /^Today$/i.test(t)) || (hasTime && !hasOlderStamp);
      // Clean the text lines: OWA prepends an avatar-initials token ("MO", "GC")
      // and sprinkles in time/day stamps — strip both so sender/subject align.
      // Lines without a single letter/digit are leftover icon glyphs — drop them.
      let parts = text.filter((t) =>
        /[\\p{L}\\p{N}]/u.test(t) &&                                // must contain real text
        !/^[A-Z]{1,3}$/.test(t) &&                                  // avatar initials
        !/^\\d{1,2}:\\d{2}\\s*(AM|PM)?$/i.test(t) &&                 // bare time
        !/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\\b/i.test(t) &&            // day-prefixed stamps
        !/^(Today|Yesterday)$/i.test(t)
      );
      let sender = (parts[0] || '').trim();
      let subject = (parts[1] || '').trim();
      let preview = (parts[2] || '').trim();
      // Fallback: OWA aria-labels read "Unread, Sender, Subject, Preview, …" —
      // parse them when the child spans are missing or empty.
      if (!sender || !subject) {
        const lp = label.replace(/^(un)?read[,.]?\\s*/i, '').split(',').map(strip).filter(s => s && /[\\p{L}\\p{N}]/u.test(s));
        sender = sender || lp[0] || '';
        subject = subject || lp[1] || '';
        preview = preview || lp[2] || '';
      }
      // Sender+subject fingerprint for deduplication — null byte as separator
      const id = sender + '\\x00' + subject;
      if ((!sender && !subject) || seen.has(id)) continue;
      seen.add(id);
      items.push({
        sender: sender.slice(0, 60),
        subject: subject.slice(0, 90),
        preview: preview.slice(0, 140),
        id,
        today,
        rowIdx: i
      });
    }
    return { state: items.length ? 'ok' : 'empty', items };
  } catch (e) { return { state: 'error', items: [] }; }
})()`

// Asana uses CSS modules with hashed class names, so we try several selector
// strategies and fall back to broader role-based matching if none hit.
const ASANA_SCRAPE = `(() => {
  try {
    // Detect login / unauthenticated state before trying to scrape tasks.
    const isLoginPage = document.querySelector('[class*="LoginForm" i], [data-testid="login-form"]')
      || /^\\/auth|^\\/sign-?in/i.test(location.pathname)
      || location.hostname === 'account.asana.com';
    if (isLoginPage) return { state: 'login', items: [] };

    const seen = new Set();
    const items = [];
    // Try specific selectors first; fall back to generic [role="row"].
    const selectors = [
      '[data-testid="task-row-content"]',
      '[data-testid*="TaskRow"]',
      '[class*="taskRow" i]',
      '[class*="TaskRow" i]',
      '.TaskRow'
    ];
    let rows = [];
    for (const sel of selectors) {
      rows = Array.from(document.querySelectorAll(sel));
      if (rows.length) break;
    }
    if (!rows.length) rows = Array.from(document.querySelectorAll('[role="row"]'));

    for (let ri = 0; ri < rows.length; ri++) {
      const row = rows[ri];
      const nameEl = row.querySelector(
        '[data-testid*="task-name"], [class*="TaskName" i] textarea, ' +
        '[class*="TaskName" i] [contenteditable], [class*="taskName" i], ' +
        'a[href*="/task/"], [data-task-name]'
      );
      let name = '';
      if (nameEl) name = (nameEl.value || nameEl.getAttribute('data-task-name') || nameEl.innerText || '').trim();
      if (!name) {
        const texts = (row.innerText || '').split('\\n').map(s => s.trim()).filter(s => s && s.length > 1 && s.length < 200);
        name = texts[0] || '';
      }
      name = name.slice(0, 90);
      if (!name || seen.has(name)) continue;
      // Skip rows already marked complete — not actionable, and they'd churn
      // the diff if Asana keeps them on screen for a while after completion.
      if (row.querySelector('[role="checkbox"][aria-checked="true"], [class*="isCompleted" i]')) continue;
      const isSub = !!row.closest('[class*="Subtask" i]')
        || (row.getAttribute('aria-level') && row.getAttribute('aria-level') !== '1');
      seen.add(name);
      const taskLinkEl = row.querySelector('a[href*="/task/"], a[href*="/0/"]');
      const taskUrl = taskLinkEl ? taskLinkEl.href : null;
      // Stable id: the numeric task gid from a permalink or data attribute,
      // falling back to the name so diff-based notifications still work.
      let gid = row.getAttribute('data-task-id') || '';
      if (!gid && taskUrl) {
        const m = taskUrl.match(/\\/task\\/(\\d+)|\\/item\\/(\\d+)|\\/0\\/\\d+\\/(\\d+)/);
        if (m) gid = m[1] || m[2] || m[3] || '';
      }
      if (isSub && items.length) { (items[items.length - 1].subtasks ||= []).push(name); }
      else items.push({ id: gid || name, name, subtasks: [], rowIdx: ri, taskUrl });
      if (items.length >= 8) break;
    }
    return { state: items.length ? 'ok' : 'empty', items };
  } catch (e) { return { state: 'error', items: [] }; }
})()`

// OWA calendar: try automation-id selectors first, then fall back to
// aria-label parsing on buttons/listitem children that look like events.
const CALENDAR_SCRAPE = `(() => {
  try {
    // Not signed in: report the login page so the sidebar says so instead of
    // showing "No upcoming events" for an inbox that simply isn't loaded.
    if (/^(login|account)\\./i.test(location.hostname) || document.querySelector('input[name="loginfmt"]')) {
      return { state: 'login', items: [] };
    }
    const items = [];
    const seen = new Set();
    let cards = Array.from(document.querySelectorAll(
      '[data-automationid="CalendarEventItem"], [data-automationid="calendarAgendaItem"], ' +
      '[data-automationid="calendarListItem"]'
    ));
    if (!cards.length) {
      // Day/week view: events are buttons with time-containing aria-labels.
      // Calendar chrome (navigation arrows, view switchers, date cells) also
      // carries aria-labels, so require a full h:mm time plus real title text
      // and reject the known chrome verbs.
      cards = Array.from(document.querySelectorAll('button[aria-label], [role="button"][aria-label]'))
        .filter(el => {
          const label = el.getAttribute('aria-label') || '';
          if (!/\\d{1,2}[:.]\\d{2}/.test(label)) return false;
          if (/^(next|previous|go to|today|search|close|new event|switch|open|month view|week view|day view)/i.test(label)) return false;
          return /[\\p{L}]{2,}/u.test(label);
        });
    }
    const now = new Date();
    const todayName = now.toLocaleDateString(undefined, { weekday: 'long' }).toLowerCase();
    // All seven weekday names in the page's locale (2024-01-07 was a Sunday),
    // so the "names a different day" guard below works beyond English OWA.
    const dayNames = Array.from({ length: 7 }, (_, d) =>
      new Date(2024, 0, 7 + d).toLocaleDateString(undefined, { weekday: 'long' }).toLowerCase());
    for (let i = 0; i < cards.length; i++) {
      const card = cards[i];
      const label = (card.getAttribute('aria-label') || '').trim();
      const parts = label.split(',').map(s => s.trim()).filter(Boolean);
      const title = parts[0] || (card.innerText || '').split('\\n')[0].trim();
      const time = parts.find(p => /\\d{1,2}[:.]/i.test(p)) || parts[1] || '';
      if (!title || title.length < 2) continue;
      // Best-effort start time for reminders: first h:mm in the label plus
      // today's date. Only trusted when the label doesn't name a DIFFERENT
      // weekday, so next week's agenda can never fire a reminder today.
      let startIso = null;
      const t = time.match(/(\\d{1,2})[:.](\\d{2})\\s*(AM|PM)?/i);
      const labelLc = label.toLowerCase();
      const otherDay = dayNames.some((d) => d !== todayName && labelLc.includes(d));
      if (t && !otherDay) {
        let h = parseInt(t[1], 10);
        const m = parseInt(t[2], 10);
        const ap = (t[3] || '').toUpperCase();
        if (ap === 'PM' && h < 12) h += 12;
        if (ap === 'AM' && h === 12) h = 0;
        if (h >= 0 && h < 24 && m >= 0 && m < 60) {
          startIso = new Date(now.getFullYear(), now.getMonth(), now.getDate(), h, m, 0, 0).toISOString();
        }
      }
      // Date-stamped id: a recurring event ("Standup", "9:30 AM") must get a
      // fresh id each day or it reminds only once per app session.
      const id = now.toDateString() + '\\x00' + title + '\\x00' + time;
      if (seen.has(id)) continue;
      seen.add(id);
      items.push({ id, title: title.slice(0, 80), time: time.slice(0, 40), startIso, rowIdx: i });
      if (items.length >= 20) break;
    }
    return { state: items.length ? 'ok' : 'empty', items };
  } catch(e) { return { state: 'error', items: [] }; }
})()`

// Stable identity for a feed list: which actual items it shows, ignoring
// volatile fields like rowIdx that shift as OWA virtualizes its rows.
function feedSignature(kind, items) {
  return (items || [])
    .map((i) =>
      kind === 'mail'
        ? i.id || `${i.sender}\u0000${i.subject}`
        : kind === 'asana'
          ? i.id || i.name
          : i.id || `${i.title}\u0000${i.time || ''}`
    )
    .join('\n')
}

// Apply a fetched { state, items } result to a feed using the cache rule:
// good data always wins; an empty/error/auth result only overwrites the cache
// when this tab is active or there's no last-known-good list yet. API results
// pass trusted=true — an authoritative "empty" (inbox zero, no events) always
// replaces the cache; the active-only heuristic exists for flaky DOM scrapes.
// Returns true only when the feed visibly changed — when the same items come
// back, the cached list is refreshed in place (rowIdx/links) but never
// replaced, so repeated polls don't churn the sidebar.
function applyFeedResult(key, feed, result, { trusted = false } = {}) {
  const isActive = key === activeServiceKey
  const hadCache = Array.isArray(feed.items) && feed.items.length > 0
  const gotItems = result.items.length > 0
  const prevState = feed.state
  if (gotItems) {
    feed.state = result.state || 'ok'
    if (hadCache && feedSignature(feed.kind, feed.items) === feedSignature(feed.kind, result.items)) {
      feed.items.forEach((item, i) => Object.assign(item, result.items[i]))
      return feed.state !== prevState
    }
    feed.items = result.items
    return true
  }
  if (trusted || isActive || !hadCache) {
    feed.state = result.state || 'empty'
    feed.items = result.items
    return feed.state !== prevState || hadCache
  }
  return false
}

// API path: the owning provider is connected, so Graph/Asana is the source of
// truth. Polls even while the panel is hidden (background notifications).
async function refreshFeedFromApi(key, feed) {
  const gen = feed.gen || 0
  const result = await connections.getFeed(feed.kind)
  // The feed's source of truth changed while we were fetching (connect/
  // disconnect cutover) — this result belongs to the old source, drop it.
  if ((feed.gen || 0) !== gen) return
  if (!result) {
    // No token / auth error — keep the cached list, but if we have nothing to
    // show, signal that a (re)connect is needed.
    if ((!Array.isArray(feed.items) || !feed.items.length) && feed.state !== 'auth') {
      feed.state = 'auth'
      pushSnapshot()
    }
    return
  }
  let changed = applyFeedResult(key, feed, result, { trusted: true })

  if (feed.kind === 'mail') {
    const unread = await connections.getMailUnreadCount()
    if ((feed.gen || 0) !== gen) return
    if (serviceState[key] && serviceState[key].unreadCount !== unread) {
      serviceState[key].unreadCount = unread
      changed = true
    }
    diffAndNotifyMail(feed.items, unread)
  } else if (feed.kind === 'asana') {
    if (serviceState[key] && serviceState[key].unreadCount !== feed.items.length) {
      serviceState[key].unreadCount = feed.items.length
      changed = true
    }
    diffAndNotifyAsana(feed.items)
  } else if (feed.kind === 'calendar') {
    maybeNotifyCalendar(feed.items)
  }
  if (changed) pushSnapshot()
}

function refreshFeed(key) {
  const feed = serviceFeeds[key]
  if (!feed) {
    return
  }
  // When the provider behind this feed is connected, prefer the API entirely.
  if (connections.feedIsLive(feed.kind)) {
    refreshFeedFromApi(key, feed).catch(() => {
      if (!Array.isArray(feed.items) || !feed.items.length) {
        feed.state = 'error'
        pushSnapshot()
      }
    })
    return
  }

  const view = serviceViews.get(key)
  if (!view || view.webContents.isDestroyed() || view.webContents.isLoading()) {
    return
  }
  const script = feed.kind === 'mail' ? MAIL_SCRAPE : feed.kind === 'asana' ? ASANA_SCRAPE : feed.kind === 'calendar' ? CALENDAR_SCRAPE : null
  if (!script) {
    return
  }
  const gen = feed.gen || 0
  view.webContents
    .executeJavaScript(script, true)
    .then((result) => {
      // Drop scrape results that resolve after the provider connected (or the
      // feed was otherwise reset) — the API owns the feed from that moment.
      if ((feed.gen || 0) !== gen || connections.feedIsLive(feed.kind)) return
      if (result && Array.isArray(result.items)) {
        // Cache rule lives in applyFeedResult: good data always wins, an
        // empty/error result is only trusted when this tab is ACTIVE, and an
        // identical scrape leaves the cached list untouched (no churn).
        let changed = applyFeedResult(key, feed, result)

        // Scraped items now carry the same fields the API path produces
        // (ids, startIso), so the same notification diffing applies pre-connect.
        if (feed.kind === 'mail') {
          const mailState = serviceState[key]
          diffAndNotifyMail(feed.items, mailState ? mailState.unreadCount : 0)
        } else if (feed.kind === 'asana') {
          // Badge mirrors the visible task count, same as the API path.
          if (serviceState[key] && serviceState[key].unreadCount !== feed.items.length) {
            serviceState[key].unreadCount = feed.items.length
            changed = true
          }
          diffAndNotifyAsana(feed.items)
        } else if (feed.kind === 'calendar') {
          maybeNotifyCalendar(feed.items)
        }
        if (changed) pushSnapshot()
      }
    })
    .catch(() => {
      // Network/scrape failure: keep cached items, only flag error if empty.
      if ((!Array.isArray(feed.items) || !feed.items.length) && feed.state !== 'error') {
        feed.state = 'error'
        pushSnapshot()
      }
    })
}

function refreshFeeds() {
  // Snoozes expire lazily; reconcile tab audio and the zZ badge here so a
  // naturally lapsed snooze unmutes within one tick instead of staying muted
  // until the next manual reload.
  let snoozeLapsed = false
  const now = Date.now()
  for (const [key, exp] of serviceSnooze) {
    if (now >= exp) {
      serviceSnooze.delete(key)
      snoozeLapsed = true
    }
  }
  for (const key of serviceViews.keys()) applyServiceMute(key)
  if (snoozeLapsed) pushSnapshot()

  const visible = Boolean(panelWindow && panelWindow.isVisible() && !settingsOpen)
  for (const key of Object.keys(serviceFeeds)) {
    const kind = serviceFeeds[key].kind
    // API-backed feeds poll continuously — no view needed, fire every tick.
    if (connections.feedIsLive(kind)) {
      refreshFeed(key)
      continue
    }
    // Scrape feeds need a loaded view. Previously only the visible panel's views
    // scraped, so background tabs never notified — the "only the first tab gets
    // notifications" bug. Now every loaded view scrapes even with the panel
    // hidden; hidden scrapes are throttled (HIDDEN_SCRAPE_MS) to spare CPU.
    if (!loadedServiceKeys.has(key)) continue
    const view = serviceViews.get(key)
    if (!view || view.webContents.isDestroyed()) continue
    if (!visible && now - (lastScrapeAt[key] || 0) < HIDDEN_SCRAPE_MS) continue
    lastScrapeAt[key] = now
    refreshFeed(key)
  }
}

function startFeedTimer() {
  if (feedTimer) {
    clearInterval(feedTimer)
  }
  feedTimer = setInterval(refreshFeeds, FEED_REFRESH_MS)
}

/* ---------- Tray ---------- */
function mailUnread() {
  for (const key of Object.keys(serviceFeeds)) {
    if (serviceFeeds[key].kind === 'mail') {
      return serviceState[key] ? serviceState[key].unreadCount : 0
    }
  }
  return serviceState.mail ? serviceState.mail.unreadCount : 0
}

function createTray() {
  const image = nativeImage.createFromPath(getTrayIconPath(0))
  image.setTemplateImage(true)
  tray = new Tray(image)
  tray.setToolTip(APP_NAME)
  tray.on('click', () => toggleMenuWindow())
  tray.on('right-click', () => tray.popUpContextMenu(buildTrayContextMenu()))
  updateTray()
}

function buildTrayContextMenu() {
  const unread = mailUnread()
  return Menu.buildFromTemplate([
    { label: 'Open MailStudio', click: () => togglePanelWindow(true) },
    { label: 'Menu Panel', click: () => toggleMenuWindow(true) },
    { type: 'separator' },
    ...visibleServices().map((service) => ({
      label: serviceState[service.key] && serviceState[service.key].unreadCount > 0
        ? `${service.label} (${serviceState[service.key].unreadCount})`
        : service.label,
      type: 'radio',
      checked: service.key === activeServiceKey,
      click: () => showService(service.key)
    })),
    { type: 'separator' },
    { label: 'Reload Active View', click: () => navigate('reload') },
    { label: 'Open Active in Browser', click: () => openExternalSafe(serviceState[activeServiceKey].href) },
    { type: 'separator' },
    { label: unread > 0 ? `${unread} unread` : 'Inbox caught up', enabled: false },
    { label: 'Report an Issue…', click: () => openExternalSafe(`${REPO_URL}/issues/new?template=bug_report.yml`) },
    { label: 'Quit MailStudio', click: () => { app.isQuitting = true; app.quit() } }
  ])
}

function updateTray() {
  const unread = mailUnread()

  // Dock badge (macOS "app notification number") — set this regardless of the
  // tray so the count always reflects on the app icon.
  if (typeof app.setBadgeCount === 'function') {
    app.setBadgeCount(unread)
  }
  if (process.platform === 'darwin' && app.dock && typeof app.dock.setBadge === 'function') {
    app.dock.setBadge(unread > 0 ? String(unread) : '')
  }

  if (!tray) {
    return
  }
  const trayImage = nativeImage.createFromPath(getTrayIconPath(unread))
  trayImage.setTemplateImage(true)
  tray.setImage(trayImage)
  tray.setTitle(unread > 0 ? ` ${unread}` : '')
}

/* ---------- Main window ---------- */
// The panel + dropdown are trusted local pages that should never navigate away
// or spawn windows. Block both; stray links route in-app when a tab owns them
// (Microsoft/Asana), otherwise to the browser.
function hardenLocalWindow(win) {
  const wc = win.webContents
  wc.setWindowOpenHandler(({ url }) => {
    openLinkInApp(url)
    return { action: 'deny' }
  })
  wc.on('will-navigate', (event, url) => {
    if (!url.startsWith('file://')) {
      event.preventDefault()
      openLinkInApp(url)
    }
  })
}

function createPanelWindow() {
  panelWindow = new BrowserWindow({
    width: WINDOW_SIZE.width,
    height: WINDOW_SIZE.height,
    minWidth: WINDOW_SIZE.minWidth,
    minHeight: WINDOW_SIZE.minHeight,
    show: false,
    center: true,
    resizable: true,
    maximizable: true,
    minimizable: true,
    fullscreenable: true,
    title: APP_NAME,
    backgroundColor: IS_GLASS_MODE ? '#00000000' : '#16191e',
    ...(IS_GLASS_MODE ? { vibrancy: 'under-window' } : {}),
    autoHideMenuBar: true,
    frame: false,
    titleBarStyle: 'hiddenInset',
    movable: true,
    trafficLightPosition: { x: 18, y: 18 },
    webPreferences: {
      preload: path.join(__dirname, 'panel-preload.js'),
      sandbox: true,
      contextIsolation: true,
      nodeIntegration: false
    }
  })

  panelWindow.webContents.setMaxListeners(20)
  panelWindow.loadFile(path.join(__dirname, '..', 'renderer', 'panel.html'))
  hardenLocalWindow(panelWindow)
  panelWindow.on('resize', () => attachServiceView())
  panelWindow.on('swipe', (_event, direction) => {
    if (direction === 'right') navigate('back')
    else if (direction === 'left') navigate('forward')
  })
  panelWindow.on('show', () => {
    attachServiceView()
    updateVisibleStates()
    pushSnapshot()
    refreshFeeds()
  })

  panelWindow.on('close', (event) => {
    if (!app.isQuitting) {
      event.preventDefault()
      panelWindow.hide()
      updateVisibleStates()
      pushSnapshot()
    }
  })
}

function showPanelWindow() {
  if (!panelWindow) {
    createPanelWindow()
  }
  panelWindow.show()
  panelWindow.focus()
  attachServiceView()
  pushSnapshot()
}

function togglePanelWindow(forceShow = false) {
  if (!panelWindow) {
    createPanelWindow()
  }
  if (forceShow || !panelWindow.isVisible()) {
    showPanelWindow()
    return
  }
  panelWindow.hide()
  updateVisibleStates()
  pushSnapshot()
}

/* ---------- Menu bar dropdown ---------- */
function createMenuWindow() {
  menuWindow = new BrowserWindow({
    width: MENU_SIZE.width,
    height: MENU_SIZE.height,
    show: false,
    frame: false,
    transparent: true,
    backgroundColor: '#00000000',
    ...(IS_GLASS_MODE ? { vibrancy: 'popover' } : {}),
    hasShadow: true,
    resizable: false,
    movable: false,
    minimizable: false,
    maximizable: false,
    fullscreenable: false,
    skipTaskbar: true,
    alwaysOnTop: true,
    webPreferences: {
      preload: path.join(__dirname, 'panel-preload.js'),
      sandbox: true,
      contextIsolation: true,
      nodeIntegration: false
    }
  })

  menuWindow.webContents.setMaxListeners(20)
  menuWindow.loadFile(path.join(__dirname, '..', 'renderer', 'menu.html'))
  hardenLocalWindow(menuWindow)
  menuWindow.setVisibleOnAllWorkspaces(true, { visibleOnFullScreen: true })
  menuWindow.on('blur', () => {
    if (!menuWindow.webContents.isDevToolsOpened()) {
      menuWindow.hide()
    }
  })
}

function positionMenuWindow() {
  if (!tray || !menuWindow) {
    return
  }
  const trayBounds = tray.getBounds()
  const { width } = menuWindow.getBounds()
  const x = Math.round(trayBounds.x + trayBounds.width / 2 - width / 2)
  const y = Math.round(trayBounds.y + trayBounds.height + 2)
  menuWindow.setPosition(Math.max(8, x), Math.max(8, y), false)
}

function showMenuWindow() {
  if (!menuWindow) {
    createMenuWindow()
  }
  positionMenuWindow()
  menuWindow.show()
  menuWindow.focus()
  pushSnapshot()
}

function toggleMenuWindow(forceShow = false) {
  if (!menuWindow) {
    createMenuWindow()
  }
  if (forceShow || !menuWindow.isVisible()) {
    showMenuWindow()
    return
  }
  menuWindow.hide()
}

function hideMenuWindow() {
  if (menuWindow && menuWindow.isVisible()) {
    menuWindow.hide()
  }
}


/* ---------- Snapshot ---------- */
function getSnapshot() {
  const activeWebContents = serviceViews.get(activeServiceKey)?.webContents
  return {
    appName: APP_NAME,
    activeServiceKey,
    splitKeys: inSplit() ? splitKeys.slice() : [],
    splitOrientation,
    splitRatio,
    sidebarCollapsed,
    collapseMode: settings.collapseMode,
    taskProvider: settings.taskProvider,
    notif: settings.notif,
    firstBoot,
    onboarded: Boolean(settings.onboarded),
    notifSetupSkipped: Boolean(settings.notifSetupSkipped),
    connections: connections.getStatus(),
    connConfig: settings.connections,
    scratch: settings.scratch || '',
    settingsOpen,
    theme: settings.theme,
    glassMode: IS_GLASS_MODE,
    nav: {
      canGoBack: activeWebContents ? canGoBack(activeWebContents) : false,
      canGoForward: activeWebContents ? canGoForward(activeWebContents) : false
    },
    globalSnoozed: isGlobalSnoozed(),
    services: settings.services.map((service) => {
      const state = serviceState[service.key] || {}
      const feed = serviceFeeds[service.key]
      return {
        key: service.key,
        label: service.label,
        icon: service.icon,
        url: service.url,
        home: service.home,
        builtin: service.builtin,
        visible: service.visible,
        title: state.title || service.label,
        unreadCount: state.unreadCount || 0,
        href: state.href || service.url,
        snoozed: isSnoozed(service.key),
        feed: feed ? { kind: feed.kind, state: feed.state, items: feed.items } : null
      }
    })
  }
}

function pushSnapshot() {
  updateTray()
  const snapshot = getSnapshot()
  if (panelWindow && !panelWindow.isDestroyed()) {
    panelWindow.webContents.send('panel:status-updated', snapshot)
  }
  if (menuWindow && !menuWindow.isDestroyed()) {
    menuWindow.webContents.send('panel:status-updated', snapshot)
  }
}

/* ---------- Settings application ---------- */
function applySettings(next) {
  settings = store.save(next)
  syncServiceViews()
  buildAppMenu()
  if (panelWindow) {
    attachServiceView()
  }
  pushSnapshot()
  setTimeout(refreshFeeds, 1500)
}

/* ---------- IPC ---------- */
function registerIpc() {
  ipcMain.handle('panel:get-snapshot', () => getSnapshot())

  ipcMain.on('panel:command', (_event, command) => {
    switch (command.type) {
      case 'switch-service':
        showService(command.serviceKey)
        hideMenuWindow()
        break
      case 'split-select':
        if (typeof command.serviceKey === 'string') {
          splitSelect(command.serviceKey)
          hideMenuWindow()
        }
        break
      case 'toggle-split-orientation':
        // Flip side-by-side ↔ stacked. Only meaningful while a split is up.
        if (inSplit()) {
          splitOrientation = splitOrientation === 'vertical' ? 'horizontal' : 'vertical'
          attachServiceView()
          pushSnapshot()
        }
        break
      case 'split-drag-start':
        // Detach the panes so the renderer owns the content area and can paint a
        // live preview while the user drags the divider.
        if (inSplit()) {
          splitDragging = true
          detachAllViews()
        }
        break
      case 'split-drag-end':
        // Re-attach at the dropped ratio (clamped so both panes stay usable).
        splitDragging = false
        if (typeof command.ratio === 'number' && Number.isFinite(command.ratio)) {
          splitRatio = Math.min(SPLIT_RATIO_MAX, Math.max(SPLIT_RATIO_MIN, command.ratio))
        }
        attachServiceView()
        pushSnapshot()
        break
      case 'open-app':
        togglePanelWindow(true)
        hideMenuWindow()
        break
      case 'hide-menu':
        hideMenuWindow()
        break
      case 'set-theme':
        if (command.theme === 'dark' || command.theme === 'light') {
          applySettings({ ...settings, theme: command.theme })
        }
        break
      case 'menu-resize':
        if (menuWindow && !menuWindow.isDestroyed() && Number.isFinite(command.height)) {
          const height = Math.max(160, Math.min(720, Math.round(command.height)))
          menuWindow.setContentSize(MENU_SIZE.width, height)
          positionMenuWindow()
        }
        break
      case 'quit':
        app.isQuitting = true
        app.quit()
        break
      case 'toggle-sidebar':
        sidebarCollapsed = !sidebarCollapsed
        attachServiceView()
        pushSnapshot()
        break
      case 'open-settings':
        settingsOpen = true
        attachServiceView()
        updateVisibleStates()
        pushSnapshot()
        break
      case 'close-settings':
        settingsOpen = false
        attachServiceView()
        updateVisibleStates()
        pushSnapshot()
        break
      case 'update-settings':
        if (command.settings) {
          applySettings({
            ...settings,
            collapseMode: command.settings.collapseMode || settings.collapseMode,
            taskProvider: command.settings.taskProvider || settings.taskProvider,
            notif: command.settings.notif || settings.notif,
            services: Array.isArray(command.settings.services) ? command.settings.services : settings.services
          })
        }
        break
      case 'go-home':
        goHome()
        break
      case 'nav-back':
        navigate('back')
        break
      case 'nav-forward':
        navigate('forward')
        break
      case 'nav-reload':
        navigate('reload')
        break
      case 'refresh-feeds':
        refreshFeeds()
        break
      case 'open-feed-item': {
        const targetView = serviceViews.get(command.serviceKey)
        if (targetView && !targetView.webContents.isDestroyed()) {
          const feedKind = serviceFeeds[command.serviceKey] ? serviceFeeds[command.serviceKey].kind : null
          // API feed items carry a deep link (webLink/permalink) — open it
          // directly in the owning view rather than clicking a scraped row.
          const deepLink = (typeof command.webLink === 'string' && command.webLink) ? command.webLink : null
          if (deepLink) {
            loadInServiceView(targetView, command.serviceKey, deepLink)
            revealFeedTarget(command.serviceKey, command.split)
            break
          }
          if (feedKind === 'mail' && typeof command.rowIdx === 'number') {
            const idx = Math.trunc(command.rowIdx)
            setTimeout(() => {
              targetView.webContents
                .executeJavaScript(
                  `(() => { const rows = Array.from(document.querySelectorAll('div[role="option"], div[role="listitem"]')); const row = rows[${idx}]; if (row) { row.click(); row.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); } })()`,
                  true
                )
                .catch(() => {})
            }, 200)
          } else if (feedKind === 'calendar' && typeof command.rowIdx === 'number') {
            const idx = Math.trunc(command.rowIdx)
            setTimeout(() => {
              targetView.webContents
                .executeJavaScript(
                  `(() => { let cards = Array.from(document.querySelectorAll('[data-automationid="CalendarEventItem"],[data-automationid="calendarAgendaItem"],[data-automationid="calendarListItem"]')); if (!cards.length) cards = Array.from(document.querySelectorAll('button[aria-label],[role="button"][aria-label]')).filter(el => /\\d{1,2}[:.]/i.test(el.getAttribute('aria-label')||'')); const card = cards[${idx}]; if (card) { card.click(); card.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); } })()`,
                  true
                )
                .catch(() => {})
            }, 200)
          } else if (feedKind === 'asana') {
            if (command.taskUrl && typeof command.taskUrl === 'string') {
              loadInServiceView(targetView, command.serviceKey, command.taskUrl)
            } else if (typeof command.rowIdx === 'number') {
              const idx = Math.trunc(command.rowIdx)
              setTimeout(() => {
                targetView.webContents
                  .executeJavaScript(
                    `(() => { const sels = ['[data-testid="task-row-content"]','[data-testid*="TaskRow"]','[class*="taskRow" i]','[class*="TaskRow" i]','.TaskRow','[role="row"]']; let rows = []; for (const s of sels) { rows = Array.from(document.querySelectorAll(s)); if (rows.length) break; } const row = rows[${idx}]; if (row) { const link = row.querySelector('a[href*="/task/"],a[href*="/0/"]'); if (link) link.click(); else row.click(); row.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); } })()`,
                    true
                  )
                  .catch(() => {})
              }, 200)
            }
          }
          revealFeedTarget(command.serviceKey, command.split)
        }
        break
      }
      case 'open-external':
        openExternalSafe(serviceState[activeServiceKey].href)
        hideMenuWindow()
        break
      case 'open-url':
        if (typeof command.url === 'string') openLinkInApp(command.url)
        break
      case 'compose':
        if (typeof command.kind === 'string') compose(command.kind)
        break
      case 'focus-done':
        // User-started timer: fires regardless of the onboarding notification
        // gate — the user explicitly asked for this one.
        fireNotification(
          { title: 'Focus session complete', body: '25 minutes — take a break.' },
          () => showPanelWindow()
        )
        break
      case 'start-signin':
        finishFirstBoot()
        wantPrewarm = true
        servicesPrewarmed = false
        showService('mail')
        break
      case 'skip-signin':
        finishFirstBoot()
        showService(activeServiceKey)
        break
      // Reserved for tray/automation use — no renderer sender today.
      case 'snooze-service':
        if (typeof command.serviceKey === 'string') {
          const mins = Number(command.minutes)
          setSnooze(command.serviceKey, mins > 0 ? Date.now() + mins * 60000 : 0)
        }
        break
      case 'global-snooze-menu':
        showGlobalSnoozeMenu()
        break
      case 'tab-context-menu':
        if (typeof command.serviceKey === 'string') {
          showTabContextMenu(command.serviceKey)
        }
        break
      case 'save-scratch':
        // Lightweight persist — don't churn views/menu/feeds like applySettings.
        settings = store.save({ ...settings, scratch: typeof command.text === 'string' ? command.text.slice(0, 20000) : '' })
        break
      case 'connect-provider':
        if (typeof command.provider === 'string') {
          // State transitions (connecting → connected/error) flow back through
          // the connections onChange callback → pushSnapshot.
          console.log(`[connect] starting ${command.provider}`)
          // Detach the on-top service views so the sign-in popup is never
          // hidden behind a service web page; restore them when sign-in settles.
          detachAllViews()
          connections
            .connect(command.provider, { parentWindow: panelWindow })
            .then(() => {
              // The OAuth popup signs into the shared MS session too, so a
              // Microsoft connect also warms the logged-in web views.
              if (command.provider === 'microsoft') {
                finishFirstBoot()
                prewarmMicrosoftServices()
              }
              // Scrapers are a pre-connect fallback only: drop their cached
              // items and baselines so the API owns the feeds from here on.
              resetProviderFeeds(command.provider)
              // Connecting any provider arms notifications — the onboarding
              // sheet's "Enable notifications" button is an explicit confirm
              // of the same thing, so users who connect via Settings (or close
              // the sheet early) aren't left permanently silent.
              if (!settings.onboarded) {
                settings = store.save({ ...settings, onboarded: true })
              }
              pushSnapshot()
              setTimeout(refreshFeeds, 800)
            })
            .catch((e) => console.error(`[connect] ${command.provider} failed:`, e && e.message))
            .finally(() => {
              if (panelWindow && !panelWindow.isDestroyed()) {
                attachServiceView()
              }
            })
        }
        break
      case 'disconnect-provider':
        if (typeof command.provider === 'string') {
          connections.disconnect(command.provider)
          // Drop the cached API feed and baselines for that provider's kinds
          // so the sidebar falls back to scrape/empty immediately.
          resetProviderFeeds(command.provider)
          pushSnapshot()
          setTimeout(refreshFeeds, 500)
        }
        break
      case 'save-connections':
        if (command.connections && typeof command.connections === 'object') {
          settings = store.save({ ...settings, connections: command.connections })
          connections.setConfig(settings.connections)
          pushSnapshot()
        }
        break
      case 'open-onboarding':
        onboardingOpen = true
        attachServiceView()
        updateVisibleStates()
        pushSnapshot()
        break
      case 'close-onboarding':
        onboardingOpen = false
        attachServiceView()
        updateVisibleStates()
        pushSnapshot()
        break
      case 'finish-onboarding':
        settings = store.save({ ...settings, onboarded: true })
        onboardingOpen = false
        finishFirstBoot()
        attachServiceView()
        pushSnapshot()
        break
      case 'reset-onboarding':
        settings = store.save({ ...settings, onboarded: false, notifSetupSkipped: false })
        pushSnapshot()
        break
      case 'skip-notif-setup':
        settings = store.save({ ...settings, notifSetupSkipped: true })
        pushSnapshot()
        break
      case 'find-in-page': {
        const fView = serviceViews.get(activeServiceKey)
        if (fView && !fView.webContents.isDestroyed() && typeof command.text === 'string' && command.text) {
          fView.webContents.findInPage(command.text, { forward: command.forward !== false, findNext: !!command.findNext })
        } else if (!command.text) {
          const fViewStop = serviceViews.get(activeServiceKey)
          if (fViewStop && !fViewStop.webContents.isDestroyed()) fViewStop.webContents.stopFindInPage('clearSelection')
          sendPanelEvent({ type: 'find-result', activeMatchOrdinal: 0, matches: 0 })
        }
        break
      }
      case 'stop-find-in-page': {
        const sView = serviceViews.get(activeServiceKey)
        if (sView && !sView.webContents.isDestroyed()) sView.webContents.stopFindInPage('clearSelection')
        sendPanelEvent({ type: 'find-result', activeMatchOrdinal: 0, matches: 0 })
        break
      }
      case 'network-online':
        // Network came back — reload any feed-enabled views that are empty/errored
        // so the sidebar repopulates without waiting for the next 25 s poll.
        for (const nKey of Object.keys(serviceFeeds)) {
          const nView = serviceViews.get(nKey)
          if (nView && !nView.webContents.isDestroyed()) {
            const nFeed = serviceFeeds[nKey]
            if (nFeed && (nFeed.state === 'error' || !nFeed.items.length)) {
              nView.webContents.reload()
            }
          }
        }
        setTimeout(refreshFeeds, 4000)
        break
      default:
        break
    }
  })
}

// Native right-click menu for a sidebar tab: snooze notifications, reload, open
// externally. Snooze only applies to mail/teams (the notifying services).
function showTabContextMenu(serviceKey) {
  const service = findService(serviceKey)
  if (!service) return
  const canSnooze = SNOOZABLE_SERVICES.includes(serviceKey)
  const snoozed = isSnoozed(serviceKey)
  const items = []

  if (canSnooze) {
    if (snoozed) {
      items.push({ label: 'Resume notifications', click: () => setSnooze(serviceKey, 0) })
    } else {
      items.push({ label: 'Snooze notifications & sound', enabled: false })
      items.push({ label: 'For 1 hour', click: () => setSnooze(serviceKey, Date.now() + 60 * 60000) })
      items.push({ label: 'For 4 hours', click: () => setSnooze(serviceKey, Date.now() + 240 * 60000) })
      items.push({ label: 'Until tomorrow', click: () => {
        const t = new Date(); t.setHours(24, 0, 0, 0)
        setSnooze(serviceKey, t.getTime())
      } })
    }
    items.push({ type: 'separator' })
  }

  items.push({ label: `Reload ${service.label}`, click: () => {
    const v = serviceViews.get(serviceKey)
    if (v && !v.webContents.isDestroyed()) v.webContents.reload()
  } })
  items.push({ label: 'Open in browser', click: () => openExternalSafe((serviceState[serviceKey] || {}).href || service.url) })

  Menu.buildFromTemplate(items).popup({ window: panelWindow })
}

// Dropdown for the global snooze (zzz) button in the topbar: snooze every tab and
// mute all sound, or resume everything.
function showGlobalSnoozeMenu() {
  const items = []
  if (isGlobalSnoozed()) {
    items.push({ label: 'Resume all notifications', click: () => setGlobalSnooze(0) })
  } else {
    items.push({ label: 'Snooze all & mute sound', enabled: false })
    items.push({ label: 'For 1 hour', click: () => setGlobalSnooze(Date.now() + 60 * 60000) })
    items.push({ label: 'For 4 hours', click: () => setGlobalSnooze(Date.now() + 240 * 60000) })
    items.push({ label: 'Until tomorrow', click: () => {
      const t = new Date(); t.setHours(24, 0, 0, 0)
      setGlobalSnooze(t.getTime())
    } })
  }
  Menu.buildFromTemplate(items).popup({ window: panelWindow })
}

/* ---------- Lifecycle ---------- */
// Only one instance may run: a second process would fight over persistent
// browser session partitions, corrupting the File System LOCK and spamming the
// quota/sandbox-DB errors. A second launch just surfaces the existing window.
const gotSingleInstanceLock = app.requestSingleInstanceLock()
if (!gotSingleInstanceLock) {
  app.quit()
}

app.on('second-instance', () => {
  if (panelWindow) {
    showPanelWindow()
  } else {
    togglePanelWindow(true)
  }
})

app.whenReady().then(() => {
  if (!gotSingleInstanceLock) return
  app.setName(APP_NAME)
  if (typeof app.setAppUserModelId === 'function') {
    app.setAppUserModelId('com.mailstudio.mailstudio')
  }
  // Keep the Dock icon visible so the unread count shows as a Dock badge
  // (macOS "app notification number"). The app still lives in the menu bar too.
  // In dev the `electron` CLI launches the unpackaged binary as a child of node,
  // which LaunchServices classifies as a UIElement agent (no Dock icon); calling
  // dock.show() forces a regular activation policy. No-op for the packaged app.
  if (process.platform === 'darwin' && app.dock && typeof app.dock.show === 'function') {
    app.dock.show()
  }

  settings = store.load()
  firstBoot = Boolean(settings.firstBoot)
  buildAllowedHosts()
  buildAppMenu()
  configureSession()
  registerIpc()

  // Restore API connections from the encrypted vault; status changes (connect,
  // refresh failures) push a fresh snapshot so the account rail stays live.
  connections.init({
    config: settings.connections,
    partitionForProvider,
    onChange: () => pushSnapshot()
  })

  for (const service of settings.services) {
    ensureServiceState(service)
    createServiceView(service)
  }
  syncServiceViews()

  createTray()
  createPanelWindow()
  createMenuWindow()
  startFeedTimer()
  startReaperTimer()
  showService(activeServiceKey)
  // Once the active tab has settled, prewarm the rest in the background so every
  // tab's notifications and sidebar go live without being clicked — staggered to
  // avoid spawning all renderers at once.
  setTimeout(startPrewarmQueue, PREWARM_BOOT_DELAY_MS)

  // Check GitHub Releases for updates (packaged builds only; no-op in dev).
  updater.init({ app })
})

app.on('activate', () => {
  if (panelWindow && !panelWindow.isVisible()) {
    showPanelWindow()
  }
})

app.on('before-quit', () => {
  app.isQuitting = true
})

app.on('window-all-closed', (event) => {
  event.preventDefault()
})
