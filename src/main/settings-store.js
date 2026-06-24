const fs = require('fs')
const path = require('path')
const { app } = require('electron')

// Built-in services. `feed` enables the live sidebar section (scraped from the
// logged-in web view). `home` is the pinned URL the Home button returns to.
const DEFAULT_SERVICES = [
  {
    key: 'teams',
    label: 'Teams',
    url: 'https://teams.cloud.microsoft/',
    home: 'https://teams.cloud.microsoft/',
    icon: 'teams',
    builtin: true,
    visible: true
  },
  {
    key: 'mail',
    label: 'Mail',
    url: 'https://outlook.office.com/mail/',
    home: 'https://outlook.office.com/mail/inbox',
    icon: 'mail',
    builtin: true,
    visible: true,
    feed: 'mail'
  },
  {
    key: 'calendar',
    label: 'Calendar',
    url: 'https://outlook.office.com/calendar/',
    home: 'https://outlook.office.com/calendar/',
    icon: 'calendar',
    builtin: true,
    visible: true,
    feed: 'calendar'
  },
  {
    key: 'todo',
    label: 'To Do',
    url: 'https://to-do.office.com/tasks/',
    home: 'https://to-do.office.com/tasks/',
    icon: 'todo',
    builtin: true,
    visible: true
  },
  {
    key: 'asana',
    label: 'Asana',
    url: 'https://app.asana.com/',
    home: 'https://app.asana.com/',
    icon: 'asana',
    builtin: true,
    visible: true,
    feed: 'asana'
  },
  // Extended Office suite — hidden by default, enable via Settings.
  // www.office.com now redirects to the Microsoft 365 Copilot home.
  {
    key: 'office',
    label: 'Copilot',
    url: 'https://m365.cloud.microsoft/',
    home: 'https://m365.cloud.microsoft/',
    icon: 'copilot',
    builtin: true,
    visible: false
  },
  {
    key: 'word',
    label: 'Word',
    url: 'https://www.office.com/launch/word',
    home: 'https://www.office.com/launch/word',
    icon: 'word',
    builtin: true,
    visible: false
  },
  {
    key: 'excel',
    label: 'Excel',
    url: 'https://www.office.com/launch/excel',
    home: 'https://www.office.com/launch/excel',
    icon: 'excel',
    builtin: true,
    visible: false
  },
  {
    key: 'powerpoint',
    label: 'PowerPoint',
    url: 'https://www.office.com/launch/powerpoint',
    home: 'https://www.office.com/launch/powerpoint',
    icon: 'powerpoint',
    builtin: true,
    visible: false
  },
  {
    key: 'onenote',
    label: 'OneNote',
    url: 'https://www.office.com/launch/onenote',
    home: 'https://www.office.com/launch/onenote',
    icon: 'onenote',
    builtin: true,
    visible: false
  },
  {
    key: 'onedrive',
    label: 'OneDrive',
    url: 'https://www.office.com/launch/onedrive',
    home: 'https://www.office.com/launch/onedrive',
    icon: 'onedrive',
    builtin: true,
    visible: false
  },
  {
    key: 'planner',
    label: 'Planner',
    url: 'https://planner.cloud.microsoft/',
    home: 'https://planner.cloud.microsoft/',
    icon: 'planner',
    builtin: true,
    visible: false
  },
  {
    key: 'sharepoint',
    label: 'SharePoint',
    url: 'https://www.office.com/launch/sharepoint',
    home: 'https://www.office.com/launch/sharepoint',
    icon: 'sharepoint',
    builtin: true,
    visible: false
  }
]

const DEFAULTS = {
  theme: 'dark',
  uiDensity: 'comfortable',
  debugging: {
    enabled: false
  },
  collapseMode: 'vanish', // 'vanish' (fully hidden) | 'rail' (icon strip)
  taskProvider: 'microsoft', // where "New task" composes: 'microsoft' (To Do) | 'asana'
  firstBoot: true, // cleared after first successful Microsoft sign-in or skip
  onboarded: false, // notifications stay silent until the user finishes onboarding
  notifSetupSkipped: false, // user dismissed the "Set up notifications" nudge
  scratch: '', // sidebar scratch-pad contents
  notif: {
    mail: true, // Graph — new email
    calendar: true, // Graph — upcoming event reminders
    asana: true, // Asana — newly assigned tasks
    teams: true, // title watcher — new Teams messages
    preview: true,
    quietStart: '', // "HH:MM" 24-hour, or '' to disable
    quietEnd: ''
  },
  notificationState: {
    mail: {},
    asana: { ready: false, ids: [] },
    calendar: { ids: [] },
    teams: { ready: false, lastCount: 0 },
    history: []
  },
  // Bring-your-own OAuth app registrations. clientIds are NOT secrets (PKCE
  // flow), so they live in plain settings; tokens live in the encrypted vault.
  connections: {
    microsoft: { clientId: '', tenant: 'common' },
    asana: { clientId: '' }
  },
  // Per-service feed (notification preview) collapse — true = hidden.
  feedCollapsed: {},
  feedPrefs: {
    mailTodayOnly: false,
    tasksTodayOnly: false,
    hidePreviews: false
  },
  layout: {
    activeServiceKey: 'mail',
    splitKeys: [],
    splitOrientation: 'vertical',
    splitRatio: 0.5,
    zoomLevels: {}
  },
  workspaces: [],
  recentItems: [],
  downloadHistory: [],
  serviceFailureLog: [],
  downloads: {
    rememberHistory: true,
    clearOnQuit: false
  },
  // User-dragged sidebar width in px (clamped to [180, 480]).
  sidebarWidth: 280,
  services: DEFAULT_SERVICES
}

const VALID_MANAGED_FEEDS = new Set(['mail'])
const MANAGED_MAILBOX_HOSTS = new Set([
  'outlook.office.com',
  'outlook.office365.com',
  'outlook.live.com'
])
const MAX_CUSTOM_SERVICES = 30

function filePath() {
  return path.join(app.getPath('userData'), 'mailstudio-settings.json')
}

function clone(value) {
  return JSON.parse(JSON.stringify(value))
}

function isManagedMailboxUrl(value) {
  try {
    const host = new URL(value).hostname.toLowerCase()
    return MANAGED_MAILBOX_HOSTS.has(host) || host.endsWith('.outlook.office.com') || host.endsWith('.outlook.office365.com')
  } catch {
    return false
  }
}

function cleanServiceKey(value) {
  return (typeof value === 'string' ? value : '')
    .replace(/[\u0000-\u001F\u007F]/g, '')
    .trim()
    .slice(0, 80)
}

function sanitizeService(input, builtinDefaults) {
  if (!input || typeof input.url !== 'string') {
    return null
  }

  let url
  try {
    const parsed = new URL(input.url)
    // Only http(s) sites may be pinned — blocks javascript:, file:, data:, etc.
    if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
      return null
    }
    url = parsed.toString()
  } catch {
    return null
  }

  const cleanKey = cleanServiceKey(input.key)
  const builtin = builtinDefaults.get(cleanKey)
  const cleanLabel = (value) =>
    String(value || '')
      .replace(/[\u200B-\u200D\u2060\uFEFF\uFFFD\uE000-\uF8FF\u25A0-\u25A1]/g, '')
      .trim()
      .slice(0, 80)
  let home = url
  if (builtin) {
    home = builtin.home
  } else {
    try {
      const h = new URL(input.home || input.url)
      if (h.protocol === 'http:' || h.protocol === 'https:') home = h.toString()
    } catch {
      home = url
    }
  }
  const managedMailbox = Boolean(
    !builtin &&
    input.mailboxManaged &&
    VALID_MANAGED_FEEDS.has(input.feed) &&
    isManagedMailboxUrl(url)
  )
  if (managedMailbox && !isManagedMailboxUrl(home)) {
    home = url
  }

  return {
    key: builtin ? builtin.key : (cleanKey || `site-${Math.abs(hash(url))}`),
    // Built-in labels are fixed to the default (they can't be renamed in the UI),
    // so a renamed default (e.g. Office → Copilot) reaches existing users. Only
    // pinned custom sites keep a user-supplied label.
    label: builtin ? builtin.label : (cleanLabel(input.label) || 'Site'),
    // Built-in services have fixed URLs that cannot be overridden via settings.
    // Pinning them to the hardcoded defaults prevents a tampered settings file
    // (or a compromised renderer) from redirecting the Mail tab to a phishing domain.
    url: builtin ? builtin.url : url,
    home,
    icon: builtin ? builtin.icon : (input.icon === 'mail' ? 'mail' : 'link'),
    builtin: Boolean(builtin),
    visible: input.visible !== false,
    ...(builtin && builtin.feed ? { feed: builtin.feed } : {}),
    ...(managedMailbox ? { feed: input.feed, mailboxManaged: true } : {})
  }
}

function hash(str) {
  let h = 0
  for (let i = 0; i < str.length; i += 1) {
    h = (h << 5) - h + str.charCodeAt(i)
    h |= 0
  }
  return h
}

function normalizeTime(value) {
  if (typeof value !== 'string') return ''
  const match = value.trim().match(/^(\d{1,2}):(\d{2})$/)
  if (!match) return ''
  const hours = Number(match[1])
  const minutes = Number(match[2])
  if (!Number.isInteger(hours) || !Number.isInteger(minutes)) return ''
  if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) return ''
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`
}

function normalize(raw) {
  const settings = { ...clone(DEFAULTS), ...(raw || {}) }

  if (settings.theme !== 'light' && settings.theme !== 'dark') {
    settings.theme = DEFAULTS.theme
  }
  if (!['comfortable', 'compact'].includes(settings.uiDensity)) {
    settings.uiDensity = DEFAULTS.uiDensity
  }
  const rawDebugging = (raw && typeof raw.debugging === 'object' && raw.debugging) ? raw.debugging : {}
  settings.debugging = {
    enabled: Boolean(rawDebugging.enabled)
  }
  if (settings.collapseMode !== 'rail' && settings.collapseMode !== 'vanish') {
    settings.collapseMode = DEFAULTS.collapseMode
  }
  if (settings.taskProvider !== 'microsoft' && settings.taskProvider !== 'asana') {
    settings.taskProvider = DEFAULTS.taskProvider
  }

  // firstBoot: true only on genuine fresh install (normalize(null)); false for existing users
  settings.firstBoot = raw === null ? true : Boolean(raw.firstBoot)
  settings.onboarded = Boolean(raw && raw.onboarded)
  settings.notifSetupSkipped = Boolean(raw && raw.notifSetupSkipped)

  settings.scratch = typeof (raw && raw.scratch) === 'string' ? raw.scratch.slice(0, 20000) : ''

  const rawSbWidth = raw && typeof raw.sidebarWidth === 'number' ? raw.sidebarWidth : 280
  settings.sidebarWidth = Math.round(Math.min(480, Math.max(180, rawSbWidth)))

  const rawConn = (raw && typeof raw.connections === 'object' && raw.connections) ? raw.connections : {}
  const rawMs = (rawConn.microsoft && typeof rawConn.microsoft === 'object') ? rawConn.microsoft : {}
  const rawAs = (rawConn.asana && typeof rawConn.asana === 'object') ? rawConn.asana : {}
  settings.connections = {
    microsoft: {
      // clientIds are public PKCE identifiers (not secrets); length-cap to prevent
      // excessively large values from being persisted or forwarded to the token endpoint.
      clientId: typeof rawMs.clientId === 'string' ? rawMs.clientId.trim().slice(0, 256) : '',
      tenant: (typeof rawMs.tenant === 'string' && rawMs.tenant.trim().slice(0, 256)) || 'common'
    },
    asana: {
      clientId: typeof rawAs.clientId === 'string' ? rawAs.clientId.trim().slice(0, 256) : ''
    }
  }

  const rawNotif = (raw && typeof raw.notif === 'object' && raw.notif) ? raw.notif : {}
  settings.notif = {
    mail: rawNotif.mail !== false,
    calendar: rawNotif.calendar !== false,
    asana: rawNotif.asana !== false,
    teams: rawNotif.teams !== false,
    preview: rawNotif.preview !== false,
    quietStart: normalizeTime(rawNotif.quietStart),
    quietEnd: normalizeTime(rawNotif.quietEnd),
    quietWeekends: Boolean(rawNotif.quietWeekends),
    quietAllowCalendar: Boolean(rawNotif.quietAllowCalendar)
  }

  const rawFeedPrefs = (raw && typeof raw.feedPrefs === 'object' && raw.feedPrefs) ? raw.feedPrefs : {}
  settings.feedPrefs = {
    mailTodayOnly: Boolean(rawFeedPrefs.mailTodayOnly),
    tasksTodayOnly: Boolean(rawFeedPrefs.tasksTodayOnly),
    hidePreviews: Boolean(rawFeedPrefs.hidePreviews)
  }

  const rawDownloads = (raw && typeof raw.downloads === 'object' && raw.downloads) ? raw.downloads : {}
  settings.downloads = {
    rememberHistory: rawDownloads.rememberHistory !== false,
    clearOnQuit: Boolean(rawDownloads.clearOnQuit)
  }

  const rawLayout = (raw && typeof raw.layout === 'object' && raw.layout) ? raw.layout : {}
  const rawSplitKeys = Array.isArray(rawLayout.splitKeys) ? rawLayout.splitKeys : []
  const rawZoom = (rawLayout && typeof rawLayout.zoomLevels === 'object' && rawLayout.zoomLevels) ? rawLayout.zoomLevels : {}
  settings.layout = {
    activeServiceKey: typeof rawLayout.activeServiceKey === 'string' ? rawLayout.activeServiceKey.slice(0, 80) : 'mail',
    splitKeys: rawSplitKeys.filter((k) => typeof k === 'string').slice(0, 2),
    splitOrientation: rawLayout.splitOrientation === 'horizontal' ? 'horizontal' : 'vertical',
    splitRatio: typeof rawLayout.splitRatio === 'number'
      ? Math.min(0.85, Math.max(0.15, rawLayout.splitRatio))
      : 0.5,
    zoomLevels: Object.fromEntries(
      Object.entries(rawZoom)
        .filter(([k, v]) => typeof k === 'string' && typeof v === 'number' && Number.isFinite(v))
        .map(([k, v]) => [k, Math.min(3, Math.max(-2.5, v))])
    )
  }

  const cleanRecents = (items, limit = 50) => (Array.isArray(items) ? items : [])
    .filter((item) => item && typeof item === 'object')
    .map((item) => ({
      id: typeof item.id === 'string' ? item.id.slice(0, 160) : `recent-${Math.abs(hash(JSON.stringify(item))).toString(36)}`,
      kind: typeof item.kind === 'string' ? item.kind.slice(0, 32) : 'link',
      title: typeof item.title === 'string' ? item.title.slice(0, 180) : 'Item',
      subtitle: typeof item.subtitle === 'string' ? item.subtitle.slice(0, 220) : '',
      url: typeof item.url === 'string' ? item.url.slice(0, 2000) : '',
      serviceKey: typeof item.serviceKey === 'string' ? item.serviceKey.slice(0, 80) : '',
      at: typeof item.at === 'number' ? item.at : Date.now(),
      code: typeof item.code === 'number' ? item.code : 0
    }))
    .slice(0, limit)
  settings.recentItems = cleanRecents(raw && raw.recentItems)
  settings.downloadHistory = cleanRecents(raw && raw.downloadHistory)
  settings.serviceFailureLog = cleanRecents(raw && raw.serviceFailureLog, 80)
  const cleanIdList = (items, limit = 2000) => (Array.isArray(items) ? items : [])
    .filter((item) => typeof item === 'string' && item)
    .map((item) => item.slice(0, 240))
    .slice(0, limit)
  const rawNotificationState = (raw && typeof raw.notificationState === 'object' && raw.notificationState) ? raw.notificationState : {}
  const rawMailState = (rawNotificationState.mail && typeof rawNotificationState.mail === 'object') ? rawNotificationState.mail : {}
  settings.notificationState = {
    mail: Object.fromEntries(
      Object.entries(rawMailState)
        .filter(([key, value]) => typeof key === 'string' && value && typeof value === 'object')
        .slice(0, 60)
        .map(([key, value]) => [
          cleanServiceKey(key),
          {
            ready: Boolean(value.ready),
            lastCount: Math.max(0, Math.min(9999, Number.isFinite(value.lastCount) ? Math.trunc(value.lastCount) : 0)),
            ids: cleanIdList(value.ids)
          }
        ])
        .filter(([key]) => key)
    ),
    asana: {
      ready: Boolean(rawNotificationState.asana && rawNotificationState.asana.ready),
      ids: cleanIdList(rawNotificationState.asana && rawNotificationState.asana.ids)
    },
    calendar: {
      ids: cleanIdList(rawNotificationState.calendar && rawNotificationState.calendar.ids)
    },
    teams: {
      ready: Boolean(rawNotificationState.teams && rawNotificationState.teams.ready),
      lastCount: Math.max(0, Math.min(9999, Number.isFinite(rawNotificationState.teams && rawNotificationState.teams.lastCount) ? Math.trunc(rawNotificationState.teams.lastCount) : 0))
    },
    history: cleanRecents(rawNotificationState.history, 120)
  }

  settings.workspaces = (Array.isArray(raw && raw.workspaces) ? raw.workspaces : [])
    .filter((item) => item && typeof item === 'object')
    .map((item) => ({
      id: typeof item.id === 'string' ? item.id.slice(0, 80) : `workspace-${Math.abs(hash(JSON.stringify(item))).toString(36)}`,
      name: typeof item.name === 'string' && item.name.trim() ? item.name.trim().slice(0, 80) : 'Workspace',
      icon: typeof item.icon === 'string' && /^[A-Za-z0-9]{1,3}$/.test(item.icon.trim()) ? item.icon.trim().slice(0, 3) : '',
      color: typeof item.color === 'string' && /^#[0-9a-f]{6}$/i.test(item.color) ? item.color.toLowerCase() : '#3b82f6',
      activeServiceKey: typeof item.activeServiceKey === 'string' ? item.activeServiceKey.slice(0, 80) : 'mail',
      splitKeys: Array.isArray(item.splitKeys) ? item.splitKeys.filter((k) => typeof k === 'string').slice(0, 2) : [],
      splitOrientation: item.splitOrientation === 'horizontal' ? 'horizontal' : 'vertical',
      splitRatio: typeof item.splitRatio === 'number' ? Math.min(0.85, Math.max(0.15, item.splitRatio)) : 0.5,
      sidebarCollapsed: Boolean(item.sidebarCollapsed),
      collapseMode: item.collapseMode === 'rail' ? 'rail' : 'vanish',
      services: Array.isArray(item.services)
        ? item.services.map((s) => ({ key: String(s.key || '').slice(0, 80), visible: s.visible !== false })).filter((s) => s.key)
        : []
    }))
    .slice(0, 20)

  const builtinDefaults = new Map(DEFAULT_SERVICES.map((s) => [s.key, s]))
  const seen = new Set()
  const services = []
  let customCount = 0

  if (Array.isArray(raw && raw.services)) {
    for (const entry of raw.services) {
      const clean = sanitizeService(entry, builtinDefaults)
      if (clean && !clean.builtin && customCount >= MAX_CUSTOM_SERVICES) {
        continue
      }
      if (clean && !seen.has(clean.key)) {
        seen.add(clean.key)
        services.push(clean)
        if (!clean.builtin) customCount += 1
      }
    }
  }

  // Make sure every built-in still exists (built-ins can be hidden but never
  // deleted). Insert a missing one near its default position so newly added
  // built-ins (e.g. Teams at the top) land where intended for existing configs.
  DEFAULT_SERVICES.forEach((builtin, defaultIndex) => {
    if (!seen.has(builtin.key)) {
      services.splice(Math.min(defaultIndex, services.length), 0, clone(builtin))
      seen.add(builtin.key)
    }
  })

  const rawCollapsed = (raw && typeof raw.feedCollapsed === 'object' && raw.feedCollapsed) ? raw.feedCollapsed : {}
  settings.feedCollapsed = {}
  for (const [key, val] of Object.entries(rawCollapsed)) {
    if (typeof key === 'string' && key) settings.feedCollapsed[key] = Boolean(val)
  }

  settings.services = services
  return settings
}

function writeSettingsFile(target, payload) {
  const dir = path.dirname(target)
  const tmp = `${target}.tmp`
  fs.mkdirSync(dir, { recursive: true })
  try { fs.unlinkSync(tmp) } catch { /* file may not exist */ }
  fs.writeFileSync(tmp, payload, { encoding: 'utf8', mode: 0o600 })
  fs.renameSync(tmp, target)
  try {
    fs.chmodSync(target, 0o600)
  } catch {
    /* ignore chmod failures on restrictive filesystems */
  }
}

function load() {
  try {
    const raw = JSON.parse(fs.readFileSync(filePath(), 'utf8'))
    return normalize(raw)
  } catch {
    return normalize(null)
  }
}

function save(settings) {
  const normalized = normalize(settings)
  try {
    // Owner read/write only — settings include OAuth client IDs and pinned
    // intranet URLs that shouldn't be world-readable on shared systems. Write
    // through a temp file + rename so a crash cannot leave truncated JSON.
    writeSettingsFile(filePath(), JSON.stringify(normalized, null, 2))
  } catch {
    // Non-fatal: settings just won't persist this session.
  }
  return normalized
}

module.exports = { load, save, normalize, DEFAULTS, DEFAULT_SERVICES }
