const fs = require('fs')
const path = require('path')
const { app } = require('electron')

// Built-in services. `feed` enables the live sidebar section (scraped from the
// logged-in web view). `home` is the pinned URL the Home button returns to.
const DEFAULT_SERVICES = [
  {
    key: 'teams',
    label: 'Teams',
    url: 'https://teams.microsoft.com/',
    home: 'https://teams.microsoft.com/',
    icon: 'teams',
    builtin: true,
    visible: true
  },
  {
    key: 'mail',
    label: 'Mail',
    url: 'https://outlook.office.com/mail/',
    home: 'https://outlook.office.com/mail/',
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
    url: 'https://planner.microsoft.com/',
    home: 'https://planner.microsoft.com/',
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
  // Bring-your-own OAuth app registrations. clientIds are NOT secrets (PKCE
  // flow), so they live in plain settings; tokens live in the encrypted vault.
  connections: {
    microsoft: { clientId: '', tenant: 'common' },
    asana: { clientId: '' }
  },
  // Per-service feed (notification preview) collapse — true = hidden.
  feedCollapsed: {},
  // User-dragged sidebar width in px (clamped to [180, 480]).
  sidebarWidth: 280,
  services: DEFAULT_SERVICES
}

function filePath() {
  return path.join(app.getPath('userData'), 'mailstudio-settings.json')
}

function clone(value) {
  return JSON.parse(JSON.stringify(value))
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

  const builtin = builtinDefaults.get(input.key)
  return {
    key: typeof input.key === 'string' && input.key ? input.key : `site-${Math.abs(hash(url))}`,
    // Built-in labels are fixed to the default (they can't be renamed in the UI),
    // so a renamed default (e.g. Office → Copilot) reaches existing users. Only
    // pinned custom sites keep a user-supplied label.
    label: builtin ? builtin.label : ((typeof input.label === 'string' && input.label.trim()) || 'Site'),
    // Built-in services have fixed URLs that cannot be overridden via settings.
    // Pinning them to the hardcoded defaults prevents a tampered settings file
    // (or a compromised renderer) from redirecting the Mail tab to a phishing domain.
    url: builtin ? builtin.url : url,
    home: builtin ? builtin.home : (() => {
      try {
        const h = new URL(input.home || input.url)
        if (h.protocol !== 'http:' && h.protocol !== 'https:') return url
        return h.toString()
      } catch {
        return url
      }
    })(),
    icon: builtin ? builtin.icon : (input.icon === 'mail' ? 'mail' : 'link'),
    builtin: Boolean(builtin),
    visible: input.visible !== false,
    ...(builtin && builtin.feed ? { feed: builtin.feed } : {}),
    ...(typeof input.feed === 'string' && input.feed ? { feed: input.feed } : {}),
    ...(input.mailboxManaged ? { mailboxManaged: true } : {})
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

function normalize(raw) {
  const settings = { ...clone(DEFAULTS), ...(raw || {}) }

  if (settings.theme !== 'light' && settings.theme !== 'dark') {
    settings.theme = DEFAULTS.theme
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
    quietStart: typeof rawNotif.quietStart === 'string' ? rawNotif.quietStart.trim() : '',
    quietEnd: typeof rawNotif.quietEnd === 'string' ? rawNotif.quietEnd.trim() : ''
  }

  const builtinDefaults = new Map(DEFAULT_SERVICES.map((s) => [s.key, s]))
  const seen = new Set()
  const services = []

  if (Array.isArray(raw && raw.services)) {
    for (const entry of raw.services) {
      const clean = sanitizeService(entry, builtinDefaults)
      if (clean && !seen.has(clean.key)) {
        seen.add(clean.key)
        services.push(clean)
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
    // mode 0o600: owner read/write only — settings include OAuth client IDs and
    // pinned intranet URLs that shouldn't be world-readable on shared systems.
    fs.writeFileSync(filePath(), JSON.stringify(normalized, null, 2), { encoding: 'utf8', mode: 0o600 })
  } catch {
    // Non-fatal: settings just won't persist this session.
  }
  return normalized
}

module.exports = { load, save, normalize, DEFAULTS, DEFAULT_SERVICES }
