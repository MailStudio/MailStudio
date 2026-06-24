'use strict'

const fs = require('fs')
const path = require('path')

const ROOT = path.resolve(__dirname, '..', '..')
const OUT = path.join(__dirname, 'panel-browser-harness.generated.html')

function read(rel) {
  return fs.readFileSync(path.join(ROOT, rel), 'utf8')
}

const snapshot = {
  appName: 'MailStudio',
  activeServiceKey: 'mail',
  splitKeys: [],
  splitOrientation: 'vertical',
  splitRatio: 0.5,
  sidebarCollapsed: false,
  sidebarExpandedWidth: 280,
  collapseMode: 'vanish',
  uiDensity: 'comfortable',
  debugging: { enabled: false },
  taskProvider: 'microsoft',
  notif: { mail: true, calendar: true, asana: true, teams: true, preview: true, quietStart: '', quietEnd: '' },
  feedPrefs: {},
  workspaces: [{ id: 'ws-focus', name: 'Focus stack', icon: 'F', color: '#3b82f6', activeServiceKey: 'mail', splitKeys: ['mail', 'calendar'] }],
  recentItems: [{ id: 'recent-mail', kind: 'routed-link', title: 'Budget thread', subtitle: 'Mail', url: 'https://outlook.office.com/mail/inbox/id/abc', serviceKey: 'mail', at: Date.now() }],
  downloadHistory: [{ id: 'recent-download', kind: 'download', title: 'report.pdf', subtitle: 'Downloads', url: 'file:///tmp/report.pdf', serviceKey: 'mail', at: Date.now() - 1000 }],
  downloadPrefs: { rememberHistory: true, clearOnQuit: false },
  diagnostics: {
    debugging: { enabled: false },
    networkOnline: true,
    prewarm: { queued: 1, active: false, currentKey: '', loaded: 3, maxLoaded: 6 },
    services: [
      { key: 'mail', label: 'Mail', visible: true, loaded: true, hibernated: false, snoozed: false, feedKind: 'mail', feedState: 'ok', feedItems: 2, source: 'api', lastRefreshAt: Date.now(), lastError: '', stale: false, prewarmState: 'loaded', visibleState: 'onscreen' },
      { key: 'teams', label: 'Teams', visible: true, loaded: true, hibernated: false, snoozed: false, feedKind: '', feedState: '', feedItems: 0, source: '', lastRefreshAt: 0, lastError: '', stale: false, prewarmState: 'queued', visibleState: 'background' },
      { key: 'asana', label: 'Asana', visible: true, loaded: true, hibernated: false, snoozed: false, feedKind: 'asana', feedState: 'ok', feedItems: 1, source: 'scraper', lastRefreshAt: Date.now(), lastError: '', stale: false, prewarmState: 'sleeping', visibleState: 'background' }
    ],
    sessionPartitions: [
      { kind: 'microsoft', label: 'Microsoft', services: ['mail', 'calendar', 'teams'], recentFailures: 1 },
      { kind: 'asana', label: 'Asana', services: ['asana'], recentFailures: 0 }
    ],
    microsoftAuth: {
      loginHint: 'qa@example.com',
      webSession: true,
      states: [
        { key: 'mail', label: 'Mail', authState: 'app', loaded: true, visible: true, href: 'https://outlook.office.com/mail/' },
        { key: 'teams', label: 'Teams', authState: 'login', loaded: true, visible: true, href: 'https://login.microsoftonline.com/' }
      ],
      events: [
        { id: 'auth-1', type: 'silent-failed', stage: 'silent', prompt: 'none', oauthError: 'interaction_required', message: 'Multiple accounts available', at: Date.now() - 30000 },
        { id: 'auth-2', type: 'soft-repair', stage: 'timer', prompt: '', oauthError: '', message: 'Teams', at: Date.now() - 20000 }
      ]
    },
    notifications: {
      baselines: { mailboxes: 2, mailIds: 4, asanaIds: 1, calendarIds: 1, teamsReady: true, teamsLastCount: 1 },
      cooldowns: [{ key: 'teams:title-count', remainingMs: 5000 }],
      history: [
        { id: 'notif-1', kind: 'notification-shown', title: 'Mail: New email', subtitle: 'mail · Budget thread', serviceKey: 'mail', at: Date.now() - 15000 },
        { id: 'notif-2', kind: 'notification-suppressed', title: 'Teams: New Teams message', subtitle: 'teams · focused', serviceKey: 'teams', at: Date.now() - 12000 }
      ]
    },
    serviceFailureLog: [
      { id: 'failure-1', kind: 'service-failure', title: 'Teams page failed', subtitle: 'blank: No visible content', url: 'https://teams.cloud.microsoft/', serviceKey: 'teams', at: Date.now() - 60000 }
    ],
    connections: { microsoft: { status: 'connected' }, asana: { status: 'connected' } }
  },
  setupHealth: {
    status: 'attention',
    failed: 1,
    checks: [
      { key: 'microsoft', label: 'Microsoft account', ok: true, detail: 'Connected' },
      { key: 'notifications', label: 'Notification arming', ok: false, detail: 'Connected providers will stay quiet until setup is enabled.', action: 'notifications' }
    ]
  },
  firstBoot: false,
  onboarded: false,
  notifSetupSkipped: false,
  connections: {
    microsoft: { status: 'connected', account: { email: 'qa@example.com', name: 'QA User' } },
    asana: { status: 'connected', account: { email: 'qa@example.com', name: 'QA User' } },
    encryptionAvailable: true
  },
  connConfig: { microsoft: { clientId: '', tenant: 'common' }, asana: { clientId: '' } },
  scratch: '',
  settingsOpen: false,
  theme: 'dark',
  glassMode: false,
  nav: { canGoBack: false, canGoForward: false },
  globalSnoozed: false,
  feedCollapsed: {},
  services: [
    {
      key: 'teams',
      label: 'Teams',
      icon: 'teams',
      url: 'https://teams.cloud.microsoft/',
      home: 'https://teams.cloud.microsoft/',
      builtin: true,
      visible: true,
      unreadCount: 1,
      feedCollapsed: false
    },
    {
      key: 'mail',
      label: 'Mail',
      icon: 'mail',
      url: 'https://outlook.office.com/mail/',
      home: 'https://outlook.office.com/mail/',
      builtin: true,
      visible: true,
      unreadCount: 2,
      feedCollapsed: false,
      health: { state: 'ok', message: '', code: 0, at: Date.now() },
      feed: {
        kind: 'mail',
        state: 'ok',
        items: [
          { id: 'm1', sender: 'A', subject: 'One', preview: 'P', today: true, isRead: false },
          { id: 'm2', sender: 'B', subject: 'Two', preview: 'Q', today: true, isRead: true }
        ]
      }
    },
    {
      key: 'shared',
      label: 'Shared',
      icon: 'mail',
      url: 'https://outlook.office.com/mail/shared/',
      home: 'https://outlook.office.com/mail/shared/',
      builtin: false,
      visible: true,
      mailboxManaged: true,
      unreadCount: 3,
      feedCollapsed: false,
      feed: {
        kind: 'mail',
        state: 'ok',
        items: [
          { id: 's1', sender: 'C', subject: 'Three', preview: 'R', today: true, isRead: false },
          { id: 's2', sender: 'D', subject: 'Four', preview: 'S', today: true, isRead: false },
          { id: 'old', sender: 'E', subject: 'Old', preview: 'T', today: false, isRead: false }
        ]
      }
    },
    {
      key: 'calendar',
      label: 'Calendar',
      icon: 'calendar',
      url: 'https://outlook.office.com/calendar/',
      home: 'https://outlook.office.com/calendar/',
      builtin: true,
      visible: true,
      unreadCount: 0,
      feedCollapsed: false,
      feed: { kind: 'calendar', state: 'ok', items: [{ id: 'c1', title: 'Standup', time: '9:30 AM' }] }
    },
    {
      key: 'asana',
      label: 'Asana',
      icon: 'asana',
      url: 'https://app.asana.com/',
      home: 'https://app.asana.com/',
      builtin: true,
      visible: true,
      unreadCount: 1,
      feedCollapsed: false,
      feed: { kind: 'asana', state: 'ok', items: [{ id: 'a1', name: 'Task', dueOn: null }] }
    }
  ]
}

const stub = `
window.__commands = [];
window.__snapshotHandlers = [];
window.__eventHandlers = [];
window.__downloadHandlers = [];
window.__findHandlers = [];
window.__snapshot = ${JSON.stringify(snapshot)};
window.__downloads = {
  activeCount: 1,
  list: [
    {
      id: 1,
      filename: 'report.pdf',
      state: 'progressing',
      receivedBytes: 512000,
      totalBytes: 1024000,
      speed: 64000,
      startTime: Date.now()
    }
  ]
};
window.__recordCommand = (command) => {
  window.__commands.push(command);
  document.documentElement.dataset.commands = JSON.stringify(window.__commands);
  if (command && (command.type === 'open-settings' || command.type === 'close-settings')) {
    window.__snapshot = { ...window.__snapshot, settingsOpen: command.type === 'open-settings' };
    for (const callback of window.__snapshotHandlers) callback(window.__snapshot);
  }
};
window.panelApi = {
  getSnapshot: () => Promise.resolve(window.__snapshot),
  sendCommand: window.__recordCommand,
  onStatusUpdated: (callback) => { window.__snapshotHandlers.push(callback); },
  onEvent: (callback) => { window.__eventHandlers.push(callback); },
  onDownloadsUpdated: (callback) => {
    window.__downloadHandlers.push(callback);
    callback(window.__downloads);
  },
  onFindResult: (callback) => { window.__findHandlers.push(callback); }
};
window.__emitSnapshot = (snapshot) => {
  window.__snapshot = snapshot;
  for (const callback of window.__snapshotHandlers) callback(snapshot);
};
window.__emitEvent = (data) => {
  for (const callback of window.__eventHandlers) callback(data);
  if (data && data.type === 'find-result') {
    for (const callback of window.__findHandlers) callback(data);
  }
};
window.__emitDownloads = (data) => {
  for (const callback of window.__downloadHandlers) callback(data);
};
`

let html = read('src/renderer/panel.html')
html = html
  .replace(/<meta http-equiv="Content-Security-Policy"[^>]+>/, '')
  .replace('<link rel="stylesheet" href="./panel.css" />', `<style>${read('src/renderer/panel.css')}</style>`)
  .replace('<script src="./panel.js"></script>', `<script>${stub}</script><script>${read('src/renderer/panel.js')}</script>`)

fs.writeFileSync(OUT, html, 'utf8')
console.log(OUT)
