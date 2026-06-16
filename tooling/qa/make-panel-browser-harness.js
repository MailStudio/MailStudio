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
  taskProvider: 'microsoft',
  notif: { mail: true, calendar: true, asana: true, teams: true, preview: true, quietStart: '', quietEnd: '' },
  feedPrefs: {},
  workspaces: [],
  recentItems: [],
  downloadHistory: [],
  downloadPrefs: { rememberHistory: true, clearOnQuit: false },
  diagnostics: {
    services: [],
    connections: { microsoft: { status: 'disconnected' }, asana: { status: 'disconnected' } }
  },
  firstBoot: false,
  onboarded: false,
  notifSetupSkipped: false,
  connections: {
    microsoft: { status: 'disconnected' },
    asana: { status: 'disconnected' },
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
      key: 'mail',
      label: 'Mail',
      icon: 'mail',
      url: 'https://outlook.office.com/mail/',
      home: 'https://outlook.office.com/mail/',
      builtin: true,
      visible: true,
      unreadCount: 2,
      feedCollapsed: false,
      feed: {
        kind: 'mail',
        state: 'ok',
        items: [
          { id: 'm1', sender: 'A', subject: 'One', preview: 'P', today: true },
          { id: 'm2', sender: 'B', subject: 'Two', preview: 'Q', today: true }
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
          { id: 's1', sender: 'C', subject: 'Three', preview: 'R', today: true },
          { id: 's2', sender: 'D', subject: 'Four', preview: 'S', today: true },
          { id: 'old', sender: 'E', subject: 'Old', preview: 'T', today: false }
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
      id: 'dl1',
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
