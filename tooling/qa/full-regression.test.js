'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')

const ROOT = path.resolve(__dirname, '..', '..')
const MAIN_JS = fs.readFileSync(path.join(ROOT, 'src', 'main', 'main.js'), 'utf8')
const PANEL_JS = fs.readFileSync(path.join(ROOT, 'src', 'renderer', 'panel.js'), 'utf8')
const PANEL_HTML = fs.readFileSync(path.join(ROOT, 'src', 'renderer', 'panel.html'), 'utf8')

// settings-store requires electron's `app`.
const electronPath = require.resolve('electron')
require.cache[electronPath] = {
  id: electronPath,
  filename: electronPath,
  loaded: true,
  exports: { app: { getPath: () => '/tmp/mailstudio-test' } }
}

const store = require(path.join(ROOT, 'src', 'main', 'settings-store.js'))
const apiFeeds = require(path.join(ROOT, 'src', 'main', 'api-feeds.js'))

function withMockedFetch(mock, fn) {
  const prevFetch = global.fetch
  const prevAbortSignal = global.AbortSignal
  global.fetch = mock
  global.AbortSignal = { timeout: () => undefined }
  return Promise.resolve()
    .then(fn)
    .finally(() => {
      global.fetch = prevFetch
      global.AbortSignal = prevAbortSignal
    })
}

function mockJsonResponse(status, body, headers) {
  const bag = new Map(Object.entries(headers || {}))
  return {
    status,
    ok: status >= 200 && status < 300,
    headers: { get: (key) => bag.get(String(key).toLowerCase()) || null },
    text: async () => JSON.stringify(body)
  }
}

test('settings normalize persists the expanded QoL surface', () => {
  const settings = store.normalize({
    notif: {
      quietStart: '22:00',
      quietEnd: '08:00',
      quietWeekends: true,
      quietAllowCalendar: true
    },
    feedPrefs: {
      mailTodayOnly: true,
      tasksTodayOnly: true,
      hidePreviews: true
    },
    downloads: {
      rememberHistory: false,
      clearOnQuit: true
    },
    layout: {
      activeServiceKey: 'calendar',
      splitKeys: ['mail', 'calendar'],
      splitOrientation: 'horizontal',
      splitRatio: 0.7,
      zoomLevels: { mail: 1.5 }
    },
    workspaces: [{ id: 'w1', name: 'Focus', splitKeys: ['mail', 'calendar'] }],
    recentItems: [{ id: 'r1', title: 'Recent', url: 'https://example.com' }],
    downloadHistory: [{ id: 'd1', title: 'Archive', url: 'https://example.com/file.zip' }]
  })

  assert.equal(settings.notif.quietWeekends, true)
  assert.equal(settings.notif.quietAllowCalendar, true)
  assert.equal(settings.feedPrefs.hidePreviews, true)
  assert.equal(settings.downloads.rememberHistory, false)
  assert.equal(settings.downloads.clearOnQuit, true)
  assert.equal(settings.layout.activeServiceKey, 'calendar')
  assert.deepEqual(settings.layout.splitKeys, ['mail', 'calendar'])
  assert.equal(settings.layout.splitOrientation, 'horizontal')
  assert.equal(settings.layout.splitRatio, 0.7)
  assert.equal(settings.layout.zoomLevels.mail, 1.5)
  assert.equal(settings.workspaces[0].name, 'Focus')
  assert.equal(settings.recentItems[0].title, 'Recent')
  assert.equal(settings.downloadHistory[0].title, 'Archive')
})

test('fetchMail shapes Outlook-owned deep links for inbox items', async () => {
  await withMockedFetch(async () => mockJsonResponse(200, {
    value: [{
      id: 'AAMkAGI2TAAA=',
      subject: 'Status update',
      bodyPreview: 'Body preview',
      receivedDateTime: '2026-06-15T10:00:00Z',
      webLink: 'https://outlook.office.com/mail/deeplink/read/foo',
      from: { emailAddress: { name: 'Casey' } }
    }]
  }), async () => {
    const result = await apiFeeds.fetchMail('token')
    assert.equal(result.state, 'ok')
    assert.equal(result.items.length, 1)
    assert.equal(result.items[0].sender, 'Casey')
    assert.equal(
      result.items[0].deepLink,
      'https://outlook.office.com/mail/inbox/id/AAMkAGI2TAAA%3D'
    )
  })
})

test('fetchAsanaTasks preserves a per-task permalink for every API item', async () => {
  await withMockedFetch(async () => mockJsonResponse(200, {
    data: [
      { gid: '1', name: 'First task', permalink_url: 'https://app.asana.com/0/1/1', due_on: null },
      { gid: '2', name: 'Second task', permalink_url: 'https://app.asana.com/0/1/2', due_on: '2026-06-16' }
    ]
  }), async () => {
    const result = await apiFeeds.fetchAsanaTasks('token', 'workspace')
    assert.equal(result.state, 'ok')
    assert.equal(result.items.length, 2)
    assert.equal(result.items[0].taskUrl, 'https://app.asana.com/0/1/1')
    assert.equal(result.items[1].taskUrl, 'https://app.asana.com/0/1/2')
    assert.notEqual(result.items[0].taskUrl, result.items[1].taskUrl)
  })
})

test('renderer feed clicks pass stable item identifiers for mail, calendar, and asana', () => {
  const itemIdCount = (PANEL_JS.match(/itemId: item\.id \|\| null/g) || []).length
  assert.equal(itemIdCount, 3)
  assert.match(PANEL_JS, /deepLink: item\.deepLink \|\| null/)
  assert.match(PANEL_JS, /taskUrl: item\.taskUrl \|\| null/)
})

test('main feed click handler resolves clicked items by stable id before row fallback', () => {
  assert.match(MAIN_JS, /function findFeedItem\(feed, command\)/)
  assert.match(MAIN_JS, /const recentItem = findFeedItem\(feed, command\)/)
  assert.match(MAIN_JS, /stringOrNull\(command\.deepLink\)/)
  assert.match(MAIN_JS, /stringOrNull\(command\.taskUrl\)/)
  assert.match(MAIN_JS, /stringOrNull\(recentItem && recentItem\.taskUrl\)/)
})

test('settings pages declared in HTML match PAGE_TITLES entries in the renderer', () => {
  const htmlPages = [...PANEL_HTML.matchAll(/id="set-page-([^"]+)"/g)].map((m) => m[1]).sort()
  const pageTitlesBlock = PANEL_JS.match(/const PAGE_TITLES = \{([\s\S]*?)\n\}/)
  assert.ok(pageTitlesBlock, 'PAGE_TITLES block missing')
  const jsPages = [...pageTitlesBlock[1].matchAll(/^\s*'?(?<key>[a-z-]+)'?:\s'[^']+',?$/gm)]
    .map((m) => m.groups.key)
    .sort()
  assert.deepEqual(htmlPages, jsPages)
})

test('top-level renderer commands are all handled in main', () => {
  const explicitCommands = new Set(
    [...PANEL_JS.matchAll(/type:\s'([^']+)'/g)]
      .map((m) => m[1])
      .filter((name) => !name.startsWith('download-'))
  )
  const handledCommands = new Set([...MAIN_JS.matchAll(/case\s'([^']+)':/g)].map((m) => m[1]))
  for (const cmd of explicitCommands) {
    assert.ok(handledCommands.has(cmd), `missing IPC handler for ${cmd}`)
  }
})

test('non-secret settings export excludes OAuth connection config and secrets', () => {
  const exportBlock = MAIN_JS.match(/const portable = \{[\s\S]*?settings: \{([\s\S]*?)\n\s*\}\n\s*\}/)
  assert.ok(exportBlock, 'portable export block missing')
  const block = exportBlock[1]
  assert.ok(!block.includes('connections'))
  assert.ok(!block.includes('clientSecret'))
  assert.ok(!block.includes('secureStore'))
})

test('imported layouts are reattached and snapshots refreshed immediately', () => {
  const importBlock = MAIN_JS.match(/function importPortableSettings\(\) \{([\s\S]*?)\n\}/)
  assert.ok(importBlock, 'importPortableSettings missing')
  const block = importBlock[1]
  assert.ok(block.includes('restorePersistedLayout()'))
  assert.ok(block.includes('attachServiceView()'))
  assert.ok(block.includes('pushSnapshot()'))
})

test('applying a saved workspace persists it as the current restored layout', () => {
  const workspaceBlock = MAIN_JS.match(/function applyWorkspace\(id\) \{([\s\S]*?)\n\}/)
  assert.ok(workspaceBlock, 'applyWorkspace missing')
  assert.ok(workspaceBlock[1].includes('persistLayout()'))
})

test('transient overlays have both open and close handlers so BrowserViews can detach and restore', () => {
  assert.match(MAIN_JS, /case 'open-transient-overlay':/)
  assert.match(MAIN_JS, /case 'close-transient-overlay':/)
  assert.match(PANEL_JS, /window\.panelApi\.sendCommand\(\{ type: 'open-transient-overlay' \}\)/)
  assert.match(PANEL_JS, /window\.panelApi\.sendCommand\(\{ type: 'close-transient-overlay' \}\)/)
})
