'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')

const ROOT = path.resolve(__dirname, '..', '..')
const MAIN_JS = fs.readFileSync(path.join(ROOT, 'src', 'main', 'main.js'), 'utf8')
const PANEL_JS = fs.readFileSync(path.join(ROOT, 'src', 'renderer', 'panel.js'), 'utf8')
const PANEL_HTML = fs.readFileSync(path.join(ROOT, 'src', 'renderer', 'panel.html'), 'utf8')
const CONNECTIONS_JS = fs.readFileSync(path.join(ROOT, 'src', 'main', 'connections.js'), 'utf8')

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

test('mailbox-managed labels strip Outlook icon glyphs and square placeholders', () => {
  const settings = store.normalize({
    services: [{
      key: 'mailbox',
      label: '\uE000\u25A1support@joinparade.app',
      url: 'https://outlook.office.com/mail/support%40joinparade.app/',
      mailboxManaged: true,
      feed: 'mail'
    }]
  })
  const mailbox = settings.services.find((service) => service.key === 'mailbox')
  assert.ok(mailbox)
  assert.equal(mailbox.label, 'support@joinparade.app')
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
  // The three feed kinds share one dispatch helper, so the stable identifiers are
  // built once and the kind only selects taskUrl (asana) vs webLink (mail/cal).
  assert.match(PANEL_JS, /itemId: item\.id \|\| null/)
  assert.match(PANEL_JS, /deepLink: item\.deepLink \|\| null/)
  assert.match(PANEL_JS, /if \(feed\.kind === 'asana'\) command\.taskUrl = item\.taskUrl \|\| null/)
  assert.match(PANEL_JS, /else command\.webLink = item\.webLink \|\| null/)
})

test('feed rows are keyboard-operable buttons, not mouse-only divs', () => {
  assert.match(PANEL_JS, /row\.setAttribute\('role', 'button'\)/)
  assert.match(PANEL_JS, /row\.tabIndex = 0/)
  assert.match(PANEL_JS, /row\.addEventListener\('keydown'/)
  assert.match(PANEL_JS, /event\.key === 'Enter' \|\| event\.key === ' '/)
})

test('main feed click handler resolves clicked items by stable id before row fallback', () => {
  assert.match(MAIN_JS, /function findFeedItem\(feed, command\)/)
  assert.match(MAIN_JS, /const recentItem = findFeedItem\(feed, command\)/)
  assert.match(MAIN_JS, /stringOrNull\(command\.deepLink\)/)
  assert.match(MAIN_JS, /stringOrNull\(command\.taskUrl\)/)
  assert.match(MAIN_JS, /stringOrNull\(recentItem && recentItem\.taskUrl\)/)
})

test('generic Microsoft 365 URLs no longer forward clicks to the Copilot tab', () => {
  assert.match(MAIN_JS, /const appMatch = route\.match\(\/\\\/launch\\\/\(word\|excel\|powerpoint\|onenote\|onedrive\|sharepoint\)\/i\)/)
  assert.match(MAIN_JS, /if \(appMatch\) return findOrReveal\(appMatch\[1\]\.toLowerCase\(\)\)/)
  assert.doesNotMatch(MAIN_JS, /return findOrReveal\('office'\)/)
})

test('shared mailbox feeds bypass the primary Microsoft Graph mail feed', () => {
  assert.match(MAIN_JS, /mailboxManaged/)
  assert.match(
    MAIN_JS,
    /connections\.feedIsLive\(feed\.kind\) && !\(feed\.kind === 'mail' && service && service\.mailboxManaged\)/
  )
})

test('mail notification baselines are tracked per service key', () => {
  assert.match(MAIN_JS, /const mailNotificationState = new Map\(\)/)
  assert.match(MAIN_JS, /function mailNotifyState\(serviceKey\)/)
  assert.match(MAIN_JS, /diffAndNotifyMail\(items, unreadCount, serviceKey = 'mail'\)/)
  assert.match(MAIN_JS, /showService\(serviceKey\)/)
})

test('mail notifications require a newly identified message, not just count churn', () => {
  assert.match(MAIN_JS, /if \(newItems\.length === 0\) \{[\s\S]*?return[\s\S]*?\}/)
  assert.ok(!MAIN_JS.includes('New email in ${serviceLabel}.'))
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

/* ---------- Audit-fix regressions ---------- */

test('shared mailbox views are kept resident (never hibernated)', () => {
  // needsLiveView must treat mailboxManaged mail feeds as live, otherwise the
  // reaper hibernates them and their scrape-only feed/notifications freeze.
  const block = MAIN_JS.match(/function needsLiveView\(key\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'needsLiveView missing')
  assert.match(block[1], /feed\.kind === 'mail' && service && service\.mailboxManaged/)
})

test('tray/dock badge sums unread across every mail feed', () => {
  const block = MAIN_JS.match(/function mailUnread\(\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'mailUnread missing')
  // Must accumulate rather than return the first mail feed's count.
  assert.match(block[1], /total \+= count/)
  assert.doesNotMatch(block[1], /return serviceState\[key\] \? serviceState\[key\]\.unreadCount : 0/)
})

test('mail unread count reports null (not 0) when it cannot be determined', () => {
  const block = CONNECTIONS_JS.match(/async function getMailUnreadCount\(\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'getMailUnreadCount missing')
  assert.match(block[1], /typeof count === 'number' \? count : null/)
  // The old code coerced a missing token to 0, flickering the badge.
  assert.doesNotMatch(block[1], /: 0\b/)
})

test('api refresh keeps the cached badge when unread count is null', () => {
  assert.match(MAIN_JS, /typeof unread === 'number' && serviceState\[key\]/)
  assert.match(MAIN_JS, /const effectiveUnread = typeof unread === 'number'/)
})

test('a bare 400 from the token endpoint is not treated as a dead grant', () => {
  const block = CONNECTIONS_JS.match(/function isGrantDead\(err\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'isGrantDead missing')
  assert.match(block[1], /DEAD_GRANT_OAUTH_ERRORS\.has\(err\.oauthError\)/)
  assert.match(block[1], /return err\.status === 401/)
  // The blanket 400 catch-all must be gone.
  assert.doesNotMatch(block[1], /err\.status === 400/)
})

test('generic API failures get exponential backoff that resets on success', () => {
  assert.match(CONNECTIONS_JS, /function noteApiFailure\(provider\)/)
  assert.match(CONNECTIONS_JS, /function noteApiSuccess\(provider\)/)
  assert.match(CONNECTIONS_JS, /GENERIC_BACKOFF_BASE_MS \* 2 \*\* \(streak - 1\)/)
  // withToken must record success and failure around the API call.
  assert.match(CONNECTIONS_JS, /noteApiSuccess\(provider\)/)
  assert.match(CONNECTIONS_JS, /noteApiFailure\(provider\)/)
})

test('downloads drawer clears rows and closes when the list empties', () => {
  const block = PANEL_JS.match(/function renderDownloads\(\{ list, activeCount \}\) \{([\s\S]*?)\n  dlToggleBtn\.hidden = false/)
  assert.ok(block, 'renderDownloads empty branch missing')
  assert.match(block[1], /if \(dlList\) dlList\.innerHTML = ''/)
  assert.match(block[1], /if \(dlDrawer\) dlDrawer\.hidden = true/)
})

test('scratchpad rejects stale snapshots that predate an in-flight save', () => {
  assert.match(PANEL_JS, /let lastSavedScratch = null/)
  assert.match(PANEL_JS, /lastSavedScratch === null \|\| snapshot\.scratch === lastSavedScratch/)
  // Both save paths record what was persisted.
  assert.ok((PANEL_JS.match(/lastSavedScratch = scratchArea\.value/g) || []).length >= 2)
})

test('notification toggles are disabled while their provider is disconnected', () => {
  assert.match(PANEL_JS, /toggle\.disabled = !connected/)
  assert.match(PANEL_JS, /if \(btn\.disabled\) return/)
})

test('discovered mailbox URLs are pinned to the primary Outlook origin', () => {
  const block = MAIN_JS.match(/function syncDiscoveredMailboxes\(mailboxes\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'syncDiscoveredMailboxes missing')
  assert.match(block[1], /const safeMailboxUrl = \(value\) =>/)
  assert.match(block[1], /parsed\.protocol !== 'https:'/)
  assert.match(block[1], /parsed\.hostname\.toLowerCase\(\) !== mailHost/)
  // A mailbox whose URL fails validation is dropped, not persisted.
  assert.match(block[1], /const url = safeMailboxUrl\(mb\.url\)\n\s*if \(!url\) continue/)
})
