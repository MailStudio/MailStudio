'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')

const ROOT = path.resolve(__dirname, '..', '..')
const lf = (s) => s.replace(/\r\n/g, '\n')
const MAIN_JS = lf(fs.readFileSync(path.join(ROOT, 'src', 'main', 'main.js'), 'utf8'))
const PANEL_JS = lf(fs.readFileSync(path.join(ROOT, 'src', 'renderer', 'panel.js'), 'utf8'))
const PANEL_CSS = lf(fs.readFileSync(path.join(ROOT, 'src', 'renderer', 'panel.css'), 'utf8'))
const PANEL_HTML = lf(fs.readFileSync(path.join(ROOT, 'src', 'renderer', 'panel.html'), 'utf8'))
const API_FEEDS_JS = lf(fs.readFileSync(path.join(ROOT, 'src', 'main', 'api-feeds.js'), 'utf8'))
const SETTINGS_STORE_JS = lf(fs.readFileSync(path.join(ROOT, 'src', 'main', 'settings-store.js'), 'utf8'))
const CONNECTIONS_JS = lf(fs.readFileSync(path.join(ROOT, 'src', 'main', 'connections.js'), 'utf8'))
const OAUTH_JS = lf(fs.readFileSync(path.join(ROOT, 'src', 'main', 'oauth.js'), 'utf8'))
const MENU_JS = lf(fs.readFileSync(path.join(ROOT, 'src', 'renderer', 'menu.js'), 'utf8'))
const UPDATER_JS = lf(fs.readFileSync(path.join(ROOT, 'src', 'main', 'updater.js'), 'utf8'))
const LIVE_CDP_CHECK_JS = lf(fs.readFileSync(path.join(ROOT, 'tooling', 'qa', 'live-cdp-check.js'), 'utf8'))
const BUILD_SH = lf(fs.readFileSync(path.join(ROOT, 'build.sh'), 'utf8'))
const DEV_DOCS = lf(fs.readFileSync(path.join(ROOT, 'docs', 'development.md'), 'utf8'))
const SECURITY_DOCS = lf(fs.readFileSync(path.join(ROOT, 'docs', 'security.md'), 'utf8'))
const RELEASE_WORKFLOW = lf(fs.readFileSync(path.join(ROOT, '.github', 'workflows', 'release.yml'), 'utf8'))
const PACKAGE_JSON = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'))

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
    uiDensity: 'compact',
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
    workspaces: [{ id: 'w1', name: 'Focus', icon: 'FX', color: '#16A34A', splitKeys: ['mail', 'calendar'] }],
    recentItems: [{ id: 'r1', title: 'Recent', url: 'https://example.com' }],
    downloadHistory: [{ id: 'd1', title: 'Archive', url: 'https://example.com/file.zip' }]
  })

  assert.equal(settings.notif.quietWeekends, true)
  assert.equal(settings.uiDensity, 'compact')
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
  assert.equal(settings.workspaces[0].icon, 'FX')
  assert.equal(settings.workspaces[0].color, '#16a34a')
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

test('mail sidebar sources fetch only the latest 10 messages', () => {
  assert.match(API_FEEDS_JS, /\$top=10&\$orderby=receivedDateTime desc/)
  const block = MAIN_JS.match(/const MAIL_SCRAPE = `\(\(\) => \{([\s\S]*?)\n\}\)\(\)`/)
  assert.ok(block, 'MAIL_SCRAPE missing')
  assert.match(block[1], /if \(items\.length >= 10\) break/)
  assert.match(block[1], /row\.getAttribute\('aria-selected'\) === 'true'/)
  assert.doesNotMatch(block[1], /row\.hasAttribute\('aria-selected'\)/)
  assert.match(block[1], /const selectedInboxUnread = \(\) => \{/)
  assert.match(block[1], /return \{ state: items\.length \? 'ok' : 'empty', items, unreadCount: folderUnread \}/)
  assert.match(block[1], /id: 'inbox-unread-count\\\\x00' \+ folderUnread/)
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

test('icon-only renderer controls expose accessible labels and switch state', () => {
  assert.match(PANEL_JS, /themeToggle\.setAttribute\('aria-label', themeToggle\.title\)/)
  assert.match(PANEL_JS, /themeToggle\.setAttribute\('aria-pressed', theme === 'light' \? 'true' : 'false'\)/)
  assert.match(PANEL_JS, /snoozeBtn\.setAttribute\('aria-label', snoozeBtn\.title\)/)
  assert.match(PANEL_JS, /snoozeBtn\.setAttribute\('aria-pressed', active \? 'true' : 'false'\)/)
  assert.match(PANEL_JS, /homeBtn\.setAttribute\('aria-label', railMode \? `Switch to \$\{service\.label\}` : `Open \$\{service\.label\} home`\)/)
  assert.match(PANEL_JS, /collapseBtn\.setAttribute\('aria-label', feedCollapsed \? `Show \$\{service\.label\} notifications` : `Hide \$\{service\.label\} notifications`\)/)
  assert.match(PANEL_JS, /sw\.setAttribute\('role', 'switch'\)/)
  assert.match(PANEL_JS, /sw\.setAttribute\('aria-checked', service\.visible \? 'true' : 'false'\)/)
  assert.match(PANEL_JS, /sw\.setAttribute\('aria-label', `\$\{service\.visible \? 'Hide' : 'Show'\} \$\{service\.label\} in sidebar`\)/)
  assert.match(PANEL_JS, /rename\.setAttribute\('aria-label', `Rename \$\{service\.label\}`\)/)
  assert.match(PANEL_JS, /editUrl\.setAttribute\('aria-label', `Edit \$\{service\.label\} URL`\)/)
  assert.match(PANEL_JS, /del\.setAttribute\('aria-label', `Remove \$\{service\.label\}`\)/)
  assert.match(PANEL_JS, /toolsTab\.setAttribute\('aria-expanded', toolsDrawer\.classList\.contains\('open'\) \? 'true' : 'false'\)/)
  assert.match(PANEL_JS, /toolsTab\.setAttribute\('aria-controls', 'tools-drawer'\)/)
})

test('collapsed rail service icons switch tabs instead of resetting home', () => {
  assert.match(PANEL_JS, /const railMode = Boolean\(snapshot\.sidebarCollapsed && snapshot\.collapseMode === 'rail' && !snapshot\.settingsOpen\)/)
  assert.match(PANEL_JS, /railMode \? `Switch to \$\{service\.label\}` : `Open \$\{service\.label\} home`/)
  assert.match(PANEL_JS, /const type = railMode && \(event\.metaKey \|\| event\.ctrlKey\) \? 'split-select' : railMode \? 'switch-service' : 'go-service-home'/)
  assert.match(PANEL_CSS, /body\.collapsed\.mode-rail:not\(\.settings-open\) \.service-home\.active/)
  assert.match(PANEL_CSS, /body\.collapsed\.mode-rail:not\(\.settings-open\) \.service-home\.in-split/)
})

test('resize handles are keyboard-adjustable separators', () => {
  assert.match(PANEL_HTML, /id="sidebar-resize-handle"[\s\S]*role="separator"[\s\S]*tabindex="0"[\s\S]*aria-orientation="vertical"/)
  assert.match(PANEL_HTML, /id="split-divider" role="separator" tabindex="0" aria-label="Resize split panes"/)
  assert.match(PANEL_JS, /function paintSidebarWidth\(width\)/)
  assert.match(PANEL_JS, /sbResizeHandle\.setAttribute\('aria-valuenow', String\(newWidth\)\)/)
  assert.match(PANEL_JS, /sbResizeHandle\.addEventListener\('keydown'/)
  assert.match(PANEL_JS, /e\.key === 'ArrowLeft'/)
  assert.match(PANEL_JS, /e\.key === 'ArrowRight'/)
  assert.match(PANEL_JS, /type: 'set-sidebar-width', width, save: true/)
  assert.match(PANEL_JS, /splitDivider\.setAttribute\('aria-valuenow', String\(Math\.round\(ratio \* 100\)\)\)/)
  assert.match(PANEL_JS, /splitDivider\.addEventListener\('keydown'/)
  assert.match(PANEL_JS, /type: 'split-drag-end', ratio/)
  assert.match(PANEL_CSS, /\.sidebar-resize-handle:focus-visible/)
  assert.match(PANEL_CSS, /\.split-divider:focus-visible/)
})

test('main feed click handler resolves clicked items by stable id before row fallback', () => {
  assert.match(MAIN_JS, /function findFeedItem\(feed, command\)/)
  assert.match(MAIN_JS, /const recentItem = findFeedItem\(feed, command\)/)
  assert.match(MAIN_JS, /stringOrNull\(command\.deepLink\)/)
  assert.match(MAIN_JS, /stringOrNull\(command\.taskUrl\)/)
  assert.match(MAIN_JS, /stringOrNull\(recentItem && recentItem\.taskUrl\)/)
})

test('feed item clicks wake hibernated API-backed views before opening links', () => {
  const block = MAIN_JS.match(/case 'open-feed-item': \{([\s\S]*?)\n\s*\}\n\s*case 'open-external':/)
  assert.ok(block, 'open-feed-item handler missing')
  assert.match(block[1], /if \(typeof command\.serviceKey !== 'string'\) break/)
  assert.match(block[1], /const targetView = ensureServiceView\(command\.serviceKey\)/)
  assert.doesNotMatch(block[1], /const targetView = serviceViews\.get\(command\.serviceKey\)/)
})

test('mail feed item opens prefer Graph webLink over constructed deep links', () => {
  const block = MAIN_JS.match(/const itemUrl = feedKind === 'mail'([\s\S]*?)if \(recentItem\)/)
  assert.ok(block, 'feed item URL selection missing')
  assert.match(block[1], /stringOrNull\(command\.webLink\)[\s\S]*stringOrNull\(command\.deepLink\)/)
  assert.match(block[1], /stringOrNull\(recentItem && recentItem\.webLink\)[\s\S]*stringOrNull\(recentItem && recentItem\.deepLink\)/)
})

test('scraped Outlook mail rows are opened, not only selected', () => {
  const mailClickBlock = MAIN_JS.match(/if \(feedKind === 'mail' && typeof command\.rowIdx === 'number'\) \{([\s\S]*?)\n\s*\} else if \(feedKind === 'calendar'/)
  assert.ok(mailClickBlock, 'mail feed click fallback missing')
  assert.match(mailClickBlock[1], /MouseEvent\('dblclick'/)
  assert.match(mailClickBlock[1], /KeyboardEvent\('keydown', \{ key: 'Enter'/)
  assert.match(mailClickBlock[1], /\[data-convid\]/)
})

test('Asana scraper requires task signals and filters metadata chips', () => {
  const block = MAIN_JS.match(/const ASANA_SCRAPE = `\(\(\) => \{([\s\S]*?)\n\}\)\(\)`/)
  assert.ok(block, 'ASANA_SCRAPE missing')
  assert.match(block[1], /const looksLikeMetadata = \(text\) =>/)
  assert.match(block[1], /const cleanTaskName = \(text\) =>/)
  assert.match(block[1], /\.replace\(\/\^\(modal\|dialog\)\\\\s\+\/i, ''\)/)
  assert.match(block[1], /taskSignal\(row\)/)
  assert.match(block[1], /if \(!row \|\| seenRows\.has\(row\) \|\| !taskSignal\(row\)\) continue/)
  assert.match(block[1], /\^due date/i)
  assert.doesNotMatch(block[1], /if \(!rows\.length\) rows = Array\.from\(document\.querySelectorAll\('\[role="row"\]'\)\)/)
})

test('Asana feed names are normalized before caching or rendering', () => {
  assert.match(MAIN_JS, /function cleanAsanaTaskName\(value\)/)
  assert.match(MAIN_JS, /\.replace\(\/\^\(modal\|dialog\)\\s\+\/i, ''\)/)
  assert.match(MAIN_JS, /function normalizeFeedResult\(feed, result\)/)
  assert.match(MAIN_JS, /if \(feed\.kind !== 'asana'\) return result/)
  assert.match(MAIN_JS, /name: cleanAsanaTaskName\(item\.name\)/)
  const applyBlock = MAIN_JS.match(/function applyFeedResult\(key, feed, result, \{ trusted = false \} = \{\}\) \{([\s\S]*?)\n\}/)
  assert.ok(applyBlock, 'applyFeedResult missing')
  assert.match(applyBlock[1], /result = normalizeFeedResult\(feed, result\)/)
})

test('OAuth redirect cancellation does not race loadURL rejection', () => {
  assert.match(OAUTH_JS, /let sawRedirect = false/)
  assert.match(OAUTH_JS, /sawRedirect = true[\s\S]*event\.preventDefault\(\)/)
  assert.match(OAUTH_JS, /authWindow\.loadURL\(authUrl\.toString\(\)\)\.catch\(\(e\) => \{[\s\S]*if \(sawRedirect\) return[\s\S]*finish\(reject, e\)/)
})

test('Microsoft suite auth keeps legacy and new login hosts inside the shared session', () => {
  assert.match(MAIN_JS, /'microsoft365\.com'/)
  assert.match(MAIN_JS, /host === 'login\.windows\.net'/)
  assert.match(MAIN_JS, /host\.endsWith\('\.login\.windows\.net'\)/)
  assert.match(MAIN_JS, /popupHost === 'login\.windows\.net'/)
})

test('Microsoft OAuth prefers silent and hinted SSO before account picker', () => {
  assert.match(OAUTH_JS, /class AuthorizationError extends Error/)
  assert.match(OAUTH_JS, /params\.prompt = authPrompt/)
  assert.match(OAUTH_JS, /params\.login_hint = loginHint\.trim\(\)/)
  assert.match(OAUTH_JS, /const quiet = provider === 'microsoft' && authPrompt === 'none'/)
  assert.match(OAUTH_JS, /silent_timeout/)
  assert.doesNotMatch(OAUTH_JS, /params\.prompt = 'select_account'/)
  assert.match(CONNECTIONS_JS, /prompt: 'none'/)
  assert.match(CONNECTIONS_JS, /prompt: loginHint \? '' : 'select_account'/)
  assert.match(CONNECTIONS_JS, /SILENT_FALLBACK_AUTH_ERRORS/)
  assert.match(MAIN_JS, /loginHint: command\.provider === 'microsoft' \? microsoftLoginHint\(\) : ''/)
  assert.match(MAIN_JS, /onAuthEvent: recordMicrosoftAuthEvent/)
  assert.match(MAIN_JS, /function withMicrosoftLoginHint\(service, url\)/)
  assert.match(MAIN_JS, /parsed\.searchParams\.set\('login_hint', hint\)/)
  assert.match(MAIN_JS, /function serviceHomeTarget\(service\)/)
  assert.match(MAIN_JS, /function serviceInitialTarget\(service\)/)
  assert.match(MAIN_JS, /safeLoadURL\(webContents, serviceHomeTarget\(service\)\)/)
  assert.match(MAIN_JS, /safeLoadURL\(contents, serviceInitialTarget\(service\)\)/)
})

test('setup health treats a Microsoft web session as signed in even without Graph API', () => {
  assert.match(MAIN_JS, /function hasMicrosoftWebSession\(\)/)
  assert.match(MAIN_JS, /microsoftAuthState\.get\(service\.key\) === 'app'/)
  assert.match(MAIN_JS, /const microsoftApiConnected = conn\.microsoft && conn\.microsoft\.status === 'connected'/)
  assert.match(MAIN_JS, /const microsoftWebSession = hasMicrosoftWebSession\(\)/)
  assert.match(MAIN_JS, /microsoftApiConnected \|\| microsoftWebSession/)
  assert.match(MAIN_JS, /Signed in through Microsoft web session/)
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
  assert.match(MAIN_JS, /diffAndNotifyMail\(items, unreadCount, serviceKey = 'mail', options = \{\}\)/)
  assert.match(MAIN_JS, /showService\(serviceKey\)/)
})

test('notification engine centralizes dispatch, audit, cooldowns, and persisted baselines', () => {
  assert.match(SETTINGS_STORE_JS, /notificationState: \{[\s\S]*mail: \{\}[\s\S]*history: \[\]/)
  assert.match(SETTINGS_STORE_JS, /settings\.notificationState = \{[\s\S]*history: cleanRecents\(rawNotificationState\.history, 120\)/)
  assert.match(MAIN_JS, /const NOTIFICATION_KIND_COOLDOWN_MS = \{[\s\S]*teams: 12000/)
  assert.match(MAIN_JS, /const notificationCooldownUntil = new Map\(\)/)
  assert.match(MAIN_JS, /function dispatchNotification\(candidate, onClick\)/)
  assert.match(MAIN_JS, /function recordNotificationEvent\(status, candidate, reason = ''\)/)
  assert.match(MAIN_JS, /function serializeNotificationState\(\)/)
  assert.match(MAIN_JS, /function restoreNotificationEngineState\(\)/)
  assert.match(MAIN_JS, /restoreNotificationEngineState\(\)/)
  assert.match(MAIN_JS, /function getNotificationDiagnostics\(\)/)
  assert.match(MAIN_JS, /notifications: settings\.debugging && settings\.debugging\.enabled \? getNotificationDiagnostics\(\)/)
  assert.match(PANEL_HTML, /id="debug-notification-list"/)
  assert.match(PANEL_JS, /const notificationDiag = diag\.notifications \|\| \{\}/)
  const teamsBlock = MAIN_JS.match(/function maybeNotifyTeams\(unreadCount\) \{([\s\S]*?)\n\}/)
  assert.ok(teamsBlock, 'maybeNotifyTeams missing')
  assert.match(teamsBlock[1], /activeServiceKey === 'teams'[\s\S]*lastTeamsNotificationCount = 0[\s\S]*persistNotificationState\(\)/)
})

test('mail notifications require a newly identified message, not just count churn', () => {
  assert.match(MAIN_JS, /if \(newItems\.length === 0\) \{[\s\S]*?return[\s\S]*?\}/)
  assert.ok(!MAIN_JS.includes('New email in ${serviceLabel}.'))
})

test('mail notification baselines treat inbox zero and suppressed mail as seen', () => {
  const block = MAIN_JS.match(/function diffAndNotifyMail\(items, unreadCount, serviceKey = 'mail', options = \{\}\) \{([\s\S]*?)\nfunction maybeNotifyTeams/)
  assert.ok(block, 'diffAndNotifyMail missing')
  assert.match(block[1], /const knownEmptyBaseline = Boolean\(options\.knownEmptyBaseline\)/)
  assert.match(block[1], /if \(!state\.ready && !hasItems && !knownEmptyBaseline\) return/)
  assert.match(block[1], /if \(unreadCount === 0\) \{[\s\S]*state\.ready = true/)
  assert.match(MAIN_JS, /function markMailItemsSeen\(state, items\)/)
  assert.match(block[1], /const suppress = \(reason\) => \{[\s\S]*markMailItemsSeen\(state, items\)/)
  assert.match(block[1], /persistNotificationState\(\)/)
  assert.match(block[1], /dispatchNotification\(/)
  assert.match(MAIN_JS, /candidate\.suppressWhenFocused !== false && panelWindow && panelWindow\.isFocused\(\) && activeServiceKey === candidate\.serviceKey/)
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

test('settings export uses owner-only atomic writes', () => {
  assert.match(MAIN_JS, /function writeOwnerOnlyFile\(target, payload\)/)
  assert.match(MAIN_JS, /fs\.writeFileSync\(tmp, payload, \{ encoding: 'utf8', mode: 0o600 \}\)/)
  assert.match(MAIN_JS, /fs\.renameSync\(tmp, target\)/)
  assert.match(MAIN_JS, /fs\.chmodSync\(target, 0o600\)/)
  assert.match(MAIN_JS, /writeOwnerOnlyFile\(filePath, JSON\.stringify\(portable, null, 2\)\)/)
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

test('command search and shortcut guide popups are centered in the viewport', () => {
  const overlayRule = PANEL_CSS.match(/\.overlay-panel \{([\s\S]*?)\n\}/)
  assert.ok(overlayRule, 'overlay panel rule missing')
  assert.match(overlayRule[1], /place-items: center;/)
  assert.match(overlayRule[1], /padding: 24px;/)
  assert.match(overlayRule[1], /box-sizing: border-box;/)

  const commandRule = PANEL_CSS.match(/\.command-card \{([\s\S]*?)\n\}/)
  assert.ok(commandRule, 'command card rule missing')
  assert.match(commandRule[1], /width: min\(560px, 100%\);/)

  const shortcutRule = PANEL_CSS.match(/\.shortcut-card \{([\s\S]*?)\n\}/)
  assert.ok(shortcutRule, 'shortcut card rule missing')
  assert.match(shortcutRule[1], /width: min\(680px, 100%\);/)
})

test('transient overlay commands are emitted only when overlay visibility changes', () => {
  const commandPaletteBlock = PANEL_JS.match(/function openCommandPalette\(\) \{([\s\S]*?)function openShortcutGuide/)
  assert.ok(commandPaletteBlock, 'command palette overlay helpers missing')
  assert.match(commandPaletteBlock[1], /const wasHidden = commandPalette\.hidden/)
  assert.match(commandPaletteBlock[1], /if \(wasHidden\) window\.panelApi\.sendCommand\(\{ type: 'open-transient-overlay' \}\)/)
  assert.match(commandPaletteBlock[1], /const wasOpen = commandPalette && !commandPalette\.hidden/)
  assert.match(commandPaletteBlock[1], /if \(wasOpen && !hasOtherTransientOverlay\('command'\)\)/)

  const shortcutBlock = PANEL_JS.match(/function openShortcutGuide\(\) \{([\s\S]*?)const commandPaletteBtn/)
  assert.ok(shortcutBlock, 'shortcut overlay helpers missing')
  assert.match(shortcutBlock[1], /const wasHidden = shortcutGuide\.hidden/)
  assert.match(shortcutBlock[1], /if \(wasHidden\) window\.panelApi\.sendCommand\(\{ type: 'open-transient-overlay' \}\)/)
  assert.match(shortcutBlock[1], /const wasOpen = shortcutGuide && !shortcutGuide\.hidden/)
  assert.match(shortcutBlock[1], /if \(wasOpen && !hasOtherTransientOverlay\('shortcut'\)\)/)
})

test('downloads drawer detaches BrowserViews while visible', () => {
  assert.match(PANEL_JS, /function setDownloadsDrawerOpen\(open\)/)
  const block = PANEL_JS.match(/function setDownloadsDrawerOpen\(open\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'setDownloadsDrawerOpen missing')
  assert.match(block[1], /const wasOpen = !dlDrawer\.hidden/)
  assert.match(block[1], /if \(open && !wasOpen\) \{[\s\S]*open-transient-overlay/)
  assert.match(block[1], /else if \(!open && wasOpen && !hasOtherTransientOverlay\('downloads'\)\) \{[\s\S]*close-transient-overlay/)
  assert.match(PANEL_JS, /setDownloadsDrawerOpen\(dlDrawer\.hidden\)/)
  assert.match(PANEL_JS, /setDownloadsDrawerOpen\(true\)/)
  assert.match(PANEL_JS, /setDownloadsDrawerOpen\(false\)/)
  assert.doesNotMatch(PANEL_JS, /dlDrawer\.hidden = !dlDrawer\.hidden/)
})

test('Teams presence controls use Graph preferred presence from a right-side settings popover', () => {
  assert.match(OAUTH_JS, /Presence\.ReadWrite/)
  assert.match(API_FEEDS_JS, /presence\/setUserPreferredPresence/)
  assert.match(API_FEEDS_JS, /presence\/clearUserPreferredPresence/)
  assert.match(CONNECTIONS_JS, /function setTeamsPreferredPresence\(presence\)/)
  assert.match(CONNECTIONS_JS, /function clearTeamsPreferredPresence\(\)/)
  assert.match(MAIN_JS, /const TEAMS_PRESENCE_OPTIONS = \{[\s\S]*available:[\s\S]*busy:[\s\S]*dnd:[\s\S]*brb:[\s\S]*away:[\s\S]*offline:/)
  assert.match(MAIN_JS, /teamsPresence: teamsPresenceState/)
  assert.match(MAIN_JS, /case 'set-teams-presence':/)
  assert.match(MAIN_JS, /case 'reset-teams-presence':/)
  assert.match(PANEL_HTML, /id="teams-presence-open"/)
  assert.match(PANEL_HTML, /id="teams-presence-popover"/)
  assert.match(PANEL_HTML, /data-teams-presence="available"/)
  assert.match(PANEL_HTML, /data-teams-presence="offline"/)
  assert.match(PANEL_HTML, /data-teams-duration="4h"/)
  assert.match(PANEL_CSS, /\.teams-presence-popover \{[\s\S]*left: calc\(var\(--sb-expanded\) \+ 10px\)/)
  assert.match(PANEL_JS, /function setTeamsPresencePopoverOpen\(open\)/)
  assert.match(PANEL_JS, /type: 'set-teams-presence'/)
  assert.match(PANEL_JS, /type: 'reset-teams-presence'/)
  assert.match(PANEL_JS, /if \(open && !wasOpen\) \{[\s\S]*open-transient-overlay/)
  assert.match(PANEL_JS, /hasOtherTransientOverlay\('teams'\)/)
})

test('focus completion respects reduced-motion preference', () => {
  const block = PANEL_JS.match(/function spawnBurst\(\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'spawnBurst missing')
  assert.match(block[1], /matchMedia\('\(prefers-reduced-motion: reduce\)'\)\.matches\) return/)
})

test('local panel and menu windows only allow their packaged file URLs', () => {
  assert.match(MAIN_JS, /const \{ pathToFileURL \} = require\('url'\)/)
  const block = MAIN_JS.match(/function hardenLocalWindow\(win, allowedFile\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'hardenLocalWindow missing')
  assert.match(block[1], /const allowedFileUrl = pathToFileURL\(allowedFile\)\.href/)
  assert.match(block[1], /if \(url === allowedFileUrl\) return/)
  assert.match(block[1], /event\.preventDefault\(\)/)
  assert.match(MAIN_JS, /hardenLocalWindow\(panelWindow, panelHtml\)[\s\S]*panelWindow\.loadFile\(panelHtml\)/)
  assert.match(MAIN_JS, /hardenLocalWindow\(menuWindow, menuHtml\)[\s\S]*menuWindow\.loadFile\(menuHtml\)/)
})

test('custom pinned hosts are allowed only for their owning service', () => {
  assert.match(MAIN_JS, /let allowedCustomHostsByService = new Map\(\)/)
  const buildBlock = MAIN_JS.match(/function buildAllowedHosts\(\) \{([\s\S]*?)\n\}/)
  assert.ok(buildBlock, 'buildAllowedHosts missing')
  assert.match(buildBlock[1], /if \(service\.builtin \|\| isMailboxService\(service\)\) continue/)
  assert.match(buildBlock[1], /hostsByService\.set\(service\.key, hosts\)/)
  const allowBlock = MAIN_JS.match(/function isAllowedHost\(urlString, serviceKey = null\) \{([\s\S]*?)\n\}/)
  assert.ok(allowBlock, 'isAllowedHost missing')
  assert.match(allowBlock[1], /const serviceHosts = serviceKey \? allowedCustomHostsByService\.get\(serviceKey\) : null/)
  assert.match(allowBlock[1], /serviceHosts && serviceHosts\.has\(host\)/)
  assert.match(MAIN_JS, /isAllowedHost\(url, serviceKey\)/)
  assert.match(MAIN_JS, /isAllowedHost\(url, service\.key\)/)
  assert.match(MAIN_JS, /isAllowedHost\(meta\.href, serviceKey\)/)
})

/* ---------- Audit-fix regressions ---------- */

test('shared mailbox views are kept resident (never hibernated)', () => {
  // needsLiveView must treat mailboxManaged mail feeds as live, otherwise the
  // reaper hibernates them and their scrape-only feed/notifications freeze.
  const block = MAIN_JS.match(/function needsLiveView\(key\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'needsLiveView missing')
  assert.match(block[1], /feed\.kind === 'mail' && service && service\.mailboxManaged/)
})

test('shared Outlook mailboxes participate in Microsoft SSO instead of isolated custom sessions', () => {
  const isMicrosoftBlock = MAIN_JS.match(/function isMicrosoftService\(service\) \{([\s\S]*?)\n\}/)
  assert.ok(isMicrosoftBlock, 'isMicrosoftService missing')
  assert.match(isMicrosoftBlock[1], /isMailboxService\(service\)/)

  const partitionBlock = MAIN_JS.match(/function partitionFor\(service\) \{([\s\S]*?)\n\}/)
  assert.ok(partitionBlock, 'partitionFor missing')
  assert.match(partitionBlock[1], /if \(isMailboxService\(service\)\) return MICROSOFT_SESSION_PARTITION/)
  assert.match(partitionBlock[1], /if \(!service\.builtin\) return `persist:mailstudio-site-\$\{safeCustomPartitionKey\(service\.key\)\}`/)
})

test('custom site partition names are sanitized before Electron session use', () => {
  const safeKeyBlock = MAIN_JS.match(/function safeCustomPartitionKey\(key\) \{([\s\S]*?)\n\}/)
  assert.ok(safeKeyBlock, 'safeCustomPartitionKey missing')
  assert.match(safeKeyBlock[1], /\^\[A-Za-z0-9_-\]\{1,80\}\$/)
  assert.match(safeKeyBlock[1], /replace\(\/\[\^A-Za-z0-9_-\]\/g, '-'\)/)
  assert.match(safeKeyBlock[1], /stablePartitionHash\(raw\)/)
  assert.match(MAIN_JS, /persist:mailstudio-site-\$\{safeCustomPartitionKey\(service\.key\)\}/)
})

test('Outlook mailbox URLs prefer the matching managed mailbox before primary Mail', () => {
  const block = MAIN_JS.match(/function coreServiceForUrl\(urlString\) \{([\s\S]*?)\nfunction resolveServiceByUrl/)
  assert.ok(block, 'coreServiceForUrl missing')
  assert.match(MAIN_JS, /function isOutlookMailHost\(host\)/)
  assert.match(MAIN_JS, /value === 'outlook\.cloud\.microsoft'/)
  assert.match(MAIN_JS, /value\.endsWith\('\.outlook\.cloud\.microsoft'\)/)
  assert.match(block[1], /const isOutlookHost = isOutlookMailHost\(host\)/)
  assert.match(block[1], /const mailbox = settings\.services\.find/)
  assert.match(block[1], /isMailboxService\(service\)/)
  assert.match(block[1], /targetSameOriginAndPathPrefix\(parsed, current\)/)
  assert.match(block[1], /if \(mailbox\) return mailbox/)
  assert.match(block[1], /return find\('mail'\)/)
  assert.match(MAIN_JS, /function targetSameOriginAndPathPrefix\(target, current\)/)
})

test('discovered primary Outlook folder URLs are not duplicated as managed mailboxes', () => {
  const block = MAIN_JS.match(/function syncDiscoveredMailboxes\(mailboxes, discoveredPrimaryEmail = ''\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'syncDiscoveredMailboxes missing')
  assert.match(block[1], /const isPrimaryMailboxUrl = \(value\) =>/)
  assert.match(block[1], /const candidates = \[\]/)
  assert.match(block[1], /if \(!discoveredUrl \|\| isPrimaryMailboxUrl\(discoveredUrl\)\) continue/)
  assert.match(block[1], /mb\.primaryAccount \|\| \(mbEmail && primaryEmails\.has\(mbEmail\)\) \|\| looksLikeNamedPrimary/)
  assert.match(block[1], /\['inbox', 'deeplink', 'id', 'sentitems', 'drafts', 'archive', 'deleteditems', 'junkemail'\]/)
})

test('generic Microsoft 365 home redirects cannot be claimed by the Copilot tab fallback', () => {
  assert.match(MAIN_JS, /function isMicrosoft365HomeHost\(host\)/)
  const coreBlock = MAIN_JS.match(/function coreServiceForUrl\(urlString\) \{([\s\S]*?)\nfunction resolveServiceByUrl/)
  assert.ok(coreBlock, 'coreServiceForUrl missing')
  assert.match(coreBlock[1], /if \(isMicrosoft365HomeHost\(host\)\) \{[\s\S]*return null/)
  const resolveBlock = MAIN_JS.match(/function resolveServiceByUrl\(urlString\) \{([\s\S]*?)\n\}/)
  assert.ok(resolveBlock, 'resolveServiceByUrl missing')
  assert.match(resolveBlock[1], /if \(isMicrosoft365HomeHost\(target\.hostname\)\) return null/)
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
  assert.match(block[1], /setDownloadsDrawerOpen\(false\)/)
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

test('notifications are armed automatically before snapshots', () => {
  assert.match(MAIN_JS, /function armNotificationsForConnectedProviders\(\)/)
  assert.match(MAIN_JS, /store\.save\(\{ \.\.\.settings, onboarded: true, notifSetupSkipped: false \}\)/)
  const pushBlock = MAIN_JS.match(/function pushSnapshot\(\) \{([\s\S]*?)\n  const snapshot = getSnapshot\(\)/)
  assert.ok(pushBlock, 'pushSnapshot missing')
  assert.match(pushBlock[1], /armNotificationsForConnectedProviders\(\)/)
  assert.doesNotMatch(MAIN_JS, /if \(!settings\.onboarded\) return 'not onboarded'/)
  assert.match(MAIN_JS, /Notifications are armed automatically\./)
  assert.match(PANEL_JS, /warn\.hidden = true/)
  assert.match(PANEL_JS, /el\.hidden = true/)
})

test('Asana API failures are classified and can fall back to web scrape', () => {
  assert.match(API_FEEDS_JS, /class ApiError extends Error/)
  assert.match(API_FEEDS_JS, /throw new ApiError\([^)]*res\.status/)
  assert.match(CONNECTIONS_JS, /Asana API task access is blocked/)
  assert.match(CONNECTIONS_JS, /trial or workspace permission changed/)
  assert.match(CONNECTIONS_JS, /Asana API timed out/)
  assert.match(MAIN_JS, /function refreshFeedFromScrape\(key, feed, options = \{\}\)/)
  assert.match(MAIN_JS, /feed\.kind === 'asana' && result\.state === 'error'/)
  assert.match(MAIN_JS, /refreshFeedFromScrape\(key, feed, \{ ignoreApiOwnership: true, apiFallback: true \}\)/)
  assert.match(MAIN_JS, /recordServiceFailure\(key, 'asana-api'/)
  assert.match(MAIN_JS, /recordServiceFailure\(key, 'asana-api-fallback'/)
})

test('provider connect buttons disable immediately and main ignores duplicate connects', () => {
  assert.match(PANEL_JS, /action\.disabled = true/)
  assert.match(PANEL_JS, /action\.textContent = 'Connecting\.\.\.'/)
  const connectBlock = MAIN_JS.match(/case 'connect-provider':([\s\S]*?)case 'disconnect-provider':/)
  assert.ok(connectBlock, 'connect-provider handler missing')
  assert.match(connectBlock[1], /const connStatus = connections\.getStatus\(\)\[command\.provider\]/)
  assert.match(connectBlock[1], /connStatus\.status === 'connecting'/)
  assert.match(connectBlock[1], /command\.provider === 'asana'[\s\S]*activeServiceKey = 'asana'/)
})

test('discovered mailbox URLs are pinned to the primary Outlook origin', () => {
  const block = MAIN_JS.match(/function syncDiscoveredMailboxes\(mailboxes, discoveredPrimaryEmail = ''\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'syncDiscoveredMailboxes missing')
  assert.match(block[1], /const mailHosts = new Set\(\)/)
  assert.match(block[1], /addMailHost\(serviceState\.mail && serviceState\.mail\.href\)/)
  assert.match(block[1], /const safeMailboxUrl = \(value\) =>/)
  assert.match(block[1], /parsed\.protocol !== 'https:'/)
  assert.match(block[1], /!mailHosts\.has\(parsed\.hostname\.toLowerCase\(\)\)/)
  // A mailbox whose URL fails validation is dropped, not persisted.
  assert.match(block[1], /const discoveredUrl = safeMailboxUrl\(mb\.url\)/)
  assert.match(block[1], /if \(!discoveredUrl \|\| isPrimaryMailboxUrl\(discoveredUrl\)\) continue/)
})

test('mailbox discovery timer tolerates stale BrowserViews', () => {
  assert.match(MAIN_JS, /function liveWebContents\(view\)/)
  const block = MAIN_JS.match(/function discoverMailboxes\(\) \{([\s\S]*?)\nfunction scheduleMailboxDiscover/)
  assert.ok(block, 'discoverMailboxes missing')
  assert.match(block[1], /const contents = liveWebContents\(view\)/)
  assert.match(block[1], /if \(!contents\) return/)
  assert.match(block[1], /contents[\s\S]*\.executeJavaScript\(MAILBOX_DISCOVER, true\)/)
  assert.doesNotMatch(block[1], /view\.webContents\.isDestroyed/)
  assert.doesNotMatch(block[1], /view\.webContents[\s\S]*executeJavaScript\(MAILBOX_DISCOVER/)
})

test('partial mailbox discovery never deletes existing shared inbox tabs', () => {
  const block = MAIN_JS.match(/function syncDiscoveredMailboxes\(mailboxes, discoveredPrimaryEmail = ''\) \{([\s\S]*?)\nfunction prunePrimaryMailboxDuplicates/)
  assert.ok(block, 'syncDiscoveredMailboxes missing')
  assert.match(block[1], /const services = settings\.services\.slice\(\)/)
  assert.doesNotMatch(block[1], /settings\.services\.filter\(\(service\) => \{[\s\S]*service\.mailboxManaged[\s\S]*return false/)
  assert.doesNotMatch(block[1], /discoveredKeys/)
})

test('managed mailbox tabs keep broad mailbox routing but canonicalize home to inbox', () => {
  assert.match(MAIN_JS, /function mailboxHomeUrl\(origin, mailboxEmail\)/)
  assert.match(MAIN_JS, /function mailboxRootUrl\(origin, mailboxEmail\)/)
  const block = MAIN_JS.match(/function syncDiscoveredMailboxes\(mailboxes, discoveredPrimaryEmail = ''\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'syncDiscoveredMailboxes missing')
  assert.match(block[1], /const origin = \(\(\) => \{/)
  assert.match(block[1], /const url = mailboxRootUrl\(origin, mbEmail\) \|\| discoveredUrl/)
  assert.match(block[1], /const home = mailboxHomeUrl\(origin, mbEmail\) \|\| safeMailboxUrl\(mb\.home\) \|\| discoveredUrl/)
})

test('mail first load prefers inbox home over the broader mailbox root', () => {
  const block = MAIN_JS.match(/function ensureServiceLoaded\(key\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'ensureServiceLoaded missing')
  assert.match(block[1], /serviceInitialTarget\(service\)/)
  const initialBlock = MAIN_JS.match(/function serviceInitialTarget\(service\) \{([\s\S]*?)\n\}/)
  assert.ok(initialBlock, 'serviceInitialTarget missing')
  assert.match(initialBlock[1], /service\.key === 'mail' \|\| service\.mailboxManaged/)
  assert.match(SETTINGS_STORE_JS, /home: 'https:\/\/outlook\.office\.com\/mail\/inbox'/)
})

test('shared mailbox scrape results are accepted even when the primary Graph mail feed is live', () => {
  const block = MAIN_JS.match(/function refreshFeed\(key\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'refreshFeed missing')
  assert.match(MAIN_JS, /function apiOwnsFeed\(key\)/)
  assert.match(
    block[1],
    /if \(apiOwnsFeed\(key\)\)/
  )
  const ownerBlock = MAIN_JS.match(/function apiOwnsFeed\(key\) \{([\s\S]*?)\n\}/)
  assert.ok(ownerBlock, 'apiOwnsFeed missing')
  assert.match(ownerBlock[1], /connections\.feedIsLive\(feed\.kind\)/)
  assert.match(ownerBlock[1], /!\(feed\.kind === 'mail' && service && service\.mailboxManaged\)/)
  const refreshFeedsBlock = MAIN_JS.match(/function refreshFeeds\(\) \{([\s\S]*?)\nfunction refreshApiFeedsNow/)
  assert.ok(refreshFeedsBlock, 'refreshFeeds missing')
  assert.match(refreshFeedsBlock[1], /if \(apiOwnsFeed\(key\)\)/)
  assert.doesNotMatch(refreshFeedsBlock[1], /connections\.feedIsLive\(kind\)/)
})

test('scraped mail unread count uses unread rows, not title count or all recent rows', () => {
  const block = MAIN_JS.match(/const scrapedUnread = feed\.items\.filter\(\(item\) => item && item\.isRead === false\)\.length([\s\S]*?)diffAndNotifyMail\(feed\.items, visibleUnread, key, \{/)
  assert.ok(block, 'scraped mail branch missing')
  assert.match(block[1], /const folderUnread = result && typeof result\.unreadCount === 'number'/)
  assert.match(block[1], /const visibleUnread = folderUnread === null \? scrapedUnread : Math\.max\(folderUnread, scrapedUnread\)/)
  assert.doesNotMatch(block[1], /service && service\.mailboxManaged\s*\?\s*scrapedUnread/)
  assert.match(block[1], /mailState\.unreadCount = visibleUnread/)
  assert.match(MAIN_JS, /knownEmptyBaseline: key === activeServiceKey && panelWindow && panelWindow\.isFocused\(\)/)
  assert.match(block[1], /changed = true/)
})

test('API-backed feeds refresh immediately on startup and connection changes', () => {
  assert.match(MAIN_JS, /function refreshApiFeedsNow\(\) \{[\s\S]*?if \(apiOwnsFeed\(key\)\) refreshFeed\(key\)/)
  const timerBlock = MAIN_JS.match(/function startFeedTimer\(\) \{([\s\S]*?)\n\}/)
  assert.ok(timerBlock, 'startFeedTimer missing')
  assert.match(timerBlock[1], /refreshApiFeedsNow\(\)/)
  const initBlock = MAIN_JS.match(/connections\.init\(\{([\s\S]*?)\n\s*\}\)/)
  assert.ok(initBlock, 'connections init missing')
  assert.match(initBlock[1], /onChange: \(\) => \{[\s\S]*refreshApiFeedsNow\(\)[\s\S]*pushSnapshot\(\)/)
  const connectBlock = MAIN_JS.match(/case 'connect-provider':([\s\S]*?)case 'disconnect-provider':/)
  assert.ok(connectBlock, 'connect-provider handler missing')
  assert.match(connectBlock[1], /resetProviderFeeds\(command\.provider\)[\s\S]*refreshApiFeedsNow\(\)/)
  const applyBlock = MAIN_JS.match(/function applySettings\(next\) \{([\s\S]*?)\n\}/)
  assert.ok(applyBlock, 'applySettings missing')
  assert.match(applyBlock[1], /setTimeout\(refreshApiFeedsNow, 250\)/)
})

test('scrape-owned mail refreshes promptly without requiring a click', () => {
  assert.match(MAIN_JS, /const HIDDEN_MAIL_SCRAPE_MS = 12000/)
  const finishBlock = MAIN_JS.match(/view\.webContents\.on\('did-finish-load', \(\) => \{([\s\S]*?)\n  \}\)/)
  assert.ok(finishBlock, 'did-finish-load handler missing')
  assert.match(finishBlock[1], /refreshFeed\(service\.key\)/)
  assert.match(finishBlock[1], /setTimeout\(\(\) => refreshFeed\(service\.key\), 1000\)/)
  const feedsBlock = MAIN_JS.match(/function refreshFeeds\(\) \{([\s\S]*?)\nfunction refreshApiFeedsNow/)
  assert.ok(feedsBlock, 'refreshFeeds missing')
  assert.match(feedsBlock[1], /const hiddenGap = serviceFeeds\[key\]\.kind === 'mail' \? HIDDEN_MAIL_SCRAPE_MS : HIDDEN_SCRAPE_MS/)
})

test('new scrape-owned inboxes are prewarmed after settings changes', () => {
  const prewarmBlock = MAIN_JS.match(/function prewarmLiveScrapeViews\(\) \{([\s\S]*?)\n\}/)
  assert.ok(prewarmBlock, 'prewarmLiveScrapeViews missing')
  assert.match(prewarmBlock[1], /needsLiveView\(service\.key\)/)
  assert.match(prewarmBlock[1], /enqueuePrewarm\(keys\)/)
  const applyBlock = MAIN_JS.match(/function applySettings\(next\) \{([\s\S]*?)\n\}/)
  assert.ok(applyBlock, 'applySettings missing')
  assert.match(applyBlock[1], /syncServiceViews\(\)[\s\S]*prewarmLiveScrapeViews\(\)/)
})

test('the Copilot tab is preserved as a hidden built-in service', () => {
  const fresh = store.normalize(null)
  const office = fresh.services.find((s) => s.key === 'office')
  assert.ok(office, 'Copilot/office tab missing from defaults')
  assert.equal(office.label, 'Copilot')
  assert.equal(office.visible, false)
})

test('launcher-capable apps may use Microsoft 365 bootstrap while core tabs stay pinned', () => {
  assert.match(MAIN_JS, /function isMicrosoft365HomeUrl\(urlString\)/)
  assert.match(MAIN_JS, /const MICROSOFT_365_LAUNCHER_KEYS = new Set/)
  assert.match(MAIN_JS, /function canUseMicrosoft365Launcher\(service\)/)
  assert.match(MAIN_JS, /const MICROSOFT_365_LAUNCHER_KEYS = new Set\(\[\s*'teams'/)
  // Direct navigations and server redirects are still cancelled for Mail/Calendar
  // style tabs, but Teams and Office apps can pass through the launcher.
  const navBlock = MAIN_JS.match(/const handleNavRequest = \(event, url\) => \{([\s\S]*?)\n  \}/)
  assert.ok(navBlock, 'handleNavRequest missing')
  assert.match(navBlock[1], /if \(!canUseMicrosoft365Launcher\(service\) && isMicrosoft365HomeUrl\(url\)\) \{\s*event\.preventDefault\(\)\s*return/)
  // New-window/pop-up requests from core tabs are denied.
  assert.match(MAIN_JS, /if \(!canUseMicrosoft365Launcher\(service\) && isMicrosoft365HomeUrl\(url\)\) \{\s*return \{ action: 'deny' \}/)
  // Explicit Copilot-home clicks route to the Copilot tab, never the active view.
  const openBlock = MAIN_JS.match(/function openLinkInApp\(url\) \{([\s\S]*?)\n\}/)
  assert.ok(openBlock, 'openLinkInApp missing')
  assert.match(openBlock[1], /const office = findService\('office'\)\s*\n\s*if \(office\) routeToService\(office, url\)/)
})

test('Planner default uses the current cloud Microsoft host', () => {
  const fresh = store.normalize(null)
  const planner = fresh.services.find((s) => s.key === 'planner')
  assert.ok(planner, 'Planner tab missing from defaults')
  assert.equal(planner.url, 'https://planner.cloud.microsoft/')
  assert.equal(planner.home, 'https://planner.cloud.microsoft/')
  assert.doesNotMatch(SETTINGS_STORE_JS, /https:\/\/planner\.microsoft\.com\//)
})

test('Teams default uses the cloud Microsoft host', () => {
  const fresh = store.normalize(null)
  const teams = fresh.services.find((s) => s.key === 'teams')
  assert.ok(teams, 'Teams tab missing from defaults')
  assert.equal(teams.url, 'https://teams.cloud.microsoft/')
  assert.equal(teams.home, 'https://teams.cloud.microsoft/')
})

test('Teams email-share links from mail route directly to the Teams tab', () => {
  assert.match(MAIN_JS, /function isTeamsEmailShareLink\(urlString\)/)
  assert.match(MAIN_JS, /\/\^\\\/l\\\/share-email\(\?:\\\/\|\$\)\/i/)
  assert.match(MAIN_JS, /\/\^\\\/dl\\\/launcher\\\/launcher\\\.html\$\/i/)
  assert.match(MAIN_JS, /launcherType === 'share-email'/)
  assert.match(MAIN_JS, /function shouldConfirmTeamsNavigation\(url, sourceService\)/)
  assert.match(MAIN_JS, /return !isTeamsEmailShareLink\(url\)/)
  const guardedRoutes = [...MAIN_JS.matchAll(/internalService\.key === 'teams' && shouldConfirmTeamsNavigation\(url, service\)/g)]
  assert.equal(guardedRoutes.length, 2)
  assert.match(MAIN_JS, /else \{\s*routeToService\(internalService, url\)\s*\}/)
})

test('blank-page health only reports for onscreen tabs or Teams', () => {
  assert.match(MAIN_JS, /const BLANK_HEALTH_SWEEP_MS = 10000/)
  assert.match(MAIN_JS, /const lastBlankHealthAt = \{\}/)
  assert.match(MAIN_JS, /function shouldReportBlankLoad\(serviceKey\)/)
  assert.match(MAIN_JS, /if \(serviceKey === 'teams'\) return true/)
  assert.match(MAIN_JS, /return Boolean\(serviceState\[serviceKey\] && serviceState\[serviceKey\]\.visible\)/)
  assert.match(MAIN_JS, /function checkServiceBlankHealth\(service, reason = 'timer'\)/)
  assert.match(MAIN_JS, /if \(!service \|\| !shouldReportBlankLoad\(service\.key\)\) return false[\s\S]*state: 'blank'/)
  assert.match(MAIN_JS, /const currentView = serviceViews\.get\(service\.key\)/)
  assert.match(MAIN_JS, /const currentContents = currentView && currentView\.webContents/)
  assert.match(MAIN_JS, /function maybeClickTeamsRecoveryAction\(serviceKey, health\)/)
  assert.match(MAIN_JS, /document\.getElementById\('error-action-clear-cache'\)/)
  assert.match(MAIN_JS, /recordServiceFailure\([\s\S]*'teams-recovery'/)
  assert.match(MAIN_JS, /const preserveTeamsBlank =[\s\S]*service\.key === 'teams'[\s\S]*\['blank', 'failed', 'safe-mode'\]\.includes\(previousHealth\.state\)/)
  assert.match(MAIN_JS, /if \(service\.key !== 'teams'\) clearServiceHealth\(service\.key\)/)
  assert.match(MAIN_JS, /serviceState\[service\.key\]\.health\.state !== 'ok'/)
  assert.match(MAIN_JS, /const visibleContentArea = Array\.from/)
  assert.match(MAIN_JS, /hasVisibleContent: text\.length > 0 \|\| visibleContentArea > 1000/)
  assert.match(MAIN_JS, /const teamsBlankShell =[\s\S]*service\.key === 'teams'[\s\S]*!health\.hasVisibleContent \|\| health\.hasTeamsClearCache \|\| health\.hasTeamsRetry/)
  assert.match(MAIN_JS, /const healthyContent = service\.key === 'teams'[\s\S]*\? health\.hasVisibleContent/)
  assert.match(MAIN_JS, /if \(teamsBlankShell && maybeClickTeamsRecoveryAction\(service\.key, health\)\) return/)
  assert.match(MAIN_JS, /Teams loaded its shell without chat content\./)
  assert.match(MAIN_JS, /Teams could not be inspected during blank-page health check\./)
  assert.match(MAIN_JS, /service\.key !== 'teams' && currentContents\.isLoading\(\)/)
  assert.match(MAIN_JS, /now - \(lastBlankHealthAt\[service\.key\] \|\| 0\) < BLANK_HEALTH_SWEEP_MS/)
  assert.match(MAIN_JS, /checkServiceBlankHealth\(service, 'timer'\)/)
  assert.doesNotMatch(MAIN_JS, /setTimeout\(\(\) => \{\s*if \(view\.webContents\.isDestroyed\(\)\) return/)
  assert.doesNotMatch(LIVE_CDP_CHECK_JS, /!service\.host\.includes\('teams\.'\)/)
})

test('stability debugging exposes tab observability behind an explicit setting', () => {
  assert.match(SETTINGS_STORE_JS, /debugging: \{\s*enabled: false\s*\}/)
  assert.match(SETTINGS_STORE_JS, /settings\.serviceFailureLog = cleanRecents\(raw && raw\.serviceFailureLog, 80\)/)
  assert.match(MAIN_JS, /const STALE_FEED_MS = 7 \* 60 \* 1000/)
  assert.match(MAIN_JS, /serviceFailureLog: settings\.debugging && settings\.debugging\.enabled \?/)
  assert.match(MAIN_JS, /function recordDebugEvent\(kind, title/)
  assert.match(MAIN_JS, /recordServiceFailure\(key, 'feed-stale'/)
  assert.match(MAIN_JS, /recordServiceFailure\(key, 'prewarm-timeout'/)
  assert.match(MAIN_JS, /recordDebugEvent\('network-offline'/)
  assert.match(MAIN_JS, /function buildDebugReport\(\)/)
  assert.match(MAIN_JS, /microsoftAuthEvents: microsoftAuthEvents\.slice/)
  assert.match(MAIN_JS, /microsoftAuth: settings\.debugging && settings\.debugging\.enabled \? getMicrosoftAuthDiagnostics\(\)/)
  assert.match(MAIN_JS, /case 'export-debug-report':/)
  assert.match(MAIN_JS, /prewarmState: prewarmCurrentKey === service\.key/)
  assert.match(MAIN_JS, /visibleState: \(serviceState\[service\.key\] \|\| \{\}\)\.visible \? 'onscreen' : 'background'/)
  assert.match(PANEL_HTML, /id="set-page-debugging"/)
  assert.match(PANEL_HTML, /id="debug-auth-list"/)
  assert.match(PANEL_HTML, /id="export-debug-report"/)
  assert.match(PANEL_JS, /function renderDebugging\(snapshot\)/)
  assert.match(PANEL_JS, /const auth = diag\.microsoftAuth \|\| \{\}/)
  assert.match(PANEL_JS, /repair-partition', partition: 'microsoft'/)
  assert.match(PANEL_JS, /repair-partition', partition: 'asana'/)
  assert.match(PANEL_JS, /clear-failure-log/)
  assert.match(PANEL_JS, /export-debug-report/)
})

test('Microsoft session repair is soft before destructive clear-session repair', () => {
  assert.match(MAIN_JS, /const MICROSOFT_SESSION_HEALTH_MS = 2 \* 60 \* 1000/)
  assert.match(MAIN_JS, /function runMicrosoftSessionHealthSweep\(reason = 'timer'\)/)
  assert.match(MAIN_JS, /function softRepairMicrosoftService\(service, reason = 'manual'/)
  assert.match(MAIN_JS, /serviceHomeTarget\(service\)/)
  assert.match(MAIN_JS, /case 'network-online':[\s\S]*runMicrosoftSessionHealthSweep\('network-online'\)/)
  assert.match(MAIN_JS, /startMicrosoftSessionHealthTimer\(\)/)
  const partitionBlock = MAIN_JS.match(/function repairPartition\(kind\) \{([\s\S]*?)const partition =/)
  assert.ok(partitionBlock, 'repairPartition missing')
  assert.match(partitionBlock[1], /if \(kind === 'microsoft'\)/)
  assert.match(partitionBlock[1], /runMicrosoftSessionHealthSweep\('manual-repair'\)/)
  assert.doesNotMatch(partitionBlock[1], /clearStorageData/)
  assert.match(MAIN_JS, /function forceUnloadServiceView\(key\) \{/)
  const clearSessionBlock = MAIN_JS.match(/if \(action === 'clear-session'\) \{([\s\S]*?)\n  \}/)
  assert.ok(clearSessionBlock, 'clear-session repair action missing')
  assert.match(clearSessionBlock[1], /forceUnloadServiceView\(serviceKey\)/)
})

test('the signed-in mailbox address is captured so it can be de-duplicated from shared mailboxes', () => {
  const block = CONNECTIONS_JS.match(/async function loadAccount\(provider, tokenSet\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'loadAccount missing')
  assert.match(block[1], /userPrincipalName/)
  assert.match(block[1], /email: email \|\| null/)
})

test('the primary account is not duplicated as a shared mailbox tab', () => {
  const block = MAIN_JS.match(/function syncDiscoveredMailboxes\(mailboxes, discoveredPrimaryEmail = ''\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'syncDiscoveredMailboxes missing')
  assert.match(MAIN_JS, /function syncDiscoveredMailboxes\(mailboxes, discoveredPrimaryEmail = ''\)/)
  assert.match(MAIN_JS, /function extractEmail\(value\)/)
  assert.match(MAIN_JS, /function connectedMicrosoftEmail\(\)/)
  assert.match(MAIN_JS, /function connectedMicrosoftNameToken\(\)/)
  assert.match(block[1], /const primaryEmails = new Set\(\)/)
  assert.match(block[1], /connectedMicrosoftEmail\(\), discoveredPrimaryEmail/)
  assert.match(block[1], /extractEmail\(candidateEmail\)/)
  assert.match(block[1], /const primaryNameToken = connectedMicrosoftNameToken\(\)/)
  assert.match(block[1], /const looksLikeNamedPrimary =/)
  assert.match(block[1], /primaryNameToken\.startsWith\(localPart\) \|\| localPart\.startsWith\(primaryNameToken\)/)
  assert.match(block[1], /mb\.primaryAccount \|\| \(mbEmail && primaryEmails\.has\(mbEmail\)\) \|\| looksLikeNamedPrimary/)
  assert.match(MAIN_JS, /syncDiscoveredMailboxes\(result\.mailboxes, result\.primaryEmail\)/)
})

test('persisted primary mailbox duplicates are pruned on startup before views are built', () => {
  assert.match(MAIN_JS, /function prunePrimaryMailboxDuplicates\(\)/)
  const pruneBlock = MAIN_JS.match(/function prunePrimaryMailboxDuplicates\(\) \{([\s\S]*?)\n\}/)
  assert.ok(pruneBlock, 'prunePrimaryMailboxDuplicates missing')
  assert.match(pruneBlock[1], /service\.mailboxManaged/)
  assert.match(pruneBlock[1], /primaryNameToken\.startsWith\(localPart\) \|\| localPart\.startsWith\(primaryNameToken\)/)
  assert.match(pruneBlock[1], /settings = store\.save\(\{ \.\.\.settings, services \}\)/)
  const bootBlock = MAIN_JS.match(/connections\.init\(\{([\s\S]*?)\n\s*for \(const service of settings\.services\)/)
  assert.ok(bootBlock, 'startup connection/view sequence missing')
  assert.match(bootBlock[1], /if \(prunePrimaryMailboxDuplicates\(\)\)/)
})

test('the built-in Mail row displays the signed-in primary mailbox identity', () => {
  assert.match(MAIN_JS, /let discoveredPrimaryMailEmail = ''/)
  assert.match(MAIN_JS, /function rememberDiscoveredPrimaryMailEmail\(value\)/)
  assert.match(MAIN_JS, /function primaryMailDisplayLabel\(\)/)
  assert.match(MAIN_JS, /function serviceDisplayLabel\(service\)/)
  assert.match(MAIN_JS, /if \(service\.key === 'mail'\) return primaryMailDisplayLabel\(\) \|\| service\.label/)

  const syncBlock = MAIN_JS.match(/function syncDiscoveredMailboxes\(mailboxes, discoveredPrimaryEmail = ''\) \{([\s\S]*?)\n\}/)
  assert.ok(syncBlock, 'syncDiscoveredMailboxes missing')
  assert.match(syncBlock[1], /const primaryIdentityChanged = rememberDiscoveredPrimaryMailEmail\(discoveredPrimaryEmail\)/)
  assert.match(syncBlock[1], /else if \(primaryIdentityChanged\) pushSnapshot\(\)/)

  const snapshotBlock = MAIN_JS.match(/function getSnapshot\(\) \{([\s\S]*?)\nfunction armNotificationsForConnectedProviders/)
  assert.ok(snapshotBlock, 'getSnapshot missing')
  assert.match(snapshotBlock[1], /const label = serviceDisplayLabel\(service\)/)
  assert.match(snapshotBlock[1], /label,/)
  assert.match(snapshotBlock[1], /title: state\.title \|\| label/)

  const pushBlock = MAIN_JS.match(/function pushSnapshot\(\) \{([\s\S]*?)\nfunction serviceKeyForWebContents/)
  assert.ok(pushBlock, 'pushSnapshot missing')
  assert.match(pushBlock[1], /const svcLabel = serviceDisplayLabel\(svc\)/)
  assert.match(pushBlock[1], /\$\{svcLabel\}/)
})

test('Outlook mailbox discovery marks the profile account as primary', () => {
  const block = MAIN_JS.match(/const MAILBOX_DISCOVER = `\(\(\) => \{([\s\S]*?)\n\}\)\(\)`/)
  assert.ok(block, 'MAILBOX_DISCOVER missing')
  assert.match(block[1], /let primaryEmail = ''/)
  assert.match(block[1], /O365_MainLink_Me/)
  assert.match(block[1], /outlook\\\\\.cloud\\\\\.microsoft/)
  assert.match(block[1], /\[aria-label\*="Account manager" i\]/)
  assert.match(block[1], /const allTreeItems = Array\.from\(document\.querySelectorAll\('\[role="treeitem"\]'\)\)/)
  assert.match(block[1], /selected && currentRootEmail && .*inbox/)
  assert.match(block[1], /primaryEmail = currentRootEmail/)
  assert.match(block[1], /encodeURIComponent\(fallbackEmail\) \+ '\/inbox'/)
  assert.match(block[1], /primaryAccount: Boolean\(options\.primaryAccount\)/)
  assert.match(block[1], /let firstMailboxRoot = true/)
  assert.match(block[1], /primaryAccount: primaryEmail \? email === primaryEmail : firstMailboxRoot/)
  assert.match(block[1], /firstMailboxRoot = false/)
  assert.match(block[1], /return \{ mailboxes, primaryEmail \}/)
})

test('summary strip accumulates unread mail across every visible mail feed', () => {
  const block = PANEL_JS.match(/function renderSummary\(snapshot\) \{([\s\S]*?)\n\}/)
  assert.ok(block, 'renderSummary missing')
  assert.match(block[1], /mail \+= \(service\.feed\.items \|\| \[\]\)\.filter/)
  assert.match(block[1], /item\.isRead === false/)
  assert.doesNotMatch(block[1], /mail = \(service\.feed\.items \|\| \[\]\)\.filter/)
})

test('renderer settings round-trip preserves managed mailbox metadata', () => {
  assert.match(MAIN_JS, /mailboxManaged: Boolean\(service\.mailboxManaged\)/)
  const workingBlock = PANEL_JS.match(/function workingServices\(\) \{([\s\S]*?)\n\}/)
  assert.ok(workingBlock, 'workingServices missing')
  assert.match(workingBlock[1], /icon: s\.icon/)
  assert.match(workingBlock[1], /mailboxManaged: Boolean\(s\.mailboxManaged\)/)
  assert.match(workingBlock[1], /feed: s\.feed \? s\.feed\.kind : undefined/)
})

test('service preload page metadata is handled by main and sender-scoped to BrowserViews', () => {
  assert.match(MAIN_JS, /ipcMain\.on\('service:page-meta'/)
  assert.match(MAIN_JS, /function serviceKeyForWebContents\(webContents\)/)
  assert.match(MAIN_JS, /const serviceKey = serviceKeyForWebContents\(event\.sender\)/)
  assert.match(MAIN_JS, /meta\.serviceKey !== serviceKey/)
  assert.match(MAIN_JS, /handleMicrosoftNavigation\(service, meta\.href\)/)
})

test('panel commands are shape-checked before dispatch', () => {
  const block = MAIN_JS.match(/ipcMain\.on\('panel:command', \(event, command\) => \{([\s\S]*?)\n\s*switch \(command\.type\)/)
  assert.ok(block, 'panel command handler missing')
  assert.match(block[1], /event\.sender !== panelWindow\?\.webContents && event\.sender !== menuWindow\?\.webContents/)
  assert.match(block[1], /if \(!command \|\| typeof command\.type !== 'string'\) return/)
})

test('partial settings updates merge nested preference groups', () => {
  assert.match(MAIN_JS, /function mergeObjectSetting\(current, patch\)/)
  const block = MAIN_JS.match(/case 'update-settings':([\s\S]*?)\n\s*break/)
  assert.ok(block, 'update-settings handler missing')
  assert.match(block[1], /notif: mergeObjectSetting\(settings\.notif, command\.settings\.notif\)/)
  assert.match(block[1], /feedPrefs: mergeObjectSetting\(settings\.feedPrefs, command\.settings\.feedPrefs\)/)
  assert.match(block[1], /downloads: mergeObjectSetting\(settings\.downloads, command\.settings\.downloads\)/)
  assert.doesNotMatch(block[1], /notif: command\.settings\.notif \|\| settings\.notif/)
})

test('tray dropdown status sums unread across every mail service', () => {
  const statusBlock = MENU_JS.match(/function renderStatus\(mailService\) \{([\s\S]*?)\n\}/)
  assert.ok(statusBlock, 'renderStatus missing')
  assert.match(statusBlock[1], /reduce\(\(total, service\) => total \+ \(Number\(service\.unreadCount\) \|\| 0\), 0\)/)
  assert.match(statusBlock[1], /feedStates\.includes\('error'\)/)
  const renderBlock = MENU_JS.match(/function render\(snapshot\) \{([\s\S]*?)\n\}/)
  assert.ok(renderBlock, 'render missing')
  assert.match(renderBlock[1], /snapshot\.services\.filter\(\(service\) => service\.feed && service\.feed\.kind === 'mail'\)/)
  assert.doesNotMatch(renderBlock[1], /snapshot\.services\.find\(\(service\) => service\.feed && service\.feed\.kind === 'mail'\)/)
})

test('home and external-open commands tolerate missing active service state', () => {
  assert.match(MAIN_JS, /function activeServiceHref\(\)/)
  assert.match(MAIN_JS, /safeLoadURL\(webContents, serviceHomeTarget\(service\)\)/)
  assert.match(MAIN_JS, /openExternalSafe\(activeServiceHref\(\)\)/)
  assert.doesNotMatch(MAIN_JS, /openExternalSafe\(serviceState\[activeServiceKey\]\.href\)/)
})

test('panel BrowserWindow listener cap covers repeated OAuth child windows and heavy navigation', () => {
  const block = MAIN_JS.match(/function createPanelWindow\(\) \{([\s\S]*?)\n  panelWindow\.webContents\.setMaxListeners\(40\)/)
  assert.ok(block, 'createPanelWindow listener setup missing')
  assert.match(block[1], /panelWindow\.setMaxListeners\(80\)/)
})

test('build wrapper does not invoke macOS GUI tooling unless explicitly requested', () => {
  assert.match(BUILD_SH, /MAILSTUDIO_OPEN_DIST/)
  assert.match(BUILD_SH, /\$\{MAILSTUDIO_OPEN_DIST:-\}/)
  assert.doesNotMatch(BUILD_SH, /if \[\[ "\$\(uname\)" == "Darwin" \]\]; then\n\s*open dist\//)
})

test('release and build paths run checks and tests before packaging', () => {
  assert.match(PACKAGE_JSON.scripts.release, /npm run check && npm test && electron-builder --publish always/)
  assert.equal(PACKAGE_JSON.scripts['start:debug'], 'electron --remote-debugging-port=9333 .')
  assert.match(RELEASE_WORKFLOW, /name: Syntax check[\s\S]*run: npm run check/)
  assert.match(RELEASE_WORKFLOW, /name: Unit tests[\s\S]*run: npm test/)
  assert.match(RELEASE_WORKFLOW, /name: Build and publish[\s\S]*npx electron-builder --publish always/)
  assert.match(BUILD_SH, /npm ci --prefer-offline/)
  assert.match(BUILD_SH, /MAILSTUDIO_SKIP_TESTS/)
  assert.match(BUILD_SH, /npm test/)
  assert.ok(BUILD_SH.indexOf('npm run check') < BUILD_SH.indexOf('npm test'))
})

test('development docs match release workflow platform split', () => {
  assert.match(DEV_DOCS, /builds the Windows and Linux\s+artifacts/)
  assert.match(DEV_DOCS, /Build the macOS artifact locally with `npm run dist:mac`/)
  assert.doesNotMatch(DEV_DOCS, /builds the macOS, Windows, and\s+Linux artifacts/)
})

test('manual update-check failures always surface user-visible feedback', () => {
  assert.match(UPDATER_JS, /function showManualCheckFailed\(err\)/)
  assert.match(UPDATER_JS, /autoUpdater\.on\('error'[\s\S]*showManualCheckFailed\(err\)/)
  assert.match(UPDATER_JS, /checkForUpdates\(\)\?\.catch\(\(e\) => \{[\s\S]*if \(manual\) showManualCheckFailed\(e\)/)
  assert.match(UPDATER_JS, /catch \(err\) \{[\s\S]*if \(manual\) showManualCheckFailed\(err\)/)
})

test('Teams meetings can request camera and microphone without broadening device permissions', () => {
  assert.match(MAIN_JS, /function isTeamsHost\(host\)/)
  assert.match(MAIN_JS, /if \(permission === 'media'\) \{[\s\S]*return isTeamsHost\(new URL\(url\)\.hostname\)/)
  assert.doesNotMatch(MAIN_JS, /ALLOWED_PERMISSIONS = new Set\(\[[^\]]*'media'/)
  assert.doesNotMatch(MAIN_JS, /permission === 'display-capture'[\s\S]*return true/)
  const extendInfo = PACKAGE_JSON.build.mac.extendInfo
  assert.match(extendInfo.NSCameraUsageDescription, /Teams meeting or call/)
  assert.match(extendInfo.NSMicrophoneUsageDescription, /Teams meeting or call/)
  assert.match(SECURITY_DOCS, /camera\/microphone media only for\s+Teams meetings\/calls/)
  assert.match(SECURITY_DOCS, /screen capture, USB, etc\. are denied/)
})
