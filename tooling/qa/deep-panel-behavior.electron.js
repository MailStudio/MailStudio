'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const { app, BrowserWindow } = require('electron')

const ROOT = path.resolve(__dirname, '..', '..')
const HARNESS = path.join(__dirname, 'panel-browser-harness.generated.html')
const USER_DATA = path.join(__dirname, '.tmp-deep-panel-behavior')

app.setPath('userData', USER_DATA)
app.disableHardwareAcceleration()

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

function hasCommand(commands, partial) {
  return commands.some((command) =>
    Object.entries(partial).every(([key, value]) => command && command[key] === value)
  )
}

async function main() {
  const issues = []
  const consoleMessages = []
  fs.rmSync(USER_DATA, { recursive: true, force: true })
  fs.mkdirSync(USER_DATA, { recursive: true })

  await app.whenReady()
  const win = new BrowserWindow({
    show: false,
    width: 1180,
    height: 900,
    webPreferences: {
      contextIsolation: false,
      nodeIntegration: false,
      sandbox: false
    }
  })

  win.webContents.on('console-message', (_event, level, message) => {
    if (/Electron Security Warning \(Insecure Content-Security-Policy\)/.test(message)) return
    if (level >= 2) consoleMessages.push(message)
  })

  await win.loadFile(HARNESS)
  await sleep(250)

  const evalPage = (expression) => win.webContents.executeJavaScript(expression, true)
  const click = (selector) => evalPage(`(() => {
    const el = document.querySelector(${JSON.stringify(selector)});
    if (!el) throw new Error('missing selector: ${selector}');
    el.click();
    return true;
  })()`)

  const typeValue = (selector, value, eventName = 'input') => evalPage(`(() => {
    const el = document.querySelector(${JSON.stringify(selector)});
    if (!el) throw new Error('missing selector: ${selector}');
    el.focus();
    el.value = ${JSON.stringify(value)};
    el.dispatchEvent(new Event(${JSON.stringify(eventName)}, { bubbles: true }));
    return true;
  })()`)

  const initial = await evalPage(`(() => ({
    bodyTextLength: document.body.textContent.trim().length,
    serviceRows: document.querySelectorAll('.service-row').length,
    feedRows: document.querySelectorAll('.feed-item').length,
    teamsRows: Array.from(document.querySelectorAll('.service-row')).filter((row) => /Teams/.test(row.innerText)).length,
    downloadButtonHidden: document.getElementById('dl-toggle-btn').hidden,
    downloadBadge: document.getElementById('dl-badge').textContent,
    focusTitleCount: document.querySelectorAll('.focus-title').length,
    focusClass: document.getElementById('focus').className
  }))()`)
  assert.ok(initial.bodyTextLength > 1000, 'panel rendered with unexpectedly little text')
  assert.equal(initial.serviceRows, 5, 'visible service count did not render')
  assert.ok(initial.feedRows >= 4, 'feed previews did not render')
  assert.equal(initial.teamsRows, 1, 'Teams service row missing')
  assert.equal(initial.downloadButtonHidden, false, 'active download button hidden')
  assert.equal(initial.downloadBadge, '1', 'active download badge missing')
  assert.equal(initial.focusTitleCount, 0, 'old focus title is still in the timer tile')
  assert.equal(/\brunning\b|\bwarning\b|\burgent\b/.test(initial.focusClass), false, 'idle focus tile is colored')

  await click('[data-compose="mail"]')
  await click('[data-compose="calendar"]')
  await click('[data-compose="todo"]')
  await click('#refresh-mail-btn')
  await click('.feed-item')
  await click('#summary-strip')
  let commands = await evalPage('window.__commands')
  assert.ok(hasCommand(commands, { type: 'compose', kind: 'mail' }), 'mail compose command missing')
  assert.ok(hasCommand(commands, { type: 'compose', kind: 'calendar' }), 'calendar compose command missing')
  assert.ok(hasCommand(commands, { type: 'compose', kind: 'todo' }), 'task compose command missing')
  assert.ok(hasCommand(commands, { type: 'refresh-feeds' }), 'feed refresh command missing')
  assert.ok(hasCommand(commands, { type: 'open-feed-item' }), 'feed item click did not dispatch')
  assert.ok(hasCommand(commands, { type: 'switch-service', serviceKey: 'mail' }), 'summary strip did not switch to mail')

  await click('#dl-toggle-btn')
  const downloadOpen = await evalPage(`(() => ({
    open: !document.getElementById('dl-drawer').hidden,
    rows: document.querySelectorAll('.dl-item').length,
    progressWidth: document.querySelector('.dl-progress-fill')?.style.width || ''
  }))()`)
  assert.equal(downloadOpen.open, true, 'download drawer did not open')
  assert.equal(downloadOpen.rows, 1, 'download row missing')
  assert.equal(downloadOpen.progressWidth, '50%', 'download progress did not render')
  await click('.dl-act.danger')
  commands = await evalPage('window.__commands')
  assert.ok(hasCommand(commands, { type: 'download-cancel', id: 1 }), 'download cancel command missing')
  await evalPage('window.__emitDownloads({ list: [], activeCount: 0 })')
  const downloadsCleared = await evalPage(`(() => ({
    buttonHidden: document.getElementById('dl-toggle-btn').hidden,
    drawerHidden: document.getElementById('dl-drawer').hidden,
    rows: document.querySelectorAll('.dl-item').length
  }))()`)
  assert.deepEqual(downloadsCleared, { buttonHidden: true, drawerHidden: true, rows: 0 }, 'empty download state left stale UI')
  await evalPage(`window.__emitDownloads({ activeCount: 0, list: [{ id: 2, filename: 'archive.zip', url: 'https://example.com/archive.zip', state: 'interrupted', receivedBytes: 0, totalBytes: 100, speed: 0, startedAt: Date.now(), retryable: true }] })`)
  await click('#dl-toggle-btn')
  await click('.dl-act[data-action="retry"]')
  commands = await evalPage('window.__commands')
  assert.ok(hasCommand(commands, { type: 'download-retry', id: 2 }), 'download retry command missing')

  await click('#open-settings')
  await click('[data-goto="diagnostics"]')
  let settingsState = await evalPage(`(() => ({
    settingsOpen: document.body.classList.contains('settings-open'),
    title: document.getElementById('set-head-title').textContent,
    diagnosticsRows: document.querySelectorAll('#diagnostics-list .set-list-row').length
  }))()`)
  assert.equal(settingsState.settingsOpen, true, 'settings did not open')
  assert.equal(settingsState.title, 'Diagnostics', 'diagnostics page did not become active')
  assert.ok(settingsState.diagnosticsRows >= 3, 'diagnostics did not render services')

  await click('#set-back')
  await click('[data-goto="debugging"]')
  let debuggingState = await evalPage(`(() => ({
    title: document.getElementById('set-head-title').textContent,
    enabled: document.getElementById('debugging-enabled').checked,
    tabText: document.getElementById('debug-tab-list').innerText,
    authText: document.getElementById('debug-auth-list').innerText,
    notificationText: document.getElementById('debug-notification-list').innerText,
    failureText: document.getElementById('debug-failure-list').innerText,
    repairDisabled: document.getElementById('repair-microsoft-session').disabled,
    clearDisabled: document.getElementById('clear-failure-log').disabled,
    exportDisabled: document.getElementById('export-debug-report').disabled
  }))()`)
  assert.equal(debuggingState.title, 'Debugging', 'debugging page did not become active')
  assert.equal(debuggingState.enabled, false, 'debugging toggle should default off')
  assert.match(debuggingState.tabText, /Debugging tools are off/, 'debugging-off tab state missing')
  assert.match(debuggingState.authText, /Enable debugging/, 'debugging-off auth state missing')
  assert.match(debuggingState.notificationText, /Enable debugging/, 'debugging-off notification state missing')
  assert.match(debuggingState.failureText, /Enable debugging/, 'debugging-off failure state missing')
  assert.equal(debuggingState.repairDisabled, true, 'session repair should be disabled while debugging is off')
  assert.equal(debuggingState.clearDisabled, true, 'failure log clear should be disabled while debugging is off')
  assert.equal(debuggingState.exportDisabled, true, 'debug report export should be disabled while debugging is off')

  await click('#debugging-enabled')
  commands = await evalPage('window.__commands')
  assert.ok(commands.some((command) => command.type === 'update-settings' && command.settings?.debugging?.enabled === true), 'debugging enable setting was not sent')
  await evalPage(`window.__emitSnapshot({
    ...window.__snapshot,
    debugging: { enabled: true },
    diagnostics: { ...window.__snapshot.diagnostics, debugging: { enabled: true } }
  })`)
  debuggingState = await evalPage(`(() => ({
    enabled: document.getElementById('debugging-enabled').checked,
    tabText: document.getElementById('debug-tab-list').innerText,
    authText: document.getElementById('debug-auth-list').innerText,
    notificationText: document.getElementById('debug-notification-list').innerText,
    failureText: document.getElementById('debug-failure-list').innerText,
    repairDisabled: document.getElementById('repair-microsoft-session').disabled,
    clearDisabled: document.getElementById('clear-failure-log').disabled,
    exportDisabled: document.getElementById('export-debug-report').disabled
  }))()`)
  assert.equal(debuggingState.enabled, true, 'debugging toggle did not reflect enabled snapshot')
  assert.equal(debuggingState.repairDisabled, false, 'session repair stayed disabled after enabling debugging')
  assert.equal(debuggingState.clearDisabled, false, 'failure log clear stayed disabled after enabling debugging')
  assert.equal(debuggingState.exportDisabled, false, 'debug report export stayed disabled after enabling debugging')
  assert.match(debuggingState.tabText, /Microsoft session/, 'Microsoft partition debug row missing')
  assert.match(debuggingState.tabText, /Mail/, 'Mail tab debug row missing')
  assert.match(debuggingState.tabText, /queued/, 'per-tab prewarm state missing')
  assert.match(debuggingState.authText, /qa@example\.com/, 'Microsoft auth hint missing')
  assert.match(debuggingState.authText, /silent-failed/, 'Microsoft auth event missing')
  assert.match(debuggingState.authText, /interaction_required/, 'Microsoft auth OAuth error missing')
  assert.match(debuggingState.notificationText, /Baselines/, 'notification baseline row missing')
  assert.match(debuggingState.notificationText, /notification-shown/, 'notification audit event missing')
  assert.match(debuggingState.notificationText, /teams:title-count/, 'notification cooldown row missing')
  assert.match(debuggingState.failureText, /Teams page failed/, 'failure log row missing')
  await click('#repair-microsoft-session')
  await click('#repair-asana-session')
  await click('#clear-failure-log')
  await click('#export-debug-report')
  commands = await evalPage('window.__commands')
  assert.ok(hasCommand(commands, { type: 'repair-partition', partition: 'microsoft' }), 'Microsoft partition repair command missing')
  assert.ok(hasCommand(commands, { type: 'repair-partition', partition: 'asana' }), 'Asana partition repair command missing')
  assert.ok(hasCommand(commands, { type: 'clear-failure-log' }), 'failure log clear command missing')
  assert.ok(hasCommand(commands, { type: 'export-debug-report' }), 'debug report export command missing')

  await click('#set-back')
  await click('[data-goto="connections"]')
  const healthRows = await evalPage(`document.querySelectorAll('#setup-health-list .health-row').length`)
  assert.ok(healthRows >= 2, 'setup health checks did not render')
  await click('#set-back')
  await click('[data-goto="appearance"]')
  await click('input[name="uiDensity"][value="compact"]')
  commands = await evalPage('window.__commands')
  assert.ok(hasCommand(commands, { type: 'update-settings' }), 'density update command missing')
  assert.ok(commands.some((command) => command.type === 'update-settings' && command.settings?.uiDensity === 'compact'), 'compact density was not sent')

  await click('#set-back')
  await click('[data-goto="workspaces"]')
  await typeValue('#workspace-icon', 'D')
  await typeValue('#workspace-color', '#16a34a', 'change')
  await typeValue('#workspace-name', 'Daily triage')
  await click('#save-workspace')
  await evalPage(`Array.from(document.querySelectorAll('#workspace-list .set-mini-btn')).find((btn) => btn.textContent === 'Duplicate').click()`)
  commands = await evalPage('window.__commands')
  assert.ok(commands.some((command) => command.type === 'save-workspace' && command.name === 'Daily triage' && command.icon === 'D' && command.color === '#16a34a'), 'workspace save metadata command missing')
  assert.ok(hasCommand(commands, { type: 'duplicate-workspace', id: 'ws-focus' }), 'workspace duplicate command missing')

  await click('#set-back')
  await click('[data-goto="recents"]')
  await click('#recent-list .recent-row')
  commands = await evalPage('window.__commands')
  assert.ok(hasCommand(commands, { type: 'open-url', url: 'https://outlook.office.com/mail/inbox/id/abc' }), 'recent item did not open URL')

  await click('#set-back')
  await click('[data-goto="downloads"]')
  const downloadHistoryRows = await evalPage(`document.querySelectorAll('#download-history-list .set-list-row').length`)
  assert.ok(downloadHistoryRows >= 1, 'download history did not render in settings')
  await click('#set-back')
  await click('[data-goto="portable"]')
  await click('#export-settings')
  await click('#import-settings')
  commands = await evalPage('window.__commands')
  assert.ok(hasCommand(commands, { type: 'export-settings' }), 'export command missing')
  assert.ok(hasCommand(commands, { type: 'import-settings' }), 'import command missing')

  await click('#command-palette-btn')
  await typeValue('#command-input', 'teams status')
  const paletteRowsBeforeEnter = await evalPage(`Array.from(document.querySelectorAll('.command-row')).map((row) => row.innerText)`)
  await evalPage(`document.querySelector('#command-input').dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))`)
  const teamsPopover = await evalPage(`(() => ({
    paletteHidden: document.getElementById('command-palette').hidden,
    popoverHidden: document.getElementById('teams-presence-popover').hidden,
    firstPresenceFocused: document.activeElement?.matches('[data-teams-presence]') || false,
    activeElement: document.activeElement?.id || document.activeElement?.textContent || '',
    commands: window.__commands
  }))()`)
  assert.equal(teamsPopover.paletteHidden, true, 'command palette did not close after selection')
  assert.equal(teamsPopover.popoverHidden, false, `Teams status popover did not open from command palette. Rows: ${JSON.stringify(paletteRowsBeforeEnter)} State: ${JSON.stringify(teamsPopover)}`)
  assert.equal(teamsPopover.firstPresenceFocused, true, 'Teams status did not focus the first presence control')
  await click('[data-teams-presence="busy"]')
  commands = await evalPage('window.__commands')
  assert.ok(hasCommand(commands, { type: 'set-teams-presence', statusKey: 'busy' }), 'Teams presence command missing')
  await click('#teams-presence-reset')
  commands = await evalPage('window.__commands')
  assert.ok(hasCommand(commands, { type: 'reset-teams-presence' }), 'Teams reset command missing')

  await click('#search-toggle')
  await typeValue('#search-input', 'budget')
  await evalPage(`window.__emitEvent({ type: 'find-result', activeMatchOrdinal: 1, matches: 3 })`)
  const searchState = await evalPage(`(() => ({
    expanded: document.getElementById('search').classList.contains('open'),
    count: document.getElementById('search-count').textContent
  }))()`)
  assert.equal(searchState.expanded, true, 'find UI did not open')
  assert.equal(searchState.count, '1/3', 'find result count did not render')
  commands = await evalPage('window.__commands')
  assert.ok(hasCommand(commands, { type: 'find-in-page', text: 'budget' }), 'find command missing')

  await click('[data-command="close-settings"]')
  await sleep(50)
  await evalPage(`window.__emitSnapshot({
    ...window.__snapshot,
    settingsOpen: false,
    activeServiceKey: 'mail',
    services: window.__snapshot.services.map((service) => service.key === 'mail'
      ? { ...service, health: { state: 'blank', message: 'No visible content', code: 0, at: Date.now() } }
      : service)
  })`)
  const recoveryState = await evalPage(`(() => ({
    hidden: document.getElementById('service-recovery').hidden,
    title: document.getElementById('service-recovery-title').textContent
  }))()`)
  assert.equal(recoveryState.hidden, false, 'service recovery panel did not show for blank service')
  assert.match(recoveryState.title, /Mail needs attention/, 'service recovery title is wrong')
  await click('#recovery-reload')
  await click('#recovery-reset')
  await click('#recovery-browser')
  commands = await evalPage('window.__commands')
  assert.ok(hasCommand(commands, { type: 'repair-service', serviceKey: 'mail', action: 'reload' }), 'recovery reload command missing')
  assert.ok(hasCommand(commands, { type: 'repair-service', serviceKey: 'mail', action: 'reset-home' }), 'recovery reset command missing')
  assert.ok(hasCommand(commands, { type: 'repair-service', serviceKey: 'mail', action: 'open-external' }), 'recovery browser command missing')
  await evalPage(`window.__emitSnapshot({
    ...window.__snapshot,
    services: window.__snapshot.services.map((service) => service.key === 'mail'
      ? { ...service, health: { state: 'ok', message: '', code: 0, at: Date.now() } }
      : service)
  })`)

  await click('#tools-tab')
  await click('[data-focus-minutes="15"]')
  let focusState = await evalPage(`(() => ({
    time: document.getElementById('focus-time').textContent,
    activePreset: document.querySelector('[data-focus-minutes="15"]').classList.contains('active'),
    className: document.getElementById('focus').className
  }))()`)
  assert.equal(focusState.time, '15:00', '15 minute preset did not update timer')
  assert.equal(focusState.activePreset, true, '15 minute preset did not become active')
  assert.equal(/\brunning\b|\bwarning\b|\burgent\b/.test(focusState.className), false, 'unused focus time is colored')

  await click('#focus-toggle')
  focusState = await evalPage(`(() => ({
    className: document.getElementById('focus').className,
    customDisabled: document.getElementById('focus-custom').disabled,
    presetDisabled: document.querySelector('[data-focus-minutes="30"]').disabled
  }))()`)
  assert.match(focusState.className, /\brunning\b/, 'focus timer did not enter running state')
  assert.equal(focusState.customDisabled, true, 'custom focus input stayed enabled while running')
  assert.equal(focusState.presetDisabled, true, 'focus presets stayed enabled while running')

  focusState = await evalPage(`(() => {
    focusRemaining = Math.ceil(focusTotal * 0.2);
    paintFocus();
    const warningClass = document.getElementById('focus').className;
    focusRemaining = Math.ceil(focusTotal * 0.1);
    paintFocus();
    const urgentClass = document.getElementById('focus').className;
    return { warningClass, urgentClass };
  })()`)
  assert.match(focusState.warningClass, /\bwarning\b/, '20% remaining state is not warning')
  assert.doesNotMatch(focusState.warningClass, /\burgent\b/, '20% remaining state became urgent too early')
  assert.match(focusState.urgentClass, /\burgent\b/, '10% remaining state is not urgent/rainbow')
  await click('#focus-reset')
  focusState = await evalPage(`(() => ({
    time: document.getElementById('focus-time').textContent,
    className: document.getElementById('focus').className,
    customDisabled: document.getElementById('focus-custom').disabled
  }))()`)
  assert.equal(focusState.time, '15:00', 'focus reset did not restore selected duration')
  assert.equal(/\brunning\b|\bwarning\b|\burgent\b/.test(focusState.className), false, 'reset focus tile stayed colored')
  assert.equal(focusState.customDisabled, false, 'custom focus input stayed disabled after reset')

  const focusLayout = await evalPage(`(() => {
    const focus = document.getElementById('focus');
    const ring = document.querySelector('.focus-ring-wrap');
    const controls = document.querySelector('.focus-controls');
    const buttons = document.querySelector('.focus-btns');
    const rects = [ring, controls, buttons].map((el) => {
      const r = el.getBoundingClientRect();
      return { left: r.left, right: r.right, top: r.top, bottom: r.bottom, width: r.width, height: r.height };
    });
    const overlap = (a, b) => !(a.right <= b.left || b.right <= a.left || a.bottom <= b.top || b.bottom <= a.top);
    const style = getComputedStyle(focus);
    return {
      display: style.display,
      columns: style.gridTemplateColumns,
      focusWidth: focus.getBoundingClientRect().width,
      rects,
      overlaps: [overlap(rects[0], rects[1]), overlap(rects[1], rects[2]), overlap(rects[0], rects[2])]
    };
  })()`)
  assert.equal(focusLayout.display, 'grid', 'focus tile is not using grid layout')
  assert.ok(focusLayout.focusWidth > 220, `focus tile is too cramped in normal sidebar width: ${JSON.stringify(focusLayout)}`)
  assert.deepEqual(focusLayout.overlaps, [false, false, false], 'focus timer columns overlap')

  await click('#scratch-toggle')
  await typeValue('#scratch-area', 'Remember to follow up')
  await sleep(700)
  commands = await evalPage('window.__commands')
  assert.ok(hasCommand(commands, { type: 'save-scratch', text: 'Remember to follow up' }), 'scratchpad save command missing')

  if (consoleMessages.length) {
    issues.push(...consoleMessages.map((message) => `console: ${message}`))
  }

  const report = {
    behaviorChecks: [
      'initial render',
      'compose/feed/summary commands',
      'downloads drawer lifecycle and retry',
      'settings diagnostics/setup health/workspaces/density/recents/import-export',
      'debugging toggle, tab observability, session repair, failure log, and debug export',
      'command palette to Teams status',
      'find in page',
      'blank-service recovery actions',
      'focus timer thresholds and layout',
      'scratchpad save'
    ],
    issues
  }

  console.log(JSON.stringify(report, null, 2))
  assert.equal(issues.length, 0, 'renderer console errors occurred')

  win.destroy()
  app.quit()
}

main().catch((err) => {
  console.error(err && err.stack ? err.stack : err)
  app.quit()
  process.exit(1)
})
