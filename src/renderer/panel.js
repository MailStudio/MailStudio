const serviceList = document.getElementById('service-list')
const toggleButton = document.getElementById('toggle-sidebar')
const themeToggle = document.getElementById('theme-toggle')
const navBack = document.getElementById('nav-back')
const navForward = document.getElementById('nav-forward')
const setItems = document.getElementById('set-items')
const addName = document.getElementById('add-name')
const addUrl = document.getElementById('add-url')
const addBtn = document.getElementById('add-site')
const addError = document.getElementById('add-error')
const summaryStrip = document.getElementById('summary-strip')
const searchBox = document.getElementById('search')
const searchInput = document.getElementById('search-input')
const searchCount = document.getElementById('search-count')
const scratchArea = document.getElementById('scratch-area')

// Default theme until the first snapshot arrives.
document.documentElement.setAttribute('data-theme', 'dark')

const SERVICE_ICONS = {
  teams:
    '<svg viewBox="0 0 24 24"><circle cx="9" cy="8.5" r="3.2"/><path d="M3.6 19a5.4 5.4 0 0 1 10.8 0"/><circle cx="17.6" cy="9.5" r="2.3"/><path d="M16.2 19a4.4 4.4 0 0 1 5-4.4"/></svg>',
  mail: '<svg viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="14" rx="2.5"/><path d="M4 7l8 6 8-6"/></svg>',
  calendar:
    '<svg viewBox="0 0 24 24"><rect x="3.5" y="5" width="17" height="15" rx="2.5"/><path d="M3.5 9.5h17M8 3.5v3M16 3.5v3"/></svg>',
  todo:
    '<svg viewBox="0 0 24 24"><path d="M9 6h11M9 12h11M9 18h11"/><path d="M4 6.2l1.1 1.1L7 5.4M4 17.2l1.1 1.1L7 16.4"/></svg>',
  asana:
    '<svg viewBox="0 0 24 24"><circle cx="12" cy="6.5" r="3"/><circle cx="6.5" cy="15" r="3"/><circle cx="17.5" cy="15" r="3"/></svg>',
  office:
    '<svg viewBox="0 0 24 24"><path d="M13 3 5 6v12l8 3 6-2.5V5.5L13 3zM13 3v18M13 7l6-1.5M13 17l6 1.5"/></svg>',
  word:
    '<svg viewBox="0 0 24 24"><rect x="4" y="4" width="16" height="16" rx="2.5"/><path d="M7.5 9l1.6 6 1.4-4.2L11.9 15l1.6-6"/></svg>',
  excel:
    '<svg viewBox="0 0 24 24"><rect x="4" y="4" width="16" height="16" rx="2.5"/><path d="M8 9l4 6M12 9l-4 6M15 9v6"/></svg>',
  powerpoint:
    '<svg viewBox="0 0 24 24"><rect x="4" y="4" width="16" height="16" rx="2.5"/><path d="M9 16V8h3a2.5 2.5 0 0 1 0 5H9"/></svg>',
  onenote:
    '<svg viewBox="0 0 24 24"><rect x="4" y="4" width="16" height="16" rx="2.5"/><path d="M8.5 16V8l4 8V8M16 4v16"/></svg>',
  onedrive:
    '<svg viewBox="0 0 24 24"><path d="M7 17h10.5a3 3 0 0 0 .3-6A4.5 4.5 0 0 0 9 9.5 3.8 3.8 0 0 0 7 17z"/></svg>',
  planner:
    '<svg viewBox="0 0 24 24"><rect x="4" y="5" width="16" height="15" rx="2.5"/><path d="M4 9h16M8 3v4M16 3v4M8 13l1.5 1.5L12 12M8 17l1.5 1.5L12 16"/></svg>',
  sharepoint:
    '<svg viewBox="0 0 24 24"><circle cx="9" cy="7" r="3.2"/><circle cx="16.5" cy="11" r="2.6"/><circle cx="11" cy="16.5" r="2.8"/><path d="M11.5 9.5l2.7 0.8M13.8 13l-1 1.8"/></svg>',
  link: '<svg viewBox="0 0 24 24"><path d="M10 13a5 5 0 0 0 7 0l2-2a5 5 0 0 0-7-7l-1 1"/><path d="M14 11a5 5 0 0 0-7 0l-2 2a5 5 0 0 0 7 7l1-1"/></svg>'
}

const SUN_ICON =
  '<svg viewBox="0 0 24 24" class="icon"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>'
const MOON_ICON =
  '<svg viewBox="0 0 24 24" class="icon"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg>'
const ARROW_UP = '<svg viewBox="0 0 24 24" class="icon"><path d="M12 19V6M6 12l6-6 6 6"/></svg>'
const ARROW_DOWN = '<svg viewBox="0 0 24 24" class="icon"><path d="M12 5v13M6 12l6 6 6-6"/></svg>'
const TRASH = '<svg viewBox="0 0 24 24" class="icon"><path d="M4 7h16M9 7V5h6v2M6 7l1 13h10l1-13"/></svg>'
// Split-pane glyphs: a framed rectangle with the occupied half filled, so a tab
// in split view shows whether it's the left or right pane.
const SPLIT_LEFT_ICON =
  '<svg viewBox="0 0 24 24"><rect x="3.5" y="5.5" width="17" height="13" rx="2"/><rect x="4.3" y="6.3" width="7.3" height="11.4" rx="1.2" fill="currentColor" stroke="none"/></svg>'
const SPLIT_RIGHT_ICON =
  '<svg viewBox="0 0 24 24"><rect x="3.5" y="5.5" width="17" height="13" rx="2"/><rect x="12.4" y="6.3" width="7.3" height="11.4" rx="1.2" fill="currentColor" stroke="none"/></svg>'
// Multi-select modifier for building a split: Cmd on macOS, Ctrl elsewhere.
const SPLIT_MOD_LABEL = navigator.platform.toUpperCase().includes('MAC') ? '⌘' : 'Ctrl'

let latest = null
let dragIndex = null
// Pending scratch-pad save debounce (null when no save is in flight).
let scratchTimer = null
// Provider keys whose "Developer setup" panel the user explicitly opened, so
// re-renders don't snap it shut while a saved client ID is being edited.
const devOpen = new Set()
// Paged settings: which sub-page is showing, and the prior open-state so we can
// reset to the root list each time Settings is reopened.
let settingsPage = 'root'
let prevSettingsOpen = false

function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme)
  themeToggle.innerHTML = theme === 'dark' ? MOON_ICON : SUN_ICON
  themeToggle.title = theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode'
}

/* ---------- Live feeds ---------- */
function renderFeed(service) {
  const feed = service.feed
  const wrap = document.createElement('div')
  wrap.className = `feed feed-${feed.kind}`

  const note = (text) => {
    const n = document.createElement('div')
    n.className = 'feed-note'
    n.textContent = text
    wrap.appendChild(n)
  }

  if (feed.state === 'loading') {
    note('Loading…')
    return wrap
  }
  if (feed.state === 'login') {
    // Scrape-backed feed is signed out — jump to the tab for a plain web sign-in.
    const btn = document.createElement('button')
    btn.className = 'feed-connect'
    btn.textContent = 'Sign in →'
    btn.addEventListener('click', () => {
      window.panelApi.sendCommand({ type: 'switch-service', serviceKey: service.key })
    })
    wrap.appendChild(btn)
    return wrap
  }
  if (feed.state === 'auth') {
    // API provider is configured but not (or no longer) connected.
    const btn = document.createElement('button')
    btn.className = 'feed-connect'
    btn.textContent = feed.kind === 'asana' ? 'Connect Asana' : 'Connect Microsoft'
    btn.addEventListener('click', () => openOnboarding())
    wrap.appendChild(btn)
    return wrap
  }
  if (feed.state === 'error') {
    note(feed.kind === 'mail' ? "Couldn't read inbox" : feed.kind === 'calendar' ? "Couldn't read calendar" : "Couldn't read tasks")
    const retry = document.createElement('button')
    retry.className = 'feed-retry'
    retry.textContent = 'Retry'
    retry.addEventListener('click', () => {
      window.panelApi.sendCommand({ type: 'refresh-feeds' })
    })
    wrap.appendChild(retry)
    return wrap
  }
  if (feed.state === 'empty' || !feed.items.length) {
    note(feed.kind === 'mail' ? 'No unread mail' : feed.kind === 'calendar' ? 'No upcoming events' : 'No tasks found')
    return wrap
  }

  for (const item of feed.items) {
    const row = document.createElement('div')
    row.className = 'feed-item'

    if (feed.kind === 'mail') {
      row.classList.add('feed-mail')
      row.innerHTML = `
        <div class="feed-sender">${escapeHtml(cleanText(item.sender) || 'Unknown')}</div>
        <div class="feed-subject">${escapeHtml(cleanText(item.subject) || '(no subject)')}</div>
      `
      row.addEventListener('click', (event) => {
        window.panelApi.sendCommand({
          type: 'open-feed-item',
          serviceKey: service.key,
          rowIdx: item.rowIdx,
          webLink: item.webLink || null,
          split: event.metaKey || event.ctrlKey
        })
      })
    } else if (feed.kind === 'calendar') {
      row.classList.add('feed-event')
      row.innerHTML = `
        <div class="feed-event-title">${escapeHtml(cleanText(item.title) || 'Event')}</div>
        ${item.time ? `<div class="feed-event-time">${escapeHtml(item.time)}</div>` : ''}
      `
      row.addEventListener('click', (event) => {
        window.panelApi.sendCommand({
          type: 'open-feed-item',
          serviceKey: service.key,
          rowIdx: item.rowIdx,
          webLink: item.webLink || null,
          split: event.metaKey || event.ctrlKey
        })
      })
    } else {
      row.classList.add('feed-task')
      const subs = (item.subtasks || [])
        .map((s) => cleanText(s))
        .filter(Boolean)
        .slice(0, 4)
        .map((s) => `<div class="feed-subtask">${escapeHtml(s)}</div>`)
        .join('')
      row.innerHTML = `<div class="feed-task-name">${escapeHtml(cleanText(item.name) || 'Task')}</div>${subs}`
      row.addEventListener('click', (event) => {
        window.panelApi.sendCommand({
          type: 'open-feed-item',
          serviceKey: service.key,
          rowIdx: item.rowIdx,
          taskUrl: item.taskUrl || null,
          split: event.metaKey || event.ctrlKey
        })
      })
    }

    wrap.appendChild(row)
  }
  return wrap
}

// Services whose right-click context menu offers snooze options.
const SNOOZABLE_KEYS = new Set(['mail', 'teams', 'calendar', 'asana'])

function renderServices(snapshot) {
  serviceList.innerHTML = ''
  const splitKeys = Array.isArray(snapshot.splitKeys) ? snapshot.splitKeys : []
  for (const service of snapshot.services) {
    if (!service.visible) continue

    const group = document.createElement('div')
    group.className = 'svc-group'

    const splitIndex = splitKeys.indexOf(service.key)
    const inSplit = splitIndex !== -1
    const isActive = service.key === snapshot.activeServiceKey

    const button = document.createElement('button')
    const classes = ['service']
    if (isActive || inSplit) classes.push('active')
    if (inSplit) classes.push('in-split')
    if (inSplit && isActive) classes.push('split-focus')
    button.className = classes.join(' ')
    button.type = 'button'
    button.title = inSplit
      ? `${service.label} · Split pane ${splitIndex + 1} (${SPLIT_MOD_LABEL}-click to remove)`
      : `${service.label} · ${SPLIT_MOD_LABEL}-click to open beside the current tab`
    if (SNOOZABLE_KEYS.has(service.key)) button.title += ' · right-click for snooze & more'

    const hasUnread = service.unreadCount > 0
    const count = service.unreadCount
    const snoozeIcon = service.snoozed
      ? '<span class="service-snooze" title="Notifications snoozed"><svg viewBox="0 0 24 24"><path d="M6 5h6L6 12h6M14 11h5l-5 6h5"/></svg></span>'
      : ''
    const splitIcon = inSplit
      ? `<span class="service-split">${splitIndex === 0 ? SPLIT_LEFT_ICON : SPLIT_RIGHT_ICON}</span>`
      : ''
    button.innerHTML = `
      <span class="service-icon">${SERVICE_ICONS[service.icon] || SERVICE_ICONS.link}</span>
      <span class="service-name">${escapeHtml(service.label)}</span>
      ${splitIcon}
      ${snoozeIcon}
      <span class="service-badge${hasUnread ? '' : ' hidden'}">${count}</span>
    `
    button.addEventListener('click', (event) => {
      // Cmd-click (macOS) / Ctrl-click (Windows/Linux) pairs this tab with the
      // active one (or toggles it within an existing split); a plain click
      // switches to it and exits split view.
      const type = event.metaKey || event.ctrlKey ? 'split-select' : 'switch-service'
      window.panelApi.sendCommand({ type, serviceKey: service.key })
    })
    button.addEventListener('contextmenu', (e) => {
      e.preventDefault()
      window.panelApi.sendCommand({ type: 'tab-context-menu', serviceKey: service.key })
    })
    group.appendChild(button)

    if (service.feed) {
      group.appendChild(renderFeed(service))
    }
    serviceList.appendChild(group)
  }
}

/* ---------- Settings ---------- */
function workingServices() {
  // A plain, editable copy of the configured services from the last snapshot.
  return (latest ? latest.services : []).map((s) => ({
    key: s.key,
    label: s.label,
    url: s.url,
    home: s.home,
    visible: s.visible,
    builtin: s.builtin
  }))
}

function sendSettings(services, collapseMode) {
  window.panelApi.sendCommand({
    type: 'update-settings',
    settings: {
      collapseMode: collapseMode || (latest ? latest.collapseMode : 'vanish'),
      notif: latest ? latest.notif : null,
      services
    }
  })
}

function sendNotif(notif) {
  window.panelApi.sendCommand({
    type: 'update-settings',
    settings: {
      collapseMode: latest ? latest.collapseMode : 'vanish',
      services: workingServices(),
      notif
    }
  })
}

function renderSettings(snapshot) {
  // Radios
  for (const radio of document.querySelectorAll('input[name="collapseMode"]')) {
    radio.checked = radio.value === snapshot.collapseMode
  }
  for (const radio of document.querySelectorAll('input[name="taskProvider"]')) {
    radio.checked = radio.value === (snapshot.taskProvider || 'microsoft')
  }

  // Quiet hours: reflect the persisted range (skip an input being edited so a
  // feed-poll snapshot can't clobber a half-typed time).
  const notif = snapshot.notif || {}
  for (const [id, key] of [['quiet-start', 'quietStart'], ['quiet-end', 'quietEnd']]) {
    const el = document.getElementById(id)
    if (el && document.activeElement !== el) el.value = notif[key] || ''
  }
  const quietSub = document.querySelector('.set-nav-row[data-goto="quiet"] small')
  if (quietSub) {
    quietSub.textContent = notif.quietStart && notif.quietEnd
      ? 'Silence within a time range · active'
      : 'Silence within a time range'
  }

  setItems.innerHTML = ''
  const services = snapshot.services
  services.forEach((service, index) => {
    const row = document.createElement('div')
    row.className = `set-item${service.visible ? '' : ' hidden-item'}`
    row.draggable = true
    row.dataset.index = index

    // HTML5 drag-to-reorder. The custom type marks the drag as one of our own
    // rows, so stray external drags can't trigger a reorder.
    row.addEventListener('dragstart', (e) => {
      dragIndex = index
      row.classList.add('dragging')
      e.dataTransfer.effectAllowed = 'move'
      e.dataTransfer.setData('application/x-orbit-row', String(index))
    })
    row.addEventListener('dragend', () => {
      dragIndex = null
      row.classList.remove('dragging')
      document.querySelectorAll('.set-item.drag-over').forEach((el) => el.classList.remove('drag-over'))
    })
    row.addEventListener('dragover', (e) => {
      if (!e.dataTransfer.types.includes('application/x-orbit-row')) return
      e.preventDefault()
      if (dragIndex !== null && dragIndex !== index) row.classList.add('drag-over')
    })
    row.addEventListener('dragleave', () => row.classList.remove('drag-over'))
    row.addEventListener('drop', (e) => {
      if (!e.dataTransfer.types.includes('application/x-orbit-row')) return
      e.preventDefault()
      row.classList.remove('drag-over')
      if (dragIndex === null || dragIndex === index) return
      const list = workingServices()
      const [moved] = list.splice(dragIndex, 1)
      list.splice(index, 0, moved)
      dragIndex = null
      sendSettings(list)
    })

    const up = document.createElement('button')
    up.className = 'set-icon-btn'
    up.innerHTML = ARROW_UP
    up.title = 'Move up'
    up.disabled = index === 0
    up.addEventListener('click', () => {
      const list = workingServices()
      ;[list[index - 1], list[index]] = [list[index], list[index - 1]]
      sendSettings(list)
    })

    const down = document.createElement('button')
    down.className = 'set-icon-btn'
    down.innerHTML = ARROW_DOWN
    down.title = 'Move down'
    down.disabled = index === services.length - 1
    down.addEventListener('click', () => {
      const list = workingServices()
      ;[list[index + 1], list[index]] = [list[index], list[index + 1]]
      sendSettings(list)
    })

    const reorder = document.createElement('div')
    reorder.className = 'set-reorder'
    reorder.append(up, down)

    const name = document.createElement('span')
    name.className = 'set-item-name'
    name.textContent = service.label
    name.title = service.builtin ? `${service.label} (built-in)` : `${service.label} (pinned)`

    const sw = document.createElement('button')
    sw.className = `set-switch${service.visible ? ' on' : ''}`
    sw.title = service.visible ? 'Hide from sidebar' : 'Show in sidebar'
    sw.addEventListener('click', () => {
      const list = workingServices()
      list[index].visible = !list[index].visible
      sendSettings(list)
    })

    row.append(reorder, name, sw)

    if (!service.builtin) {
      const del = document.createElement('button')
      del.className = 'set-icon-btn danger'
      del.innerHTML = TRASH
      del.title = 'Remove pinned site'
      del.addEventListener('click', () => {
        const list = workingServices().filter((s) => s.key !== service.key)
        sendSettings(list)
      })
      row.appendChild(del)
    }

    setItems.appendChild(row)
  })
}

function addSite() {
  addError.textContent = ''
  const label = addName.value.trim()
  let url = addUrl.value.trim()
  if (!url) {
    addError.textContent = 'Enter a site address.'
    return
  }
  if (!/^https?:\/\//i.test(url)) {
    url = `https://${url}`
  }
  let valid
  try {
    valid = new URL(url).toString()
  } catch {
    addError.textContent = "That doesn't look like a valid URL."
    return
  }
  const list = workingServices()
  list.push({
    key: `site-${Date.now().toString(36)}`,
    label: label || new URL(valid).hostname.replace(/^www\./, ''),
    url: valid,
    home: valid,
    visible: true,
    builtin: false
  })
  sendSettings(list)
  addName.value = ''
  addUrl.value = ''
}

/* ---------- Settings pages ---------- */
const PAGE_TITLES = {
  root: 'Settings',
  connections: 'Connections',
  notifications: 'Notifications',
  tasks: 'Tasks',
  quiet: 'Quiet hours',
  sidebar: 'Sidebar items',
  appearance: 'Sidebar behavior'
}

function showSetPage(page) {
  settingsPage = PAGE_TITLES[page] ? page : 'root'
  for (const el of document.querySelectorAll('.set-page')) {
    el.hidden = el.id !== `set-page-${settingsPage}`
  }
  const title = document.getElementById('set-head-title')
  if (title) title.textContent = PAGE_TITLES[settingsPage]
  const back = document.getElementById('set-back')
  if (back) back.hidden = settingsPage === 'root'
  if (settingsPage === 'connections' && latest) {
    renderConnCards(latest, document.getElementById('set-conn-cards'))
  }
}

/* ---------- Notification settings ---------- */
const NOTIF_PROVIDER_STATUS = {
  microsoft: { sub: 'Requires Microsoft' },
  asana: { sub: 'Requires Asana' },
  teams: { sub: 'From the Teams tab' }
}

function renderNotifSettings(snapshot) {
  const notif = snapshot.notif || {}
  for (const btn of document.querySelectorAll('[data-notif]')) {
    const on = notif[btn.dataset.notif] !== false
    btn.classList.toggle('on', on)
  }
  // Reflect live connection state per row: a service whose provider isn't
  // connected shows a muted "connect first" hint and can't actually fire.
  const conn = snapshot.connections || {}
  for (const row of document.querySelectorAll('.set-notif-row[data-notif-provider]')) {
    const provider = row.dataset.notifProvider
    const connected = provider === 'teams' ? true : Boolean(conn[provider] && conn[provider].status === 'connected')
    row.classList.toggle('disconnected', !connected)
    const sub = row.querySelector('[data-notif-sub]')
    if (sub) {
      // Capture the original copy once so we can restore it on reconnect.
      if (sub.dataset.default === undefined) sub.dataset.default = sub.textContent
      sub.textContent = connected
        ? sub.dataset.default
        : `${(NOTIF_PROVIDER_STATUS[provider] || {}).sub || 'Not connected'} — connect to enable`
    }
  }
  // A provider is connected but onboarding never finished, so nothing fires.
  const warn = document.getElementById('notif-arm-warn')
  if (warn) {
    const anyConnected = PROVIDERS.some((p) => conn[p] && conn[p].status === 'connected')
    warn.hidden = !(anyConnected && !snapshot.onboarded)
  }
}

/* ---------- Summary strip ---------- */
function renderSummary(snapshot) {
  let mail = 0
  let events = 0
  let tasks = 0
  const todayStr = new Date().toDateString()
  for (const service of snapshot.services) {
    if (!service.visible || !service.feed) continue
    if (service.feed.kind === 'mail') {
      // Unread emails received today: scraped rows carry a `today` flag, API
      // rows the Graph receivedDateTime.
      mail = (service.feed.items || []).filter(
        (item) => item.today === true || (item.receivedIso && new Date(item.receivedIso).toDateString() === todayStr)
      ).length
    } else if (service.feed.kind === 'calendar') events = (service.feed.items || []).length
    else if (service.feed.kind === 'asana') tasks = (service.feed.items || []).length
  }
  const parts = []
  if (mail > 0) parts.push(`<b>${mail > 9 ? '10+' : mail}</b> email${mail === 1 ? '' : 's'} today`)
  if (events > 0) parts.push(`<b>${events}</b> event${events === 1 ? '' : 's'}`)
  if (tasks > 0) parts.push(`<b>${tasks}</b> task${tasks === 1 ? '' : 's'}`)

  if (!parts.length) {
    summaryStrip.hidden = true
    return
  }
  summaryStrip.hidden = false
  summaryStrip.innerHTML = `<span class="summary-dot"></span><span class="summary-text">${parts.join(' · ')}</span>`
}

/* ---------- Main render ---------- */
function render(snapshot) {
  latest = snapshot
  applyTheme(snapshot.theme === 'light' ? 'light' : 'dark')
  document.documentElement.toggleAttribute('data-glass', Boolean(snapshot.glassMode))

  document.body.classList.toggle('collapsed', Boolean(snapshot.sidebarCollapsed))
  document.body.classList.toggle('mode-rail', snapshot.collapseMode === 'rail')
  document.body.classList.toggle('mode-vanish', snapshot.collapseMode !== 'rail')
  document.body.classList.toggle('settings-open', Boolean(snapshot.settingsOpen))
  document.body.classList.toggle('first-boot', Boolean(snapshot.firstBoot))

  toggleButton.title = snapshot.sidebarCollapsed ? 'Show sidebar' : 'Hide sidebar'

  const nav = snapshot.nav || {}
  navBack.disabled = !nav.canGoBack
  navForward.disabled = !nav.canGoForward

  const snoozeBtn = document.getElementById('global-snooze')
  if (snoozeBtn) {
    const active = Boolean(snapshot.globalSnoozed)
    snoozeBtn.classList.toggle('active', active)
    snoozeBtn.title = active ? 'Notifications snoozed — click to resume' : 'Snooze all notifications & sound'
  }

  renderServices(snapshot)
  renderSummary(snapshot)
  renderNotifSetup(snapshot)
  renderSplitControls(snapshot)

  // Keep the onboarding sheet live while open — but don't yank an input the
  // user is mid-edit (frequent feed-poll snapshots would clobber typed IDs).
  if (onboardOverlay && !onboardOverlay.hidden) {
    const ae = document.activeElement
    const typing = ae && ae.closest && ae.closest('.conn-card')
    if (!typing) renderConnCards(snapshot)
  }

  // Sync scratch pad from persisted state (only when not actively editing and
  // no save is pending — a snapshot taken before the debounce fires still
  // carries the old text and would revert what was just typed).
  if (document.activeElement !== scratchArea && scratchTimer === null && typeof snapshot.scratch === 'string') {
    scratchArea.value = snapshot.scratch
  }

  if (snapshot.settingsOpen) {
    // Reset to the root list each time Settings is (re)opened.
    if (!prevSettingsOpen) showSetPage('root')
    renderSettings(snapshot)
    renderNotifSettings(snapshot)
    // Keep the Connections page cards live (without yanking a field being typed).
    if (settingsPage === 'connections') {
      const ae = document.activeElement
      const typing = ae && ae.closest && ae.closest('.conn-card')
      if (!typing) renderConnCards(snapshot, document.getElementById('set-conn-cards'))
      const warn = document.getElementById('set-encwarn')
      if (warn) warn.hidden = !(snapshot.connections && snapshot.connections.encryptionAvailable === false)
    }
  }
  prevSettingsOpen = Boolean(snapshot.settingsOpen)
}

/* ---------- Split view: orientation toggle + draggable divider ---------- */
// These constants mirror main.js (SPLIT_GUTTER / ratio bounds) and the CSS
// sidebar/topbar sizes so the divider lands exactly in the gutter between panes.
const SPLIT_GUTTER = 8
const SPLIT_MIN = 0.15
const SPLIT_MAX = 0.85
const SB_EXPANDED = 280
const SB_RAIL = 76
const TOPBAR_H = 46

const splitOrientBtn = document.getElementById('split-orient')
const splitDivider = document.getElementById('split-divider')
const splitGhost = document.getElementById('split-ghost')
const splitGhostA = document.getElementById('split-ghost-a')
const splitGhostB = document.getElementById('split-ghost-b')
const splitGhostBar = document.getElementById('split-ghost-bar')

const ORIENT_VERTICAL_ICON =
  '<svg viewBox="0 0 24 24" class="icon"><rect x="3.5" y="5" width="17" height="14" rx="2.5"/><path d="M12 5v14"/></svg>'
const ORIENT_HORIZONTAL_ICON =
  '<svg viewBox="0 0 24 24" class="icon"><rect x="3.5" y="5" width="17" height="14" rx="2.5"/><path d="M3.5 12h17"/></svg>'

let splitDrag = null // { horizontal, rect, ratio } while the divider is dragged

function splitActive(snap) {
  return Boolean(snap && Array.isArray(snap.splitKeys) && snap.splitKeys.length === 2)
}

function clampRatio(r) {
  return Math.min(SPLIT_MAX, Math.max(SPLIT_MIN, r))
}

// Content area to the right of the sidebar / below the topbar, in CSS px.
// Mirrors main's sidebarWidth() so the divider aligns with the real gutter.
function contentRect(snap) {
  let left
  if (snap.settingsOpen || !snap.sidebarCollapsed) left = SB_EXPANDED
  else left = snap.collapseMode === 'rail' ? SB_RAIL : 0
  return {
    left,
    top: TOPBAR_H,
    width: Math.max(1, window.innerWidth - left),
    height: Math.max(1, window.innerHeight - TOPBAR_H)
  }
}

function positionSplitDivider() {
  if (!splitDivider || splitDrag) return
  const onboardOpen = onboardOverlay && !onboardOverlay.hidden
  if (!latest || !splitActive(latest) || latest.firstBoot || onboardOpen) {
    splitDivider.hidden = true
    return
  }
  const horizontal = latest.splitOrientation === 'horizontal'
  const r = contentRect(latest)
  const ratio = clampRatio(typeof latest.splitRatio === 'number' ? latest.splitRatio : 0.5)
  splitDivider.classList.toggle('horizontal', horizontal)
  splitDivider.classList.toggle('vertical', !horizontal)
  if (horizontal) {
    const topH = Math.round((r.height - SPLIT_GUTTER) * ratio)
    splitDivider.style.cssText = `left:${r.left}px;top:${r.top + topH}px;width:${r.width}px;height:${SPLIT_GUTTER}px`
  } else {
    const leftW = Math.round((r.width - SPLIT_GUTTER) * ratio)
    splitDivider.style.cssText = `left:${r.left + leftW}px;top:${r.top}px;width:${SPLIT_GUTTER}px;height:${r.height}px`
  }
  splitDivider.hidden = false
}

function renderSplitControls(snapshot) {
  const active = splitActive(snapshot)
  if (splitOrientBtn) {
    const horizontal = snapshot.splitOrientation === 'horizontal'
    splitOrientBtn.innerHTML = horizontal ? ORIENT_HORIZONTAL_ICON : ORIENT_VERTICAL_ICON
    splitOrientBtn.disabled = !active
    splitOrientBtn.title = !active
      ? `Split layout — ${SPLIT_MOD_LABEL}-click two tabs to split`
      : horizontal
        ? 'Split: stacked — click for side by side'
        : 'Split: side by side — click for stacked'
  }
  positionSplitDivider()
}

// Live preview drawn over the (detached) content area while dragging.
function layoutGhost(ratio) {
  if (!splitDrag) return
  const r = splitDrag.rect
  splitGhost.style.cssText = `left:${r.left}px;top:${r.top}px;width:${r.width}px;height:${r.height}px`
  splitGhost.classList.toggle('horizontal', splitDrag.horizontal)
  splitGhost.classList.toggle('vertical', !splitDrag.horizontal)
  if (splitDrag.horizontal) {
    const topH = Math.round((r.height - SPLIT_GUTTER) * ratio)
    splitGhostA.style.cssText = `left:0;top:0;width:${r.width}px;height:${topH}px`
    splitGhostBar.style.cssText = `left:0;top:${topH}px;width:${r.width}px;height:${SPLIT_GUTTER}px`
    splitGhostB.style.cssText = `left:0;top:${topH + SPLIT_GUTTER}px;width:${r.width}px;height:${r.height - SPLIT_GUTTER - topH}px`
  } else {
    const leftW = Math.round((r.width - SPLIT_GUTTER) * ratio)
    splitGhostA.style.cssText = `left:0;top:0;width:${leftW}px;height:${r.height}px`
    splitGhostBar.style.cssText = `left:${leftW}px;top:0;width:${SPLIT_GUTTER}px;height:${r.height}px`
    splitGhostB.style.cssText = `left:${leftW + SPLIT_GUTTER}px;top:0;width:${r.width - SPLIT_GUTTER - leftW}px;height:${r.height}px`
  }
}

function onSplitDragMove(e) {
  if (!splitDrag) return
  const r = splitDrag.rect
  const ratio = clampRatio(
    splitDrag.horizontal ? (e.clientY - r.top) / r.height : (e.clientX - r.left) / r.width
  )
  splitDrag.ratio = ratio
  layoutGhost(ratio)
}

function endSplitDrag() {
  if (!splitDrag) return
  const ratio = splitDrag.ratio
  splitDrag = null
  document.removeEventListener('mousemove', onSplitDragMove, true)
  document.removeEventListener('mouseup', endSplitDrag, true)
  window.removeEventListener('blur', endSplitDrag)
  if (splitGhost) splitGhost.hidden = true
  document.body.classList.remove('split-dragging', 'split-dragging-h', 'split-dragging-v')
  window.panelApi.sendCommand({ type: 'split-drag-end', ratio })
  positionSplitDivider()
}

if (splitOrientBtn) {
  splitOrientBtn.addEventListener('click', () => {
    window.panelApi.sendCommand({ type: 'toggle-split-orientation' })
  })
}

if (splitDivider && splitGhost) {
  splitDivider.addEventListener('mousedown', (e) => {
    if (!latest || !splitActive(latest)) return
    e.preventDefault()
    const horizontal = latest.splitOrientation === 'horizontal'
    const ratio = clampRatio(typeof latest.splitRatio === 'number' ? latest.splitRatio : 0.5)
    splitDrag = { horizontal, rect: contentRect(latest), ratio }
    document.body.classList.add('split-dragging')
    document.body.classList.toggle('split-dragging-h', horizontal)
    document.body.classList.toggle('split-dragging-v', !horizontal)
    splitGhost.hidden = false
    layoutGhost(ratio)
    // Ask main to detach the panes so this drag owns the whole content area.
    window.panelApi.sendCommand({ type: 'split-drag-start' })
    document.addEventListener('mousemove', onSplitDragMove, true)
    document.addEventListener('mouseup', endSplitDrag, true)
    window.addEventListener('blur', endSplitDrag)
  })
}

// Main repositions the web views on window resize but doesn't push a snapshot,
// so keep the divider handle in sync from here.
window.addEventListener('resize', positionSplitDivider)

/* ---------- Wiring ---------- */
toggleButton.addEventListener('click', () => {
  window.panelApi.sendCommand({ type: 'toggle-sidebar' })
})

for (const btn of document.querySelectorAll('[data-notif]')) {
  btn.addEventListener('click', () => {
    const notif = Object.assign({ mail: true, calendar: true, asana: true, teams: true, preview: true, quietStart: '', quietEnd: '' }, latest ? latest.notif : {})
    notif[btn.dataset.notif] = !notif[btn.dataset.notif]
    sendNotif(notif)
  })
}

// Settings page navigation (root list rows + the "Manage connections" link).
for (const row of document.querySelectorAll('[data-goto]')) {
  row.addEventListener('click', () => showSetPage(row.dataset.goto))
}
const setBackBtn = document.getElementById('set-back')
if (setBackBtn) setBackBtn.addEventListener('click', () => showSetPage('root'))

function saveQuietHours() {
  const notif = Object.assign({ mail: true, teams: true, preview: true, quietStart: '', quietEnd: '' }, latest ? latest.notif : {})
  notif.quietStart = (document.getElementById('quiet-start') || {}).value || ''
  notif.quietEnd = (document.getElementById('quiet-end') || {}).value || ''
  sendNotif(notif)
}

for (const id of ['quiet-start', 'quiet-end']) {
  const el = document.getElementById(id)
  if (el) el.addEventListener('change', saveQuietHours)
}

// "Notifications not armed" warning: a provider is connected but onboarding
// never finished — arm them in place.
const notifArmBtn = document.getElementById('notif-arm-btn')
if (notifArmBtn) {
  notifArmBtn.addEventListener('click', () => {
    window.panelApi.sendCommand({ type: 'finish-onboarding' })
  })
}

themeToggle.addEventListener('click', () => {
  const next = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark'
  window.panelApi.sendCommand({ type: 'set-theme', theme: next })
})

for (const radio of document.querySelectorAll('input[name="collapseMode"]')) {
  radio.addEventListener('change', () => {
    if (radio.checked) sendSettings(workingServices(), radio.value)
  })
}

for (const radio of document.querySelectorAll('input[name="taskProvider"]')) {
  radio.addEventListener('change', () => {
    if (!radio.checked) return
    window.panelApi.sendCommand({
      type: 'update-settings',
      settings: {
        collapseMode: latest ? latest.collapseMode : 'vanish',
        notif: latest ? latest.notif : null,
        services: workingServices(),
        taskProvider: radio.value
      }
    })
  })
}

addBtn.addEventListener('click', addSite)
addUrl.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') addSite()
})

for (const button of document.querySelectorAll('[data-command]')) {
  button.addEventListener('click', () => {
    window.panelApi.sendCommand({ type: button.dataset.command })
  })
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]))
}

// Strip icon-font glyphs (Private Use Area), zero-width characters, and the
// replacement char — scraped OWA text picks these up and they render as □.
function cleanText(value) {
  return String(value || '').replace(/[\u200B-\u200D\u2060\uFEFF\uFFFD\uE000-\uF8FF]/g, '').trim()
}

/* ---------- Compose bar ---------- */
for (const btn of document.querySelectorAll('[data-compose]')) {
  btn.addEventListener('click', () => {
    window.panelApi.sendCommand({ type: 'compose', kind: btn.dataset.compose })
  })
}

/* ---------- Refresh mail button ---------- */
const refreshMailBtn = document.getElementById('refresh-mail-btn')
if (refreshMailBtn) {
  refreshMailBtn.addEventListener('click', (e) => {
    e.stopPropagation()
    refreshMailBtn.classList.add('spinning')
    window.panelApi.sendCommand({ type: 'refresh-feeds' })
    setTimeout(() => refreshMailBtn.classList.remove('spinning'), 700)
  })
}

/* ---------- Summary strip → jump to mail ---------- */
summaryStrip.addEventListener('click', () => {
  const mail = (latest ? latest.services : []).find((s) => s.feed && s.feed.kind === 'mail')
  if (mail) window.panelApi.sendCommand({ type: 'switch-service', serviceKey: mail.key })
})

/* ---------- Connections: account rail + onboarding ---------- */
const MS_ICON =
  '<svg viewBox="0 0 24 24"><rect x="3" y="3" width="8.4" height="8.4"/><rect x="12.6" y="3" width="8.4" height="8.4"/><rect x="3" y="12.6" width="8.4" height="8.4"/><rect x="12.6" y="12.6" width="8.4" height="8.4"/></svg>'
const PROVIDER_META = {
  microsoft: { label: 'Microsoft', sub: 'Mail · Calendar', icon: MS_ICON, help: 'https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade' },
  asana: { label: 'Asana', sub: 'Tasks', icon: SERVICE_ICONS.asana, help: 'https://app.asana.com/0/my-apps' }
}
const STATUS_TEXT = {
  disconnected: 'Not connected',
  connecting: 'Connecting…',
  connected: 'Connected',
  error: 'Error'
}
const PROVIDERS = ['microsoft', 'asana']

const onboardOverlay = document.getElementById('onboard-overlay')

// Light "Set up notifications" nudge below the compose buttons. Shown only
// until the user finishes onboarding or skips it.
function renderNotifSetup(snapshot) {
  const el = document.getElementById('notif-setup')
  if (!el) return
  const show = !snapshot.onboarded && !snapshot.notifSetupSkipped
  el.hidden = !show
}

// Render the provider cards into a container (the onboarding sheet OR the
// settings Connections page). Inputs are scoped to each card by class — never
// by global id — so both containers can coexist without id collisions.
function renderConnCards(snapshot, container) {
  const wrap = container || document.getElementById('conn-cards')
  if (!wrap) return
  const conn = snapshot.connections || {}
  const cfg = snapshot.connConfig || {}
  wrap.innerHTML = ''
  for (const provider of PROVIDERS) {
    const st = conn[provider] || { status: 'disconnected', configured: false }
    const meta = PROVIDER_META[provider]
    const pcfg = cfg[provider] || {}
    const connected = st.status === 'connected'
    const connecting = st.status === 'connecting'
    const card = document.createElement('div')
    card.className = `conn-card conn-${st.status}`
    card.dataset.provider = provider

    const devFields =
      provider === 'microsoft'
        ? `<label class="conn-field"><span>Azure client ID</span><input type="text" class="cfg-clientId" spellcheck="false" placeholder="00000000-0000-0000-0000-000000000000" value="${escapeHtml(pcfg.clientId || '')}"></label>
           <label class="conn-field"><span>Tenant</span><input type="text" class="cfg-tenant" spellcheck="false" placeholder="common" value="${escapeHtml(pcfg.tenant || 'common')}"></label>`
        : `<label class="conn-field"><span>Asana client ID</span><input type="text" class="cfg-clientId" spellcheck="false" placeholder="1200000000000000" value="${escapeHtml(pcfg.clientId || '')}"></label>`

    card.innerHTML = `
      <div class="conn-head">
        <span class="conn-icon">${meta.icon}</span>
        <div class="conn-titles">
          <div class="conn-name">${meta.label}</div>
          <div class="conn-sub">${escapeHtml(connected && st.account && st.account.name ? st.account.name : meta.sub)}</div>
        </div>
        <span class="conn-state conn-state-${st.status}">${escapeHtml(STATUS_TEXT[st.status])}</span>
      </div>
      ${st.error ? `<div class="conn-error">${escapeHtml(st.error)}</div>` : ''}
      ${!st.configured ? `<div class="conn-hint">Add your ${meta.label} client ID below to enable Connect.</div>` : ''}
      <div class="conn-actions">
        <button class="conn-btn ${connected ? 'ghost' : 'primary'} conn-action" data-action="${connected ? 'disconnect' : 'connect'}" data-configured="${st.configured ? '1' : ''}" ${connecting ? 'disabled' : ''}>
          ${connected ? 'Disconnect' : connecting ? 'Connecting…' : 'Connect'}
        </button>
        <button class="conn-btn link conn-toggle">Developer setup</button>
      </div>
      <div class="conn-dev" ${st.configured && !devOpen.has(provider) ? 'hidden' : ''}>
        ${devFields}
        <div class="conn-dev-row">
          <button class="conn-btn small conn-save">Save</button>
          <a class="conn-help" href="#" data-ext="${meta.help}">Where do I get this?</a>
        </div>
      </div>
    `
    wrap.appendChild(card)
  }

  wrap.querySelectorAll('.conn-card').forEach((card) => {
    const provider = card.dataset.provider
    const action = card.querySelector('.conn-action')
    const dev = card.querySelector('.conn-dev')
    if (action) {
      action.addEventListener('click', () => {
        if (action.dataset.action === 'disconnect') {
          window.panelApi.sendCommand({ type: 'disconnect-provider', provider })
          return
        }
        // No client ID yet → reveal the setup fields rather than firing a
        // sign-in that can't start.
        if (action.dataset.configured !== '1') {
          if (dev) dev.hidden = false
          devOpen.add(provider)
          const input = card.querySelector('.cfg-clientId')
          if (input) input.focus()
          return
        }
        window.panelApi.sendCommand({ type: 'connect-provider', provider })
      })
    }
    const toggle = card.querySelector('.conn-toggle')
    if (toggle && dev) {
      toggle.addEventListener('click', () => {
        dev.hidden = !dev.hidden
        // Remember the open state so re-renders don't snap the panel shut.
        if (dev.hidden) devOpen.delete(provider)
        else devOpen.add(provider)
      })
    }
    const save = card.querySelector('.conn-save')
    if (save) {
      save.addEventListener('click', () => {
        devOpen.delete(provider)
        saveConnConfig(wrap)
      })
    }
    const help = card.querySelector('.conn-help')
    if (help) {
      help.addEventListener('click', (e) => {
        e.preventDefault()
        window.panelApi.sendCommand({ type: 'open-url', url: help.dataset.ext })
      })
    }
  })
}

// Collect both providers' client IDs from whichever container holds the cards.
function saveConnConfig(container) {
  const wrap = container || document.getElementById('conn-cards')
  const read = (provider, cls) => {
    const card = wrap && wrap.querySelector(`.conn-card[data-provider="${provider}"]`)
    const el = card ? card.querySelector(cls) : null
    return el ? el.value.trim() : ''
  }
  window.panelApi.sendCommand({
    type: 'save-connections',
    connections: {
      microsoft: { clientId: read('microsoft', '.cfg-clientId'), tenant: read('microsoft', '.cfg-tenant') || 'common' },
      asana: { clientId: read('asana', '.cfg-clientId') }
    }
  })
  // Release focus from the dev panel so the incoming snapshot re-renders the
  // card (the typing-guard in render() suppresses it while focus stays inside).
  if (document.activeElement && document.activeElement.blur) {
    document.activeElement.blur()
  }
}

function openOnboarding() {
  onboardOverlay.hidden = false
  // Ask main to detach the on-top service web view so the centered sheet is
  // visible and clickable (otherwise the BrowserView paints over it).
  window.panelApi.sendCommand({ type: 'open-onboarding' })
  if (latest) renderConnCards(latest)
  const warn = document.getElementById('onboard-encwarn')
  if (warn) warn.hidden = !(latest && latest.connections && latest.connections.encryptionAvailable === false)
}
function closeOnboarding() {
  onboardOverlay.hidden = true
  window.panelApi.sendCommand({ type: 'close-onboarding' })
}

// The "Set up notifications" nudge: main area opens onboarding, inline Skip
// dismisses it for good (persisted).
document.getElementById('notif-setup-main').addEventListener('click', openOnboarding)
document.getElementById('notif-setup-skip').addEventListener('click', (e) => {
  e.stopPropagation()
  document.getElementById('notif-setup').hidden = true
  window.panelApi.sendCommand({ type: 'skip-notif-setup' })
})

document.getElementById('onboard-close').addEventListener('click', closeOnboarding)
document.getElementById('onboard-later').addEventListener('click', closeOnboarding)
document.getElementById('onboard-finish').addEventListener('click', () => {
  window.panelApi.sendCommand({ type: 'finish-onboarding' })
  closeOnboarding()
})
onboardOverlay.addEventListener('mousedown', (e) => {
  if (e.target === onboardOverlay) closeOnboarding()
})

/* ---------- First-boot welcome ---------- */
// Primary path: plain web sign-in on the mail tab (finishes first boot).
// API onboarding is the advanced, optional route.
document.getElementById('welcome-signin').addEventListener('click', () => {
  window.panelApi.sendCommand({ type: 'start-signin' })
})
document.getElementById('welcome-advanced').addEventListener('click', openOnboarding)
document.getElementById('welcome-skip').addEventListener('click', () => {
  window.panelApi.sendCommand({ type: 'skip-signin' })
})

/* ---------- Find in page ---------- */
function isSearchOpen() {
  return searchBox.classList.contains('open')
}

function openSearch() {
  searchBox.classList.add('open')
  searchInput.tabIndex = 0
  document.getElementById('search-close').tabIndex = 0
  searchInput.focus()
  searchInput.select()
  if (searchInput.value) {
    window.panelApi.sendCommand({ type: 'find-in-page', text: searchInput.value })
  }
}

function closeSearch() {
  searchBox.classList.remove('open')
  searchInput.tabIndex = -1
  document.getElementById('search-close').tabIndex = -1
  searchCount.textContent = ''
  searchCount.classList.remove('no-results')
  window.panelApi.sendCommand({ type: 'stop-find-in-page' })
  searchInput.blur()
}

document.getElementById('search-toggle').addEventListener('click', () => {
  isSearchOpen() ? closeSearch() : openSearch()
})
document.getElementById('search-close').addEventListener('click', closeSearch)

searchInput.addEventListener('input', () => {
  window.panelApi.sendCommand({ type: 'find-in-page', text: searchInput.value })
})
searchInput.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') { closeSearch(); return }
  if (e.key === 'Enter') {
    e.preventDefault()
    window.panelApi.sendCommand({ type: 'find-in-page', text: searchInput.value, forward: !e.shiftKey, findNext: true })
  }
})
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && isSearchOpen()) closeSearch()
})

window.panelApi.onFindResult((data) => {
  if (data.matches === 0) {
    searchCount.textContent = searchInput.value ? 'No results' : ''
    searchCount.classList.toggle('no-results', !!searchInput.value)
  } else {
    searchCount.textContent = `${data.activeMatchOrdinal}/${data.matches}`
    searchCount.classList.remove('no-results')
  }
})

/* ---------- Focus timer ---------- */
const FOCUS_TOTAL = 25 * 60
const PLAY_ICON = '<path d="M8 5v14l11-7z"/>'
const PAUSE_ICON = '<path d="M8 5h3v14H8zM13 5h3v14h-3z"/>'
const focusTimeEl = document.getElementById('focus-time')
const focusLabelEl = document.getElementById('focus-label')
const focusProg = document.getElementById('focus-prog')
const focusToggleIcon = document.getElementById('focus-toggle-icon')
const RING_CIRC = 2 * Math.PI * 19
focusProg.style.strokeDasharray = String(RING_CIRC)
let focusRemaining = FOCUS_TOTAL
let focusRunning = false
let focusInterval = null
// Absolute deadline (ms epoch) while running — ticks derive the remaining time
// from it, so throttled timers in a hidden window can't stall the countdown.
let focusEndsAt = null

function paintFocus() {
  const m = Math.floor(focusRemaining / 60)
  const s = focusRemaining % 60
  focusTimeEl.textContent = `${m}:${String(s).padStart(2, '0')}`
  const frac = 1 - focusRemaining / FOCUS_TOTAL
  focusProg.style.strokeDashoffset = String(RING_CIRC * (1 - frac))
  focusToggleIcon.innerHTML = focusRunning ? PAUSE_ICON : PLAY_ICON
  focusLabelEl.textContent = focusRemaining === 0 ? 'Done!' : focusRunning ? 'Focusing' : focusRemaining === FOCUS_TOTAL ? 'Focus' : 'Paused'
  document.getElementById('focus').classList.toggle('running', focusRunning)
}

function tickFocus() {
  focusRemaining = Math.max(0, Math.round((focusEndsAt - Date.now()) / 1000))
  if (focusRemaining === 0) {
    focusRunning = false
    clearInterval(focusInterval)
    // Tell main so completion is noticeable even with the window hidden.
    window.panelApi.sendCommand({ type: 'focus-done' })
  }
  paintFocus()
}

document.getElementById('focus-toggle').addEventListener('click', () => {
  focusRunning = !focusRunning
  if (focusRunning) {
    if (focusRemaining === 0) focusRemaining = FOCUS_TOTAL
    focusEndsAt = Date.now() + focusRemaining * 1000
    focusInterval = setInterval(tickFocus, 1000)
  } else {
    clearInterval(focusInterval)
  }
  paintFocus()
})
document.getElementById('focus-reset').addEventListener('click', () => {
  focusRunning = false
  clearInterval(focusInterval)
  focusRemaining = FOCUS_TOTAL
  paintFocus()
})
paintFocus()

/* ---------- Tools drawer ---------- */
const toolsDrawer = document.getElementById('tools-drawer')
const toolsTab = document.getElementById('tools-tab')
if (toolsTab && toolsDrawer) {
  if (localStorage.getItem('orbit-tools-open')) toolsDrawer.classList.add('open')
  toolsTab.addEventListener('click', () => {
    const open = toolsDrawer.classList.toggle('open')
    localStorage.setItem('orbit-tools-open', open ? '1' : '')
  })
}

/* ---------- Scratch pad ---------- */
document.getElementById('scratch-toggle').addEventListener('click', () => {
  const open = scratchArea.hidden
  scratchArea.hidden = !open
  document.getElementById('scratch-chevron').style.transform = open ? 'rotate(180deg)' : ''
  if (open) scratchArea.focus()
})
scratchArea.addEventListener('input', () => {
  clearTimeout(scratchTimer)
  scratchTimer = setTimeout(() => {
    scratchTimer = null
    window.panelApi.sendCommand({ type: 'save-scratch', text: scratchArea.value })
  }, 500)
})
// Flush the pending debounce on blur — otherwise clicking away within 500ms
// lets a snapshot revert the textarea before the stale save fires.
scratchArea.addEventListener('blur', () => {
  if (scratchTimer === null) return
  clearTimeout(scratchTimer)
  scratchTimer = null
  window.panelApi.sendCommand({ type: 'save-scratch', text: scratchArea.value })
})

/* ---------- Main-process events ---------- */
window.panelApi.onEvent((data) => {
  if (data && data.type === 'open-search') {
    isSearchOpen() ? closeSearch() : openSearch()
  }
})

window.panelApi.onStatusUpdated(render)
window.panelApi.getSnapshot().then(render)

/* ---------- Network reconnect ---------- */
window.addEventListener('online', () => {
  window.panelApi.sendCommand({ type: 'network-online' })
})
