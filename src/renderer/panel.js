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
  copilot:
    '<svg viewBox="0 0 24 24"><path d="M12 3c.7 3.5 2 4.8 5.5 5.5C14 9.2 12.7 10.5 12 14c-.7-3.5-2-4.8-5.5-5.5C10 7.8 11.3 6.5 12 3z"/><path d="M18.5 14c.4 1.8 1 2.5 2.8 2.9-1.8.4-2.4 1-2.8 2.8-.4-1.8-1-2.4-2.8-2.8 1.8-.4 2.4-1.1 2.8-2.9z"/></svg>',
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
    note(feed.kind === 'mail' ? 'Inbox zero' : feed.kind === 'calendar' ? 'No upcoming events' : 'No tasks assigned')
    return wrap
  }

  for (const item of feed.items) {
    const row = document.createElement('div')
    row.className = 'feed-item'

    if (feed.kind === 'mail') {
      row.classList.add('feed-mail')
      const timeStr = item.receivedIso
        ? relativeTime(item.receivedIso)
        : item.today === true ? 'Today' : ''
      const preview = cleanText(item.preview || '')
      row.innerHTML = `
        <div class="feed-mail-header">
          <div class="feed-sender">${escapeHtml(cleanText(item.sender) || 'Unknown')}</div>
          ${timeStr ? `<div class="feed-time">${escapeHtml(timeStr)}</div>` : ''}
        </div>
        <div class="feed-subject">${escapeHtml(cleanText(item.subject) || '(no subject)')}</div>
        ${preview ? `<div class="feed-preview">${escapeHtml(preview)}</div>` : ''}
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
      if (item.cancelled) row.classList.add('feed-event-cancelled')
      row.innerHTML = `
        <div class="feed-event-title">${escapeHtml(cleanText(item.title) || 'Event')}</div>
        ${item.cancelled ? '<div class="feed-event-status">Cancelled</div>' : ''}
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
      const dueLabel = item.dueOn ? formatDueDate(item.dueOn) : ''
      const dueClass = dueLabel === 'Overdue' ? 'feed-due overdue' : dueLabel === 'Due today' ? 'feed-due due-today' : 'feed-due'
      row.innerHTML = `
        <div class="feed-task-name">${escapeHtml(cleanText(item.name) || 'Task')}</div>
        ${subs}
        ${dueLabel ? `<div class="${dueClass}">${escapeHtml(dueLabel)}</div>` : ''}
      `
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
const FEED_COLLAPSE_ICON =
  '<svg viewBox="0 0 24 24"><path d="M6 9l6 6 6-6"/></svg>'

function isSnoozableService(service) {
  return SNOOZABLE_KEYS.has(service.key) || Boolean(service.mailboxManaged)
}

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
    const feedCollapsed = Boolean(service.feedCollapsed)

    const row = document.createElement('div')
    row.className = 'service-row'

    const homeBtn = document.createElement('button')
    homeBtn.type = 'button'
    homeBtn.className = 'service-home'
    homeBtn.title = `Home — ${service.label}`
    homeBtn.innerHTML = `<span class="service-icon">${SERVICE_ICONS[service.icon] || SERVICE_ICONS.link}</span>`
    homeBtn.addEventListener('click', (event) => {
      event.stopPropagation()
      window.panelApi.sendCommand({ type: 'go-service-home', serviceKey: service.key })
    })
    row.appendChild(homeBtn)

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
    if (isSnoozableService(service)) button.title += ' · right-click for snooze & more'

    const hasUnread = service.unreadCount > 0
    const count = service.unreadCount
    const snoozeIcon = service.snoozed
      ? '<span class="service-snooze" title="Notifications snoozed"><svg viewBox="0 0 24 24"><path d="M6 5h6L6 12h6M14 11h5l-5 6h5"/></svg></span>'
      : ''
    const splitIcon = inSplit
      ? `<span class="service-split">${splitIndex === 0 ? SPLIT_LEFT_ICON : SPLIT_RIGHT_ICON}</span>`
      : ''
    button.innerHTML = `
      <span class="service-name">${escapeHtml(service.label)}</span>
      ${splitIcon}
      ${snoozeIcon}
    `
    button.addEventListener('click', (event) => {
      const type = event.metaKey || event.ctrlKey ? 'split-select' : 'switch-service'
      window.panelApi.sendCommand({ type, serviceKey: service.key })
    })
    button.addEventListener('contextmenu', (e) => {
      e.preventDefault()
      window.panelApi.sendCommand({ type: 'tab-context-menu', serviceKey: service.key })
    })
    row.appendChild(button)

    const badge = document.createElement('span')
    badge.className = `service-badge${hasUnread ? '' : ' hidden'}`
    badge.textContent = String(count)
    row.appendChild(badge)

    if (service.feed) {
      const collapseBtn = document.createElement('button')
      collapseBtn.type = 'button'
      collapseBtn.className = `feed-collapse${feedCollapsed ? ' collapsed' : ''}`
      collapseBtn.title = feedCollapsed ? 'Show notifications' : 'Hide notifications'
      collapseBtn.setAttribute('aria-expanded', feedCollapsed ? 'false' : 'true')
      collapseBtn.innerHTML = FEED_COLLAPSE_ICON
      collapseBtn.addEventListener('click', (event) => {
        event.stopPropagation()
        window.panelApi.sendCommand({ type: 'toggle-feed-collapse', serviceKey: service.key })
      })
      row.appendChild(collapseBtn)
    }

    group.appendChild(row)

    if (service.feed && !feedCollapsed) {
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
      e.dataTransfer.setData('application/x-mailstudio-row', String(index))
    })
    row.addEventListener('dragend', () => {
      dragIndex = null
      row.classList.remove('dragging')
      document.querySelectorAll('.set-item.drag-over').forEach((el) => el.classList.remove('drag-over'))
    })
    row.addEventListener('dragover', (e) => {
      if (!e.dataTransfer.types.includes('application/x-mailstudio-row')) return
      e.preventDefault()
      if (dragIndex !== null && dragIndex !== index) row.classList.add('drag-over')
    })
    row.addEventListener('dragleave', () => row.classList.remove('drag-over'))
    row.addEventListener('drop', (e) => {
      if (!e.dataTransfer.types.includes('application/x-mailstudio-row')) return
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

  // Apply the persisted sidebar width from settings. Don't clobber it while
  // the user is actively dragging the resize handle (sbDrag is live then).
  if (!sbDrag && typeof snapshot.sidebarExpandedWidth === 'number') {
    SB_EXPANDED = snapshot.sidebarExpandedWidth
    document.documentElement.style.setProperty('--sb-expanded', `${SB_EXPANDED}px`)
  }

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
let SB_EXPANDED = 280  // kept in sync with snapshot.sidebarExpandedWidth
const SB_RAIL = 76
const TOPBAR_H = 38
const SB_MIN = 180
const SB_MAX = 480

/* ---------- Sidebar resize ---------- */
const sbResizeHandle = document.getElementById('sidebar-resize-handle')
let sbDrag = null   // { startX, startWidth } while dragging
let sbLastSentWidth = null

function onSbMove(e) {
  if (!sbDrag) return
  const newWidth = Math.min(SB_MAX, Math.max(SB_MIN, sbDrag.startWidth + e.clientX - sbDrag.startX))
  // Update CSS variable immediately so the sidebar and handle move live.
  document.documentElement.style.setProperty('--sb-expanded', `${newWidth}px`)
  SB_EXPANDED = newWidth
  if (latest) latest.sidebarExpandedWidth = newWidth
  positionSplitDivider()
  // Throttle IPC: only send when width changed by ≥2px to avoid flooding main.
  if (sbLastSentWidth === null || Math.abs(newWidth - sbLastSentWidth) >= 2) {
    sbLastSentWidth = newWidth
    window.panelApi.sendCommand({ type: 'set-sidebar-width', width: newWidth })
  }
}

function onSbUp() {
  if (!sbDrag) return
  document.removeEventListener('mousemove', onSbMove, true)
  document.removeEventListener('mouseup', onSbUp, true)
  window.removeEventListener('blur', onSbUp)
  document.body.classList.remove('sb-resizing')
  const finalWidth = SB_EXPANDED
  window.panelApi.sendCommand({ type: 'set-sidebar-width', width: finalWidth, save: true })
  sbDrag = null
  sbLastSentWidth = null
}

if (sbResizeHandle) {
  sbResizeHandle.addEventListener('mousedown', (e) => {
    // Only drag when sidebar is actually expanded.
    if (document.body.classList.contains('collapsed') && !document.body.classList.contains('settings-open')) return
    e.preventDefault()
    sbDrag = { startX: e.clientX, startWidth: SB_EXPANDED }
    document.body.classList.add('sb-resizing')
    document.addEventListener('mousemove', onSbMove, true)
    document.addEventListener('mouseup', onSbUp, true)
    window.addEventListener('blur', onSbUp)
  })
}

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
  // Use the live SB_EXPANDED (updated from snapshot and during drag) so the
  // split divider handle always sits in the real gutter, not at 280px.
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

// Relative time label from a received-at ISO timestamp.
function relativeTime(isoString) {
  if (!isoString) return ''
  const d = new Date(isoString)
  if (Number.isNaN(d.getTime())) return ''
  const diff = Date.now() - d.getTime()
  if (diff < 60000) return 'Just now'
  if (diff < 3600000) return `${Math.floor(diff / 60000)}m ago`
  if (diff < 86400000) return `${Math.floor(diff / 3600000)}h ago`
  const today = new Date()
  const yesterday = new Date(today)
  yesterday.setDate(today.getDate() - 1)
  if (d.toDateString() === today.toDateString()) return 'Today'
  if (d.toDateString() === yesterday.toDateString()) return 'Yesterday'
  return d.toLocaleDateString([], { weekday: 'short', month: 'short', day: 'numeric' })
}

// Human-friendly due-date label for Asana task items.
function formatDueDate(dateStr) {
  if (!dateStr) return ''
  const d = new Date(`${dateStr}T00:00:00`)
  if (Number.isNaN(d.getTime())) return ''
  const today = new Date()
  today.setHours(0, 0, 0, 0)
  const tomorrow = new Date(today)
  tomorrow.setDate(today.getDate() + 1)
  if (d.getTime() < today.getTime()) return 'Overdue'
  if (d.toDateString() === today.toDateString()) return 'Due today'
  if (d.toDateString() === tomorrow.toDateString()) return 'Due tomorrow'
  return `Due ${d.toLocaleDateString([], { month: 'short', day: 'numeric' })}`
}

/* ---------- Downloads ---------- */
const dlDrawer = document.getElementById('dl-drawer')
const dlList = document.getElementById('dl-list')
const dlEmpty = document.getElementById('dl-empty')
const dlToggleBtn = document.getElementById('dl-toggle-btn')
const dlBadge = document.getElementById('dl-badge')

// Classify a filename to drive icon colour and the file-type icon glyph.
function dlFileKind(filename) {
  const ext = (filename.split('.').pop() || '').toLowerCase()
  if (/^(jpg|jpeg|png|gif|webp|svg|heic|avif|bmp|tiff?)$/.test(ext)) return 'image'
  if (/^(mp4|mov|mkv|avi|webm|m4v|flv|wmv)$/.test(ext)) return 'video'
  if (/^(zip|tar|gz|bz2|xz|7z|rar|dmg|iso|pkg)$/.test(ext)) return 'archive'
  if (/^(pdf|doc|docx|xls|xlsx|ppt|pptx|odt|ods|odp|pages|numbers|key|txt|md|csv)$/.test(ext)) return 'doc'
  return 'other'
}

// Icon SVG paths by kind.
const DL_KIND_ICON = {
  image: '<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="M21 15l-5-5L5 21"/>',
  video: '<polygon points="5 3 19 12 5 21 5 3"/>',
  archive: '<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>',
  doc: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>',
  other: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>'
}

function formatBytes(bytes) {
  if (!bytes || bytes <= 0) return '0 B'
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`
  if (bytes < 1073741824) return `${(bytes / 1048576).toFixed(1)} MB`
  return `${(bytes / 1073741824).toFixed(2)} GB`
}

function renderDownloads({ list, activeCount }) {
  // Show/hide button and badge
  if (!dlToggleBtn) return
  if (list.length === 0) {
    dlToggleBtn.hidden = true
    if (dlBadge) dlBadge.hidden = true
    return
  }
  dlToggleBtn.hidden = false
  if (dlBadge) {
    if (activeCount > 0) {
      dlBadge.textContent = String(activeCount)
      dlBadge.hidden = false
    } else {
      dlBadge.hidden = true
    }
  }

  if (!dlList) return
  dlList.innerHTML = ''
  dlEmpty.hidden = list.length > 0

  for (const dl of list) {
    const kind = dlFileKind(dl.filename)
    const ext = dl.filename.includes('.') ? dl.filename.split('.').pop().toLowerCase().slice(0, 4) : ''
    const progress = dl.totalBytes > 0 ? dl.receivedBytes / dl.totalBytes : 0
    const pct = Math.min(100, Math.round(progress * 100))

    let metaText
    if (dl.state === 'progressing') {
      const received = formatBytes(dl.receivedBytes)
      const total = dl.totalBytes > 0 ? ` of ${formatBytes(dl.totalBytes)}` : ''
      const speed = dl.speed > 0 ? ` · ${formatBytes(dl.speed)}/s` : ''
      metaText = `${received}${total}${speed}`
    } else if (dl.state === 'completed') {
      metaText = `${formatBytes(dl.totalBytes)} · Done`
    } else if (dl.state === 'cancelled') {
      metaText = 'Cancelled'
    } else {
      metaText = 'Failed'
    }

    const item = document.createElement('div')
    item.className = `dl-item dl-${dl.state}`
    item.dataset.kind = kind

    item.innerHTML = `
      <div class="dl-file-icon">
        <svg viewBox="0 0 24 24">${DL_KIND_ICON[kind] || DL_KIND_ICON.other}</svg>
        ${ext ? `<span class="dl-ext">${escapeHtml(ext)}</span>` : ''}
      </div>
      <div class="dl-details">
        <div class="dl-name" title="${escapeHtml(dl.filename)}">${escapeHtml(dl.filename)}</div>
        <div class="dl-meta">${escapeHtml(metaText)}</div>
        ${dl.state === 'progressing'
          ? `<div class="dl-progress-wrap"><div class="dl-progress-fill" style="width:${pct}%"></div></div>`
          : ''}
      </div>
      <div class="dl-actions">
        ${dl.state === 'completed' && dl.savePath ? `
          <button class="dl-act" data-action="open" data-id="${dl.id}" title="Open file">
            <svg viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
          </button>
          <button class="dl-act" data-action="show" data-id="${dl.id}" title="Show in Finder">
            <svg viewBox="0 0 24 24"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
          </button>` : ''}
        ${dl.state === 'progressing' ? `
          <button class="dl-act danger" data-action="cancel" data-id="${dl.id}" title="Cancel">
            <svg viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>` : ''}
      </div>
    `
    dlList.appendChild(item)
  }

  // Wire action buttons
  for (const btn of dlList.querySelectorAll('.dl-act')) {
    btn.addEventListener('click', (e) => {
      e.stopPropagation()
      window.panelApi.sendCommand({ type: `download-${btn.dataset.action}`, id: Number(btn.dataset.id) })
    })
  }
}

// Toggle drawer open/close
if (dlToggleBtn) {
  dlToggleBtn.addEventListener('click', (e) => {
    e.stopPropagation()
    dlDrawer.hidden = !dlDrawer.hidden
  })
}
if (document.getElementById('dl-drawer-close')) {
  document.getElementById('dl-drawer-close').addEventListener('click', () => {
    dlDrawer.hidden = true
  })
}
if (document.getElementById('dl-clear-btn')) {
  document.getElementById('dl-clear-btn').addEventListener('click', () => {
    window.panelApi.sendCommand({ type: 'download-clear' })
  })
}
// Click outside the drawer to dismiss it
document.addEventListener('click', (e) => {
  if (dlDrawer && !dlDrawer.hidden && !dlDrawer.contains(e.target) && e.target !== dlToggleBtn) {
    dlDrawer.hidden = true
  }
})

// Auto-open drawer when a new download starts
window.panelApi.onEvent((data) => {
  if (data && data.type === 'open-search') {
    isSearchOpen() ? closeSearch() : openSearch()
  }
  if (data && data.type === 'download-started') {
    if (dlDrawer) dlDrawer.hidden = false
  }
})

window.panelApi.onDownloadsUpdated((data) => {
  renderDownloads(data)
})

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
  microsoft: {
    label: 'Microsoft',
    sub: 'Mail · Calendar',
    icon: MS_ICON,
    helpClientId: 'https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade',
    helpTenant: 'https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/Overview'
  },
  asana: {
    label: 'Asana',
    sub: 'Tasks',
    icon: SERVICE_ICONS.asana,
    helpClientId: 'https://app.asana.com/0/my-apps'
  }
}
const README_SETUP_URL = 'https://github.com/MailStudio/MailStudio#api-setup--byo-credentials'
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
           <a class="conn-help conn-field-help" href="#" data-ext="${meta.helpClientId}">Where do I get the client ID?</a>
           <label class="conn-field"><span>Tenant ID</span><input type="text" class="cfg-tenant" spellcheck="false" placeholder="common" value="${escapeHtml(pcfg.tenant || 'common')}"></label>
           <p class="conn-field-hint">Leave as <code>common</code> for personal and work/school accounts. If sign-in fails, paste your Directory (tenant) ID from Azure → App registration → Overview.</p>
           <a class="conn-help conn-field-help" href="#" data-ext="${meta.helpTenant}">Find my tenant ID</a>`
        : `<label class="conn-field"><span>Asana client ID</span><input type="text" class="cfg-clientId" spellcheck="false" placeholder="1200000000000000" value="${escapeHtml(pcfg.clientId || '')}"></label>
           <a class="conn-help conn-field-help" href="#" data-ext="${meta.helpClientId}">Where do I get the client ID?</a>
           <label class="conn-field"><span>Asana client secret</span><input type="password" class="cfg-clientSecret" spellcheck="false" autocomplete="off" placeholder="${st.secretSet ? '•••••••• saved — leave blank to keep' : "Paste your app's client secret"}" value=""></label>
           <p class="conn-field-hint">Asana requires the app's <strong>client secret</strong> too (Microsoft doesn't). Find it in your Asana app under <em>Basic information</em>. It's sealed in your OS keychain — never stored in plaintext.</p>`

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
      ${!st.configured ? `<div class="conn-hint">Add your ${meta.label} ${provider === 'asana' ? 'client ID and secret' : 'client ID'} below to enable Connect.</div>` : ''}
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
        // Persist any unsaved client ID / tenant edits before starting OAuth.
        const container = card.closest('.conn-cards')
        window.panelApi.sendCommand({
          type: 'save-connections',
          connections: readConnConfig(container)
        })
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
    card.querySelectorAll('.conn-help').forEach((help) => {
      help.addEventListener('click', (e) => {
        e.preventDefault()
        window.panelApi.sendCommand({ type: 'open-url', url: help.dataset.ext, external: true })
      })
    })
  })
}

// Collect both providers' client IDs from whichever container holds the cards.
function readConnConfig(container) {
  const wrap = container || document.getElementById('conn-cards')
  const read = (provider, cls) => {
    const card = wrap && wrap.querySelector(`.conn-card[data-provider="${provider}"]`)
    const el = card ? card.querySelector(cls) : null
    return el ? el.value.trim() : ''
  }
  return {
    microsoft: { clientId: read('microsoft', '.cfg-clientId'), tenant: read('microsoft', '.cfg-tenant') || 'common' },
    // clientSecret is blank unless the user typed a new one; main.js treats an
    // empty value as "keep the existing stored secret" (never wiped on re-save).
    asana: { clientId: read('asana', '.cfg-clientId'), clientSecret: read('asana', '.cfg-clientSecret') }
  }
}

function saveConnConfig(container) {
  window.panelApi.sendCommand({
    type: 'save-connections',
    connections: readConnConfig(container)
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
const onboardReadme = document.getElementById('onboard-readme')
if (onboardReadme) {
  onboardReadme.addEventListener('click', (e) => {
    e.preventDefault()
    window.panelApi.sendCommand({ type: 'open-url', url: README_SETUP_URL, external: true })
  })
}
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

/* ---------- Focus complete particle burst ---------- */
const BURST_CANVAS = document.getElementById('burst-canvas')
const BURST_COLORS = [
  '#FF6B6B', '#FF8C42', '#FFC300', '#ADFF2F',
  '#2FD88A', '#00CFFF', '#6C63FF', '#D66FFF',
  '#FF6CAD', '#FF9A3C', '#57E8B0', '#74C0FC',
  '#FFA8C5', '#A9FF68', '#FFDE59', '#5CE1E6'
]

let burstParticles = []
let burstRaf = null

function makeBurstParticle(W, H) {
  // Spawn on a random edge, fire inward with spread.
  const edge = (Math.random() * 4) | 0
  let x, y, vx, vy
  const spd = 2.5 + Math.random() * 7
  const spread = (Math.random() - 0.5) * Math.PI * 0.55

  if (edge === 0) {        // top
    x = Math.random() * W; y = 0
    vx = Math.sin(spread) * spd; vy = Math.cos(spread) * spd
  } else if (edge === 1) { // right
    x = W; y = Math.random() * H
    vx = -Math.cos(spread) * spd; vy = Math.sin(spread) * spd
  } else if (edge === 2) { // bottom
    x = Math.random() * W; y = H
    vx = Math.sin(spread) * spd; vy = -Math.cos(spread) * spd
  } else {                 // left
    x = 0; y = Math.random() * H
    vx = Math.cos(spread) * spd; vy = Math.sin(spread) * spd
  }

  const isRect = Math.random() < 0.55
  return {
    x, y, vx, vy,
    // rect fields
    w: 5 + Math.random() * 8,
    h: 3 + Math.random() * 4,
    rot: Math.random() * Math.PI * 2,
    rotSpd: (Math.random() - 0.5) * 0.35,
    // circle fields
    r: 2.5 + Math.random() * 4,
    // common
    shape: isRect ? 'rect' : 'dot',
    color: BURST_COLORS[(Math.random() * BURST_COLORS.length) | 0],
    life: 0.85 + Math.random() * 0.15,
    decay: 0.007 + Math.random() * 0.01,
    grav: 0.04 + Math.random() * 0.09
  }
}

function drawBurstParticle(ctx, p) {
  const a = p.life * p.life // quadratic — fast start, gentle tail-off
  ctx.globalAlpha = a
  ctx.fillStyle = p.color
  if (p.shape === 'rect') {
    ctx.save()
    ctx.translate(p.x, p.y)
    ctx.rotate(p.rot)
    const h = p.h * (0.4 + 0.6 * p.life)
    ctx.fillRect(-p.w * 0.5, -h * 0.5, p.w, h)
    ctx.restore()
  } else {
    const r = p.r * (0.5 + 0.5 * p.life)
    ctx.beginPath()
    ctx.arc(p.x, p.y, r, 0, 6.283185)
    ctx.fill()
  }
}

function animateBurst() {
  if (!BURST_CANVAS) return
  const ctx = BURST_CANVAS.getContext('2d')
  const W = BURST_CANVAS.width
  const H = BURST_CANVAS.height
  ctx.clearRect(0, 0, W, H)

  let alive = false
  for (const p of burstParticles) {
    if (p.life <= 0) continue
    alive = true
    p.x += p.vx; p.y += p.vy
    p.vy += p.grav; p.vx *= 0.992
    p.rot += p.rotSpd
    p.life -= p.decay
    drawBurstParticle(ctx, p)
  }
  ctx.globalAlpha = 1

  if (alive) {
    burstRaf = requestAnimationFrame(animateBurst)
  } else {
    BURST_CANVAS.classList.remove('active')
    burstParticles = []
    burstRaf = null
  }
}

function spawnBurst() {
  if (!BURST_CANVAS) return
  const W = window.innerWidth
  const H = window.innerHeight
  BURST_CANVAS.width = W
  BURST_CANVAS.height = H
  BURST_CANVAS.classList.add('active')

  burstParticles = []
  // 480 particles distributed along all four edges for a dense rim burst.
  for (let i = 0; i < 480; i++) burstParticles.push(makeBurstParticle(W, H))

  if (burstRaf) cancelAnimationFrame(burstRaf)
  burstRaf = requestAnimationFrame(animateBurst)
}

/* ---------- Focus timer ---------- */
const PLAY_ICON = '<path d="M8 5v14l11-7z"/>'
const PAUSE_ICON = '<path d="M8 5h3v14H8zM13 5h3v14h-3z"/>'
const focusTimeEl = document.getElementById('focus-time')
const focusLabelEl = document.getElementById('focus-label')
const focusProg = document.getElementById('focus-prog')
const focusToggleIcon = document.getElementById('focus-toggle-icon')
const RING_CIRC = 2 * Math.PI * 19
focusProg.style.strokeDasharray = String(RING_CIRC)
const FOCUS_MIN_MINUTES = 1
const FOCUS_MAX_MINUTES = 240
let focusTotal = 30 * 60
let focusRemaining = focusTotal
let focusRunning = false
let focusInterval = null
// Absolute deadline (ms epoch) while running — ticks derive the remaining time
// from it, so throttled timers in a hidden window can't stall the countdown.
let focusEndsAt = null

function paintFocus() {
  const m = Math.floor(focusRemaining / 60)
  const s = focusRemaining % 60
  focusTimeEl.textContent = `${m}:${String(s).padStart(2, '0')}`
  const frac = 1 - focusRemaining / focusTotal
  focusProg.style.strokeDashoffset = String(RING_CIRC * (1 - frac))
  focusToggleIcon.innerHTML = focusRunning ? PAUSE_ICON : PLAY_ICON
  focusLabelEl.textContent = focusRemaining === 0 ? 'Done!' : focusRunning ? 'Focusing' : focusRemaining === focusTotal ? 'Focus' : 'Paused'
  document.getElementById('focus').classList.toggle('running', focusRunning)
  // Keep the duration input in sync with the current session length (unless the
  // user is mid-edit). It's the only duration control, so the row always fits
  // any sidebar width. Disabled while running so the countdown isn't edited live.
  const customEl = document.getElementById('focus-custom')
  if (customEl) {
    if (document.activeElement !== customEl) customEl.value = String(Math.round(focusTotal / 60))
    customEl.disabled = focusRunning
  }
  for (const id of ['focus-minus', 'focus-plus']) {
    const b = document.getElementById(id)
    if (b) b.disabled = focusRunning
  }
}

function tickFocus() {
  focusRemaining = Math.max(0, Math.round((focusEndsAt - Date.now()) / 1000))
  if (focusRemaining === 0) {
    focusRunning = false
    clearInterval(focusInterval)
    window.panelApi.sendCommand({ type: 'focus-done', minutes: Math.round(focusTotal / 60) })
    // 🎉 Colorful edge particle burst — fires from the window border inward.
    spawnBurst()
  }
  paintFocus()
}

document.getElementById('focus-toggle').addEventListener('click', () => {
  focusRunning = !focusRunning
  if (focusRunning) {
    if (focusRemaining === 0) focusRemaining = focusTotal
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
  focusRemaining = focusTotal
  paintFocus()
})
// Set the session length to a given minute count (clamped), resetting the timer.
function setFocusMinutes(mins) {
  const clamped = Math.max(FOCUS_MIN_MINUTES, Math.min(FOCUS_MAX_MINUTES, Math.round(mins)))
  focusRunning = false
  clearInterval(focusInterval)
  focusTotal = clamped * 60
  focusRemaining = focusTotal
  paintFocus()
}

// Custom duration input — type minutes and press Enter or tab away. This is the
// only duration control, so it fits the sidebar at every drag width.
const focusCustomEl = document.getElementById('focus-custom')
if (focusCustomEl) {
  const applyCustom = () => {
    const mins = parseInt(focusCustomEl.value, 10)
    if (!mins) { paintFocus(); return } // empty/invalid → restore current value
    setFocusMinutes(mins)
  }
  focusCustomEl.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); applyCustom(); focusCustomEl.blur() }
    if (e.key === 'Escape') { focusCustomEl.blur(); paintFocus() }
  })
  focusCustomEl.addEventListener('change', applyCustom)
}

// −/+ steppers adjust by 5 minutes (snapping to the nearest multiple of 5).
const focusMinus = document.getElementById('focus-minus')
const focusPlus = document.getElementById('focus-plus')
if (focusMinus) focusMinus.addEventListener('click', () => {
  const cur = Math.round(focusTotal / 60)
  setFocusMinutes(Math.floor((cur - 1) / 5) * 5 || FOCUS_MIN_MINUTES)
})
if (focusPlus) focusPlus.addEventListener('click', () => {
  const cur = Math.round(focusTotal / 60)
  setFocusMinutes((Math.floor(cur / 5) + 1) * 5)
})

paintFocus()

/* ---------- Tools drawer ---------- */
const toolsDrawer = document.getElementById('tools-drawer')
const toolsTab = document.getElementById('tools-tab')
if (toolsTab && toolsDrawer) {
  if (localStorage.getItem('mailstudio-tools-open')) toolsDrawer.classList.add('open')
  toolsTab.addEventListener('click', () => {
    const open = toolsDrawer.classList.toggle('open')
    localStorage.setItem('mailstudio-tools-open', open ? '1' : '')
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
window.panelApi.onStatusUpdated(render)
window.panelApi.getSnapshot().then(render)

/* ---------- Network reconnect ---------- */
window.addEventListener('online', () => {
  window.panelApi.sendCommand({ type: 'network-online' })
})
