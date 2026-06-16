const card = document.getElementById('card')
const servicesEl = document.getElementById('services')
const statusIcon = document.getElementById('status-icon')
const statusNum = document.getElementById('status-num')
const statusText = document.getElementById('status-text')

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
  link: '<svg viewBox="0 0 24 24"><path d="M10 13a5 5 0 0 0 7 0l2-2a5 5 0 0 0-7-7l-1 1"/><path d="M14 11a5 5 0 0 0-7 0l-2 2a5 5 0 0 0 7 7l1-1"/></svg>'
}

const FALLBACK_ICON = SERVICE_ICONS.link
const MAIL_GLYPH = '<svg viewBox="0 0 24 24" class="icon"><rect x="3" y="5" width="18" height="14" rx="2.5"/><path d="M4 7l8 6 8-6"/></svg>'
const CHECK_GLYPH = '<svg viewBox="0 0 24 24" class="icon"><path d="M5 12.5l4.5 4.5L19 7"/></svg>'

// Keep the transparent window sized to its content (Asana-style auto-fit).
function syncHeight() {
  requestAnimationFrame(() => {
    const height = Math.ceil(card.getBoundingClientRect().height) + 20 // body padding (10 × 2)
    window.panelApi.sendCommand({ type: 'menu-resize', height })
  })
}

function renderStatus(mailService) {
  const mailServices = Array.isArray(mailService) ? mailService : (mailService ? [mailService] : [])
  const feedStates = mailServices.map((service) => service.feed && service.feed.state).filter(Boolean)
  const feed = mailServices.find((service) => service.feed)?.feed || null
  const state = feed ? feed.state : null
  const mailUnread = mailServices.reduce((total, service) => total + (Number(service.unreadCount) || 0), 0)
  if (mailUnread > 0) {
    statusIcon.className = 'm-status-icon unread'
    statusIcon.innerHTML = MAIL_GLYPH
    statusNum.textContent = mailUnread > 100 ? '100+' : mailUnread
    statusText.textContent = mailUnread === 1 ? 'unread email' : 'unread emails'
  } else if (feedStates.includes('login') || feedStates.includes('auth') || state === 'login' || state === 'auth') {
    // Signed out — "Inbox zero" here would be a false all-clear.
    statusIcon.className = 'm-status-icon'
    statusIcon.innerHTML = MAIL_GLYPH
    statusNum.textContent = 'Not signed in'
    statusText.textContent = 'Sign in to see your inbox'
  } else if (state === 'error') {
    statusIcon.className = 'm-status-icon'
    statusIcon.innerHTML = MAIL_GLYPH
    statusNum.textContent = 'Unavailable'
    statusText.textContent = "Couldn't read your inbox"
  } else if (state === 'ok' || state === 'empty') {
    statusIcon.className = 'm-status-icon'
    statusIcon.innerHTML = CHECK_GLYPH
    statusNum.textContent = 'Inbox zero'
    statusText.textContent = 'You are all caught up'
  } else {
    // 'loading' or no mail feed yet.
    statusIcon.className = 'm-status-icon'
    statusIcon.innerHTML = MAIL_GLYPH
    statusNum.textContent = '—'
    statusText.textContent = 'Loading…'
  }
}

function renderServices(snapshot) {
  servicesEl.innerHTML = ''

  for (const service of snapshot.services) {
    if (service.visible === false) continue
    const button = document.createElement('button')
    button.className = `m-service${service.key === snapshot.activeServiceKey ? ' active' : ''}`
    button.type = 'button'

    const hasUnread = service.unreadCount > 0
    const count = service.unreadCount > 100 ? '100+' : service.unreadCount

    button.innerHTML = `
      <span class="m-service-icon">${SERVICE_ICONS[service.icon] || FALLBACK_ICON}</span>
      <span class="m-service-name">${escapeHtml(service.label)}</span>
      <span class="m-service-badge${hasUnread ? '' : ' hidden'}">${count}</span>
    `

    button.addEventListener('click', () => {
      window.panelApi.sendCommand({ type: 'switch-service', serviceKey: service.key })
    })

    servicesEl.appendChild(button)
  }
}

function render(snapshot) {
  document.documentElement.setAttribute('data-theme', snapshot.theme === 'light' ? 'light' : 'dark')
  document.documentElement.toggleAttribute('data-glass', Boolean(snapshot.glassMode))

  const mail = snapshot.services.filter((service) => service.feed && service.feed.kind === 'mail')
  renderStatus(mail)
  renderServices(snapshot)
  syncHeight()
}

for (const button of document.querySelectorAll('[data-command]')) {
  button.addEventListener('click', () => {
    window.panelApi.sendCommand({ type: button.dataset.command })
  })
}

// Esc closes the dropdown, like a native menu.
window.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') {
    window.panelApi.sendCommand({ type: 'hide-menu' })
  }
})

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]))
}

window.panelApi.onStatusUpdated(render)
window.panelApi.getSnapshot().then(render)
