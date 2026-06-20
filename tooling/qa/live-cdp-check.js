'use strict'

const assert = require('node:assert/strict')

const DEBUG_BASE = process.env.MAILSTUDIO_CDP_URL || 'http://127.0.0.1:9333'
let idSeq = 0

async function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

async function connect(wsUrl) {
  const ws = new WebSocket(wsUrl)
  const pending = new Map()
  const events = []

  ws.addEventListener('message', (event) => {
    const msg = JSON.parse(event.data)
    if (msg.id && pending.has(msg.id)) {
      const { resolve, reject, timer } = pending.get(msg.id)
      clearTimeout(timer)
      pending.delete(msg.id)
      if (msg.error) reject(new Error(msg.error.message || JSON.stringify(msg.error)))
      else resolve(msg.result || {})
      return
    }
    events.push(msg)
  })

  await new Promise((resolve, reject) => {
    ws.addEventListener('open', resolve, { once: true })
    ws.addEventListener('error', reject, { once: true })
  })

  return {
    events,
    send(method, params) {
      const id = ++idSeq
      ws.send(JSON.stringify({ id, method, params: params || {} }))
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
          if (pending.delete(id)) reject(new Error(`CDP timeout: ${method}`))
        }, 10000)
        pending.set(id, { resolve, reject, timer })
      })
    },
    close() {
      ws.close()
    }
  }
}

async function evaluate(client, expression) {
  const out = await client.send('Runtime.evaluate', {
    expression,
    awaitPromise: true,
    returnByValue: true,
    timeout: 10000
  })
  if (out.exceptionDetails) {
    throw new Error(out.exceptionDetails.text || 'Runtime evaluation failed')
  }
  return out.result.value
}

async function targetHealth(target) {
  const client = await connect(target.webSocketDebuggerUrl)
  try {
    await client.send('Runtime.enable')
    await client.send('Log.enable').catch(() => {})
    await sleep(150)
    const health = await evaluate(client, `(() => ({
      title: document.title || '',
      href: location.href,
      readyState: document.readyState,
      bodyTextLength: document.body ? document.body.innerText.length : 0,
      buttonCount: document.querySelectorAll('button').length,
      inputCount: document.querySelectorAll('input, textarea, [contenteditable="true"]').length,
      hasLoginInput: Boolean(document.querySelector('input[name="loginfmt"], input[type="password"]')),
      blankLike: !document.body || document.body.innerText.trim().length === 0,
      errorEvents: 0
    }))()`)
    health.errorEvents = client.events.filter((event) =>
      event.method === 'Runtime.exceptionThrown' ||
      (event.method === 'Log.entryAdded' && ['error', 'warning'].includes(event.params?.entry?.level))
    ).length
    return health
  } finally {
    client.close()
  }
}

function hostOf(url) {
  try { return new URL(url).hostname } catch { return '' }
}

async function main() {
  const targets = await fetch(`${DEBUG_BASE}/json/list`).then((res) => res.json())
  const pages = targets.filter((target) => target.type === 'page')
  const panel = pages.find((target) => target.url.endsWith('/src/renderer/panel.html'))
  const menu = pages.find((target) => target.url.endsWith('/src/renderer/menu.html'))
  const servicePages = pages.filter((target) => /^https?:/.test(target.url))
  const outlookPages = servicePages.filter((target) => hostOf(target.url).includes('outlook.'))
  const asanaPages = servicePages.filter((target) => hostOf(target.url).endsWith('asana.com'))
  const teamsPages = servicePages.filter((target) => hostOf(target.url).includes('teams.'))

  assert.ok(panel, 'MailStudio panel target not found')
  assert.ok(menu, 'MailStudio menu target not found')
  assert.ok(outlookPages.length > 0, 'No Outlook service target found')

  const panelClient = await connect(panel.webSocketDebuggerUrl)
  try {
    await panelClient.send('Runtime.enable')
    await panelClient.send('Log.enable').catch(() => {})
    const snapshot = await evaluate(panelClient, 'window.panelApi.getSnapshot()')
    assert.equal(snapshot.appName, 'MailStudio')
    assert.ok(Array.isArray(snapshot.services), 'snapshot.services missing')
    assert.ok(snapshot.services.some((service) => service.key === 'asana'), 'asana service missing from snapshot')

    const visible = snapshot.services.filter((service) => service.visible)
    const feedSummary = visible
      .filter((service) => service.feed)
      .map((service) => ({
        key: service.key,
        label: service.label,
        kind: service.feed.kind,
        state: service.feed.state,
        items: Array.isArray(service.feed.items) ? service.feed.items.length : 0,
        unread: Number(service.unreadCount) || 0
      }))
    const visibleKeys = visible.map((service) => service.key)
    const mailServices = snapshot.services.filter((service) => service.feed?.kind === 'mail')
    const mailUnreadTotal = mailServices.reduce((total, service) => total + (Number(service.unreadCount) || 0), 0)

    const panelBefore = await evaluate(panelClient, `(() => ({
      bodyClass: document.body.className,
      serviceRows: document.querySelectorAll('.service-row').length,
      feedRows: document.querySelectorAll('.feed-item').length,
      settingsOpen: document.body.classList.contains('settings-open'),
      firstBoot: document.body.classList.contains('first-boot'),
      theme: document.documentElement.getAttribute('data-theme'),
      splitDividerHidden: document.getElementById('split-divider')?.hidden ?? null,
      downloadDrawerHidden: document.getElementById('dl-drawer')?.hidden ?? null
    }))()`)

    await evaluate(panelClient, "window.panelApi.sendCommand({ type: 'refresh-feeds' }); true")
    await sleep(1500)
    const refreshed = await evaluate(panelClient, 'window.panelApi.getSnapshot()')
    const refreshedFeeds = refreshed.services
      .filter((service) => service.visible && service.feed)
      .map((service) => ({ key: service.key, state: service.feed.state, items: service.feed.items.length }))

    await evaluate(panelClient, "window.panelApi.sendCommand({ type: 'open-settings' }); true")
    await sleep(400)
    const settingsOpened = await evaluate(panelClient, "document.body.classList.contains('settings-open')")
    await evaluate(panelClient, "window.panelApi.sendCommand({ type: 'close-settings' }); true")
    await sleep(400)
    const settingsClosed = await evaluate(panelClient, "!document.body.classList.contains('settings-open')")
    assert.equal(settingsOpened, true, 'settings did not open')
    assert.equal(settingsClosed, true, 'settings did not close')

    const menuHealth = await targetHealth(menu)
    const serviceHealth = []
    for (const target of [...outlookPages, ...asanaPages, ...teamsPages].slice(0, 8)) {
      const health = await targetHealth(target)
      serviceHealth.push({
        host: hostOf(target.url),
        title: health.title,
        readyState: health.readyState,
        bodyTextLength: health.bodyTextLength,
        buttonCount: health.buttonCount,
        inputCount: health.inputCount,
        hasLoginInput: health.hasLoginInput,
        blankLike: health.blankLike,
        errorEvents: health.errorEvents
      })
    }

    const issues = []
    const visibleFeedErrors = feedSummary.filter((feed) => feed.state === 'error')
    if (visibleFeedErrors.length) issues.push(`feeds in error state: ${visibleFeedErrors.map((f) => f.key).join(', ')}`)
    if (panelBefore.serviceRows < visible.length) issues.push('panel rendered fewer service rows than visible services')
    if (panelBefore.firstBoot) issues.push('first-boot overlay still active after login')
    const blankServices = serviceHealth.filter((service) =>
      service.blankLike &&
      !service.hasLoginInput &&
      service.host !== 'teams.microsoft.com'
    )
    if (blankServices.length) issues.push(`blank service pages: ${blankServices.map((s) => s.host).join(', ')}`)

    const report = {
      targetCount: targets.length,
      pageCount: pages.length,
      servicePageHosts: servicePages.map((target) => hostOf(target.url)).filter(Boolean),
      visibleKeys,
      activeServiceKey: snapshot.activeServiceKey,
      connections: {
        microsoft: snapshot.connections?.microsoft?.status,
        asana: snapshot.connections?.asana?.status,
        encryptionAvailable: snapshot.connections?.encryptionAvailable
      },
      feedSummary,
      refreshedFeeds,
      mailUnreadTotal,
      panelBefore,
      menuHealth,
      serviceHealth,
      issues
    }

    console.log(JSON.stringify(report, null, 2))
    if (issues.length) process.exitCode = 2
  } finally {
    panelClient.close()
  }
}

main().catch((err) => {
  console.error(err && err.stack ? err.stack : err)
  process.exit(1)
})
