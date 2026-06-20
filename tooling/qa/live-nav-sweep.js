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
        }, 12000)
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
    timeout: 12000
  })
  if (out.exceptionDetails) {
    throw new Error(out.exceptionDetails.text || 'Runtime evaluation failed')
  }
  return out.result.value
}

function hostOf(raw) {
  try { return new URL(raw).hostname } catch { return '' }
}

function normalizeHost(raw) {
  const host = hostOf(raw).replace(/^www\./, '')
  if (host === 'outlook.office.com') return 'outlook.cloud.microsoft'
  if (host === 'office.com') return 'm365.cloud.microsoft'
  return host
}

function pageScore(service, target) {
  const targetHost = normalizeHost(target.url)
  const targetUrl = target.url || ''
  const serviceHost = normalizeHost(service.href || service.url || '')
  const serviceUrl = service.href || service.url || ''
  let score = 0

  if (targetHost && targetHost === serviceHost) score += 3
  if (service.feedKind === 'mail' && targetHost.includes('outlook.') && targetUrl.includes('/mail/')) score += 3
  if (service.feedKind === 'calendar' && targetHost.includes('outlook.') && targetUrl.includes('/calendar')) score += 3
  if (service.key === 'teams' && targetHost.includes('teams.')) score += 4
  if (service.key === 'asana' && targetHost.endsWith('asana.com')) score += 4
  if (service.key === 'sharepoint' && targetHost.includes('sharepoint.com') && !targetHost.includes('-my.')) score += 4
  if (service.key === 'onedrive' && targetHost.includes('-my.sharepoint.com')) score += 4
  if (service.key === 'office' && targetHost.includes('m365.cloud.microsoft')) score += 4
  if (service.key === 'word' && targetHost.includes('word.cloud.microsoft')) score += 4
  if (service.key === 'excel' && targetHost.includes('excel.cloud.microsoft')) score += 4
  if (service.key === 'powerpoint' && targetHost.includes('powerpoint.cloud.microsoft')) score += 4
  if (service.key === 'onenote' && targetUrl.includes('onenote')) score += 4
  if (service.key === 'planner' && targetHost.includes('planner.cloud.microsoft')) score += 4
  if (serviceUrl && targetUrl.startsWith(serviceUrl)) score += 2

  return score
}

async function pageHealth(target) {
  const client = await connect(target.webSocketDebuggerUrl)
  try {
    await client.send('Runtime.enable')
    await client.send('Log.enable').catch(() => {})
    await sleep(200)
    const health = await evaluate(client, `(() => ({
      title: document.title || '',
      href: location.href,
      readyState: document.readyState,
      bodyTextLength: document.body ? document.body.innerText.trim().length : 0,
      buttonCount: document.querySelectorAll('button').length,
      inputCount: document.querySelectorAll('input, textarea, [contenteditable="true"]').length,
      hasLoginInput: Boolean(document.querySelector('input[name="loginfmt"], input[type="password"]')),
      hasAppShell: Boolean(document.body && document.body.children.length),
      blankLike: !document.body || document.body.innerText.trim().length === 0
    }))()`)
    health.errorEvents = client.events.filter((event) =>
      event.method === 'Runtime.exceptionThrown' ||
      (event.method === 'Log.entryAdded' && event.params?.entry?.level === 'error')
    ).length
    return health
  } finally {
    client.close()
  }
}

async function targets() {
  return fetch(`${DEBUG_BASE}/json/list`).then((res) => res.json())
}

async function waitForBestTarget(service, attempts = 16) {
  let best = null
  for (let i = 0; i < attempts; i += 1) {
    const pages = (await targets()).filter((target) => target.type === 'page' && /^https?:/.test(target.url || ''))
    best = pages
      .map((target) => ({ target, score: pageScore(service, target) }))
      .sort((a, b) => b.score - a.score)[0] || null
    if (best && best.score > 0) return best.target
    await sleep(750)
  }
  return best && best.score > 0 ? best.target : null
}

async function waitForHealthyPage(service, target, attempts = 16) {
  let health = null
  for (let i = 0; i < attempts; i += 1) {
    health = await pageHealth(target)
    if (health.hasLoginInput) return health
    if (service.key === 'teams') return health
    if (!health.blankLike && health.readyState === 'complete') return health
    await sleep(1000)
    const nextTarget = await waitForBestTarget(service, 2)
    if (nextTarget) target = nextTarget
  }
  return health
}

async function main() {
  const allTargets = await targets()
  const panel = allTargets.find((target) => target.type === 'page' && target.url.endsWith('/src/renderer/panel.html'))
  assert.ok(panel, 'MailStudio panel target not found')

  const panelClient = await connect(panel.webSocketDebuggerUrl)
  const results = []
  const issues = []

  try {
    await panelClient.send('Runtime.enable')
    await panelClient.send('Log.enable').catch(() => {})

    const initial = await evaluate(panelClient, `window.panelApi.getSnapshot().then((snapshot) => ({
      activeServiceKey: snapshot.activeServiceKey,
      services: snapshot.services
        .filter((service) => service.visible)
        .map((service) => ({
          key: service.key,
          label: service.label,
          url: service.url,
          href: service.href,
          feedKind: service.feed && service.feed.kind,
          feedState: service.feed && service.feed.state,
          feedItems: service.feed && Array.isArray(service.feed.items) ? service.feed.items.length : null
        }))
    }))`)

    for (const service of initial.services) {
      await evaluate(panelClient, `window.panelApi.sendCommand({ type: 'switch-service', serviceKey: ${JSON.stringify(service.key)} }); true`)
      await sleep(900)
      const afterSwitch = await evaluate(panelClient, 'window.panelApi.getSnapshot()')
      if (afterSwitch.activeServiceKey !== service.key) {
        issues.push(`${service.key}: active service stayed ${afterSwitch.activeServiceKey}`)
      }

      const target = await waitForBestTarget(service)
      if (!target) {
        issues.push(`${service.key}: no matching service page target`)
        results.push({ key: service.key, label: service.label, active: afterSwitch.activeServiceKey, target: null })
        continue
      }

      const health = await waitForHealthyPage(service, target)
      const allowedBlank = service.key === 'teams'
      if (health.hasLoginInput) issues.push(`${service.key}: sign-in form visible`)
      if (health.blankLike && !allowedBlank) issues.push(`${service.key}: blank service page`)
      if (health.readyState !== 'complete') issues.push(`${service.key}: document not complete (${health.readyState})`)

      results.push({
        key: service.key,
        label: service.label,
        active: afterSwitch.activeServiceKey,
        targetHost: hostOf(target.url),
        targetTitle: health.title,
        readyState: health.readyState,
        blankLike: health.blankLike,
        hasLoginInput: health.hasLoginInput,
        bodyTextLength: health.bodyTextLength,
        buttonCount: health.buttonCount,
        inputCount: health.inputCount,
        errorEvents: health.errorEvents,
        feedState: service.feedState,
        feedItems: service.feedItems
      })
    }

    const panelHealth = await evaluate(panelClient, `(() => ({
      className: document.body.className,
      serviceRows: document.querySelectorAll('.service-row').length,
      settingsOpen: document.body.classList.contains('settings-open'),
      firstBoot: document.body.classList.contains('first-boot'),
      blankLike: !document.body || document.body.innerText.trim().length === 0
    }))()`)
    if (panelHealth.blankLike) issues.push('panel: blank body')
    if (panelHealth.firstBoot) issues.push('panel: first boot overlay active')

    await evaluate(panelClient, `window.panelApi.sendCommand({ type: 'switch-service', serviceKey: ${JSON.stringify(initial.activeServiceKey)} }); true`)
    await sleep(500)

    const report = {
      checkedServices: results.length,
      initialActiveServiceKey: initial.activeServiceKey,
      finalActiveServiceKey: (await evaluate(panelClient, 'window.panelApi.getSnapshot()')).activeServiceKey,
      panelHealth,
      results,
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
