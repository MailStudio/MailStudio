'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')
const path = require('node:path')

const ROOT = path.resolve(__dirname, '..', '..')

function mockJsonResponse(status, body, headers) {
  const bag = new Map(Object.entries(headers || {}).map(([key, value]) => [key.toLowerCase(), value]))
  return {
    status,
    ok: status >= 200 && status < 300,
    headers: { get: (key) => bag.get(String(key).toLowerCase()) || null },
    text: async () => JSON.stringify(body),
    json: async () => body
  }
}

async function withMockedFetch(mock, fn) {
  const prevFetch = global.fetch
  const prevAbortSignal = global.AbortSignal
  global.fetch = mock
  global.AbortSignal = { timeout: () => ({ aborted: false }) }
  try {
    return await fn()
  } finally {
    global.fetch = prevFetch
    global.AbortSignal = prevAbortSignal
  }
}

function loadFresh(modulePath, mocks) {
  const resolved = require.resolve(modulePath)
  const previous = new Map()
  for (const [target, exports] of Object.entries(mocks || {})) {
    const key = require.resolve(target)
    previous.set(key, require.cache[key])
    require.cache[key] = {
      id: key,
      filename: key,
      loaded: true,
      exports
    }
  }
  delete require.cache[resolved]
  const loaded = require(modulePath)
  return {
    loaded,
    restore() {
      delete require.cache[resolved]
      for (const [key, entry] of previous.entries()) {
        if (entry) require.cache[key] = entry
        else delete require.cache[key]
      }
    }
  }
}

function parseForm(body) {
  return Object.fromEntries(new URLSearchParams(String(body)).entries())
}

test('simulator: Graph calendar fetch sends auth, JSON, timezone, and shapes events', async () => {
  const apiFeeds = require(path.join(ROOT, 'src', 'main', 'api-feeds.js'))
  const calls = []
  await withMockedFetch(async (url, options) => {
    calls.push({ url, options })
    return mockJsonResponse(200, {
      value: [{
        id: 'evt1',
        subject: 'Planning',
        start: { dateTime: '2026-06-16T09:30:00.0000000' },
        isAllDay: false,
        webLink: 'https://outlook.office.com/calendar/item/evt1',
        isCancelled: false
      }]
    })
  }, async () => {
    const result = await apiFeeds.fetchCalendar('graph-token')
    assert.equal(result.state, 'ok')
    assert.equal(result.items[0].id, 'evt1')
    assert.equal(result.items[0].title, 'Planning')
    assert.equal(result.items[0].startIso, '2026-06-16T09:30:00.0000000')
    assert.equal(result.items[0].webLink, 'https://outlook.office.com/calendar/item/evt1')
  })

  assert.equal(calls.length, 1)
  assert.match(calls[0].url, /^https:\/\/graph\.microsoft\.com\/v1\.0\/me\/calendarView\?/)
  assert.match(calls[0].url, /startDateTime=/)
  assert.match(calls[0].url, /endDateTime=/)
  assert.equal(calls[0].options.headers.Authorization, 'Bearer graph-token')
  assert.equal(calls[0].options.headers.Accept, 'application/json')
  assert.match(calls[0].options.headers.Prefer, /^outlook\.timezone="/)
})

test('simulator: API wrappers surface auth and throttle states as typed errors', async () => {
  const apiFeeds = require(path.join(ROOT, 'src', 'main', 'api-feeds.js'))

  await withMockedFetch(async () => mockJsonResponse(401, { error: { message: 'nope' } }), async () => {
    await assert.rejects(() => apiFeeds.fetchMail('bad-token'), apiFeeds.AuthError)
  })

  await withMockedFetch(async () => mockJsonResponse(429, { error: 'slow down' }, { 'retry-after': '7' }), async () => {
    await assert.rejects(
      async () => apiFeeds.fetchMailUnreadCount('busy-token'),
      (err) => err instanceof apiFeeds.ThrottledError && err.retryAfter === 7
    )
  })
})

test('simulator: Teams presence controls post preferred presence and reset through Graph', async () => {
  const apiFeeds = require(path.join(ROOT, 'src', 'main', 'api-feeds.js'))
  const calls = []
  await withMockedFetch(async (url, options) => {
    calls.push({ url, options })
    return mockJsonResponse(204, {})
  }, async () => {
    await apiFeeds.setTeamsPreferredPresence('graph-token', 'user id/1', {
      availability: 'DoNotDisturb',
      activity: 'DoNotDisturb',
      expirationDuration: 'PT1H'
    })
    await apiFeeds.clearTeamsPreferredPresence('graph-token', 'user id/1')
  })

  assert.equal(calls.length, 2)
  assert.equal(
    calls[0].url,
    'https://graph.microsoft.com/v1.0/users/user%20id%2F1/presence/setUserPreferredPresence'
  )
  assert.equal(calls[0].options.method, 'POST')
  assert.equal(calls[0].options.headers.Authorization, 'Bearer graph-token')
  assert.equal(calls[0].options.headers['Content-Type'], 'application/json')
  assert.deepEqual(JSON.parse(calls[0].options.body), {
    availability: 'DoNotDisturb',
    activity: 'DoNotDisturb',
    expirationDuration: 'PT1H'
  })
  assert.equal(
    calls[1].url,
    'https://graph.microsoft.com/v1.0/users/user%20id%2F1/presence/clearUserPreferredPresence'
  )
  assert.deepEqual(JSON.parse(calls[1].options.body), {})
})

test('simulator: Asana task fetch uses bearer auth, encoded workspace, and preserves permalinks', async () => {
  const apiFeeds = require(path.join(ROOT, 'src', 'main', 'api-feeds.js'))
  const calls = []
  await withMockedFetch(async (url, options) => {
    calls.push({ url, options })
    return mockJsonResponse(200, {
      data: [
        { gid: '1', name: '  ', permalink_url: 'https://app.asana.com/0/1/1', due_on: null },
        { gid: '2', name: 'Launch task', permalink_url: 'https://app.asana.com/0/1/2', due_on: '2026-06-16' }
      ]
    })
  }, async () => {
    const result = await apiFeeds.fetchAsanaTasks('asana-token', 'workspace/with space')
    assert.equal(result.state, 'ok')
    assert.deepEqual(result.items, [{
      id: '2',
      name: 'Launch task',
      subtasks: [],
      taskUrl: 'https://app.asana.com/0/1/2',
      dueOn: '2026-06-16'
    }])
  })

  assert.equal(calls.length, 1)
  assert.match(calls[0].url, /^https:\/\/app\.asana\.com\/api\/1\.0\/tasks\?/)
  assert.match(calls[0].url, /workspace=workspace%2Fwith%20space/)
  assert.match(calls[0].url, /completed_since=now/)
  assert.match(calls[0].url, /opt_fields=name,permalink_url,due_on/)
  assert.equal(calls[0].options.headers.Authorization, 'Bearer asana-token')
  assert.equal(calls[0].options.headers.Accept, 'application/json')
})

test('simulator: OAuth refresh posts provider-correct token forms', async () => {
  const electronPath = require.resolve('electron')
  const { loaded: oauth, restore } = loadFresh(path.join(ROOT, 'src', 'main', 'oauth.js'), {
    [electronPath]: { BrowserWindow: function BrowserWindow() {}, session: {} }
  })
  try {
    const calls = []
    await withMockedFetch(async (url, options) => {
      calls.push({ url, options, form: parseForm(options.body) })
      return mockJsonResponse(200, {
        access_token: `${calls.length}-access`,
        refresh_token: `${calls.length}-refresh`,
        expires_in: 3600,
        scope: 'scope'
      })
    }, async () => {
      const microsoft = await oauth.refreshTokens({
        provider: 'microsoft',
        clientId: 'ms-client',
        tenant: 'organizations',
        refreshToken: 'ms-refresh'
      })
      const asana = await oauth.refreshTokens({
        provider: 'asana',
        clientId: 'asana-client',
        clientSecret: 'asana-secret',
        refreshToken: 'asana-refresh'
      })
      assert.equal(microsoft.accessToken, '1-access')
      assert.equal(asana.accessToken, '2-access')
    })

    assert.match(calls[0].url, /login\.microsoftonline\.com\/organizations\/oauth2\/v2\.0\/token$/)
    assert.deepEqual(calls[0].form, {
      grant_type: 'refresh_token',
      refresh_token: 'ms-refresh',
      client_id: 'ms-client',
      scope: oauth.PROVIDERS.microsoft.scope
    })
    assert.ok(!Object.prototype.hasOwnProperty.call(calls[0].form, 'client_secret'))

    assert.equal(calls[1].url, 'https://app.asana.com/-/oauth_token')
    assert.deepEqual(calls[1].form, {
      grant_type: 'refresh_token',
      refresh_token: 'asana-refresh',
      client_id: 'asana-client',
      scope: oauth.PROVIDERS.asana.scope,
      client_secret: 'asana-secret'
    })
  } finally {
    restore()
  }
})

test('simulator: OAuth token failures preserve status and OAuth error code', async () => {
  const electronPath = require.resolve('electron')
  const { loaded: oauth, restore } = loadFresh(path.join(ROOT, 'src', 'main', 'oauth.js'), {
    [electronPath]: { BrowserWindow: function BrowserWindow() {}, session: {} }
  })
  try {
    await withMockedFetch(async () => mockJsonResponse(400, {
      error: 'invalid_grant',
      error_description: 'expired refresh token'
    }), async () => {
      await assert.rejects(
        () => oauth.refreshTokens({
          provider: 'microsoft',
          clientId: 'ms-client',
          refreshToken: 'dead-refresh'
        }),
        (err) => err instanceof oauth.TokenError && err.status === 400 && err.oauthError === 'invalid_grant'
      )
    })

    await withMockedFetch(async () => {
      throw new Error('network offline')
    }, async () => {
      await assert.rejects(
        () => oauth.refreshTokens({
          provider: 'asana',
          clientId: 'asana-client',
          clientSecret: 'secret',
          refreshToken: 'refresh'
        }),
        (err) => err instanceof oauth.TokenError && err.status === 0 && err.oauthError === null
      )
    })
  } finally {
    restore()
  }
})

function loadConnectionsSimulator({ tokens, secrets, oauthImpl, apiFeedsImpl }) {
  const secureStorePath = path.join(ROOT, 'src', 'main', 'secure-store.js')
  const oauthPath = path.join(ROOT, 'src', 'main', 'oauth.js')
  const apiFeedsPath = path.join(ROOT, 'src', 'main', 'api-feeds.js')
  const secureStore = {
    getToken: (provider) => tokens[provider] || null,
    setToken: (provider, tokenSet) => {
      tokens[provider] = tokenSet || null
      return tokens[provider]
    },
    clearToken: (provider) => { tokens[provider] = null },
    getSecret: (name) => secrets[name] || null,
    setSecret: (name, value) => {
      if (value) secrets[name] = value
      else delete secrets[name]
    },
    hasSecret: (name) => Boolean(secrets[name]),
    encryptionAvailable: () => true
  }
  return loadFresh(path.join(ROOT, 'src', 'main', 'connections.js'), {
    [secureStorePath]: secureStore,
    [oauthPath]: oauthImpl,
    [apiFeedsPath]: apiFeedsImpl
  })
}

test('simulator: connections refreshes once after a Graph 401 and retries the feed', async () => {
  class AuthError extends Error {}
  class ThrottledError extends Error {
    constructor(message, retryAfter) {
      super(message)
      this.retryAfter = retryAfter
    }
  }
  class TokenError extends Error {}

  const tokens = {
    microsoft: {
      accessToken: 'old-access',
      refreshToken: 'refresh-token',
      expiresAt: Date.now() + 60_000,
      account: { name: 'Mina' }
    }
  }
  const fetchMailTokens = []
  let refreshCalls = 0
  const oauthImpl = {
    TokenError,
    CancelledError: class CancelledError extends Error {},
    PROVIDERS: { microsoft: { label: 'Microsoft' }, asana: { label: 'Asana' } },
    isProvider: () => true,
    refreshTokens: async () => {
      refreshCalls += 1
      return {
        accessToken: 'fresh-access',
        refreshToken: 'fresh-refresh',
        expiresAt: Date.now() + 60_000,
        scope: 'scope'
      }
    }
  }
  const apiFeedsImpl = {
    AuthError,
    ThrottledError,
    fetchMail: async (token) => {
      fetchMailTokens.push(token)
      if (token === 'old-access') throw new AuthError('stale')
      return { state: 'ok', items: [{ id: 'm1' }] }
    }
  }

  const { loaded: connections, restore } = loadConnectionsSimulator({
    tokens,
    secrets: {},
    oauthImpl,
    apiFeedsImpl
  })
  try {
    connections.init({ config: { microsoft: { clientId: 'ms-client', tenant: 'common' } } })
    const result = await connections.getFeed('mail')
    assert.deepEqual(result, { state: 'ok', items: [{ id: 'm1' }] })
    assert.deepEqual(fetchMailTokens, ['old-access', 'fresh-access'])
    assert.equal(refreshCalls, 1)
    assert.equal(tokens.microsoft.accessToken, 'fresh-access')
    assert.deepEqual(tokens.microsoft.account, { name: 'Mina' })
    assert.equal(connections.isConnected('microsoft'), true)
  } finally {
    restore()
  }
})

test('simulator: connections throttling returns null and suppresses immediate API retries', async () => {
  class AuthError extends Error {}
  class ThrottledError extends Error {
    constructor(message, retryAfter) {
      super(message)
      this.retryAfter = retryAfter
    }
  }
  class TokenError extends Error {}

  const tokens = {
    microsoft: {
      accessToken: 'calendar-access',
      refreshToken: 'calendar-refresh',
      expiresAt: Date.now() + 60_000,
      account: { name: 'Mina' }
    }
  }
  let fetchCalendarCalls = 0
  const { loaded: connections, restore } = loadConnectionsSimulator({
    tokens,
    secrets: {},
    oauthImpl: {
      TokenError,
      CancelledError: class CancelledError extends Error {},
      PROVIDERS: { microsoft: { label: 'Microsoft' }, asana: { label: 'Asana' } },
      isProvider: () => true,
      refreshTokens: async () => {
        throw new Error('should not refresh')
      }
    },
    apiFeedsImpl: {
      AuthError,
      ThrottledError,
      fetchCalendar: async () => {
        fetchCalendarCalls += 1
        throw new ThrottledError('slow down', 60)
      }
    }
  })
  try {
    connections.init({ config: { microsoft: { clientId: 'ms-client', tenant: 'common' } } })
    assert.equal(await connections.getFeed('calendar'), null)
    assert.equal(await connections.getFeed('calendar'), null)
    assert.equal(fetchCalendarCalls, 1)
    assert.equal(connections.isConnected('microsoft'), true)
  } finally {
    restore()
  }
})

test('simulator: Asana feed lazily discovers workspace and persists account metadata', async () => {
  class AuthError extends Error {}
  class ThrottledError extends Error {
    constructor(message, retryAfter) {
      super(message)
      this.retryAfter = retryAfter
    }
  }
  class TokenError extends Error {}

  const tokens = {
    asana: {
      accessToken: 'asana-access',
      refreshToken: 'asana-refresh',
      expiresAt: Date.now() + 60_000,
      account: { name: 'Asana' }
    }
  }
  const calls = []
  const { loaded: connections, restore } = loadConnectionsSimulator({
    tokens,
    secrets: { asana: 'asana-secret' },
    oauthImpl: {
      TokenError,
      CancelledError: class CancelledError extends Error {},
      PROVIDERS: { microsoft: { label: 'Microsoft' }, asana: { label: 'Asana' } },
      isProvider: () => true,
      refreshTokens: async () => {
        throw new Error('should not refresh')
      }
    },
    apiFeedsImpl: {
      AuthError,
      ThrottledError,
      fetchAsanaMe: async (token) => {
        calls.push(['me', token])
        return { name: 'Asha', workspaceGid: 'workspace-1' }
      },
      fetchAsanaTasks: async (token, workspaceGid) => {
        calls.push(['tasks', token, workspaceGid])
        return { state: 'ok', items: [{ id: 'task-1' }] }
      }
    }
  })
  try {
    connections.init({ config: { asana: { clientId: 'asana-client' } } })
    const result = await connections.getFeed('asana')
    assert.deepEqual(result, { state: 'ok', items: [{ id: 'task-1' }] })
    assert.deepEqual(calls, [
      ['me', 'asana-access'],
      ['tasks', 'asana-access', 'workspace-1']
    ])
    assert.deepEqual(tokens.asana.account, { name: 'Asha', workspaceGid: 'workspace-1' })
    assert.equal(connections.getStatus().asana.secretSet, true)
  } finally {
    restore()
  }
})
