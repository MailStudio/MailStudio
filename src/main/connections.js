const secureStore = require('./secure-store')
const oauth = require('./oauth')
const apiFeeds = require('./api-feeds')

// Orchestrates the API providers: OAuth lifecycle, token refresh, connection
// state, and feed access. main.js talks ONLY to this module — it never touches
// raw tokens or the OAuth engine directly.
//
// Provider → the feed kinds it powers once connected:
//   microsoft → mail, calendar
//   asana     → asana
// (Teams stays on title-scrape; it has no API path here.)

const STATUS = {
  DISCONNECTED: 'disconnected',
  CONNECTING: 'connecting',
  CONNECTED: 'connected',
  ERROR: 'error'
}

const PROVIDER_FEEDS = {
  microsoft: ['mail', 'calendar'],
  asana: ['asana']
}

// Live, in-memory state per provider (account is the human-readable identity).
const state = {
  microsoft: { status: STATUS.DISCONNECTED, account: null, error: null },
  asana: { status: STATUS.DISCONNECTED, account: null, error: null }
}

let config = {
  microsoft: { clientId: '', tenant: 'common' },
  asana: { clientId: '' }
}
let onChange = () => {}
let partitionForProvider = () => 'persist:mailstudio'

function setStatus(provider, status, extra) {
  state[provider] = { ...state[provider], status, ...(extra || {}) }
  try {
    onChange(provider)
  } catch {
    /* ignore listener errors */
  }
}

function normalizeConfig(raw) {
  const r = raw && typeof raw === 'object' ? raw : {}
  const ms = (r.microsoft && typeof r.microsoft === 'object') ? r.microsoft : {}
  const as = (r.asana && typeof r.asana === 'object') ? r.asana : {}
  return {
    microsoft: {
      clientId: typeof ms.clientId === 'string' ? ms.clientId.trim() : '',
      tenant: (typeof ms.tenant === 'string' && ms.tenant.trim()) || 'common'
    },
    asana: {
      clientId: typeof as.clientId === 'string' ? as.clientId.trim() : ''
    }
  }
}

// Restore connection state from the encrypted vault on launch: a stored token
// means we start "connected" (optimistically) and let the first refresh/feed
// call correct us to 'error' if the grant was revoked.
function init({ config: cfg, partitionForProvider: partitionResolver, onChange: cb }) {
  config = normalizeConfig(cfg)
  if (typeof partitionResolver === 'function') partitionForProvider = partitionResolver
  if (typeof cb === 'function') onChange = cb

  for (const provider of Object.keys(state)) {
    const token = secureStore.getToken(provider)
    if (token && token.refreshToken) {
      state[provider] = {
        status: STATUS.CONNECTED,
        account: token.account || null,
        error: null
      }
    }
  }
}

function setConfig(cfg) {
  config = normalizeConfig(cfg)
}

function getConfig() {
  return config
}

function isConfigured(provider) {
  if (!(config[provider] && config[provider].clientId)) return false
  // Asana's token endpoint requires a client secret (even with PKCE), so it's
  // only fully configured once both the client ID and the secret are present.
  if (provider === 'asana') return secureStore.hasSecret('asana')
  return true
}

// The Asana OAuth client secret lives in the encrypted vault, never in the
// plaintext settings file and never sent back to the renderer.
function setAsanaSecret(value) {
  secureStore.setSecret('asana', typeof value === 'string' ? value.trim() : '')
}

function hasAsanaSecret() {
  return secureStore.hasSecret('asana')
}

function isConnected(provider) {
  return state[provider] && state[provider].status === STATUS.CONNECTED
}

// Snapshot slice for the renderer's account rail.
function getStatus() {
  return {
    microsoft: {
      ...state.microsoft,
      configured: isConfigured('microsoft'),
      feeds: PROVIDER_FEEDS.microsoft
    },
    asana: {
      ...state.asana,
      configured: isConfigured('asana'),
      // Asana needs a client secret in addition to the client ID; surface whether
      // one is stored so the setup UI can reflect "saved" without exposing it.
      secretRequired: true,
      secretSet: hasAsanaSecret(),
      feeds: PROVIDER_FEEDS.asana
    },
    encryptionAvailable: secureStore.encryptionAvailable()
  }
}

// Which provider (if any) owns a given feed kind, and whether it's connected.
function providerForFeed(feedKind) {
  for (const [provider, feeds] of Object.entries(PROVIDER_FEEDS)) {
    if (feeds.includes(feedKind)) return provider
  }
  return null
}

function feedIsLive(feedKind) {
  const provider = providerForFeed(feedKind)
  return provider ? isConnected(provider) : false
}

/* ---------- OAuth lifecycle ---------- */
async function connect(provider, { parentWindow } = {}) {
  if (!oauth.isProvider(provider)) {
    throw new Error(`Unknown provider: ${provider}`)
  }
  if (state[provider].status === STATUS.CONNECTING) {
    // A popup is already open for this provider — don't stack a second flow.
    throw new Error('Sign-in already in progress')
  }
  if (!isConfigured(provider)) {
    setStatus(provider, STATUS.ERROR, { error: `Add a ${provider} client ID first.` })
    throw new Error(`${provider} is not configured`)
  }
  setStatus(provider, STATUS.CONNECTING, { error: null })
  try {
    const tokenSet = await oauth.authorize({
      provider,
      clientId: config[provider].clientId,
      clientSecret: provider === 'asana' ? secureStore.getSecret('asana') : undefined,
      tenant: provider === 'microsoft' ? config.microsoft.tenant : undefined,
      partition: partitionForProvider(provider),
      parentWindow
    })
    const account = await loadAccount(provider, tokenSet)
    tokenSet.account = account
    secureStore.setToken(provider, tokenSet)
    setStatus(provider, STATUS.CONNECTED, { account, error: null })
    return account
  } catch (err) {
    if (err instanceof oauth.CancelledError) {
      // User closed the popup or declined consent — back to a clean slate, not
      // an error state.
      setStatus(provider, STATUS.DISCONNECTED, { error: null })
    } else {
      setStatus(provider, STATUS.ERROR, { error: err.message || 'Sign-in failed' })
    }
    throw err
  }
}

function disconnect(provider) {
  secureStore.clearToken(provider)
  setStatus(provider, STATUS.DISCONNECTED, { account: null, error: null })
}

// Fetch a human-readable identity (and, for Asana, the workspace gid we need
// for task queries) right after a successful grant.
async function loadAccount(provider, tokenSet) {
  try {
    if (provider === 'asana') {
      const me = await apiFeeds.fetchAsanaMe(tokenSet.accessToken)
      return { name: me.name, workspaceGid: me.workspaceGid }
    }
    if (provider === 'microsoft') {
      // /me is covered by User.Read; reuse the Graph helper inline.
      const res = await fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,userPrincipalName', {
        headers: { Authorization: `Bearer ${tokenSet.accessToken}`, Accept: 'application/json' },
        signal: AbortSignal.timeout(15000)
      })
      if (res.ok) {
        const json = await res.json()
        // userPrincipalName is the signed-in mailbox address. We keep it so the
        // mailbox-discovery scrape can tell the primary account apart from the
        // shared mailboxes it sits beside in the OWA folder tree (otherwise the
        // primary's own email-labeled root gets re-added as a duplicate tab).
        const email = typeof json.userPrincipalName === 'string' ? json.userPrincipalName.trim().toLowerCase() : ''
        return { name: json.displayName || json.userPrincipalName || 'Microsoft', email: email || null }
      }
    }
  } catch {
    /* non-fatal — identity is cosmetic */
  }
  return { name: oauth.PROVIDERS[provider] ? oauth.PROVIDERS[provider].label : provider }
}

/* ---------- Token access (auto-refresh) ---------- */
// How long past expiresAt we still hand out a stale access token when a
// refresh fails transiently (expiresAt already carries a 60s safety margin,
// so "slightly expired" tokens usually still work).
const TOKEN_GRACE_MS = 5 * 60 * 1000

// In-flight refresh promise per provider — mail and calendar poll the same
// Microsoft token from one loop, and a refresh_token grant rotates the token,
// so two concurrent refreshes would invalidate each other.
const refreshPromises = {}

// A refresh failure only means the grant is dead when the token endpoint says
// so explicitly. Network blips, timeouts (TokenError status 0), 429s and 5xx
// are transient and must not tear the connection down.
const DEAD_GRANT_OAUTH_ERRORS = new Set([
  'invalid_grant',        // refresh token expired/revoked
  'interaction_required', // user must re-consent / re-auth
  'invalid_client',       // client credentials no longer valid (e.g. Asana secret)
  'unauthorized_client'   // client not allowed this grant anymore
])

function isGrantDead(err) {
  if (!(err instanceof oauth.TokenError)) return false
  // Only tear the connection down on a definitive auth rejection. The token
  // endpoint signals these via an OAuth `error` code in the body (always a 400
  // per the spec) or a 401. A bare 400 with no recognized error code is treated
  // as transient — a malformed/throttled/server blip must not force a reconnect.
  if (err.oauthError && DEAD_GRANT_OAUTH_ERRORS.has(err.oauthError)) return true
  return err.status === 401
}

// Returns a valid access token, refreshing if expired. Returns null (and moves
// the provider to 'error') when no token exists or the refresh grant is dead.
// Transient refresh failures keep the provider connected so the next poll
// retries, returning the stale token while it's within the grace window.
async function getAccessToken(provider) {
  const token = secureStore.getToken(provider)
  if (!token || !token.accessToken) {
    return null
  }
  if (Date.now() < token.expiresAt) {
    return token.accessToken
  }
  // Expired → refresh (single-flight: concurrent callers share one grant).
  if (!refreshPromises[provider]) {
    refreshPromises[provider] = refreshAccessToken(provider, token).finally(() => {
      delete refreshPromises[provider]
    })
  }
  return refreshPromises[provider]
}

async function refreshAccessToken(provider, token) {
  try {
    const refreshed = await oauth.refreshTokens({
      provider,
      clientId: config[provider].clientId,
      clientSecret: provider === 'asana' ? secureStore.getSecret('asana') : undefined,
      tenant: provider === 'microsoft' ? config.microsoft.tenant : undefined,
      refreshToken: token.refreshToken
    })
    refreshed.account = token.account || null
    secureStore.setToken(provider, refreshed)
    if (state[provider].status !== STATUS.CONNECTED) {
      setStatus(provider, STATUS.CONNECTED, { error: null })
    }
    return refreshed.accessToken
  } catch (err) {
    if (isGrantDead(err)) {
      setStatus(provider, STATUS.ERROR, { error: 'Session expired — reconnect.' })
      return null
    }
    // Transient failure — leave status alone (the next poll retries) and hand
    // out the slightly-stale token if it's still within the grace window.
    if (Date.now() < token.expiresAt + TOKEN_GRACE_MS) {
      return token.accessToken
    }
    return null
  }
}

// Per-provider cooldown after a 429/503 throttle: feed accessors return null
// (main.js keeps its cache) until this timestamp passes. Connection status is
// untouched — throttling is not an auth problem.
const nextAllowedAt = { microsoft: 0, asana: 0 }

// Generic (non-auth, non-throttle) failures — 5xx without Retry-After, network
// drops, parse errors — get exponential backoff so a flapping/unreachable API
// isn't hammered every feed tick. The streak resets on the first success.
const GENERIC_BACKOFF_BASE_MS = 30000
const GENERIC_BACKOFF_MAX_MS = 300000
const apiFailureStreak = { microsoft: 0, asana: 0 }

function isThrottled(provider) {
  return Boolean(provider) && Date.now() < (nextAllowedAt[provider] || 0)
}

function noteThrottle(provider, err) {
  if (!provider) return
  nextAllowedAt[provider] = Date.now() + (err.retryAfter || 60) * 1000
  // An explicit server throttle isn't part of a generic failure streak.
  apiFailureStreak[provider] = 0
}

function noteApiSuccess(provider) {
  if (provider) apiFailureStreak[provider] = 0
}

function noteApiFailure(provider) {
  if (!provider) return
  const streak = (apiFailureStreak[provider] || 0) + 1
  apiFailureStreak[provider] = streak
  const backoff = Math.min(GENERIC_BACKOFF_BASE_MS * 2 ** (streak - 1), GENERIC_BACKOFF_MAX_MS)
  nextAllowedAt[provider] = Math.max(nextAllowedAt[provider] || 0, Date.now() + backoff)
}

// Run an authed API call, transparently refreshing once on a 401.
async function withToken(provider, fn) {
  let accessToken = await getAccessToken(provider)
  if (!accessToken) return null
  try {
    const out = await fn(accessToken)
    noteApiSuccess(provider)
    return out
  } catch (err) {
    if (err instanceof apiFeeds.ThrottledError) {
      noteThrottle(provider, err)
      return null
    }
    if (err instanceof apiFeeds.AuthError) {
      // Force a refresh by expiring our cached token, then retry once — but
      // only stamp the vault if it still holds the token that just 401'd; a
      // parallel caller may already have persisted a fresh tokenSet.
      const token = secureStore.getToken(provider)
      if (token && token.accessToken === accessToken) {
        secureStore.setToken(provider, { ...token, expiresAt: 0 })
      }
      accessToken = await getAccessToken(provider)
      if (!accessToken) return null
      try {
        const out2 = await fn(accessToken)
        noteApiSuccess(provider)
        return out2
      } catch (err2) {
        if (err2 instanceof apiFeeds.ThrottledError) {
          noteThrottle(provider, err2)
          return null
        }
        if (err2 instanceof apiFeeds.AuthError) {
          setStatus(provider, STATUS.ERROR, { error: 'Session expired — reconnect.' })
          return null
        }
        console.warn(`[api] ${provider} request failed after refresh:`, err2 && err2.message)
        noteApiFailure(provider)
        throw err2
      }
    }
    // Non-auth, non-throttle failure (e.g. 403 missing consent, 5xx, network).
    // Log it and apply exponential backoff so we stop hammering a flapping API.
    console.warn(`[api] ${provider} request failed:`, err && err.message)
    noteApiFailure(provider)
    throw err
  }
}

/* ---------- Feed accessors (used by main.js refreshFeed) ---------- */
async function getMailFeed() {
  return withToken('microsoft', (t) => apiFeeds.fetchMail(t))
}

async function getMailUnreadCount() {
  if (isThrottled('microsoft')) return null
  const count = await withToken('microsoft', (t) => apiFeeds.fetchMailUnreadCount(t))
  // A real count (including 0) comes back as a number. Anything else means we
  // couldn't determine it — missing/expired token, a fresh throttle, or a
  // transient API failure. Report null so the caller keeps its cached badge
  // rather than flickering the unread count to zero on a blip.
  return typeof count === 'number' ? count : null
}

async function getCalendarFeed() {
  return withToken('microsoft', (t) => apiFeeds.fetchCalendar(t))
}

async function getAsanaFeed() {
  const token = secureStore.getToken('asana')
  let workspaceGid = token && token.account ? token.account.workspaceGid : null
  return withToken('asana', async (t) => {
    if (!workspaceGid) {
      // Connect-time fetchAsanaMe failed, so the stored account never got its
      // workspace gid — recover it lazily and persist so the UI gets a name.
      const me = await apiFeeds.fetchAsanaMe(t)
      if (!me.workspaceGid) {
        return { state: 'error', items: [] }
      }
      workspaceGid = me.workspaceGid
      // Re-read the vault entry: the refresh above may have rotated the tokenSet.
      const fresh = secureStore.getToken('asana')
      if (fresh) {
        const account = { ...(fresh.account || {}), name: me.name, workspaceGid: me.workspaceGid }
        secureStore.setToken('asana', { ...fresh, account })
        setStatus('asana', state.asana.status, { account })
      }
    }
    return apiFeeds.fetchAsanaTasks(t, workspaceGid)
  })
}

// kind → its accessor, so main.js can fetch generically. Returns null (keep
// cache) while the owning provider is in a throttle cooldown.
async function getFeed(kind) {
  if (isThrottled(providerForFeed(kind))) return null
  if (kind === 'mail') return getMailFeed()
  if (kind === 'calendar') return getCalendarFeed()
  if (kind === 'asana') return getAsanaFeed()
  return null
}

module.exports = {
  STATUS,
  PROVIDER_FEEDS,
  init,
  setConfig,
  getConfig,
  setAsanaSecret,
  hasAsanaSecret,
  isConfigured,
  isConnected,
  getStatus,
  providerForFeed,
  feedIsLive,
  connect,
  disconnect,
  getAccessToken,
  getMailUnreadCount,
  getFeed
}
