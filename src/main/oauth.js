const crypto = require('crypto')
const { BrowserWindow, session } = require('electron')

// Generic OAuth2 Authorization-Code-with-PKCE engine for desktop. No client
// secret is ever used or stored — both Microsoft (public client) and Asana
// (PKCE-enabled native app) issue tokens to a confidential-secret-free flow,
// which is the only safe model for code that ships in a public GitHub repo.
//
// The redirect is captured by intercepting the auth popup's navigation to the
// redirect URI (RFC 8252 style, but in-window) — so there's no loopback HTTP
// server and no custom protocol registration to maintain.

// Shared, app-wide redirect target. Registered identically in the Azure app
// ("Mobile and desktop applications" platform) and the Asana app. The popup
// never actually loads it; we read the ?code= off the navigation and stop it.
const REDIRECT_URI = 'http://localhost/mailstudio-auth'

// Typed failure from the token endpoint, so callers can tell a definitive
// grant death (e.g. invalid_grant) apart from a transient network/server blip.
// `status` is the HTTP status (0 for network-level failures and timeouts) and
// `oauthError` is the OAuth `error` code from the response body when present.
class TokenError extends Error {
  constructor(message, { status, oauthError } = {}) {
    super(message)
    this.name = 'TokenError'
    this.status = typeof status === 'number' ? status : 0
    this.oauthError = oauthError || null
  }
}

// The user backed out of sign-in (closed the window or denied consent) — not
// an error condition, so callers can avoid surfacing it as one.
class CancelledError extends Error {
  constructor(message) {
    super(message || 'Sign-in was cancelled.')
    this.name = 'CancelledError'
  }
}

// Provider definitions. clientId/tenant are injected from user config at call
// time so each person who clones the repo supplies their own registrations.
const PROVIDERS = {
  microsoft: {
    label: 'Microsoft',
    // Mail + Calendar in one grant; offline_access yields the refresh token.
    scope: 'openid profile offline_access User.Read Mail.Read Calendars.Read',
    authorizePath: (tenant) =>
      `https://login.microsoftonline.com/${tenant || 'common'}/oauth2/v2.0/authorize`,
    tokenPath: (tenant) =>
      `https://login.microsoftonline.com/${tenant || 'common'}/oauth2/v2.0/token`
  },
  asana: {
    label: 'Asana',
    scope: 'default',
    authorizePath: () => 'https://app.asana.com/-/oauth_authorize',
    tokenPath: () => 'https://app.asana.com/-/oauth_token'
  }
}

function base64url(buffer) {
  return buffer.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
}

function createPkce() {
  const verifier = base64url(crypto.randomBytes(32))
  const challenge = base64url(crypto.createHash('sha256').update(verifier).digest())
  return { verifier, challenge }
}

// Normalize a raw token endpoint response into our stored tokenSet shape.
// expiresAt is an absolute epoch-ms timestamp with a 60s safety margin shaved
// off so we refresh slightly early rather than mid-request.
function toTokenSet(json, fallbackRefresh) {
  const expiresInMs = (Number(json.expires_in) || 3600) * 1000
  return {
    accessToken: json.access_token,
    // Asana/MS both rotate refresh tokens; keep the old one if a refresh
    // response omits a new one.
    refreshToken: json.refresh_token || fallbackRefresh || null,
    expiresAt: Date.now() + expiresInMs - 60000,
    scope: json.scope || '',
    account: null
  }
}

function isProvider(name) {
  return Object.prototype.hasOwnProperty.call(PROVIDERS, name)
}

// Open the consent popup and resolve with a tokenSet. The popup shares the
// given session partition, so for Microsoft the existing signed-in web session
// makes consent a single click (or fully silent when already consented).
function authorize({ provider, clientId, tenant, partition, parentWindow }) {
  return new Promise((resolve, reject) => {
    if (!isProvider(provider)) {
      reject(new Error(`Unknown provider: ${provider}`))
      return
    }
    if (!clientId) {
      reject(new Error(`${PROVIDERS[provider].label} client ID is not configured.`))
      return
    }

    const def = PROVIDERS[provider]
    const pkce = createPkce()
    const state = base64url(crypto.randomBytes(16))

    const authUrl = new URL(def.authorizePath(tenant))
    const params = {
      client_id: clientId,
      response_type: 'code',
      redirect_uri: REDIRECT_URI,
      scope: def.scope,
      state,
      code_challenge: pkce.challenge,
      code_challenge_method: 'S256'
    }
    // Microsoft: prompt=select_account avoids silently reusing the wrong tenant
    // identity; response_mode=query keeps the code on the URL we intercept.
    if (provider === 'microsoft') {
      params.response_mode = 'query'
      params.prompt = 'select_account'
    }
    for (const [k, v] of Object.entries(params)) {
      authUrl.searchParams.set(k, v)
    }

    // A standalone, framed, centered window — NOT a modal child. On macOS a
    // modal child of the frameless main window renders as a titlebar "sheet"
    // that dims the parent but never shows a usable popup, which is exactly the
    // "everything goes dark but no popup" symptom. A normal top-level window
    // shown on ready-to-show is reliably visible and focused.
    const hasParent = parentWindow && !parentWindow.isDestroyed()
    const authWindow = new BrowserWindow({
      width: 520,
      height: 700,
      parent: hasParent ? parentWindow : undefined,
      modal: false,
      show: false,
      center: true,
      resizable: true,
      minimizable: false,
      maximizable: false,
      fullscreenable: false,
      alwaysOnTop: true,
      autoHideMenuBar: true,
      title: `Sign in to ${def.label}`,
      backgroundColor: '#ffffff',
      webPreferences: {
        partition,
        sandbox: true,
        contextIsolation: true,
        nodeIntegration: false
      }
    })

    // Show once content is painted and keep it pinned above the app for its
    // whole life — the main window paints service pages as on-top BrowserViews,
    // so a transient sign-in window must stay above them to be usable.
    authWindow.once('ready-to-show', () => {
      if (authWindow.isDestroyed()) return
      authWindow.show()
      authWindow.focus()
      authWindow.moveTop()
      authWindow.setAlwaysOnTop(true, 'modal-panel')
    })

    let settled = false
    const finish = (fn, arg) => {
      if (settled) return
      settled = true
      try {
        if (!authWindow.isDestroyed()) authWindow.destroy()
      } catch {
        /* ignore */
      }
      fn(arg)
    }

    // Inspect every navigation for the redirect; capture code or error, then stop.
    // Use exact protocol+host+pathname match rather than startsWith() — the prefix
    // check would also accept URLs like http://localhost/mailstudio-auth-extra which
    // could confuse the handler on a misconfigured redirect registration.
    const redirectUrl = new URL(REDIRECT_URI)
    const handleNavigation = (event, url) => {
      let parsed
      try {
        parsed = new URL(url)
      } catch {
        return
      }
      if (
        parsed.protocol !== redirectUrl.protocol ||
        parsed.host !== redirectUrl.host ||
        parsed.pathname !== redirectUrl.pathname
      ) return
      event.preventDefault()
      const err = parsed.searchParams.get('error')
      if (err === 'access_denied') {
        // The user declined consent — a cancel, not a failure.
        finish(reject, new CancelledError('Sign-in was cancelled.'))
        return
      }
      if (err) {
        finish(reject, new Error(parsed.searchParams.get('error_description') || err))
        return
      }
      const code = parsed.searchParams.get('code')
      const returnedState = parsed.searchParams.get('state')
      if (returnedState !== state) {
        finish(reject, new Error('OAuth state mismatch — aborting for safety.'))
        return
      }
      if (!code) {
        finish(reject, new Error('No authorization code returned.'))
        return
      }
      exchangeCode({ def, tenant, clientId, code, verifier: pkce.verifier })
        .then((tokenSet) => finish(resolve, tokenSet))
        .catch((e) => finish(reject, e))
    }

    console.log(`[oauth] opening ${def.label} auth window → ${authUrl.origin}`)
    authWindow.webContents.on('will-redirect', handleNavigation)
    authWindow.webContents.on('will-navigate', handleNavigation)
    authWindow.on('closed', () => {
      if (!settled) {
        settled = true
        reject(new CancelledError('Sign-in window was closed before completing.'))
      }
    })

    authWindow.loadURL(authUrl.toString()).catch((e) => finish(reject, e))

    // Safety net: if ready-to-show is slow to fire (some MS pages defer it),
    // force the window visible so it can never get stuck hidden behind the app.
    setTimeout(() => {
      if (!settled && !authWindow.isDestroyed() && !authWindow.isVisible()) {
        authWindow.show()
        authWindow.focus()
      }
    }, 1500)
  })
}

async function postForm(url, form) {
  const body = new URLSearchParams(form).toString()
  let res
  try {
    res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json' },
      body,
      signal: AbortSignal.timeout(30000)
    })
  } catch (err) {
    // Network failure or timeout — typed with status 0 so callers can tell it
    // apart from a definitive rejection by the token endpoint.
    throw new TokenError((err && err.message) || 'Network error during token request.', { status: 0 })
  }
  const text = await res.text()
  let json
  try {
    json = text ? JSON.parse(text) : {}
  } catch {
    json = {}
  }
  if (!res.ok) {
    const msg = json.error_description || json.error || `Token request failed (${res.status}).`
    throw new TokenError(msg, {
      status: res.status,
      oauthError: typeof json.error === 'string' ? json.error : null
    })
  }
  return json
}

function exchangeCode({ def, tenant, clientId, code, verifier }) {
  return postForm(def.tokenPath(tenant), {
    grant_type: 'authorization_code',
    code,
    redirect_uri: REDIRECT_URI,
    client_id: clientId,
    code_verifier: verifier
  }).then((json) => toTokenSet(json))
}

// Exchange a stored refresh token for a fresh access token. Throws if the grant
// has been revoked/expired so the caller can drop to a disconnected state.
async function refreshTokens({ provider, clientId, tenant, refreshToken }) {
  if (!isProvider(provider)) {
    throw new Error(`Unknown provider: ${provider}`)
  }
  if (!refreshToken) {
    throw new Error('No refresh token available.')
  }
  const def = PROVIDERS[provider]
  const json = await postForm(def.tokenPath(tenant), {
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    client_id: clientId,
    scope: def.scope
  })
  return toTokenSet(json, refreshToken)
}

module.exports = {
  PROVIDERS,
  REDIRECT_URI,
  TokenError,
  CancelledError,
  authorize,
  refreshTokens,
  isProvider
}
