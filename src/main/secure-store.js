const fs = require('fs')
const path = require('path')
const { app, safeStorage } = require('electron')

// Encrypted credential vault. OAuth access/refresh tokens are written here via
// Electron's safeStorage (backed by the macOS Keychain / Windows DPAPI / libsecret),
// NEVER to the plaintext settings JSON. If OS-level encryption is unavailable we
// fall back to keeping tokens in memory only for the session — refusing to ever
// persist secrets in the clear.
//
// Shape on disk (after decryption):
//   { microsoft: <tokenSet>|null, asana: <tokenSet>|null, secrets: { <name>: <string> } }
// where a tokenSet is { accessToken, refreshToken, expiresAt, scope, account }.
// `secrets` holds non-token credentials that still must never touch plaintext —
// notably the Asana OAuth client secret (Asana's token endpoint requires it even
// under PKCE, unlike Microsoft's public-client flow).

function vaultPath() {
  return path.join(app.getPath('userData'), 'mailstudio-credentials.bin')
}

function encryptionAvailable() {
  try {
    return safeStorage.isEncryptionAvailable()
  } catch {
    return false
  }
}

// In-memory mirror so reads are cheap and so we can still function for the
// session even when the OS keychain is unavailable (nothing hits disk then).
let cache = null

function emptyVault() {
  return { microsoft: null, asana: null, secrets: {} }
}

function load() {
  if (cache) {
    return cache
  }
  if (!encryptionAvailable()) {
    cache = emptyVault()
    return cache
  }
  try {
    const blob = fs.readFileSync(vaultPath())
    const json = safeStorage.decryptString(blob)
    const parsed = JSON.parse(json)
    cache = { ...emptyVault(), ...(parsed && typeof parsed === 'object' ? parsed : {}) }
  } catch {
    // Missing file, corrupt blob, or a vault written under a different OS key —
    // start clean rather than throwing on launch.
    cache = emptyVault()
  }
  return cache
}

function persist() {
  if (!encryptionAvailable()) {
    // No safe place to write — keep tokens in memory only.
    return false
  }
  try {
    const blob = safeStorage.encryptString(JSON.stringify(cache))
    const target = vaultPath()
    const tmp = target + '.tmp'
    fs.mkdirSync(path.dirname(target), { recursive: true })
    // Remove any pre-existing .tmp file before writing — a symlink planted at
    // this path by another process could otherwise redirect the write to an
    // attacker-controlled target. Unlink breaks the symlink; the subsequent
    // write creates a fresh, owned, 0o600 regular file.
    try { fs.unlinkSync(tmp) } catch { /* file may not exist */ }
    // Atomic write: never leave a torn/partial vault behind if we crash
    // mid-write. Owner-only permissions on the temp file from the start.
    fs.writeFileSync(tmp, blob, { mode: 0o600 })
    fs.renameSync(tmp, target)
    try {
      // Best-effort tighten in case an older vault was created more loosely.
      fs.chmodSync(target, 0o600)
    } catch {
      /* ignore */
    }
    return true
  } catch {
    return false
  }
}

// Warn (once per session) when tokens can't be persisted, so the silent
// memory-only fallback is at least visible in logs.
let warnedPersistFailure = false
function warnPersistFailed() {
  if (warnedPersistFailure) return
  warnedPersistFailure = true
  console.warn('[secure-store] could not persist tokens — keeping them in memory for this session only')
}

function getToken(provider) {
  const vault = load()
  return vault[provider] || null
}

function setToken(provider, tokenSet) {
  const vault = load()
  vault[provider] = tokenSet || null
  if (!persist()) warnPersistFailed()
  return vault[provider]
}

function clearToken(provider) {
  const vault = load()
  vault[provider] = null
  if (!persist()) warnPersistFailed()
}

// Encrypted, non-token credentials (e.g. an OAuth client secret). Stored in the
// same vault so they share safeStorage encryption and the memory-only fallback.
function getSecret(name) {
  const vault = load()
  const secrets = vault.secrets && typeof vault.secrets === 'object' ? vault.secrets : null
  return secrets && typeof secrets[name] === 'string' ? secrets[name] : null
}

function setSecret(name, value) {
  const vault = load()
  if (!vault.secrets || typeof vault.secrets !== 'object') vault.secrets = {}
  if (typeof value === 'string' && value) {
    vault.secrets[name] = value
  } else {
    delete vault.secrets[name]
  }
  if (!persist()) warnPersistFailed()
}

function hasSecret(name) {
  return Boolean(getSecret(name))
}

function clearAll() {
  cache = emptyVault()
  try {
    fs.unlinkSync(vaultPath())
  } catch {
    /* file may not exist */
  }
}

module.exports = {
  encryptionAvailable,
  getToken,
  setToken,
  clearToken,
  getSecret,
  setSecret,
  hasSecret,
  clearAll
}
