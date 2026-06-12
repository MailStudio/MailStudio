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
//   { microsoft: <tokenSet>|null, asana: <tokenSet>|null }
// where a tokenSet is { accessToken, refreshToken, expiresAt, scope, account }.

function vaultPath() {
  return path.join(app.getPath('userData'), 'orbit-credentials.bin')
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
  return { microsoft: null, asana: null }
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
  clearAll
}
