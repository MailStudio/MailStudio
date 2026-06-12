'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')

// settings-store requires electron's `app` (only used for the on-disk settings
// path). Inject a minimal stub into the require cache so the module loads under
// a plain `node --test` run, outside the Electron runtime.
const electronPath = require.resolve('electron')
require.cache[electronPath] = {
  id: electronPath,
  filename: electronPath,
  loaded: true,
  exports: { app: { getPath: () => '/tmp/outlook-orbit-test' } }
}

const store = require('../src/main/settings-store')

test('normalize(null) produces a fresh-install config', () => {
  const settings = store.normalize(null)
  assert.equal(settings.firstBoot, true)
  assert.equal(settings.onboarded, false)
  // Every built-in service is present even with no saved services.
  const keys = settings.services.map((s) => s.key)
  for (const builtin of store.DEFAULT_SERVICES) {
    assert.ok(keys.includes(builtin.key), `missing built-in ${builtin.key}`)
  }
})

test('normalize coerces invalid theme and collapseMode to defaults', () => {
  const settings = store.normalize({ theme: 'rainbow', collapseMode: 'fold' })
  assert.equal(settings.theme, store.DEFAULTS.theme)
  assert.equal(settings.collapseMode, store.DEFAULTS.collapseMode)
})

test('taskProvider defaults to microsoft and rejects unknown values', () => {
  assert.equal(store.normalize(null).taskProvider, 'microsoft')
  assert.equal(store.normalize({ taskProvider: 'asana' }).taskProvider, 'asana')
  assert.equal(store.normalize({ taskProvider: 'jira' }).taskProvider, 'microsoft')
})

test('existing config is not flagged as first boot', () => {
  const settings = store.normalize({ firstBoot: false, onboarded: true })
  assert.equal(settings.firstBoot, false)
  assert.equal(settings.onboarded, true)
})

test('scratch is clamped to 20000 chars and non-strings dropped', () => {
  const long = 'x'.repeat(25000)
  assert.equal(store.normalize({ scratch: long }).scratch.length, 20000)
  assert.equal(store.normalize({ scratch: { not: 'a string' } }).scratch, '')
})

test('connections trim client IDs and default the Microsoft tenant', () => {
  const settings = store.normalize({
    connections: {
      microsoft: { clientId: '  abc-123  ' },
      asana: { clientId: ' xyz ' }
    }
  })
  assert.equal(settings.connections.microsoft.clientId, 'abc-123')
  assert.equal(settings.connections.microsoft.tenant, 'common')
  assert.equal(settings.connections.asana.clientId, 'xyz')
})

test('notification toggles default on, explicit false respected', () => {
  const settings = store.normalize({ notif: { mail: false } })
  assert.equal(settings.notif.mail, false)
  assert.equal(settings.notif.calendar, true)
  assert.equal(settings.notif.preview, true)
})

// --- URL sanitization (exercised through the public normalize() surface) ---

test('custom pinned site with a dangerous protocol is dropped', () => {
  for (const url of ['javascript:alert(1)', 'file:///etc/passwd', 'data:text/html,<x>']) {
    const settings = store.normalize({ services: [{ key: 'evil', url }] })
    assert.ok(!settings.services.some((s) => s.key === 'evil'), `accepted ${url}`)
  }
})

test('custom pinned site with an http(s) URL is kept', () => {
  const settings = store.normalize({
    services: [{ key: 'docs', label: 'Docs', url: 'https://example.com/wiki' }]
  })
  const docs = settings.services.find((s) => s.key === 'docs')
  assert.ok(docs, 'expected the custom site to be kept')
  assert.equal(docs.label, 'Docs')
  assert.match(docs.url, /^https:\/\/example\.com\/wiki/)
  assert.equal(docs.builtin, false)
})

test('a pinned site with no URL is rejected', () => {
  const settings = store.normalize({ services: [{ key: 'bad' }] })
  assert.ok(!settings.services.some((s) => s.key === 'bad'))
})
