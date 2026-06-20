'use strict'

const { after, test } = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')

const TEST_USER_DATA = path.join(__dirname, '..', 'tooling', 'qa', '.tmp-settings-store-test')
fs.rmSync(TEST_USER_DATA, { recursive: true, force: true })
fs.mkdirSync(TEST_USER_DATA, { recursive: true })
after(() => {
  fs.rmSync(TEST_USER_DATA, { recursive: true, force: true })
})

// settings-store requires electron's `app` (only used for the on-disk settings
// path). Inject a minimal stub into the require cache so the module loads under
// a plain `node --test` run, outside the Electron runtime.
const electronPath = require.resolve('electron')
require.cache[electronPath] = {
  id: electronPath,
  filename: electronPath,
  loaded: true,
  exports: { app: { getPath: () => TEST_USER_DATA } }
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

test('connections never persist an Asana client secret in plaintext', () => {
  // The secret is routed to the encrypted vault by main.js; normalize() must
  // strip it so it can never leak into the plaintext settings file even if a
  // caller (or a tampered file) includes one.
  const settings = store.normalize({
    connections: {
      asana: { clientId: 'xyz', clientSecret: 'super-secret-value' }
    }
  })
  assert.equal(settings.connections.asana.clientId, 'xyz')
  assert.equal('clientSecret' in settings.connections.asana, false)
  assert.equal(JSON.stringify(settings).includes('super-secret-value'), false)
})

test('notification toggles default on, explicit false respected', () => {
  const settings = store.normalize({ notif: { mail: false } })
  assert.equal(settings.notif.mail, false)
  assert.equal(settings.notif.calendar, true)
  assert.equal(settings.notif.preview, true)
})

test('quiet hours keep only valid HH:MM values', () => {
  const settings = store.normalize({
    notif: {
      quietStart: '8:05',
      quietEnd: '24:99'
    }
  })

  assert.equal(settings.notif.quietStart, '08:05')
  assert.equal(settings.notif.quietEnd, '')
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

test('saved service feed kinds are constrained to known feed providers', () => {
  const settings = store.normalize({
    services: [
      { key: 'mail', url: 'https://evil.example', feed: 'asana' },
      { key: 'docs', label: 'Docs', url: 'https://example.com/wiki', feed: 'javascript:alert(1)' },
      { key: 'fake-mail', label: 'Fake Mail', url: 'https://example.com/mail', feed: 'mail', mailboxManaged: true },
      { key: 'shared', label: 'Shared', url: 'https://outlook.office.com/mail/shared/', feed: 'mail' }
    ]
  })

  assert.equal(settings.services.find((s) => s.key === 'mail').feed, 'mail')
  assert.equal('feed' in settings.services.find((s) => s.key === 'docs'), false)
  assert.equal('feed' in settings.services.find((s) => s.key === 'fake-mail'), false)
  assert.equal('mailboxManaged' in settings.services.find((s) => s.key === 'fake-mail'), false)
  assert.equal('feed' in settings.services.find((s) => s.key === 'shared'), false)
})

test('tampered built-in service keys are canonicalized before URL pinning', () => {
  const settings = store.normalize({
    services: [
      { key: ' mail ', label: 'Mail', url: 'https://phish.example/login' },
      { key: 'calendar\u0000', label: 'Calendar', url: 'https://phish.example/calendar' }
    ]
  })

  const mail = settings.services.find((s) => s.key === 'mail')
  const calendar = settings.services.find((s) => s.key === 'calendar')
  assert.equal(mail.builtin, true)
  assert.equal(mail.url, 'https://outlook.office.com/mail/')
  assert.equal(calendar.builtin, true)
  assert.equal(calendar.url, 'https://outlook.office.com/calendar/')
  assert.equal(settings.services.filter((s) => s.key === 'mail').length, 1)
})

test('custom pinned services are capped on import', () => {
  const services = Array.from({ length: 40 }, (_v, i) => ({
    key: `custom-${i}`,
    label: `Custom ${i}`,
    url: `https://example-${i}.com/`
  }))
  const settings = store.normalize({ services })
  const custom = settings.services.filter((s) => !s.builtin && !s.mailboxManaged)

  assert.equal(custom.length, 30)
  for (const builtin of store.DEFAULT_SERVICES) {
    assert.ok(settings.services.some((s) => s.key === builtin.key), `missing built-in ${builtin.key}`)
  }
})

test('managed shared mailbox services may keep their mail feed', () => {
  const settings = store.normalize({
    services: [
      {
        key: 'shared',
        label: 'Shared',
        url: 'https://outlook.office.com/mail/shared/',
        feed: 'mail',
        mailboxManaged: true
      }
    ]
  })

  assert.equal(settings.services.find((s) => s.key === 'shared').feed, 'mail')
  assert.equal(settings.services.find((s) => s.key === 'shared').mailboxManaged, true)
})

test('managed shared mailbox services cannot route home outside Outlook', () => {
  const settings = store.normalize({
    services: [
      {
        key: 'shared',
        label: 'Shared',
        url: 'https://outlook.office.com/mail/shared/',
        home: 'https://example.com/phish',
        feed: 'mail',
        mailboxManaged: true
      }
    ]
  })
  const shared = settings.services.find((s) => s.key === 'shared')

  assert.equal(shared.mailboxManaged, true)
  assert.equal(shared.home, shared.url)
})

test('custom service key and label metadata are clamped', () => {
  const settings = store.normalize({
    services: [
      {
        key: `\u0000${'k'.repeat(120)}`,
        label: 'l'.repeat(120),
        url: 'https://example.com/wiki'
      }
    ]
  })
  const custom = settings.services.find((s) => s.url === 'https://example.com/wiki')

  assert.equal(custom.key.length, 80)
  assert.equal(custom.label.length, 80)
  assert.equal(custom.key.includes('\u0000'), false)
})

test('save tightens an existing settings file to owner-only permissions', () => {
  const target = path.join(TEST_USER_DATA, 'mailstudio-settings.json')
  fs.writeFileSync(target, '{}', { mode: 0o644 })

  store.save({ theme: 'light' })

  const mode = fs.statSync(target).mode & 0o777
  if (process.platform !== 'win32') assert.equal(mode, 0o600)
  assert.equal(JSON.parse(fs.readFileSync(target, 'utf8')).theme, 'light')
})
