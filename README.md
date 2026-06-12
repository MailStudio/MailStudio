# Outlook Orbit

**Outlook and Teams, finally in one place.**

> **Beta** — Outlook Orbit is in active development. Things may break, selectors
> may drift, edges may be rough. Feedback, bug reports, and pull requests are
> very welcome — please open an issue.
>
> Developed with the help of [Claude Code](https://claude.com/claude-code) —
> thank you, Anthropic.

> Outlook Orbit is an **unofficial** client and is not affiliated with or
> endorsed by Microsoft or Asana. Outlook and Teams are trademarks of Microsoft
> Corporation; Asana is a trademark of Asana, Inc.

Outlook, Teams, Calendar, To Do, Asana — and your whole Office 365 suite — living
together in one fast, persistent window with real desktop notifications, a focus
timer, a scratchpad, and a sidebar that actually tells you
what's happening. No browser tabs. No context switching. One sign-in.

---

## Why it exists

Microsoft's web apps are excellent. The browser experience of _using_ them is not.
Outlook Orbit wraps everything in a dedicated macOS application that:

- Keeps every surface warm and logged in at all times
- Delivers native desktop notifications the moment mail or tasks arrive
- Puts a live unread count in your menu bar and Dock badge
- Lets you switch between Mail, Teams, Calendar, To Do, and Asana with a single
  keystroke or click — no hunting for tabs

If you've ever wished Microsoft made a real Mac app, this is that.

---

## What's inside

### Microsoft + Asana in one window

| Surface | What you get |
|---|---|
| **Outlook Mail** | Full OWA, live unread feed in the sidebar, compose shortcut |
| **Microsoft Teams** | Real Teams web app, native notifications via title watcher |
| **Calendar** | Full calendar, upcoming events in the sidebar, new-event shortcut |
| **Microsoft To Do** | Task lists alongside your mail |
| **Asana** | Your assigned tasks in the sidebar, powered by the Asana API |
| **Office 365 suite** | Word, Excel, PowerPoint, OneNote, OneDrive, Planner, SharePoint — all one click away, hidden by default and enabled in Settings |

One Microsoft sign-in carries across every Microsoft surface. Shared session — you
sign in once and Mail, Teams, Calendar, To Do, and the entire Office suite are all
authenticated.

### Live sidebar feeds

The sidebar isn't just navigation. It actively pulls in:

- **Unread emails** — sender, subject, and preview, refreshed every 25 seconds
  (faster on new mail via title change detection)
- **Upcoming calendar events** — today's agenda at a glance
- **Asana tasks** — your assigned work, with task permalinks

An aggregate strip at the top summarizes everything: _3 emails, 2 events, 4 tasks_.
Click it to jump straight to mail.

**Two sources, automatic switchover.** Before you connect an account, feeds and
notifications come from a built-in scraper that reads the logged-in web views —
best-effort, but it covers unread mail, today's events (including pre-start
reminders), and newly assigned Asana tasks. The moment you connect Microsoft or
Asana (see [API setup](#api-setup--byo-credentials)), the scrapers for that
provider's feeds **disable completely** and the Graph/Asana APIs become the only
source of truth — more reliable, richer data, and polling continues even while
the window is hidden. Disconnect and the scraper fallback takes over again.

> **Known limitation:** the scraper fallback parses Microsoft's web UI and
> currently assumes an **English** interface language for unread detection and
> date stamps. Non-English Microsoft accounts get degraded scrape feeds — the
> API path (once connected) is locale-independent and unaffected.

### Real notifications — not browser alerts

Once connected, notifications arrive from the **Microsoft Graph API** and
**Asana API**, which means they fire even when the window is hidden (before
connecting, the scraper fallback drives them while the app is open). Each
feature is independently toggleable:

- Per-email notifications with sender, subject, and message preview
- Batching: 4+ emails collapse into one summary notification instead of a flood
- **Quiet hours** — set a start and end time; notifications go silent overnight
- **Per-service snooze** — right-click any sidebar tab to snooze that service's
  alerts without affecting anything else
- Click a notification to open the exact email or task, not just the app

### Token security — OS keychain sandboxing

Credentials never touch disk in plaintext. Orbit uses Electron's `safeStorage` API
to encrypt tokens with the OS keychain before writing them to disk. If encryption
is unavailable (rare), tokens are kept in memory for the session only and discarded
on quit. Tokens refresh automatically; sign in once and stay signed in.

### Session isolation for custom sites

Built-in Microsoft and Asana services share one persistent session — that's how SSO
works. Every **custom pinned site** you add gets its own isolated session partition.
An untrusted pinned site cannot read cookies, storage, or credentials from your
Microsoft or Asana surfaces. Period.

### Collapsible sidebar — two modes

| Mode | What it does |
|---|---|
| **Vanish** (default) | Sidebar collapses to zero width, full screen for the active service |
| **Rail** | Collapses to a 76px icon strip with unread badges visible at all times |

Toggle with the collapse button or a keyboard shortcut. The sidebar springs back to
its full 280px panel on demand.

### Compose shortcuts

| Action | Shortcut |
|---|---|
| New email | ⌘N |
| New calendar event | ⌘⇧E |
| New task | ⌘⇧T |

Compose buttons also live in the sidebar for one-click access. Each one clicks the
native "new item" button inside the target surface and falls back to a deep-link
compose URL if the tab isn't ready.

### Focus timer (Pomodoro)

A 25-minute countdown ring lives at the bottom of the sidebar. Hit play to start a
focus session — the ring fills as time passes and the label switches to _Focusing_.
Pause, reset, repeat. No external app needed.

### Persistent scratchpad

A small text area in the sidebar that survives app restarts. Stored in your
settings file via a dedicated IPC channel that doesn't reload the views on every
keystroke. Use it for quick notes, draft subject lines, or anything you want
close to your inbox without opening a separate app.

### Custom pinned sites

Add any URL as a pinned tab alongside your Microsoft and Asana services. Each pin
gets its own isolated session (see above), its own icon, and appears in the
sidebar. Manage, reorder, and remove pins in the Settings page.

### Settings page — no separate preferences window

Open Settings from the sidebar. From there you can:

- Show or hide any built-in service tab
- Enable Office suite tabs (Word, Excel, PowerPoint, OneDrive, etc.)
- Add, remove, and reorder custom pinned sites
- Drag to reorder all tabs
- Switch sidebar collapse mode (vanish vs. rail)
- Configure notification toggles and quiet hours
- Connect Microsoft and Asana accounts

### Split view

⌘-click (Ctrl-click on Windows) a second tab to view two services side by side —
Mail next to Calendar, or Asana next to To Do. Click any tab normally to return to
a single view.

Drag the divider between the panes to resize them. The split-layout button in the
top bar (before the search box) toggles between side-by-side and stacked layouts;
it's enabled whenever a split is active.

### Always there, never in the way

The red close button hides the window to the menu bar tray — it does not quit the
app. The tray icon shows your live unread count and a menu with quick actions. The
Dock badge mirrors the mail unread count. Left-click the tray icon to open the
quick-access menu; right-click for native controls.

---

## Security model

Orbit was built with the assumption that it's wrapping trusted-but-complex
Microsoft surfaces alongside potentially untrusted custom pins.

- `contextIsolation: true`, `sandbox: true`, `nodeIntegration: false` on every
  web view
- Allowlist of trusted base domains (microsoft.com, office.com, sharepoint.com,
  teams.microsoft.com, asana.com, etc.) — suffix-matched, not prefix
- All external navigation checked against the allowlist; unknown hosts open in
  the default browser
- Custom pins isolated in per-site session partitions
- `file://` and `javascript:` navigation blocked; only `http://`, `https://`, and
  `mailto:` may open externally
- Panel and menu renderer windows hardened against navigation and `window.open`
- Feed content (email subjects, sender names) escaped before DOM insertion to
  prevent XSS from scraped content
- Tokens encrypted in OS keychain via `safeStorage`; never written in clear

To report a vulnerability privately, see [SECURITY.md](SECURITY.md).

### Privacy

- **No telemetry, no analytics, no third-party servers.** Your account data
  flows directly between your machine and Microsoft / Asana — nothing else
  sees it.
- Mail senders, subjects, and previews are held in memory for the sidebar and
  shown in OS notifications (preview text is toggleable in Settings).
- OAuth tokens are encrypted with the OS keychain (`safeStorage`). Settings and
  the scratchpad are stored locally as plaintext JSON in the app's user-data
  folder — don't keep secrets in the scratchpad.
- The only other network contact is the update check against GitHub Releases
  (packaged builds only, on launch and every six hours).

---

## API setup — BYO credentials

Because Orbit is open source, no API keys are baked in. Each person supplies their
own free OAuth app registrations (a one-time ~5-minute setup). You paste the
client IDs into **Connect accounts → Developer setup** inside the app — no code
editing. The app uses OAuth2 authorization-code + **PKCE**, which means no client
secret is ever needed or stored.

### Microsoft (Mail + Calendar)

1. [Azure Portal](https://portal.azure.com) → **App registrations** → **New registration**
2. Name it (e.g. "Outlook Orbit"); set supported account types to fit your setup
3. **Redirect URI**: platform _Mobile and desktop applications_ → `http://localhost/orbit-auth`
4. **API permissions** → Microsoft Graph → Delegated → `Mail.Read`, `Calendars.Read`, `offline_access`, `User.Read`
5. Copy the **Application (client) ID** → paste it into the app under **Microsoft → Developer setup**

### Asana (Tasks)

1. [Asana → My apps](https://app.asana.com/0/my-apps) → **Create new app**
2. Add OAuth redirect URL: `http://localhost/orbit-auth`
3. Copy the **Client ID** → paste it into the app under **Asana → Developer setup**

After setup, one Microsoft sign-in authenticates the entire Microsoft suite. One
Asana sign-in does the same for tasks. Refresh tokens keep you signed in
indefinitely.

> Teams notifications come from the lightweight title watcher rather than the
> Graph API. Graph Teams requires protected scopes, admin consent, and a public
> webhook endpoint — not viable for a desktop app that runs locally.

---

## Getting started

```bash
npm install
npm start          # run in development
npm run check      # syntax-check every source file
npm test           # run the unit tests (settings normalization + URL sanitization)
npm run dist:mac   # build a .dmg and .zip for macOS
npm run dist       # build for the current platform (mac / win / linux)
```

Continuous integration runs `npm run check` and `npm test` on every push and pull
request via [`.github/workflows/ci.yml`](.github/workflows/ci.yml).

### Cross-platform builds

Outlook Orbit builds for macOS, Windows, and Linux:

| Platform | Output | Notes |
|---|---|---|
| **macOS** | `.dmg`, `.zip` | Ad-hoc signed (see Gatekeeper note below) |
| **Windows** | `.exe` (NSIS installer) | Build on Windows |
| **Linux** | `.AppImage`, `.deb` | Build on Linux |

electron-builder packages for the host OS, so build each platform's artifact on
that platform (or in CI). The macOS-only icon/signing step in `scripts/after-pack.js`
automatically no-ops on Windows and Linux.

### Auto-update

Packaged builds check **GitHub Releases** for updates on launch and every six
hours, download new versions in the background, and post a native notification
prompting a restart (handled by [`src/main/updater.js`](src/main/updater.js) via
`electron-updater`). There's also a **Check for Updates…** item in the app menu.

To publish an update:

1. Bump `version` in `package.json` and push a matching tag (e.g. `v0.9.1`).
   [`release.yml`](.github/workflows/release.yml) builds the macOS, Windows,
   and Linux artifacts plus the update metadata (`latest*.yml`) on the tag and
   uploads everything to a draft GitHub Release.
2. Publish the release. Installed apps pick it up on their next check.

(Manual alternative: `npm run release` with a `GH_TOKEN` that can create
releases on the repo configured under `build.publish` — builds the current
platform only.)

> Auto-update is a no-op in development (unpackaged). On macOS, Squirrel.Mac only
> applies updates to a **notarized** build — until the app is notarized, macOS
> users update by downloading the new `.dmg`. Windows and Linux auto-update work
> with the current ad-hoc/unsigned builds.

### Installing on macOS (Gatekeeper)

The build is code-signed but not notarized. On first launch macOS will block it.

**Option A — System Settings (macOS 13+):**
1. Drag **Outlook Orbit** to `/Applications` and double-click (it will be blocked)
2. Open **System Settings → Privacy & Security** → **Open Anyway**
3. Launch again and confirm — remembered after that

**Option B — Terminal:**
```bash
xattr -dr com.apple.quarantine "/Applications/Outlook Orbit.app"
```

### Signing in to Microsoft

Microsoft may prompt for a passkey or security key. Inside Orbit:

- **Touch ID / platform passkey**: not supported — Chromium requires Apple's
  restricted browser entitlement, which is only granted to notarized browsers
- **Hardware security keys (YubiKey, etc.)**: supported
- **Recommended**: use the back arrow or _"Sign in another way"_ to switch to
  Microsoft Authenticator push or a one-time code — both render in-page and work
  fully

Because all Microsoft tabs share one persistent session, you sign in **once** and
everything — Teams, Mail, Calendar, To Do, and the Office suite — is authenticated.

---

## Project layout

```
src/
  main/
    main.js              — window management, BrowserViews, IPC, notifications
    settings-store.js    — persistent settings (services, theme, scratch, notif)
    secure-store.js      — safeStorage vault for OAuth tokens
    oauth.js             — PKCE engine + provider definitions
    api-feeds.js         — Graph + Asana API fetchers
    connections.js       — token lifecycle + feed accessors
    updater.js           — GitHub Releases auto-update (electron-updater)
    panel-preload.js     — sandboxed bridge between panel renderer and main
    service-preload.js   — title/URL reporter for web view tabs
  renderer/
    panel.html/css/js    — sidebar UI (tabs, feeds, settings, focus, scratchpad)
    menu.html/css/js     — tray dropdown menu
assets/
  icon.png
  trayTemplate.png       — menu bar icon (inactive)
  trayUnreadTemplate.png — menu bar icon (unread)
```

---

## Status & feedback

Outlook Orbit is in **beta**. It works, it's used daily, and it will also have
rough edges — Microsoft and Asana ship UI changes that can break the scraper
fallback (the API paths are immune), and notarization is still on the roadmap.
If something breaks or you have an idea, please open an issue. Feedback of any
kind is genuinely welcome.

## Acknowledgements

Developed with the help of [Claude Code](https://claude.com/claude-code) —
thank you, Anthropic.

## License

MIT
