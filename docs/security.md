# Security & privacy

MailStudio is built on the assumption that it's wrapping powerful, signed-in
accounts alongside potentially untrusted custom pins — so it's locked down by
default. To report a vulnerability privately, see [SECURITY.md](../SECURITY.md).

## Least-privilege access

Both providers are connected with the **minimum scopes** the features need. Mail,
calendar, and task access is read-only. Microsoft presence is the one writeable
scope: it lets MailStudio set or clear your Teams availability from the sidebar.
The tokens cannot send mail, edit files, change tasks, or delete anything.

| Provider | Scopes requested | Access |
|---|---|---|
| Microsoft | `User.Read`, `Mail.Read`, `Calendars.Read`, `Presence.ReadWrite` (+ `openid`, `profile`, `offline_access`) | mail/calendar read-only; presence set/reset |
| Asana | `users:read`, `tasks:read`, `workspaces:read` (+ `openid`, `email`, `profile`) | read-only |

The "New task" / compose actions never use these tokens — they drive each service's
own web UI in its tab, using that surface's normal session.

## No secrets, no servers

- **PKCE OAuth** — the authorization-code + PKCE flow means **no client secret is
  shipped in MailStudio**. Microsoft needs no secret at all. Asana's API requires
  a client secret, so you supply your own (from your own free app registration);
  it's sealed in the OS keychain via `safeStorage`, never written in plaintext and
  never embedded in the repo.
- **No telemetry, no analytics, no third-party backend.** Your account data flows
  directly between your machine and Microsoft / Asana — nothing else sees it. The
  only other network contact is the update check against GitHub Releases (packaged
  builds, on launch and every six hours).

## Encrypted token storage

Credentials never touch disk in plaintext. MailStudio seals tokens with the OS
keychain via Electron's `safeStorage` (macOS Keychain / Windows DPAPI / libsecret)
before writing them. If encryption is unavailable (rare), tokens are kept in memory
for the session only and discarded on quit. The token file is written owner-only
(`0600`) with a symlink-safe atomic write.

## Hard session isolation

- Microsoft services share **one** persistent SSO partition — that's how one
  sign-in unlocks Outlook, Teams, Calendar, To Do, and the Office suite.
- Asana runs in its **own** partition (it doesn't need Microsoft SSO).
- Every **custom pinned site** gets its **own isolated partition**. An untrusted
  pin cannot read cookies, storage, tokens, or credentials from your Microsoft or
  Asana surfaces.

## Sandboxed, allowlisted web views

- `contextIsolation: true`, `sandbox: true`, `nodeIntegration: false` on every view
- Suffix-matched allowlist of trusted base domains (microsoft.com, office.com,
  cloud.microsoft, sharepoint.com, asana.com, ...)
- All navigation checked against the allowlist; trusted links route to their owning
  tab in-app, unknown hosts open in the default browser
- `file://` and `javascript:` navigation blocked — only `http://`, `https://`, and
  `mailto:` may open externally
- A **native confirmation** before any cross-app link is loaded into the
  authenticated Teams session
- Web-view permission requests restricted to a small allowlist: notifications and
  clipboard write on trusted app hosts, plus camera/microphone media only for
  Teams meetings/calls. Geolocation, screen capture, USB, etc. are denied
- Auth popups (Microsoft login/MFA) are only permitted from Microsoft service tabs,
  never from custom pins
- Panel and menu renderer windows hardened against navigation and `window.open`;
  IPC commands are accepted only from the app's own windows
- Feed content (email subjects, sender names, task names) is escaped before DOM
  insertion to prevent XSS from scraped or API content

## Privacy notes

- Mail senders, subjects, and previews are held in memory for the sidebar and shown
  in OS notifications (preview text is toggleable in Settings).
- Settings and the scratchpad are stored locally as plaintext JSON (owner-only) in
  the app's user-data folder — **don't keep secrets in the scratchpad.**
- OAuth tokens are the only credentials stored, and they're encrypted as above.
