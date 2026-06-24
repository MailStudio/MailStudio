# API setup — bring your own credentials

Because MailStudio is open source, no API keys are baked in. Each person supplies
their own free OAuth app registrations (a one-time ~5-minute setup per provider).
You paste the client IDs into **Settings → Connections → Developer setup** inside
the app — no code editing.

The app uses OAuth2 authorization-code + **PKCE**. **Microsoft** is a public
client, so it needs **no client secret at all**. **Asana** is different: its token
endpoint requires a **client secret even with PKCE**, so you paste your Asana app's
secret in too — it's sealed in your OS keychain (never the plaintext settings file,
never the repo). Both providers are requested with **least-privilege scopes** —
MailStudio can read your mail, calendar, and assigned tasks for the sidebar and
notifications. Microsoft also grants `Presence.ReadWrite` so the Teams status
control can set or clear your availability. The tokens cannot send mail, edit
files, modify tasks, or delete anything.

Until you connect an account, a built-in scraper drives the sidebar feeds and
notifications from the logged-in web views. The moment you connect, the official
APIs take over (more reliable, richer data, and they keep working while the window
is hidden).

Both providers share one redirect URI, which never actually loads — the app
intercepts the sign-in popup's navigation to it and reads the authorization code
off the URL:

```
http://localhost/mailstudio-auth
```

---

## Microsoft (Mail + Calendar)

1. **Register the app** — [Azure Portal](https://portal.azure.com) →
   **App registrations** → **New registration**.
2. **Name** it (e.g. "MailStudio").
3. **Supported account types** — choose **Accounts in any organizational
   directory and personal Microsoft accounts** for the simplest setup. If you
   pick **single tenant** (your org only), you must also paste your tenant ID
   into the app later (step 7).
4. **Redirect URI** — set the platform to **Mobile and desktop applications**
   *(not "Web" — Web requires a client secret, which this app deliberately
   avoids)* and enter:
   ```
   http://localhost/mailstudio-auth
   ```
   *(Already created the app without this? Open **Authentication → + Add a
   platform → Mobile and desktop applications** and add it there.)*
5. **API permissions** → **+ Add a permission** → **Microsoft Graph** →
   **Delegated permissions** → add all of:
   `openid`, `profile`, `offline_access`, `User.Read`, `Mail.Read`,
   `Calendars.Read`, `Presence.ReadWrite`. If your tenant shows a **Grant admin consent** button,
   click it.
6. **Authentication** → **Advanced settings** → set **Allow public client flows**
   to **Yes**.
7. **Copy the credentials into MailStudio** — from the app's **Overview** page:
   - **Application (client) ID** → **Microsoft → Developer setup → Azure client ID**
   - **Tenant ID** → leave as `common` for the multi-tenant/personal option above;
     for a single-tenant app, paste the **Directory (tenant) ID** into the
     **Tenant ID** field.

### Troubleshooting Microsoft

| Error | Fix |
|---|---|
| `AADSTS500113: No reply address is registered` | Redirect URI missing — add `http://localhost/mailstudio-auth` under **Mobile and desktop applications** (step 4). |
| "Sorry, but we're having trouble signing you in" | Usually missing Graph permissions or consent — complete step 5 (add delegated scopes, grant admin consent). |
| Account-type / tenant error on a single-tenant app | Paste your **Directory (tenant) ID** into the **Tenant ID** field instead of `common` (step 7). |

---

## Asana (Tasks)

1. **Create the app** — [Asana → My apps](https://app.asana.com/0/my-apps) →
   **Create new app**.
2. **Redirect URL** — in the app's **OAuth** section, add:
   ```
   http://localhost/mailstudio-auth
   ```
3. **Scopes** — MailStudio uses **granular, read-only** scopes (not "Full
   permissions"). In the **OAuth** section, enable read access for:
   - **Tasks** → read
   - **Users** → read
   - **Workspaces** → read
   - Identity: **openid**, **email**, **profile**
4. **Copy the credentials into MailStudio** — from the app's **Basic information**
   tab:
   - **Client ID** → **Asana → Developer setup → Asana client ID**
   - **Client secret** → **Asana → Developer setup → Asana client secret**

   Unlike Microsoft, **Asana's token endpoint requires the client secret even with
   PKCE**, so both fields are needed before Connect is enabled. The secret is
   stored encrypted in your OS keychain (via Electron `safeStorage`) — never in the
   plaintext settings file and never in this repo.

### Troubleshooting Asana

| Error | Fix |
|---|---|
| `The `client_id` and `client_secret` must authorize the app` | The client secret is missing or wrong. Paste your Asana app's **client secret** (Basic information tab) into **Developer setup → Asana client secret** (step 4). |
| `Your app does not have any scopes registered` | Enable the granular read scopes in step 3 (Tasks/Users/Workspaces read + identity). |
| `invalid_request: redirect_uri does not match` | Add `http://localhost/mailstudio-auth` exactly (lowercase, no trailing slash) in the OAuth section (step 2). |

---

After setup, one Microsoft sign-in authenticates the entire Microsoft suite, and
one Asana sign-in does the same for tasks. Refresh tokens keep you signed in
indefinitely.

> **Teams notifications** come from a lightweight title watcher rather than the
> Graph API. Graph Teams requires protected scopes, admin consent, and a public
> webhook endpoint — not viable for a desktop app that runs locally.
