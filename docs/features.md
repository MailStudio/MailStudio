# Feature tour

A deeper look at what MailStudio does. For the short pitch, see the
[README](../README.md).

## Microsoft + Asana in one window

| Surface | What you get |
|---|---|
| **Outlook Mail** | Full OWA, live unread feed in the sidebar, compose shortcut |
| **Microsoft Teams** | Real Teams web app, native notifications via title watcher |
| **Calendar** | Full calendar, upcoming events in the sidebar, new-event shortcut |
| **Microsoft To Do** | Task lists alongside your mail |
| **Asana** | Your assigned tasks in the sidebar, powered by the Asana API |
| **Office 365 suite** | Word, Excel, PowerPoint, OneNote, OneDrive, Planner, SharePoint — hidden by default, enabled in Settings |
| **Copilot** | Microsoft 365 Copilot, one click away |

One Microsoft sign-in carries across every Microsoft surface — sign in once and
Mail, Teams, Calendar, To Do, and the entire Office suite are all authenticated.

## Smart link routing

Click any Microsoft or Asana link, anywhere in MailStudio — an email, a Teams
chat, a sidebar feed item, a notification — and it opens **in the tab that owns
it**, not in your default browser:

- A Teams meeting link in an email → jumps to the **Teams** tab
- An Asana task permalink in a chat message → jumps to the **Asana** tab
- A shared Word/Excel/PowerPoint document → opens on its **app tab** (auto-revealed
  if hidden), falling back to **OneDrive/SharePoint**
- Calendar invites, To Do lists, Planner boards, OneNote pages, and Microsoft 365
  launch links → each lands on its respective tab
- Even `target="_blank"` popup links are intercepted and routed

The router recognizes the full constellation of Microsoft hosts —
`outlook.office.com`, `teams.microsoft.com`, `*.sharepoint.com`,
`onedrive.live.com`, `1drv.ms`, `planner.microsoft.com`, the new
`*.cloud.microsoft` app hosts, and more — plus `app.asana.com`.

Tabs you've hidden never get surfaced by a link unless the link is for them; only
genuinely external (non-Microsoft, non-Asana) links go to your default browser.
Opening a link into **Teams** from another tab triggers a native confirmation
first, so nothing loads into your authenticated Teams session without your say-so.

## Live sidebar feeds

The sidebar actively pulls in:

- **Unread emails** — sender, subject, preview, and relative timestamps, refreshed
  every 25 seconds (faster on new mail via title-change detection)
- **Upcoming calendar events** — today/tomorrow's agenda at a glance
- **Asana tasks** — your assigned work with due dates and task permalinks

An aggregate strip summarizes everything — _3 emails, 2 events, 4 tasks_ — and
jumps you to mail on click.

**Two sources, automatic switchover.** Before you connect an account, feeds and
notifications come from a built-in scraper that reads the logged-in web views —
best-effort, covering unread mail, today's events, and newly assigned Asana tasks.
The moment you connect Microsoft or Asana (see [setup](setup.md)), the scrapers
for that provider's feeds **disable completely** and the Graph/Asana APIs become
the only source of truth — more reliable, richer data, and polling continues even
while the window is hidden. Disconnect and the scraper fallback takes over again.

> **Known limitation:** the scraper fallback parses Microsoft's web UI and assumes
> an **English** interface for unread detection and date stamps. Non-English
> accounts get degraded scrape feeds — the API path (once connected) is
> locale-independent and unaffected.

## Real notifications — not browser alerts

Once connected, notifications arrive from the **Microsoft Graph API** and **Asana
API**, so they fire even when the window is hidden (before connecting, the scraper
drives them while the app is open). Each feature is independently toggleable:

- Per-email notifications with sender, subject, and message preview
- Batching: 4+ emails collapse into one summary instead of a flood
- Calendar reminders shortly before an event starts
- **Quiet hours** — set a start/end time; notifications go silent overnight
- **Per-service snooze** — right-click any sidebar tab to snooze just that service
- Click a notification to open the exact email or task, not just the app

## Downloads, your way

Files downloaded from any tab funnel through one place: a **Downloads** button in
the top bar opens a drawer with live progress, speed, and color-coded file types.
Every download uses the **OS-native Save dialog**, so you choose exactly where each
file lands, and a notification on completion reveals it in Finder/Explorer.

## Split view

⌘-click (Ctrl-click on Windows) a second tab to view two services side by side —
Mail next to Calendar, or Asana next to To Do. Drag the divider to resize; the
split-layout button toggles side-by-side vs. stacked. Click any tab normally to
return to a single view.

## Resizable, collapsible sidebar

Drag the right edge of the sidebar to any width. Collapse it two ways:

| Mode | What it does |
|---|---|
| **Vanish** (default) | Collapses to zero width — full screen for the active service |
| **Rail** | Collapses to a 76px icon strip with unread badges always visible |

## Focus timer + scratchpad

- **Focus timer** — a countdown ring (default 30 minutes; set any length with the
  duration input or −/+ steppers). The ring fills as time passes; finishing a
  session triggers a colorful particle burst from the window edges.
- **Scratchpad** — a small text area that survives restarts, stored locally via a
  dedicated channel that doesn't reload your views on every keystroke.

## Compose shortcuts

| Action | Shortcut |
|---|---|
| New email | ⌘N |
| New calendar event | ⌘⇧E |
| New task | ⌘⇧T |

Compose buttons also live in the sidebar. Each clicks the native "new item" button
inside the target surface, falling back to a deep-link compose URL if needed.

## Custom pinned sites

Add any URL as a pinned tab alongside Microsoft and Asana. Each pin gets its own
**isolated session** (it can't read your Microsoft/Asana cookies or storage), its
own icon, and a sidebar slot. Manage, reorder, and remove pins in Settings.

## Settings — no separate window

Open Settings from the sidebar to show/hide tabs, enable Office-suite apps, add and
reorder pinned sites, switch collapse mode, configure notifications and quiet
hours, and connect Microsoft and Asana accounts (⌘, opens it directly).

## Always there, never in the way

The red close button hides the window to the menu-bar tray — it doesn't quit. The
tray shows your live unread count with quick actions; the Dock badge mirrors it.
Left-click the tray to open the quick menu; right-click for native controls.
