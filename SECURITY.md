# Security Policy

Outlook Orbit stores OAuth tokens for your Microsoft and Asana accounts
(encrypted via the OS keychain through Electron's `safeStorage`) and renders
your mailbox inside embedded web views, so we take security reports seriously.

## Reporting a vulnerability

Please **do not** open a public issue for security problems.

Instead, report privately via
[GitHub Security Advisories](https://github.com/MailStudio/MailStudio/security/advisories/new)
("Report a vulnerability" on the repository's Security tab).

Include:

- What you found and where (file/line if you have it)
- Steps to reproduce or a proof of concept
- Impact as you understand it (e.g. token exposure, navigation of an
  authenticated view to an untrusted origin, XSS in the panel renderer)

You should hear back within a week. Since this is a beta maintained in spare
time, fixes for confirmed issues are prioritized over all other work.

## Scope notes

- The app never transmits your data anywhere except Microsoft, Asana, and
  GitHub (update checks) — see the Privacy section of the README.
- Custom pinned sites run in isolated session partitions by design; reports
  about a pinned site reading *its own* partition are not vulnerabilities.
- The DOM-scraper fallback parses untrusted page content; escaping bugs in how
  scraped text reaches the sidebar UI are in scope and very welcome.
