# Capturing the hero screenshot

The README hero image should be a **real** capture of MailStudio wrapping Outlook —
with **no real account data**. Two ways to do it; pick one, save the result as
`docs/assets/screenshot-overview.png`, then uncomment the hero `<img>` block in
`README.md`.

Capture on a Retina display, window ~1200–1440px wide, dark theme.

---

## Option A — Outlook loading splash (easiest, zero data)

The cleanest data-free shot: catch the Outlook web app while it's still loading.

1. `npm start`, open the **Mail** tab.
2. Press **⌘R** to reload, and immediately take the shot during the ~1s Outlook
   splash (logo + spinner) — no inbox content is visible yet.
3. macOS: **⌘⇧4 → Space → click the window** (captures with rounded corners +
   shadow). Save to `docs/assets/screenshot-overview.png`.

This shows the real MailStudio chrome (top bar + sidebar) around the genuine
Outlook surface, with nothing private on screen.

---

## Option B — Demo data injected via DevTools (richest look)

Show a populated inbox with fake data.

1. Open DevTools on the **mail** web view. Quickest: temporarily add this line in
   `src/main/main.js` inside `createServiceView`, right after the view is built,
   then `npm start` (remove it afterward):
   ```js
   if (service.key === 'mail') view.webContents.openDevTools({ mode: 'detach' })
   ```
2. Sign in (ideally a throwaway/test account), open **Mail**, let the inbox render.
3. In the detached DevTools **Console** (it targets the Outlook page), paste:

   ```js
   (() => {
     const demo = [
       ['Alex Morgan', 'Q3 roadmap review — agenda', 'Sharing the deck ahead of Thursday'],
       ['Design Team', 'Weekly standup notes', 'Recap and action items from this week'],
       ['GitHub', 'Build passed on main', 'All checks have completed successfully'],
       ['Notion', 'Your weekly digest', "Here's what happened in your workspace"],
       ['Figma', '3 new comments on Mockups', 'Reviewers left feedback on your frames'],
       ['Stripe', 'Your monthly receipt', 'Thanks — your invoice is attached'],
     ];
     const rows = [...document.querySelectorAll('div[role="option"]')].slice(0, demo.length);
     rows.forEach((row, i) => {
       const leaves = [...row.querySelectorAll('*')]
         .filter(el => el.children.length === 0 && el.textContent.trim());
       const [sender, subject, preview] = demo[i];
       if (leaves[0]) leaves[0].textContent = sender;
       if (leaves[1]) leaves[1].textContent = subject;
       if (leaves[2]) leaves[2].textContent = preview;
     });
     // Scrub identity in the top bar (name, email, initials).
     document.querySelectorAll('[aria-label*="@"], [aria-label*="Account manager" i]')
       .forEach(el => el.setAttribute('aria-label', 'Demo User'));
   })();
   ```

   OWA's DOM shifts over time — if a field doesn't swap, tweak the selectors. Leave
   the reading pane on its empty "select an item" placeholder so no real message
   body shows (or open a demo row only after re-running the snippet).
4. Screenshot the window (**⌘⇧4 → Space → click**) → `docs/assets/screenshot-overview.png`.
5. **Remove the temporary `openDevTools` line.**

---

## Optional extra shots

The same approach captures the split view, the focus timer, the downloads drawer,
and a notification. Save as `docs/assets/screenshot-<name>.png` and reference them
in `docs/features.md`.

## Wire it into the README

Uncomment the hero block near the top of `README.md`:

```html
<p align="center">
  <img src="docs/assets/screenshot-overview.png" alt="MailStudio main window: live sidebar feeds beside the embedded Outlook web app" width="100%" />
</p>
```
