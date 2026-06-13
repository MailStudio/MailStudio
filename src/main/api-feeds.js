// API-backed feed fetchers. Each returns the SAME { state, items } shape the
// sidebar renderer already draws for scraped feeds, so swapping the data source
// is transparent to the UI. `state` is one of: 'ok' | 'empty' | 'auth' | 'error'.
//
// Items intentionally mirror the scraped item fields, plus a `webLink` so a
// click can open the real item in the owning web view:
//   mail:     { id, sender, subject, preview, webLink }
//   calendar: { id, title, time, webLink, cancelled? }
//   asana:    { id, name, subtasks, taskUrl }

const GRAPH = 'https://graph.microsoft.com/v1.0'
const ASANA = 'https://app.asana.com/api/1.0'

// A 401 means the access token is stale/revoked — surfaced as a typed error so
// the caller can refresh-and-retry (or drop the provider to disconnected).
class AuthError extends Error {}

// A 429 (or 503 carrying a Retry-After header) means the API wants us to back
// off — surfaced with the requested pause so the caller can cool down without
// touching connection state.
class ThrottledError extends Error {
  constructor(message, retryAfter) {
    super(message)
    this.name = 'ThrottledError'
    this.retryAfter = retryAfter
  }
}

// Retry-After in seconds; default to 60 when the header is missing/unparseable.
function parseRetryAfter(res) {
  const secs = Number(res.headers.get('retry-after'))
  return Number.isFinite(secs) && secs > 0 ? secs : 60
}

async function apiGet(url, accessToken, extraHeaders) {
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}`, Accept: 'application/json', ...(extraHeaders || {}) },
    signal: AbortSignal.timeout(15000)
  })
  if (res.status === 401) {
    throw new AuthError('Unauthorized')
  }
  if (res.status === 429 || (res.status === 503 && res.headers.get('retry-after'))) {
    throw new ThrottledError(`Throttled (${res.status})`, parseRetryAfter(res))
  }
  const text = await res.text()
  let json
  try {
    json = text ? JSON.parse(text) : {}
  } catch {
    json = {}
  }
  if (!res.ok) {
    const msg = (json.error && (json.error.message || json.error)) || `Request failed (${res.status})`
    throw new Error(typeof msg === 'string' ? msg : 'Request failed')
  }
  return json
}

/* ---------- Microsoft Graph: Mail ---------- */
async function fetchMail(accessToken) {
  // Unread inbox, newest first. $select keeps the payload small.
  const url =
    `${GRAPH}/me/mailFolders/inbox/messages` +
    `?$filter=isRead eq false&$top=10&$orderby=receivedDateTime desc` +
    `&$select=id,subject,bodyPreview,from,receivedDateTime,webLink`
  const json = await apiGet(url, accessToken)
  const rows = Array.isArray(json.value) ? json.value : []
  const items = rows.map((m) => {
    const addr = m.from && m.from.emailAddress ? m.from.emailAddress : {}
    return {
      id: m.id,
      sender: (addr.name || addr.address || 'Unknown').slice(0, 80),
      subject: (m.subject || '(no subject)').slice(0, 120),
      preview: (m.bodyPreview || '').replace(/\s+/g, ' ').trim().slice(0, 160),
      // Raw Graph timestamp so the renderer can count "unread today" itself.
      receivedIso: m.receivedDateTime || null,
      webLink: m.webLink || null
    }
  })
  return { state: items.length ? 'ok' : 'empty', items }
}

// Total unread count for the inbox badge (cheap, separate from the list).
async function fetchMailUnreadCount(accessToken) {
  const json = await apiGet(`${GRAPH}/me/mailFolders/inbox?$select=unreadItemCount`, accessToken)
  return Number(json.unreadItemCount) || 0
}

/* ---------- Microsoft Graph: Calendar ---------- */
async function fetchCalendar(accessToken) {
  const now = new Date()
  const end = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000)
  const url =
    `${GRAPH}/me/calendarView` +
    `?startDateTime=${encodeURIComponent(now.toISOString())}` +
    `&endDateTime=${encodeURIComponent(end.toISOString())}` +
    `&$select=id,subject,start,end,isAllDay,webLink,isCancelled&$orderby=start/dateTime&$top=10`
  // Ask Graph to return start/end in the user's local zone for clean formatting.
  const tz = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC'
  const json = await apiGet(url, accessToken, { Prefer: `outlook.timezone="${tz}"` })
  const rows = Array.isArray(json.value) ? json.value : []
  const items = rows.map((e) => ({
    id: e.id,
    title: (e.subject || '(no title)').slice(0, 90),
    time: e.isAllDay ? 'All day' : formatEventTime(e.start && e.start.dateTime),
    // Raw local start string (no zone suffix when a Prefer tz is set) used for
    // upcoming-event reminder timing; all-day events get no reminder.
    startIso: e.isAllDay ? null : (e.start && e.start.dateTime) || null,
    webLink: e.webLink || null,
    cancelled: Boolean(e.isCancelled)
  }))
  return { state: items.length ? 'ok' : 'empty', items }
}

function formatEventTime(dateTime) {
  if (!dateTime) return ''
  // Graph returns 'YYYY-MM-DDTHH:mm:ss.0000000' with no zone suffix when a
  // Prefer timezone is set — treat it as local wall-clock time.
  const d = new Date(dateTime.replace(/(\.\d+)?$/, ''))
  if (Number.isNaN(d.getTime())) return ''
  const today = new Date()
  const tomorrow = new Date(today)
  tomorrow.setDate(today.getDate() + 1)
  const time = d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })
  if (d.toDateString() === today.toDateString()) return time
  if (d.toDateString() === tomorrow.toDateString()) return `Tomorrow ${time}`
  const day = d.toLocaleDateString([], { weekday: 'short' })
  return `${day} ${time}`
}

/* ---------- Asana ---------- */
// Returns { gid, name, workspaceGid } for the signed-in user — workspaceGid is
// needed to scope the task query and is cached on the stored account.
async function fetchAsanaMe(accessToken) {
  const json = await apiGet(`${ASANA}/users/me?opt_fields=name,email,workspaces.name`, accessToken)
  const data = json.data || {}
  const workspaces = Array.isArray(data.workspaces) ? data.workspaces : []
  return {
    gid: data.gid || null,
    name: data.name || data.email || 'Asana',
    workspaceGid: workspaces[0] ? workspaces[0].gid : null
  }
}

async function fetchAsanaTasks(accessToken, workspaceGid) {
  if (!workspaceGid) {
    return { state: 'error', items: [] }
  }
  // Incomplete tasks assigned to me in the primary workspace.
  const url =
    `${ASANA}/tasks?assignee=me&workspace=${encodeURIComponent(workspaceGid)}` +
    `&completed_since=now&opt_fields=name,permalink_url,due_on&limit=20`
  const json = await apiGet(url, accessToken)
  const rows = Array.isArray(json.data) ? json.data : []
  const items = rows
    .filter((t) => t.name && t.name.trim())
    .slice(0, 12)
    .map((t) => ({
      id: t.gid,
      name: t.name.slice(0, 110),
      subtasks: [],
      taskUrl: t.permalink_url || null,
      dueOn: t.due_on || null
    }))
  return { state: items.length ? 'ok' : 'empty', items }
}

module.exports = {
  AuthError,
  ThrottledError,
  fetchMail,
  fetchMailUnreadCount,
  fetchCalendar,
  fetchAsanaMe,
  fetchAsanaTasks
}
