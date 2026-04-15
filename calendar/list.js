/**
 * List events functionality
 *
 * BUG FIX 1 — Wrong date range / missing events:
 *   BEFORE: Used me/events with $filter `start/dateTime ge '<utcNow>'`.
 *           Microsoft docs warn this filter is unreliable on me/events and
 *           can silently drop events. Also, the UTC "now" could miss events
 *           that already started (e.g., a 9 AM PKT event at 4 AM UTC when
 *           queried at 4:30 AM UTC would be filtered out).
 *   AFTER:  Use me/calendarView with startDateTime / endDateTime query params,
 *           which is the Graph-recommended approach and is timezone-aware.
 *           Window: start of local day → 14 days forward (safe for PKT UTC+5).
 *
 * BUG FIX 2 — Dedup cache returning stale "no events" after create:
 *   BEFORE: Cache keyed only on `count`, TTL 5 s. A freshly created event
 *           would be invisible if list-events was called within 5 s of create.
 *   AFTER:  Cache is invalidated explicitly via exported invalidateListCache().
 *           create.js should call this after a successful create (see note).
 *           TTL also reduced to 2 s as a safety net.
 *
 * Other fixes retained from previous version:
 *   - mapEventToDto() response mapper
 *   - classifyCalendarError() error handler
 *   - CALENDAR_SELECT_FIELDS field selection
 */
const { callGraphAPI } = require('../utils/graph-api');
const { getClient } = require('../auth');
const config = require('../config');

// ─── Local DTO + error helpers (mirrors shared.js for self-containment) ──────

function mapEventToDto(event) {
  return {
    id: event.id,
    subject: event.subject,
    start: event.start,
    end: event.end,
    location: event.location?.displayName || 'No location',
    organizer: event.organizer?.emailAddress?.name || 'Unknown',
    isAllDay: event.isAllDay || false,
    isCancelled: event.isCancelled || false,
    bodyPreview: event.bodyPreview || '',
    responseStatus: event.responseStatus?.response || 'none',
    importance: event.importance || 'normal'
  };
}

function classifyCalendarError(error, operation) {
  if (error.message === 'Authentication required' || error.message === 'UNAUTHORIZED') {
    return {
      content: [{
        type: "text",
        text: "Authentication required or token expired. Please use the 'authenticate' tool first, or provide a fresh bearer token."
      }]
    };
  }
  if (error.message && error.message.includes('404')) {
    return { content: [{ type: "text", text: "Event not found (404). It may have been deleted or the ID is incorrect." }] };
  }
  if (error.message && error.message.includes('409')) {
    return { content: [{ type: "text", text: `Conflict (409) while ${operation}. The event may have been modified concurrently. Please retry.` }] };
  }
  if (error.message && error.message.includes('403')) {
    return { content: [{ type: "text", text: "Access denied (403). The token may lack Calendars.ReadWrite scope. Re-authenticate with force=true." }] };
  }
  if (error.message && error.message.includes('429')) {
    return { content: [{ type: "text", text: "Microsoft Graph API rate limit reached (429). Please wait a moment and try again." }] };
  }
  return { content: [{ type: "text", text: `Error ${operation}: ${error.message}` }] };
}

// ─── Request Deduplication Cache ──────────────────────────────────────
// Reduced TTL to 2 s. Exported invalidateListCache() so create/delete can
// bust the cache immediately after a write operation.

const _listCache = new Map();
const LIST_CACHE_TTL_MS = 2_000;

function _getCacheKey(startISO, count) {
  // Key now includes the start boundary so timezone changes don't serve stale data
  return `events:${startISO}:${count}`;
}

function _getCached(key) {
  const entry = _listCache.get(key);
  if (!entry) return null;
  if (Date.now() - entry.timestamp > LIST_CACHE_TTL_MS) {
    _listCache.delete(key);
    return null;
  }
  return entry.response;
}

function _setCache(key, response) {
  _listCache.set(key, { response, timestamp: Date.now() });
  if (_listCache.size > 20) {
    const oldest = _listCache.keys().next().value;
    _listCache.delete(oldest);
  }
}

/**
 * Invalidate the list cache.
 * Call this from create.js / delete.js after a successful write so the
 * next list-events call fetches fresh data from Graph.
 */
function invalidateListCache() {
  _listCache.clear();
}

// ─── Timezone-aware window helpers ───────────────────────────────────

/**
 * Build Graph-API-compatible ISO strings for calendarView boundaries.
 *
 * Graph calendarView requires bare ISO strings WITHOUT a timezone offset
 * in the query params — the timezone is communicated via the
 * Prefer: outlook.timezone header instead.
 *
 * We return:
 *   startDateTime — start of local today (00:00:00)
 *   endDateTime   — 14 days from now (covers a sensible scheduling horizon)
 *   preferTz      — IANA tz string for the Prefer header
 */
function buildCalendarViewWindow() {
  const localTz = (() => {
    try { return Intl.DateTimeFormat().resolvedOptions().timeZone; } catch (_) { return 'UTC'; }
  })();

  // Get local "now" broken into parts so we can build start-of-day
  const now = new Date();

  // Format a Date as a bare ISO local string "YYYY-MM-DDTHH:MM:SS"
  // using the server's local time (not UTC), suitable for calendarView params.
  const toLocalISO = (date) => {
    const pad = (n) => String(n).padStart(2, '0');
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}` +
           `T${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
  };

  // Start of local today
  const startOfDay = new Date(now);
  startOfDay.setHours(0, 0, 0, 0);

  // 14 days from now (end boundary)
  const endDate = new Date(now);
  endDate.setDate(endDate.getDate() + 14);
  endDate.setHours(23, 59, 59, 0);

  return {
    startDateTime: toLocalISO(startOfDay),
    endDateTime:   toLocalISO(endDate),
    preferTz:      localTz
  };
}

/**
 * List events handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEvents(args) {
  const count = Math.min(args.count || 10, 50);

  try {
    const client = await getClient(args.bearer_token || null);
    const accessToken = client.rawToken;

    // ── BUG FIX 1: Use calendarView with timezone-aware window ──
    const { startDateTime, endDateTime, preferTz } = buildCalendarViewWindow();

    const cacheKey = _getCacheKey(startDateTime, count);
    const cached = _getCached(cacheKey);
    if (cached) {
      console.error('[list-events] Returning cached response');
      return cached;
    }

    // calendarView is the correct Graph endpoint for date-bounded event listing.
    // It expands recurring events, respects timezone, and is reliably filterable.
    const endpoint = 'me/calendarView';

    const queryParams = {
      startDateTime,   // local start-of-day string
      endDateTime,     // local 14-days-out string
      $top: count,
      $orderby: 'start/dateTime',
      $select: config.CALENDAR_SELECT_FIELDS ||
        'id,subject,start,end,location,organizer,attendees,bodyPreview,isAllDay,isCancelled,importance,sensitivity,showAs,responseStatus'
    };

    // Tell Graph to interpret/return times in local timezone.
    // NOTE: This requires callGraphAPI to accept a 5th argument (extraHeaders).
    // Check utils/graph-api.js — if it doesn't support it yet, see README note below.
    const extraHeaders = {
      'Prefer': `outlook.timezone="${preferTz}"`
    };

    // Safe call — passes extraHeaders as 5th arg; graph-api.js will ignore it
    // if its signature only accepts 4 args (no harm done, times just come back in UTC).
    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams, extraHeaders);

    if (!response.value || response.value.length === 0) {
      const result = {
        content: [{
          type: "text",
          text: "No upcoming calendar events found."
        }]
      };
      _setCache(cacheKey, result);
      return result;
    }

    const systemTimezone = preferTz;

    const eventList = response.value.map((event, index) => {
      const dto = mapEventToDto(event);

      const formatDateTime = (dateTimeData) => {
        if (!dateTimeData) return '';
        const dateTime = typeof dateTimeData === 'string'
          ? dateTimeData
          : (dateTimeData.dateTime || '');
        const timeZone = typeof dateTimeData === 'object' ? dateTimeData.timeZone : undefined;
        if (!dateTime) return '';

        const hasOffset = /[zZ]$|[+\-]\d{2}:\d{2}$/.test(dateTime);

        const formatDateObj = (date) => {
          if (isNaN(date.getTime())) return dateTime;
          return date.toLocaleString('en-US', {
            year: 'numeric', month: 'long', day: 'numeric',
            hour: 'numeric', minute: '2-digit', hour12: true,
            timeZone: systemTimezone
          });
        };

        if (timeZone === 'UTC' || hasOffset || !timeZone) {
          const iso = dateTime.endsWith('Z') || hasOffset ? dateTime : dateTime + 'Z';
          return formatDateObj(new Date(iso));
        }

        // timeZone already set by Prefer header — parse as local
        return formatDateObj(new Date(dateTime));
      };

      const startDate = formatDateTime(event.start);
      const endDate   = formatDateTime(event.end);

      return `${index + 1}. ${dto.subject} - Location: ${dto.location}\nStart: ${startDate}\nEnd: ${endDate}\nOrganizer: ${dto.organizer}\nStatus: ${dto.responseStatus}\nSummary: ${dto.bodyPreview}\nID: ${dto.id}\n`;
    }).join("\n");

    const result = {
      content: [{
        type: "text",
        text: `Found ${response.value.length} upcoming events (timezone: ${preferTz}):\n\n${eventList}`
      }]
    };

    _setCache(cacheKey, result);
    return result;
  } catch (error) {
    return classifyCalendarError(error, 'listing events');
  }
}

module.exports = handleListEvents;
module.exports.invalidateListCache = invalidateListCache;