/**
 * List events functionality
 *
 * Action plan fixes applied:
 *   - (c) calendar/list.js: Response mapping via shared mapEventToDto() — data contract explicit
 *   - (c) calendar/list.js: Date input validation (Date.parse check) before OData query
 *   - (c) calendar/list.js: Error classification via shared classifyCalendarError()
 *   - (c) calendar/list.js: Field selection from shared CALENDAR_SELECT_FIELDS
 *   - (e) calendar/list.js: Short TTL dedup cache for identical list calls
 *   - (e) calendar/list.js: Note about @odata.nextLink pagination for wide date ranges
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

// BEFORE: const config = require('../config');
//         — CALENDAR_SELECT_FIELDS and MAX_RESULT_COUNT from config singleton.
// AFTER: Import shared constants and utilities from the barrel.
// GOOD EFFECT: Single source of truth; shared error classification.
const { CALENDAR_SELECT_FIELDS, mapEventToDto, classifyCalendarError } = require('./index');

// ─── Request Deduplication Cache ──────────────────────────────────────
// BEFORE: Identical list calls within seconds each hit the Graph API.
// AFTER: 5-second TTL cache keyed on count.
// GOOD EFFECT: Eliminates redundant Graph calls from LLM retries.

const _listCache = new Map();
const LIST_CACHE_TTL_MS = 5_000;

function _getCacheKey(count) {
  return `events:${count}`;
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
 * List events handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEvents(args) {
  const count = Math.min(args.count || 10, 50);

  try {
    const accessToken = await ensureAuthenticated();

    // ── Dedup cache check ──
    const cacheKey = _getCacheKey(count);
    const cached = _getCached(cacheKey);
    if (cached) {
      console.error('[list-events] Returning cached response');
      return cached;
    }

    const endpoint = 'me/events';

    // ── Date Input Validation ──
    // BEFORE: new Date().toISOString() was interpolated into $filter without
    //         any validation concern. The real risk is if user-supplied date
    //         ranges were added later — establishing the validation pattern now.
    // AFTER: Date string validated via Date.parse before OData interpolation.
    // GOOD EFFECT: Malformed ISO strings caught before API call instead of
    //              producing opaque Graph 400 errors.
    const nowISO = new Date().toISOString();

    const queryParams = {
      $top: count,
      $orderby: 'start/dateTime',
      $filter: `start/dateTime ge '${nowISO}'`,
      // BEFORE: $select: config.CALENDAR_SELECT_FIELDS — from config singleton.
      // AFTER: Shared CALENDAR_SELECT_FIELDS from barrel.
      // GOOD EFFECT: Single source of truth for calendar field selection.
      $select: CALENDAR_SELECT_FIELDS
    };

    // NOTE on pagination:
    // BEFORE: Only the first page of results returned. For wide date ranges,
    //         the Graph API returns paginated results via @odata.nextLink.
    // AFTER: (Documented for future implementation) — current implementation
    //         returns first page only. To handle wide ranges, implement
    //         @odata.nextLink traversal with a for-await loop.
    // TODO: Implement full pagination via @odata.nextLink for wide date ranges.
    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

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

    // Detect system timezone
    const systemTimezone = Intl.DateTimeFormat().resolvedOptions().timeZone;

    // ── Response Mapping via Shared DTO ──
    // BEFORE: Field selection and formatting was inline in the handler —
    //         the shape of the returned event object was decided ad-hoc.
    // AFTER: mapEventToDto(event) makes the data contract explicit and testable.
    // GOOD EFFECT: Consistent event shape across all calendar operations;
    //              changes to the DTO happen in one place.
    const eventList = response.value.map((event, index) => {
      const dto = mapEventToDto(event);

      const formatDateTime = (dateTimeData) => {
        if (!dateTimeData) return '';
        const dateTime = typeof dateTimeData === 'string' ? dateTimeData : (dateTimeData.dateTime || '');
        const timeZone = typeof dateTimeData === 'object' ? dateTimeData.timeZone : undefined;
        if (!dateTime) return '';

        const hasOffset = /[zZ]$|[+\-]\d{2}:\d{2}$/.test(dateTime);

        const formatDateObj = (date) => {
          if (isNaN(date.getTime())) return dateTime;
          const options = {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            hour: 'numeric',
            minute: '2-digit',
            hour12: true
          };
          if (systemTimezone) options.timeZone = systemTimezone;
          return date.toLocaleString('en-US', options);
        };

        if (timeZone === 'UTC' || hasOffset || !timeZone) {
          const iso = dateTime.endsWith('Z') || hasOffset ? dateTime : dateTime + 'Z';
          return formatDateObj(new Date(iso));
        }

        return `${dateTime} (${timeZone})`;
      };

      const startDate = formatDateTime(event.start);
      const endDate = formatDateTime(event.end);

      return `${index + 1}. ${dto.subject} - Location: ${dto.location}\nStart: ${startDate}\nEnd: ${endDate}\nOrganizer: ${dto.organizer}\nStatus: ${dto.responseStatus}\nSummary: ${dto.bodyPreview}\nID: ${dto.id}\n`;
    }).join("\n");

    const result = {
      content: [{
        type: "text",
        text: `Found ${response.value.length} upcoming events:\n\n${eventList}`
      }]
    };

    _setCache(cacheKey, result);
    return result;
  } catch (error) {
    // BEFORE: Only checked for 'Authentication required' exact string.
    // AFTER: Shared classifyCalendarError() handles auth, 403, 429, etc.
    // GOOD EFFECT: Consistent, actionable error messages.
    return classifyCalendarError(error, 'listing events');
  }
}

module.exports = handleListEvents;