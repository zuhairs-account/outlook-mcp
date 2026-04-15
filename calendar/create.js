/**
 * Create event functionality
 *
 * Action plan fixes applied:
 *   - (c) calendar/create.js: Payload construction via shared buildEventPayload() builder
 *   - (c) calendar/create.js: Date input validation (Date.parse) before API call
 *   - (c) calendar/create.js: Attendee email validation via buildEventPayload()
 *   - (c) calendar/create.js: Error classification via shared classifyCalendarError()
 *   - (e) calendar/create.js: Idempotency note — duplicate invocations create duplicate events
 *
 * BUG FIX: Timezone-aware event creation.
 *   BEFORE: buildEventPayload() always fell back to DEFAULT_TIMEZONE (often 'UTC'),
 *           so "9:00 AM" was stored as 9:00 UTC and displayed as 2:00 PM in PKT (UTC+5).
 *   AFTER:  If start/end are bare local datetime strings (no Z / no offset), we detect
 *           the server's local IANA timezone via Intl and pass it in the payload so
 *           Graph API interprets the time correctly.
 */
const { callGraphAPI } = require('../utils/graph-api');
const { getClient } = require('../auth');
const { DEFAULT_TIMEZONE } = require('../config');

// Import from shared.js (not ./index) to avoid circular dependency.
const { buildEventPayload, classifyCalendarError } = require('./shared');
// Bust the list cache after a successful create so list-events returns fresh data
const { invalidateListCache } = require('./list');

/**
 * Detect the effective timezone for a datetime string.
 *
 * If the string already carries UTC offset info (ends with Z or ±HH:MM),
 * we honour it as-is and tell Graph it's UTC (the offset is embedded).
 * If it's a bare local string like "2026-04-15T09:00:00", we use the
 * server's local IANA timezone so Graph stores it in the right zone.
 *
 * @param {string} dateTimeStr
 * @param {string} fallback  - config DEFAULT_TIMEZONE (last resort)
 * @returns {string} IANA timezone string
 */
function resolveTimezone(dateTimeStr, fallback) {
  if (!dateTimeStr) return fallback || 'UTC';
  // Already has offset info — treat as UTC (offset is self-describing)
  if (/[zZ]$|[+\-]\d{2}:\d{2}$/.test(dateTimeStr)) return 'UTC';
  // Bare local string — use server's local timezone
  try {
    const localTz = Intl.DateTimeFormat().resolvedOptions().timeZone;
    if (localTz) return localTz;
  } catch (_) {}
  return fallback || 'UTC';
}

/**
 * Create event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateEvent(args) {
  const { subject, start, end, attendees, body, location } = args;

  if (!subject || !start || !end) {
    return {
      content: [{
        type: "text",
        text: "Subject, start, and end times are required to create an event."
      }]
    };
  }

  const startDateTime = typeof start === 'string' ? start : start.dateTime;
  const endDateTime   = typeof end   === 'string' ? end   : end.dateTime;

  if (isNaN(Date.parse(startDateTime))) {
    return {
      content: [{
        type: "text",
        text: `Invalid start date format: "${startDateTime}". Please use ISO 8601 format (e.g., 2026-04-15T09:00:00).`
      }]
    };
  }

  if (isNaN(Date.parse(endDateTime))) {
    return {
      content: [{
        type: "text",
        text: `Invalid end date format: "${endDateTime}". Please use ISO 8601 format (e.g., 2026-04-15T10:00:00).`
      }]
    };
  }

  // ── Timezone Resolution ──
  // BUG FIX: Determine the correct timezone for bare local datetime strings.
  // If start already carries a timezone object, use that; otherwise resolve.
  const startTz = (typeof start === 'object' && start.timeZone)
    ? start.timeZone
    : resolveTimezone(startDateTime, DEFAULT_TIMEZONE);

  const endTz = (typeof end === 'object' && end.timeZone)
    ? end.timeZone
    : resolveTimezone(endDateTime, DEFAULT_TIMEZONE);

  try {
    const client = await getClient(args.bearer_token || null);
    const accessToken = client.rawToken;

    // Pass normalised start/end objects so buildEventPayload gets the timezone.
    const normalizedStart = { dateTime: startDateTime, timeZone: startTz };
    const normalizedEnd   = { dateTime: endDateTime,   timeZone: endTz };

    const payload = buildEventPayload(
      { subject, start: normalizedStart, end: normalizedEnd, attendees, body, location },
      startTz   // also used as the default for any field that omits its zone
    );

    const response = await callGraphAPI(accessToken, 'POST', 'me/events', payload);

    // Bust the list cache so an immediate list-events call sees the new event
    invalidateListCache();

    return {
      content: [{
        type: "text",
        text: `Event '${subject}' has been successfully created.\n\nEvent ID: ${response.id}\nStart: ${startDateTime} (${startTz})\nEnd: ${endDateTime} (${endTz})${attendees && attendees.length > 0 ? `\nAttendees: ${attendees.join(', ')}` : ''}`
      }]
    };
  } catch (error) {
    return classifyCalendarError(error, 'creating event');
  }
}

module.exports = handleCreateEvent;