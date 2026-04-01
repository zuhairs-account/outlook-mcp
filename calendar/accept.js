/**
 * Accept event functionality
 *
 * Action plan fixes applied:
 *   - (c) calendar/accept.js + decline.js: 90% identical → unified via shared
 *         sendRsvp() utility. This file is now a thin wrapper.
 *   - (c) calendar/accept.js: sendResponse flag exposed — LLM can silently RSVP
 *   - (c) calendar/accept.js: comment parameter was already present but now routes
 *         through the shared utility for consistency
 *   - (c) calendar/accept.js: Event ID validated via shared validateEventId()
 *   - (c) calendar/accept.js: Error classification via shared classifyCalendarError()
 *   - (e) calendar/accept.js: Request dedup guard (eventId + action within 5s window)
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

const { validateEventId, sendRsvp, classifyCalendarError } = require('./index');

// ─── RSVP Deduplication Guard ─────────────────────────────────────────
// BEFORE: Repeated accept calls on the same event were silently duplicated
//         — the Graph API accepts them but sends duplicate notifications.
// AFTER: 5-second window dedup guard keyed on eventId + action.
// GOOD EFFECT: Prevents spurious LLM retries from sending duplicate
//              RSVP notifications to the organiser.

const _rsvpDedup = new Map();
const RSVP_DEDUP_TTL_MS = 5_000;

function _isDuplicateRsvp(eventId, action) {
  const key = `${eventId}:${action}`;
  const entry = _rsvpDedup.get(key);
  if (entry && Date.now() - entry < RSVP_DEDUP_TTL_MS) {
    return true;
  }
  _rsvpDedup.set(key, Date.now());
  // Cleanup old entries
  if (_rsvpDedup.size > 50) {
    const oldest = _rsvpDedup.keys().next().value;
    _rsvpDedup.delete(oldest);
  }
  return false;
}

/**
 * Accept event handler
 *
 * BEFORE: This was a standalone 60-line handler with inline endpoint
 *         construction, API call, and error handling — 90% identical
 *         to decline.js. Any bug fix had to be applied twice.
 * AFTER: Thin wrapper around shared sendRsvp('accept').
 * GOOD EFFECT: Bug fixes and feature additions (sendResponse, dedup)
 *              apply to both accept and decline simultaneously.
 *
 * @param {object} args - Tool arguments
 * @param {string} args.eventId - Event ID (required)
 * @param {string} [args.comment] - Optional RSVP comment
 * @param {boolean} [args.sendResponse] - Whether to notify the organiser (default: true)
 * @returns {object} - MCP response
 */
async function handleAcceptEvent(args) {
  const { eventId, comment, sendResponse } = args;

  // ── Input Validation ──
  // BEFORE: if (!eventId) — simple null check.
  // AFTER: validateEventId() — format + path traversal check.
  // GOOD EFFECT: Malformed IDs caught before API call.
  const validation = validateEventId(eventId);
  if (!validation.valid) {
    return {
      content: [{
        type: "text",
        text: validation.error
      }]
    };
  }

  // ── Dedup Guard ──
  // BEFORE: No dedup — repeated accepts sent duplicate notifications.
  // AFTER: Skip if same eventId+accept was called within 5 seconds.
  // GOOD EFFECT: Prevents duplicate RSVP notifications from LLM retries.
  if (_isDuplicateRsvp(eventId, 'accept')) {
    return {
      content: [{
        type: "text",
        text: `Event with ID ${eventId} was already accepted moments ago. Skipping duplicate.`
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // ── Shared sendRsvp Utility ──
    // BEFORE: const endpoint = `me/events/${eventId}/accept`;
    //         const body = { comment: comment || "Accepted via API" };
    //         await callGraphAPI(accessToken, 'POST', endpoint, body);
    //         — 90% identical to decline.js.
    // AFTER: return await sendRsvp(callGraphAPI, accessToken, eventId, 'accept', { comment, sendResponse });
    // GOOD EFFECT: DRY — one function handles both accept and decline.
    return await sendRsvp(callGraphAPI, accessToken, eventId, 'accept', {
      comment,
      sendResponse
    });
  } catch (error) {
    // BEFORE: Only checked for 'Authentication required'.
    // AFTER: Shared classifyCalendarError().
    // GOOD EFFECT: Consistent, actionable error messages.
    return classifyCalendarError(error, 'accepting event');
  }
}

module.exports = handleAcceptEvent;