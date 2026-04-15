/**
 * Decline event functionality
 *
 * Action plan fixes applied:
 *   - (c) calendar/decline.js: Identical structure to accept.js → unified via
 *         shared sendRsvp() utility. Textbook WET code eliminated.
 *   - (c) calendar/decline.js: sendResponse flag exposed
 *   - (c) calendar/decline.js: Event ID validated via shared validateEventId()
 *   - (c) calendar/decline.js: Error classification via shared classifyCalendarError()
 *   - (e) calendar/decline.js: Same dedup guard as accept.js
 */
const { callGraphAPI } = require('../utils/graph-api');
const { getClient } = require('../auth');

// Import from shared.js (not ./index) to avoid circular dependency.
const { validateEventId, sendRsvp, classifyCalendarError } = require('./shared');

// ─── RSVP Deduplication Guard ─────────────────────────────────────────
// Same pattern as accept.js — prevents duplicate decline notifications.
const _rsvpDedup = new Map();
const RSVP_DEDUP_TTL_MS = 5_000;

function _isDuplicateRsvp(eventId, action) {
  const key = `${eventId}:${action}`;
  const entry = _rsvpDedup.get(key);
  if (entry && Date.now() - entry < RSVP_DEDUP_TTL_MS) {
    return true;
  }
  _rsvpDedup.set(key, Date.now());
  if (_rsvpDedup.size > 50) {
    const oldest = _rsvpDedup.keys().next().value;
    _rsvpDedup.delete(oldest);
  }
  return false;
}

/**
 * Decline event handler
 *
 * BEFORE: 60-line standalone handler — 90% identical to accept.js.
 *         The only differences were the endpoint suffix (/decline vs /accept)
 *         and the response message string. Any bug fix had to be applied twice.
 * AFTER: Thin wrapper around shared sendRsvp('decline').
 * GOOD EFFECT: Maintaining two near-identical files is eliminated.
 *
 * @param {object} args - Tool arguments
 * @param {string} args.eventId - Event ID (required)
 * @param {string} [args.comment] - Optional RSVP comment
 * @param {boolean} [args.sendResponse] - Whether to notify the organiser (default: true)
 * @returns {object} - MCP response
 */
async function handleDeclineEvent(args) {
  const { eventId, comment, sendResponse } = args;

  // ── Input Validation ──
  // BEFORE: if (!eventId) — simple null check.
  // AFTER: validateEventId() — format + path traversal check.
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
  if (_isDuplicateRsvp(eventId, 'decline')) {
    return {
      content: [{
        type: "text",
        text: `Event with ID ${eventId} was already declined moments ago. Skipping duplicate.`
      }]
    };
  }

  try {
    const client = await getClient(args.bearer_token || null);
    const accessToken = client.rawToken;

    // ── Shared sendRsvp Utility ──
    // BEFORE: const endpoint = `me/events/${eventId}/decline`;
    //         const body = { comment: comment || "Declined via API" };
    //         await callGraphAPI(accessToken, 'POST', endpoint, body);
    // AFTER: sendRsvp(callGraphAPI, accessToken, eventId, 'decline', { comment, sendResponse })
    // GOOD EFFECT: DRY — same function handles accept and decline.
    return await sendRsvp(callGraphAPI, accessToken, eventId, 'decline', {
      comment,
      sendResponse
    });
  } catch (error) {
    // BEFORE: Only checked for 'Authentication required'.
    // AFTER: Shared classifyCalendarError().
    return classifyCalendarError(error, 'declining event');
  }
}

module.exports = handleDeclineEvent;