/**
 * Cancel event functionality
 *
 * Action plan fixes applied:
 *   - (c) calendar/cancel.js: Uses shared graphAction() utility — unified with delete.js
 *   - (c) calendar/cancel.js: Attendee notification message (Comment) parameter
 *         already existed but now explicitly documented and validated
 *   - (c) calendar/cancel.js: Event ID validated via shared validateEventId()
 *   - (c) calendar/cancel.js: Error classification via shared classifyCalendarError()
 *   - (e) calendar/cancel.js: 409 Conflict handled explicitly (concurrent accept/decline)
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

const { validateEventId, graphAction, classifyCalendarError } = require('./index');

/**
 * Cancel event handler
 * Sends a cancellation notice to all attendees via /me/events/{id}/cancel
 * (distinct from outright deletion — this notifies attendees).
 *
 * @param {object} args - Tool arguments
 * @param {string} args.eventId - Event ID (required)
 * @param {string} [args.comment] - Cancellation message sent to attendees
 * @returns {object} - MCP response
 */
async function handleCancelEvent(args) {
  const { eventId, comment } = args;

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

  try {
    const accessToken = await ensureAuthenticated();

    // ── Cancellation Message ──
    // BEFORE: const body = { comment: comment || "Cancelled via API" };
    //         — comment parameter existed but was not documented in the
    //         action plan as "No attendee notification message parameter".
    // AFTER: Explicitly documented; the Graph API cancel action accepts an
    //        optional Comment body for the cancellation message.
    // GOOD EFFECT: The LLM can provide a meaningful cancellation reason
    //              that attendees will see in their notification.
    const body = {
      comment: comment || "Cancelled via M365 Assistant"
    };

    // ── Shared graphAction Utility ──
    // BEFORE: const endpoint = `me/events/${eventId}/cancel`;
    //         await callGraphAPI(accessToken, 'POST', endpoint, body);
    //         — inline, redundant with delete.js pattern.
    // AFTER: graphAction(callGraphAPI, accessToken, eventId, 'cancel', 'POST', body)
    // GOOD EFFECT: Unified utility for all calendar event actions.
    await graphAction(callGraphAPI, accessToken, eventId, 'cancel', 'POST', body);

    return {
      content: [{
        type: "text",
        text: `Event with ID ${eventId} has been successfully cancelled. Attendees have been notified.`
      }]
    };
  } catch (error) {
    // BEFORE: Only checked for 'Authentication required'.
    // AFTER: Shared classifyCalendarError() — handles 409 Conflict
    //        (race with concurrent accept/decline on same event), 403, 429.
    // GOOD EFFECT: Consistent, actionable error messages. 409 explicitly
    //              tells the user to retry, not to panic.
    return classifyCalendarError(error, 'cancelling event');
  }
}

module.exports = handleCancelEvent;