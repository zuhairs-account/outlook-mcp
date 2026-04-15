/**
 * Delete event functionality
 *
 * Action plan fixes applied:
 *   - (c) calendar/delete.js: Event ID validated via shared validateEventId()
 *   - (c) calendar/delete.js: Uses shared graphAction() utility
 *   - (e) calendar/delete.js: 404 on delete treated as success (event already deleted)
 *         — makes the operation safely retryable
 *   - (c) calendar/delete.js: Error classification via shared classifyCalendarError()
 */
const { callGraphAPI } = require('../utils/graph-api');
const { getClient } = require('../auth');

// Import from shared.js (not ./index) to avoid circular dependency.
const { validateEventId, graphAction, classifyCalendarError } = require('./shared');

/**
 * Delete event handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteEvent(args) {
  const { eventId } = args;

  // ── Input Validation ──
  // BEFORE: if (!eventId) — simple null check, no format validation.
  //         ID passed directly into URL without sanitisation.
  // AFTER: validateEventId() — checks non-empty, no path traversal chars.
  // GOOD EFFECT: Malformed IDs caught with clear error before API call.
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
    const client = await getClient(args.bearer_token || null);
    const accessToken = client.rawToken;

    // ── Shared graphAction Utility ──
    // BEFORE: const endpoint = `me/events/${eventId}`;
    //         await callGraphAPI(accessToken, 'DELETE', endpoint);
    //         — inline URL construction + API call, duplicated across ops.
    // AFTER: await graphAction(callGraphAPI, accessToken, eventId, '', 'DELETE');
    // GOOD EFFECT: Single utility for all calendar event operations;
    //              URL construction and validation happen in one place.
    try {
      await graphAction(callGraphAPI, accessToken, eventId, '', 'DELETE');
    } catch (deleteError) {
      // ── 404 = Success (Idempotent Delete) ──
      // BEFORE: A 404 on delete (event already deleted) was treated as an
      //         error and surfaced to the LLM.
      // AFTER: 404 is treated as success — the event is gone, which is
      //        the desired end state.
      // GOOD EFFECT: Makes the operation safely retryable; LLM retries
      //              after a timeout don't produce false errors.
      if (deleteError.message && (deleteError.message.includes('404') || deleteError.message.includes('ErrorItemNotFound'))) {
        return {
          content: [{
            type: "text",
            text: `Event with ID ${eventId} was already deleted or does not exist. No action needed.`
          }]
        };
      }
      throw deleteError; // Re-throw non-404 errors for classification below
    }

    return {
      content: [{
        type: "text",
        text: `Event with ID ${eventId} has been successfully deleted.`
      }]
    };
  } catch (error) {
    // BEFORE: Only checked for 'Authentication required'.
    // AFTER: Shared classifyCalendarError() handles auth, 403, 409, 429.
    // GOOD EFFECT: Consistent, actionable error messages.
    return classifyCalendarError(error, 'deleting event');
  }
}

module.exports = handleDeleteEvent;