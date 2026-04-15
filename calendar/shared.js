/**
 * Shared constants and utilities for the calendar module.
 *
 * Extracted from index.js to break the circular dependency that caused
 * validateEventId, sendRsvp, classifyCalendarError, buildEventPayload,
 * graphAction, and mapEventToDto to resolve as undefined when handler
 * files imported from ./index.
 *
 * Handler files (accept, decline, cancel, delete, create) now import
 * from ./shared instead of ./index, eliminating the cycle.
 */

// ─── Shared Constants ─────────────────────────────────────────────────

/**
 * Standard fields for calendar event list results.
 */
const CALENDAR_SELECT_FIELDS = [
  'id', 'subject', 'start', 'end', 'location', 'organizer',
  'attendees', 'bodyPreview', 'isAllDay', 'isCancelled',
  'importance', 'sensitivity', 'showAs', 'responseStatus'
].join(',');

// ─── Shared Utilities ─────────────────────────────────────────────────

/**
 * Validates an event ID before use in a Graph API URL.
 * @param {string} id - The event ID to validate
 * @returns {{ valid: boolean, error?: string }}
 */
function validateEventId(id) {
  if (!id || typeof id !== 'string') {
    return { valid: false, error: 'Event ID is required and must be a string.' };
  }
  if (id.trim().length === 0) {
    return { valid: false, error: 'Event ID cannot be empty.' };
  }
  if (/[\/\\]/.test(id)) {
    return { valid: false, error: 'Event ID contains invalid characters.' };
  }
  return { valid: true };
}

/**
 * Performs a Graph API action on a calendar event.
 * @param {Function} callGraphAPI - The Graph API call function
 * @param {string} accessToken - Valid access token
 * @param {string} eventId - Event ID
 * @param {string} action - Action suffix (e.g., 'accept', 'decline', 'cancel') or empty for the event itself
 * @param {string} method - HTTP method (POST, DELETE, etc.)
 * @param {object} [body] - Optional request body
 * @returns {Promise<object|void>} - API response
 */
async function graphAction(callGraphAPI, accessToken, eventId, action, method = 'POST', body = null) {
  const validation = validateEventId(eventId);
  if (!validation.valid) {
    throw new Error(validation.error);
  }
  const endpoint = action
    ? `me/events/${encodeURIComponent(eventId)}/${action}`
    : `me/events/${encodeURIComponent(eventId)}`;
  return callGraphAPI(accessToken, method, endpoint, body);
}

/**
 * Sends an RSVP (accept/decline) for a calendar event.
 * @param {Function} callGraphAPI - Graph API call function
 * @param {string} accessToken - Valid access token
 * @param {string} eventId - Event ID
 * @param {'accept'|'decline'} action - RSVP action
 * @param {object} [options]
 * @param {string} [options.comment] - Optional RSVP comment
 * @param {boolean} [options.sendResponse=true] - Whether to notify the organiser
 * @returns {Promise<object>} - MCP response
 */
async function sendRsvp(callGraphAPI, accessToken, eventId, action, options = {}) {
  const body = {
    comment: options.comment || `${action === 'accept' ? 'Accepted' : 'Declined'} via API`,
    sendResponse: options.sendResponse !== undefined ? options.sendResponse : true
  };

  await graphAction(callGraphAPI, accessToken, eventId, action, 'POST', body);

  const actionPastTense = action === 'accept' ? 'accepted' : 'declined';
  return {
    content: [{
      type: "text",
      text: `Event with ID ${eventId} has been successfully ${actionPastTense}.`
    }]
  };
}

/**
 * Builds a Graph API event payload from tool parameters.
 * @param {object} params - Tool parameters
 * @param {string} defaultTimezone - Fallback timezone
 * @returns {object} - Graph API event body
 */
function buildEventPayload(params, defaultTimezone = 'UTC') {
  const { subject, start, end, attendees, body, location, isAllDay } = params;

  const payload = {
    subject,
    start: {
      dateTime: start.dateTime || start,
      timeZone: start.timeZone || defaultTimezone
    },
    end: {
      dateTime: end.dateTime || end,
      timeZone: end.timeZone || defaultTimezone
    },
    body: {
      contentType: 'HTML',
      content: body || ''
    }
  };

  if (attendees && Array.isArray(attendees)) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    payload.attendees = attendees
      .filter(email => typeof email === 'string' && email.trim())
      .map(email => {
        const trimmed = email.trim();
        if (!emailRegex.test(trimmed)) {
          console.warn(`[calendar] Invalid attendee email skipped: ${trimmed}`);
        }
        return {
          emailAddress: { address: trimmed },
          type: 'required'
        };
      });
  }

  if (location) {
    payload.location = { displayName: location };
  }

  if (isAllDay !== undefined) {
    payload.isAllDay = isAllDay;
  }

  return payload;
}

/**
 * Maps a raw Graph API event to a clean DTO for display.
 * @param {object} event - Raw Graph API event
 * @returns {object} - Clean event DTO
 */
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

/**
 * Classifies a Graph API error and returns an MCP-friendly response.
 * @param {Error} error - The caught error
 * @param {string} operation - Description of what was attempted
 * @returns {object} - MCP response
 */
function classifyCalendarError(error, operation) {
  if (error.message === 'Authentication required') {
    return {
      content: [{
        type: "text",
        text: "Authentication required. Please use the 'authenticate' tool first."
      }]
    };
  }

  if (error.message && error.message.includes('404')) {
    return {
      content: [{
        type: "text",
        text: `Event not found (404). It may have been deleted or the ID is incorrect.`
      }]
    };
  }

  if (error.message && error.message.includes('409')) {
    return {
      content: [{
        type: "text",
        text: `Conflict (409) while ${operation}. The event may have been modified concurrently. Please retry.`
      }]
    };
  }

  if (error.message && error.message.includes('403')) {
    return {
      content: [{
        type: "text",
        text: `Access denied (403). The token may lack Calendars.ReadWrite scope. Re-authenticate with force=true.`
      }]
    };
  }

  if (error.message && error.message.includes('429')) {
    return {
      content: [{
        type: "text",
        text: "Microsoft Graph API rate limit reached (429). Please wait a moment and try again."
      }]
    };
  }

  return {
    content: [{
      type: "text",
      text: `Error ${operation}: ${error.message}`
    }]
  };
}

module.exports = {
  CALENDAR_SELECT_FIELDS,
  validateEventId,
  graphAction,
  sendRsvp,
  buildEventPayload,
  mapEventToDto,
  classifyCalendarError
};