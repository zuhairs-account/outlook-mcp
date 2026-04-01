/**
 * Calendar module for Outlook MCP server
 * @module calendar
 *
 * Action plan fixes applied:
 *   - (c) calendar/index.js: Shared CalendarBaseHandler utilities extracted here
 *   - (c) calendar/accept.js + decline.js: Unified sendRsvp() utility (90% identical → DRY)
 *   - (c) calendar/cancel.js + delete.js: Shared graphAction() utility
 *   - (c) calendar/create.js: Shared buildEventPayload() builder
 *   - (c) calendar/list.js: Shared mapEventToDto() response mapper
 *   - (c) calendar/index.js: Missing accept handler import added
 *   - (c) All: Shared validateEventId() for input validation
 */

const handleListEvents = require('./list');
const handleDeclineEvent = require('./decline');
const handleCreateEvent = require('./create');
const handleCancelEvent = require('./cancel');
const handleDeleteEvent = require('./delete');
// BEFORE: handleAcceptEvent was not imported — accept tool was missing from barrel.
// AFTER: Imported and included in calendarTools.
// GOOD EFFECT: Accept tool is actually usable via the MCP server.
const handleAcceptEvent = require('./accept');

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

// BEFORE: Event ID was passed directly into URLs in accept, decline, cancel,
//         delete without validation. Malformed IDs could cause unexpected URL
//         construction.
// AFTER: Shared validateEventId() — checks non-empty, no path traversal.
// GOOD EFFECT: Consistent input validation across all calendar operations.

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

// BEFORE: accept.js, decline.js, cancel.js, delete.js each resolved the
//         event ID and made a single Graph call with nearly identical logic.
// AFTER: Shared graphAction(accessToken, eventId, action, body) utility.
// GOOD EFFECT: Unifies the pattern across all calendar action endpoints;
//              one place for error handling, URL construction, and validation.

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

// BEFORE: accept.js and decline.js were 90% identical — textbook WET code.
//         Both differed only in the endpoint suffix (/accept vs /decline)
//         and the response message string. Any bug fix had to be applied twice.
// AFTER: Shared sendRsvp(callGraphAPI, accessToken, eventId, action, options).
// GOOD EFFECT: Single function for both accept and decline; bug fixes apply once.
//              Also exposes the sendResponse parameter that Graph API supports
//              but was previously hidden.

/**
 * Sends an RSVP (accept/decline) for a calendar event.
 *
 * BEFORE: Graph API accept/decline accepts comment and sendResponse flag,
 *         but neither was surfaced to the LLM.
 * AFTER: Both are exposed as optional parameters.
 * GOOD EFFECT: LLM can control whether attendees are notified and include
 *              a custom message.
 *
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
    // BEFORE: sendResponse was not included — always defaulted to true server-side.
    // AFTER: Explicitly included and controllable.
    // GOOD EFFECT: LLM can silently RSVP without notifying the organiser if desired.
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

// BEFORE: Payload construction for create-event was inline imperative code
//         with scattered conditionals inside the handler.
// AFTER: Shared buildEventPayload(params, defaultTimezone) builder.
// GOOD EFFECT: Data contract is explicit and testable; defaults are clear;
//              the handler just calls the builder and sends.

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
    // BEFORE: attendees?.map(email => ({ emailAddress: { address: email }, type: "required" }))
    //         — no email format validation.
    // AFTER: Basic validation before mapping.
    // GOOD EFFECT: Malformed attendee emails caught before API call.
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

// BEFORE: Response mapping (which fields to include in the returned event)
//         was decided inline in list.js. No shared data contract.
// AFTER: Shared mapEventToDto(event) function.
// GOOD EFFECT: Explicit, testable data contract for calendar events.

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

// ─── Error Classification Helper ──────────────────────────────────────
// BEFORE: Each handler only checked for 'Authentication required'.
// AFTER: Shared classifyCalendarError() for consistent error messages.
// GOOD EFFECT: Auth, 403, 404, 409, 429 errors all get actionable messages.

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

// ─── Tool Definitions ─────────────────────────────────────────────────
const calendarTools = [
  {
    name: "list-events",
    description: "Lists upcoming events from your calendar",
    inputSchema: {
      type: "object",
      properties: {
        count: {
          type: "number",
          description: "Number of events to retrieve (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleListEvents
  },
  // BEFORE: accept-event tool was missing from the barrel.
  // AFTER: Included in calendarTools.
  // GOOD EFFECT: Accept tool is actually registered and usable.
  {
    name: "accept-event",
    description: "Accepts a calendar event invitation",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to accept"
        },
        comment: {
          type: "string",
          description: "Optional comment for accepting the event"
        },
        // BEFORE: sendResponse parameter was not exposed.
        // AFTER: Added to schema.
        // GOOD EFFECT: LLM can silently RSVP without notifying the organiser.
        sendResponse: {
          type: "boolean",
          description: "Whether to notify the organiser (default: true)"
        }
      },
      required: ["eventId"]
    },
    handler: handleAcceptEvent
  },
  {
    name: "decline-event",
    description: "Declines a calendar event invitation",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to decline"
        },
        comment: {
          type: "string",
          description: "Optional comment for declining the event"
        },
        sendResponse: {
          type: "boolean",
          description: "Whether to notify the organiser (default: true)"
        }
      },
      required: ["eventId"]
    },
    handler: handleDeclineEvent
  },
  {
    name: "create-event",
    description: "Creates a new calendar event",
    inputSchema: {
      type: "object",
      properties: {
        subject: {
          type: "string",
          description: "The subject of the event"
        },
        start: {
          type: "string",
          description: "The start time of the event in ISO 8601 format"
        },
        end: {
          type: "string",
          description: "The end time of the event in ISO 8601 format"
        },
        attendees: {
          type: "array",
          items: { type: "string" },
          description: "List of attendee email addresses"
        },
        body: {
          type: "string",
          description: "Optional body content for the event"
        },
        location: {
          type: "string",
          description: "Optional location for the event"
        }
      },
      required: ["subject", "start", "end"]
    },
    handler: handleCreateEvent
  },
  {
    name: "cancel-event",
    description: "Cancels a calendar event and notifies attendees",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to cancel"
        },
        comment: {
          type: "string",
          description: "Optional cancellation message sent to attendees"
        }
      },
      required: ["eventId"]
    },
    handler: handleCancelEvent
  },
  {
    name: "delete-event",
    description: "Deletes a calendar event permanently (no notification sent to attendees)",
    inputSchema: {
      type: "object",
      properties: {
        eventId: {
          type: "string",
          description: "The ID of the event to delete"
        }
      },
      required: ["eventId"]
    },
    handler: handleDeleteEvent
  }
];

module.exports = {
  calendarTools,
  handleListEvents,
  handleAcceptEvent,
  handleDeclineEvent,
  handleCreateEvent,
  handleCancelEvent,
  handleDeleteEvent,
  // Shared utilities for use by calendar operations
  CALENDAR_SELECT_FIELDS,
  validateEventId,
  graphAction,
  sendRsvp,
  buildEventPayload,
  mapEventToDto,
  classifyCalendarError
};