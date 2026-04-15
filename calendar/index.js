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

// ─── Shared Constants and Utilities ──────────────────────────────────
// Imported from shared.js to avoid circular dependencies.
// Handler files import directly from shared.js — NOT from this barrel.
const {
  CALENDAR_SELECT_FIELDS,
  validateEventId,
  graphAction,
  sendRsvp,
  buildEventPayload,
  mapEventToDto,
  classifyCalendarError
} = require('./shared');

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