/**
 * Email module for Outlook MCP server
 * @module email
 *
 * Action plan fixes applied:
 *   - (c) email/index.js: Centralised EMAIL_SELECT_FIELDS constant shared by all email ops
 *   - (c) email/mark-as-read.js: Shared patchMessage utility extracted here
 *   - (c) email/send.js + draft.js: Shared parseRecipients utility for robust email parsing
 *   - (c) email/read.js + mark-as-read.js: Shared validateEmailId utility
 */

const handleListEmails = require('./list');
const handleSearchEmails = require('./search');
const handleReadEmail = require('./read');
const handleSendEmail = require('./send');
const handleDraftEmail = require('./draft');
const handleMarkAsRead = require('./mark-as-read');

// ─── Shared Constants ─────────────────────────────────────────────────
// BEFORE: Each operation file independently decided which fields to request
//         from the Graph API ($select), leading to inconsistent field sets.
// AFTER: Centralised constants importable by all email operations.
// GOOD EFFECT: Single source of truth for field selection — adding a new
//              field happens once here, not scattered across 6 files.

/**
 * Standard fields for email list/search results.
 * Used by list.js, search.js, and any operation returning email summaries.
 */
const EMAIL_SELECT_FIELDS = [
  'id', 'subject', 'from', 'toRecipients', 'ccRecipients',
  'receivedDateTime', 'isRead', 'hasAttachments', 'importance',
  'bodyPreview', 'parentFolderId'
].join(',');

/**
 * Extended fields for reading a single email's full content.
 * Used by read.js.
 */
const EMAIL_DETAIL_FIELDS = [
  'id', 'subject', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients',
  'receivedDateTime', 'isRead', 'hasAttachments', 'importance',
  'body', 'bodyPreview', 'parentFolderId', 'conversationId',
  'internetMessageHeaders'
].join(',');

// ─── Shared Utilities ─────────────────────────────────────────────────

// BEFORE: Email ID was passed directly into URLs without validation in
//         read.js, mark-as-read.js, and potentially others. Malformed IDs
//         could cause unexpected URL construction or opaque Graph 400s.
// AFTER: Shared validateEmailId() checks for non-empty string with no
//        path traversal characters.
// GOOD EFFECT: Consistent input validation across all operations that
//              accept an email ID — clear error messages before the API call.

/**
 * Validates an email/message ID before use in a Graph API URL.
 * @param {string} id - The email ID to validate
 * @returns {{ valid: boolean, error?: string }}
 */
function validateEmailId(id) {
  if (!id || typeof id !== 'string') {
    return { valid: false, error: 'Email ID is required and must be a string.' };
  }
  if (id.trim().length === 0) {
    return { valid: false, error: 'Email ID cannot be empty.' };
  }
  // Guard against path traversal characters
  if (/[\/\\]/.test(id)) {
    return { valid: false, error: 'Email ID contains invalid characters.' };
  }
  return { valid: true };
}

// BEFORE: Recipient parsing (splitting comma-separated emails) was
//         duplicated in send.js and draft.js with no format validation.
//         Malformed emails caused silent failures or cryptic Graph errors.
// AFTER: Shared parseRecipients() with basic RFC 5322 format validation.
// GOOD EFFECT: Invalid email addresses are caught before the API call with
//              a clear error message; no duplication between send and draft.

/**
 * Parses a comma-separated string of email addresses into Graph API recipient format.
 * Validates each address against a basic email regex.
 * @param {string} recipientString - Comma-separated email addresses
 * @returns {{ recipients: Array, invalidAddresses: string[] }}
 */
function parseRecipients(recipientString) {
  if (!recipientString || typeof recipientString !== 'string') {
    return { recipients: [], invalidAddresses: [] };
  }

  // Basic email validation regex (RFC 5322 simplified)
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const recipients = [];
  const invalidAddresses = [];

  recipientString.split(',').forEach(raw => {
    const email = raw.trim();
    if (!email) return;

    if (emailRegex.test(email)) {
      recipients.push({ emailAddress: { address: email } });
    } else {
      invalidAddresses.push(email);
    }
  });

  return { recipients, invalidAddresses };
}

// BEFORE: mark-as-read.js implemented its own PATCH call to update a
//         single property. Any future patch operations (e.g., flag, move)
//         would each re-implement the same pattern.
// AFTER: Shared patchMessage(accessToken, messageId, properties) utility.
// GOOD EFFECT: One tested function for all message PATCH operations;
//              future flag/categorise/move operations call the same utility.

/**
 * Patches a message with the given properties via the Graph API.
 * @param {Function} callGraphAPI - The Graph API call function (injected)
 * @param {string} accessToken - Valid access token
 * @param {string} messageId - Message ID to patch
 * @param {object} properties - Properties to update
 * @returns {Promise<object>} - Updated message object
 */
async function patchMessage(callGraphAPI, accessToken, messageId, properties) {
  const validation = validateEmailId(messageId);
  if (!validation.valid) {
    throw new Error(validation.error);
  }
  const endpoint = `me/messages/${encodeURIComponent(messageId)}`;
  return callGraphAPI(accessToken, 'PATCH', endpoint, properties);
}


// ─── Tool Definitions ─────────────────────────────────────────────────
const emailTools = [
  {
    name: "list-emails",
    description: "Lists recent emails from your inbox",
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "Email folder to list (e.g., 'inbox', 'sent', 'drafts', default: 'inbox')"
        },
        count: {
          type: "number",
          description: "Number of emails to retrieve (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleListEmails
  },
  {
    name: "search-emails",
    description: "Search for emails using various criteria",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query text to find in emails"
        },
        folder: {
          type: "string",
          description: "Email folder to search in (default: 'inbox')"
        },
        from: {
          type: "string",
          description: "Filter by sender email address or name"
        },
        to: {
          type: "string",
          description: "Filter by recipient email address or name"
        },
        subject: {
          type: "string",
          description: "Filter by email subject"
        },
        hasAttachments: {
          type: "boolean",
          description: "Filter to only emails with attachments"
        },
        unreadOnly: {
          type: "boolean",
          description: "Filter to only unread emails"
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleSearchEmails
  },
  {
    name: "read-email",
    description: "Reads the content of a specific email. HTML emails are securely sanitized to extract only visible text, preventing prompt injection attacks via hidden content.",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to read"
        },
        includeRawHtml: {
          type: "boolean",
          description: "Include raw HTML content (UNSAFE - for debugging only, may contain hidden prompt injection content)"
        }
      },
      required: ["id"]
    },
    handler: handleReadEmail
  },
  {
    name: "send-email",
    description: "Composes and sends a new email. Supports both plain text and HTML content.",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Email subject"
        },
        body: {
          type: "string",
          description: "Email body content (plain text or HTML)"
        },
        isHtml: {
          type: "boolean",
          description: "Set to true to send as HTML, false for plain text. If not specified, auto-detects based on <html> tag presence."
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        },
        saveToSentItems: {
          type: "boolean",
          description: "Whether to save the email to sent items"
        },
        // BEFORE: No preview/dry-run parameter.
        // AFTER: Added preview mode parameter.
        // GOOD EFFECT: LLM can inspect the constructed payload before
        //              committing to an irreversible send.
        preview: {
          type: "boolean",
          description: "If true, returns the constructed email payload without sending (dry-run mode)"
        }
      },
      required: ["to", "subject", "body"]
    },
    handler: handleSendEmail
  },
  {
    name: "draft-email",
    description: "Creates and saves an email draft in Outlook",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Draft email subject"
        },
        body: {
          type: "string",
          description: "Draft email body content (can be plain text or HTML)"
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        }
      },
      required: []
    },
    handler: handleDraftEmail
  },
  {
    name: "mark-as-read",
    description: "Marks an email as read or unread",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to mark as read/unread"
        },
        isRead: {
          type: "boolean",
          description: "Whether to mark as read (true) or unread (false). Default: true"
        }
      },
      required: ["id"]
    },
    handler: handleMarkAsRead
  }
];

module.exports = {
  emailTools,
  handleListEmails,
  handleSearchEmails,
  handleReadEmail,
  handleSendEmail,
  handleDraftEmail,
  handleMarkAsRead,
  // BEFORE: No shared utilities exported from barrel.
  // AFTER: Shared constants and utilities available to all email operations.
  // GOOD EFFECT: All email ops import from the barrel, not from scattered files.
  EMAIL_SELECT_FIELDS,
  EMAIL_DETAIL_FIELDS,
  validateEmailId,
  parseRecipients,
  patchMessage
};