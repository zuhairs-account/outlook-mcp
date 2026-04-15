/**
 * Shared constants and utilities for the email module.
 *
 * Extracted from index.js to break the circular dependency that caused
 * EMAIL_SELECT_FIELDS, validateEmailId, parseRecipients, and patchMessage
 * to resolve as undefined when handler files imported from ./index.
 *
 * Handler files (search.js, read.js, send.js, draft.js, mark-as-read.js)
 * now import from ./shared instead of ./index, eliminating the cycle.
 */

// ─── Shared Field Selection Constants ────────────────────────────────
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

module.exports = {
  EMAIL_SELECT_FIELDS,
  EMAIL_DETAIL_FIELDS,
  validateEmailId,
  parseRecipients,
  patchMessage
};