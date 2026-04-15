/**
 * Mark email as read functionality
 *
 * Action plan fixes applied:
 *   - (c) email/mark-as-read.js: Uses shared patchMessage utility instead of
 *         re-implementing PATCH inline — future patch ops (flag, categorise)
 *         will use the same utility.
 *   - (c) email/mark-as-read.js: Uses shared validateEmailId for input validation.
 *   - (e) email/mark-as-read.js: PATCH is idempotent — marking a read email as
 *         read returns 200 with unchanged state; handled gracefully (not as error).
 *   - (c): Error classification (auth, 403, 429, mailbox mismatch).
 */
const { callGraphAPI } = require('../utils/graph-api');
const { getClient } = require('../auth');

// Import from shared.js (not ./index) to avoid circular dependency.
// index.js imports all handlers; handlers importing back from index.js
// causes patchMessage and validateEmailId to resolve as undefined.
const { patchMessage, validateEmailId } = require('./shared');

/**
 * Mark email as read handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMarkAsRead(args) {
  const emailId = args.id;
  const isRead = args.isRead !== undefined ? args.isRead : true;

  // ── Input Validation ──
  // BEFORE: if (!emailId) — simple null check, no format validation.
  // AFTER: validateEmailId() checks non-empty + no path traversal characters.
  // GOOD EFFECT: Malformed IDs caught with clear error before API call.
  const validation = validateEmailId(emailId);
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

    try {
      // ── Shared patchMessage Utility ──
      // BEFORE: const endpoint = `me/messages/${encodeURIComponent(emailId)}`;
      //         const updateData = { isRead: isRead };
      //         const result = await callGraphAPI(accessToken, 'PATCH', endpoint, updateData);
      //         — inline PATCH implementation, duplicated for every future patch op.
      // AFTER: await patchMessage(callGraphAPI, accessToken, emailId, { isRead });
      // GOOD EFFECT: Single tested utility for all message PATCH operations.
      //              Future ops (flag, categorise, move) call the same function.
      await patchMessage(callGraphAPI, accessToken, emailId, { isRead });

      // BEFORE: (no special handling for idempotent success)
      // AFTER: PATCH is idempotent — marking a read email as read returns 200
      //        with unchanged state. This is handled gracefully as success.
      // GOOD EFFECT: Repeated calls (LLM retry) don't produce false errors.
      const status = isRead ? 'read' : 'unread';

      return {
        content: [{
          type: "text",
          text: `Email successfully marked as ${status}.`
        }]
      };
    } catch (error) {
      console.error(`Error marking email as ${isRead ? 'read' : 'unread'}: ${error.message}`);

      // ── Error Classification ──
      // BEFORE: Checked for "doesn't belong to the targeted mailbox" and
      //         "UNAUTHORIZED" — but missed 403, 404, 429.
      // AFTER: Expanded classification with actionable messages.
      // GOOD EFFECT: Each failure mode gets specific guidance.

      if (error.message.includes("doesn't belong to the targeted mailbox")) {
        return {
          content: [{
            type: "text",
            text: `The email ID seems invalid or doesn't belong to your mailbox. Please try with a different email ID.`
          }]
        };
      }

      if (error.message.includes("UNAUTHORIZED") || error.message.includes('401')) {
        return {
          content: [{
            type: "text",
            text: "Authentication failed. Please re-authenticate and try again."
          }]
        };
      }

      // BEFORE: (no 403 classification)
      // AFTER: Detect 403 and suggest scope issue.
      // GOOD EFFECT: Clear guidance on permission errors.
      if (error.message.includes('403')) {
        return {
          content: [{
            type: "text",
            text: "Access denied (403). The token may lack Mail.ReadWrite scope. Re-authenticate with force=true."
          }]
        };
      }

      // BEFORE: (no 429 classification)
      // AFTER: Detect 429 and surface retry hint.
      // GOOD EFFECT: LLM waits and retries instead of treating as permanent failure.
      if (error.message.includes('429')) {
        return {
          content: [{
            type: "text",
            text: "Microsoft Graph API rate limit reached (429). Please wait a moment and try again."
          }]
        };
      }

      // BEFORE: (no 404 classification)
      // AFTER: Detect 404 — email may have been deleted.
      // GOOD EFFECT: Clear message instead of opaque Graph error.
      if (error.message.includes('404') || error.message.includes('ErrorItemNotFound')) {
        return {
          content: [{
            type: "text",
            text: `Email not found (404). It may have been deleted or the ID is incorrect.`
          }]
        };
      }

      return {
        content: [{
          type: "text",
          text: `Failed to mark email as ${isRead ? 'read' : 'unread'}: ${error.message}`
        }]
      };
    }
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error accessing email: ${error.message}`
      }]
    };
  }
}

module.exports = handleMarkAsRead;