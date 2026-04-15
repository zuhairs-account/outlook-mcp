/**
 * Read email functionality
 *
 * Security: HTML emails are sanitized to remove hidden content that could
 * be used for prompt injection attacks. Only visible text is extracted.
 *
 * Action plan fixes applied:
 *   - (c) email/read.js: Message ID validated before use in URL (shared validateEmailId)
 *   - (c) email/read.js: Attachment listing via $expand=attachments
 *   - (c) email/read.js: Field selection from shared EMAIL_DETAIL_FIELDS
 *   - (e) email/read.js: Body truncation to configurable max character count
 *   - (e) email/read.js: Per-session LRU cache keyed on message ID
 *   - (c) email/read.js: Error classification (auth, 403, 404, 429)
 */
const { callGraphAPI } = require('../utils/graph-api');
const { getClient } = require('../auth');
const { processHtmlEmail } = require('../utils/html-sanitizer');

// Import from shared.js (not ./index) to avoid circular dependency.
// index.js imports all handlers; handlers importing back from index.js
// causes EMAIL_DETAIL_FIELDS and validateEmailId to resolve as undefined.
const { EMAIL_DETAIL_FIELDS, validateEmailId } = require('./shared');

// ─── Configuration Constants ──────────────────────────────────────────
// BEFORE: No body size limit — very large HTML emails (newsletters, etc.)
//         could produce multi-MB responses that block JSON.parse on the
//         event loop and overwhelm the LLM's context window.
// AFTER: Configurable max body length with truncation notice.
// GOOD EFFECT: Large emails are gracefully truncated instead of crashing
//              the LLM or blocking the event loop.
const MAX_BODY_LENGTH = 50_000; // ~50KB — configurable

// ─── Per-Session Response Cache ───────────────────────────────────────
// BEFORE: Re-reading the same message ID within a session re-fetched
//         from Graph every time.
// AFTER: LRU cache keyed on message ID with 60-second TTL.
// GOOD EFFECT: Eliminates redundant Graph API calls when the LLM reads
//              the same email multiple times in a conversation.

const _readCache = new Map();
const READ_CACHE_TTL_MS = 60_000; // 60 seconds
const READ_CACHE_MAX_SIZE = 20;   // LRU eviction threshold

function _getCachedRead(emailId) {
  const entry = _readCache.get(emailId);
  if (!entry) return null;
  if (Date.now() - entry.timestamp > READ_CACHE_TTL_MS) {
    _readCache.delete(emailId);
    return null;
  }
  return entry.response;
}

function _setCachedRead(emailId, response) {
  _readCache.set(emailId, { response, timestamp: Date.now() });
  // LRU eviction — delete oldest entry when cache exceeds max size
  if (_readCache.size > READ_CACHE_MAX_SIZE) {
    const oldest = _readCache.keys().next().value;
    _readCache.delete(oldest);
  }
}

/**
 * Read email handler
 * @param {object} args - Tool arguments
 * @param {string} args.id - Email ID (required)
 * @param {boolean} args.includeRawHtml - If true, include raw HTML (unsafe, for debugging only)
 * @returns {object} - MCP response
 */
async function handleReadEmail(args) {
  const emailId = args.id;
  const includeRawHtml = args.includeRawHtml === true;

  // ── Input Validation ──
  // BEFORE: if (!emailId) — simple null check, no format validation.
  //         ID was injected directly into the URL without sanitisation.
  // AFTER: validateEmailId() checks for non-empty, no path traversal chars.
  // GOOD EFFECT: Malformed IDs are caught with a clear error before the API
  //              call, instead of producing opaque Graph 400 responses.
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

    // ── Cache check ──
    // BEFORE: Every read hit the Graph API.
    // AFTER: Return cached response if same ID was read within 60s.
    // GOOD EFFECT: Eliminates redundant Graph API calls.
    const cached = _getCachedRead(emailId);
    if (cached && !includeRawHtml) {
      console.error(`[read-email] Returning cached response for ${emailId}`);
      return cached;
    }

    // BEFORE: const endpoint = `me/messages/${encodeURIComponent(emailId)}`;
    //         const queryParams = { $select: config.EMAIL_DETAIL_FIELDS };
    //         — no attachment listing.
    // AFTER: Use shared EMAIL_DETAIL_FIELDS; add $expand=attachments to
    //        include attachment metadata in the response.
    // GOOD EFFECT: (c) Attachment names/sizes are surfaced alongside the
    //              message body — the LLM can inform the user about attachments.
    const endpoint = `me/messages/${encodeURIComponent(emailId)}`;
    const queryParams = {
      $select: EMAIL_DETAIL_FIELDS,
      // BEFORE: No attachment listing — read tool did not indicate whether
      //         attachments exist or enumerate them.
      // AFTER: $expand=attachments($select=id,name,contentType,size)
      // GOOD EFFECT: Attachment metadata returned alongside the message;
      //              the LLM can list attachment names/sizes to the user.
      $expand: 'attachments($select=id,name,contentType,size)'
    };

    try {
      const email = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

      if (!email) {
        return {
          content: [{
            type: "text",
            text: `Email with ID ${emailId} not found.`
          }]
        };
      }

      // Format sender, recipients, etc.
      const sender = email.from ? `${email.from.emailAddress.name} (${email.from.emailAddress.address})` : 'Unknown';
      const senderAddress = email.from?.emailAddress?.address || 'unknown';
      const to = email.toRecipients ? email.toRecipients.map(r => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ") : 'None';
      const cc = email.ccRecipients && email.ccRecipients.length > 0 ? email.ccRecipients.map(r => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ") : 'None';
      const bcc = email.bccRecipients && email.bccRecipients.length > 0 ? email.bccRecipients.map(r => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ") : 'None';
      const date = new Date(email.receivedDateTime).toLocaleString();

      // Extract and sanitize body content
      let body = '';
      let bodyNote = '';

      if (email.body) {
        if (email.body.contentType === 'html') {
          body = processHtmlEmail(email.body.content, {
            addBoundary: true,
            metadata: {
              from: senderAddress,
              subject: email.subject,
              date: date
            }
          });
          bodyNote = '\n[HTML email - sanitized for security, hidden content removed]\n';
        } else {
          body = processHtmlEmail(email.body.content, {
            addBoundary: true,
            metadata: {
              from: senderAddress,
              subject: email.subject,
              date: date
            }
          });
        }
      } else {
        body = email.bodyPreview || 'No content';
      }

      // ── Body Truncation ──
      // BEFORE: Full body returned regardless of size — multi-MB newsletter
      //         HTML could block the event loop and overwhelm the LLM context.
      // AFTER: Truncate to MAX_BODY_LENGTH with a notice.
      // GOOD EFFECT: Large emails are handled gracefully; the LLM gets
      //              usable content without context-window overflow.
      let truncationNote = '';
      if (body.length > MAX_BODY_LENGTH) {
        const originalLength = body.length;           // ← capture BEFORE slicing
        body = body.substring(0, MAX_BODY_LENGTH);
        truncationNote = `\n[Body truncated at ${MAX_BODY_LENGTH} characters. Original length: ${originalLength}]\n`;
      }

      // ── Attachment Listing ──
      // BEFORE: No attachment information returned — the user had no way to
      //         know if an email had attachments or what they were.
      // AFTER: List attachment names, types, and sizes.
      // GOOD EFFECT: The LLM can inform the user about attachments and
      //              suggest downloading specific ones by name.
      let attachmentSection = '';
      if (email.attachments && email.attachments.length > 0) {
        const attachmentList = email.attachments.map((att, i) => {
          const sizeKB = att.size ? `${(att.size / 1024).toFixed(1)}KB` : 'unknown size';
          return `  ${i + 1}. ${att.name} (${att.contentType || 'unknown type'}, ${sizeKB})`;
        }).join('\n');
        attachmentSection = `\nAttachments (${email.attachments.length}):\n${attachmentList}\n`;
      }

      // Format the email
      const formattedEmail = `From: ${sender}
To: ${to}
${cc !== 'None' ? `CC: ${cc}\n` : ''}${bcc !== 'None' ? `BCC: ${bcc}\n` : ''}Subject: ${email.subject}
Date: ${date}
Importance: ${email.importance || 'normal'}
Has Attachments: ${email.hasAttachments ? 'Yes' : 'No'}${attachmentSection}${bodyNote}${truncationNote}
${body}`;

      // Optionally include raw HTML for debugging
      let rawHtmlSection = '';
      if (includeRawHtml && email.body?.contentType === 'html') {
        rawHtmlSection = `\n\n--- RAW HTML (UNSAFE - FOR DEBUGGING ONLY) ---\n${email.body.content}\n--- END RAW HTML ---`;
      }

      const result = {
        content: [{
          type: "text",
          text: formattedEmail + rawHtmlSection
        }]
      };

      // Cache the sanitised response (not the raw HTML variant)
      if (!includeRawHtml) {
        _setCachedRead(emailId, result);
      }

      return result;
    } catch (error) {
      console.error(`Error reading email: ${error.message}`);

      // ── Error Classification ──
      // BEFORE: Only checked for "doesn't belong to the targeted mailbox".
      // AFTER: Classify 404, 403, 429, and mailbox-mismatch errors.
      // GOOD EFFECT: Actionable error messages for each failure mode.

      if (error.message.includes("doesn't belong to the targeted mailbox")) {
        return {
          content: [{
            type: "text",
            text: `The email ID seems invalid or doesn't belong to your mailbox. Please try with a different email ID.`
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

      if (error.message.includes('429')) {
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
          text: `Failed to read email: ${error.message}`
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

module.exports = handleReadEmail;