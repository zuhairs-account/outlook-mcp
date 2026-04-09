/**
 * Draft email functionality
 *
 * Action plan fixes applied:
 *   - (c) email/send.js pattern: Recipient parsing via shared parseRecipients with validation
 *   - (c) email/send.js pattern: Error classification (auth, 403, 429)
 *   - (c): Consistent recipient validation prevents silent failures from malformed addresses
 */
const { callGraphAPI } = require('../utils/graph-api');
const { getClient } = require('../auth');

// BEFORE: Recipient parsing was inline — to.split(',').map(...) with no validation.
//         Identical duplication of the pattern in send.js.
// AFTER: Import shared parseRecipients from barrel.
// GOOD EFFECT: Invalid email addresses caught before API call; no duplication
//              with send.js — both files use the same validated parsing.
const { parseRecipients } = require('./index');

/**
 * Draft email handler
 * Creates a draft in Outlook using Microsoft Graph:
 * POST /me/messages
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDraftEmail(args) {
  const { to, cc, bcc, subject = '', body = '', importance = 'normal' } = args || {};

  try {
    const client = await getClient(args.bearer_token || null);
    const accessToken = client.rawToken;

    // ── Recipient Parsing with Validation ──
    // BEFORE: const toRecipients = to
    //           ? to.split(',').map(email => ({
    //               emailAddress: { address: email.trim() }
    //             })).filter(r => r.emailAddress.address)
    //           : [];
    //         — duplicated inline parsing with no format validation.
    // AFTER: const toParsed = parseRecipients(to);
    // GOOD EFFECT: RFC 5322 validation catches malformed addresses before
    //              the API call; shared utility eliminates duplication with send.js.
    const toParsed = parseRecipients(to);
    const ccParsed = parseRecipients(cc);
    const bccParsed = parseRecipients(bcc);

    // BEFORE: (no validation — malformed emails forwarded to Graph API)
    // AFTER: Check for invalid addresses and return early with a clear error.
    // GOOD EFFECT: User sees exactly which addresses are malformed instead
    //              of getting a cryptic Graph API 400 error.
    const allInvalid = [
      ...toParsed.invalidAddresses.map(a => `To: ${a}`),
      ...ccParsed.invalidAddresses.map(a => `CC: ${a}`),
      ...bccParsed.invalidAddresses.map(a => `BCC: ${a}`)
    ];
    if (allInvalid.length > 0) {
      return {
        content: [{
          type: "text",
          text: `Invalid email addresses detected:\n${allInvalid.join('\n')}\n\nPlease correct these addresses and try again.`
        }]
      };
    }

    // Create message payload for draft creation
    const messageObject = {
      subject,
      body: {
        contentType: typeof body === 'string' && body.toLowerCase().includes('<html') ? 'html' : 'text',
        content: body
      },
      toRecipients: toParsed.recipients.length > 0 ? toParsed.recipients : undefined,
      ccRecipients: ccParsed.recipients.length > 0 ? ccParsed.recipients : undefined,
      bccRecipients: bccParsed.recipients.length > 0 ? bccParsed.recipients : undefined,
      importance
    };

    // Create draft message
    const draft = await callGraphAPI(accessToken, 'POST', 'me/messages', messageObject);

    return {
      content: [{
        type: "text",
        text: `Draft created successfully!\n\nDraft ID: ${draft.id}\nSubject: ${draft.subject || '(no subject)'}\nRecipients: ${toParsed.recipients.length}${ccParsed.recipients.length > 0 ? ` + ${ccParsed.recipients.length} CC` : ''}${bccParsed.recipients.length > 0 ? ` + ${bccParsed.recipients.length} BCC` : ''}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    // BEFORE: if (error.message && error.message.includes('status 403')) — partial classification.
    // AFTER: Expanded classification including 429 throttle errors.
    // GOOD EFFECT: Actionable error messages for each failure mode.
    if (error.message && error.message.includes('403')) {
      return {
        content: [{
          type: "text",
          text: "Draft creation was denied by Microsoft Graph (403). The token likely lacks Mail.ReadWrite scope. Re-authenticate with force=true to refresh consent, then try again."
        }]
      };
    }

    // BEFORE: (no 429 classification)
    // AFTER: Detect 429 and surface retry hint.
    // GOOD EFFECT: LLM knows to wait and retry rather than treating as permanent failure.
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
        text: `Error creating draft email: ${error.message}`
      }]
    };
  }
}

module.exports = handleDraftEmail;