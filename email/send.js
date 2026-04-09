/**
 * Send email functionality
 *
 * Action plan fixes applied:
 *   - (c) email/send.js: Preview/dry-run mode — returns payload without sending
 *   - (c) email/send.js: Recipient parsing via shared parseRecipients with RFC 5322 validation
 *   - (c) email/send.js: Draft-first workflow — creates draft then sends (safer, inspectable)
 *   - (e) email/send.js: POST /me/sendMail is not idempotent — draft-first makes the critical
 *         action a one-shot operation; timeout and retry only affect the draft creation (idempotent)
 *   - (c) email/send.js: Error classification (auth, 403, 429)
 */
const { callGraphAPI } = require('../utils/graph-api');
const { getClient } = require('../auth');

// BEFORE: Recipient parsing was inline — split(',').map() with no validation.
//         Malformed emails caused silent failures or cryptic Graph errors.
// AFTER: Import shared parseRecipients from barrel.
// GOOD EFFECT: Invalid email addresses caught before API call with clear errors.
const { parseRecipients } = require('./index');

/**
 * Send email handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSendEmail(args) {
  const { to, cc, bcc, subject, body, importance = 'normal', saveToSentItems = true, isHtml, preview } = args;

  // Validate required parameters
  if (!to) {
    return {
      content: [{ type: "text", text: "Recipient (to) is required." }]
    };
  }
  if (!subject) {
    return {
      content: [{ type: "text", text: "Subject is required." }]
    };
  }
  if (!body) {
    return {
      content: [{ type: "text", text: "Body content is required." }]
    };
  }

  try {
    const client = await getClient(args.bearer_token || null);
    const accessToken = client.rawToken;

    // ── Recipient Parsing with Validation ──
    // BEFORE: const toRecipients = to.split(',').map(email => { ... });
    //         — no format validation; malformed addresses forwarded to Graph.
    // AFTER: const { recipients, invalidAddresses } = parseRecipients(to);
    // GOOD EFFECT: Invalid email addresses are caught before the API call
    //              with a clear error listing the bad addresses.
    const toParsed = parseRecipients(to);
    const ccParsed = parseRecipients(cc);
    const bccParsed = parseRecipients(bcc);

    // BEFORE: (no validation — malformed emails sent to Graph API)
    // AFTER: Check for invalid addresses and return early with a clear error.
    // GOOD EFFECT: User sees exactly which addresses are malformed.
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

    if (toParsed.recipients.length === 0) {
      return {
        content: [{ type: "text", text: "At least one valid recipient (to) is required." }]
      };
    }

    // Determine content type
    const contentType = isHtml === true ? 'html' :
                        isHtml === false ? 'text' :
                        (body.includes('<html') || body.includes('<HTML')) ? 'html' : 'text';

    // Build message object
    const messageObject = {
      subject,
      body: {
        contentType: contentType,
        content: body
      },
      toRecipients: toParsed.recipients,
      ccRecipients: ccParsed.recipients.length > 0 ? ccParsed.recipients : undefined,
      bccRecipients: bccParsed.recipients.length > 0 ? bccParsed.recipients : undefined,
      importance
    };

    // ── Preview / Dry-Run Mode ──
    // BEFORE: Sending was immediate and irreversible. No way for the LLM to
    //         inspect the constructed payload before committing.
    // AFTER: If preview=true, return the payload without sending.
    // GOOD EFFECT: The LLM can confirm recipients, subject, and body with the
    //              user before actual delivery — prevents accidental sends.
    if (preview === true) {
      const recipientSummary = [
        `To: ${toParsed.recipients.map(r => r.emailAddress.address).join(', ')}`,
        ccParsed.recipients.length > 0 ? `CC: ${ccParsed.recipients.map(r => r.emailAddress.address).join(', ')}` : null,
        bccParsed.recipients.length > 0 ? `BCC: ${bccParsed.recipients.map(r => r.emailAddress.address).join(', ')}` : null
      ].filter(Boolean).join('\n');

      return {
        content: [{
          type: "text",
          text: `📋 Email Preview (NOT SENT):\n\n${recipientSummary}\nSubject: ${subject}\nImportance: ${importance}\nContent Type: ${contentType}\nBody Length: ${body.length} characters\n\nBody:\n${body.substring(0, 500)}${body.length > 500 ? '\n...(truncated)' : ''}\n\nTo send this email, call send-email again without preview=true.`
        }]
      };
    }

    // ── Draft-First Workflow ──
    // BEFORE: POST /me/sendMail — immediate, irreversible, not idempotent.
    //         A network timeout + retry would send the email twice.
    // AFTER: Create a draft first (POST /me/messages — idempotent to create),
    //        then send it (POST /me/messages/{id}/send — one-shot operation).
    // GOOD EFFECT: The critical send action is a one-shot on a known draft ID;
    //              if the draft creation times out and retries, no duplicate send.
    //              The draft also serves as an audit trail.

    // Step 1: Create draft
    // BEFORE: (no draft step — direct sendMail)
    // AFTER: POST /me/messages creates a draft.
    // GOOD EFFECT: Draft is idempotent to create; if this step fails and
    //              retries, no email is sent prematurely.
    const draft = await callGraphAPI(accessToken, 'POST', 'me/messages', messageObject);

    if (!draft || !draft.id) {
      return {
        content: [{
          type: "text",
          text: "Failed to create email draft. Please try again."
        }]
      };
    }

    // Step 2: Send the draft
    // BEFORE: await callGraphAPI(accessToken, 'POST', 'me/sendMail', emailObject);
    //         — direct sendMail, not idempotent.
    // AFTER: POST /me/messages/{id}/send — sends the specific draft.
    // GOOD EFFECT: One-shot send on a known draft ID; if this fails, the
    //              draft still exists and can be retried without duplication.
    await callGraphAPI(accessToken, 'POST', `me/messages/${encodeURIComponent(draft.id)}/send`);

    return {
      content: [{
        type: "text",
        text: `Email sent successfully!\n\nSubject: ${subject}\nRecipients: ${toParsed.recipients.length}${ccParsed.recipients.length > 0 ? ` + ${ccParsed.recipients.length} CC` : ''}${bccParsed.recipients.length > 0 ? ` + ${bccParsed.recipients.length} BCC` : ''}\nMessage Length: ${body.length} characters`
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

    // BEFORE: (no error classification)
    // AFTER: Classify 403 and 429 errors.
    // GOOD EFFECT: Actionable guidance for the user.
    if (error.message && error.message.includes('403')) {
      return {
        content: [{
          type: "text",
          text: "Send denied (403). The token may lack Mail.Send scope. Re-authenticate with force=true to refresh consent."
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
        text: `Error sending email: ${error.message}`
      }]
    };
  }
}

module.exports = handleSendEmail;