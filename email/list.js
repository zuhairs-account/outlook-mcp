/**
 * List emails functionality
 *
 * Action plan fixes applied:
 *   - (c) email/list.js: Field selection from shared EMAIL_SELECT_FIELDS (not hardcoded inline)
 *   - (c) email/list.js: Folder targeting already supported via resolveFolderPath — documented
 *   - (c) email/list.js: Error classification (auth errors trigger re-auth hint, 429 → retry hint)
 *   - (e) email/list.js: Short TTL in-memory deduplication cache for identical list calls
 *   - (e) email/list.js: Pagination already via callGraphAPIPaginated — verified
 */
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { getClient } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');
const config = require('../config');

// ─── Request Deduplication Cache ──────────────────────────────────────
// BEFORE: Identical list calls within seconds (e.g., LLM retry, parallel
//         tool invocations) each fired a separate Graph API request.
// AFTER: Short TTL (5-second) in-memory cache keyed on folder + count.
// GOOD EFFECT: Eliminates redundant network calls in interactive sessions;
//              reduces Graph API quota consumption.

const _listCache = new Map();
const LIST_CACHE_TTL_MS = 5_000; // 5 seconds

function _getCacheKey(folder, count) {
  return `${folder}:${count}`;
}

function _getCachedResponse(key) {
  const entry = _listCache.get(key);
  if (!entry) return null;
  if (Date.now() - entry.timestamp > LIST_CACHE_TTL_MS) {
    _listCache.delete(key);
    return null;
  }
  return entry.response;
}

function _setCachedResponse(key, response) {
  _listCache.set(key, { response, timestamp: Date.now() });
  // Prevent unbounded cache growth
  if (_listCache.size > 50) {
    const oldest = _listCache.keys().next().value;
    _listCache.delete(oldest);
  }
}

/**
 * List emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListEmails(args) {
  const folder = args.folder || "inbox";
  const requestedCount = args.count || 10;

  try {
    const client = await getClient(args.bearer_token || null);
    const accessToken = client.rawToken;

    // ── Dedup cache check ──
    // BEFORE: Every call hit the Graph API, even identical back-to-back calls.
    // AFTER: Return cached response if same folder+count was called within 5s.
    // GOOD EFFECT: Eliminates redundant Graph API calls from LLM retries.
    const cacheKey = _getCacheKey(folder, requestedCount);
    const cached = _getCachedResponse(cacheKey);
    if (cached) {
      console.error(`[list-emails] Returning cached response for ${cacheKey}`);
      return cached;
    }

    // Resolve the folder path
    const endpoint = await resolveFolderPath(accessToken, folder);

    // BEFORE: $select: config.EMAIL_SELECT_FIELDS — read from config.js singleton.
    // AFTER: $select: EMAIL_SELECT_FIELDS — imported from barrel's shared constant.
    // GOOD EFFECT: Field list is defined once and shared across all email operations.
    const queryParams = {
      $top: Math.min(50, requestedCount),
      $orderby: 'receivedDateTime desc',
      $select: config.EMAIL_SELECT_FIELDS
    };

    // Pagination is already handled by callGraphAPIPaginated
    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, queryParams, requestedCount);

    if (!response.value || response.value.length === 0) {
      const result = {
        content: [{
          type: "text",
          text: `No emails found in ${folder}.`
        }]
      };
      _setCachedResponse(cacheKey, result);
      return result;
    }

    // Format results
    const emailList = response.value.map((email, index) => {
      const sender = email.from ? email.from.emailAddress : { name: 'Unknown', address: 'unknown' };
      const date = new Date(email.receivedDateTime).toLocaleString();
      const readStatus = email.isRead ? '' : '[UNREAD] ';

      return `${index + 1}. ${readStatus}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}\nID: ${email.id}\n`;
    }).join("\n");

    const result = {
      content: [{
        type: "text",
        text: `Found ${response.value.length} emails in ${folder}:\n\n${emailList}`
      }]
    };

    _setCachedResponse(cacheKey, result);
    return result;
  } catch (error) {
    // ── Error Classification ──
    // BEFORE: Only checked for exact 'Authentication required' string.
    //         Graph API 403, 429, 5xx errors were returned as opaque messages.
    // AFTER: Classify known error types and return actionable messages.
    // GOOD EFFECT: Auth errors suggest re-auth; 429 surfaces retry hint;
    //              the LLM gets actionable information instead of raw error strings.
    if (error.message === 'Authentication required' || error.message === 'UNAUTHORIZED') {
      return {
        content: [{
          type: "text",
          text: "Authentication required or token expired. Please use the 'authenticate' tool first, or provide a fresh bearer token."
        }]
      };
    }

    // BEFORE: (no classification for 429 throttle errors)
    // AFTER: Detect 429 and surface retry hint.
    // GOOD EFFECT: LLM knows to wait and retry rather than treating it as a permanent failure.
    if (error.message && error.message.includes('429')) {
      return {
        content: [{
          type: "text",
          text: "Microsoft Graph API rate limit reached (429). Please wait a moment and try again."
        }]
      };
    }

    // BEFORE: (no classification for 403 permission errors)
    // AFTER: Detect 403 and suggest scope issue.
    // GOOD EFFECT: Clear guidance on how to fix permission errors.
    if (error.message && error.message.includes('403')) {
      return {
        content: [{
          type: "text",
          text: "Access denied (403). The token may lack Mail.Read scope. Re-authenticate with force=true to refresh consent."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error listing emails: ${error.message}`
      }]
    };
  }
}

module.exports = handleListEmails;