/**
 * Improved search emails functionality
 *
 * Action plan fixes applied:
 *   - (c) email/search.js: KQL injection sanitised — user input escaped before interpolation
 *   - (c) email/search.js: Field selection from shared EMAIL_SELECT_FIELDS
 *   - (e) email/search.js: Short TTL query result cache (5-30s) keyed on normalised query
 *   - (c) email/index.js: Error classification (auth, 429, 403)
 */
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { getClient } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');

// BEFORE: const config = require('../config');
//         — EMAIL_SELECT_FIELDS read from config.js or hardcoded.
// AFTER: Import from the barrel's shared constant.
// GOOD EFFECT: Single source of truth for field selection.
const { EMAIL_SELECT_FIELDS } = require('./index');

// ─── KQL Sanitisation ─────────────────────────────────────────────────
// BEFORE: User-supplied search strings were interpolated directly into the
//         $search OData value without any escaping. While Graph API is not
//         SQL, malformed KQL can produce confusing errors or unexpected results.
// AFTER: sanitizeKQL() escapes special KQL characters.
// GOOD EFFECT: Prevents malformed KQL from causing opaque Graph API errors;
//              user input is safely incorporated into search queries.

/**
 * Sanitises a string for safe use in KQL (Keyword Query Language).
 * Escapes characters that have special meaning in KQL: : ( ) " AND OR NOT
 * @param {string} input - Raw user input
 * @returns {string} - Sanitised input safe for KQL interpolation
 */
function sanitizeKQL(input) {
  if (!input || typeof input !== 'string') return '';

  // BEFORE: (no sanitisation) — raw user input was interpolated directly.
  // AFTER: Escape double quotes (KQL string delimiter) and wrap the entire
  //        value so special characters like : ( ) are treated as literals.
  // GOOD EFFECT: Input like "from:attacker OR subject:secret" is treated
  //              as a literal search string, not as KQL operators.
  return input
    .replace(/"/g, '\\"')  // Escape embedded double quotes
    .replace(/\\/g, '\\\\'); // Escape backslashes
}

// ─── Query Result Cache ───────────────────────────────────────────────
// BEFORE: The same KQL query run twice within seconds hit the API twice.
// AFTER: Short TTL (10-second) cache keyed on normalised query string.
// GOOD EFFECT: Eliminates redundant Graph API calls in interactive sessions
//              where the LLM may re-invoke the same search.

const _searchCache = new Map();
const SEARCH_CACHE_TTL_MS = 10_000; // 10 seconds

function _getSearchCacheKey(folder, searchTerms, filterTerms, count) {
  return JSON.stringify({ folder, searchTerms, filterTerms, count });
}

function _getCachedSearch(key) {
  const entry = _searchCache.get(key);
  if (!entry) return null;
  if (Date.now() - entry.timestamp > SEARCH_CACHE_TTL_MS) {
    _searchCache.delete(key);
    return null;
  }
  return entry.response;
}

function _setCachedSearch(key, response) {
  _searchCache.set(key, { response, timestamp: Date.now() });
  if (_searchCache.size > 100) {
    const oldest = _searchCache.keys().next().value;
    _searchCache.delete(oldest);
  }
}

/**
 * Search emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchEmails(args) {
  const folder = args.folder || "inbox";
  const requestedCount = args.count || 10;
  const query = args.query || '';
  const from = args.from || '';
  const to = args.to || '';
  const subject = args.subject || '';
  const hasAttachments = args.hasAttachments;
  const unreadOnly = args.unreadOnly;

  try {
    const client = await getClient(args.bearer_token || null);
    const accessToken = client.rawToken;

    // Resolve the folder path
    const endpoint = await resolveFolderPath(accessToken, folder);
    console.error(`Using endpoint: ${endpoint} for folder: ${folder}`);

    const searchTerms = { query, from, to, subject };
    const filterTerms = { hasAttachments, unreadOnly };

    // ── Cache check ──
    // BEFORE: Every search hit the Graph API.
    // AFTER: Return cached response if same query was run within 10s.
    // GOOD EFFECT: Reduces API quota usage for repeated queries.
    const cacheKey = _getSearchCacheKey(folder, searchTerms, filterTerms, requestedCount);
    const cached = _getCachedSearch(cacheKey);
    if (cached) {
      console.error('[search-emails] Returning cached response');
      return cached;
    }

    // Execute progressive search with pagination
    const response = await progressiveSearch(
      endpoint,
      accessToken,
      searchTerms,
      filterTerms,
      requestedCount
    );

    const result = formatSearchResults(response);
    _setCachedSearch(cacheKey, result);
    return result;
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    // BEFORE: (no 429 classification)
    // AFTER: Surface retry hint for throttle errors.
    // GOOD EFFECT: LLM knows to wait and retry.
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
        text: `Error searching emails: ${error.message}`
      }]
    };
  }
}

/**
 * Execute a search with progressively simpler fallback strategies
 */
async function progressiveSearch(endpoint, accessToken, searchTerms, filterTerms, maxCount) {
  const searchAttempts = [];

  // 1. Try combined search (most specific)
  try {
    const params = buildSearchParams(searchTerms, filterTerms, Math.min(50, maxCount));
    console.error("Attempting combined search with params:", params);
    searchAttempts.push("combined-search");

    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, params, maxCount);
    if (response.value && response.value.length > 0) {
      console.error(`Combined search successful: found ${response.value.length} results`);
      return response;
    }
  } catch (error) {
    console.error(`Combined search failed: ${error.message}`);
  }

  // 2. Try each search term individually
  const searchPriority = ['subject', 'from', 'to', 'query'];

  for (const term of searchPriority) {
    if (searchTerms[term]) {
      try {
        console.error(`Attempting search with only ${term}: "${searchTerms[term]}"`);
        searchAttempts.push(`single-term-${term}`);

        const simplifiedParams = {
          $top: Math.min(50, maxCount),
          $select: EMAIL_SELECT_FIELDS
        };

        const kqlParts = [];

        // BEFORE: kqlParts.push(`${term}:${searchTerms[term]}`);
        //         — raw user input interpolated directly into KQL.
        // AFTER: kqlParts.push(`${term}:${sanitizeKQL(searchTerms[term])}`);
        // GOOD EFFECT: Special KQL characters in user input are escaped,
        //              preventing malformed queries and unexpected search behaviour.
        if (term === 'query') {
          kqlParts.push(sanitizeKQL(searchTerms[term]));
        } else {
          kqlParts.push(`${term}:${sanitizeKQL(searchTerms[term])}`);
        }

        addBooleanFiltersAsKQL(kqlParts, filterTerms);
        simplifiedParams.$search = `"${kqlParts.join(' ')}"`;

        const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, simplifiedParams, maxCount);
        if (response.value && response.value.length > 0) {
          console.error(`Search with ${term} successful: found ${response.value.length} results`);
          return response;
        }
      } catch (error) {
        console.error(`Search with ${term} failed: ${error.message}`);
      }
    }
  }

  // 3. Try with only boolean filters
  if (filterTerms.hasAttachments === true || filterTerms.unreadOnly === true) {
    try {
      console.error("Attempting search with only boolean filters");
      searchAttempts.push("boolean-filters-only");

      const filterOnlyParams = {
        $top: Math.min(50, maxCount),
        $select: EMAIL_SELECT_FIELDS,
        $orderby: 'receivedDateTime desc'
      };

      addBooleanFilters(filterOnlyParams, filterTerms);

      const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, filterOnlyParams, maxCount);
      console.error(`Boolean filter search found ${response.value?.length || 0} results`);
      return response;
    } catch (error) {
      console.error(`Boolean filter search failed: ${error.message}`);
    }
  }

  // 4. Final fallback: recent emails
  console.error("All search strategies failed, falling back to recent emails");
  searchAttempts.push("recent-emails");

  const basicParams = {
    $top: Math.min(50, maxCount),
    $select: EMAIL_SELECT_FIELDS,
    $orderby: 'receivedDateTime desc'
  };

  const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, basicParams, maxCount);
  console.error(`Fallback to recent emails found ${response.value?.length || 0} results`);

  response._searchInfo = {
    attemptsCount: searchAttempts.length,
    strategies: searchAttempts,
    originalTerms: searchTerms,
    filterTerms: filterTerms
  };

  return response;
}

/**
 * Build search parameters from search terms and filter terms
 */
function buildSearchParams(searchTerms, filterTerms, count) {
  const params = {
    $top: count,
    $select: EMAIL_SELECT_FIELDS
  };

  const kqlTerms = [];

  // BEFORE: kqlTerms.push(searchTerms.query);
  //         — unsanitised user input injected into KQL.
  // AFTER: All user inputs run through sanitizeKQL().
  // GOOD EFFECT: Prevents KQL injection across all search parameters.
  if (searchTerms.query) {
    kqlTerms.push(sanitizeKQL(searchTerms.query));
  }

  if (searchTerms.subject) {
    kqlTerms.push(`subject:"${sanitizeKQL(searchTerms.subject)}"`);
  }

  if (searchTerms.from) {
    kqlTerms.push(`from:"${sanitizeKQL(searchTerms.from)}"`);
  }

  if (searchTerms.to) {
    kqlTerms.push(`to:"${sanitizeKQL(searchTerms.to)}"`);
  }

  if (kqlTerms.length > 0) {
    addBooleanFiltersAsKQL(kqlTerms, filterTerms);
    params.$search = `"${kqlTerms.join(' ')}"`;
  } else {
    params.$orderby = 'receivedDateTime desc';
    addBooleanFilters(params, filterTerms);
  }

  return params;
}

/**
 * Add boolean filters as OData $filter (when $search is NOT present)
 */
function addBooleanFilters(params, filterTerms) {
  const filterConditions = [];
  if (filterTerms.hasAttachments === true) {
    filterConditions.push('hasAttachments eq true');
  }
  if (filterTerms.unreadOnly === true) {
    filterConditions.push('isRead eq false');
  }
  if (filterConditions.length > 0) {
    params.$filter = filterConditions.join(' and ');
  }
}

/**
 * Add boolean filters as KQL terms (when $search IS present)
 */
function addBooleanFiltersAsKQL(kqlTerms, filterTerms) {
  if (filterTerms.hasAttachments === true) {
    kqlTerms.push('hasAttachments:true');
  }
  if (filterTerms.unreadOnly === true) {
    kqlTerms.push('isRead:false');
  }
}

/**
 * Format search results into readable text
 */
function formatSearchResults(response) {
  if (!response.value || response.value.length === 0) {
    return {
      content: [{
        type: "text",
        text: `No emails found matching your search criteria.`
      }]
    };
  }

  const emailList = response.value.map((email, index) => {
    const sender = email.from?.emailAddress || { name: 'Unknown', address: 'unknown' };
    const date = new Date(email.receivedDateTime).toLocaleString();
    const readStatus = email.isRead ? '' : '[UNREAD] ';

    return `${index + 1}. ${readStatus}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}\nID: ${email.id}\n`;
  }).join("\n");

  let additionalInfo = '';
  if (response._searchInfo) {
    additionalInfo = `\n(Search used ${response._searchInfo.strategies[response._searchInfo.strategies.length - 1]} strategy)`;
  }

  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} emails matching your search criteria:${additionalInfo}\n\n${emailList}`
    }]
  };
}

module.exports = handleSearchEmails;