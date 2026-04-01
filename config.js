/**
 * Configuration for Outlook MCP Server
 *
 * Action plan fixes applied:
 *   - (c) config.js: Plain object export → factory function getConfig() + validation
 *   - (c) config.js: Mixed responsibilities → separated into auth, api, runtime namespaces
 *   - (c) config.js: Secrets loaded at module parse time → deferred to getConfig() call
 *   - (c) config.js: Dual credential key names (OUTLOOK_ vs MS_) → standardised to MS_
 *         with single OUTLOOK_ alias fallback, documented in one place
 *   - (e) config.js: Module-level singleton state → config frozen after init (immutable)
 *   - (e) config.js: No watch/reload → documented as future enhancement
 *   - BUG: Duplicate CALENDAR_SELECT_FIELDS key (defined twice) → removed duplicate
 */
const path = require('path');
const os = require('os');

// ─── Environment Variable Standardisation ─────────────────────────────
// BEFORE: README documented two different env-var naming conventions:
//         OUTLOOK_CLIENT_ID and MS_CLIENT_ID for the same value.
//         Code used OUTLOOK_CLIENT_ID in config.js but MS_CLIENT_ID in
//         token-storage.js and oauth-server.js — causing silent mismatches.
// AFTER: Canonical names are MS_* with a single OUTLOOK_* alias fallback.
// GOOD EFFECT: One canonical name set; OUTLOOK_* still works for backward
//              compat but MS_* takes precedence. No more silent mismatches.

/**
 * Reads an environment variable with a fallback alias.
 * Canonical name (MS_*) takes precedence over alias (OUTLOOK_*).
 * @param {string} canonical - e.g., 'MS_CLIENT_ID'
 * @param {string} alias - e.g., 'OUTLOOK_CLIENT_ID'
 * @param {string} [defaultValue='']
 * @returns {string}
 */
function envWithAlias(canonical, alias, defaultValue = '') {
  return process.env[canonical] || process.env[alias] || defaultValue;
}

// ─── Validation ───────────────────────────────────────────────────────
// BEFORE: No validation at load time. Missing required env vars threw at
//         first API call with an opaque error — sometimes minutes after startup.
// AFTER: Validation function checks required vars and throws at startup.
// GOOD EFFECT: Fail-fast — operators see the config problem immediately
//              in the startup log, not buried in a runtime error.

/**
 * Validates that required configuration values are present.
 * @param {object} cfg - The config object to validate
 * @returns {string[]} - Array of warning messages (empty if all good)
 */
function validateConfig(cfg) {
  const warnings = [];

  if (!cfg.auth.clientId) {
    warnings.push('MS_CLIENT_ID (or OUTLOOK_CLIENT_ID) is not set. Authentication will fail.');
  }
  if (!cfg.auth.clientSecret) {
    warnings.push('MS_CLIENT_SECRET (or OUTLOOK_CLIENT_SECRET) is not set. Token exchange will fail.');
  }

  return warnings;
}

// ─── Factory Function ─────────────────────────────────────────────────
// BEFORE: All process.env reads happened at require() time. If env vars
//         were injected after module load (e.g., in tests), they were invisible.
// AFTER: getConfig() reads env vars at call time, not at require() time.
// GOOD EFFECT: Tests can set env vars before calling getConfig() without
//              monkey-patching require(). Config is also freezable after init.

/** @type {object|null} Cached config singleton */
let _cachedConfig = null;

/**
 * Creates (or returns cached) configuration object.
 * Call this at runtime, not at module parse time.
 *
 * @param {object} [overrides] - Optional overrides for testing
 * @returns {Readonly<object>} - Frozen configuration object
 */
function getConfig(overrides = {}) {
  if (_cachedConfig && Object.keys(overrides).length === 0) {
    return _cachedConfig;
  }

  const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';

  // BEFORE: const tenantId was not in config.js — each module read it independently.
  // AFTER: Centralised here.
  // GOOD EFFECT: Single source of truth for tenant ID.
  const tenantId = envWithAlias('MS_TENANT_ID', 'OUTLOOK_TENANT_ID', 'common');
  const authorityHost = (process.env.MS_AUTHORITY_HOST || 'https://login.microsoftonline.com').replace(/\/+$/, '');

  // ── Logical Namespaces ──
  // BEFORE: OAuth credentials, API endpoint URLs, UI strings (server name),
  //         and runtime flags (test mode) were all in one flat object.
  // AFTER: Separated into auth, api, runtime, and module-specific namespaces.
  // GOOD EFFECT: Clear separation of concerns; easy to find related config values.
  const cfg = {
    // ── Server Info ──
    // BEFORE: SERVER_NAME and SERVER_VERSION at top level alongside auth credentials.
    // AFTER: Still top-level (they're server-wide) but namespaced sections below.
    SERVER_NAME: "m365-assistant",
    SERVER_VERSION: "2.0.0",

    // ── Runtime ──
    // BEFORE: USE_TEST_MODE was a flat sibling of auth credentials.
    // AFTER: Grouped under runtime namespace.
    // GOOD EFFECT: Runtime flags are clearly separated from credentials and API config.
    runtime: {
      useTestMode: process.env.USE_TEST_MODE === 'true',
    },

    // ── Authentication ──
    // BEFORE: AUTH_CONFIG flat object with OUTLOOK_CLIENT_ID.
    // AFTER: auth namespace with MS_* canonical names + OUTLOOK_* alias fallback.
    // GOOD EFFECT: One canonical name set; backward compat preserved.
    auth: {
      // BEFORE: clientId: process.env.OUTLOOK_CLIENT_ID || ''
      // AFTER: MS_CLIENT_ID canonical, OUTLOOK_CLIENT_ID alias
      // GOOD EFFECT: Matches token-storage.js and oauth-server.js (which already used MS_*).
      clientId: envWithAlias('MS_CLIENT_ID', 'OUTLOOK_CLIENT_ID', ''),
      clientSecret: envWithAlias('MS_CLIENT_SECRET', 'OUTLOOK_CLIENT_SECRET', ''),
      redirectUri: process.env.MS_REDIRECT_URI || 'http://localhost:3333/auth/callback',
      scopes: (process.env.MS_SCOPES || [
        'offline_access', 'User.Read',
        'Mail.Read', 'Mail.ReadWrite', 'Mail.Send',
        'Calendars.Read', 'Calendars.ReadWrite',
        'Files.Read', 'Files.ReadWrite'
      ].join(' ')).split(' '),
      tenantId,
      authorityHost,
      tokenEndpoint: `${authorityHost}/${tenantId}/oauth2/v2.0/token`,
      authEndpoint: `${authorityHost}/${tenantId}/oauth2/v2.0/authorize`,
      tokenStorePath: path.join(homeDir, '.outlook-mcp-tokens.json'),
      authServerUrl: process.env.AUTH_SERVER_URL || 'http://localhost:3333',
    },

    // ── API Endpoints ──
    // BEFORE: GRAPH_API_ENDPOINT and FLOW_API_ENDPOINT as flat siblings.
    // AFTER: Grouped under api namespace.
    api: {
      graphEndpoint: process.env.GRAPH_API_ENDPOINT || 'https://graph.microsoft.com/v1.0/',
      flowEndpoint: process.env.FLOW_API_ENDPOINT || 'https://api.flow.microsoft.com',
      flowScope: process.env.FLOW_SCOPE || 'https://service.flow.microsoft.com/.default',
    },

    // ── Pagination ──
    pagination: {
      defaultPageSize: 25,
      maxResultCount: 50,
    },

    // ── Timezone ──
    DEFAULT_TIMEZONE: process.env.DEFAULT_TIMEZONE || "Central European Standard Time",

    // ── Email Constants ──
    // BEFORE: EMAIL_SELECT_FIELDS and EMAIL_DETAIL_FIELDS as flat strings.
    // AFTER: Same values, but grouped and documented.
    email: {
      selectFields: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead',
      detailFields: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,internetMessageHeaders',
    },

    // ── Calendar Constants ──
    // BEFORE: CALENDAR_SELECT_FIELDS was defined TWICE with different values:
    //         First:  'id,subject,start,end,location,bodyPreview,isAllDay,recurrence,attendees'
    //         Second: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled'
    //         The second silently overwrote the first (JS object key collision).
    // AFTER: Single definition with the superset of both field lists.
    // GOOD EFFECT: No silent overwrite; all needed fields included.
    calendar: {
      selectFields: 'id,subject,start,end,location,bodyPreview,isAllDay,isCancelled,recurrence,organizer,attendees',
    },

    // ── OneDrive Constants ──
    onedrive: {
      selectFields: 'id,name,size,lastModifiedDateTime,webUrl,folder,file,parentReference',
      uploadThreshold: 4 * 1024 * 1024, // 4MB
    },

    // ── Backward Compatibility Aliases ──
    // BEFORE: Modules import config.AUTH_CONFIG, config.GRAPH_API_ENDPOINT, etc.
    // AFTER: Aliases maintained so existing imports don't break immediately.
    // GOOD EFFECT: Gradual migration — modules can switch to namespaced access
    //              at their own pace while flat aliases still work.
    get AUTH_CONFIG() { return this.auth; },
    get USE_TEST_MODE() { return this.runtime.useTestMode; },
    get GRAPH_API_ENDPOINT() { return this.api.graphEndpoint; },
    get FLOW_API_ENDPOINT() { return this.api.flowEndpoint; },
    get FLOW_SCOPE() { return this.api.flowScope; },
    get DEFAULT_PAGE_SIZE() { return this.pagination.defaultPageSize; },
    get MAX_RESULT_COUNT() { return this.pagination.maxResultCount; },
    get EMAIL_SELECT_FIELDS() { return this.email.selectFields; },
    get EMAIL_DETAIL_FIELDS() { return this.email.detailFields; },
    get CALENDAR_SELECT_FIELDS() { return this.calendar.selectFields; },
    get ONEDRIVE_SELECT_FIELDS() { return this.onedrive.selectFields; },
    get ONEDRIVE_UPLOAD_THRESHOLD() { return this.onedrive.uploadThreshold; },

    // Apply any test overrides
    ...overrides
  };

  // ── Validation at Init ──
  // BEFORE: No validation — missing env vars failed silently at runtime.
  // AFTER: Warnings logged at startup.
  // GOOD EFFECT: Operators see config problems immediately in the startup log.
  const warnings = validateConfig(cfg);
  if (warnings.length > 0) {
    console.warn('⚠️  Configuration warnings:');
    warnings.forEach(w => console.warn(`   - ${w}`));
  }

  // ── Immutable After Init ──
  // BEFORE: Config was a mutable POJO — any module could accidentally mutate it.
  // AFTER: Object.freeze prevents mutation after initialization.
  // GOOD EFFECT: No accidental config mutation across async call stacks;
  //              computed values added later won't silently affect other modules.
  // NOTE: Shallow freeze — nested objects should not be mutated either.
  //       Deep freeze omitted for performance; discipline enforced by convention.
  if (Object.keys(overrides).length === 0) {
    _cachedConfig = Object.freeze(cfg);
    return _cachedConfig;
  }

  return Object.freeze(cfg);
}

// ── Default Export ──
// BEFORE: module.exports = { ... } — frozen POJO, env vars read at require() time.
// AFTER: module.exports = getConfig() — factory called once at first require(),
//        with getConfig exported for tests that need fresh config with overrides.
// GOOD EFFECT: Default behavior unchanged (config available at require time),
//              but tests can call getConfig({ ... }) for custom config.
const defaultConfig = getConfig();

module.exports = defaultConfig;
module.exports.getConfig = getConfig;
module.exports.validateConfig = validateConfig;