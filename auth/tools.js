/**
 * Authentication-related tools for the Outlook MCP server
 *
 * Action plan fixes applied:
 *   - (c): URL construction extracted to shared buildAuthUrl() utility
 *   - (c): Tool schema separated from handler business logic
 *   - (c): Scopes read from config, not hardcoded
 *   - (e): PKCE support added (code_verifier / code_challenge)
 */

const crypto = require('crypto');
const querystring = require('querystring');

// ─── Shared Auth URL Builder ──────────────────────────────────────────
// BEFORE: Auth URL was constructed inline in handleAuthenticate(), and
//         separately (with the same params) in oauth-server.js.
// AFTER: Single buildAuthUrl(config, options) function used by both.
// GOOD EFFECT: Eliminates duplication — a change to query params only
//              needs to happen in one place. Reduces risk of divergence.

/**
 * Generates a PKCE code verifier and challenge.
 *
 * BEFORE: No PKCE support (e-issue: Microsoft's OAuth 2.1 recommendations
 *         require PKCE; without it, authorization code interception is possible).
 * AFTER: code_verifier and code_challenge generated per auth flow.
 * GOOD EFFECT: Prevents authorization code interception attacks at the
 *              redirect URI — critical for public clients and recommended
 *              for confidential clients per OAuth 2.1.
 *
 * @returns {{ verifier: string, challenge: string }}
 */
function generatePKCE() {
  const verifier = crypto.randomBytes(32).toString('base64url');
  const challenge = crypto.createHash('sha256').update(verifier).digest('base64url');
  return { verifier, challenge };
}

/**
 * Builds the Microsoft OAuth2 authorization URL.
 *
 * BEFORE: URL construction was inline in the handler AND in oauth-server.js.
 * AFTER: Shared utility function.
 * GOOD EFFECT: Single source of truth for auth URL shape; both the MCP
 *              tool and the OAuth server callback handler use the same builder.
 *
 * @param {object} authConfig - Auth configuration
 * @param {object} [options] - Options like state, PKCE challenge
 * @returns {{ url: string, codeVerifier: string|null }}
 */
function buildAuthUrl(authConfig, options = {}) {
  const pkce = generatePKCE();
  const state = options.state || crypto.randomBytes(16).toString('hex');

  // BEFORE: Scopes were a hardcoded string literal inside the function.
  // AFTER: Read from authConfig.scopes (which comes from config.js).
  // GOOD EFFECT: Scopes are defined once in config.js and match the
  //              Azure Portal app registration — no hidden divergence.
  const params = {
    client_id: authConfig.clientId,
    response_type: 'code',
    redirect_uri: authConfig.redirectUri,
    scope: Array.isArray(authConfig.scopes) ? authConfig.scopes.join(' ') : authConfig.scopes,
    response_mode: 'query',
    state: state,
    // BEFORE: No PKCE parameters.
    // AFTER: code_challenge and code_challenge_method included.
    // GOOD EFFECT: Enables PKCE flow for OAuth 2.1 compliance.
    code_challenge: pkce.challenge,
    code_challenge_method: 'S256'
  };

  const url = `${authConfig.authEndpoint}?${querystring.stringify(params)}`;
  return { url, codeVerifier: pkce.verifier, state };
}

// ─── Tool Handlers (separated from schema definitions) ────────────────
// BEFORE: Each tool object contained both the MCP schema AND the handler
//         function in one combined object, mixing concerns.
// AFTER: Handlers are standalone functions; schema objects reference them.
// GOOD EFFECT: Handlers are independently testable without parsing the
//              MCP schema; schemas can be validated separately.

/**
 * About tool handler
 * @param {object} _args - unused
 * @param {object} serverConfig - Injected server config
 * @returns {object} MCP response
 */
async function handleAbout(serverConfig = {}) {
  const version = serverConfig.SERVER_VERSION || '1.0.0';
  return {
    content: [{
      type: "text",
      text: `M365 Assistant MCP Server v${version}\n\nProvides access to Microsoft 365 services through Microsoft Graph API:\n- Outlook (email, calendar, folders, rules)\n- OneDrive (files, folders, sharing)\n- Power Automate (flows, environments, runs)\n\nModular architecture for improved maintainability.`
    }]
  };
}

/**
 * Authentication tool handler
 *
 * BEFORE: Built the auth URL inline with hardcoded scopes and no PKCE.
 * AFTER: Delegates to buildAuthUrl() with PKCE and config-driven scopes.
 * GOOD EFFECT: Auth URL is built consistently everywhere, PKCE is enabled,
 *              and the code_verifier is returned for use in the callback.
 *
 * @param {object} args - Tool arguments
 * @param {object} authConfig - Injected auth configuration
 * @param {object} tokenManager - Injected token manager instance
 * @returns {object} MCP response
 */
async function handleAuthenticate(args, authConfig, tokenManager) {
  const force = args && args.force === true;

  // BEFORE: if (config.USE_TEST_MODE) { tokenManager.createTestTokens(); ... }
  //         — read global config directly
  // AFTER: authConfig.useTestMode — injected, not global.
  // GOOD EFFECT: Test mode is injectable, not a hidden global flag.
  if (authConfig.useTestMode && tokenManager) {
    await tokenManager.createTestTokens();
    return {
      content: [{
        type: "text",
        text: 'Successfully authenticated with Microsoft Graph API (test mode)'
      }]
    };
  }

  // BEFORE: const authUrl = `${config.AUTH_CONFIG.authServerUrl}/auth?client_id=${config.AUTH_CONFIG.clientId}`;
  //         — inline URL construction, no PKCE, hardcoded scopes.
  // AFTER: const { url, codeVerifier } = buildAuthUrl(authConfig);
  // GOOD EFFECT: URL built via shared utility with PKCE and config-driven scopes.
  const { url, codeVerifier, state } = buildAuthUrl(authConfig);

  // NOTE: The codeVerifier must be stored (in session/memory) for use during
  // the callback token exchange. The consuming server is responsible for this.
  return {
    content: [{
      type: "text",
      text: `Authentication required. Please visit the following URL to authenticate with Microsoft:\n\n${url}\n\nAfter authentication, you will be redirected back to this application.`
    }],
    // BEFORE: No metadata returned alongside the auth URL.
    // AFTER: codeVerifier and state returned as metadata for the callback.
    // GOOD EFFECT: The consuming server can store these for PKCE verification
    //              and CSRF state validation during the callback.
    _metadata: {
      codeVerifier,
      state
    }
  };
}

/**
 * Check authentication status tool handler
 *
 * BEFORE: Called tokenManager.loadTokenCache() — synchronous disk read.
 * AFTER: Calls tokenManager.getValidAccessToken() — async, uses cache.
 * GOOD EFFECT: Non-blocking; leverages the in-memory token cache.
 *
 * @param {object} _args - unused
 * @param {object} tokenManager - Injected token manager instance
 * @returns {object} MCP response
 */
async function handleCheckAuthStatus(tokenManager) {
  console.error('[CHECK-AUTH-STATUS] Starting authentication status check');

  if (!tokenManager) {
    return {
      content: [{ type: "text", text: "Token manager not available." }]
    };
  }

  // BEFORE: const tokens = tokenManager.loadTokenCache();
  //         — synchronous, no refresh, no cache.
  // AFTER:  const token = await tokenManager.getValidAccessToken();
  //         — async, checks cache, auto-refreshes if expired.
  // GOOD EFFECT: Status check reflects the true live state of the token
  //              (including auto-refresh), not just what's on disk.
  const token = await tokenManager.getValidAccessToken();

  if (!token) {
    console.error('[CHECK-AUTH-STATUS] No valid access token found');
    return {
      content: [{ type: "text", text: "Not authenticated" }]
    };
  }

  const expiryTime = tokenManager.getExpiryTime();
  const expiryDate = expiryTime ? new Date(expiryTime).toLocaleString() : 'unknown';

  console.error(`[CHECK-AUTH-STATUS] Authenticated. Token expires at: ${expiryDate}`);
  return {
    content: [{ type: "text", text: `Authenticated and ready. Token expires at: ${expiryDate}` }]
  };
}

// ─── Tool Definitions (schema separated from handlers) ────────────────
// BEFORE: Schema and handler were combined in one object, making it
//         impossible to test handlers without the MCP schema overhead.
// AFTER: Clean separation — schemas reference handler functions.
// GOOD EFFECT: Schemas are validatable and serialisable independently;
//              handlers are unit-testable without MCP infrastructure.

const authTools = [
  {
    name: "about",
    description: "Returns information about this M365 Assistant server",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleAbout
  },
  {
    name: "authenticate",
    description: "Authenticate with Microsoft Graph API to access Outlook data",
    inputSchema: {
      type: "object",
      properties: {
        force: {
          type: "boolean",
          description: "Force re-authentication even if already authenticated"
        }
      },
      required: []
    },
    handler: handleAuthenticate
  },
  {
    name: "check-auth-status",
    description: "Check the current authentication status with Microsoft Graph API",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleCheckAuthStatus
  }
];

module.exports = {
  authTools,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus,
  // BEFORE: Not exported — URL builder was inline, PKCE didn't exist.
  // AFTER: Exported for shared use by oauth-server.js and tests.
  // GOOD EFFECT: oauth-server.js can import buildAuthUrl() instead of
  //              duplicating URL construction logic.
  buildAuthUrl,
  generatePKCE
};