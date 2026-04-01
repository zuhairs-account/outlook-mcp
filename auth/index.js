/**
 * Authentication module for Outlook MCP server
 * @module auth
 */

// BEFORE: const tokenManager = require('./token-manager');
// AFTER: Import the class, not a singleton instance.
// GOOD EFFECT: Enables factory-based DI — consumers are decoupled from
//              the concrete token-manager.js implementation.
const TokenManager = require('./token-manager');
const { authTools } = require('./tools');

/**
 * Factory function to create a TokenManager instance with the given config.
 * Consumers should use this instead of importing token-manager directly.
 * @param {object} [config] - Optional config overrides
 * @returns {TokenManager} - A configured TokenManager instance
 *
 * GOOD EFFECT: Enables dependency injection — tests can pass mock config,
 * production code passes real config, and no one is coupled to the concrete
 * file path of token-manager.js.
 */
function createTokenManager(config) {
  return new TokenManager(config);
}

/**
 * Ensures the user is authenticated and returns an access token.
 *
 * @param {TokenManager} tokenManagerInstance - Injected token manager
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(tokenManagerInstance, forceNew = false) {
  // BEFORE: No error boundary — raw exceptions propagated to the MCP dispatch.
  // AFTER: Entire function wrapped in try-catch.
  // GOOD EFFECT: Prevents unhandled exceptions from crashing the MCP server process.
  try {
    if (forceNew) {
      throw new Error('Authentication required. Please re-authenticate.');
    }

    // BEFORE: const accessToken = tokenManager.getAccessToken();
    //         — called synchronous getAccessToken() on a module-level singleton
    // AFTER: await tokenManagerInstance.getValidAccessToken()
    //        — uses injected instance, async, checks expiry + auto-refreshes
    // GOOD EFFECT: Tokens are auto-refreshed if expired (no stale-token bugs),
    //              and the injected instance makes this testable.
    const accessToken = await tokenManagerInstance.getValidAccessToken();
    if (!accessToken) {
      throw new Error('Authentication required. No valid token found.');
    }

    return accessToken;
  } catch (error) {
    // BEFORE: (no catch block existed)
    // AFTER: Catch and re-throw with structured message.
    // GOOD EFFECT: MCP client receives a descriptive error rather than
    //              an opaque stack trace or process crash.
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Authentication failed: ${message}`);
  }
}

module.exports = {
  // BEFORE: tokenManager (concrete singleton export)
  // AFTER: createTokenManager (factory function)
  // GOOD EFFECT: Barrel enforces the intended interface — consumers go
  //              through the factory and never import internals directly.
  createTokenManager,
  authTools,
  ensureAuthenticated
};